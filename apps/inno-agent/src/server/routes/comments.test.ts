import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { PassThrough } from "node:stream";
import type { IncomingMessage as HttpReq, ServerResponse } from "node:http";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { handleCommentsRoutes, type CommentsRouteContext } from "./comments.js";
import { completeSideQuestion } from "../../agent/pi-runner.js";
import { parseDocument } from "../../memory/l2/document-parser.js";
import type { WorkspaceRegistry } from "../../workspace/workspace-registry.js";

vi.mock("../../agent/pi-runner.js", async (importOriginal) => {
	const original = await importOriginal<typeof import("../../agent/pi-runner.js")>();
	return { ...original, completeSideQuestion: vi.fn() };
});

vi.mock("../../memory/l2/document-parser.js", () => ({
	parseDocument: vi.fn(),
}));

const mockedComplete = vi.mocked(completeSideQuestion);
const mockedParse = vi.mocked(parseDocument);

function fakeReq(payload: unknown): HttpReq & PassThrough {
	const stream = new PassThrough() as HttpReq & PassThrough;
	stream.headers = {};
	stream.end(JSON.stringify(payload));
	return stream;
}

function fakeRes(): ServerResponse & { statusCode: number; body: () => unknown } {
	let body = "";
	const res = {
		statusCode: 0,
		writeHead(status: number) {
			res.statusCode = status;
			return res;
		},
		end(chunk?: string) {
			body = chunk ?? "";
			return res;
		},
		body: () => (body ? JSON.parse(body) : undefined),
	} as unknown as ServerResponse & { statusCode: number; body: () => unknown };
	return res;
}

const ANCHOR = { text: "勾股定理", prefix: "第二章 ", suffix: " 是基本定理" };
const DOC_TEXT = `第一章 引言。第二章 勾股定理 是基本定理。第三章 总结。`;

let dataDir: string;
let wsRoot: string;
let ctx: CommentsRouteContext;

beforeEach(() => {
	dataDir = mkdtempSync(join(tmpdir(), "comments-routes-"));
	wsRoot = join(dataDir, "ws");
	mkdirSync(wsRoot, { recursive: true });
	writeFileSync(join(wsRoot, "a.docx"), "fake");
	const registry = {
		resolveWorkspaceDir: (id: string | null | undefined) => (id === "ws1" ? wsRoot : null),
	} as unknown as WorkspaceRegistry;
	ctx = { dataDir, workspaceRegistry: registry };
	mockedComplete.mockReset();
	mockedParse.mockReset();
	mockedParse.mockResolvedValue({ text: DOC_TEXT, pageCount: 1, pages: [{ pageNumber: 1, text: DOC_TEXT }] });
});

afterEach(() => {
	rmSync(dataDir, { recursive: true, force: true });
});

describe("handleCommentsRoutes", () => {
	it("GET lists threads for a document", async () => {
		const res = fakeRes();
		const handled = await handleCommentsRoutes(
			fakeReq({}), res, "GET", "/api/comments?workspaceId=ws1&path=a.docx", ctx,
		);
		expect(handled).toBe(true);
		expect(res.statusCode).toBe(200);
		expect(res.body()).toEqual({ docPath: "a.docx", threads: [] });
	});

	it("POST ask creates a thread with the completion answer", async () => {
		mockedComplete.mockResolvedValue("直角三角形两直角边平方和等于斜边平方。");
		const res = fakeRes();
		await handleCommentsRoutes(fakeReq({
			workspaceId: "ws1", path: "a.docx", anchor: ANCHOR, question: "这句话什么意思？",
		}), res, "POST", "/api/comments/ask", ctx);
		expect(res.statusCode).toBe(200);
		const thread = (res.body() as { thread: { turns: Array<{ role: string; content: string }>; anchor: { text: string } } }).thread;
		expect(thread.anchor.text).toBe("勾股定理");
		expect(thread.turns).toHaveLength(2);
		expect(thread.turns[0]).toMatchObject({ role: "user", content: "这句话什么意思？" });
		expect(thread.turns[1].role).toBe("assistant");
		// The digest must contain the anchored quote and the document excerpt.
		const digest = mockedComplete.mock.calls[0][0].contextDigest;
		expect(digest).toContain("勾股定理");
		expect(digest).toContain("a.docx");
	});

	it("POST ask with threadId appends a turn to the existing thread", async () => {
		mockedComplete.mockResolvedValue("第一答");
		const first = fakeRes();
		await handleCommentsRoutes(fakeReq({
			workspaceId: "ws1", path: "a.docx", anchor: ANCHOR, question: "问题一",
		}), first, "POST", "/api/comments/ask", ctx);
		const threadId = (first.body() as { thread: { id: string } }).thread.id;

		mockedComplete.mockResolvedValue("第二答");
		const second = fakeRes();
		await handleCommentsRoutes(fakeReq({
			workspaceId: "ws1", path: "a.docx", threadId, question: "追问",
		}), second, "POST", "/api/comments/ask", ctx);
		expect(second.statusCode).toBe(200);
		const thread = (second.body() as { thread: { turns: unknown[] } }).thread;
		expect(thread.turns).toHaveLength(4);
		// Prior turns are forwarded to the completion as conversation thread.
		expect(mockedComplete.mock.calls[1][0].thread).toEqual([{ question: "问题一", answer: "第一答" }]);
	});

	it("POST ask rejects anchor-less new threads and missing files", async () => {
		const res = fakeRes();
		await handleCommentsRoutes(fakeReq({
			workspaceId: "ws1", path: "a.docx", question: "没有锚点",
		}), res, "POST", "/api/comments/ask", ctx);
		expect(res.statusCode).toBe(400);

		const missing = fakeRes();
		await handleCommentsRoutes(fakeReq({
			workspaceId: "ws1", path: "nope.docx", anchor: ANCHOR, question: "q",
		}), missing, "POST", "/api/comments/ask", ctx);
		expect(missing.statusCode).toBe(404);
	});

	it("POST ask rejects path traversal outside the workspace", async () => {
		const res = fakeRes();
		await handleCommentsRoutes(fakeReq({
			workspaceId: "ws1", path: "../../etc/passwd", anchor: ANCHOR, question: "q",
		}), res, "POST", "/api/comments/ask", ctx);
		expect(res.statusCode).toBe(404);
	});

	it("POST ask reports parse failures as 422", async () => {
		mockedParse.mockRejectedValue(new Error("boom"));
		const res = fakeRes();
		await handleCommentsRoutes(fakeReq({
			workspaceId: "ws1", path: "a.docx", anchor: ANCHOR, question: "q",
		}), res, "POST", "/api/comments/ask", ctx);
		expect(res.statusCode).toBe(422);
	});

	it("POST update resolves, reopens and deletes threads", async () => {
		mockedComplete.mockResolvedValue("答");
		const created = fakeRes();
		await handleCommentsRoutes(fakeReq({
			workspaceId: "ws1", path: "a.docx", anchor: ANCHOR, question: "q",
		}), created, "POST", "/api/comments/ask", ctx);
		const threadId = (created.body() as { thread: { id: string } }).thread.id;

		const resolved = fakeRes();
		await handleCommentsRoutes(fakeReq({
			workspaceId: "ws1", path: "a.docx", threadId, action: "resolve",
		}), resolved, "POST", "/api/comments/update", ctx);
		expect((resolved.body() as { thread: { status: string } }).thread.status).toBe("resolved");

		const reopened = fakeRes();
		await handleCommentsRoutes(fakeReq({
			workspaceId: "ws1", path: "a.docx", threadId, action: "reopen",
		}), reopened, "POST", "/api/comments/update", ctx);
		expect((reopened.body() as { thread: { status: string } }).thread.status).toBe("open");

		const deleted = fakeRes();
		await handleCommentsRoutes(fakeReq({
			workspaceId: "ws1", path: "a.docx", threadId, action: "delete",
		}), deleted, "POST", "/api/comments/update", ctx);
		expect(deleted.statusCode).toBe(200);

		const listed = fakeRes();
		await handleCommentsRoutes(fakeReq({}), listed, "GET", "/api/comments?workspaceId=ws1&path=a.docx", ctx);
		expect((listed.body() as { threads: unknown[] }).threads).toHaveLength(0);
	});
});
