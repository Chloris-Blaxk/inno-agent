/**
 * Tests for the office_* agent tools. The UNO sidecar is mocked; these tests
 * cover path guarding, argument validation and the save-on-success policy.
 */

import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const callMock = vi.fn();

vi.mock("../office/uno-sidecar.js", () => ({
	getUnoSidecar: () => ({ call: callMock }),
}));

vi.mock("../office/engine-locator.js", () => ({
	resolveLibreOffice: () => ({
		available: true,
		sofficePath: "/fake/soffice",
		source: "bundled",
		version: "LibreOffice 26.8.1.1 test",
	}),
}));

import { createOfficeTools } from "./office-tools.js";

let workspace: string;
let tools: ReturnType<typeof createOfficeTools>;

function toolByName(name: string) {
	const tool = tools.find((t) => t.name === name);
	if (!tool) throw new Error(`tool ${name} not registered`);
	return tool;
}

async function run(name: string, params: Record<string, unknown>) {
	return toolByName(name).execute(
		"call-1",
		params,
		undefined as never,
		undefined as never,
		{ cwd: workspace } as never,
	);
}

function textOf(result: Awaited<ReturnType<typeof run>>): string {
	return (result.content[0] as { text: string }).text;
}

beforeEach(() => {
	workspace = mkdtempSync(join(tmpdir(), "office-tools-"));
	callMock.mockReset();
	tools = createOfficeTools(() => workspace);
});

afterEach(() => {
	rmSync(workspace, { recursive: true, force: true });
});

describe("office tools", () => {
	it("registers the five-tool surface", () => {
		expect(tools.map((t) => t.name)).toEqual([
			"office_engine_status",
			"office_open_document",
			"office_edit_document",
			"office_render_document",
			"office_close_document",
		]);
	});

	it("reports engine availability", async () => {
		const res = await run("office_engine_status", {});
		expect(textOf(res)).toContain("LibreOffice 可用");
		expect((res.details as { source: string }).source).toBe("bundled");
	});

	it("rejects paths outside the workspace before touching the engine", async () => {
		const res = await run("office_open_document", { path: "/etc/passwd" });
		expect(textOf(res)).toContain("已拒绝");
		expect(callMock).not.toHaveBeenCalled();
	});

	it("opens a workspace document and returns the outline", async () => {
		writeFileSync(join(workspace, "a.docx"), "fake");
		callMock
			.mockResolvedValueOnce({ ok: true, result: { docId: "d1", kind: "doc", path: "a.docx" } })
			.mockResolvedValueOnce({ ok: true, result: { paragraphCount: 3, paragraphs: [] } });
		const res = await run("office_open_document", { path: "a.docx" });
		expect(textOf(res)).toContain("docId: d1");
		expect((res.details as { kind: string }).kind).toBe("doc");
		expect(callMock).toHaveBeenCalledTimes(2);
	});

	it("rejects empty op batches", async () => {
		const res = await run("office_edit_document", { docId: "d1", ops: [] });
		expect(textOf(res)).toContain("ops 为空");
		expect(callMock).not.toHaveBeenCalled();
	});

	it("saves only when every op succeeds", async () => {
		callMock.mockResolvedValueOnce({
			ok: true,
			result: { results: [{ op: "replace_text", ok: true }] },
		}).mockResolvedValueOnce({ ok: true, result: { saved: "/ws/a.docx" } });
		const res = await run("office_edit_document", {
			docId: "d1",
			ops: [{ op: "replace_text", find: "a", replace: "b" }],
			save: true,
		});
		expect(textOf(res)).toContain("已保存");
		expect(callMock).toHaveBeenLastCalledWith("save", { docId: "d1" });
	});

	it("skips saving when any op failed", async () => {
		callMock.mockResolvedValueOnce({
			ok: true,
			result: {
				results: [
					{ op: "replace_text", ok: true },
					{ op: "set_paragraph", ok: false, error: "index out of range" },
				],
			},
		});
		const res = await run("office_edit_document", {
			docId: "d1",
			ops: [{ op: "replace_text" }, { op: "set_paragraph" }],
			save: true,
		});
		expect(textOf(res)).toContain("已跳过保存");
		expect(callMock).toHaveBeenCalledTimes(1);
	});

	it("save-as rejects paths outside the workspace", async () => {
		callMock.mockResolvedValueOnce({
			ok: true,
			result: { results: [{ op: "replace_text", ok: true }] },
		});
		const res = await run("office_edit_document", {
			docId: "d1",
			ops: [{ op: "replace_text", find: "a", replace: "b" }],
			saveAs: "/tmp/evil.docx",
		});
		expect(textOf(res)).toContain("已拒绝");
	});

	it("closes documents and surfaces engine errors as text", async () => {
		callMock.mockResolvedValueOnce({ ok: false, error: "unknown docId d9" });
		const res = await run("office_close_document", { docId: "d9" });
		expect(textOf(res)).toContain("unknown docId d9");
	});
});
