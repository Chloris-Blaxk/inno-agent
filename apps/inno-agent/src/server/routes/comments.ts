import type { IncomingMessage as HttpReq, ServerResponse } from "node:http";
import { existsSync, statSync } from "node:fs";
import { spawnSync } from "node:child_process";
import { basename, extname } from "node:path";
import { completeSideQuestion } from "../../agent/pi-runner.js";
import { logger } from "../../logger.js";
import { parseDocument } from "../../memory/l2/document-parser.js";
import {
	MAX_COMMENT_ANCHOR_LEN,
	MAX_COMMENT_QUESTION_LEN,
	CommentStore,
	type DocCommentAnchor,
} from "../../comments/comment-store.js";
import { json, matchRoute, readBody } from "../http-helpers.js";
import { normalizeWorkspaceRelativePath, safeJoinReal } from "../file-helpers.js";
import { TEMP_WORKSPACE_ID, type WorkspaceRegistry } from "../../workspace/workspace-registry.js";

/**
 * /api/comments route domain — anchored comment conversations on previewed
 * workspace documents (ChatGPT-canvas-style "批注对话").
 *
 * Answers go through the same stateless side-channel completion as btw
 * (completeSideQuestion): no tools, never enters the main session transcript
 * or the prompt queue. The difference is the context digest: instead of recent
 * chat messages, the model gets the document name, the anchored quote, and an
 * excerpt of the document around the anchor.
 */

export interface CommentsRouteContext {
	dataDir: string;
	workspaceRegistry: WorkspaceRegistry;
}

const ANCHOR_CONTEXT_CHARS = 120;
const EXCERPT_WINDOW_CHARS = 1500;

const stores = new Map<string, CommentStore>();
function commentStore(dataDir: string): CommentStore {
	let existing = stores.get(dataDir);
	if (!existing) {
		existing = new CommentStore(dataDir);
		stores.set(dataDir, existing);
	}
	return existing;
}

function isRecord(value: unknown): value is Record<string, unknown> {
	return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

/** Same workspace path policy as the workspace file routes (symlink-safe). */
function resolveWorkspaceFile(
	registry: WorkspaceRegistry,
	workspaceId: string | null,
	userPath: string,
): { wsId: string; filePath: string } | null {
	const wsId = workspaceId ?? TEMP_WORKSPACE_ID;
	const root = registry.resolveWorkspaceDir(wsId);
	if (!root) return null;
	const normalized = normalizeWorkspaceRelativePath(userPath);
	if (normalized.startsWith("/") || /^[A-Za-z]:\//.test(normalized)) return null;
	const filePath = safeJoinReal(root, normalized);
	if (!filePath || !existsSync(filePath) || !statSync(filePath).isFile()) return null;
	return { wsId, filePath };
}

function parseAnchor(raw: unknown): DocCommentAnchor | null {
	if (!isRecord(raw)) return null;
	const { text, prefix, suffix } = raw;
	if (typeof text !== "string" || !text.trim()) return null;
	if (text.length > MAX_COMMENT_ANCHOR_LEN) return null;
	return {
		text,
		prefix: typeof prefix === "string" ? prefix.slice(-ANCHOR_CONTEXT_CHARS) : "",
		suffix: typeof suffix === "string" ? suffix.slice(0, ANCHOR_CONTEXT_CHARS) : "",
	};
}

/**
 * Crude .docx text extraction without LibreOffice: unzip word/document.xml
 * and strip OOXML tags. Paragraph boundaries become newlines. Good enough for
 * a comment context digest; the rich parser (liteparse) stays the first choice.
 */
function extractDocxTextFallback(filePath: string): string | null {
	const proc = spawnSync("unzip", ["-p", filePath, "word/document.xml"], {
		encoding: "utf8",
		maxBuffer: 64 * 1024 * 1024,
	});
	if (proc.status !== 0 || !proc.stdout) return null;
	const text = proc.stdout
		.replace(/<w:tab\b[^>]*\/>/g, "\t")
		.replace(/<\/w:p>/g, "\n")
		.replace(/<[^>]+>/g, "")
		.replace(/&lt;/g, "<")
		.replace(/&gt;/g, ">")
		.replace(/&quot;/g, '"')
		.replace(/&apos;/g, "'")
		.replace(/&amp;/g, "&")
		.replace(/\n{3,}/g, "\n\n")
		.trim();
	return text || null;
}

/** Build the document-side digest for a comment question. */function buildCommentDigest(docName: string, fullText: string, anchor: DocCommentAnchor): string {
	const anchorIdx = fullText.indexOf(anchor.text);
	let excerpt: string;
	if (anchorIdx >= 0) {
		const start = Math.max(0, anchorIdx - EXCERPT_WINDOW_CHARS);
		const end = Math.min(fullText.length, anchorIdx + anchor.text.length + EXCERPT_WINDOW_CHARS);
		excerpt = `${start > 0 ? "…" : ""}${fullText.slice(start, end)}${end < fullText.length ? "…" : ""}`;
	} else {
		// Anchor not found verbatim (document may have changed): fall back to the
		// document head so the model still has material to work with.
		excerpt = fullText.slice(0, EXCERPT_WINDOW_CHARS * 2);
	}
	return [
		`学生正在预览文档《${docName}》,并对其中一段内容发起批注提问。`,
		"",
		"被批注的原文(引用):",
		`"${anchor.text}"`,
		"",
		"原文附近的内容节选(供你理解上下文):",
		excerpt,
		"",
		anchorIdx < 0 ? "注意:被批注的原文在当前文档中未逐字找到,文档可能已被修改,回答时请说明这一不确定性。" : "",
		"以上仅供你理解上下文。请直接回答学生针对这段原文的问题,不要复述题目。",
	].filter(Boolean).join("\n");
}

export async function handleCommentsRoutes(
	req: HttpReq,
	res: ServerResponse,
	method: string,
	url: string,
	ctx: CommentsRouteContext,
): Promise<boolean> {
	if (method === "GET" && url.startsWith("/api/comments?")) {
		const params = new URL(url, "http://localhost").searchParams;
		const wsId = params.get("workspaceId");
		const docPath = params.get("path") ?? "";
		const root = ctx.workspaceRegistry.resolveWorkspaceDir(wsId ?? TEMP_WORKSPACE_ID);
		if (!root || !docPath) {
			json(res, 400, { error: "Missing workspaceId/path" });
			return true;
		}
		const normalized = normalizeWorkspaceRelativePath(docPath);
		json(res, 200, commentStore(ctx.dataDir).read(wsId ?? TEMP_WORKSPACE_ID, normalized));
		return true;
	}

	if (method === "POST" && url === "/api/comments/ask") {
		const body = (await readBody(req)) as Record<string, unknown>;
		const wsIdParam = typeof body.workspaceId === "string" ? body.workspaceId : null;
		const docPath = typeof body.path === "string" ? body.path : "";
		const question = typeof body.question === "string" ? body.question.trim() : "";
		const threadId = typeof body.threadId === "string" ? body.threadId : null;
		if (!docPath || !question) {
			json(res, 400, { error: "Missing path/question" });
			return true;
		}
		if (question.length > MAX_COMMENT_QUESTION_LEN) {
			json(res, 400, { error: `question too long (max ${MAX_COMMENT_QUESTION_LEN} chars)` });
			return true;
		}
		const resolved = resolveWorkspaceFile(ctx.workspaceRegistry, wsIdParam, docPath);
		if (!resolved) {
			json(res, 404, { error: "Workspace file not found" });
			return true;
		}
		const store = commentStore(ctx.dataDir);
		const normalizedPath = normalizeWorkspaceRelativePath(docPath);

		let anchor: DocCommentAnchor | null = null;
		let priorTurns: Array<{ question: string; answer: string }> = [];
		if (threadId) {
			const existing = store.read(resolved.wsId, normalizedPath).threads.find((t) => t.id === threadId);
			if (!existing) {
				json(res, 404, { error: "Comment thread not found" });
				return true;
			}
			anchor = existing.anchor;
			const turns = existing.turns;
			for (let i = 0; i + 1 < turns.length; i += 2) {
				if (turns[i].role === "user" && turns[i + 1].role === "assistant") {
					priorTurns.push({ question: turns[i].content, answer: turns[i + 1].content });
				}
			}
		} else {
			anchor = parseAnchor(body.anchor);
			if (!anchor) {
				json(res, 400, { error: "Missing or malformed anchor" });
				return true;
			}
		}

		let digest: string;
		try {
			const parsed = await parseDocument(resolved.filePath);
			digest = buildCommentDigest(basename(resolved.filePath), parsed.text, anchor);
		} catch (err) {
			// liteparse needs LibreOffice for office formats; for .docx fall back
			// to a crude built-in extractor so comments work without it.
			const fallbackText = extname(resolved.filePath).toLowerCase() === ".docx"
				? extractDocxTextFallback(resolved.filePath)
				: null;
			if (fallbackText) {
				digest = buildCommentDigest(basename(resolved.filePath), fallbackText, anchor);
			} else {
				logger.warn({ err, file: resolved.filePath }, "comment ask: document parse failed");
				json(res, 422, { error: "Failed to extract document text for the comment context" });
				return true;
			}
		}

		const controller = new AbortController();
		const abort = () => controller.abort();
		const onRequestClose = () => { if (req.complete === false) abort(); };
		const onResponseClose = () => { if (!res.writableEnded) abort(); };
		req.once("aborted", abort);
		req.once("close", onRequestClose);
		if (typeof res.once === "function") res.once("close", onResponseClose);
		try {
			const answer = await completeSideQuestion({
				contextDigest: digest,
				thread: priorTurns,
				question,
				signal: controller.signal,
			});
			if (controller.signal.aborted) return true;
			if (!answer) {
				json(res, 502, { error: "Comment completion failed or timed out" });
				return true;
			}
			const thread = threadId
				? store.appendTurn(resolved.wsId, normalizedPath, threadId, question, answer)
				: store.createThread(resolved.wsId, normalizedPath, anchor, question, answer);
			if (!thread) {
				json(res, 500, { error: "Failed to persist the comment thread" });
				return true;
			}
			json(res, 200, { thread });
			return true;
		} finally {
			req.removeListener("aborted", abort);
			req.removeListener("close", onRequestClose);
			if (typeof res.removeListener === "function") res.removeListener("close", onResponseClose);
		}
	}

	const updateMatch = matchRoute("POST", method, url, "/api/comments/update");
	if (updateMatch) {
		const body = (await readBody(req)) as Record<string, unknown>;
		const wsId = typeof body.workspaceId === "string" ? body.workspaceId : TEMP_WORKSPACE_ID;
		const docPath = typeof body.path === "string" ? normalizeWorkspaceRelativePath(body.path) : "";
		const threadId = typeof body.threadId === "string" ? body.threadId : "";
		const action = body.action;
		if (!docPath || !threadId || (action !== "resolve" && action !== "reopen" && action !== "delete")) {
			json(res, 400, { error: "Missing path/threadId or unknown action" });
			return true;
		}
		const store = commentStore(ctx.dataDir);
		if (action === "delete") {
			if (!store.deleteThread(wsId, docPath, threadId)) {
				json(res, 404, { error: "Comment thread not found" });
				return true;
			}
			json(res, 200, { ok: true });
			return true;
		}
		const thread = store.setStatus(wsId, docPath, threadId, action === "resolve" ? "resolved" : "open");
		if (!thread) {
			json(res, 404, { error: "Comment thread not found" });
			return true;
		}
		json(res, 200, { thread });
		return true;
	}

	return false;
}
