import { createHash } from "node:crypto";
import { existsSync, mkdirSync, readFileSync, renameSync, writeFileSync } from "node:fs";
import { join } from "node:path";

/**
 * Anchored document comment threads ("预览侧批注对话"), persisted per workspace
 * document under `<dataDir>/comments/<workspaceId>/<docKey>.json`.
 *
 * A thread is anchored to a text selection by the selected text plus a bit of
 * surrounding context (prefix/suffix), so the frontend can re-locate it in the
 * rendered document even after unrelated edits. This mirrors ChatGPT canvas
 * comments: the anchor is content-based, not offset-based.
 */

export interface DocCommentAnchor {
	/** The exact selected text. */
	text: string;
	/** Up to ~120 chars of text immediately before the selection. */
	prefix: string;
	/** Up to ~120 chars of text immediately after the selection. */
	suffix: string;
}

export interface DocCommentTurn {
	role: "user" | "assistant";
	content: string;
	ts: string;
}

export interface DocCommentThread {
	id: string;
	anchor: DocCommentAnchor;
	turns: DocCommentTurn[];
	status: "open" | "resolved";
	createdAt: string;
	updatedAt: string;
}

export interface DocCommentFile {
	docPath: string;
	threads: DocCommentThread[];
}

export const MAX_COMMENT_THREADS_PER_DOC = 100;
export const MAX_COMMENT_ANCHOR_LEN = 2000;
export const MAX_COMMENT_QUESTION_LEN = 4000;
export const MAX_COMMENT_TURNS_PER_THREAD = 40;

function safeSegment(value: string): string {
	const cleaned = value.replace(/[^A-Za-z0-9._-]/g, "_");
	return cleaned || "_";
}

function docKey(docPath: string): string {
	return createHash("sha256").update(docPath).digest("hex").slice(0, 24);
}

export class CommentStore {
	private readonly baseDir: string;

	constructor(dataDir: string) {
		this.baseDir = join(dataDir, "comments");
	}

	private fileFor(workspaceId: string, docPath: string): string {
		return join(this.baseDir, safeSegment(workspaceId), `${docKey(docPath)}.json`);
	}

	read(workspaceId: string, docPath: string): DocCommentFile {
		const file = this.fileFor(workspaceId, docPath);
		if (!existsSync(file)) return { docPath, threads: [] };
		try {
			const raw = JSON.parse(readFileSync(file, "utf8")) as DocCommentFile;
			if (!raw || !Array.isArray(raw.threads)) return { docPath, threads: [] };
			return { docPath, threads: raw.threads };
		} catch {
			return { docPath, threads: [] };
		}
	}

	private write(workspaceId: string, docPath: string, data: DocCommentFile): void {
		const file = this.fileFor(workspaceId, docPath);
		mkdirSync(join(file, ".."), { recursive: true });
		const tmp = `${file}.${process.pid}.tmp`;
		writeFileSync(tmp, JSON.stringify(data, null, 2), "utf8");
		renameSync(tmp, file);
	}

	createThread(workspaceId: string, docPath: string, anchor: DocCommentAnchor, question: string, answer: string): DocCommentThread {
		const data = this.read(workspaceId, docPath);
		if (data.threads.length >= MAX_COMMENT_THREADS_PER_DOC) {
			throw new Error(`Too many comment threads on this document (max ${MAX_COMMENT_THREADS_PER_DOC})`);
		}
		const now = new Date().toISOString();
		const thread: DocCommentThread = {
			id: createHash("sha256").update(`${docPath}${now}${Math.random()}`).digest("hex").slice(0, 16),
			anchor,
			turns: [
				{ role: "user", content: question, ts: now },
				{ role: "assistant", content: answer, ts: now },
			],
			status: "open",
			createdAt: now,
			updatedAt: now,
		};
		data.threads.push(thread);
		this.write(workspaceId, docPath, data);
		return thread;
	}

	appendTurn(workspaceId: string, docPath: string, threadId: string, question: string, answer: string): DocCommentThread | null {
		const data = this.read(workspaceId, docPath);
		const thread = data.threads.find((t) => t.id === threadId);
		if (!thread) return null;
		if (thread.turns.length >= MAX_COMMENT_TURNS_PER_THREAD) {
			throw new Error(`Comment thread is too long (max ${MAX_COMMENT_TURNS_PER_THREAD} turns)`);
		}
		const now = new Date().toISOString();
		thread.turns.push({ role: "user", content: question, ts: now });
		thread.turns.push({ role: "assistant", content: answer, ts: now });
		thread.updatedAt = now;
		this.write(workspaceId, docPath, data);
		return thread;
	}

	setStatus(workspaceId: string, docPath: string, threadId: string, status: "open" | "resolved"): DocCommentThread | null {
		const data = this.read(workspaceId, docPath);
		const thread = data.threads.find((t) => t.id === threadId);
		if (!thread) return null;
		thread.status = status;
		thread.updatedAt = new Date().toISOString();
		this.write(workspaceId, docPath, data);
		return thread;
	}

	deleteThread(workspaceId: string, docPath: string, threadId: string): boolean {
		const data = this.read(workspaceId, docPath);
		const next = data.threads.filter((t) => t.id !== threadId);
		if (next.length === data.threads.length) return false;
		this.write(workspaceId, docPath, { docPath, threads: next });
		return true;
	}
}
