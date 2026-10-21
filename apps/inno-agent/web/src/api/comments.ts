import { apiFetch } from "./client.js";

/** Client for /api/comments — anchored comment threads on previewed documents. */

export interface DocCommentAnchor {
	text: string;
	prefix: string;
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

export async function listDocComments(path: string, workspaceId?: string): Promise<DocCommentThread[]> {
	const params = new URLSearchParams({ path });
	if (workspaceId) params.set("workspaceId", workspaceId);
	const data = await apiFetch<{ threads: DocCommentThread[] }>(`/api/comments?${params.toString()}`);
	return data.threads;
}

export async function askDocComment(input: {
	path: string;
	workspaceId?: string;
	question: string;
	anchor?: DocCommentAnchor;
	threadId?: string;
}): Promise<DocCommentThread> {
	const data = await apiFetch<{ thread: DocCommentThread }>("/api/comments/ask", {
		method: "POST",
		body: JSON.stringify(input),
	});
	return data.thread;
}

export async function updateDocComment(input: {
	path: string;
	workspaceId?: string;
	threadId: string;
	action: "resolve" | "reopen" | "delete";
}): Promise<void> {
	await apiFetch("/api/comments/update", {
		method: "POST",
		body: JSON.stringify(input),
	});
}
