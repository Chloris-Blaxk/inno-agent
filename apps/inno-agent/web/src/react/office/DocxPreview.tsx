import { useCallback, useEffect, useRef, useState } from "react";
import { useTranslation } from "react-i18next";
import { Check, Download, MessageSquarePlus, RotateCcw, Trash2, X } from "lucide-react";
import { renderAsync } from "docx-preview";
import type { WorkspaceFileDetail } from "../../types/workspace.js";
import { triggerDownload } from "../../api/workspace.js";
import {
	askDocComment,
	listDocComments,
	updateDocComment,
	type DocCommentAnchor,
	type DocCommentThread,
} from "../../api/comments.js";

/**
 * Render a .docx into an HTML container with formatting preserved, via
 * docx-preview (client-side, no LibreOffice). Options mirror pi-web-ui's
 * DocxArtifact for parity.
 *
 * Anchored comments (ChatGPT-canvas-style): selecting text inside the rendered
 * document offers a comment popover; threads are stored server-side keyed by
 * workspace + document path, anchored by content (text + prefix/suffix) rather
 * than offsets so they survive unrelated edits.
 */

const ANCHOR_CONTEXT_CHARS = 120;

interface PendingAnchor {
	anchor: DocCommentAnchor;
	/** Popover position, relative to the outer wrapper. */
	top: number;
	left: number;
}

function workspaceIdFromUrl(url?: string): string | undefined {
	if (!url) return undefined;
	try {
		const params = new URL(url, window.location.origin).searchParams;
		return params.get("workspaceId") ?? undefined;
	} catch {
		return undefined;
	}
}

/** Exact surrounding context of a Range within the rendered container. */
function anchorFromRange(range: Range, container: HTMLElement): DocCommentAnchor | null {
	const text = range.toString();
	if (!text.trim() || text.length > 2000) return null;
	const before = document.createRange();
	before.selectNodeContents(container);
	before.setEnd(range.startContainer, range.startOffset);
	const after = document.createRange();
	after.selectNodeContents(container);
	after.setStart(range.endContainer, range.endOffset);
	return {
		text,
		prefix: before.toString().slice(-ANCHOR_CONTEXT_CHARS),
		suffix: after.toString().slice(0, ANCHOR_CONTEXT_CHARS),
	};
}

/** Wrap the best-matching occurrence of the anchor text in <mark> elements. */
function highlightAnchor(container: HTMLElement, anchor: DocCommentAnchor, className: string): HTMLElement[] {
	const walker = document.createTreeWalker(container, NodeFilter.SHOW_TEXT);
	const nodes: Text[] = [];
	let node = walker.nextNode();
	while (node) {
		nodes.push(node as Text);
		node = walker.nextNode();
	}
	const fullText = nodes.map((n) => n.data).join("");
	const haystack = anchor.text;
	const candidates: number[] = [];
	let idx = fullText.indexOf(haystack);
	while (idx >= 0) {
		candidates.push(idx);
		idx = fullText.indexOf(haystack, idx + 1);
	}
	if (!candidates.length) return [];
	// Prefer the occurrence whose surrounding context matches the anchor.
	const score = (at: number): number => {
		let s = 0;
		const pre = fullText.slice(Math.max(0, at - anchor.prefix.length), at);
		if (anchor.prefix && pre === anchor.prefix) s += 2;
		else if (anchor.prefix && pre.endsWith(anchor.prefix.slice(-20))) s += 1;
		const post = fullText.slice(at + haystack.length, at + haystack.length + anchor.suffix.length);
		if (anchor.suffix && post === anchor.suffix) s += 2;
		else if (anchor.suffix && post.startsWith(anchor.suffix.slice(0, 20))) s += 1;
		return s;
	};
	const start = candidates.reduce((best, at) => (score(at) > score(best) ? at : best), candidates[0]);
	const end = start + haystack.length;

	// Map character offsets back to text nodes and wrap the covered ranges.
	const marks: HTMLElement[] = [];
	let offset = 0;
	for (const textNode of nodes) {
		const nodeStart = offset;
		const nodeEnd = offset + textNode.data.length;
		offset = nodeEnd;
		const wrapStart = Math.max(start, nodeStart);
		const wrapEnd = Math.min(end, nodeEnd);
		if (wrapStart >= wrapEnd) continue;
		const target = textNode.splitText(wrapStart - nodeStart);
		if (wrapEnd - nodeStart < target.data.length) target.splitText(wrapEnd - wrapStart);
		const mark = document.createElement("mark");
		mark.className = className;
		target.parentNode?.replaceChild(mark, target);
		mark.appendChild(target);
		marks.push(mark);
	}
	return marks;
}

function clearHighlights(container: HTMLElement, className: string): void {
	for (const mark of Array.from(container.querySelectorAll(`mark.${className}`))) {
		const parent = mark.parentNode;
		if (!parent) continue;
		while (mark.firstChild) parent.insertBefore(mark.firstChild, mark);
		parent.removeChild(mark);
		parent.normalize();
	}
}

export default function DocxPreview({ file }: { file: WorkspaceFileDetail }) {
	const { t } = useTranslation();
	const wrapperRef = useRef<HTMLDivElement>(null);
	const containerRef = useRef<HTMLDivElement>(null);
	const [loading, setLoading] = useState(true);
	const [error, setError] = useState("");

	const wsId = workspaceIdFromUrl(file.url);
	const [threads, setThreads] = useState<DocCommentThread[]>([]);
	const [pending, setPending] = useState<PendingAnchor | null>(null);
	const [activeThreadId, setActiveThreadId] = useState<string | null>(null);
	const [panelOpen, setPanelOpen] = useState(false);
	const [draft, setDraft] = useState("");
	const [submitting, setSubmitting] = useState(false);
	const [commentError, setCommentError] = useState("");

	useEffect(() => {
		if (!file.url) {
			setError(t("preview.docxFailed", "Failed to render document"));
			setLoading(false);
			return;
		}
		let cancelled = false;
		setLoading(true);
		setError("");
		setThreads([]);
		setPending(null);
		setActiveThreadId(null);
		const container = containerRef.current;
		if (container) container.innerHTML = "";
		(async () => {
			try {
				const res = await fetch(file.url as string);
				if (!res.ok) throw new Error(res.statusText);
				const buf = await res.arrayBuffer();
				if (cancelled || !containerRef.current) return;
				await renderAsync(buf, containerRef.current, undefined, {
					className: "docx",
					inWrapper: true,
					ignoreWidth: true,
					ignoreHeight: false,
					ignoreFonts: false,
					breakPages: true,
					ignoreLastRenderedPageBreak: true,
					experimental: false,
					trimXmlDeclaration: true,
					useBase64URL: false,
					renderHeaders: true,
					renderFooters: true,
					renderFootnotes: true,
					renderEndnotes: true,
				});
				// Load existing comment threads and re-anchor their highlights.
				const loaded = await listDocComments(file.path, wsId).catch(() => [] as DocCommentThread[]);
				if (cancelled || !containerRef.current) return;
				setThreads(loaded);
				if (loaded.length) setPanelOpen(true);
				for (const thread of loaded) {
					if (thread.status === "open") highlightAnchor(containerRef.current, thread.anchor, "inno-comment-anchor");
				}
			} catch (e) {
				if (!cancelled) setError(e instanceof Error ? e.message : t("preview.docxFailed", "Failed to render document"));
			} finally {
				if (!cancelled) setLoading(false);
			}
		})();
		return () => { cancelled = true; };
	}, [file.url, file.path, t]); // eslint-disable-line react-hooks/exhaustive-deps

	const onMouseUp = useCallback(() => {
		const container = containerRef.current;
		const wrapper = wrapperRef.current;
		if (!container || !wrapper) return;
		const selection = window.getSelection();
		if (!selection || selection.isCollapsed || selection.rangeCount === 0) {
			setPending(null);
			return;
		}
		const range = selection.getRangeAt(0);
		if (!container.contains(range.commonAncestorContainer)) {
			setPending(null);
			return;
		}
		const anchor = anchorFromRange(range, container);
		if (!anchor) {
			setPending(null);
			return;
		}
		const rect = range.getBoundingClientRect();
		const wrapRect = wrapper.getBoundingClientRect();
		setPending({
			anchor,
			top: rect.top - wrapRect.top - 8,
			left: Math.max(8, rect.left - wrapRect.left + rect.width / 2),
		});
	}, []);

	const startComment = () => {
		if (!pending) return;
		setActiveThreadId(null);
		setDraft("");
		setCommentError("");
		setPanelOpen(true);
		window.getSelection()?.removeAllRanges();
	};

	const submitComment = async () => {
		const question = draft.trim();
		if (!question || submitting) return;
		const activeThread = threads.find((th) => th.id === activeThreadId);
		const anchor = activeThread ? undefined : pending?.anchor;
		if (!activeThread && !anchor) return;
		setSubmitting(true);
		setCommentError("");
		try {
			const thread = await askDocComment({
				path: file.path,
				workspaceId: wsId,
				question,
				...(activeThread ? { threadId: activeThread.id } : { anchor: anchor as DocCommentAnchor }),
			});
			setThreads((prev) => {
				const idx = prev.findIndex((th) => th.id === thread.id);
				if (idx >= 0) {
					const next = [...prev];
					next[idx] = thread;
					return next;
				}
				return [...prev, thread];
			});
			setActiveThreadId(thread.id);
			if (!activeThread && containerRef.current) {
				highlightAnchor(containerRef.current, thread.anchor, "inno-comment-anchor");
			}
			setPending(null);
			setDraft("");
		} catch (e) {
			setCommentError(e instanceof Error ? e.message : t("preview.commentFailed", "Comment answer failed, please retry"));
		} finally {
			setSubmitting(false);
		}
	};

	const mutateThread = async (threadId: string, action: "resolve" | "reopen" | "delete") => {
		try {
			await updateDocComment({ path: file.path, workspaceId: wsId, threadId, action });
		} catch {
			return;
		}
		if (action === "delete") {
			const removed = threads.find((th) => th.id === threadId);
			setThreads((prev) => prev.filter((th) => th.id !== threadId));
			if (removed && containerRef.current) {
				// Re-anchor everything: simplest consistent way to drop one highlight.
				clearHighlights(containerRef.current, "inno-comment-anchor");
				for (const other of threads.filter((item) => item.id !== threadId && item.status === "open")) {
					highlightAnchor(containerRef.current, other.anchor, "inno-comment-anchor");
				}
			}
			if (activeThreadId === threadId) setActiveThreadId(null);
			return;
		}
		setThreads((prev) => prev.map((th) => (th.id === threadId ? { ...th, status: action === "resolve" ? "resolved" : "open" } : th)));
	};

	const jumpToThread = (thread: DocCommentThread) => {
		const nextActive = thread.id === activeThreadId ? null : thread.id;
		setActiveThreadId(nextActive);
		const container = containerRef.current;
		if (!container) return;
		clearHighlights(container, "inno-comment-anchor-active");
		if (!nextActive) return;
		const marks = highlightAnchor(container, thread.anchor, "inno-comment-anchor-active");
		marks[0]?.scrollIntoView({ block: "center", behavior: "smooth" });
	};

	const downloadOriginal = () => {
		if (file.url) triggerDownload(`${file.url}${file.url.includes("?") ? "&" : "?"}download=1`);
	};

	const activeThread = threads.find((th) => th.id === activeThreadId) ?? null;
	const composerAnchor = activeThread?.anchor ?? pending?.anchor ?? null;

	return (
		<div ref={wrapperRef} className="relative flex h-full" onMouseUp={onMouseUp}>
			<div className="workspace-scroll h-full flex-1 overflow-auto bg-[var(--inno-surface-muted)]">
				{loading ? (
					<div className="flex h-full items-center justify-center text-sm text-[var(--inno-text-muted)]">
						{t("preview.docxLoading", "Rendering document...")}
					</div>
				) : null}
				{error ? (
					<div className="flex h-full flex-col items-center justify-center gap-3 p-6 text-center text-sm text-[var(--inno-text-muted)]">
						<div className="font-medium text-[var(--inno-text)]">{file.name}</div>
						<div className="text-xs text-[var(--inno-danger)]">{error}</div>
						<button
							className="flex items-center gap-1 rounded-md border border-[var(--inno-border)] px-3 py-1.5 text-xs text-[var(--inno-text-muted)] hover:bg-[var(--inno-surface)]"
							onClick={downloadOriginal}
						>
							<Download size={12} />
							{t("files.download", "Download")}
						</button>
					</div>
				) : null}
				<div ref={containerRef} className={`docx-host px-4 py-4 ${loading || error ? "hidden" : ""}`} />
			</div>

			{pending && !submitting ? (
				<button
					className="absolute z-20 flex -translate-x-1/2 -translate-y-full items-center gap-1 rounded-full border border-[var(--inno-border)] bg-[var(--inno-surface)] px-2.5 py-1 text-xs text-[var(--inno-text)] shadow-md hover:bg-[var(--inno-surface-muted)]"
					style={{ top: pending.top, left: pending.left }}
					onMouseDown={(e) => e.preventDefault()}
					onClick={startComment}
				>
					<MessageSquarePlus size={12} />
					{t("preview.commentAdd", "Comment")}
				</button>
			) : null}

			{panelOpen ? (
				<div className="flex w-80 shrink-0 flex-col border-l border-[var(--inno-border)] bg-[var(--inno-surface)]">
					<div className="flex items-center justify-between border-b border-[var(--inno-border)] px-3 py-2">
						<span className="text-xs font-medium text-[var(--inno-text)]">{t("preview.commentTitle", "Comments")}</span>
						<button
							className="text-[var(--inno-text-muted)] hover:text-[var(--inno-text)]"
							onClick={() => { setPanelOpen(false); setActiveThreadId(null); }}
						>
							<X size={14} />
						</button>
					</div>
					<div className="flex-1 overflow-auto p-2">
						{threads.length === 0 && !pending ? (
							<div className="p-3 text-xs text-[var(--inno-text-muted)]">
								{t("preview.commentEmpty", "Select a passage in the document to start a comment thread")}
							</div>
						) : null}
						{threads.map((thread) => (
							<div
								key={thread.id}
								className={`mb-2 cursor-pointer rounded-md border p-2 text-xs ${
									thread.id === activeThreadId
										? "border-[var(--inno-accent)] bg-[var(--inno-surface-muted)]"
										: "border-[var(--inno-border)] hover:bg-[var(--inno-surface-muted)]"
								} ${thread.status === "resolved" ? "opacity-60" : ""}`}
								onClick={() => jumpToThread(thread)}
							>
								<div className="mb-1 line-clamp-2 border-l-2 border-[var(--inno-accent)] pl-2 italic text-[var(--inno-text-muted)]">
									{thread.anchor.text}
								</div>
								{thread.turns.map((turn, i) => (
									<div key={i} className="mb-1">
										<span className="font-medium text-[var(--inno-text)]">{turn.role === "user" ? "Q: " : "A: "}</span>
										<span className="whitespace-pre-wrap text-[var(--inno-text)]">{turn.content}</span>
									</div>
								))}
								<div className="mt-1 flex items-center gap-2" onClick={(e) => e.stopPropagation()}>
									{thread.status === "open" ? (
										<button
											className="flex items-center gap-0.5 text-[var(--inno-text-muted)] hover:text-[var(--inno-text)]"
											onClick={() => mutateThread(thread.id, "resolve")}
										>
											<Check size={11} />
											{t("preview.commentResolve", "Resolve")}
										</button>
									) : (
										<button
											className="flex items-center gap-0.5 text-[var(--inno-text-muted)] hover:text-[var(--inno-text)]"
											onClick={() => mutateThread(thread.id, "reopen")}
										>
											<RotateCcw size={11} />
											{t("preview.commentReopen", "Reopen")}
										</button>
									)}
									<button
										className="flex items-center gap-0.5 text-[var(--inno-text-muted)] hover:text-[var(--inno-danger)]"
										onClick={() => mutateThread(thread.id, "delete")}
									>
										<Trash2 size={11} />
										{t("preview.commentDelete", "Delete")}
									</button>
									{thread.status === "resolved" ? (
										<span className="ml-auto text-[var(--inno-text-muted)]">{t("preview.commentResolved", "Resolved")}</span>
									) : null}
								</div>
							</div>
						))}
					</div>
					{composerAnchor ? (
						<div className="border-t border-[var(--inno-border)] p-2">
							<div className="mb-1 line-clamp-2 border-l-2 border-[var(--inno-accent)] pl-2 text-xs italic text-[var(--inno-text-muted)]">
								{composerAnchor.text}
							</div>
							<textarea
								className="mb-1 w-full resize-none rounded-md border border-[var(--inno-border)] bg-[var(--inno-surface)] p-2 text-xs text-[var(--inno-text)] outline-none focus:border-[var(--inno-accent)]"
								rows={3}
								placeholder={activeThread ? t("preview.commentReplyPlaceholder", "Ask a follow-up…") : t("preview.commentPlaceholder", "Ask about this passage…")}
								value={draft}
								onChange={(e) => setDraft(e.target.value)}
								onKeyDown={(e) => {
									if (e.key === "Enter" && (e.metaKey || e.ctrlKey)) void submitComment();
								}}
							/>
							{commentError ? <div className="mb-1 text-xs text-[var(--inno-danger)]">{commentError}</div> : null}
							<button
								className="w-full rounded-md bg-[var(--inno-accent)] px-2 py-1 text-xs text-white disabled:opacity-50"
								disabled={!draft.trim() || submitting}
								onClick={() => void submitComment()}
							>
								{submitting ? t("preview.commentSending", "Thinking…") : t("preview.commentSend", "Ask")}
							</button>
						</div>
					) : null}
				</div>
			) : null}
		</div>
	);
}
