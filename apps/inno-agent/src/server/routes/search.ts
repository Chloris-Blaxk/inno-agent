import type { IncomingMessage as HttpReq, ServerResponse } from "node:http";
import { existsSync, readdirSync, realpathSync, statSync } from "node:fs";
import { basename, join, relative } from "node:path";
import type { RuntimePaths } from "../../runtime.js";
import { getL2Memory } from "../../memory/l2/l2-memory.js";
import { parseFrontmatter } from "../../memory/l2/wiki-maintainer.js";
import { listWikiPagePaths } from "../../memory/l2/wiki-page-files.js";
import { getL3Memory } from "../../memory/l3/l3-tools.js";
import { readJson, readText } from "../../storage/file-store.js";
import type { SessionMessageSummary, SessionSummary } from "../session-model.js";
import { isWithin } from "../../utils/path-safety.js";
import { WORKSPACE_IGNORES, WORKSPACE_TREE_MAX_DEPTH } from "../file-helpers.js";
import type { SearchCategory, SearchResponse, SearchResult } from "../../shared/search.js";
import type { WorkspaceRegistry } from "../../workspace/workspace-registry.js";
import { json } from "../http-helpers.js";

const CATEGORY_ORDER = ["conversation", "workspace", "file", "knowledge"] as const;
const RESULTS_PER_CATEGORY = 12;

type ParsedSession = { summary: SessionSummary; messages: SessionMessageSummary[] };

export interface SearchRouteContext {
	dataDir: string;
	paths: RuntimePaths;
	workspaceRegistry: WorkspaceRegistry;
	parseSessionFile: (filePath: string) => ParsedSession | null;
}

function normalized(value: string): string {
	return value.toLocaleLowerCase();
}

function excerpt(text: string, query: string, maxLength = 180): string {
	const content = text.replace(/\s+/g, " ").trim();
	if (content.length <= maxLength) return content;
	const lower = normalized(content);
	const terms = query.trim().split(/\s+/).filter(Boolean);
	const index = terms.map((term) => lower.indexOf(normalized(term))).filter((at) => at >= 0).sort((a, b) => a - b)[0] ?? -1;
	if (index < 0) return `${content.slice(0, maxLength).trimEnd()}…`;
	const start = Math.max(0, index - Math.floor(maxLength / 3));
	const end = Math.min(content.length, start + maxLength);
	return `${start > 0 ? "…" : ""}${content.slice(start, end).trim()}${end < content.length ? "…" : ""}`;
}

function includesQuery(value: string, query: string): boolean {
	return normalized(value).includes(normalized(query));
}

function rankText(value: string, query: string): number {
	const target = normalized(value);
	const needle = normalized(query);
	if (target === needle) return 3;
	if (target.startsWith(needle)) return 2;
	return target.includes(needle) ? 1 : 0;
}

function searchWorkspaces(ctx: SearchRouteContext, query: string): SearchResult[] {
	return ctx.workspaceRegistry.listWorkspaces()
		.filter((workspace) => includesQuery(workspace.name, query))
		.sort((a, b) => rankText(b.name, query) - rankText(a.name, query) || Date.parse(b.updatedAt) - Date.parse(a.updatedAt))
		.slice(0, RESULTS_PER_CATEGORY)
		.map((workspace) => ({
			id: `workspace:${workspace.id}`,
			category: "workspace",
			title: workspace.name,
			subtitle: workspace.relPath,
			snippet: workspace.isTemp ? "临时工作区" : workspace.relPath,
			updatedAt: workspace.updatedAt,
			workspaceId: workspace.id,
		}));
}

function searchWorkspaceFiles(ctx: SearchRouteContext, query: string): SearchResult[] {
	const matches: Array<SearchResult & { rank: number }> = [];
	for (const workspace of ctx.workspaceRegistry.listWorkspaces()) {
		const root = ctx.workspaceRegistry.resolveWorkspaceDir(workspace.id);
		if (!root || !existsSync(root)) continue;
		let realRoot: string;
		try {
			realRoot = realpathSync(root);
		} catch {
			continue;
		}
		const visited = new Set<string>([realRoot]);
		const walk = (dir: string, depth: number): void => {
			if (depth > WORKSPACE_TREE_MAX_DEPTH) return;
			let entries;
			try {
				entries = readdirSync(dir, { withFileTypes: true });
			} catch {
				return;
			}
			for (const entry of entries) {
				if (WORKSPACE_IGNORES.has(entry.name)) continue;
				const fullPath = join(dir, entry.name);
				let stat: ReturnType<typeof statSync>;
				let realPath: string;
				try {
					stat = statSync(fullPath);
					realPath = realpathSync(fullPath);
				} catch {
					continue;
				}
				if (!isWithin(realRoot, realPath)) continue;
				if (stat.isDirectory()) {
					if (visited.has(realPath)) continue;
					visited.add(realPath);
					walk(fullPath, depth + 1);
					continue;
				}
				if (!stat.isFile()) continue;
				const path = relative(root, fullPath).replaceAll("\\", "/");
				if (!includesQuery(path, query)) continue;
				matches.push({
					id: `file:${workspace.id}:${path}`,
					category: "file",
					title: basename(path),
					subtitle: `${workspace.name} · ${path}`,
					snippet: path,
					updatedAt: stat.mtime.toISOString(),
					workspaceId: workspace.id,
					path,
					rank: rankText(basename(path), query),
				});
			}
		};
		walk(root, 0);
	}
	return matches
		.sort((a, b) => b.rank - a.rank || Date.parse(b.updatedAt ?? "") - Date.parse(a.updatedAt ?? ""))
		.slice(0, RESULTS_PER_CATEGORY)
		.map(({ rank: _rank, ...result }) => result);
}

function searchKnowledgeFallback(ctx: SearchRouteContext, query: string): SearchResult[] {
	const results: SearchResult[] = [];
	for (const path of listWikiPagePaths(ctx.paths.l2DataDir)) {
		try {
			const content = readText(join(ctx.paths.l2DataDir, path));
			const { frontmatter, body } = parseFrontmatter(content);
			const title = frontmatter?.title?.trim() || basename(path, ".md");
			const tags = frontmatter?.tags ?? [];
			if (!includesQuery(`${title} ${tags.join(" ")} ${body}`, query)) continue;
			results.push({
				id: `knowledge:${path}`,
				category: "knowledge",
				title,
				subtitle: path,
				snippet: excerpt(body || title, query),
				path,
			});
		} catch {
			// Ignore an unavailable or concurrently removed page.
		}
		if (results.length >= RESULTS_PER_CATEGORY) break;
	}
	return results;
}

async function searchKnowledge(ctx: SearchRouteContext, query: string): Promise<SearchResult[]> {
	const memory = getL2Memory(ctx.paths.l2DataDir);
	await memory.backfill();
	const store = await memory.getStore();
	if (!store) return searchKnowledgeFallback(ctx, query);
	const pages = new Map(store.getAllPages().map((page) => [page.path, page]));
	return store.searchLexical(query, RESULTS_PER_CATEGORY * 2)
		.flatMap((hit) => {
			const page = pages.get(hit.path);
			if (!page) return [];
			return [{
				id: `knowledge:${page.path}`,
				category: "knowledge" as const,
				title: page.title,
				subtitle: page.path,
				snippet: excerpt(page.body || page.title, query),
				path: page.path,
			}];
		})
		.slice(0, RESULTS_PER_CATEGORY);
}

function activeMessageForHit(messages: SessionMessageSummary[], hit: { role: "user" | "assistant"; entryId?: string; ts: number; text: string }): SessionMessageSummary | undefined {
	if (hit.entryId) return messages.find((message) => message.entryId === hit.entryId && message.role === hit.role);
	return messages.find((message) => message.role === hit.role
		&& Math.abs(message.timestamp - hit.ts) < 2000
		&& message.content.includes(hit.text));
}

function conversationFallback(ctx: SearchRouteContext, query: string): SearchResult[] {
	const sessionDir = join(ctx.dataDir, "sessions");
	const archived = readJson<Record<string, boolean>>(join(sessionDir, "archives.json"), {});
	let files: string[];
	try {
		files = readdirSync(sessionDir).filter((file) => file.endsWith(".jsonl"));
	} catch {
		return [];
	}
	const results: SearchResult[] = [];
	for (const file of files) {
		const parsed = ctx.parseSessionFile(join(sessionDir, file));
		if (!parsed) continue;
		for (const message of parsed.messages) {
			if (!includesQuery(message.content, query)) continue;
			results.push({
				id: `conversation:${parsed.summary.id}:${message.entryId ?? message.timestamp}`,
				category: "conversation",
				title: parsed.summary.name,
				subtitle: new Date(message.timestamp).toISOString(),
				snippet: excerpt(message.content, query),
				updatedAt: parsed.summary.updatedAt,
				archived: archived[parsed.summary.id] === true,
				role: message.role,
				sessionId: parsed.summary.id,
				entryId: message.entryId,
			});
		}
	}
	return results
		.sort((a, b) => rankText(b.snippet, query) - rankText(a.snippet, query) || Date.parse(b.updatedAt ?? "") - Date.parse(a.updatedAt ?? ""))
		.slice(0, RESULTS_PER_CATEGORY);
}

async function searchConversations(ctx: SearchRouteContext, query: string): Promise<SearchResult[]> {
	const sessionDir = join(ctx.dataDir, "sessions");
	const memory = getL3Memory(ctx.paths.l3DataDir, sessionDir);
	const indexedHits = await memory.searchLexical(query, RESULTS_PER_CATEGORY * 8);
	if (!indexedHits) return conversationFallback(ctx, query);
	const archived = readJson<Record<string, boolean>>(join(sessionDir, "archives.json"), {});
	const parsedById = new Map<string, ParsedSession | null>();
	const results: SearchResult[] = [];
	const usedMessages = new Set<string>();
	for (const hit of indexedHits) {
		let parsed = parsedById.get(hit.sessionId);
		if (parsed === undefined) {
			parsed = ctx.parseSessionFile(join(sessionDir, hit.sessionId));
			parsedById.set(hit.sessionId, parsed);
		}
		if (!parsed) continue;
		const message = activeMessageForHit(parsed.messages, { role: hit.role, entryId: hit.entryId, ts: hit.ts, text: hit.text });
		if (!message) continue;
		const messageKey = `${parsed.summary.id}:${message.entryId ?? message.timestamp}`;
		if (usedMessages.has(messageKey)) continue;
		usedMessages.add(messageKey);
		results.push({
			id: `conversation:${messageKey}`,
			category: "conversation",
			title: parsed.summary.name,
			subtitle: new Date(message.timestamp).toISOString(),
			snippet: excerpt(message.content, query),
			updatedAt: parsed.summary.updatedAt,
			archived: archived[parsed.summary.id] === true,
			role: message.role,
			sessionId: parsed.summary.id,
			entryId: message.entryId,
		});
		if (results.length >= RESULTS_PER_CATEGORY) break;
	}
	return results;
}

async function searchCategory(ctx: SearchRouteContext, query: string, category: SearchCategory): Promise<SearchResult[]> {
	switch (category) {
		case "conversation": return searchConversations(ctx, query);
		case "workspace": return searchWorkspaces(ctx, query);
		case "file": return searchWorkspaceFiles(ctx, query);
		case "knowledge": return searchKnowledge(ctx, query);
		case "all": {
			const results = await Promise.all(CATEGORY_ORDER.map((item) => searchCategory(ctx, query, item)));
			return results.flat();
		}
	}
}

export async function handleSearchRoutes(
	_req: HttpReq,
	res: ServerResponse,
	method: string,
	url: string,
	ctx: SearchRouteContext,
): Promise<boolean> {
	if (method !== "GET" || url.split("?")[0] !== "/api/search") return false;
	const params = new URL(url, "http://localhost").searchParams;
	const query = (params.get("query") ?? "").trim().slice(0, 200);
	const requestedCategory = params.get("category") ?? "all";
	const category: SearchCategory = ["all", ...CATEGORY_ORDER].includes(requestedCategory as SearchCategory)
		? requestedCategory as SearchCategory
		: "all";
	if (!query) {
		json(res, 200, { results: [] } satisfies SearchResponse);
		return true;
	}
	try {
		json(res, 200, { results: await searchCategory(ctx, query, category) } satisfies SearchResponse);
	} catch (err) {
		json(res, 500, { error: err instanceof Error ? err.message : "Search failed" });
	}
	return true;
}
