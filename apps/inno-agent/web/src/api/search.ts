import type { SearchCategory, SearchResponse } from "../../../src/shared/search.js";
import { apiFetch } from "./client.js";

export async function searchInnoAgent(
	query: string,
	category: SearchCategory,
	signal?: AbortSignal,
): Promise<SearchResponse> {
	const params = new URLSearchParams({ query, category });
	return apiFetch<SearchResponse>(`/api/search?${params.toString()}`, { signal });
}
