import { useEffect, useMemo, useRef, useState, type KeyboardEvent } from "react";
import { createPortal } from "react-dom";
import { BookOpen, File, Folder, LoaderCircle, MessageSquareText, Search, X } from "lucide-react";
import { useTranslation } from "react-i18next";
import type { SearchCategory, SearchResult, SearchResultCategory } from "../../../src/shared/search.js";
import { searchInnoAgent } from "../api/search.js";

const SEARCH_CATEGORIES: SearchCategory[] = ["all", "conversation", "workspace", "file", "knowledge"];

function resultIcon(category: SearchResultCategory) {
	switch (category) {
		case "conversation": return <MessageSquareText size={15} />;
		case "workspace": return <Folder size={15} />;
		case "file": return <File size={15} />;
		case "knowledge": return <BookOpen size={15} />;
	}
}

function useSearchResults(query: string, category: SearchCategory) {
	const { t } = useTranslation();
	const [results, setResults] = useState<SearchResult[]>([]);
	const [loading, setLoading] = useState(false);
	const [error, setError] = useState("");
	useEffect(() => {
		const trimmed = query.trim();
		if (!trimmed) {
			setResults([]);
			setLoading(false);
			setError("");
			return;
		}
		setResults([]);
		setLoading(true);
		setError("");
		const controller = new AbortController();
		const timer = window.setTimeout(() => {
			void searchInnoAgent(trimmed, category, controller.signal)
				.then((response) => setResults(response.results))
				.catch((reason: unknown) => {
					if (!controller.signal.aborted) {
						setResults([]);
						setError(reason instanceof Error ? reason.message : t("globalSearch.failed"));
					}
				})
				.finally(() => {
					if (!controller.signal.aborted) setLoading(false);
				});
		}, 180);
		return () => {
			window.clearTimeout(timer);
			controller.abort();
		};
	}, [query, category, t]);
	return { results, loading, error };
}

interface SearchPanelProps {
	standalone?: boolean;
	macTitlebarInset?: boolean;
	onClose: () => void;
	onNavigate: (result: SearchResult) => void | Promise<void>;
}

export function SearchPanel({ standalone = false, macTitlebarInset = false, onClose, onNavigate }: SearchPanelProps) {
	const { t } = useTranslation();
	const [query, setQuery] = useState("");
	const [category, setCategory] = useState<SearchCategory>("all");
	const [activeIndex, setActiveIndex] = useState(0);
	const inputRef = useRef<HTMLInputElement>(null);
	const listRef = useRef<HTMLDivElement>(null);
	const { results, loading, error } = useSearchResults(query, category);
	const categoryNames = useMemo<Record<SearchCategory, string>>(() => ({
		all: t("globalSearch.tabs.all"),
		conversation: t("globalSearch.tabs.conversation"),
		workspace: t("globalSearch.tabs.workspace"),
		file: t("globalSearch.tabs.file"),
		knowledge: t("globalSearch.tabs.knowledge"),
	}), [t]);
	const resultGroups = useMemo(() => {
		if (category !== "all") return [{ category, results }];
		return (["conversation", "workspace", "file", "knowledge"] as const)
			.map((item) => ({ category: item, results: results.filter((result) => result.category === item) }))
			.filter((group) => group.results.length > 0);
	}, [category, results]);
	const orderedResults = useMemo(() => resultGroups.flatMap((group) => group.results), [resultGroups]);

	useEffect(() => {
		inputRef.current?.focus();
	}, []);
	useEffect(() => setActiveIndex(0), [query, category, results.length]);
	useEffect(() => {
		const active = listRef.current?.querySelector<HTMLElement>("[data-active='true']");
		active?.scrollIntoView({ block: "nearest" });
	}, [activeIndex]);

	const selectCategory = (delta: number) => {
		const currentIndex = SEARCH_CATEGORIES.indexOf(category);
		const nextIndex = (currentIndex + delta + SEARCH_CATEGORIES.length) % SEARCH_CATEGORIES.length;
		setCategory(SEARCH_CATEGORIES[nextIndex]);
		inputRef.current?.focus();
	};

	const onKeyDown = (event: KeyboardEvent<HTMLDivElement>) => {
		if (event.key === "Escape") {
			event.preventDefault();
			event.stopPropagation();
			onClose();
			return;
		}
		if ((event.ctrlKey || event.metaKey) && event.key === "ArrowLeft") {
			event.preventDefault();
			selectCategory(-1);
			return;
		}
		if ((event.ctrlKey || event.metaKey) && event.key === "ArrowRight") {
			event.preventDefault();
			selectCategory(1);
			return;
		}
		if (event.key === "ArrowLeft" || event.key === "ArrowRight") {
			const input = event.target instanceof HTMLInputElement ? event.target : null;
			const atEdge = !input || !query || (event.key === "ArrowLeft"
				? input.selectionStart === 0 && input.selectionEnd === 0
				: input.selectionStart === input.value.length && input.selectionEnd === input.value.length);
			if (atEdge) {
				event.preventDefault();
				selectCategory(event.key === "ArrowLeft" ? -1 : 1);
				return;
			}
		}
		if (event.key === "ArrowDown" && orderedResults.length > 0) {
			event.preventDefault();
			setActiveIndex((index) => (index + 1) % orderedResults.length);
			return;
		}
		if (event.key === "ArrowUp" && orderedResults.length > 0) {
			event.preventDefault();
			setActiveIndex((index) => (index - 1 + orderedResults.length) % orderedResults.length);
			return;
		}
		if (event.key === "Enter" && orderedResults[activeIndex]) {
			event.preventDefault();
			void onNavigate(orderedResults[activeIndex]);
		}
	};

	let resultIndex = -1;
	const surface = (
		<div
			className={`inno-search-surface${standalone ? " inno-search-surface--window" : " inno-search-surface--overlay"}`}
			role="dialog"
			aria-modal="true"
			aria-label={t("globalSearch.open")}
			onKeyDown={onKeyDown}
		>
			<div className="inno-search-header">
				<div className="inno-search-input-wrap">
					<Search size={17} aria-hidden="true" />
					<input
						ref={inputRef}
						className="inno-search-input"
						value={query}
						onChange={(event) => setQuery(event.target.value)}
						placeholder={t("globalSearch.placeholder")}
						aria-label={t("globalSearch.placeholder")}
						autoComplete="off"
						spellCheck={false}
					/>
				</div>
				{!standalone ? (
					<button className="inno-search-close" type="button" aria-label={t("common.close")} onClick={onClose}>
						<X size={16} />
					</button>
				) : null}
			</div>

			<div className="inno-search-tabs" role="tablist" aria-label={t("globalSearch.open")}>
				{SEARCH_CATEGORIES.map((item) => (
					<button
						key={item}
						type="button"
						role="tab"
						aria-selected={category === item}
						className={`inno-search-tab${category === item ? " is-active" : ""}`}
						onClick={() => { setCategory(item); inputRef.current?.focus(); }}
					>
						{categoryNames[item]}
					</button>
				))}
			</div>

			<div ref={listRef} className="inno-search-results" role="listbox" aria-label={categoryNames[category]}>
				{!query.trim() ? (
					<div className="inno-search-empty">
						<Search size={48} strokeWidth={1.7} aria-hidden="true" />
						<p>{t("globalSearch.emptyPrompt")}</p>
					</div>
				) : loading && results.length === 0 ? (
					<div className="inno-search-status"><LoaderCircle className="animate-spin" size={18} />{t("globalSearch.loading")}</div>
				) : error ? (
					<div className="inno-search-status" role="alert">{error}</div>
				) : results.length === 0 ? (
					<div className="inno-search-status">{t("globalSearch.noResults")}</div>
				) : resultGroups.map((group) => (
					<section key={group.category} className="inno-search-group" aria-label={categoryNames[group.category]}>
						{category === "all" ? <h2 className="inno-search-group-title">{categoryNames[group.category]}</h2> : null}
						{group.results.map((result) => {
							resultIndex += 1;
							const index = resultIndex;
							const active = index === activeIndex;
							return (
								<button
									key={result.id}
									className="inno-search-result"
									type="button"
									role="option"
									aria-selected={active}
									data-active={active}
									onMouseEnter={() => setActiveIndex(index)}
									onClick={() => void onNavigate(result)}
								>
									<span className="inno-search-result-icon" aria-hidden="true">{resultIcon(result.category)}</span>
									<span className="inno-search-result-copy">
										<span className="inno-search-result-title">
											{result.title}
											{result.archived ? <span className="inno-search-result-badge">{t("globalSearch.archived")}</span> : null}
										</span>
										{result.subtitle ? <span className="inno-search-result-subtitle">{result.subtitle}</span> : null}
										{result.snippet ? <span className="inno-search-result-snippet">{result.snippet}</span> : null}
									</span>
								</button>
							);
						})}
					</section>
				))}
			</div>

			<footer className="inno-search-footer">
				<span>{t("globalSearch.shortcuts.move")}</span>
				<span>{t("globalSearch.shortcuts.open")}</span>
				<span>{t("globalSearch.shortcuts.switch")}</span>
				<span>{t("globalSearch.shortcuts.close")}</span>
			</footer>
		</div>
	);

	if (standalone) return (
		<div className={`inno-search-window-root${macTitlebarInset ? " inno-search-window-root--mac-titlebar" : ""}`}>
			{macTitlebarInset ? <div className="inno-search-window-drag-region" aria-hidden="true" /> : null}
			{surface}
		</div>
	);
	return createPortal(
		<div className="inno-search-overlay" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose(); }}>
			{surface}
		</div>,
		document.body,
	);
}
