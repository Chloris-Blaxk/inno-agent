type HistoryMode = "push" | "replace" | "none";

interface NavigationPosition {
	index: number;
	maxIndex: number;
}

export interface NavigationAvailability {
	canGoBack: boolean;
	canGoForward: boolean;
}

const STATE_KEY = "__innoAgentNavigation";
export const NAVIGATION_HISTORY_CHANGE_EVENT = "inno-agent-navigation-history-change";

let knownMaxIndex = 0;

function getHistoryState(): Record<string, unknown> {
	const state: unknown = window.history.state;
	return state && typeof state === "object" && !Array.isArray(state)
		? state as Record<string, unknown>
		: {};
}

function getPosition(): NavigationPosition | null {
	const position = getHistoryState()[STATE_KEY];
	if (!position || typeof position !== "object") return null;
	const { index, maxIndex } = position as Partial<NavigationPosition>;
	if (!Number.isInteger(index) || !Number.isInteger(maxIndex)) return null;
	return { index: index as number, maxIndex: maxIndex as number };
}

function savePosition(position: NavigationPosition, mode: "push" | "replace", url: URL | string): void {
	const state = { ...getHistoryState(), [STATE_KEY]: position };
	if (mode === "push") window.history.pushState(state, "", url);
	else window.history.replaceState(state, "", url);
	window.dispatchEvent(new Event(NAVIGATION_HISTORY_CHANGE_EVENT));
}

/** Mark the current page as the start of the app's in-window navigation trail. */
export function initializeNavigationHistory(): void {
	if (typeof window === "undefined") return;
	const position = getPosition();
	if (position) {
		knownMaxIndex = Math.max(knownMaxIndex, position.maxIndex);
		if (position.maxIndex !== knownMaxIndex) {
			savePosition({ ...position, maxIndex: knownMaxIndex }, "replace", window.location.href);
		}
		return;
	}
	knownMaxIndex = 0;
	savePosition({ index: 0, maxIndex: 0 }, "replace", window.location.href);
}

/** Keep page and session URL updates on one tracked browser-history timeline. */
export function syncNavigationHistory(url: URL, mode: HistoryMode): void {
	if (mode === "none" || typeof window === "undefined") return;
	const current = getPosition();
	const index = current?.index ?? 0;
	knownMaxIndex = Math.max(knownMaxIndex, current?.maxIndex ?? index);
	if (mode === "push") {
		const nextIndex = index + 1;
		knownMaxIndex = nextIndex;
		savePosition({ index: nextIndex, maxIndex: nextIndex }, "push", url);
		return;
	}
	savePosition({ index, maxIndex: knownMaxIndex }, "replace", url);
}

/** Read and persist the known forward boundary when the browser changes entries. */
export function getNavigationAvailability(): NavigationAvailability {
	if (typeof window === "undefined") return { canGoBack: false, canGoForward: false };
	const position = getPosition();
	if (!position) return { canGoBack: false, canGoForward: false };
	knownMaxIndex = Math.max(knownMaxIndex, position.maxIndex);
	if (position.maxIndex !== knownMaxIndex) {
		savePosition({ ...position, maxIndex: knownMaxIndex }, "replace", window.location.href);
	}
	return {
		canGoBack: position.index > 0,
		canGoForward: position.index < knownMaxIndex,
	};
}
