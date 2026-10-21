/**
 * PptxEditor — WYSIWYG pptx editing via an embedded, prebuilt PPTist bundle
 * (AGPL-3.0; staged under web/public/pptist by scripts/prepare-pptist.sh).
 * The iframe talks a small postMessage protocol (see PPTist src/embed.ts):
 * host sends the pptx bytes after the ready handshake, and the Save button
 * asks PPTist to export the deck back to bytes for the binary write route.
 */
import { useCallback, useEffect, useRef, useState } from "react";
import { useTranslation } from "react-i18next";
import { Save, LoaderCircle } from "lucide-react";
import type { WorkspaceFileDetail } from "../../types/workspace.js";
import { saveWorkspaceFileBinary } from "../../api/workspace.js";

const PPTIST_URL = `/pptist/index.html?embed=1&parentOrigin=${encodeURIComponent(window.location.origin)}`;

export default function PptxEditor({ file }: { file: WorkspaceFileDetail }) {
	const { t } = useTranslation();
	const iframeRef = useRef<HTMLIFrameElement | null>(null);
	const [ready, setReady] = useState(false);
	const [loaded, setLoaded] = useState(false);
	const [saving, setSaving] = useState(false);
	const [error, setError] = useState("");
	const exportedResolver = useRef<((buf: ArrayBuffer) => void) | null>(null);

	const post = useCallback((message: Record<string, unknown>, transfer?: Transferable[]) => {
		const win = iframeRef.current?.contentWindow;
		if (!win) return;
		win.postMessage(message, window.location.origin, transfer ?? []);
	}, []);

	useEffect(() => {
		const onMessage = (e: MessageEvent) => {
			if (e.origin !== window.location.origin) return;
			if (e.source !== iframeRef.current?.contentWindow) return;
			const data = e.data as { type?: string; buffer?: ArrayBuffer } | null;
			if (!data?.type) return;
			if (data.type === "pptist:ready") setReady(true);
			if (data.type === "pptist:loaded") setLoaded(true);
			if (data.type === "pptist:exported" && data.buffer) {
				exportedResolver.current?.(data.buffer);
				exportedResolver.current = null;
			}
		};
		window.addEventListener("message", onMessage);
		return () => window.removeEventListener("message", onMessage);
	}, []);

	useEffect(() => {
		if (!ready) return;
		let cancelled = false;
		(async () => {
			try {
				const res = await fetch(file.url ?? "");
				if (!res.ok) throw new Error(`Failed to load presentation (${res.status})`);
				const buffer = await res.arrayBuffer();
				if (cancelled) return;
				post({ type: "pptist:load", buffer, name: file.name }, [buffer]);
			} catch (err) {
				if (!cancelled) setError(err instanceof Error ? err.message : String(err));
			}
		})();
		return () => { cancelled = true; };
	}, [ready, file.url, file.name, post]);

	const save = async () => {
		if (saving || !loaded) return;
		setSaving(true);
		setError("");
		try {
			const buffer = await new Promise<ArrayBuffer>((resolvePromise, reject) => {
				exportedResolver.current = resolvePromise;
				post({ type: "pptist:export" });
				setTimeout(() => reject(new Error("Export timed out")), 60_000);
			});
			const query = file.url?.split("?")[1] ?? "";
			const wsId = new URLSearchParams(query).get("workspaceId") ?? undefined;
			await saveWorkspaceFileBinary(file.path, buffer, wsId);
		} catch (err) {
			setError(err instanceof Error ? err.message : "Save failed");
		} finally {
			setSaving(false);
		}
	};

	return (
		<div className="flex h-full flex-col">
			<div className="flex h-9 shrink-0 items-center justify-between border-b border-[var(--inno-border)] bg-[var(--inno-surface)] px-3">
				<span className="text-xs text-[var(--inno-text-muted)]">
					{t("office.editingHint", "Editing mode — changes are not saved until you click Save")}
				</span>
				<button
					type="button"
					disabled={!loaded || saving}
					className="flex h-6 items-center gap-1 rounded-md inno-primary-button px-2 text-xs text-white disabled:opacity-50"
					onClick={() => void save()}
				>
					{saving ? <LoaderCircle size={12} className="animate-spin" /> : <Save size={12} />}
					{t("common.save", "Save")}
				</button>
			</div>
			{error ? <div className="px-3 py-2 text-xs text-[var(--inno-danger)]">{error}</div> : null}
			<iframe ref={iframeRef} src={PPTIST_URL} title={file.name} className="min-h-0 flex-1 border-0" />
		</div>
	);
}
