/**
 * DocxEditor — WYSIWYG docx editing in the preview pane, powered by SuperDoc
 * (AGPL-3.0; see README "Third-party licenses"). Loaded lazily so the ~14 MB
 * editor bundle only downloads when a docx is actually opened for editing.
 *
 * Save exports DOCX bytes from SuperDoc and writes them back through the
 * binary workspace file route (`contentBase64` on PUT /api/workspace/file).
 */
import { useEffect, useRef, useState } from "react";
import { useTranslation } from "react-i18next";
import { Save, LoaderCircle } from "lucide-react";
import type { WorkspaceFileDetail } from "../../types/workspace.js";
import { saveWorkspaceFileBinary } from "../../api/workspace.js";
import "superdoc/style.css";

interface SuperDocInstance {
	export: (options?: { exportType?: string[]; triggerDownload?: boolean }) => Promise<Blob>;
	destroy?: () => void;
}

/** workspaceId is carried on the raw-file URL's query string (same as DocxPreview). */
function workspaceIdFromUrl(url: string | undefined): string | undefined {
	if (!url) return undefined;
	const query = url.split("?")[1];
	if (!query) return undefined;
	return new URLSearchParams(query).get("workspaceId") ?? undefined;
}

export default function DocxEditor({ file }: { file: WorkspaceFileDetail }) {
	const { t } = useTranslation();
	const hostRef = useRef<HTMLDivElement | null>(null);
	const docRef = useRef<SuperDocInstance | null>(null);
	const [dirty, setDirty] = useState(false);
	const [saving, setSaving] = useState(false);
	const [error, setError] = useState("");
	const [ready, setReady] = useState(false);

	useEffect(() => {
		let cancelled = false;
		setReady(false);
		setError("");
		setDirty(false);
		(async () => {
			try {
				const res = await fetch(file.url ?? "");
				if (!res.ok) throw new Error(`Failed to load document (${res.status})`);
				const data = await res.blob();
				const { SuperDoc } = await import("superdoc");
				if (cancelled || !hostRef.current) return;
				const instance = new SuperDoc({
					selector: hostRef.current,
					document: { data, name: file.name },
					documentMode: "editing",
					onEditorUpdate: () => setDirty(true),
					onReady: () => setReady(true),
					onException: (params: unknown) => {
						const message = (params as { error?: { message?: string } } | null)?.error?.message;
						setError(message ?? "SuperDoc error");
					},
				} as unknown as ConstructorParameters<typeof SuperDoc>[0]);
				docRef.current = instance as unknown as SuperDocInstance;
			} catch (err) {
				if (!cancelled) setError(err instanceof Error ? err.message : String(err));
			}
		})();
		return () => {
			cancelled = true;
			try { docRef.current?.destroy?.(); } catch { /* best effort */ }
			docRef.current = null;
		};
	}, [file.url, file.name]);

	const save = async () => {
		const doc = docRef.current;
		if (!doc || saving) return;
		setSaving(true);
		setError("");
		try {
			const blob = await doc.export({ exportType: ["docx"], triggerDownload: false });
			await saveWorkspaceFileBinary(file.path, blob, workspaceIdFromUrl(file.url));
			setDirty(false);
			// No metadata refresh here: workspaceStore.selectFile would toggle the
			// panel's loading state and unmount the editor mid-session.
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
					{dirty ? " · ●" : ""}
				</span>
				<button
					type="button"
					disabled={!ready || saving || !dirty}
					className="flex h-6 items-center gap-1 rounded-md inno-primary-button px-2 text-xs text-white disabled:opacity-50"
					onClick={() => void save()}
				>
					{saving ? <LoaderCircle size={12} className="animate-spin" /> : <Save size={12} />}
					{t("common.save", "Save")}
				</button>
			</div>
			{error ? <div className="px-3 py-2 text-xs text-[var(--inno-danger)]">{error}</div> : null}
			<div ref={hostRef} className="inno-superdoc-host min-h-0 flex-1 overflow-hidden" />
		</div>
	);
}
