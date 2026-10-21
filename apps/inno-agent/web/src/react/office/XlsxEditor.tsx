/**
 * XlsxEditor — WYSIWYG xlsx editing in the preview pane, powered by Univer
 * (Apache-2.0). Loaded lazily with its presets; xlsx bytes cross the
 * SheetJS ↔ Univer boundary in univer-xlsx.ts (values, formulas, merges,
 * column widths — rich cell styles are not serialized back to xlsx).
 */
import { useEffect, useRef, useState } from "react";
import { useTranslation } from "react-i18next";
import { Save, LoaderCircle } from "lucide-react";
import * as XLSX from "xlsx";
import type { WorkspaceFileDetail } from "../../types/workspace.js";
import { saveWorkspaceFileBinary } from "../../api/workspace.js";
import { sheetJsToUniver, univerSnapshotToXlsxBytes, type UniverWorkbookData } from "./univer-xlsx.js";

interface UniverHandles {
	save: () => UniverWorkbookData | null;
	dispose: () => void;
}

export default function XlsxEditor({ file }: { file: WorkspaceFileDetail }) {
	const { t, i18n } = useTranslation();
	const hostRef = useRef<HTMLDivElement | null>(null);
	const handlesRef = useRef<UniverHandles | null>(null);
	const [saving, setSaving] = useState(false);
	const [error, setError] = useState("");
	const [ready, setReady] = useState(false);

	useEffect(() => {
		let cancelled = false;
		setReady(false);
		setError("");
		(async () => {
			try {
				const res = await fetch(file.url ?? "");
				if (!res.ok) throw new Error(`Failed to load workbook (${res.status})`);
				const buf = await res.arrayBuffer();
				const wb = XLSX.read(buf, { type: "array", cellFormula: true });
				const snapshot = sheetJsToUniver(wb, file.name);

				const [{ createUniver, LocaleType, UniverInstanceType }, { UniverSheetsCorePreset }] = await Promise.all([
					import("@univerjs/presets"),
					import("@univerjs/preset-sheets-core"),
				]);
				const localeKey = i18n.language.startsWith("zh") ? LocaleType.ZH_CN : LocaleType.EN_US;
				const localePack = localeKey === LocaleType.ZH_CN
					? (await import("@univerjs/preset-sheets-core/locales/zh-CN")).default
					: (await import("@univerjs/preset-sheets-core/locales/en-US")).default;
				await import("@univerjs/preset-sheets-core/lib/index.css");
				if (cancelled || !hostRef.current) return;

				const { univer, univerAPI } = createUniver({
					locale: localeKey,
					locales: { [localeKey]: localePack },
					presets: [
						UniverSheetsCorePreset({
							container: hostRef.current,
						}),
					],
				});
				univer.createUnit(UniverInstanceType.UNIVER_SHEET, snapshot as never);
				handlesRef.current = {
					save: () => {
						const workbook = univerAPI.getActiveWorkbook();
						return (workbook?.save() ?? null) as UniverWorkbookData | null;
					},
					dispose: () => {
						try { univerAPI.dispose(); } catch { /* best effort */ }
						try { univer.dispose(); } catch { /* best effort */ }
					},
				};
				setReady(true);
			} catch (err) {
				if (!cancelled) setError(err instanceof Error ? err.message : String(err));
			}
		})();
		return () => {
			cancelled = true;
			handlesRef.current?.dispose();
			handlesRef.current = null;
		};
	}, [file.url, file.name, i18n.language]);

	const save = async () => {
		if (saving) return;
		setSaving(true);
		setError("");
		try {
			const snapshot = handlesRef.current?.save();
			if (!snapshot) throw new Error("No workbook snapshot");
			const bytes = univerSnapshotToXlsxBytes(snapshot, file.name);
			const query = file.url?.split("?")[1] ?? "";
			const wsId = new URLSearchParams(query).get("workspaceId") ?? undefined;
			await saveWorkspaceFileBinary(file.path, bytes, wsId);
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
					disabled={!ready || saving}
					className="flex h-6 items-center gap-1 rounded-md inno-primary-button px-2 text-xs text-white disabled:opacity-50"
					onClick={() => void save()}
				>
					{saving ? <LoaderCircle size={12} className="animate-spin" /> : <Save size={12} />}
					{t("common.save", "Save")}
				</button>
			</div>
			{error ? <div className="px-3 py-2 text-xs text-[var(--inno-danger)]">{error}</div> : null}
			<div ref={hostRef} className="min-h-0 flex-1 overflow-hidden" />
		</div>
	);
}
