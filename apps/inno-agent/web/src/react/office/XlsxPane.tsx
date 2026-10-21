/**
 * XlsxPane — xlsx preview/edit toggle. Preview is the read-only SheetJS
 * table; edit mode mounts the Univer spreadsheet (lazy chunk).
 */
import { lazy, Suspense, useState } from "react";
import { useTranslation } from "react-i18next";
import { Pencil, Eye, LoaderCircle } from "lucide-react";
import type { WorkspaceFileDetail } from "../../types/workspace.js";
import XlsxPreview from "./XlsxPreview.js";

const XlsxEditor = lazy(() => import("./XlsxEditor.js"));

export default function XlsxPane({ file }: { file: WorkspaceFileDetail }) {
	const { t } = useTranslation();
	const [mode, setMode] = useState<"preview" | "edit">("preview");

	return (
		<div className="flex h-full flex-col">
			<div className="flex h-9 shrink-0 items-center gap-1 border-b border-[var(--inno-border)] bg-[var(--inno-surface)] px-3">
				<button
					type="button"
					className={`flex h-6 items-center gap-1 rounded-md px-2 text-xs ${mode === "preview" ? "bg-[var(--inno-surface-muted)] text-[var(--inno-text)]" : "text-[var(--inno-text-muted)] hover:text-[var(--inno-text)]"}`}
					onClick={() => setMode("preview")}
				>
					<Eye size={12} />
					{t("office.previewMode", "Preview")}
				</button>
				<button
					type="button"
					className={`flex h-6 items-center gap-1 rounded-md px-2 text-xs ${mode === "edit" ? "bg-[var(--inno-surface-muted)] text-[var(--inno-text)]" : "text-[var(--inno-text-muted)] hover:text-[var(--inno-text)]"}`}
					onClick={() => setMode("edit")}
				>
					<Pencil size={12} />
					{t("office.editMode", "Edit")}
				</button>
			</div>
			<div className="workspace-scroll min-h-0 flex-1 overflow-auto">
				{mode === "preview" ? (
					<XlsxPreview file={file} />
				) : (
					<Suspense fallback={
						<div className="flex h-full items-center justify-center gap-2 text-sm text-[var(--inno-text-muted)]">
							<LoaderCircle size={14} className="animate-spin" />
							{t("preview.loadingFile")}
						</div>
					}>
						<XlsxEditor file={file} />
					</Suspense>
				)}
			</div>
		</div>
	);
}
