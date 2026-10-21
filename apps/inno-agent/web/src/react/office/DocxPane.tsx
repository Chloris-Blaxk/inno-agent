/**
 * DocxPane — docx preview with an Edit toggle. Preview mode is the existing
 * docx-preview renderer with anchored comment threads (批注); edit mode swaps
 * in the SuperDoc WYSIWYG editor. Comments are only available in preview mode:
 * editing mutates the text and invalidates content-based anchors, so threads
 * are reviewed/created in preview and the document is changed in edit mode.
 */
import { lazy, Suspense, useState } from "react";
import { useTranslation } from "react-i18next";
import { Pencil, Eye, LoaderCircle } from "lucide-react";
import type { WorkspaceFileDetail } from "../../types/workspace.js";
import DocxPreview from "./DocxPreview.js";

const DocxEditor = lazy(() => import("./DocxEditor.js"));

export default function DocxPane({ file }: { file: WorkspaceFileDetail }) {
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
				{mode === "edit" ? (
					<span className="ml-2 text-[10px] text-[var(--inno-text-subtle)]">
						{t("office.commentsInPreview", "Comments are available in Preview mode")}
					</span>
				) : null}
			</div>
			<div className="workspace-scroll min-h-0 flex-1 overflow-auto">
				{mode === "preview" ? (
					<DocxPreview file={file} />
				) : (
					<Suspense fallback={
						<div className="flex h-full items-center justify-center gap-2 text-sm text-[var(--inno-text-muted)]">
							<LoaderCircle size={14} className="animate-spin" />
							{t("preview.loadingFile")}
						</div>
					}>
						<DocxEditor file={file} />
					</Suspense>
				)}
			</div>
		</div>
	);
}
