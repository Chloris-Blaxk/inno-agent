/**
 * PptxPane — pptx preview/edit toggle. Preview is the server-rendered SVG
 * deck; edit mode embeds the prebuilt PPTist bundle in an iframe.
 */
import { useState } from "react";
import { useTranslation } from "react-i18next";
import { Pencil, Eye } from "lucide-react";
import type { WorkspaceFileDetail } from "../../types/workspace.js";
import PptxPreview from "./PptxPreview.js";
import PptxEditor from "./PptxEditor.js";

export default function PptxPane({ file }: { file: WorkspaceFileDetail }) {
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
				{mode === "preview" ? <PptxPreview file={file} /> : <PptxEditor file={file} />}
			</div>
		</div>
	);
}
