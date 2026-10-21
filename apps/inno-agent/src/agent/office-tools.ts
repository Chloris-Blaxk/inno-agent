/**
 * Office document tools (Office 文档引擎工具) — LibreOffice UNO backend.
 *
 * Exposes a small, schema-on-demand surface instead of one tool per atomic
 * operation (WorkBuddy's editor_sdk ships 37/51/77+ tools; we keep the tool
 * count low and let the model pull structure on demand):
 *
 *   office_engine_status   — is the bundled/system LibreOffice usable?
 *   office_open_document   — open docx/xlsx/pptx (+odt/ods/odp), returns
 *                            docId, kind and a structural outline.
 *   office_edit_document   — apply a batch of atomic ops (see OPS below),
 *                            then optionally save in place or save-as.
 *   office_render_document — render an open doc or a file to PDF (visual
 *                            render-check loop / preview asset).
 *   office_close_document  — release the document handle.
 *
 * All file access is constrained to the active workspace via
 * checkWorkspaceMutationPath.
 */

import { mkdirSync } from "node:fs";
import { defineTool, type ToolDefinition } from "@earendil-works/pi-coding-agent";
import { Type } from "typebox";
import { logger } from "../logger.js";
import { checkWorkspaceMutationPath } from "./workspace-path-guard.js";
import { resolveLibreOffice } from "../office/engine-locator.js";
import { getUnoSidecar, type UnoCallResult } from "../office/uno-sidecar.js";

const OP_VOCABULARY = [
	"doc（Word）: replace_text{find,replace,all?}、set_paragraph{index,text}、",
	"  insert_paragraph{text,style?,index?}（index 表示插到该段落之前，缺省追加到文末）、",
	"  delete_paragraph{index}、insert_table{rows:[[...]],index?}、list_paragraphs{}",
	"sheet（Excel）: list_sheets{}、get_range{sheet?,range}、set_cell{sheet?,cell,value}、",
	"  append_row{sheet?,values[]}、insert_rows{sheet?,index,count}",
	"slide（PPT）: list_slides{}、add_slide{index?,title?,body?}、",
	"  set_slide_text{slide,shape,text}、delete_slide{index}",
].join("\n");

function textResult(text: string, details?: Record<string, unknown>) {
	return {
		content: [{ type: "text" as const, text }],
		details: (details ?? {}) as Record<string, unknown>,
	};
}

function formatFailure(step: string, res: UnoCallResult): string {
	return `Office 引擎调用失败（${step}）：${res.error ?? "未知错误"}`;
}

export function createOfficeTools(getWorkspaceDir: () => string): ToolDefinition[] {
	function resolveWorkspacePath(filePath: string): { path: string } | { error: string } {
		const workspaceDir = getWorkspaceDir();
		const check = checkWorkspaceMutationPath(workspaceDir, filePath);
		if (!check.allowed || !check.resolvedPath) {
			return {
				error:
					`路径不在当前工作区内，已拒绝：${filePath}。` +
					"请使用当前工作区内的相对路径或工作区内的绝对路径。",
			};
		}
		return { path: check.resolvedPath };
	}

	const engineStatus = defineTool({
		name: "office_engine_status",
		label: "Office 引擎状态",
		description:
			"检查 LibreOffice 文档引擎是否可用（随包内置或系统安装），返回版本与来源。" +
			"在处理 Word/Excel/PPT 文档前可先调用以确认能力；不可用时会给出安装指引。",
		parameters: Type.Object({}),
		async execute() {
			const status = resolveLibreOffice();
			if (!status.available) {
				return textResult(
					"未找到 LibreOffice 引擎。桌面安装包已内置；当前环境请安装 LibreOffice" +
					"（macOS: brew install --cask libreoffice，Ubuntu: apt-get install libreoffice）" +
					"或设置 INNO_LIBREOFFICE_PATH 指向 soffice。",
					{ available: false },
				);
			}
			return textResult(
				`LibreOffice 可用：${status.version ?? "版本未知"}（来源：${status.source}，路径：${status.sofficePath}）`,
				{ ...status },
			);
		},
	});

	const openDocument = defineTool({
		name: "office_open_document",
		label: "打开 Office 文档",
		description:
			"用 LibreOffice 引擎打开工作区内的 docx/xlsx/pptx（也支持 odt/ods/odp）文档，" +
			"返回 docId、文档类型和结构大纲（段落列表 / 工作表列表 / 幻灯片列表）。" +
			"之后用 office_edit_document 基于 docId 做原子编辑，用 office_close_document 关闭。",
		parameters: Type.Object({
			path: Type.String({ description: "工作区内文档路径（相对或绝对）" }),
			readOnly: Type.Optional(Type.Boolean({ description: "只读打开（默认 false）" })),
		}),
		async execute(_toolCallId, params) {
			const { path: rawPath, readOnly } = params as { path: string; readOnly?: boolean };
			const resolved = resolveWorkspacePath(String(rawPath ?? ""));
			if ("error" in resolved) return textResult(resolved.error, { error: "path_rejected" });
			const sidecar = getUnoSidecar();
			const opened = await sidecar.call<{ docId: string; kind: string; path: string }>(
				"open", { path: resolved.path, readOnly: readOnly === true });
			if (!opened.ok) return textResult(formatFailure("open", opened), { error: opened.error });
			const outline = await sidecar.call("inspect", { docId: opened.result!.docId });
			if (!outline.ok) {
				await sidecar.call("close", { docId: opened.result!.docId });
				return textResult(formatFailure("inspect", outline), { error: outline.error });
			}
			const { docId, kind } = opened.result!;
			logger.info({ docId, kind, path: resolved.path }, "office_open_document: opened");
			return textResult(
				`已打开 ${kind} 文档（docId: ${docId}）。结构大纲：\n` +
				JSON.stringify(outline.result, null, 2),
				{ docId, kind, outline: outline.result },
			);
		},
	});

	const editDocument = defineTool({
		name: "office_edit_document",
		label: "编辑 Office 文档",
		description:
			"对已打开的文档执行一批原子编辑操作（按顺序执行，单个失败不中断后续），" +
			"可选择保存。先用 office_open_document 拿到 docId 和结构大纲再编辑。\n" +
			"支持的操作：\n" + OP_VOCABULARY,
		parameters: Type.Object({
			docId: Type.String({ description: "office_open_document 返回的 docId" }),
			ops: Type.Array(Type.Record(Type.String(), Type.Unknown()), {
				description: "原子操作列表，如 [{\"op\":\"replace_text\",\"find\":\"旧\",\"replace\":\"新\"}]",
			}),
			save: Type.Optional(Type.Boolean({ description: "编辑后保存到原文件（默认 false）" })),
			saveAs: Type.Optional(Type.String({ description: "另存为工作区内路径（优先于 save）" })),
		}),
		async execute(_toolCallId, params) {
			const { docId, ops, save, saveAs } = params as {
				docId: string; ops: Record<string, unknown>[]; save?: boolean; saveAs?: string;
			};
			if (!docId) return textResult("缺少 docId。", { error: "missing_doc_id" });
			if (!Array.isArray(ops) || ops.length === 0) {
				return textResult("ops 为空：至少需要一个原子操作。", { error: "empty_ops" });
			}
			if (ops.length > 100) {
				return textResult("单次最多 100 个操作，请分批提交。", { error: "too_many_ops" });
			}
			const sidecar = getUnoSidecar();
			const edited = await sidecar.call<{ results: { op: string; ok: boolean; error?: string }[] }>(
				"exec", { docId, ops });
			if (!edited.ok) return textResult(formatFailure("exec", edited), { error: edited.error });

			const results = edited.result!.results;
			const failed = results.filter((r) => !r.ok);

			let saveNote = "";
			if (failed.length === 0 && (save === true || saveAs)) {
				let target: string | undefined;
				if (saveAs) {
					const resolved = resolveWorkspacePath(String(saveAs));
					if ("error" in resolved) return textResult(resolved.error, { error: "path_rejected" });
					target = resolved.path;
				}
				const saved = await sidecar.call("save", target ? { docId, path: target } : { docId });
				saveNote = saved.ok
					? `\n已保存：${(saved.result as { saved: string }).saved}`
					: `\n保存失败：${saved.error}`;
			} else if (failed.length > 0 && (save === true || saveAs)) {
				saveNote = "\n存在失败操作，已跳过保存（修正后可再次保存）。";
			}

			const lines = results.map((r, i) =>
				r.ok ? `✓ #${i + 1} ${r.op}` : `✗ #${i + 1} ${r.op}：${r.error}`);
			return textResult(
				`执行 ${results.length} 个操作，成功 ${results.length - failed.length}，失败 ${failed.length}：\n` +
				lines.join("\n") + saveNote,
				{ results, saved: saveNote.includes("已保存") },
			);
		},
	});

	const renderDocument = defineTool({
		name: "office_render_document",
		label: "渲染 Office 文档为 PDF",
		description:
			"把工作区内的 Office 文档渲染成 PDF，用于生成后的视觉校验或预览。" +
			"返回 PDF 的工作区内路径（默认写到 .inno-previews/ 下）。",
		parameters: Type.Object({
			path: Type.String({ description: "工作区内文档路径（相对或绝对）" }),
			outDir: Type.Optional(Type.String({ description: "输出目录（工作区内，默认 .inno-previews）" })),
		}),
		async execute(_toolCallId, params) {
			const { path: rawPath, outDir } = params as { path: string; outDir?: string };
			const resolved = resolveWorkspacePath(String(rawPath ?? ""));
			if ("error" in resolved) return textResult(resolved.error, { error: "path_rejected" });
			const dirResolved = resolveWorkspacePath(outDir?.trim() || ".inno-previews");
			if ("error" in dirResolved) return textResult(dirResolved.error, { error: "path_rejected" });
			mkdirSync(dirResolved.path, { recursive: true });
			const sidecar = getUnoSidecar();
			const converted = await sidecar.call<{ output: string }>(
				"convert", { path: resolved.path, to: "pdf", outDir: dirResolved.path });
			if (!converted.ok) return textResult(formatFailure("convert", converted), { error: converted.error });
			return textResult(`已渲染 PDF：${converted.result!.output}`, { output: converted.result!.output });
		},
	});

	const closeDocument = defineTool({
		name: "office_close_document",
		label: "关闭 Office 文档",
		description:
			"关闭 office_open_document 打开的文档句柄，释放 LibreOffice 内存。" +
			"编辑任务结束后调用；未保存的修改会丢失（需要保存请用 office_edit_document 的 save/saveAs）。",
		parameters: Type.Object({
			docId: Type.String({ description: "要关闭的 docId" }),
		}),
		async execute(_toolCallId, params) {
			const { docId } = params as { docId: string };
			if (!docId) return textResult("缺少 docId。", { error: "missing_doc_id" });
			const sidecar = getUnoSidecar();
			const closed = await sidecar.call("close", { docId });
			if (!closed.ok) return textResult(formatFailure("close", closed), { error: closed.error });
			return textResult(`已关闭文档 ${docId}。`, { closed: docId });
		},
	});

	return [engineStatus, openDocument, editDocument, renderDocument, closeDocument];
}
