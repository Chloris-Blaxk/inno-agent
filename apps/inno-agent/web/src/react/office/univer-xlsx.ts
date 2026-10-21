/**
 * SheetJS ↔ Univer snapshot conversion for the xlsx editor.
 *
 * Univer's xlsx import/export is a Pro feature; this converter covers the
 * open-source subset we need: cell values (string/number/boolean), formulas,
 * merged ranges and column widths. Rich styles are preserved by Univer while
 * editing in-memory but are NOT written back to xlsx (the vendored SheetJS
 * community build cannot serialize styles).
 */
import * as XLSX from "xlsx";

/** Univer CellValueType: STRING=1, NUMBER=2, BOOLEAN=3, FORCE_STRING=4. */
const T_STRING = 1;
const T_NUMBER = 2;
const T_BOOLEAN = 3;
const T_FORCE_STRING = 4;

interface UniverCell {
	v?: string | number | boolean;
	t?: number;
	f?: string;
}

interface UniverSheetData {
	id: string;
	name: string;
	rowCount: number;
	columnCount: number;
	cellData: Record<number, Record<number, UniverCell>>;
	mergeData?: Array<{ startRow: number; endRow: number; startColumn: number; endColumn: number }>;
	columnData?: Record<number, { w?: number }>;
}

export interface UniverWorkbookData {
	id: string;
	name: string;
	sheetOrder: string[];
	sheets: Record<string, UniverSheetData>;
}

const MAX_ROWS = 5000;
const MAX_COLS = 128;

export function sheetJsToUniver(wb: XLSX.WorkBook, name: string): UniverWorkbookData {
	const sheets: Record<string, UniverSheetData> = {};
	const sheetOrder: string[] = [];
	wb.SheetNames.forEach((sheetName, i) => {
		const ws = wb.Sheets[sheetName];
		const id = `sheet-${i + 1}`;
		const ref = ws["!ref"] ?? "A1";
		const range = XLSX.utils.decode_range(ref);
		const rowCount = Math.min(range.e.r + 1, MAX_ROWS);
		const columnCount = Math.min(range.e.c + 1, MAX_COLS);
		const cellData: Record<number, Record<number, UniverCell>> = {};
		for (let r = 0; r < rowCount; r++) {
			for (let c = 0; c < columnCount; c++) {
				const cell = ws[XLSX.utils.encode_cell({ r, c })];
				if (!cell || cell.v == null) continue;
				const out: UniverCell = {};
				if (cell.t === "n" && typeof cell.v === "number") {
					out.v = cell.v;
					out.t = T_NUMBER;
				} else if (cell.t === "b") {
					out.v = cell.v ? 1 : 0;
					out.t = T_BOOLEAN;
				} else if (cell.t === "s" || cell.t === "str") {
					out.v = String(cell.v);
					out.t = T_STRING;
				} else {
					// dates/errors/unknown — keep the formatted text
					const text = cell.w ?? String(cell.v);
					out.v = text;
					out.t = T_FORCE_STRING;
				}
				if (cell.f) out.f = `=${cell.f}`;
				(cellData[r] ??= {})[c] = out;
			}
		}
		const mergeData = (ws["!merges"] ?? []).map((m) => ({
			startRow: m.s.r,
			endRow: m.e.r,
			startColumn: m.s.c,
			endColumn: m.e.c,
		}));
		const columnData: Record<number, { w?: number }> = {};
		(ws["!cols"] ?? []).forEach((col, c) => {
			if (col?.wpx) columnData[c] = { w: col.wpx };
		});
		sheets[id] = {
			id,
			name: sheetName.slice(0, 31),
			rowCount: Math.max(rowCount, 100),
			columnCount: Math.max(columnCount, 26),
			cellData,
			mergeData,
			columnData,
		};
		sheetOrder.push(id);
	});
	return { id: `wb-${Date.now()}`, name, sheetOrder, sheets };
}

export function univerToSheetJs(snapshot: UniverWorkbookData, name: string): XLSX.WorkBook {
	const wb = XLSX.utils.book_new();
	for (const sheetId of snapshot.sheetOrder) {
		const sheet = snapshot.sheets[sheetId];
		if (!sheet) continue;
		const aoa: Array<Array<string | number | boolean | null>> = [];
		for (const [rStr, row] of Object.entries(sheet.cellData ?? {})) {
			const r = Number(rStr);
			aoa[r] ??= [];
			for (const [cStr, cell] of Object.entries(row)) {
				const c = Number(cStr);
				if (cell.v == null && !cell.f) continue;
				aoa[r][c] = cell.t === T_BOOLEAN ? Boolean(cell.v) : (cell.v ?? null);
			}
		}
		const ws = XLSX.utils.aoa_to_sheet(aoa);
		// Re-attach formulas: aoa_to_sheet only writes values.
		for (const [rStr, row] of Object.entries(sheet.cellData ?? {})) {
			for (const [cStr, cell] of Object.entries(row)) {
				if (!cell.f) continue;
				const addr = XLSX.utils.encode_cell({ r: Number(rStr), c: Number(cStr) });
				ws[addr] = { ...(ws[addr] ?? { t: "s", v: "" }), f: cell.f.replace(/^=/, "") } as XLSX.CellObject;
			}
		}
		if (sheet.mergeData?.length) {
			ws["!merges"] = sheet.mergeData.map((m) => ({
				s: { r: m.startRow, c: m.startColumn },
				e: { r: m.endRow, c: m.endColumn },
			}));
		}
		const cols = Object.entries(sheet.columnData ?? {})
			.map(([cStr, col]) => ({ c: Number(cStr), wpx: col.w }))
			.filter((col) => col.wpx != null);
		if (cols.length) {
			const maxC = Math.max(...cols.map((col) => col.c));
			ws["!cols"] = Array.from({ length: maxC + 1 }, (_, c) => {
				const found = cols.find((col) => col.c === c);
				return found ? { wpx: found.wpx } : {};
			});
		}
		XLSX.utils.book_append_sheet(wb, ws, sheet.name.slice(0, 31) || `Sheet${sheetId}`);
	}
	if (wb.SheetNames.length === 0) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([]), "Sheet1");
	return wb;
}

export function univerSnapshotToXlsxBytes(snapshot: UniverWorkbookData, name: string): Uint8Array {
	const wb = univerToSheetJs(snapshot, name);
	return XLSX.write(wb, { bookType: "xlsx", type: "array" }) as Uint8Array;
}
