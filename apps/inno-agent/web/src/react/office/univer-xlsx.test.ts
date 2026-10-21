import { describe, expect, it } from "vitest";
import * as XLSX from "xlsx";
import { sheetJsToUniver, univerSnapshotToXlsxBytes, univerToSheetJs } from "./univer-xlsx.js";

describe("univer-xlsx converter", () => {
	it("round-trips values, types and formulas", () => {
		const wb = XLSX.utils.book_new();
		const ws = XLSX.utils.aoa_to_sheet([
			["名称", "数量", "达标"],
			["甲班", 3, true],
			["乙班", 4, false],
		]);
		ws["C2"].f = "B2>2";
		XLSX.utils.book_append_sheet(wb, ws, "统计");

		const snapshot = sheetJsToUniver(wb, "test.xlsx");
		const sheet = snapshot.sheets[snapshot.sheetOrder[0]];
		expect(sheet.cellData[0][0]).toEqual({ v: "名称", t: 1 });
		expect(sheet.cellData[1][1]).toEqual({ v: 3, t: 2 });
		expect(sheet.cellData[1][2]).toEqual({ v: 1, t: 3, f: "=B2>2" });
		expect(sheet.cellData[2][2]).toMatchObject({ v: 0, t: 3 });

		const back = univerToSheetJs(snapshot, "test.xlsx");
		const backWs = back.Sheets["统计"];
		expect(backWs["A1"].v).toBe("名称");
		expect(backWs["B2"].v).toBe(3);
		expect(backWs["C2"].v).toBe(true);
		expect(backWs["C2"].f).toBe("B2>2");
	});

	it("preserves merges and column widths", () => {
		const wb = XLSX.utils.book_new();
		const ws = XLSX.utils.aoa_to_sheet([["a", "b"], ["c", "d"]]);
		ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }];
		ws["!cols"] = [{ wpx: 120 }, {}];
		XLSX.utils.book_append_sheet(wb, ws, "S");

		const snapshot = sheetJsToUniver(wb, "m.xlsx");
		const sheet = snapshot.sheets[snapshot.sheetOrder[0]];
		expect(sheet.mergeData).toEqual([{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 }]);
		expect(sheet.columnData?.[0]).toEqual({ w: 120 });

		const back = univerToSheetJs(snapshot, "m.xlsx");
		expect(back.Sheets["S"]["!merges"]).toEqual([{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }]);
		expect(back.Sheets["S"]["!cols"]?.[0]).toEqual({ wpx: 120 });
	});

	it("keeps multi-sheet order and produces valid xlsx bytes", () => {
		const wb = XLSX.utils.book_new();
		XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["一"]]), "第一个");
		XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["二"]]), "第二个");

		const snapshot = sheetJsToUniver(wb, "multi.xlsx");
		expect(snapshot.sheetOrder).toHaveLength(2);

		const bytes = univerSnapshotToXlsxBytes(snapshot, "multi.xlsx");
		const reparsed = XLSX.read(bytes, { type: "array" });
		expect(reparsed.SheetNames).toEqual(["第一个", "第二个"]);
		expect(reparsed.Sheets["第二个"]["A1"].v).toBe("二");
	});

	it("caps huge ranges", () => {
		const wb = XLSX.utils.book_new();
		const ws = XLSX.utils.aoa_to_sheet([["x"]]);
		ws["!ref"] = "A1:ZZ100000";
		XLSX.utils.book_append_sheet(wb, ws, "big");
		const snapshot = sheetJsToUniver(wb, "big.xlsx");
		const sheet = snapshot.sheets[snapshot.sheetOrder[0]];
		expect(sheet.rowCount).toBeLessThanOrEqual(5000);
		expect(sheet.columnCount).toBeGreaterThanOrEqual(26);
	});
});
