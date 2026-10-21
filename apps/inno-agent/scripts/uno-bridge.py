# uno-bridge.py — JSON-lines stdio bridge between the Inno Agent backend and a
# headless LibreOffice process.
#
# This script must run under LibreOffice's bundled python (macOS:
# LibreOffice.app/Contents/Resources/python, Windows: program/python.exe,
# Linux: program/python or a system python with python3-uno). It launches
# `soffice --headless` with a private user profile and a loopback UNO socket,
# connects to it, and then serves newline-delimited JSON requests on stdin:
#
#   ->  {"id": 1, "cmd": "open",  "path": "/abs/file.docx"}
#   <-  {"id": 1, "ok": true, "result": {"docId": "...", "kind": "doc", ...}}
#   ->  {"id": 2, "cmd": "exec",  "docId": "...", "ops": [{"op": "replace_text", ...}]}
#   <-  {"id": 2, "ok": true, "result": {"results": [...]}}
#
# Every failure comes back as {"id": N, "ok": false, "error": "..."} — the
# bridge never raises across the wire. Logs go to stderr only.
#
# Commands: ping, open, close, inspect, exec, save, convert, shutdown.
# See OPS below for the atomic edit vocabulary (doc / sheet / slide).

import argparse
import json
import os
import signal
import subprocess
import sys
import tempfile
import time
import traceback
import uuid

import uno
from com.sun.star.beans import PropertyValue


def log(msg):
    sys.stderr.write(f"[uno-bridge] {msg}\n")
    sys.stderr.flush()


def send(obj):
    sys.stdout.write(json.dumps(obj, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


# Filters for saving back to OOXML / exporting PDF, keyed by document kind.
SAVE_FILTERS = {
    ("doc", ".docx"): "MS Word 2007 XML",
    ("doc", ".odt"): "writer8",
    ("sheet", ".xlsx"): "Calc MS Excel 2007 XML",
    ("sheet", ".ods"): "calc8",
    ("slide", ".pptx"): "Impress MS PowerPoint 2007 XML",
    ("slide", ".odp"): "impress8",
}
PDF_FILTERS = {
    "doc": "writer_pdf_Export",
    "sheet": "calc_pdf_Export",
    "slide": "impress_pdf_Export",
}


def file_url(path):
    return uno.systemPathToFileUrl(os.path.abspath(path))


class OfficeEngine:
    """Owns the soffice process, the UNO connection and the open documents."""

    def __init__(self, soffice_path, port):
        self.soffice_path = soffice_path
        self.port = port
        self.profile_dir = tempfile.mkdtemp(prefix="inno-uno-profile-")
        self.proc = None
        self.ctx = None
        self.desktop = None
        self.docs = {}  # docId -> {"component": ..., "kind": ..., "path": ...}

    # ── lifecycle ──────────────────────────────────────────────────────────
    def start(self):
        args = [
            self.soffice_path,
            "--headless", "--invisible", "--norestore", "--nologo", "--nofirststartwizard",
            f"-env:UserInstallation={uno.systemPathToFileUrl(self.profile_dir)}",
            f"--accept=socket,host=127.0.0.1,port={self.port};urp;",
        ]
        self.proc = subprocess.Popen(
            args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        local_ctx = uno.getComponentContext()
        resolver = local_ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_ctx)
        url = (f"uno:socket,host=127.0.0.1,port={self.port};urp;"
               "StarOffice.ComponentContext")
        last_err = None
        for _ in range(60):  # up to ~60s for cold start
            if self.proc.poll() is not None:
                raise RuntimeError(
                    f"soffice exited during startup (code {self.proc.returncode})")
            try:
                self.ctx = resolver.resolve(url)
                break
            except Exception as err:  # NoConnectException while booting
                last_err = err
                time.sleep(1)
        if self.ctx is None:
            raise RuntimeError(f"could not connect to soffice: {last_err}")
        smgr = self.ctx.ServiceManager
        self.desktop = smgr.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.ctx)

    def shutdown(self):
        for doc_id in list(self.docs):
            try:
                self.close(doc_id)
            except Exception:
                pass
        try:
            if self.desktop is not None:
                self.desktop.terminate()
        except Exception:
            pass
        if self.proc is not None:
            try:
                self.proc.wait(timeout=10)
            except Exception:
                self.proc.kill()

    # ── documents ──────────────────────────────────────────────────────────
    @staticmethod
    def _kind_of(component):
        if component.supportsService("com.sun.star.text.TextDocument"):
            return "doc"
        if component.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
            return "sheet"
        if component.supportsService("com.sun.star.presentation.PresentationDocument"):
            return "slide"
        return None

    def open(self, path, read_only=False):
        if not os.path.isfile(path):
            raise FileNotFoundError(path)
        props = [prop("Hidden", True)]
        if read_only:
            props.append(prop("ReadOnly", True))
        component = self.desktop.loadComponentFromURL(
            file_url(path), "_blank", 0, tuple(props))
        if component is None:
            raise RuntimeError(f"LibreOffice could not open {path}")
        kind = self._kind_of(component)
        if kind is None:
            component.close(False)
            raise RuntimeError("unsupported document type (not doc/sheet/slide)")
        doc_id = uuid.uuid4().hex[:12]
        self.docs[doc_id] = {"component": component, "kind": kind, "path": path}
        return {"docId": doc_id, "kind": kind, "path": path}

    def _doc(self, doc_id):
        entry = self.docs.get(doc_id)
        if entry is None:
            raise KeyError(f"unknown docId {doc_id} (closed or engine restarted)")
        return entry

    def close(self, doc_id):
        entry = self.docs.pop(doc_id)
        entry["component"].close(False)
        return {"closed": doc_id}

    def save(self, doc_id, path=None):
        entry = self._doc(doc_id)
        component = entry["component"]
        if path is None:
            component.store()
            return {"saved": entry["path"]}
        ext = os.path.splitext(path)[1].lower()
        filter_name = SAVE_FILTERS.get((entry["kind"], ext))
        if filter_name is None:
            raise ValueError(f"no save filter for kind={entry['kind']} ext={ext}")
        component.storeToURL(file_url(path), (prop("FilterName", filter_name),))
        return {"saved": path}

    def convert(self, path, target, out_dir):
        entry_id = self.open(path, read_only=True)["docId"]
        try:
            entry = self._doc(entry_id)
            ext = "." + target.lstrip(".").lower()
            if ext == ".pdf":
                filter_name = PDF_FILTERS.get(entry["kind"])
            else:
                filter_name = SAVE_FILTERS.get((entry["kind"], ext))
            if filter_name is None:
                raise ValueError(f"cannot convert {entry['kind']} to {ext}")
            base = os.path.splitext(os.path.basename(path))[0]
            out_path = os.path.join(out_dir, base + ext)
            entry["component"].storeToURL(
                file_url(out_path), (prop("FilterName", filter_name),))
            return {"output": out_path}
        finally:
            self.close(entry_id)

    # ── inspection (schema-on-demand outline) ──────────────────────────────
    def inspect(self, doc_id):
        entry = self._doc(doc_id)
        kind = entry["kind"]
        if kind == "doc":
            return {"kind": kind, **self._inspect_doc(entry["component"])}
        if kind == "sheet":
            return {"kind": kind, **self._inspect_sheet(entry["component"])}
        return {"kind": kind, **self._inspect_slide(entry["component"])}

    @staticmethod
    def _doc_paragraphs(doc):
        """All text paragraphs in document order (main text body only)."""
        result = []
        enum = doc.Text.createEnumeration()
        while enum.hasMoreElements():
            el = enum.nextElement()
            if el.supportsService("com.sun.star.text.Paragraph"):
                result.append(el)
        return result

    def _inspect_doc(self, doc):
        paragraphs = self._doc_paragraphs(doc)
        items = []
        for i, p in enumerate(paragraphs[:500]):
            items.append({
                "index": i,
                "style": p.ParaStyleName,
                "text": p.getString()[:200],
            })
        table_count = doc.TextTables.Count
        return {
            "paragraphCount": len(paragraphs),
            "tableCount": table_count,
            "truncated": len(paragraphs) > 500,
            "paragraphs": items,
        }

    def _inspect_sheet(self, doc):
        sheets = doc.Sheets
        out = []
        for name in sheets.ElementNames:
            sheet = sheets.getByName(name)
            cursor = sheet.createCursor()
            cursor.gotoEndOfUsedArea(False)
            addr = cursor.RangeAddress
            out.append({
                "name": name,
                "rows": addr.EndRow + 1,
                "cols": addr.EndColumn + 1,
            })
        return {"sheets": out}

    def _inspect_slide(self, doc):
        pages = doc.getDrawPages()
        out = []
        for i in range(pages.Count):
            page = pages.getByIndex(i)
            texts = []
            for j in range(page.Count):
                shape = page.getByIndex(j)
                try:
                    text = shape.getString()
                except Exception:
                    text = ""
                if text:
                    texts.append({"shape": j, "text": text[:120]})
            out.append({"index": i, "texts": texts})
        return {"slideCount": pages.Count, "slides": out}

    # ── atomic edits ───────────────────────────────────────────────────────
    def exec_ops(self, doc_id, ops):
        entry = self._doc(doc_id)
        results = []
        for op in ops:
            name = op.get("op")
            handler = getattr(self, f"_op_{name}", None)
            if handler is None:
                results.append({"op": name, "ok": False,
                                "error": f"unknown op for engine: {name}"})
                continue
            try:
                kwargs = {k: v for k, v in op.items() if k != "op"}
                results.append({"op": name, "ok": True,
                                "result": handler(entry, **kwargs)})
            except Exception as err:
                results.append({"op": name, "ok": False, "error": str(err)})
        return {"results": results}

    # doc ops ----------------------------------------------------------------
    def _op_list_paragraphs(self, entry):
        if entry["kind"] != "doc":
            raise ValueError("list_paragraphs is a doc op")
        return self._inspect_doc(entry["component"])

    def _op_replace_text(self, entry, find="", replace="", all=True):
        doc = entry["component"]
        if entry["kind"] != "doc" or not find:
            raise ValueError("replace_text needs kind=doc and a find string")
        rd = doc.createReplaceDescriptor()
        rd.SearchString = find
        rd.ReplaceString = replace
        if all:
            return {"replaced": doc.replaceAll(rd)}
        found = doc.findFirst(rd)
        if found is None:
            return {"replaced": 0}
        found.setString(replace)
        return {"replaced": 1}

    def _op_set_paragraph(self, entry, index=-1, text=""):
        paragraphs = self._doc_paragraphs(entry["component"])
        if not 0 <= index < len(paragraphs):
            raise IndexError(f"paragraph index {index} out of range "
                             f"(0..{len(paragraphs) - 1})")
        paragraphs[index].setString(text)
        return {"index": index}

    def _op_delete_paragraph(self, entry, index=-1):
        doc = entry["component"]
        paragraphs = self._doc_paragraphs(doc)
        if not 0 <= index < len(paragraphs):
            raise IndexError(f"paragraph index {index} out of range "
                             f"(0..{len(paragraphs) - 1})")
        target = paragraphs[index]
        cursor = doc.Text.createTextCursorByRange(target.getStart())
        cursor.gotoRange(target.getEnd(), True)
        # Swallow the trailing paragraph break when there is a next paragraph.
        if index + 1 < len(paragraphs):
            cursor.goRight(1, True)
        cursor.setString("")
        return {"deleted": index}

    def _op_insert_paragraph(self, entry, text="", style=None, index=None):
        doc = entry["component"]
        paragraphs = self._doc_paragraphs(doc)
        text_obj = doc.Text
        if index is None or index >= len(paragraphs):
            cursor = text_obj.createTextCursor()
            cursor.gotoEnd(False)
            if paragraphs:
                text_obj.insertString(cursor, "\n", False)
            text_obj.insertString(cursor, text, False)
            new_range = cursor.getEnd()
        else:
            if not 0 <= index < len(paragraphs):
                raise IndexError(f"paragraph index {index} out of range")
            anchor = paragraphs[index]
            cursor = text_obj.createTextCursorByRange(anchor.getStart())
            text_obj.insertString(cursor, text + "\n", False)
            new_range = anchor.getStart()
        if style:
            style_cursor = text_obj.createTextCursorByRange(new_range)
            style_cursor.ParaStyleName = style
        return {"inserted": text[:80]}

    def _op_insert_table(self, entry, rows=None, index=None):
        doc = entry["component"]
        rows = rows or []
        if not rows or not all(isinstance(r, list) and r for r in rows):
            raise ValueError("insert_table needs rows: [[cell, ...], ...]")
        n_rows = len(rows)
        n_cols = max(len(r) for r in rows)
        table = doc.createInstance("com.sun.star.text.TextTable")
        table.initialize(n_rows, n_cols)
        paragraphs = self._doc_paragraphs(doc)
        text_obj = doc.Text
        if index is None or index >= len(paragraphs):
            cursor = text_obj.createTextCursor()
            cursor.gotoEnd(False)
        else:
            anchor = paragraphs[index]
            cursor = text_obj.createTextCursorByRange(anchor.getStart())
        text_obj.insertTextContent(cursor, table, False)
        for r, row in enumerate(rows):
            for c, cell in enumerate(row):
                table.getCellByPosition(c, r).setString(str(cell))
        return {"rows": n_rows, "cols": n_cols}

    # sheet ops --------------------------------------------------------------
    def _sheet(self, entry, name):
        if entry["kind"] != "sheet":
            raise ValueError("op requires a spreadsheet document")
        sheets = entry["component"].Sheets
        if name is None:
            return sheets.getByIndex(0)
        if name not in sheets.ElementNames:
            raise KeyError(f"no sheet named {name!r}; have {list(sheets.ElementNames)}")
        return sheets.getByName(name)

    def _op_get_range(self, entry, sheet=None, range="A1:C10"):
        s = self._sheet(entry, sheet)
        cells = s.getCellRangeByName(range)
        return {"range": range, "values": [list(row) for row in cells.getDataArray()]}

    def _op_set_cell(self, entry, sheet=None, cell="", value=None):
        s = self._sheet(entry, sheet)
        if not cell:
            raise ValueError("set_cell needs a cell address like B2")
        target = s.getCellRangeByName(cell)
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            target.setValue(value)
        elif value is None:
            target.setString("")
        else:
            target.setString(str(value))
        return {"cell": cell}

    def _op_append_row(self, entry, sheet=None, values=None):
        s = self._sheet(entry, sheet)
        values = values or []
        cursor = s.createCursor()
        cursor.gotoEndOfUsedArea(False)
        row_index = cursor.RangeAddress.EndRow + 1
        for col, value in enumerate(values):
            cell = s.getCellByPosition(col, row_index)
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                cell.setValue(value)
            else:
                cell.setString("" if value is None else str(value))
        return {"row": row_index + 1}  # 1-based for the caller

    def _op_insert_rows(self, entry, sheet=None, index=0, count=1):
        s = self._sheet(entry, sheet)
        s.Rows.insertByIndex(int(index), int(count))
        return {"inserted": int(count), "at": int(index)}

    # slide ops --------------------------------------------------------------
    def _slide_page(self, entry, index):
        if entry["kind"] != "slide":
            raise ValueError("op requires a presentation document")
        pages = entry["component"].getDrawPages()
        if not 0 <= index < pages.Count:
            raise IndexError(f"slide index {index} out of range (0..{pages.Count - 1})")
        return pages.getByIndex(index)

    def _op_add_slide(self, entry, index=None, title=None, body=None):
        doc = entry["component"]
        if entry["kind"] != "slide":
            raise ValueError("add_slide requires a presentation document")
        pages = doc.getDrawPages()
        at = pages.Count if index is None else int(index)
        page = pages.insertNewByIndex(at)
        texts = [t for t in (title, body) if t]
        placed = 0
        for j in range(page.Count):
            if placed >= len(texts):
                break
            shape = page.getByIndex(j)
            try:
                shape.setString(texts[placed])
                placed += 1
            except Exception:
                continue
        return {"index": at, "textShapesFilled": placed}

    def _op_set_slide_text(self, entry, slide=-1, shape=-1, text=""):
        page = self._slide_page(entry, int(slide))
        if not 0 <= int(shape) < page.Count:
            raise IndexError(f"shape index {shape} out of range (0..{page.Count - 1})")
        page.getByIndex(int(shape)).setString(text)
        return {"slide": int(slide), "shape": int(shape)}

    def _op_delete_slide(self, entry, index=-1):
        doc = entry["component"]
        page = self._slide_page(entry, int(index))
        doc.getDrawPages().remove(page)
        return {"deleted": int(index)}

    def _op_list_slides(self, entry):
        if entry["kind"] != "slide":
            raise ValueError("list_slides is a slide op")
        return self._inspect_slide(entry["component"])

    def _op_list_sheets(self, entry):
        if entry["kind"] != "sheet":
            raise ValueError("list_sheets is a sheet op")
        return self._inspect_sheet(entry["component"])


def handle(engine, request):
    cmd = request.get("cmd")
    if cmd == "ping":
        return {"pong": True, "docs": len(engine.docs)}
    if cmd == "open":
        return engine.open(request["path"], bool(request.get("readOnly")))
    if cmd == "close":
        return engine.close(request["docId"])
    if cmd == "inspect":
        return engine.inspect(request["docId"])
    if cmd == "exec":
        return engine.exec_ops(request["docId"], request.get("ops") or [])
    if cmd == "save":
        return engine.save(request["docId"], request.get("path"))
    if cmd == "convert":
        return engine.convert(request["path"], request["to"], request["outDir"])
    raise ValueError(f"unknown cmd {cmd!r}")


def main():
    parser = argparse.ArgumentParser(description="Inno Agent UNO bridge")
    parser.add_argument("--soffice", required=True)
    parser.add_argument("--port", type=int, default=2002)
    args = parser.parse_args()

    engine = OfficeEngine(args.soffice, args.port)
    try:
        engine.start()
    except Exception as err:
        log(f"engine start failed: {err}")
        send({"type": "error", "error": str(err)})
        return 1

    def on_signal(_sig, _frame):
        engine.shutdown()
        os._exit(0)

    signal.signal(signal.SIGTERM, on_signal)
    signal.signal(signal.SIGINT, on_signal)

    send({"type": "ready", "port": args.port})
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            request = json.loads(line)
        except json.JSONDecodeError:
            send({"id": None, "ok": False, "error": "invalid JSON"})
            continue
        req_id = request.get("id")
        try:
            if request.get("cmd") == "shutdown":
                send({"id": req_id, "ok": True, "result": {"bye": True}})
                break
            result = handle(engine, request)
            send({"id": req_id, "ok": True, "result": result})
        except Exception as err:
            log(f"request {req_id} failed: {traceback.format_exc(limit=3)}")
            send({"id": req_id, "ok": False, "error": str(err)})
    engine.shutdown()
    return 0


if __name__ == "__main__":
    sys.exit(main())
