#!/usr/bin/env python3
from __future__ import annotations

import json
import os
import sys
import time
from pathlib import Path
from typing import Any


def write_json(path: str | None, value: Any) -> None:
    if not path:
        return
    Path(path).write_text(json.dumps(value) + "\n", encoding="utf-8")


def append_json(path: str | None, value: Any) -> None:
    if not path:
        return
    with Path(path).open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(value) + "\n")


def value(address: str, raw: Any = 2) -> dict[str, Any]:
    sheet, _, cell = address.partition("!")
    return {
        "address": address,
        "sheet": sheet or "Sheet1",
        "row": 1,
        "col": 1,
        "colLetter": "A",
        "value": raw,
        "type": "number" if isinstance(raw, (int, float)) else "string",
        "text": str(raw),
        "visibility": "visible",
    }


def write_result() -> dict[str, Any]:
    return {
        "touched": {"Sheet1!A1": "2"},
        "changed": ["Sheet1!A1"],
        "errors": [],
        "invalidatedTiles": [],
        "updatedSheets": [],
    }


def result_for(op: str, args: dict[str, Any]) -> Any:
    if op == "utf8":
        return {"text": "Café 📈 東京"}
    if op == "listSheets":
        return {
            "sheets": [
                {"sheet": "Sheet1", "address": "Sheet1!A1:B2", "rows": 2, "cols": 2},
                {"sheet": "Hidden", "address": "Hidden!A1:A1", "rows": 1, "cols": 1, "hidden": True},
            ]
        }
    if op == "readRange":
        address = args.get("address", "Sheet1!A1")
        if isinstance(address, str) and ":" not in address:
            return [[value(address)]]
        return [[value("Sheet1!A1", 2), value("Sheet1!B1", 3)]]
    if op == "readRangeTsv":
        return "A\tB\n1\t2"
    if op in {"readColumn", "readRow"}:
        return [value("Sheet1!A1", 2), value("Sheet1!A2", 3)]
    if op in {"readColumnTsv", "readRowTsv"}:
        return "A\n1\n2"
    if op == "findCells":
        return {"matches": []}
    if op == "findRows":
        return {"matches": []}
    if op == "findAndReplace":
        return {"replaced": 1, "cells": ["Sheet1!A1"], "errors": []}
    if op == "describeSheet":
        return {"tables": {}, "structure": f"{args.get('sheet', 'Sheet1')}: nn"}
    if op == "tableLookup":
        return []
    if op == "getWorkbookProperties":
        return {"activeSheetIndex": 0, "defaultFont": {"name": "Calibri", "size": 11}}
    if op == "listDefinedNames":
        return []
    if op in {"addDefinedName", "deleteDefinedName"}:
        return {"name": args.get("name"), "range": args.get("range", "Sheet1!A1"), "scope": args.get("scope")}
    if op == "getSheetProperties":
        return {"visibility": "visible", "columns": {}, "rows": {}}
    if op == "getListObject":
        return {"name": args.get("name"), "sheet": "Sheet1", "ref": "A1:B2"}
    if op in {"addListObject", "setListObject"}:
        return {**write_result(), "listObject": {"name": args.get("name", "Table1")}}
    if op == "getDataTable":
        return {"type": "oneVariableColumn", "ref": args.get("address")}
    if op == "addDataTable":
        return {**write_result(), "dataTable": {"type": "oneVariableColumn"}}
    if op in {"deleteListObject", "deleteDataTable", "setCells"}:
        return write_result()
    if op in {"getCellPrecedents", "getCellDependents"}:
        return {"cells": [], "warnings": []}
    if op in {"traceToInputs", "traceToOutputs"}:
        return []
    if op == "sweepInputs":
        return {"tsv": "", "sweeps": [], "sweepCount": 1, "inputCount": 1, "outputCount": 1}
    if op == "evaluateFormulas":
        return [{"formula": formula, "value": 42} for formula in args.get("formulas", [])]
    if op == "lint":
        return {"diagnostics": [], "total": 0}
    if op == "previewStyles":
        return {"contentType": "image/png", "data": "AAA="}
    if op == "listCharts":
        return {"charts": []}
    if op in {"getChart", "addChart", "setChart"}:
        return {"sheet": args.get("sheet"), "chart": {"name": args.get("name", "Chart1")}}
    if op == "getConditionalFormatting":
        return {"rules": []}
    if op == "autoFitColumns":
        return {"columns": {"A": {"width": 12, "previousWidth": 8}}}
    if op == "autoFitRows":
        return {"rows": {"1": {"height": 15, "previousHeight": 12, "hidden": False, "previousHidden": False}}}
    if op == "copyRange":
        return {"destination": args.get("destination"), "cellsCopied": 4}
    if op == "reduceAddresses":
        return ["Sheet1!A1:B2"]
    if op == "getStyle":
        return {}
    if op == "save":
        save_marker = os.environ.get("WITAN_FAKE_SAVE_FILE")
        if save_marker:
            Path(save_marker).write_text("saved\n", encoding="utf-8")
        return True
    return None


def main() -> int:
    write_json(os.environ.get("WITAN_FAKE_ARGV_FILE"), sys.argv[1:])
    mode = os.environ.get("WITAN_FAKE_MODE", "ok")
    request_log = os.environ.get("WITAN_FAKE_REQUESTS_FILE")

    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            request = json.loads(line)
        except json.JSONDecodeError:
            print(json.dumps({"ok": False, "code": "INVALID_JSON", "message": "bad request"}), flush=True)
            continue
        append_json(request_log, request)

        if mode == "hang":
            time.sleep(60)
            continue
        if mode == "invalid-json":
            print("not-json", flush=True)
            continue
        if mode == "wrong-id":
            print(json.dumps({"id": "wrong", "ok": True, "result": None}), flush=True)
            continue
        if mode == "rpc-error":
            print(json.dumps({"id": request.get("id"), "ok": False, "code": "BOOM", "message": "boom"}), flush=True)
            continue
        if mode == "exit":
            return 3

        args = request.get("args")
        response = {
            "id": request.get("id"),
            "ok": True,
            "result": result_for(str(request.get("op")), args if isinstance(args, dict) else {}),
        }
        if mode == "utf8":
            raw = json.dumps(response, separators=(",", ":"), ensure_ascii=False)
            sys.stdout.buffer.write((raw + "\n").encode("utf-8"))
            sys.stdout.buffer.flush()
            continue
        print(json.dumps(response, separators=(",", ":")), flush=True)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
