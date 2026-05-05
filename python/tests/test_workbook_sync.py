from __future__ import annotations

import json
import re
import stat
import subprocess
from pathlib import Path

import pytest

from witan import Regex, Workbook, WitanProcessError, WitanRPCError, WitanTimeoutError


def fake_binary() -> Path:
    path = Path(__file__).with_name("fake_witan_rpc.py")
    path.chmod(path.stat().st_mode | stat.S_IXUSR)
    return path


def json_lines(path: Path) -> list[dict[str, object]]:
    return [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines() if line]


def fake_env(tmp_path: Path, *, mode: str = "ok") -> tuple[dict[str, str], Path, Path, Path]:
    argv_file = tmp_path / "argv.jsonl"
    requests_file = tmp_path / "requests.jsonl"
    save_file = tmp_path / "saved.txt"
    env = {
        "WITAN_FAKE_ARGV_FILE": str(argv_file),
        "WITAN_FAKE_REQUESTS_FILE": str(requests_file),
        "WITAN_FAKE_SAVE_FILE": str(save_file),
        "WITAN_FAKE_MODE": mode,
    }
    return env, argv_file, requests_file, save_file


def test_workbook_invokes_witan_xlsx_rpc_and_maps_pythonic_methods(tmp_path: Path) -> None:
    env, argv_file, requests_file, save_file = fake_env(tmp_path)
    workbook_path = tmp_path / "created.xlsx"

    with Workbook(
        workbook_path,
        create=True,
        stateless=True,
        hint="Sheet1!A1:B2",
        locale="en-US",
        binary=fake_binary(),
        env=env,
    ) as wb:
        assert wb.read_range_tsv("Sheet1!A1:B2") == "A\tB\n1\t2"
        assert wb.read_cell("Sheet1!A1")["value"] == 2
        assert wb.list_sheets()[0]["sheet"] == "Sheet1"
        assert wb.describe_sheets() == {"Sheet1": {"tables": {}, "structure": "Sheet1: nn"}}
        assert wb.find_cells(Regex("rev", "i"), in_="Sheet1!A:A") == []
        assert wb.find_cells(re.compile("cost", re.I), formulas=True) == []
        assert wb.find_and_replace(Regex("old"), "new")["replaced"] == 1
        assert wb.preview_styles("Sheet1!A1") == "data:image/png;base64,AAA="
        assert wb.scenarios([{"address": "Sheet1!A1", "values": [1]}], ["Sheet1!B1"])["sweepCount"] == 1
        assert wb.reduce_addresses(["Sheet1!A:B"]) == ["Sheet1!A1:B2"]
        assert wb.copy_range("Sheet1!A1:B2", "Sheet1!C1")["cellsCopied"] == 4
        assert wb.auto_fit_columns("Sheet1")["A"]["width"] == 12
        assert not hasattr(wb, "call")
        assert wb.save() is True

    assert save_file.read_text(encoding="utf-8") == "saved\n"

    argv = json.loads(argv_file.read_text(encoding="utf-8"))
    assert argv[:2] == ["--stateless", "xlsx"]
    assert argv[2:5] == ["rpc", str(workbook_path), "--create"]
    assert "--hint" in argv
    assert "--locale" in argv

    requests = json_lines(requests_file)
    ops_and_args = [(request["op"], request["args"]) for request in requests]
    assert ops_and_args[0] == (
        "readRangeTsv",
        {"address": "Sheet1!A1:B2"},
    )
    assert ops_and_args[1] == (
        "readRange",
        {"address": "Sheet1!A1"},
    )
    assert ("findCells", {"matcher": {"source": "rev", "flags": "i"}, "in": "Sheet1!A:A", "context": 2, "limit": 20, "offset": 0}) in ops_and_args
    assert ("reduceAddresses", {"addresses": ["Sheet1!A:B"]}) in ops_and_args
    assert requests[-1]["op"] == "save"


def test_workbook_sync_transport_uses_utf8_bytes(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    real_popen = subprocess.Popen
    popen_kwargs: dict[str, object] = {}

    def popen(*args: object, **kwargs: object) -> subprocess.Popen[bytes]:
        popen_kwargs.update(kwargs)
        return real_popen(*args, **kwargs)

    monkeypatch.setattr("witan._process.subprocess.Popen", popen)

    env, _, _, _ = fake_env(tmp_path, mode="utf8")
    wb = Workbook(tmp_path / "book.xlsx", binary=fake_binary(), env=env)
    try:
        assert wb._request("utf8", "utf8", {}) == {"text": "Café 📈 東京"}
    finally:
        wb.close()

    assert "text" not in popen_kwargs
    assert "encoding" not in popen_kwargs


def test_all_sync_operation_wrappers_emit_documented_rpc_ops(tmp_path: Path) -> None:
    env, _, requests_file, _ = fake_env(tmp_path)

    with Workbook(tmp_path / "book.xlsx", binary=fake_binary(), env=env) as wb:
        wb.get_workbook_properties()
        wb.set_workbook_properties({"metadata": {"title": "Plan"}})
        wb.list_sheets()
        wb.get_sheet_properties("Sheet1", columns=["A"], rows=[1])
        wb.set_sheet_properties("Sheet1", {"visibility": "visible"})
        wb.set_row_properties("Sheet1", 1, 2, {"height": 12})
        wb.set_column_properties("Sheet1", "A", "B", {"width": 10})
        wb.list_defined_names()
        wb.add_defined_name("Input", "Sheet1!A1")
        wb.delete_defined_name("Input")
        wb.add_sheet("New")
        wb.delete_sheet("New")
        wb.rename_sheet("Old", "New")
        wb.read_cell("Sheet1!A1")
        wb.read_range("Sheet1!A1:B2")
        wb.read_column("Sheet1", "A", start_row=1, end_row=2)
        wb.read_row("Sheet1", 1, start_col=1, end_col=2)
        wb.read_range_tsv("Sheet1!A1:B2", include_empty=True, include_formulas=True)
        wb.read_column_tsv("Sheet1", "A")
        wb.read_row_tsv("Sheet1", 1)
        wb.find_cells("Revenue")
        wb.find_rows("Revenue")
        wb.find_and_replace("old", "new")
        wb.describe_sheet("Sheet1")
        wb.describe_sheets()
        wb.table_lookup("Sheet1!A1:B2", "row", "col")
        wb.get_list_object("Table1")
        wb.add_list_object("Sheet1", {"name": "Table1", "ref": "A1:B2", "columns": []})
        wb.set_list_object("Table1", {"showTotalsRow": True})
        wb.delete_list_object("Table1")
        wb.get_data_table("Sheet1!E1:F4")
        wb.add_data_table("Sheet1", {"type": "oneVariableColumn"})
        wb.delete_data_table("Sheet1!E1:F4")
        wb.get_cell_precedents("Sheet1!B1")
        wb.get_cell_dependents("Sheet1!A1", depth=float("inf"))
        wb.trace_to_inputs("Sheet1!B1")
        wb.trace_to_outputs("Sheet1!A1")
        wb.sweep_inputs([{"address": "Sheet1!A1", "values": [1]}], ["Sheet1!B1"])
        wb.scenarios([{"address": "Sheet1!A1", "values": [1]}], ["Sheet1!B1"])
        wb.evaluate_formulas("Sheet1", ["=1+1"])
        wb.evaluate_formula("Sheet1", "=1+1")
        wb.lint(range_addresses=["Sheet1!A1:B2"], skip_rule_ids=["D003"])
        wb.preview_styles("Sheet1!A1:B2")
        wb.list_charts(sheet="Sheet1")
        wb.get_chart("Sheet1", "Chart1")
        wb.add_chart("Sheet1", {"name": "Chart1"})
        wb.set_chart("Sheet1", "Chart1", {"name": "Chart1"})
        wb.delete_chart("Sheet1", "Chart1")
        wb.get_conditional_formatting("Sheet1")
        wb.set_conditional_formatting("Sheet1", [{"type": "expression", "address": "A1", "formula": "TRUE"}], clear=True)
        wb.remove_conditional_formatting("Sheet1", [0])
        wb.set_cells([{"address": "Sheet1!A1", "value": 1}])
        wb.scale_range("Sheet1!A1:B1", 2)
        wb.insert_row_after("Sheet1", 1)
        wb.delete_rows("Sheet1", 2)
        wb.insert_column_after("Sheet1", "A")
        wb.delete_columns("Sheet1", "B")
        wb.auto_fit_columns("Sheet1", ["A"])
        wb.auto_fit_rows("Sheet1", [1])
        wb.sort_range("Sheet1!A1:B2", [{"col": "A", "order": "asc"}], has_header=True)
        wb.copy_range("Sheet1!A1:B2", "Sheet1!C1", paste_type="values")
        wb.reduce_addresses(["Sheet1!A:B"])
        wb.get_style("Sheet1!A1")
        wb.set_style("Sheet1!A1", {"font": {"bold": True}})
        wb.save()

    assert [request["op"] for request in json_lines(requests_file)] == [
        "getWorkbookProperties",
        "setWorkbookProperties",
        "listSheets",
        "getSheetProperties",
        "setSheetProperties",
        "setRowProperties",
        "setColumnProperties",
        "listDefinedNames",
        "addDefinedName",
        "deleteDefinedName",
        "addSheet",
        "deleteSheet",
        "renameSheet",
        "readRange",
        "readRange",
        "readColumn",
        "readRow",
        "readRangeTsv",
        "readColumnTsv",
        "readRowTsv",
        "findCells",
        "findRows",
        "findAndReplace",
        "describeSheet",
        "listSheets",
        "describeSheet",
        "tableLookup",
        "getListObject",
        "addListObject",
        "setListObject",
        "deleteListObject",
        "getDataTable",
        "addDataTable",
        "deleteDataTable",
        "getCellPrecedents",
        "getCellDependents",
        "traceToInputs",
        "traceToOutputs",
        "sweepInputs",
        "sweepInputs",
        "evaluateFormulas",
        "evaluateFormulas",
        "lint",
        "previewStyles",
        "listCharts",
        "getChart",
        "addChart",
        "setChart",
        "deleteChart",
        "getConditionalFormatting",
        "setConditionalFormatting",
        "removeConditionalFormatting",
        "setCells",
        "readRange",
        "setCells",
        "insertRowAfter",
        "deleteRows",
        "insertColumnAfter",
        "deleteColumns",
        "autoFitColumns",
        "autoFitRows",
        "sortRange",
        "copyRange",
        "reduceAddresses",
        "getStyle",
        "setStyle",
        "save",
    ]


def test_workbook_raises_rpc_error(tmp_path: Path) -> None:
    env, _, _, _ = fake_env(tmp_path, mode="rpc-error")
    wb = Workbook(tmp_path / "book.xlsx", binary=fake_binary(), env=env)
    try:
        with pytest.raises(WitanRPCError) as raised:
            wb.list_sheets()
    finally:
        wb.close()
    assert raised.value.code == "BOOM"
    assert raised.value.method == "list_sheets"
    assert raised.value.op == "listSheets"


def test_workbook_timeout_terminates_session(tmp_path: Path) -> None:
    env, _, _, _ = fake_env(tmp_path, mode="hang")
    wb = Workbook(tmp_path / "book.xlsx", binary=fake_binary(), env=env, request_timeout=0.05)
    with pytest.raises(WitanTimeoutError):
        wb.list_sheets()
    assert wb._process.closed


def test_workbook_rejects_bad_process_responses(tmp_path: Path) -> None:
    env, _, _, _ = fake_env(tmp_path, mode="invalid-json")
    wb = Workbook(tmp_path / "book.xlsx", binary=fake_binary(), env=env)
    try:
        with pytest.raises(WitanProcessError):
            wb.list_sheets()
    finally:
        wb.close()

    env, _, _, _ = fake_env(tmp_path, mode="wrong-id")
    wb = Workbook(tmp_path / "book.xlsx", binary=fake_binary(), env=env)
    try:
        with pytest.raises(WitanProcessError):
            wb.list_sheets()
    finally:
        wb.close()
