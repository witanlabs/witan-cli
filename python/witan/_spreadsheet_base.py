from __future__ import annotations

import itertools
import math
import re
from collections.abc import Mapping, Sequence
from typing import Any, Literal, cast

from .types import (
    AutoFitColumnResult,
    AutoFitRowResult,
    CellRef,
    ChartInfo,
    ChartSpec,
    ChartSummary,
    ConditionalFormattingRule,
    CopyRangeResult,
    DataTable,
    DataTableMutationResult,
    DataTableSpec,
    DefinedName,
    DependencyResult,
    FindAndReplaceResult,
    FormulaResult,
    ColumnProperties,
    JsonDict,
    JsonMapping,
    LintResult,
    ListObject,
    ListObjectMutationResult,
    ListObjectSpec,
    ListObjectUpdate,
    Matcher,
    RangeRef,
    Regex,
    ReplaceMatcher,
    RowProperties,
    SearchCell,
    SearchRow,
    SetCellsValidationMode,
    SheetDescription,
    SheetInfo,
    SheetProperties,
    SheetPropertiesUpdate,
    Style,
    SweepResult,
    TableLookupResult,
    TraceInput,
    TraceOutput,
    Value,
    WorkbookProperties,
    WorkbookPropertiesUpdate,
    WriteResult,
    regex_from_pattern,
)


def drop_none(values: Mapping[str, Any]) -> JsonDict:
    return {key: value for key, value in values.items() if value is not None}


def regex_payload(value: Regex | re.Pattern[str]) -> JsonDict:
    regex = regex_from_pattern(value) if isinstance(value, re.Pattern) else value
    return {"source": regex.source, "flags": regex.flags}


def serialize_matcher(value: Matcher | ReplaceMatcher) -> Any:
    if isinstance(value, Regex) or isinstance(value, re.Pattern):
        return regex_payload(value)
    if isinstance(value, list):
        return [
            regex_payload(item) if isinstance(item, Regex) or isinstance(item, re.Pattern) else item
            for item in value
        ]
    return value


def preview_data_url(result: Mapping[str, Any]) -> str:
    content_type = result.get("contentType")
    data = result.get("data")
    if not isinstance(content_type, str) or not isinstance(data, str):
        msg = f"invalid previewStyles result: {result!r}"
        raise TypeError(msg)
    return f"data:{content_type};base64,{data}"


def is_json_number(value: Any) -> bool:
    return isinstance(value, int | float) and not isinstance(value, bool)


class _SpreadsheetSessionBase:
    def __init__(self) -> None:
        self._ids = itertools.count(1)

    def _next_id(self) -> str:
        return str(next(self._ids))

    def _request(self, method: str, op: str, args: Mapping[str, Any] | None = None) -> Any:
        raise NotImplementedError

    def get_workbook_properties(self) -> WorkbookProperties:
        return cast(WorkbookProperties, self._request("get_workbook_properties", "getWorkbookProperties", {}))

    def set_workbook_properties(self, properties: WorkbookPropertiesUpdate) -> None:
        self._request("set_workbook_properties", "setWorkbookProperties", properties)

    def list_sheets(self) -> list[SheetInfo]:
        result = cast(Mapping[str, Any], self._request("list_sheets", "listSheets", {}))
        return cast(list[SheetInfo], result.get("sheets", []))

    def get_sheet_properties(
        self,
        sheet_name: str,
        *,
        columns: Sequence[int | str] | None = None,
        rows: Sequence[int] | None = None,
    ) -> SheetProperties:
        filter_value = drop_none({"columns": list(columns) if columns is not None else None, "rows": list(rows) if rows is not None else None})
        args = {"sheet": sheet_name, **({"filter": filter_value} if filter_value else {})}
        result = cast(SheetProperties, self._request("get_sheet_properties", "getSheetProperties", args))
        result.setdefault("columns", {})
        result.setdefault("rows", {})
        return result

    def set_sheet_properties(self, sheet_name: str, properties: SheetPropertiesUpdate) -> None:
        self._request("set_sheet_properties", "setSheetProperties", {"sheet": sheet_name, "properties": properties})

    def set_row_properties(self, sheet_name: str, from_row: int, to_row: int, properties: RowProperties) -> None:
        self._request("set_row_properties", "setRowProperties", {"sheet": sheet_name, "fromRow": from_row, "toRow": to_row, "properties": properties})

    def set_column_properties(self, sheet_name: str, from_col: int | str, to_col: int | str, properties: ColumnProperties) -> None:
        self._request("set_column_properties", "setColumnProperties", {"sheet": sheet_name, "fromCol": from_col, "toCol": to_col, "properties": properties})

    def list_defined_names(self) -> list[DefinedName]:
        return cast(list[DefinedName], self._request("list_defined_names", "listDefinedNames", {}))

    def add_defined_name(self, name: str, range: str, *, scope: str | None = None) -> DefinedName:
        return cast(DefinedName, self._request("add_defined_name", "addDefinedName", drop_none({"name": name, "range": range, "scope": scope})))

    def delete_defined_name(self, name: str, *, scope: str | None = None) -> DefinedName:
        return cast(DefinedName, self._request("delete_defined_name", "deleteDefinedName", drop_none({"name": name, "scope": scope})))

    def add_sheet(self, name: str) -> str:
        self._request("add_sheet", "addSheet", {"name": name})
        return name

    def delete_sheet(self, name: str) -> None:
        self._request("delete_sheet", "deleteSheet", {"name": name})

    def rename_sheet(self, old_name: str, new_name: str) -> None:
        self._request("rename_sheet", "renameSheet", {"oldName": old_name, "newName": new_name})

    def read_cell(self, cell: CellRef, *, context: int | None = None) -> Value:
        data = cast(list[list[Value]], self._request("read_cell", "readRange", drop_none({"address": cell, "context": context})))
        return data[0][0]

    def read_range(self, range: RangeRef) -> list[list[Value]]:
        return cast(list[list[Value]], self._request("read_range", "readRange", {"address": range}))

    def read_column(self, sheet_name: str, col: int | str, *, start_row: int | None = None, end_row: int | None = None) -> list[Value]:
        return cast(list[Value], self._request("read_column", "readColumn", drop_none({"sheet": sheet_name, "col": col, "startRow": start_row, "endRow": end_row})))

    def read_row(self, sheet_name: str, row: int, *, start_col: int | None = None, end_col: int | None = None) -> list[Value]:
        return cast(list[Value], self._request("read_row", "readRow", drop_none({"sheet": sheet_name, "row": row, "startCol": start_col, "endCol": end_col})))

    def read_range_tsv(self, range: RangeRef, *, include_empty: bool | None = None, include_formulas: bool | None = None) -> str:
        return cast(str, self._request("read_range_tsv", "readRangeTsv", drop_none({"address": range, "includeEmpty": include_empty, "includeFormulas": include_formulas})))

    def read_column_tsv(self, sheet_name: str, col: int | str, *, start_row: int | None = None, end_row: int | None = None, include_empty: bool | None = None, include_formulas: bool | None = None) -> str:
        return cast(str, self._request("read_column_tsv", "readColumnTsv", drop_none({"sheet": sheet_name, "col": col, "startRow": start_row, "endRow": end_row, "includeEmpty": include_empty, "includeFormulas": include_formulas})))

    def read_row_tsv(self, sheet_name: str, row: int, *, start_col: int | None = None, end_col: int | None = None, include_empty: bool | None = None, include_formulas: bool | None = None) -> str:
        return cast(str, self._request("read_row_tsv", "readRowTsv", drop_none({"sheet": sheet_name, "row": row, "startCol": start_col, "endCol": end_col, "includeEmpty": include_empty, "includeFormulas": include_formulas})))

    def find_cells(self, matcher: Matcher, *, in_: RangeRef | None = None, context: int = 2, limit: int = 20, offset: int = 0, formulas: bool | None = None) -> list[SearchCell]:
        result = cast(Mapping[str, Any], self._request("find_cells", "findCells", drop_none({"matcher": serialize_matcher(matcher), "in": in_, "context": context, "limit": limit, "offset": offset, "formulas": formulas})))
        return cast(list[SearchCell], result.get("matches", []))

    def find_rows(self, matcher: Matcher, *, in_: RangeRef | None = None, context: int | None = None, limit: int = 20, offset: int = 0) -> list[SearchRow]:
        result = cast(Mapping[str, Any], self._request("find_rows", "findRows", drop_none({"matcher": serialize_matcher(matcher), "in": in_, "context": context, "limit": limit, "offset": offset})))
        return cast(list[SearchRow], result.get("matches", []))

    def find_and_replace(self, find: ReplaceMatcher, replace: str, *, in_: RangeRef | None = None, match_case: bool | None = None, whole_cell: bool | None = None, in_formulas: bool | None = None, limit: int | None = None) -> FindAndReplaceResult:
        return cast(FindAndReplaceResult, self._request("find_and_replace", "findAndReplace", drop_none({"find": serialize_matcher(find), "replace": replace, "in": in_, "matchCase": match_case, "wholeCell": whole_cell, "inFormulas": in_formulas, "limit": limit})))

    def describe_sheet(self, sheet_name: str) -> SheetDescription:
        return cast(SheetDescription, self._request("describe_sheet", "describeSheet", {"sheet": sheet_name}))

    def describe_sheets(self) -> dict[str, SheetDescription]:
        result: dict[str, SheetDescription] = {}
        for sheet in self.list_sheets():
            if not sheet.get("hidden"):
                result[sheet["sheet"]] = self.describe_sheet(sheet["sheet"])
        return result

    def table_lookup(self, table: str, row_label: str | int | float | bool, column_label: str | int | float | bool) -> list[TableLookupResult]:
        return cast(list[TableLookupResult], self._request("table_lookup", "tableLookup", {"table": table, "rowLabel": row_label, "columnLabel": column_label}))

    def get_list_object(self, name: str) -> ListObject:
        return cast(ListObject, self._request("get_list_object", "getListObject", {"name": name}))

    def add_list_object(self, sheet_name: str, list_object: ListObjectSpec) -> ListObjectMutationResult:
        return cast(ListObjectMutationResult, self._request("add_list_object", "addListObject", {"sheet": sheet_name, "listObject": list_object}))

    def set_list_object(self, name: str, list_object: ListObjectUpdate) -> ListObjectMutationResult:
        return cast(ListObjectMutationResult, self._request("set_list_object", "setListObject", {"name": name, "listObject": list_object}))

    def delete_list_object(self, name: str) -> WriteResult:
        return cast(WriteResult, self._request("delete_list_object", "deleteListObject", {"name": name}))

    def get_data_table(self, address: str) -> DataTable:
        return cast(DataTable, self._request("get_data_table", "getDataTable", {"address": address}))

    def add_data_table(self, sheet_name: str, data_table: DataTableSpec) -> DataTableMutationResult:
        return cast(DataTableMutationResult, self._request("add_data_table", "addDataTable", {"sheet": sheet_name, "dataTable": data_table}))

    def delete_data_table(self, address: str) -> WriteResult:
        return cast(WriteResult, self._request("delete_data_table", "deleteDataTable", {"address": address}))

    def get_cell_precedents(self, address: CellRef, depth: int | float = 1) -> DependencyResult:
        rpc_depth = int(depth) if math.isfinite(depth) else -1
        return cast(DependencyResult, self._request("get_cell_precedents", "getCellPrecedents", {"address": address, "depth": rpc_depth}))

    def get_cell_dependents(self, address: CellRef, depth: int | float = 1) -> DependencyResult:
        rpc_depth = int(depth) if math.isfinite(depth) else -1
        return cast(DependencyResult, self._request("get_cell_dependents", "getCellDependents", {"address": address, "depth": rpc_depth}))

    def trace_to_inputs(self, cell: CellRef) -> list[TraceInput]:
        return cast(list[TraceInput], self._request("trace_to_inputs", "traceToInputs", {"address": cell}))

    def trace_to_outputs(self, cell: CellRef) -> list[TraceOutput]:
        return cast(list[TraceOutput], self._request("trace_to_outputs", "traceToOutputs", {"address": cell}))

    def sweep_inputs(self, inputs: Sequence[JsonMapping], outputs: Sequence[str | CellRef], *, mode: Literal["cartesian", "parallel"] | None = None, include_stats: bool | None = None) -> SweepResult:
        return cast(SweepResult, self._request("sweep_inputs", "sweepInputs", drop_none({"inputs": list(inputs), "outputs": list(outputs), "mode": mode, "includeStats": include_stats})))

    def scenarios(self, inputs: Sequence[JsonMapping], outputs: Sequence[str | CellRef], *, mode: Literal["cartesian", "parallel"] | None = None, include_stats: bool | None = None) -> SweepResult:
        return self.sweep_inputs(inputs, outputs, mode=mode, include_stats=include_stats)

    def evaluate_formulas(self, sheet: str, formulas: Sequence[str]) -> list[FormulaResult]:
        return cast(list[FormulaResult], self._request("evaluate_formulas", "evaluateFormulas", {"sheet": sheet, "formulas": list(formulas)}))

    def evaluate_formula(self, sheet: str, formula: str) -> FormulaResult:
        return self.evaluate_formulas(sheet, [formula])[0]

    def lint(self, *, range_addresses: Sequence[str] | None = None, skip_rule_ids: Sequence[str] | None = None, only_rule_ids: Sequence[str] | None = None) -> LintResult:
        return cast(LintResult, self._request("lint", "lint", drop_none({"rangeAddresses": list(range_addresses) if range_addresses is not None else None, "skipRuleIds": list(skip_rule_ids) if skip_rule_ids is not None else None, "onlyRuleIds": list(only_rule_ids) if only_rule_ids is not None else None})))

    def preview_styles(self, range: RangeRef) -> str:
        result = cast(Mapping[str, Any], self._request("preview_styles", "previewStyles", {"address": range}))
        return preview_data_url(result)

    def list_charts(self, *, sheet: str | None = None) -> list[ChartSummary]:
        result = cast(Mapping[str, Any], self._request("list_charts", "listCharts", drop_none({"sheet": sheet})))
        return cast(list[ChartSummary], result.get("charts", []))

    def get_chart(self, sheet: str, name: str) -> ChartInfo:
        result = cast(Mapping[str, Any], self._request("get_chart", "getChart", {"sheet": sheet, "name": name}))
        return cast(ChartInfo, result.get("chart", {}))

    def add_chart(self, sheet: str, chart: ChartSpec) -> ChartSpec:
        result = cast(Mapping[str, Any], self._request("add_chart", "addChart", {"sheet": sheet, "chart": chart}))
        return cast(ChartSpec, result.get("chart", {}))

    def set_chart(self, sheet: str, name: str, chart: ChartSpec) -> ChartSpec:
        result = cast(Mapping[str, Any], self._request("set_chart", "setChart", {"sheet": sheet, "name": name, "chart": chart}))
        return cast(ChartSpec, result.get("chart", {}))

    def delete_chart(self, sheet: str, name: str) -> None:
        self._request("delete_chart", "deleteChart", {"sheet": sheet, "name": name})

    def get_conditional_formatting(self, sheet_name: str) -> list[ConditionalFormattingRule]:
        result = cast(Mapping[str, Any], self._request("get_conditional_formatting", "getConditionalFormatting", {"sheet": sheet_name}))
        return cast(list[ConditionalFormattingRule], result.get("rules", []))

    def set_conditional_formatting(self, sheet_name: str, rules: Sequence[ConditionalFormattingRule], *, clear: bool | None = None) -> None:
        self._request("set_conditional_formatting", "setConditionalFormatting", drop_none({"sheet": sheet_name, "rules": list(rules), "clear": clear}))

    def remove_conditional_formatting(self, sheet_name: str, indices: Sequence[int]) -> None:
        self._request("remove_conditional_formatting", "removeConditionalFormatting", {"sheet": sheet_name, "indices": list(indices)})

    def set_cells(
        self,
        cells: Sequence[JsonMapping],
        *,
        validation_mode: SetCellsValidationMode | None = None,
    ) -> WriteResult:
        return cast(
            WriteResult,
            self._request("set_cells", "setCells", drop_none({"cells": list(cells), "validationMode": validation_mode})),
        )

    def scale_range(self, range: RangeRef, factor: float, *, skip_formulas: bool = True) -> WriteResult | None:
        data = self.read_range(range)
        assignments: list[JsonDict] = []
        for row in data:
            for cell in row:
                value = cell.get("value")
                has_formula = bool(cell.get("formula"))
                if is_json_number(value) and (not has_formula or not skip_formulas):
                    assignments.append({"address": cell["address"], "value": value * factor})
        if not assignments:
            return None
        return self.set_cells(assignments)

    def insert_row_after(self, sheet_name: str, row: int, count: int = 1) -> None:
        self._request("insert_row_after", "insertRowAfter", {"sheet": sheet_name, "row": row, "count": count})

    def delete_rows(self, sheet_name: str, row: int, count: int = 1) -> None:
        self._request("delete_rows", "deleteRows", {"sheet": sheet_name, "row": row, "count": count})

    def insert_column_after(self, sheet_name: str, column: int | str, count: int = 1) -> None:
        self._request("insert_column_after", "insertColumnAfter", {"sheet": sheet_name, "column": column, "count": count})

    def delete_columns(self, sheet_name: str, column: int | str, count: int = 1) -> None:
        self._request("delete_columns", "deleteColumns", {"sheet": sheet_name, "column": column, "count": count})

    def auto_fit_columns(self, sheet_name: str, columns: Sequence[int | str] | None = None, *, min_width: float | None = None, max_width: float | None = None, padding: float | None = None) -> dict[str, AutoFitColumnResult]:
        result = cast(Mapping[str, Any], self._request("auto_fit_columns", "autoFitColumns", drop_none({"sheet": sheet_name, "columns": list(columns) if columns is not None else None, "minWidth": min_width, "maxWidth": max_width, "padding": padding})))
        return cast(dict[str, AutoFitColumnResult], result.get("columns", {}))

    def auto_fit_rows(self, sheet_name: str, rows: Sequence[int] | None = None, *, min_height: float | None = None, max_height: float | None = None) -> dict[str, AutoFitRowResult]:
        result = cast(Mapping[str, Any], self._request("auto_fit_rows", "autoFitRows", drop_none({"sheet": sheet_name, "rows": list(rows) if rows is not None else None, "minHeight": min_height, "maxHeight": max_height})))
        return cast(dict[str, AutoFitRowResult], result.get("rows", {}))

    def sort_range(self, range: RangeRef, keys: Sequence[JsonMapping], *, has_header: bool | None = None) -> None:
        self._request("sort_range", "sortRange", drop_none({"range": range, "keys": list(keys), "hasHeader": has_header}))

    def copy_range(self, source: RangeRef, destination: CellRef, *, paste_type: Literal["all", "values", "formulas", "formats"] | None = None) -> CopyRangeResult:
        return cast(CopyRangeResult, self._request("copy_range", "copyRange", drop_none({"source": source, "destination": destination, "pasteType": paste_type})))

    def reduce_addresses(self, addresses: Sequence[CellRef | RangeRef]) -> list[str]:
        return cast(list[str], self._request("reduce_addresses", "reduceAddresses", {"addresses": list(addresses)}))

    def get_style(self, cell: CellRef) -> Style:
        return cast(Style, self._request("get_style", "getStyle", {"address": cell}))

    def set_style(self, target: CellRef | RangeRef, style: Style) -> None:
        self._request("set_style", "setStyle", {"address": target, "style": style})


class _AsyncSpreadsheetSessionBase:
    def __init__(self) -> None:
        self._ids = itertools.count(1)

    def _next_id(self) -> str:
        return str(next(self._ids))

    async def _request(self, method: str, op: str, args: Mapping[str, Any] | None = None) -> Any:
        raise NotImplementedError

    async def get_workbook_properties(self) -> WorkbookProperties:
        return cast(WorkbookProperties, await self._request("get_workbook_properties", "getWorkbookProperties", {}))

    async def set_workbook_properties(self, properties: WorkbookPropertiesUpdate) -> None:
        await self._request("set_workbook_properties", "setWorkbookProperties", properties)

    async def list_sheets(self) -> list[SheetInfo]:
        result = cast(Mapping[str, Any], await self._request("list_sheets", "listSheets", {}))
        return cast(list[SheetInfo], result.get("sheets", []))

    async def get_sheet_properties(
        self,
        sheet_name: str,
        *,
        columns: Sequence[int | str] | None = None,
        rows: Sequence[int] | None = None,
    ) -> SheetProperties:
        filter_value = drop_none({"columns": list(columns) if columns is not None else None, "rows": list(rows) if rows is not None else None})
        args = {"sheet": sheet_name, **({"filter": filter_value} if filter_value else {})}
        result = cast(SheetProperties, await self._request("get_sheet_properties", "getSheetProperties", args))
        result.setdefault("columns", {})
        result.setdefault("rows", {})
        return result

    async def set_sheet_properties(self, sheet_name: str, properties: SheetPropertiesUpdate) -> None:
        await self._request("set_sheet_properties", "setSheetProperties", {"sheet": sheet_name, "properties": properties})

    async def set_row_properties(self, sheet_name: str, from_row: int, to_row: int, properties: RowProperties) -> None:
        await self._request("set_row_properties", "setRowProperties", {"sheet": sheet_name, "fromRow": from_row, "toRow": to_row, "properties": properties})

    async def set_column_properties(self, sheet_name: str, from_col: int | str, to_col: int | str, properties: ColumnProperties) -> None:
        await self._request("set_column_properties", "setColumnProperties", {"sheet": sheet_name, "fromCol": from_col, "toCol": to_col, "properties": properties})

    async def list_defined_names(self) -> list[DefinedName]:
        return cast(list[DefinedName], await self._request("list_defined_names", "listDefinedNames", {}))

    async def add_defined_name(self, name: str, range: str, *, scope: str | None = None) -> DefinedName:
        return cast(DefinedName, await self._request("add_defined_name", "addDefinedName", drop_none({"name": name, "range": range, "scope": scope})))

    async def delete_defined_name(self, name: str, *, scope: str | None = None) -> DefinedName:
        return cast(DefinedName, await self._request("delete_defined_name", "deleteDefinedName", drop_none({"name": name, "scope": scope})))

    async def add_sheet(self, name: str) -> str:
        await self._request("add_sheet", "addSheet", {"name": name})
        return name

    async def delete_sheet(self, name: str) -> None:
        await self._request("delete_sheet", "deleteSheet", {"name": name})

    async def rename_sheet(self, old_name: str, new_name: str) -> None:
        await self._request("rename_sheet", "renameSheet", {"oldName": old_name, "newName": new_name})

    async def read_cell(self, cell: CellRef, *, context: int | None = None) -> Value:
        data = cast(list[list[Value]], await self._request("read_cell", "readRange", drop_none({"address": cell, "context": context})))
        return data[0][0]

    async def read_range(self, range: RangeRef) -> list[list[Value]]:
        return cast(list[list[Value]], await self._request("read_range", "readRange", {"address": range}))

    async def read_column(self, sheet_name: str, col: int | str, *, start_row: int | None = None, end_row: int | None = None) -> list[Value]:
        return cast(list[Value], await self._request("read_column", "readColumn", drop_none({"sheet": sheet_name, "col": col, "startRow": start_row, "endRow": end_row})))

    async def read_row(self, sheet_name: str, row: int, *, start_col: int | None = None, end_col: int | None = None) -> list[Value]:
        return cast(list[Value], await self._request("read_row", "readRow", drop_none({"sheet": sheet_name, "row": row, "startCol": start_col, "endCol": end_col})))

    async def read_range_tsv(self, range: RangeRef, *, include_empty: bool | None = None, include_formulas: bool | None = None) -> str:
        return cast(str, await self._request("read_range_tsv", "readRangeTsv", drop_none({"address": range, "includeEmpty": include_empty, "includeFormulas": include_formulas})))

    async def read_column_tsv(self, sheet_name: str, col: int | str, *, start_row: int | None = None, end_row: int | None = None, include_empty: bool | None = None, include_formulas: bool | None = None) -> str:
        return cast(str, await self._request("read_column_tsv", "readColumnTsv", drop_none({"sheet": sheet_name, "col": col, "startRow": start_row, "endRow": end_row, "includeEmpty": include_empty, "includeFormulas": include_formulas})))

    async def read_row_tsv(self, sheet_name: str, row: int, *, start_col: int | None = None, end_col: int | None = None, include_empty: bool | None = None, include_formulas: bool | None = None) -> str:
        return cast(str, await self._request("read_row_tsv", "readRowTsv", drop_none({"sheet": sheet_name, "row": row, "startCol": start_col, "endCol": end_col, "includeEmpty": include_empty, "includeFormulas": include_formulas})))

    async def find_cells(self, matcher: Matcher, *, in_: RangeRef | None = None, context: int = 2, limit: int = 20, offset: int = 0, formulas: bool | None = None) -> list[SearchCell]:
        result = cast(Mapping[str, Any], await self._request("find_cells", "findCells", drop_none({"matcher": serialize_matcher(matcher), "in": in_, "context": context, "limit": limit, "offset": offset, "formulas": formulas})))
        return cast(list[SearchCell], result.get("matches", []))

    async def find_rows(self, matcher: Matcher, *, in_: RangeRef | None = None, context: int | None = None, limit: int = 20, offset: int = 0) -> list[SearchRow]:
        result = cast(Mapping[str, Any], await self._request("find_rows", "findRows", drop_none({"matcher": serialize_matcher(matcher), "in": in_, "context": context, "limit": limit, "offset": offset})))
        return cast(list[SearchRow], result.get("matches", []))

    async def find_and_replace(self, find: ReplaceMatcher, replace: str, *, in_: RangeRef | None = None, match_case: bool | None = None, whole_cell: bool | None = None, in_formulas: bool | None = None, limit: int | None = None) -> FindAndReplaceResult:
        return cast(FindAndReplaceResult, await self._request("find_and_replace", "findAndReplace", drop_none({"find": serialize_matcher(find), "replace": replace, "in": in_, "matchCase": match_case, "wholeCell": whole_cell, "inFormulas": in_formulas, "limit": limit})))

    async def describe_sheet(self, sheet_name: str) -> SheetDescription:
        return cast(SheetDescription, await self._request("describe_sheet", "describeSheet", {"sheet": sheet_name}))

    async def describe_sheets(self) -> dict[str, SheetDescription]:
        result: dict[str, SheetDescription] = {}
        for sheet in await self.list_sheets():
            if not sheet.get("hidden"):
                result[sheet["sheet"]] = await self.describe_sheet(sheet["sheet"])
        return result

    async def table_lookup(self, table: str, row_label: str | int | float | bool, column_label: str | int | float | bool) -> list[TableLookupResult]:
        return cast(list[TableLookupResult], await self._request("table_lookup", "tableLookup", {"table": table, "rowLabel": row_label, "columnLabel": column_label}))

    async def get_list_object(self, name: str) -> ListObject:
        return cast(ListObject, await self._request("get_list_object", "getListObject", {"name": name}))

    async def add_list_object(self, sheet_name: str, list_object: ListObjectSpec) -> ListObjectMutationResult:
        return cast(ListObjectMutationResult, await self._request("add_list_object", "addListObject", {"sheet": sheet_name, "listObject": list_object}))

    async def set_list_object(self, name: str, list_object: ListObjectUpdate) -> ListObjectMutationResult:
        return cast(ListObjectMutationResult, await self._request("set_list_object", "setListObject", {"name": name, "listObject": list_object}))

    async def delete_list_object(self, name: str) -> WriteResult:
        return cast(WriteResult, await self._request("delete_list_object", "deleteListObject", {"name": name}))

    async def get_data_table(self, address: str) -> DataTable:
        return cast(DataTable, await self._request("get_data_table", "getDataTable", {"address": address}))

    async def add_data_table(self, sheet_name: str, data_table: DataTableSpec) -> DataTableMutationResult:
        return cast(DataTableMutationResult, await self._request("add_data_table", "addDataTable", {"sheet": sheet_name, "dataTable": data_table}))

    async def delete_data_table(self, address: str) -> WriteResult:
        return cast(WriteResult, await self._request("delete_data_table", "deleteDataTable", {"address": address}))

    async def get_cell_precedents(self, address: CellRef, depth: int | float = 1) -> DependencyResult:
        rpc_depth = int(depth) if math.isfinite(depth) else -1
        return cast(DependencyResult, await self._request("get_cell_precedents", "getCellPrecedents", {"address": address, "depth": rpc_depth}))

    async def get_cell_dependents(self, address: CellRef, depth: int | float = 1) -> DependencyResult:
        rpc_depth = int(depth) if math.isfinite(depth) else -1
        return cast(DependencyResult, await self._request("get_cell_dependents", "getCellDependents", {"address": address, "depth": rpc_depth}))

    async def trace_to_inputs(self, cell: CellRef) -> list[TraceInput]:
        return cast(list[TraceInput], await self._request("trace_to_inputs", "traceToInputs", {"address": cell}))

    async def trace_to_outputs(self, cell: CellRef) -> list[TraceOutput]:
        return cast(list[TraceOutput], await self._request("trace_to_outputs", "traceToOutputs", {"address": cell}))

    async def sweep_inputs(self, inputs: Sequence[JsonMapping], outputs: Sequence[str | CellRef], *, mode: Literal["cartesian", "parallel"] | None = None, include_stats: bool | None = None) -> SweepResult:
        return cast(SweepResult, await self._request("sweep_inputs", "sweepInputs", drop_none({"inputs": list(inputs), "outputs": list(outputs), "mode": mode, "includeStats": include_stats})))

    async def scenarios(self, inputs: Sequence[JsonMapping], outputs: Sequence[str | CellRef], *, mode: Literal["cartesian", "parallel"] | None = None, include_stats: bool | None = None) -> SweepResult:
        return await self.sweep_inputs(inputs, outputs, mode=mode, include_stats=include_stats)

    async def evaluate_formulas(self, sheet: str, formulas: Sequence[str]) -> list[FormulaResult]:
        return cast(list[FormulaResult], await self._request("evaluate_formulas", "evaluateFormulas", {"sheet": sheet, "formulas": list(formulas)}))

    async def evaluate_formula(self, sheet: str, formula: str) -> FormulaResult:
        return (await self.evaluate_formulas(sheet, [formula]))[0]

    async def lint(self, *, range_addresses: Sequence[str] | None = None, skip_rule_ids: Sequence[str] | None = None, only_rule_ids: Sequence[str] | None = None) -> LintResult:
        return cast(LintResult, await self._request("lint", "lint", drop_none({"rangeAddresses": list(range_addresses) if range_addresses is not None else None, "skipRuleIds": list(skip_rule_ids) if skip_rule_ids is not None else None, "onlyRuleIds": list(only_rule_ids) if only_rule_ids is not None else None})))

    async def preview_styles(self, range: RangeRef) -> str:
        result = cast(Mapping[str, Any], await self._request("preview_styles", "previewStyles", {"address": range}))
        return preview_data_url(result)

    async def list_charts(self, *, sheet: str | None = None) -> list[ChartSummary]:
        result = cast(Mapping[str, Any], await self._request("list_charts", "listCharts", drop_none({"sheet": sheet})))
        return cast(list[ChartSummary], result.get("charts", []))

    async def get_chart(self, sheet: str, name: str) -> ChartInfo:
        result = cast(Mapping[str, Any], await self._request("get_chart", "getChart", {"sheet": sheet, "name": name}))
        return cast(ChartInfo, result.get("chart", {}))

    async def add_chart(self, sheet: str, chart: ChartSpec) -> ChartSpec:
        result = cast(Mapping[str, Any], await self._request("add_chart", "addChart", {"sheet": sheet, "chart": chart}))
        return cast(ChartSpec, result.get("chart", {}))

    async def set_chart(self, sheet: str, name: str, chart: ChartSpec) -> ChartSpec:
        result = cast(Mapping[str, Any], await self._request("set_chart", "setChart", {"sheet": sheet, "name": name, "chart": chart}))
        return cast(ChartSpec, result.get("chart", {}))

    async def delete_chart(self, sheet: str, name: str) -> None:
        await self._request("delete_chart", "deleteChart", {"sheet": sheet, "name": name})

    async def get_conditional_formatting(self, sheet_name: str) -> list[ConditionalFormattingRule]:
        result = cast(Mapping[str, Any], await self._request("get_conditional_formatting", "getConditionalFormatting", {"sheet": sheet_name}))
        return cast(list[ConditionalFormattingRule], result.get("rules", []))

    async def set_conditional_formatting(self, sheet_name: str, rules: Sequence[ConditionalFormattingRule], *, clear: bool | None = None) -> None:
        await self._request("set_conditional_formatting", "setConditionalFormatting", drop_none({"sheet": sheet_name, "rules": list(rules), "clear": clear}))

    async def remove_conditional_formatting(self, sheet_name: str, indices: Sequence[int]) -> None:
        await self._request("remove_conditional_formatting", "removeConditionalFormatting", {"sheet": sheet_name, "indices": list(indices)})

    async def set_cells(
        self,
        cells: Sequence[JsonMapping],
        *,
        validation_mode: SetCellsValidationMode | None = None,
    ) -> WriteResult:
        return cast(
            WriteResult,
            await self._request("set_cells", "setCells", drop_none({"cells": list(cells), "validationMode": validation_mode})),
        )

    async def scale_range(self, range: RangeRef, factor: float, *, skip_formulas: bool = True) -> WriteResult | None:
        data = await self.read_range(range)
        assignments: list[JsonDict] = []
        for row in data:
            for cell in row:
                value = cell.get("value")
                has_formula = bool(cell.get("formula"))
                if is_json_number(value) and (not has_formula or not skip_formulas):
                    assignments.append({"address": cell["address"], "value": value * factor})
        if not assignments:
            return None
        return await self.set_cells(assignments)

    async def insert_row_after(self, sheet_name: str, row: int, count: int = 1) -> None:
        await self._request("insert_row_after", "insertRowAfter", {"sheet": sheet_name, "row": row, "count": count})

    async def delete_rows(self, sheet_name: str, row: int, count: int = 1) -> None:
        await self._request("delete_rows", "deleteRows", {"sheet": sheet_name, "row": row, "count": count})

    async def insert_column_after(self, sheet_name: str, column: int | str, count: int = 1) -> None:
        await self._request("insert_column_after", "insertColumnAfter", {"sheet": sheet_name, "column": column, "count": count})

    async def delete_columns(self, sheet_name: str, column: int | str, count: int = 1) -> None:
        await self._request("delete_columns", "deleteColumns", {"sheet": sheet_name, "column": column, "count": count})

    async def auto_fit_columns(self, sheet_name: str, columns: Sequence[int | str] | None = None, *, min_width: float | None = None, max_width: float | None = None, padding: float | None = None) -> dict[str, AutoFitColumnResult]:
        result = cast(Mapping[str, Any], await self._request("auto_fit_columns", "autoFitColumns", drop_none({"sheet": sheet_name, "columns": list(columns) if columns is not None else None, "minWidth": min_width, "maxWidth": max_width, "padding": padding})))
        return cast(dict[str, AutoFitColumnResult], result.get("columns", {}))

    async def auto_fit_rows(self, sheet_name: str, rows: Sequence[int] | None = None, *, min_height: float | None = None, max_height: float | None = None) -> dict[str, AutoFitRowResult]:
        result = cast(Mapping[str, Any], await self._request("auto_fit_rows", "autoFitRows", drop_none({"sheet": sheet_name, "rows": list(rows) if rows is not None else None, "minHeight": min_height, "maxHeight": max_height})))
        return cast(dict[str, AutoFitRowResult], result.get("rows", {}))

    async def sort_range(self, range: RangeRef, keys: Sequence[JsonMapping], *, has_header: bool | None = None) -> None:
        await self._request("sort_range", "sortRange", drop_none({"range": range, "keys": list(keys), "hasHeader": has_header}))

    async def copy_range(self, source: RangeRef, destination: CellRef, *, paste_type: Literal["all", "values", "formulas", "formats"] | None = None) -> CopyRangeResult:
        return cast(CopyRangeResult, await self._request("copy_range", "copyRange", drop_none({"source": source, "destination": destination, "pasteType": paste_type})))

    async def reduce_addresses(self, addresses: Sequence[CellRef | RangeRef]) -> list[str]:
        return cast(list[str], await self._request("reduce_addresses", "reduceAddresses", {"addresses": list(addresses)}))

    async def get_style(self, cell: CellRef) -> Style:
        return cast(Style, await self._request("get_style", "getStyle", {"address": cell}))

    async def set_style(self, target: CellRef | RangeRef, style: Style) -> None:
        await self._request("set_style", "setStyle", {"address": target, "style": style})