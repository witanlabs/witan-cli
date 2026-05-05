from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class OperationSpec:
    method: str
    op: str
    params: tuple[str, ...] = ()
    options: tuple[str, ...] = ()
    composite: bool = False


OPERATIONS: tuple[OperationSpec, ...] = (
    OperationSpec("get_workbook_properties", "getWorkbookProperties"),
    OperationSpec("set_workbook_properties", "setWorkbookProperties", ("properties",)),
    OperationSpec("list_sheets", "listSheets"),
    OperationSpec("get_sheet_properties", "getSheetProperties", ("sheet",), ("filter",)),
    OperationSpec("set_sheet_properties", "setSheetProperties", ("sheet", "properties")),
    OperationSpec("set_row_properties", "setRowProperties", ("sheet", "fromRow", "toRow", "properties")),
    OperationSpec("set_column_properties", "setColumnProperties", ("sheet", "fromCol", "toCol", "properties")),
    OperationSpec("list_defined_names", "listDefinedNames"),
    OperationSpec("add_defined_name", "addDefinedName", ("name", "range"), ("scope",)),
    OperationSpec("delete_defined_name", "deleteDefinedName", ("name",), ("scope",)),
    OperationSpec("add_sheet", "addSheet", ("name",)),
    OperationSpec("delete_sheet", "deleteSheet", ("name",)),
    OperationSpec("rename_sheet", "renameSheet", ("oldName", "newName")),
    OperationSpec("read_cell", "readRange", ("address",), ("context",), composite=True),
    OperationSpec("read_range", "readRange", ("address",)),
    OperationSpec("read_column", "readColumn", ("sheet", "col"), ("startRow", "endRow")),
    OperationSpec("read_row", "readRow", ("sheet", "row"), ("startCol", "endCol")),
    OperationSpec("read_range_tsv", "readRangeTsv", ("address",), ("includeEmpty", "includeFormulas")),
    OperationSpec("read_column_tsv", "readColumnTsv", ("sheet", "col"), ("startRow", "endRow", "includeEmpty", "includeFormulas")),
    OperationSpec("read_row_tsv", "readRowTsv", ("sheet", "row"), ("startCol", "endCol", "includeEmpty", "includeFormulas")),
    OperationSpec("find_cells", "findCells", ("matcher",), ("in", "context", "limit", "offset", "formulas")),
    OperationSpec("find_rows", "findRows", ("matcher",), ("in", "context", "limit", "offset")),
    OperationSpec("find_and_replace", "findAndReplace", ("find", "replace"), ("in", "matchCase", "wholeCell", "inFormulas", "limit")),
    OperationSpec("describe_sheet", "describeSheet", ("sheet",)),
    OperationSpec("describe_sheets", "describeSheet", composite=True),
    OperationSpec("table_lookup", "tableLookup", ("table", "rowLabel", "columnLabel")),
    OperationSpec("get_list_object", "getListObject", ("name",)),
    OperationSpec("add_list_object", "addListObject", ("sheet", "listObject")),
    OperationSpec("set_list_object", "setListObject", ("name", "listObject")),
    OperationSpec("delete_list_object", "deleteListObject", ("name",)),
    OperationSpec("get_data_table", "getDataTable", ("address",)),
    OperationSpec("add_data_table", "addDataTable", ("sheet", "dataTable")),
    OperationSpec("delete_data_table", "deleteDataTable", ("address",)),
    OperationSpec("get_cell_precedents", "getCellPrecedents", ("address", "depth")),
    OperationSpec("get_cell_dependents", "getCellDependents", ("address", "depth")),
    OperationSpec("trace_to_inputs", "traceToInputs", ("address",)),
    OperationSpec("trace_to_outputs", "traceToOutputs", ("address",)),
    OperationSpec("sweep_inputs", "sweepInputs", ("inputs", "outputs"), ("mode", "includeStats")),
    OperationSpec("scenarios", "sweepInputs", ("inputs", "outputs"), ("mode", "includeStats"), composite=True),
    OperationSpec("evaluate_formulas", "evaluateFormulas", ("sheet", "formulas")),
    OperationSpec("evaluate_formula", "evaluateFormulas", ("sheet", "formulas"), composite=True),
    OperationSpec("lint", "lint", (), ("rangeAddresses", "skipRuleIds", "onlyRuleIds")),
    OperationSpec("preview_styles", "previewStyles", ("address",)),
    OperationSpec("list_charts", "listCharts", (), ("sheet",)),
    OperationSpec("get_chart", "getChart", ("sheet", "name")),
    OperationSpec("add_chart", "addChart", ("sheet", "chart")),
    OperationSpec("set_chart", "setChart", ("sheet", "name", "chart")),
    OperationSpec("delete_chart", "deleteChart", ("sheet", "name")),
    OperationSpec("get_conditional_formatting", "getConditionalFormatting", ("sheet",)),
    OperationSpec("set_conditional_formatting", "setConditionalFormatting", ("sheet", "rules"), ("clear",)),
    OperationSpec("remove_conditional_formatting", "removeConditionalFormatting", ("sheet", "indices")),
    OperationSpec("set_cells", "setCells", ("cells",)),
    OperationSpec("scale_range", "setCells", ("cells",), composite=True),
    OperationSpec("insert_row_after", "insertRowAfter", ("sheet", "row", "count")),
    OperationSpec("delete_rows", "deleteRows", ("sheet", "row", "count")),
    OperationSpec("insert_column_after", "insertColumnAfter", ("sheet", "column", "count")),
    OperationSpec("delete_columns", "deleteColumns", ("sheet", "column", "count")),
    OperationSpec("auto_fit_columns", "autoFitColumns", ("sheet",), ("columns", "minWidth", "maxWidth", "padding")),
    OperationSpec("auto_fit_rows", "autoFitRows", ("sheet",), ("rows", "minHeight", "maxHeight")),
    OperationSpec("sort_range", "sortRange", ("range", "keys"), ("hasHeader",)),
    OperationSpec("copy_range", "copyRange", ("source", "destination"), ("pasteType",)),
    OperationSpec("reduce_addresses", "reduceAddresses", ("addresses",)),
    OperationSpec("get_style", "getStyle", ("address",)),
    OperationSpec("set_style", "setStyle", ("address", "style")),
    OperationSpec("save", "save"),
)

OPS: dict[str, OperationSpec] = {op.method: op for op in OPERATIONS}
