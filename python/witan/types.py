from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Literal, Mapping, Pattern, TypeAlias, TypedDict

from typing_extensions import NotRequired

JsonDict: TypeAlias = dict[str, Any]
JsonMapping: TypeAlias = Mapping[str, Any]
CellRef: TypeAlias = str | Mapping[str, Any]
RangeRef: TypeAlias = str | Mapping[str, Any]
ScalarCellValue: TypeAlias = str | int | float | bool | None
Visibility: TypeAlias = Literal["visible", "outsidePrintArea", "collapsed", "hidden"]
SheetVisibility: TypeAlias = Literal["visible", "hidden", "veryHidden"]


@dataclass(frozen=True)
class Regex:
    """JavaScript-compatible regex matcher for search operations."""

    source: str
    flags: str = ""


Matcher: TypeAlias = str | list[str] | int | float | bool | Regex | Pattern[str] | list[Regex] | list[Pattern[str]]
ReplaceMatcher: TypeAlias = str | Regex | Pattern[str]


class CellCoordinates(TypedDict):
    sheet: str
    row: int
    col: int | str


class SheetInfo(TypedDict):
    address: str
    rows: int
    cols: int
    sheet: str
    hidden: NotRequired[bool]
    printArea: NotRequired[str]
    listObjects: NotRequired[list[str]]
    dataTables: NotRequired[list[str]]
    precedents: NotRequired[list[str]]
    dependents: NotRequired[list[str]]


class Note(TypedDict):
    author: str
    text: str


class Link(TypedDict):
    type: Literal["internal", "external"]
    target: str
    tooltip: NotRequired[str]


class ThreadComment(TypedDict):
    authorId: str
    text: str
    createdAt: str


class ThreadInfo(TypedDict):
    resolved: bool
    comments: list[ThreadComment]


class Value(TypedDict):
    address: str
    sheet: str
    row: int
    col: int
    colLetter: str
    value: Any
    formula: NotRequired[str]
    type: Literal["string", "number", "bool", "date", "error", "blank"]
    text: str
    format: NotRequired[str]
    numberType: NotRequired[Literal["currency", "percent", "fraction", "exponential", "date", "text", "number"]]
    visibility: Visibility
    context: NotRequired[str]
    note: NotRequired[Note]
    link: NotRequired[Link]
    thread: NotRequired[ThreadInfo]


class Diagnostic(TypedDict):
    code: str
    address: str
    detail: NotRequired[str]
    formula: NotRequired[str]


class UpdatedSheetInfo(TypedDict):
    name: str
    usedRange: Any
    tileRowCount: int
    tileColCount: int


class WriteResult(TypedDict):
    touched: dict[str, str]
    changed: list[str]
    errors: list[Diagnostic]
    invalidatedTiles: list[JsonDict]
    updatedSheets: list[UpdatedSheetInfo]


class FindAndReplaceResult(TypedDict):
    replaced: int
    cells: list[str]
    errors: list[Diagnostic]


class ListObjectMutationResult(WriteResult):
    listObject: JsonDict


class DataTableMutationResult(WriteResult):
    dataTable: JsonDict


class DefinedName(TypedDict):
    name: str
    range: str
    scope: str | None


class SearchCell(TypedDict):
    type: Literal["cell"]
    address: str
    value: Any
    text: str
    formula: NotRequired[str]
    row: int
    col: int
    colLetter: str
    sheet: str
    visibility: Visibility
    context: NotRequired[str]
    role: str


class SearchRow(TypedDict):
    type: Literal["row"]
    row: int
    sheet: str
    matchedAt: str
    range: str
    tsv: str
    visibility: Visibility
    context: NotRequired[str]


class TableLookupResult(TypedDict):
    address: str
    value: Any
    text: str
    row: int
    col: int
    colLetter: str
    sheet: str
    visibility: Visibility
    rowLabelFoundAt: str
    rowLabelFound: str
    columnLabelFoundAt: str
    columnLabelFound: str


class DependencyCell(TypedDict):
    address: str
    depth: int
    formula: NotRequired[str]
    referenceType: NotRequired[Literal["direct", "range", "named", "table"]]


class DependencyResult(TypedDict):
    cells: list[DependencyCell]
    warnings: NotRequired[list[Diagnostic]]


class TraceInput(TypedDict):
    address: str
    referenceCount: int
    text: NotRequired[str]
    nearbyLabel: NotRequired[str]
    context: NotRequired[str]


class TraceOutput(TypedDict):
    address: str
    formula: NotRequired[str]
    text: NotRequired[str]
    visibility: Visibility
    nearbyLabel: NotRequired[str]
    context: NotRequired[str]


class FormulaResult(TypedDict):
    formula: str
    value: Any
    error: NotRequired[JsonDict]


class SweepEntry(TypedDict):
    inputs: dict[str, str]
    outputs: dict[str, str]
    errors: list[Diagnostic]


class OutputStats(TypedDict):
    min: float
    max: float
    mean: float
    count: int


class SweepResult(TypedDict):
    tsv: str
    sweeps: list[SweepEntry]
    stats: NotRequired[dict[str, OutputStats]]
    sweepCount: int
    inputCount: int
    outputCount: int


class CopyRangeResult(TypedDict):
    destination: str
    cellsCopied: int


class AutoFitColumnResult(TypedDict):
    width: float
    previousWidth: float


class AutoFitRowResult(TypedDict):
    height: float
    previousHeight: float
    hidden: bool
    previousHidden: bool


class LintDiagnostic(TypedDict):
    severity: Literal["Info", "Warning", "Error"]
    ruleId: str
    message: str
    location: str | None
    visibility: Visibility | None


class LintResult(TypedDict):
    diagnostics: list[LintDiagnostic]
    total: int


WorkbookProperties: TypeAlias = JsonDict
WorkbookPropertiesUpdate: TypeAlias = JsonMapping
SheetProperties: TypeAlias = JsonDict
SheetPropertiesUpdate: TypeAlias = JsonMapping
Style: TypeAlias = JsonMapping
ChartSpec: TypeAlias = JsonMapping
ChartInfo: TypeAlias = JsonDict
ChartSummary: TypeAlias = JsonDict
ConditionalFormattingRule: TypeAlias = JsonMapping
ListObject: TypeAlias = JsonDict
ListObjectSpec: TypeAlias = JsonMapping
ListObjectUpdate: TypeAlias = JsonMapping
DataTable: TypeAlias = JsonDict
DataTableSpec: TypeAlias = JsonMapping


def regex_from_pattern(pattern: Pattern[str]) -> Regex:
    flags = ""
    if pattern.flags & re.IGNORECASE:
        flags += "i"
    if pattern.flags & re.MULTILINE:
        flags += "m"
    if pattern.flags & re.DOTALL:
        flags += "s"
    return Regex(pattern.pattern, flags)
