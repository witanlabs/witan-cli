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
SetCellsValidationMode: TypeAlias = Literal["ignore", "reject"]


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


class ViewRangeBounds(TypedDict):
    startRow: int
    startCol: int
    endRow: int
    endCol: int


class UpdatedSheetInfo(TypedDict):
    name: str
    usedRange: ViewRangeBounds | None
    tileRowCount: int
    tileColCount: int


class InvalidatedTile(TypedDict):
    sheet: str
    tileRow: int
    tileCol: int


class WriteResult(TypedDict):
    touched: dict[str, str]
    changed: list[str]
    errors: list[Diagnostic]
    invalidatedTiles: list[InvalidatedTile]
    updatedSheets: list[UpdatedSheetInfo]


class FindAndReplaceResult(TypedDict):
    replaced: int
    cells: list[str]
    errors: list[Diagnostic]


class ListObjectMutationResult(WriteResult):
    listObject: ListObject


class DataTableMutationResult(WriteResult):
    dataTable: DataTable


DataValidationOperator: TypeAlias = Literal[
    "Between",
    "NotBetween",
    "EqualTo",
    "NotEqualTo",
    "GreaterThan",
    "LessThan",
    "GreaterThanOrEqualTo",
    "LessThanOrEqualTo",
]
DataValidationAlertStyle: TypeAlias = Literal["Stop", "Warning", "Information"]
DataValidationStatus: TypeAlias = Literal["Valid", "Invalid", "NoValidation", "Mixed", "Unknown"]


class BasicDataValidationRule(TypedDict):
    operator: DataValidationOperator
    formula1: str | int | float
    formula2: NotRequired[str | int | float | None]


class ListDataValidationRule(TypedDict):
    source: str
    inCellDropDown: NotRequired[bool]


class CustomDataValidationRule(TypedDict):
    formula: str


class DataValidationRulePayload(TypedDict, total=False):
    wholeNumber: BasicDataValidationRule | None
    decimal: BasicDataValidationRule | None
    list: ListDataValidationRule | None
    date: BasicDataValidationRule | None
    time: BasicDataValidationRule | None
    textLength: BasicDataValidationRule | None
    custom: CustomDataValidationRule | None


class DataValidationPrompt(TypedDict, total=False):
    showPrompt: bool
    title: str | None
    message: str | None


class DataValidationErrorAlert(TypedDict, total=False):
    showAlert: bool
    style: DataValidationAlertStyle
    title: str | None
    message: str | None


class DataValidationSpec(TypedDict):
    address: str
    rule: DataValidationRulePayload
    ignoreBlanks: NotRequired[bool]
    prompt: NotRequired[DataValidationPrompt | None]
    errorAlert: NotRequired[DataValidationErrorAlert | None]


class DataValidationInfo(TypedDict):
    address: str
    rule: DataValidationRulePayload
    index: int
    sheet: str
    type: Literal["None", "WholeNumber", "Decimal", "List", "Date", "Time", "TextLength", "Custom"]
    ignoreBlanks: bool
    prompt: DataValidationPrompt
    errorAlert: DataValidationErrorAlert


class DataValidationDiagnostic(TypedDict):
    code: str
    message: str
    details: NotRequired[dict[str, str] | None]


class DataValidationResult(TypedDict):
    status: DataValidationStatus
    invalidCells: list[str]
    truncated: bool
    diagnostics: list[DataValidationDiagnostic]


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


class FontOptions(TypedDict, total=False):
    name: str | None
    size: float | None


class WorkbookDefaultFont(TypedDict):
    name: str
    size: float


class ThemeColors(TypedDict):
    dark1: str
    light1: str
    dark2: str
    light2: str
    accent1: str
    accent2: str
    accent3: str
    accent4: str
    accent5: str
    accent6: str
    hyperlink: str
    followedHyperlink: str
    majorFont: NotRequired[str | None]
    minorFont: NotRequired[str | None]


class WorkbookMetadata(TypedDict, total=False):
    author: str | None
    title: str | None
    subject: str | None
    company: str | None
    created: str | None
    modified: str | None


class IterativeCalculationSettings(TypedDict):
    enabled: bool
    maxIterations: int
    maxChange: float


class WorkbookProperties(TypedDict):
    activeSheetIndex: int
    defaultFont: WorkbookDefaultFont
    metadata: WorkbookMetadata | None
    themeColors: ThemeColors | None
    iterativeCalculation: IterativeCalculationSettings


class WorkbookPropertiesUpdate(TypedDict, total=False):
    activeSheetIndex: int
    defaultFont: FontOptions
    metadata: WorkbookMetadata
    themeColors: ThemeColors
    iterativeCalculation: IterativeCalculationSettings


class SheetViewProperties(TypedDict):
    showGridLines: bool
    zoomScale: int
    freezeRows: int
    freezeColumns: int


class SheetOutlineProperties(TypedDict):
    summaryRowsBelow: bool
    summaryColumnsRight: bool
    showSymbols: bool


class SheetFormatProperties(TypedDict):
    defaultRowHeight: float
    defaultColWidth: float
    font: FontOptions | None


class RowProperties(TypedDict, total=False):
    height: float
    hidden: bool
    outlineLevel: int
    collapsed: bool


class ColumnProperties(TypedDict, total=False):
    width: float
    hidden: bool
    outlineLevel: int
    collapsed: bool


class RowDimension(TypedDict, total=False):
    row: int
    height: float
    hidden: bool
    outlineLevel: int
    collapsed: bool


class ColumnDimension(TypedDict, total=False):
    col: str
    width: float
    hidden: bool
    outlineLevel: int
    collapsed: bool


class SheetProperties(TypedDict):
    visibility: SheetVisibility
    view: SheetViewProperties
    outline: SheetOutlineProperties
    format: SheetFormatProperties
    columns: dict[str, ColumnDimension]
    rows: dict[int, RowDimension]
    merges: NotRequired[list[str] | None]


class SheetPropertiesUpdate(TypedDict, total=False):
    visibility: SheetVisibility
    view: SheetViewProperties
    outline: SheetOutlineProperties
    format: SheetFormatProperties
    merges: list[str]


ThemeColorName: TypeAlias = Literal[
    "background1",
    "light1",
    "text1",
    "dark1",
    "background2",
    "light2",
    "text2",
    "dark2",
    "accent1",
    "accent2",
    "accent3",
    "accent4",
    "accent5",
    "accent6",
    "hyperlink",
    "followedHyperlink",
]


class GradientFill(TypedDict):
    type: str
    color1: str
    color2: str
    degree: NotRequired[float]
    top: NotRequired[float]
    bottom: NotRequired[float]
    left: NotRequired[float]
    right: NotRequired[float]


class FillStyle(TypedDict, total=False):
    color: str
    pattern: str
    patternColor: str
    gradient: GradientFill


class FontStyle(TypedDict, total=False):
    name: str
    size: float
    color: str
    bold: bool
    italic: bool
    strike: bool
    underline: str
    verticalAlign: str


class RichTextRun(TypedDict):
    text: str
    style: NotRequired[FontStyle]


class AlignmentStyle(TypedDict, total=False):
    horizontal: str
    vertical: str
    rotation: int
    wrapText: bool
    shrinkToFit: bool
    indent: float
    readingOrder: Literal["context", "leftToRight", "rightToLeft"]
    autoIndent: bool


class ProtectionStyle(TypedDict, total=False):
    locked: bool
    formulaHidden: bool


class BorderEdgeStyle(TypedDict, total=False):
    style: str
    color: str
    themeColor: ThemeColorName
    tintAndShade: float


class DiagonalBorderStyle(BorderEdgeStyle, total=False):
    up: bool
    down: bool


class BorderStyle(TypedDict, total=False):
    top: BorderEdgeStyle
    bottom: BorderEdgeStyle
    left: BorderEdgeStyle
    right: BorderEdgeStyle
    diagonal: DiagonalBorderStyle


class Style(TypedDict, total=False):
    fill: FillStyle | None
    font: FontStyle | None
    alignment: AlignmentStyle | None
    protection: ProtectionStyle | None
    border: BorderStyle | None
    numberFormat: str | None
    centerContinuousSpan: int | None
    richText: list[RichTextRun] | None


class ListObjectColumn(TypedDict, total=False):
    name: str
    totalsRowFunction: str | None
    totalsRowLabel: str | None
    totalsRowFormula: str | None
    calculatedColumnFormula: str | None


class CellWrite(TypedDict, total=False):
    value: ScalarCellValue
    formula: str
    format: str
    note: JsonMapping | None
    link: JsonMapping | None
    thread: JsonMapping | None


class ListObject(TypedDict):
    name: str
    sheet: str
    ref: str
    showHeaderRow: bool
    showTotalsRow: bool
    showAutoFilter: bool
    tableStyleName: str | None
    showFirstColumn: bool
    showLastColumn: bool
    showRowStripes: bool
    showColumnStripes: bool
    headerRowRange: str | None
    dataRange: str | None
    totalsRowRange: str | None
    columns: list[ListObjectColumn]


class ListObjectSpec(TypedDict):
    name: str
    ref: str
    columns: list[ListObjectColumn]
    showHeaderRow: NotRequired[bool]
    showTotalsRow: NotRequired[bool]
    showAutoFilter: NotRequired[bool]
    tableStyleName: NotRequired[str | None]
    showFirstColumn: NotRequired[bool]
    showLastColumn: NotRequired[bool]
    showRowStripes: NotRequired[bool]
    showColumnStripes: NotRequired[bool]
    rows: NotRequired[list[list[CellWrite]]]


class ListObjectUpdate(TypedDict, total=False):
    ref: str
    showHeaderRow: bool
    showTotalsRow: bool
    showAutoFilter: bool
    tableStyleName: str | None
    showFirstColumn: bool
    showLastColumn: bool
    showRowStripes: bool
    showColumnStripes: bool
    columns: list[ListObjectColumn]
    rows: list[list[CellWrite]]


class DataTableSourceFormula(TypedDict):
    formula: str


class OneVariableColumnDataTable(TypedDict):
    type: Literal["oneVariableColumn"]
    sheet: str
    ref: str
    dataTableRange: str
    rowInputCell: None
    columnInputCell: str
    inputValues: list[ScalarCellValue]
    rowInputValues: None
    columnInputValues: None
    formulas: list[DataTableSourceFormula]
    formula: None


class OneVariableRowDataTable(TypedDict):
    type: Literal["oneVariableRow"]
    sheet: str
    ref: str
    dataTableRange: str
    rowInputCell: str
    columnInputCell: None
    inputValues: list[ScalarCellValue]
    rowInputValues: None
    columnInputValues: None
    formulas: list[DataTableSourceFormula]
    formula: None


class TwoVariableDataTable(TypedDict):
    type: Literal["twoVariable"]
    sheet: str
    ref: str
    dataTableRange: str
    rowInputCell: str
    columnInputCell: str
    inputValues: None
    rowInputValues: list[ScalarCellValue]
    columnInputValues: list[ScalarCellValue]
    formulas: None
    formula: str


DataTable: TypeAlias = OneVariableColumnDataTable | OneVariableRowDataTable | TwoVariableDataTable


class OneVariableColumnDataTableSpec(TypedDict):
    type: Literal["oneVariableColumn"]
    ref: str
    columnInputCell: str
    inputValues: list[ScalarCellValue]
    formulas: list[str]


class OneVariableRowDataTableSpec(TypedDict):
    type: Literal["oneVariableRow"]
    ref: str
    rowInputCell: str
    inputValues: list[ScalarCellValue]
    formulas: list[str]


class TwoVariableDataTableSpec(TypedDict):
    type: Literal["twoVariable"]
    ref: str
    rowInputCell: str
    columnInputCell: str
    rowInputValues: list[ScalarCellValue]
    columnInputValues: list[ScalarCellValue]
    formula: str


DataTableSpec: TypeAlias = OneVariableColumnDataTableSpec | OneVariableRowDataTableSpec | TwoVariableDataTableSpec


ChartType: TypeAlias = Literal[
    "column",
    "bar",
    "line",
    "area",
    "pie",
    "doughnut",
    "scatter",
    "bubble",
    "stockHLC",
    "stockOHLC",
    "waterfall",
]
ChartGrouping: TypeAlias = Literal["standard", "stacked", "percentStacked"]
ChartAxisBinding: TypeAlias = Literal["primary", "secondary"]
ChartStockRole: TypeAlias = Literal["volume", "open", "high", "low", "close"]
ChartLegendPosition: TypeAlias = Literal["left", "right", "top", "bottom", "topRight"]
ChartMarkerStyle: TypeAlias = Literal[
    "auto",
    "none",
    "circle",
    "dash",
    "diamond",
    "dot",
    "picture",
    "plus",
    "square",
    "star",
    "triangle",
    "x",
]
ChartDataLabelPosition: TypeAlias = Literal[
    "bestFit",
    "center",
    "insideBase",
    "insideEnd",
    "outsideEnd",
    "left",
    "right",
    "top",
    "bottom",
]
ChartTimeUnit: TypeAlias = Literal["days", "months", "years"]
ChartScatterStyle: TypeAlias = Literal["line", "lineMarker", "marker", "smooth", "smoothMarker"]


class ChartTextSource(TypedDict, total=False):
    text: str
    ref: str


class ChartPositionAnchor(TypedDict, total=False):
    cell: str
    xOffsetPts: float
    yOffsetPts: float


ChartPositionInput = TypedDict("ChartPositionInput", {"from": ChartPositionAnchor, "to": ChartPositionAnchor})


class ChartPosition(ChartPositionInput, total=False):
    sheet: str


class ChartFillFormatSpec(TypedDict, total=False):
    noFill: bool
    color: str


class ChartBorderFormatSpec(TypedDict, total=False):
    noLine: bool
    color: str
    weight: float
    lineStyle: str


ChartLineFormatSpec: TypeAlias = ChartBorderFormatSpec


class ChartFontFormatSpec(TypedDict, total=False):
    bold: bool
    color: str
    italic: bool
    name: str
    size: float
    underline: str


class ChartDataLabelFormatSpec(TypedDict, total=False):
    fill: ChartFillFormatSpec
    border: ChartBorderFormatSpec
    font: ChartFontFormatSpec


class ChartAreaFormatSpec(ChartDataLabelFormatSpec, total=False):
    pass


class ChartPlotAreaFormatSpec(TypedDict, total=False):
    fill: ChartFillFormatSpec
    border: ChartBorderFormatSpec


class ChartPlotAreaSpec(TypedDict, total=False):
    format: ChartPlotAreaFormatSpec


class ChartSeriesFormatSpec(TypedDict, total=False):
    fill: ChartFillFormatSpec
    line: ChartLineFormatSpec


class ChartMarkerSpec(TypedDict, total=False):
    style: ChartMarkerStyle
    size: float
    fillColor: str
    borderColor: str


class ChartDataLabelsSpec(TypedDict, total=False):
    showLegendKey: bool
    showValue: bool
    showCategory: bool
    showSeriesName: bool
    showPercent: bool
    showBubbleSize: bool
    showLeaderLines: bool
    position: ChartDataLabelPosition
    numberFormat: str
    numberFormatLinked: bool
    separator: str
    format: ChartDataLabelFormatSpec


class ChartSeriesSpec(TypedDict, total=False):
    name: ChartTextSource
    stockRole: ChartStockRole
    categories: str
    categoriesRefType: Literal["string", "number", "multiLevelString"]
    values: str
    xValues: str
    yValues: str
    bubbleSizes: str
    fillColor: str
    lineColor: str
    lineWidth: float
    lineDashStyle: str
    smooth: bool
    invertIfNegative: bool
    totalIndexes: list[int]
    showConnectorLines: bool
    marker: ChartMarkerSpec
    dataLabels: ChartDataLabelsSpec


class ChartGroupSpec(TypedDict):
    type: ChartType
    series: list[ChartSeriesSpec]
    scatterStyle: NotRequired[ChartScatterStyle]
    grouping: NotRequired[ChartGrouping]
    axis: NotRequired[ChartAxisBinding]
    gapWidth: NotRequired[int]
    overlap: NotRequired[int]
    varyColors: NotRequired[bool]
    smooth: NotRequired[bool]
    bubbleScale: NotRequired[int]
    showNegativeBubbles: NotRequired[bool]
    sizeRepresents: NotRequired[Literal["area", "width"]]
    firstSliceAngle: NotRequired[int]
    holeSize: NotRequired[int]
    dataLabels: NotRequired[ChartDataLabelsSpec]


class ChartTitleSpec(ChartTextSource, total=False):
    overlay: bool


class ChartLegendSpec(TypedDict, total=False):
    visible: bool
    position: ChartLegendPosition
    overlay: bool


class ChartAxisSpec(TypedDict, total=False):
    title: ChartTextSource
    visible: bool
    categoryType: Literal["category", "date"]
    min: float
    max: float
    majorUnit: float
    minorUnit: float
    baseTimeUnit: ChartTimeUnit
    majorTimeUnit: ChartTimeUnit
    minorTimeUnit: ChartTimeUnit
    numberFormat: str
    numberFormatLinked: bool
    reversed: bool
    majorGridlines: bool
    minorGridlines: bool
    position: Literal["left", "right", "top", "bottom"]


class ChartAxesSpec(TypedDict, total=False):
    category: ChartAxisSpec
    value: ChartAxisSpec
    secondaryCategory: ChartAxisSpec
    secondaryValue: ChartAxisSpec


class ChartSpec(TypedDict):
    name: str
    position: ChartPositionInput
    groups: list[ChartGroupSpec]
    id: NotRequired[int]
    title: NotRequired[ChartTitleSpec]
    legend: NotRequired[ChartLegendSpec]
    axes: NotRequired[ChartAxesSpec]
    format: NotRequired[ChartAreaFormatSpec]
    plotArea: NotRequired[ChartPlotAreaSpec]
    displayBlanksAs: NotRequired[Literal["gap", "span", "zero"]]
    plotVisibleOnly: NotRequired[bool]
    showDataLabelsOverMaximum: NotRequired[bool]
    roundedCorners: NotRequired[bool]
    styleId: NotRequired[int]


class ChartInfo(TypedDict):
    name: str
    groups: list[ChartGroupSpec]
    position: ChartPosition
    id: NotRequired[int]
    title: NotRequired[ChartTitleSpec]
    legend: NotRequired[ChartLegendSpec]
    axes: NotRequired[ChartAxesSpec]
    format: NotRequired[ChartAreaFormatSpec]
    plotArea: NotRequired[ChartPlotAreaSpec]
    displayBlanksAs: NotRequired[Literal["gap", "span", "zero"]]
    plotVisibleOnly: NotRequired[bool]
    showDataLabelsOverMaximum: NotRequired[bool]
    roundedCorners: NotRequired[bool]
    styleId: NotRequired[int]


class ChartGroupSummary(TypedDict):
    type: str
    seriesCount: int
    axis: NotRequired[ChartAxisBinding]


class ChartSummary(TypedDict):
    sheet: str
    name: str
    type: str
    groups: list[ChartGroupSummary]
    groupCount: int
    seriesCount: int
    position: ChartPosition
    id: NotRequired[int]


CfThresholdType: TypeAlias = Literal["formula", "max", "min", "num", "percent", "percentile", "autoMin", "autoMax"]


class CfColorScalePoint(TypedDict):
    type: CfThresholdType
    value: NotRequired[float]
    formula: NotRequired[str]
    color: NotRequired[str]


class CfDataBarThreshold(TypedDict):
    type: CfThresholdType
    value: NotRequired[float]
    formula: NotRequired[str]


class CfDataBarConfig(TypedDict, total=False):
    showValue: bool
    gradient: bool
    border: bool
    negativeBarColorSameAsPositive: bool
    negativeBarBorderColorSameAsPositive: bool
    axisPosition: Literal["automatic", "middle", "none"]
    direction: Literal["context", "leftToRight", "rightToLeft"]
    fillColor: str
    borderColor: str
    negativeFillColor: str
    negativeBorderColor: str
    axisColor: str
    lowValue: CfDataBarThreshold
    highValue: CfDataBarThreshold


ConditionalFormattingRuleType: TypeAlias = Literal[
    "cellValue",
    "containsText",
    "notContainsText",
    "beginsWith",
    "endsWith",
    "containsBlanks",
    "notContainsBlanks",
    "containsErrors",
    "notContainsErrors",
    "expression",
    "timePeriod",
    "top",
    "bottom",
    "aboveAverage",
    "belowAverage",
    "duplicateValues",
    "uniqueValues",
    "twoColorScale",
    "threeColorScale",
    "dataBar",
    "iconSet",
]
ConditionalFormattingOperator: TypeAlias = Literal[
    "equal",
    "notEqual",
    "greaterThan",
    "greaterThanOrEqual",
    "lessThan",
    "lessThanOrEqual",
    "between",
    "notBetween",
]
ConditionalFormattingTimePeriod: TypeAlias = Literal[
    "today",
    "yesterday",
    "tomorrow",
    "last7Days",
    "thisWeek",
    "lastWeek",
    "nextWeek",
    "thisMonth",
    "lastMonth",
    "nextMonth",
]


class ConditionalFormattingRule(TypedDict, total=False):
    address: str
    type: ConditionalFormattingRuleType
    index: int
    priority: int
    stopIfTrue: bool
    style: Style
    operator: ConditionalFormattingOperator
    formula: str
    formula2: str
    text: str
    rank: int
    percent: bool
    equalAverage: bool
    stdDev: int
    lowValue: CfColorScalePoint
    midValue: CfColorScalePoint
    highValue: CfColorScalePoint
    dataBar: CfDataBarConfig
    iconSetStyle: str
    timePeriod: ConditionalFormattingTimePeriod


class DetectedTable(TypedDict, total=False):
    address: str
    headerRows: str
    headerCols: str | None
    tableName: str | None


class SheetDescription(TypedDict):
    tables: dict[str, DetectedTable]
    structure: str


def regex_from_pattern(pattern: Pattern[str]) -> Regex:
    flags = ""
    if pattern.flags & re.IGNORECASE:
        flags += "i"
    if pattern.flags & re.MULTILINE:
        flags += "m"
    if pattern.flags & re.DOTALL:
        flags += "s"
    return Regex(pattern.pattern, flags)
