---
name: xlsx-code-mode
description: Use this skill for any Excel workbook task involving .xls, .xlsx, or .xlsm files: read, inspect, search, understand, create, modify, repair, author, or verify spreadsheets. Trigger for workbook creation from scratch; edits to cells, formulas, sheets, styles, tables, charts, images, conditional formatting, data validation, defined names, metadata, or sheet properties; dependency tracing; what-if analysis; and Excel-grade verification with lint, calc, render. Also trigger whenever the user references a spreadsheet file by name or path, even casually. The tool runs sandboxed JavaScript against the workbook server-side via `witan xlsx exec`."
---

> **Running in Claude Cowork?** The `witan` CLI isn't preinstalled in the sandbox. Read [references/cowork-setup.md](references/cowork-setup.md) first for install, PATH and network-allowlist steps.

> **No `witan` on PATH?** Prefix commands with `npx witan` (e.g. `npx witan xlsx exec ...`).

## Setup

The CLI supports `.xls`, `.xlsx`, and `.xlsm`; legacy `.xls` files are converted to `.xlsx` when needed. New workbook creation is `.xlsx` only.

## Quick Reference

```bash
# Create a workbook (.xlsx only)
witan xlsx exec model.xlsx --create --save --stdin <<'WITAN'
await xlsx.addSheet(wb, "Inputs")
await xlsx.setCells(wb, [{ address: "Inputs!A1", value: "Revenue" }])
return await xlsx.listSheets(wb)
WITAN

# Read awkward sheet names; apostrophes inside sheet names are doubled
witan xlsx exec model.xlsx --stdin <<'WITAN'
const a = await xlsx.readCell(wb, "'Workers'' Compensation'!B50")
const b = await xlsx.readRangeTsv(wb, { sheet: "Reserve Summary (Net)", from: {row:1,col:1}, to: {row:10,col:5} })
return { a: a.value, b }
WITAN

# Sensitivity sweep
witan xlsx exec model.xlsx --stdin <<'WITAN'
const result = await xlsx.sweepInputs(wb, {
	inputs: [
		{ address: "Inputs!B5", values: [0.02, 0.04, 0.06] },
		{ address: "Inputs!B6", values: [0.08, 0.10, 0.12] },
	],
	outputs: ["Output!C30", "Output!C45"],
	mode: "cartesian",
	includeStats: true,
})
console.log(result.tsv)
WITAN

# Author a chart and preview it
witan xlsx exec model.xlsx --stdin <<'WITAN'
await xlsx.addChart(wb, "Sheet1", {
	name: "Revenue",
	position: { from: { cell: "F2" }, to: { cell: "N18" } },
	groups: [
		{
			type: "column",
			series: [
				{
					name: { ref: "Sheet1!B1" },
					categories: "Sheet1!A2:A9",
					values: "Sheet1!B2:B9"
				}
			]
		}
	],
	title: { text: "Revenue" },
	legend: { position: "right" }
})
await xlsx.previewStyles(wb, "Sheet1!F2:N18")
WITAN

# Simple one-liner
witan xlsx exec model.xlsx --expr 'xlsx.listSheets(wb)'
```

## exec — Workbook Scripting

Runs server-side JavaScript against a workbook through globals `xlsx`, `wb`, `input`, and `print`. Top-level `await` is supported; imports are not.

Use `--create` only with new `.xlsx` paths. Add `--save` to write bytes to disk; without `--save`, creation and edits are session-only.

### Invocation patterns

Prefer `--stdin` with a single-quoted heredoc for multi-line scripts and any sheet name with spaces, apostrophes, parentheses, quotes, or glob characters:

```bash
witan xlsx exec report.xlsx --stdin <<'WITAN'
const sheets = await xlsx.listSheets(wb)
const cell = await xlsx.readCell(wb, "'My Sheet'!A1")
return { sheets, cell }
WITAN
```

Use exactly one code source: `--stdin`, `--expr`, `--code`, or `--script`. `--expr` is fine for simple one-liners; use `--script file.js --input-json ...` for reusable scripts.

### Flags

- Code source: exactly one of `--stdin`, `--expr`, `--code`, `--script`
- Inputs/control: `--input-json`, `--input-file key=@path`, `--timeout-ms`, `--max-output-chars`, `--stdin-timeout-ms`, `--locale`
- Lifecycle/output: `--create`, `--save`, `--json`

### Runtime globals

- `xlsx`: curated API; pass `wb` as first arg to all functions
- `wb`: opened workbook handle
- `input`: parsed `--input-json`; `--input-file logo=@./logo.png` adds `input.logo` as `data:image/...;base64,...`
- `print`: stdout helper like `console.log`

### API reference

All functions are async and take `wb` first.

- Read: `listSheets`, `getWorkbookProperties`, `getSheetProperties`, `listDefinedNames`, `readCell`, `readRange`, `readRow`, `readColumn`, TSV variants, `getStyle`
- Search: `findCells`, `findRows`, `describeSheet`, `tableLookup`, `getListObject`, `getDataTable`; `matcher` accepts string/string[]/number/boolean/RegExp/RegExp[]
- Trace/compute/render: `getCellPrecedents`, `getCellDependents`, `traceToInputs`, `traceToOutputs`, `sweepInputs`, `evaluateFormula(s)`, `lint`, `previewStyles`
- Write/structure: `setCells`, `findAndReplace`, `scaleRange`, `copyRange`, `sortRange`, row/column insert/delete, `autoFitColumns`, sheet/name/property/style setters
- Objects: `add/set/deleteListObject`, `add/deleteDataTable`, conditional formatting and data validation APIs
- Charts/images: `list/get/add/set/deleteChart`, `list/get/add/set/deleteImage`; charts support combo, secondary axes, stock, bubble, radar, top-view surface, waterfall, histogram/Pareto, funnel, formatting, labels, number formats, style IDs

Common options: searches use `opts.in`, `context`, `limit`, `offset`; TSV reads use `includeEmpty`/`includeFormulas`; `scaleRange.skipFormulas` defaults true; `sortRange.hasHeader` defaults true; `copyRange.pasteType` supports `all`, `values`, `formulas`, `formats`; `setConditionalFormatting` and `setDataValidations` use `opts.clear` to replace existing rules.

For local images, prefer `--input-file logo=@./logo.png` then `source:{base64:input.logo}`. `source.base64` also accepts raw base64 or data URLs; image responses return metadata only.

### Write and What-If Rules

`exec` is ephemeral unless `--save` is passed. Each invocation starts from the original workbook; writes such as `setCells`, `scaleRange`, inserts, deletes, formatting, charts, and images affect only the server-side session unless saved. `setCells` creates a missing referenced sheet and returns:

```ts
{ touched:Record<string,string>; changed:string[]; errors:Diag[] }
```

Read output values from `result.touched["Sheet!Address"]`; do not recompute answers in JavaScript.

For "what happens to Y if X changes?":

1. Separate exec: find the output cell first with `findCells(wb, matcher, { context: 2 })`, synonyms, or `readRangeTsv`; choose the formula/output cell, not the label.
2. Separate exec: `traceToInputs(wb, outputAddr)`, confirm the user-named input drives it, `setCells`, then read `result.touched[outputAddr]`.
3. Report baseline and changed values, and check `result.errors`.

Use `sweepInputs` for multi-value sensitivities. If tracing returns hundreds of cells, filter by `nearbyLabel` or context before printing. For iterative/circular models, `setCells` uses workbook iterative settings and reports convergence errors.

### calc — Full-workbook verification

`setCells` already recalculates edited cells and downstream dependents. Use `calc` for standalone workbook-wide checks or final audits:

```bash
witan xlsx calc model.xlsx
witan xlsx calc model.xlsx --verify
witan xlsx calc model.xlsx --show-touched
```

It prints all formula errors, not just a count. Exit code `2` means formula errors; with `--verify`, exit code `2` can also mean computed values changed. `--show-touched` prints every recalculated cell.

### Response format

With `--json`, success includes `ok`, `stdout`, `result`, `writes_detected`, and `accesses`; failure includes `ok:false`, `stdout`, and `error:{type,code,message}`. `accesses` records cell reads/writes.

## render — Visual Screenshot

Renders a sheet range as PNG/WebP for layout, merged cells, formatting, charts, and labels.

```bash
witan xlsx render report.xlsx -r "Sheet1!A1:Z50"
witan xlsx render report.xlsx -r "'My Sheet'!B5:H20" -o snapshot.png
witan xlsx render report.xlsx -r "Sheet1!A1:F10" --dpr 2 --diff before.png
```

Flags: `--range/-r` required, `--output/-o` optional, `--dpr` 1-3 or `auto`, `--format png|webp`, `--diff baseline.png`. Inside `exec`, use `previewStyles`.

## Error Guide

- Code source errors: provide exactly one non-empty `--stdin`, `--expr`, `--code`, or `--script`.
- `Import statements are not allowed`: use the `xlsx` global, no static/dynamic imports.
- `EXEC_SYNTAX_ERROR` / `EXEC_RUNTIME_ERROR`: fix script syntax/runtime issue.
- `EXEC_RESULT_TOO_LARGE`: return less data; print/log large output.
- `invalid --input-json`, `--timeout-ms must be > 0`: fix flag value.
- `Sheet 'X' not found`: run `listSheets`; for shell quoting issues use `--stdin <<'WITAN'`.
- `findCells` empty: broaden search, use synonyms, or inspect nearby ranges.
- `setCells` missing expected output in `touched`: wrong output cell or not a dependent; trace the chain.

### Full Type Definitions

```ts
type CellRef=string|{sheet:string;row:number;col:number|string}
type RangeRef=string|{sheet:string}|{sheet:string;from:{row?:number;col?:number|string};to:{row?:number;col?:number|string}}
type Visibility="visible"|"outsidePrintArea"|"collapsed"|"hidden"
interface Diag {code:string;detail?:string;address:string;formula?:string}
type PartialDeep<T>={[K in keyof T]?:NonNullable<T[K]> extends object?PartialDeep<NonNullable<T[K]>>:T[K]}

interface WorkbookProperties {
	activeSheetIndex:number;
	defaultFont:{name:string;size:number};
	iterativeCalculation:{enabled:boolean;maxIterations:number;maxChange:number};
	metadata?:{author?:string;title?:string;subject?:string;company?:string;created?:string;modified?:string};
	themeColors?:{
		dark1:string;light1:string;dark2:string;light2:string;
		accent1:string;accent2:string;accent3:string;accent4:string;accent5:string;accent6:string;
		hyperlink:string;followedHyperlink:string;majorFont?:string;minorFont?:string;
	};
}
function getWorkbookProperties(wb):Promise<WorkbookProperties>
function setWorkbookProperties(wb,properties:PartialDeep<WorkbookProperties>):Promise<void>
function listSheets(wb):Promise<Array<{
	address:string;rows:number;cols:number;sheet:string;hidden?:boolean;printArea?:string;
	listObjects?:string[];dataTables?:string[];precedents?:string[];dependents?:string[];
}>>

interface NameDef {name:string;range:string;scope:string|null}
function listDefinedNames(wb):Promise<NameDef[]>
function addDefinedName(wb,name:string,range:string,scope?:string):Promise<NameDef>
function deleteDefinedName(wb,name:string,scope?:string):Promise<NameDef>
function addSheet(wb,name:string):Promise<string>
function deleteSheet(wb,name:string):Promise<void>
function renameSheet(wb,oldName:string,newName:string):Promise<void>

type SheetVisibility="visible"|"hidden"|"veryHidden";
interface SheetView {showGridLines:boolean;zoomScale:number}
interface SheetOutline {summaryRowsBelow:boolean;summaryColumnsRight:boolean;showSymbols:boolean}
interface SheetFormat {defaultRowHeight:number;defaultColWidth:number;font?:{name?:string;size?:number}|null}
interface RowProps {height?:number;hidden?:boolean;outlineLevel?:number;collapsed?:boolean}
interface ColProps {width?:number;hidden?:boolean;outlineLevel?:number;collapsed?:boolean}
interface SheetProperties {
	visibility:SheetVisibility;
	view:SheetView;
	outline:SheetOutline;
	format:SheetFormat;
	columns:Record<string,ColProps&{col:string;width:number}>;
	rows:Record<number,RowProps&{row:number;height:number}>;
	merges?:string[]|null;
}
function setSheetProperties(wb,sheetName:string,properties:PartialDeep<Pick<SheetProperties,"visibility"|"view"|"outline"|"format">>&{merges?:string[]}):Promise<void>
function setRowProperties(wb,sheetName:string,fromRow:number,toRow:number,properties:RowProps):Promise<void>
function setColumnProperties(wb,sheetName:string,fromCol:number|string,toCol:number|string,properties:ColProps):Promise<void>
function getSheetProperties(wb,sheetName:string,filter?:{columns?:(number|string)[];rows?:number[]}):Promise<SheetProperties>

interface CellAddr {
	address:string;sheet:string;row:number;col:number;colLetter:string;visibility:Visibility;
}
interface CellHit extends CellAddr {value:any;text:string}
interface Value extends CellAddr {
	value:string|number|boolean|null;
	formula?:string;
	type:"string"|"number"|"bool"|"date"|"error"|"blank";
	text:string;
	format?:string;
	numberType?:"currency"|"percent"|"fraction"|"exponential"|"date"|"text"|"number";
	context?:string;
	note?:{author:string;text:string};
	link?:{type:"internal"|"external";target:string;tooltip?:string};
	thread?:{resolved:boolean;comments:{authorId:string;text:string;createdAt:string}[]};
}
function readCell(wb,cell:CellRef,opts?:{context?:number}):Promise<Value>
function readRange(wb,range:RangeRef):Promise<Value[][]>
function readColumn(wb,sheetName:string,col:number|string,opts?:{startRow?:number;endRow?:number}):Promise<Value[]>
function readRow(wb,sheetName:string,row:number,opts?:{startCol?:number;endCol?:number}):Promise<Value[]>
function readRangeTsv(wb,range:RangeRef,opts?:{includeEmpty?:boolean;includeFormulas?:boolean}):Promise<string>
function readColumnTsv(wb,sheetName:string,col:number|string,opts?:{startRow?:number;endRow?:number;includeEmpty?:boolean;includeFormulas?:boolean}):Promise<string>
function readRowTsv(wb,sheetName:string,row:number,opts?:{startCol?:number;endCol?:number;includeEmpty?:boolean;includeFormulas?:boolean}):Promise<string>
declare class SearchResults<T> extends Array<T> {truncated?:boolean}
type Matcher=string|string[]|number|boolean|RegExp|RegExp[]
function findCells(wb,matcher:Matcher,opts?:{in?:RangeRef|string;context?:number;limit?:number;offset?:number;formulas?:boolean}):Promise<Array<CellHit&{
	type:"cell";formula?:string;context?:string;role:string;
}>>
function findRows(wb,matcher:Matcher,opts?:{in?:RangeRef|string;context?:number;limit?:number;offset?:number}):Promise<{
	type:"row";row:number;sheet:string;matchedAt:string;range:string;tsv:string;visibility:Visibility;context?:string;
}[]>
function findAndReplace(wb,find:string|RegExp,replace:string,opts?:{
	in?:RangeRef|string;matchCase?:boolean;wholeCell?:boolean;inFormulas?:boolean;limit?:number;
}):Promise<{replaced:number;cells:string[];errors:Diag[]}>
function describeSheet(wb,sheetName:string):Promise<{
	tables:Record<string,{address:string;headerRows:string;headerCols:string|null;tableName?:string}>;
	structure:string;
}>
function tableLookup(wb,args:{table:string;rowLabel:string|number|boolean;columnLabel:string|number|boolean}):Promise<Array<CellHit&{
	rowLabelFoundAt:string;rowLabelFound:string;columnLabelFoundAt:string;columnLabelFound:string;
}>>

interface WriteResult {
	touched:Record<string,string>;
	changed:string[];
	errors:Diag[];
	invalidatedTiles:{sheet:string;tileRow:number;tileCol:number}[];
	updatedSheets:{name:string;usedRange:{startRow:number;startCol:number;endRow:number;endCol:number}|null;tileRowCount:number;tileColCount:number}[];
}
function setCells(wb,cells:Array<{
	address:CellRef;
	value?:unknown;
	formula?:string;
	format?:string;
	note?:{text?:string;author?:string}|null;
	link?:{url?:string;ref?:string;tooltip?:string}|null;
	thread?:{add?:Array<{author?:string;text:string}>;resolved?:boolean;delete?:boolean}|null;
}>,opts?:{validationMode?:"ignore"|"reject"}):Promise<WriteResult>
function scaleRange(wb,range:RangeRef,factor:number,opts?:{skipFormulas?:boolean}):Promise<WriteResult|null>
function insertRowAfter(wb,sheetName:string,row:number,count?:number):Promise<void>
function deleteRows(wb,sheetName:string,row:number,count?:number):Promise<void>
function insertColumnAfter(wb,sheetName:string,column:number|string,count?:number):Promise<void>
function deleteColumns(wb,sheetName:string,column:number|string,count?:number):Promise<void>
function autoFitColumns(wb,sheetName:string,columns?:Array<number|string>,opts?:{minWidth?:number;maxWidth?:number;padding?:number}):Promise<Record<string,{width:number;previousWidth:number}>>
function sortRange(wb,range:RangeRef,keys:Array<{col:number|string;order?:"asc"|"desc"}>,opts?:{hasHeader?:boolean}):Promise<void>
function copyRange(wb,source:RangeRef,destination:CellRef,opts?:{pasteType?:"all"|"values"|"formulas"|"formats"}):Promise<{destination:string;cellsCopied:number;}>

type FontStyle={name?:string;size?:number;color?:string;bold?:boolean;italic?:boolean;strike?:boolean;underline?:string;verticalAlign?:string}
type BorderSide={style:string;color:string}
type DiagonalBorder={style:string;color?:string;up?:boolean;down?:boolean}
type StyleObj={
	fill?:{color?:string;pattern?:string;patternColor?:string;gradient?:{type:string;degree?:number;color1:string;color2:string;top?:number;bottom?:number;left?:number;right?:number}};
	font?:FontStyle;
	alignment?:{horizontal?:string;vertical?:string;rotation?:number;wrapText?:boolean;shrinkToFit?:boolean;indent?:number};
	border?:{top?:BorderSide;bottom?:BorderSide;left?:BorderSide;right?:BorderSide;diagonal?:DiagonalBorder};
	numberFormat?:string;
	centerContinuousSpan?:number;
	richText?:{text:string;style?:FontStyle}[];
};
function getStyle(wb,cell:CellRef):Promise<StyleObj>
function setStyle(wb,target:CellRef|RangeRef,style:StyleObj):Promise<void>
interface CfColorScalePoint {type:"formula"|"max"|"min"|"num"|"percent"|"percentile"|"autoMin"|"autoMax";value?:number;formula?:string;color?:string}
interface CfDataBarThreshold {type:"formula"|"max"|"min"|"num"|"percent"|"percentile"|"autoMin"|"autoMax";value?:number;formula?:string}
type CfWritableRuleType="cellValue"|"containsText"|"notContainsText"|"beginsWith"|"endsWith"|"containsBlanks"|"notContainsBlanks"|"containsErrors"|"notContainsErrors"|"expression"|"timePeriod"|"top"|"bottom"|"aboveAverage"|"belowAverage"|"duplicateValues"|"uniqueValues"|"twoColorScale"|"threeColorScale"|"dataBar"
interface CfRuleShared {
	address:string;
	priority?:number;
	stopIfTrue?:boolean;
	style?:StyleObj;
	operator?:"equal"|"notEqual"|"greaterThan"|"greaterThanOrEqual"|"lessThan"|"lessThanOrEqual"|"between"|"notBetween"|"above"|"aboveOrEqual"|"below"|"belowOrEqual";
	formula?:string;
	formula2?:string;
	text?:string;
	rank?:number;
	percent?:boolean;
	bottom?:boolean;
	stdDev?:number;
	lowValue?:CfColorScalePoint;
	midValue?:CfColorScalePoint;
	highValue?:CfColorScalePoint;
	dataBar?:{
		showValue?:boolean;gradient?:boolean;border?:boolean;
		negativeBarColorSameAsPositive?:boolean;negativeBarBorderColorSameAsPositive?:boolean;
		axisPosition?:"automatic"|"middle"|"none";direction?:"context"|"leftToRight"|"rightToLeft";
		fillColor?:string;borderColor?:string;negativeFillColor?:string;negativeBorderColor?:string;axisColor?:string;
		lowValue?:CfDataBarThreshold;highValue?:CfDataBarThreshold;
	};
	timePeriod?:"today"|"yesterday"|"tomorrow"|"last7Days"|"thisWeek"|"lastWeek"|"nextWeek"|"thisMonth"|"lastMonth"|"nextMonth";
}
function getConditionalFormatting(wb,sheetName:string):Promise<Array<CfRuleShared&{index?:number;type:CfWritableRuleType|"iconSet"}>>
function setConditionalFormatting(wb,sheetName:string,rules:Array<CfRuleShared&{type:CfWritableRuleType}>,opts?:{clear?:boolean}):Promise<void>
function removeConditionalFormatting(wb,sheetName:string,indices:number[]):Promise<void>
type DvOperator="Between"|"NotBetween"|"EqualTo"|"NotEqualTo"|"GreaterThan"|"LessThan"|"GreaterThanOrEqualTo"|"LessThanOrEqualTo";
type DvBasic={operator:DvOperator;formula1:string|number;formula2?:string|number|null};
type DvRule={
	wholeNumber?:DvBasic|null;decimal?:DvBasic|null;list?:{source:string;inCellDropDown?:boolean}|null;
	date?:DvBasic|null;time?:DvBasic|null;textLength?:DvBasic|null;custom?:{formula:string}|null;
};
type DvSpec={
	address:string;
	rule:DvRule;
	ignoreBlanks?:boolean;
	prompt?:{showPrompt?:boolean;title?:string|null;message?:string|null}|null;
	errorAlert?:{showAlert?:boolean;style?:"Stop"|"Warning"|"Information";title?:string|null;message?:string|null}|null;
};
function getDataValidations(wb,opts?:{sheet?:string;address?:string}):Promise<Array<DvSpec&{
	index:number;sheet:string;type:"None"|"WholeNumber"|"Decimal"|"List"|"Date"|"Time"|"TextLength"|"Custom";
}>>
function validateCells(wb,address:RangeRef,opts?:{maxCellsToScan?:number;maxInvalidCells?:number;treatUnsupportedAsInvalid?:boolean}):Promise<{
	status:"Valid"|"Invalid"|"NoValidation"|"Mixed"|"Unknown";
	invalidCells:string[];
	truncated:boolean;
	diagnostics:{code:string;message:string;details?:Record<string,string>|null}[];
}>
function setDataValidations(wb,sheetName:string,rules:DvSpec[],opts?:{clear?:boolean}):Promise<void>
function removeDataValidations(wb,sheetName:string,target:{indices:number[]}|{address:string}):Promise<void>

function sweepInputs(wb,args:{
	inputs:{address:CellRef;values:(number|string|boolean|null)[]}[];
	outputs:(string|CellRef)[];
	mode?:"cartesian"|"parallel";
	includeStats?:boolean;
}):Promise<{
	tsv:string;
	sweeps:{inputs:Record<string,string>;outputs:Record<string,string>;errors:Diag[];}[];
	stats?:Record<string,{min:number;max:number;mean:number;count:number}>;
	sweepCount:number;
	inputCount:number;
	outputCount:number;
}>
interface Deps {cells:{address:string;depth:number;formula?:string;referenceType?:"direct"|"range"|"named"|"table"}[];warnings?:Diag[]}
function getCellPrecedents(wb,address:CellRef,depth?:number):Promise<Deps>
function getCellDependents(wb,address:CellRef,depth?:number):Promise<Deps>
function traceToInputs(wb,cell:CellRef):Promise<{address:string;referenceCount:number;text?:string;nearbyLabel?:string;context?:string}[]>
function traceToOutputs(wb,cell:CellRef):Promise<{address:string;formula?:string;text?:string;visibility:Visibility;nearbyLabel?:string;context?:string}[]>
interface FormulaEval {formula:string;value:number|string|boolean|null|unknown[][];error?:{code:string;detail?:string}}
function evaluateFormulas(wb,sheet:string,formulas:string[]):Promise<FormulaEval[]>
function evaluateFormula(wb,sheet:string,formula:string):Promise<FormulaEval>
function lint(wb,options?:{rangeAddresses?:string[];skipRuleIds?:string[];onlyRuleIds?:string[]}):Promise<{
	diagnostics:{severity:"Info"|"Warning"|"Error";ruleId:string;message:string;location:string|null;visibility:Visibility|null}[];
	total:number;
}>
function previewStyles(wb,range:RangeRef):Promise<void>

type TimeUnit="days"|"months"|"years"
type ChartDataLabelPosition="bestFit"|"center"|"insideBase"|"insideEnd"|"outsideEnd"|"left"|"right"|"top"|"bottom"
interface ChartTextSource {text?:string;ref?:string} // Use exactly one of text/ref.
interface ChartPositionAnchor {cell:string;xOffsetPts?:number;yOffsetPts?:number}
interface ChartPosIn {from:ChartPositionAnchor;to:ChartPositionAnchor}
interface ChartPos extends ChartPosIn {sheet:string}
interface ChartFillFormat {noFill?:boolean;color?:string}
interface ChartLineFormat {noLine?:boolean;color?:string;weight?:number;lineStyle?:string}
interface ChartFontFormat {bold?:boolean;color?:string;italic?:boolean;name?:string;size?:number;underline?:string}
interface ChartDataLabelFormat {fill?:ChartFillFormat;border?:ChartLineFormat;font?:ChartFontFormat}
interface ChartPlotAreaSpec {format?:{fill?:ChartFillFormat;border?:ChartLineFormat}}
interface ChartDataLabels {
	showLegendKey?:boolean;showValue?:boolean;showCategory?:boolean;showSeriesName?:boolean;showPercent?:boolean;
	showBubbleSize?:boolean;showLeaderLines?:boolean;position?:ChartDataLabelPosition;
	numberFormat?:string;numberFormatLinked?:boolean;separator?:string;format?:ChartDataLabelFormat;
}
interface ChartAxisSpec {
	title?:ChartTextSource;
	visible?:boolean;
	categoryType?:"category"|"date";
	min?:number;max?:number;majorUnit?:number;minorUnit?:number;
	baseTimeUnit?:TimeUnit;majorTimeUnit?:TimeUnit;minorTimeUnit?:TimeUnit;
	numberFormat?:string;numberFormatLinked?:boolean;reversed?:boolean;
	majorGridlines?:boolean;minorGridlines?:boolean;position?:"left"|"right"|"top"|"bottom";
}
interface ChartSpec {
	name:string;
	position:ChartPosIn;
	groups:{
		type:"column"|"bar"|"line"|"area"|"pie"|"doughnut"|"scatter"|"bubble"|"radar"|"surface"|"stockHLC"|"stockOHLC"|"waterfall"|"histogram"|"pareto"|"funnel";
		scatterStyle?:"line"|"lineMarker"|"marker"|"smooth"|"smoothMarker";
		radarStyle?:"standard"|"marker"|"filled";
		surfaceVariant?:"topView"|"topViewWireframe";
		grouping?:"standard"|"stacked"|"percentStacked";
		axis?:"primary"|"secondary";
		gapWidth?:number;overlap?:number;varyColors?:boolean;smooth?:boolean;firstSliceAngle?:number;holeSize?:number;
		bubbleScale?:number;showNegativeBubbles?:boolean;sizeRepresents?:"area"|"width";
		dataLabels?:ChartDataLabels;
		series:{
			name?:ChartTextSource;
			stockRole?:"volume"|"open"|"high"|"low"|"close";
			categories?:string;categoriesRefType?:"string"|"number"|"multiLevelString";
			values?:string;xValues?:string;yValues?:string;bubbleSizes?:string;
			fillColor?:string;lineColor?:string;lineWidth?:number;lineDashStyle?:string;smooth?:boolean;invertIfNegative?:boolean;
			totalIndexes?:number[];showConnectorLines?:boolean;
			binOptions?:{type?:"auto"|"binCount"|"binWidth"|"category";count?:number;width?:number;allowOverflow?:boolean;overflowValue?:number;allowUnderflow?:boolean;underflowValue?:number};
			marker?:{style?:"auto"|"none"|"circle"|"dash"|"diamond"|"dot"|"picture"|"plus"|"square"|"star"|"triangle"|"x";size?:number;fillColor?:string;borderColor?:string};
			dataLabels?:ChartDataLabels; // bubble position only supports center/left/right/top/bottom
		}[];
	}[];
	title?:ChartTextSource&{overlay?:boolean};
	legend?:{visible?:boolean;position?:"left"|"right"|"top"|"bottom"|"topRight";overlay?:boolean};
	axes?:{category?:ChartAxisSpec;value?:ChartAxisSpec;secondaryCategory?:ChartAxisSpec;secondaryValue?:ChartAxisSpec};
	format?:ChartDataLabelFormat;
	plotArea?:ChartPlotAreaSpec;
	displayBlanksAs?:"gap"|"span"|"zero";
	plotVisibleOnly?:boolean;
	showDataLabelsOverMaximum?:boolean;
	roundedCorners?:boolean;
	styleId?:number;
}
function listCharts(wb,options?:{ sheet?:string }):Promise<Array<{
	id?:number;sheet:string;name:string;type:string;
	groups:{type:string;axis?:string;seriesCount:number}[];
	groupCount:number;seriesCount:number;position:ChartPos;
}>>
function getChart(wb,sheet:string,name:string):Promise<Omit<ChartSpec,"position">&{ position:ChartPos }>
function addChart(wb,sheet:string,chart:ChartSpec):Promise<ChartSpec>
function setChart(wb,sheet:string,name:string,chart:ChartSpec):Promise<ChartSpec>
function deleteChart(wb,sheet:string,name:string):Promise<void>

type ImageFormat="png"|"jpeg"
interface ImagePositionAnchor {cell:string;xOffsetPts?:number;yOffsetPts?:number}
interface ImagePositionInput {from:ImagePositionAnchor;to:ImagePositionAnchor}
interface ImagePosition extends ImagePositionInput {sheet?:string}
interface ImageSource {base64:string}
interface ImageAlt {altText?:string|null;altTextTitle?:string|null}
interface ImagePayload extends ImageAlt {format?:ImageFormat;preserveAspectRatio?:boolean}
interface ImageSpec extends ImagePayload {name:string;position:ImagePositionInput;source:ImageSource}
interface ImageUpdate extends ImagePayload {name?:string;position?:ImagePositionInput;source?:ImageSource}
type ImageInfo=ImageAlt&{id?:number;sheet:string;name:string;position:ImagePosition;format?:ImageFormat;widthPts?:number;heightPts?:number;naturalWidthPx?:number;naturalHeightPx?:number}
function listImages(wb,options?:{ sheet?:string }):Promise<ImageInfo[]>
function getImage(wb,sheet:string,selector:{ name?:string;id?:number }):Promise<ImageInfo>
function addImage(wb,sheet:string,image:ImageSpec):Promise<ImageInfo>
function setImage(wb,sheet:string,selector:{ name?:string;id?:number },image:ImageUpdate):Promise<ImageInfo>
function deleteImage(wb,sheet:string,selector:{ name?:string;id?:number }):Promise<void>
```
