---
name: xlsx-code-mode
description: "Use for any .xls/.xlsx/.xlsm workbook task: read, inspect, search, create, modify, repair, author, or verify spreadsheets. Trigger for workbook creation; edits to cells, formulas, sheets, styles, tables, charts, images, conditional formatting, data validation, names, metadata, or sheet properties; dependency tracing; what-if analysis; lint/calc/render verification; and any casual spreadsheet file reference. Runs sandboxed JavaScript via `witan xlsx exec`."
---

> **Claude Cowork?** `witan` is not preinstalled. Read [references/cowork-setup.md](references/cowork-setup.md).

## Setup

Supports `.xls`, `.xlsx`, `.xlsm`; legacy `.xls` is converted as needed. New workbooks are `.xlsx` only.

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
	groups: [{ type: "column", series: [{ name: { ref: "Sheet1!B1" }, categories: "Sheet1!A2:A9", values: "Sheet1!B2:B9" }] }],
	title: { text: "Revenue" },
	legend: { position: "right" }
})
await xlsx.previewStyles(wb, "Sheet1!F2:N18")
WITAN

# Add an image from a local file
witan xlsx exec model.xlsx --save --input-file logo=@./logo.png --stdin <<'WITAN'
await xlsx.addImage(wb, "Sheet1", {
	name: "Logo",
	position: { from: { cell: "A1" }, to: { cell: "D6" } },
	source: { base64: input.logo }
})
WITAN

# Simple one-liner
witan xlsx exec model.xlsx --expr 'xlsx.listSheets(wb)'
```

## exec — Workbook Scripting

Runs server-side JavaScript against a workbook through globals `xlsx`, `wb`, `input`, and `print`. Top-level `await` is supported; imports are not.

Use `--create` only with new `.xlsx` paths. Add `--save` to write bytes to disk; without `--save`, creation and edits are session-only.

### Invocation patterns

Prefer `--stdin <<'WITAN'` for multi-line scripts and awkward sheet names; the quoted delimiter prevents shell expansion. Use exactly one code source: `--stdin`, `--expr`, `--code`, or `--script`.

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

Use the full type block below for exact signatures. Non-obvious defaults: `scaleRange.skipFormulas` and `sortRange.hasHeader` default true; conditional formatting and data validation `opts.clear` replaces existing rules.

For images, `source.base64` accepts raw base64 or data URLs; responses return metadata only.

### Write and What-If Rules

- `exec` is ephemeral unless `--save` is passed; each invocation starts from the original workbook. `setCells` creates missing referenced sheets.
- Read changed outputs from `result.touched["Sheet!Address"]`; do not recompute answers in JavaScript.
- What-if: separate exec to find the output with `findCells(wb, matcher, { context: 2 })`/synonyms/`readRangeTsv`; choose the formula/output cell, not the label.
- Then separate exec: `traceToInputs(wb, outputAddr)`, confirm the named input drives it, `setCells`, read `result.touched[outputAddr]`, report baseline/new values, check `result.errors`.
- Use `sweepInputs` for multi-value sensitivities. Filter large traces by `nearbyLabel`/context. Iterative models report convergence errors through `setCells`.

### calc — Full-workbook verification

`setCells` already recalculates edited cells and downstream dependents. Use `calc` for standalone workbook-wide checks or final audits:

```bash
witan xlsx calc model.xlsx [--verify] [--show-touched]
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

- Code source/import errors: use exactly one non-empty code source; no imports, use `xlsx`.
- Syntax/runtime/large-result/flag errors: fix script or flag; return less data, print/log large output.
- Workbook lookup issues: `Sheet 'X' not found` -> `listSheets`; shell quoting -> `--stdin <<'WITAN'`.
- Search/touched issues: broaden `findCells`, use synonyms/nearby ranges; if output missing from `touched`, trace the chain or pick the right output cell.

### Full Type Definitions

```ts
type CellRef=string|{sheet:string;row:number;col:number|string}
type RangeRef=string|{sheet:string}|{sheet:string;from:{row?:number;col?:number|string};to:{row?:number;col?:number|string}}
type Visibility="visible"|"outsidePrintArea"|"collapsed"|"hidden"
type Diag={code:string;detail?:string;address:string;formula?:string}
type PartialDeep<T>={[K in keyof T]?:NonNullable<T[K]> extends object?PartialDeep<NonNullable<T[K]>>:T[K]}

type WorkbookProperties={
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

type NameDef={name:string;range:string;scope:string|null}
function listDefinedNames(wb):Promise<NameDef[]>
function addDefinedName(wb,name:string,range:string,scope?:string):Promise<NameDef>
function deleteDefinedName(wb,name:string,scope?:string):Promise<NameDef>
function addSheet(wb,name:string):Promise<string>
function deleteSheet(wb,name:string):Promise<void>
function renameSheet(wb,oldName:string,newName:string):Promise<void>

type SheetVisibility="visible"|"hidden"|"veryHidden";
type SheetView={showGridLines:boolean;zoomScale:number}
type SheetOutline={summaryRowsBelow:boolean;summaryColumnsRight:boolean;showSymbols:boolean}
type SheetFormat={defaultRowHeight:number;defaultColWidth:number;font?:{name?:string;size?:number}|null}
type RowProps={height?:number;hidden?:boolean;outlineLevel?:number;collapsed?:boolean}
type ColProps={width?:number;hidden?:boolean;outlineLevel?:number;collapsed?:boolean}
type SheetProperties={
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

type CellAddr={
	address:string;sheet:string;row:number;col:number;colLetter:string;visibility:Visibility;
}
type CellHit=CellAddr&{value:any;text:string}
type Value=CellAddr&{value:string|number|boolean|null;formula?:string;type:"string"|"number"|"bool"|"date"|"error"|"blank";text:string;format?:string;numberType?:"currency"|"percent"|"fraction"|"exponential"|"date"|"text"|"number";context?:string;note?:{author:string;text:string};link?:{type:"internal"|"external";target:string;tooltip?:string};thread?:{resolved:boolean;comments:{authorId:string;text:string;createdAt:string}[]}}
function readCell(wb,cell:CellRef,opts?:{context?:number}):Promise<Value>
function readRange(wb,range:RangeRef):Promise<Value[][]>
function readColumn(wb,sheetName:string,col:number|string,opts?:{startRow?:number;endRow?:number}):Promise<Value[]>
function readRow(wb,sheetName:string,row:number,opts?:{startCol?:number;endCol?:number}):Promise<Value[]>
function readRangeTsv(wb,range:RangeRef,opts?:{includeEmpty?:boolean;includeFormulas?:boolean}):Promise<string>
function readColumnTsv(wb,sheetName:string,col:number|string,opts?:{startRow?:number;endRow?:number;includeEmpty?:boolean;includeFormulas?:boolean}):Promise<string>
function readRowTsv(wb,sheetName:string,row:number,opts?:{startCol?:number;endCol?:number;includeEmpty?:boolean;includeFormulas?:boolean}):Promise<string>
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

type UsedRange={startRow:number;startCol:number;endRow:number;endCol:number}
type TileRef={sheet:string;tileRow:number;tileCol:number}
type SheetUpdate={name:string;usedRange:UsedRange|null;tileRowCount:number;tileColCount:number}
type WriteResult={touched:Record<string,string>;changed:string[];errors:Diag[];invalidatedTiles:TileRef[];updatedSheets:SheetUpdate[]}
type CellPatch={address:CellRef;value?:unknown;formula?:string;format?:string;note?:{text?:string;author?:string}|null;link?:{url?:string;ref?:string;tooltip?:string}|null;thread?:{add?:Array<{author?:string;text:string}>;resolved?:boolean;delete?:boolean}|null}
function setCells(wb,cells:CellPatch[],opts?:{validationMode?:"ignore"|"reject"}):Promise<WriteResult>
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
type CfThreshold="formula"|"max"|"min"|"num"|"percent"|"percentile"|"autoMin"|"autoMax"
type CfColorPoint={type:CfThreshold;value?:number;formula?:string;color?:string}
type CfBarThreshold={type:CfThreshold;value?:number;formula?:string}
type CfRuleType="cellValue"|"containsText"|"notContainsText"|"beginsWith"|"endsWith"|"containsBlanks"|"notContainsBlanks"|"containsErrors"|"notContainsErrors"|"expression"|"timePeriod"|"top"|"bottom"|"aboveAverage"|"belowAverage"|"duplicateValues"|"uniqueValues"|"twoColorScale"|"threeColorScale"|"dataBar"
type CfOp="equal"|"notEqual"|"greaterThan"|"greaterThanOrEqual"|"lessThan"|"lessThanOrEqual"|"between"|"notBetween"|"above"|"aboveOrEqual"|"below"|"belowOrEqual"
type CfPeriod="today"|"yesterday"|"tomorrow"|"last7Days"|"thisWeek"|"lastWeek"|"nextWeek"|"thisMonth"|"lastMonth"|"nextMonth"
type CfBar={showValue?:boolean;gradient?:boolean;border?:boolean;negativeBarColorSameAsPositive?:boolean;negativeBarBorderColorSameAsPositive?:boolean;axisPosition?:"automatic"|"middle"|"none";direction?:"context"|"leftToRight"|"rightToLeft";fillColor?:string;borderColor?:string;negativeFillColor?:string;negativeBorderColor?:string;axisColor?:string;lowValue?:CfBarThreshold;highValue?:CfBarThreshold}
type CfRule={address:string;priority?:number;stopIfTrue?:boolean;style?:StyleObj;operator?:CfOp;formula?:string;formula2?:string;text?:string;rank?:number;percent?:boolean;bottom?:boolean;stdDev?:number;lowValue?:CfColorPoint;midValue?:CfColorPoint;highValue?:CfColorPoint;dataBar?:CfBar;timePeriod?:CfPeriod}
function getConditionalFormatting(wb,sheetName:string):Promise<Array<CfRule&{index?:number;type:CfRuleType|"iconSet"}>>
function setConditionalFormatting(wb,sheetName:string,rules:Array<CfRule&{type:CfRuleType}>,opts?:{clear?:boolean}):Promise<void>
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
type Deps={cells:{address:string;depth:number;formula?:string;referenceType?:"direct"|"range"|"named"|"table"}[];warnings?:Diag[]}
function getCellPrecedents(wb,address:CellRef,depth?:number):Promise<Deps>
function getCellDependents(wb,address:CellRef,depth?:number):Promise<Deps>
function traceToInputs(wb,cell:CellRef):Promise<{address:string;referenceCount:number;text?:string;nearbyLabel?:string;context?:string}[]>
function traceToOutputs(wb,cell:CellRef):Promise<{address:string;formula?:string;text?:string;visibility:Visibility;nearbyLabel?:string;context?:string}[]>
type FormulaEval={formula:string;value:number|string|boolean|null|unknown[][];error?:{code:string;detail?:string}}
function evaluateFormulas(wb,sheet:string,formulas:string[]):Promise<FormulaEval[]>
function evaluateFormula(wb,sheet:string,formula:string):Promise<FormulaEval>
function lint(wb,options?:{rangeAddresses?:string[];skipRuleIds?:string[];onlyRuleIds?:string[]}):Promise<{
	diagnostics:{severity:"Info"|"Warning"|"Error";ruleId:string;message:string;location:string|null;visibility:Visibility|null}[];
	total:number;
}>
function previewStyles(wb,range:RangeRef):Promise<void>

type TUnit="days"|"months"|"years"
type Axis="left"|"right"|"top"|"bottom"
type CDLPos="bestFit"|"center"|"insideBase"|"insideEnd"|"outsideEnd"|"left"|"right"|"top"|"bottom"
type CType="column"|"bar"|"line"|"area"|"pie"|"doughnut"|"scatter"|"bubble"|"radar"|"surface"|"stockHLC"|"stockOHLC"|"waterfall"|"histogram"|"pareto"|"funnel"|"boxWhisker"
type CTxt={text?:string;ref?:string}
type CAnchor={cell:string;xOffsetPts?:number;yOffsetPts?:number}
type CPosIn={from:CAnchor;to:CAnchor}
type CPos=CPosIn&{sheet:string}
type CFill={noFill?:boolean;color?:string}
type CLine={noLine?:boolean;color?:string;weight?:number;lineStyle?:string}
type CFont={bold?:boolean;color?:string;italic?:boolean;name?:string;size?:number;underline?:string}
type CLabelFmt={fill?:CFill;border?:CLine;font?:CFont}
type CPlot={format?:{fill?:CFill;border?:CLine}}
type CLabels={showLegendKey?:boolean;showValue?:boolean;showCategory?:boolean;showSeriesName?:boolean;showPercent?:boolean;showBubbleSize?:boolean;showLeaderLines?:boolean;position?:CDLPos;numberFormat?:string;numberFormatLinked?:boolean;separator?:string;format?:CLabelFmt}
type CAxis={title?:CTxt;visible?:boolean;categoryType?:"category"|"date";min?:number;max?:number;majorUnit?:number;minorUnit?:number;baseTimeUnit?:TUnit;majorTimeUnit?:TUnit;minorTimeUnit?:TUnit;numberFormat?:string;numberFormatLinked?:boolean;reversed?:boolean;majorGridlines?:boolean;minorGridlines?:boolean;position?:Axis}
type CBins={type?:"auto"|"binCount"|"binWidth"|"category";count?:number;width?:number;allowOverflow?:boolean;overflowValue?:number;allowUnderflow?:boolean;underflowValue?:number}
type CSeries={name?:CTxt;stockRole?:"volume"|"open"|"high"|"low"|"close";categories?:string;categoriesRefType?:"string"|"number"|"multiLevelString";values?:string;xValues?:string;yValues?:string;bubbleSizes?:string;fillColor?:string;lineColor?:string;lineWidth?:number;lineDashStyle?:string;smooth?:boolean;invertIfNegative?:boolean;totalIndexes?:number[];showConnectorLines?:boolean;binOptions?:CBins;quartileCalculation?:"exclusive"|"inclusive";showInnerPoints?:boolean;showMeanLine?:boolean;showMeanMarker?:boolean;showOutlierPoints?:boolean;marker?:{style?:"auto"|"none"|"circle"|"dash"|"diamond"|"dot"|"picture"|"plus"|"square"|"star"|"triangle"|"x";size?:number;fillColor?:string;borderColor?:string};dataLabels?:CLabels}
type CGroup={type:CType;scatterStyle?:"line"|"lineMarker"|"marker"|"smooth"|"smoothMarker";radarStyle?:"standard"|"marker"|"filled";surfaceVariant?:"topView"|"topViewWireframe";grouping?:"standard"|"stacked"|"percentStacked";axis?:"primary"|"secondary";gapWidth?:number;overlap?:number;varyColors?:boolean;smooth?:boolean;firstSliceAngle?:number;holeSize?:number;bubbleScale?:number;showNegativeBubbles?:boolean;sizeRepresents?:"area"|"width";dataLabels?:CLabels;series:CSeries[]}
type ChartSpec={
	name:string;
	position:CPosIn;
	groups:CGroup[];
	title?:CTxt&{overlay?:boolean};
	legend?:{visible?:boolean;position?:Axis|"topRight";overlay?:boolean};
	axes?:{category?:CAxis;value?:CAxis;secondaryCategory?:CAxis;secondaryValue?:CAxis};
	format?:CLabelFmt;
	plotArea?:CPlot;
	displayBlanksAs?:"gap"|"span"|"zero";
	plotVisibleOnly?:boolean;
	showDataLabelsOverMaximum?:boolean;
	roundedCorners?:boolean;
	styleId?:number;
}
function listCharts(wb,options?:{ sheet?:string }):Promise<Array<{
	id?:number;sheet:string;name:string;type:string;
	groups:{type:string;axis?:string;seriesCount:number}[];
	groupCount:number;seriesCount:number;position:CPos;
}>>
function getChart(wb,sheet:string,name:string):Promise<Omit<ChartSpec,"position">&{ position:CPos }>
function addChart(wb,sheet:string,chart:ChartSpec):Promise<ChartSpec>
function setChart(wb,sheet:string,name:string,chart:ChartSpec):Promise<ChartSpec>
function deleteChart(wb,sheet:string,name:string):Promise<void>

type ImageFormat="png"|"jpeg"
type ImagePositionAnchor={cell:string;xOffsetPts?:number;yOffsetPts?:number}
type ImagePositionInput={from:ImagePositionAnchor;to:ImagePositionAnchor}
type ImagePosition=ImagePositionInput&{sheet?:string}
type ImageSource={base64:string}
type ImageAlt={altText?:string|null;altTextTitle?:string|null}
type ImagePayload=ImageAlt&{format?:ImageFormat;preserveAspectRatio?:boolean}
type ImageSelector={name?:string;id?:number}
type ImageSpec=ImagePayload&{name:string;position:ImagePositionInput;source:ImageSource}
type ImageUpdate=ImagePayload&{name?:string;position?:ImagePositionInput;source?:ImageSource}
type ImageInfo=ImageAlt&{id?:number;sheet:string;name:string;position:ImagePosition;format?:ImageFormat;widthPts?:number;heightPts?:number;naturalWidthPx?:number;naturalHeightPx?:number}
function listImages(wb,options?:{ sheet?:string }):Promise<ImageInfo[]>
function getImage(wb,sheet:string,selector:ImageSelector):Promise<ImageInfo>
function addImage(wb,sheet:string,image:ImageSpec):Promise<ImageInfo>
function setImage(wb,sheet:string,selector:ImageSelector,image:ImageUpdate):Promise<ImageInfo>
function deleteImage(wb,sheet:string,selector:ImageSelector):Promise<void>
```
