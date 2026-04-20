---
name: xlsx-code-mode
description: Use this skill any time an Excel file (.xls, .xlsx, .xlsm) needs to be read, explored, understood, or modified. You cannot read excel files with cat, head, or normal file-reading tools — this is the only way to inspect them. Trigger when you or the user need to open, look at, or explore a workbook; find out what sheets it has or where specific data lives; read cells, rows, columns, or ranges; search for values, labels, or patterns; trace formula dependencies or understand how a cell is calculated; run what-if scenarios by changing inputs and reading recalculated outputs; or edit cells, rows, columns, and sheets. Trigger when the user references a spreadsheet file by name or path — even casually (e.g. 'check the xlsx', 'what's in report.xlsx') — and also when you need to inspect a workbook yourself as part of a larger task. The tool runs sandboxed JavaScript against the workbook server-side via `witan xlsx exec`."
---

## Setup

Files are cached server-side by content hash so repeated operations skip re-upload.

The CLI automatically applies per-attempt request timeouts and retries transient API failures (`408`, `429`, `500`, `502`, `503`, `504`, plus timeout/network errors). Non-retryable `4xx` responses fail immediately.

The CLI automatically converts older .xls files to .xlsx, so it fully supports all Excel file formats.

## Quick Reference

```bash
# Create a new workbook from scratch (.xlsx only)
witan xlsx exec model.xlsx --create --save --stdin <<'WITAN'
await xlsx.addSheet(wb, "Inputs")
await xlsx.setCells(wb, [{ address: "Inputs!A1", value: "Revenue" }])
return await xlsx.listSheets(wb)
WITAN

# Read from sheets with spaces, apostrophes, or parentheses — all safe
# Note: inner apostrophes in a sheet name must be doubled (Excel convention).
# "Workers' Compensation" → "'Workers'' Compensation'!B50"
witan xlsx exec model.xlsx --stdin <<'WITAN'
const a = await xlsx.readCell(wb, "'Workers'' Compensation'!B50")
const b = await xlsx.readRangeTsv(wb, { sheet: "Reserve Summary (Net)", from: {row:1,col:1}, to: {row:10,col:5} })
return { a: a.value, b }
WITAN

# Multi-input sweep — compare several input values at once
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

# Conditional formatting — add a highlight rule and a color scale
witan xlsx exec model.xlsx --stdin <<'WITAN'
await xlsx.setConditionalFormatting(wb, "Sheet1", [
	{
		type: "cellValue",
		address: "A1:A100",
		operator: "greaterThan",
		formula: "100",
		style: { fill: { color: "#FF0000" } }
	},
	{
		type: "twoColorScale",
		address: "B1:B100",
		lowValue: { type: "min", color: "#FFFFFF" },
		highValue: { type: "max", color: "#FF0000" }
	}
])
WITAN

# Chart authoring — add an embedded chart and verify the rendered result
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

# ListObject authoring — create an Excel table and read it back by table name
witan xlsx exec model.xlsx --stdin <<'WITAN'
await xlsx.addListObject(wb, "Sheet1", {
	name: "SalesTable",
	ref: "A1:C4",
	showTotalsRow: true,
	columns: [
		{ name: "Region", totalsRowLabel: "Total" },
		{ name: "Sales", totalsRowFunction: "sum" },
		{ name: "DoubleSales", calculatedColumnFormula: "=B2*2" }
	],
	rows: [
		[{ value: "North" }, { value: 10 }, {}],
		[{ value: "South" }, { value: 20 }, {}]
	]
})
return {
	meta: await xlsx.getListObject(wb, "SalesTable"),
	data: await xlsx.readRange(wb, "SalesTable")
}
WITAN

# What-If Data Table authoring — create a visible one-variable table block
witan xlsx exec model.xlsx --stdin <<'WITAN'
await xlsx.addDataTable(wb, "Sheet1", {
	type: "oneVariableColumn",
	ref: "E1:F4",
	columnInputCell: "H1",
	inputValues: [5, 10, 15],
	formulas: ["=H1*2"]
})
return await xlsx.getDataTable(wb, "Sheet1!E1:F4")
WITAN

# Simple one-liner (--expr is fine when there are no special characters)
witan xlsx exec model.xlsx --expr 'xlsx.listSheets(wb)'
```

## Exit Codes

- `exec` `0`: script completed successfully (`ok: true`)
- `exec` `1`: transport/API error, invalid request, or script error (`ok: false`)

## exec — Workbook Scripting

Runs JavaScript against a workbook via the Witan API. The workbook is opened server-side; scripts interact through the `xlsx` and `wb` globals.

If the target workbook does not exist yet, use `--create` with a new `.xlsx` path.

- Use `--create --save` when you want to produce a real workbook file on disk.
- Use `--create` without `--save` when you want to prototype workbook structure, test generation logic, inspect returned data, or validate formulas/layout without leaving a file behind.

### Invocation patterns

**Recommended: `--stdin` with heredoc** — safe for all sheet names, supports multi-line scripts, and batches multiple operations into a single CLI invocation:

```bash
witan xlsx exec report.xlsx --stdin <<'WITAN'
const sheets = await xlsx.listSheets(wb)
const cell = await xlsx.readCell(wb, "'My Sheet'!A1")
return { sheets, cell }
WITAN
```

The single-quoted heredoc delimiter (`<<'WITAN'`) prevents all shell expansion. Apostrophes, parentheses, double quotes, and glob characters in sheet names pass through verbatim to JavaScript — no escaping needed.

**Other invocation patterns** (use only when `--stdin` is impractical):

```bash
# Expression — simple one-liners with no special characters
witan xlsx exec report.xlsx --expr 'xlsx.listSheets(wb)'

# Script file — reusable scripts, e.g. parameterized scenarios
witan xlsx exec report.xlsx --script scenario.js --input-json '{"rate": 1.05}'
```

Provide exactly one code source: `--expr`, `--code`, `--script`, or `--stdin`. They are mutually exclusive.

### Flags

Exec-specific:

- `--expr`: expression shorthand; wraps as `return (<expr>);`
- `--code`: inline JavaScript source
- `--script`: path to a JavaScript file
- `--stdin`: read JavaScript source from stdin
- `--input-json` (default `{}`): JSON value passed as `input`
- `--timeout-ms`: execution timeout in milliseconds (> 0); omit for server default
- `--max-output-chars`: maximum stdout characters to capture (> 0); omit for server default
- `--stdin-timeout-ms` (default `2000`): abort `--stdin` reads that never reach EOF; `0` disables
- `--locale`: execution locale; falls back to `WITAN_LOCALE`, then `LC_ALL` / `LC_MESSAGES` / `LANG`
- `--create` (default `false`): create a new `.xlsx` workbook; target path must not exist
- `--save` (default `false`): persist changes to the workbook file

Global (apply to every `witan` subcommand):

- `--json`: print the full response envelope as JSON instead of the human summary

### Runtime globals

- `xlsx` (`object`): curated API surface; all functions listed below
- `wb` (`WorkbookContext`): opened workbook handle; pass as first arg to all `xlsx.*` calls
- `input` (`any`): parsed value from `--input-json` (defaults to `{}`)
- `print` (`function`): output to stdout, like `console.log` but captured in the response

Top-level `await` is supported. No imports allowed (static or dynamic).

### API reference

Functions are grouped by purpose. All are async and take `wb` as the first argument.

**Reading**

- `listSheets`: sheet inventory with used ranges, visibility, and cross-sheet dependencies
- `getWorkbookProperties`, `getSheetProperties`: workbook/sheet metadata, formatting, view, outline, and merge info
- `listDefinedNames`: workbook and sheet-scoped names
- `readCell`, `readRange`, `readRow`, `readColumn`: structured cell reads
- `readRangeTsv`, `readRowTsv`, `readColumnTsv`: TSV reads for prompt-friendly extraction
- `getStyle`: resolved style properties for a cell

**Searching**

- `findCells`, `findRows`: fuzzy search by value or pattern
- `describeSheet`: sheet structure map with detected tables for a single sheet
- `tableLookup`: lookup by row and column labels inside a table
- `getListObject`, `getDataTable`: metadata for existing Excel table / What-If Data Table objects

`matcher` may be a string, string array, number, boolean, `RegExp`, or `RegExp[]`. Searches are fuzzy and case-insensitive by default.

**Tracing**

- `getCellPrecedents`, `getCellDependents`: local dependency traversal; `depth` defaults to `1`
- `traceToInputs`, `traceToOutputs`: full graph tracing to leaf inputs or terminal outputs

**Computing**

- `sweepInputs`: batch what-if sweeps with TSV and structured outputs
- `evaluateFormula`, `evaluateFormulas`: evaluate one or more formulas in sheet context

**Validating**

- `lint`: potential workbook issues and diagnostics

**Rendering**

- `previewStyles`: generate a PNG screenshot for a range; image is auto-registered

**Charts**

- `listCharts`: chart summaries for the workbook or a single sheet
- `getChart`: canonical spec for an existing chart
- `addChart`, `setChart`, `deleteChart`: create, replace, or remove embedded charts

**Conditional Formatting**

- `getConditionalFormatting`: read all sheet rules (returns `iconSet` rules too)
- `setConditionalFormatting`: author rules of any `CfWritableRuleType` (see type def); `opts.clear` replaces all existing rules. `iconSet` is not a writable type — if you need icon-style thresholds, author a `threeColorScale` instead.
- `removeConditionalFormatting`: remove rules by index

**Writing (ephemeral)**

- Cells/ranges: `setCells`, `findAndReplace`, `scaleRange`, `copyRange`, `sortRange`
- Structure: `insertRowAfter`, `deleteRows`, `insertColumnAfter`, `deleteColumns`, `autoFitColumns`
- Sheets/names: `addSheet`, `deleteSheet`, `renameSheet`, `addDefinedName`, `deleteDefinedName`
- Properties/styles: `setWorkbookProperties`, `setSheetProperties`, `setRowProperties`, `setColumnProperties`, `setStyle`
- Objects: `addListObject`, `setListObject`, `deleteListObject`, `addDataTable`, `deleteDataTable`

Common option patterns:

- Search: `opts.in`, `context`, `limit`, `offset`; `findCells` also supports `formulas`
- Row/column reads: `startRow/endRow` or `startCol/endCol`
- TSV reads: `includeEmpty`, `includeFormulas`
- `scaleRange`: `opts.skipFormulas` defaults to `true`
- `sortRange`: `opts.hasHeader` defaults to `true`
- `copyRange`: `opts.pasteType` supports `all`, `values`, `formulas`, `formats`
- `setConditionalFormatting`: `opts.clear` replaces all existing rules

### The ephemeral write contract

By default, `exec` **does not write workbook bytes back to disk**. All write operations (`setCells`, `scaleRange`, inserts, deletes) take effect in the server-side session only. The `result.touched` map contains the recalculated formatted text values — read answers from there.

This means:

- No risk of corrupting the original file
- No `reset()` needed — each invocation starts clean
- For independent what-if tests, each `exec` invocation starts from the original file

To persist changes back to the workbook file, pass the `--save` flag.

When `--create` is set, the same ephemeral rule applies to the new workbook session: no local file is created unless `--save` is also passed. New workbook creation is `.xlsx` only.

### setCells result shape

```ts
{
  touched: Record<string, string>      // address → formatted text value
  changed: string[]                    // addresses whose values changed
  errors: Diag[]                       // cells that errored after recalc
  invalidatedTiles: { sheet, tileRow, tileCol }[]   // server-side render invalidations
  updatedSheets: { name, usedRange, tileRowCount, tileColCount }[]  // post-write sheet state
}
```

Read the output value from `result.touched["Sheet!Address"]`. Never compute the answer in JavaScript. `invalidatedTiles` and `updatedSheets` are informational — they exist so callers can refresh UI tiles — you can usually ignore them. `setCells` also implicitly creates a sheet if `address` references one that does not yet exist.

### What-if / sensitivity workflow

For questions like "what happens to Y if X changes?", follow this sequence. **Steps 1 and 2 should be separate `exec` calls** — review the output of step 1 before proceeding.

1. **Find the output cell (separate exec call)** — search for what the user is asking about (e.g., "net income", "reserve"). Use `xlsx.findCells` with `context: 2` and review the candidates. Labels and formula cells often share the same text — pick the formula cell, not the label. This step may take more than one attempt: try synonym arrays, read surrounding rows with `xlsx.readRangeTsv`, or check multiple sheets.
2. **Trace + run the what-if (second exec call)** — once you have the output address, call `xlsx.traceToInputs(wb, outputAddr)` to confirm the user's input actually drives it. Trace results can contain hundreds of cells — filter by `nearbyLabel` matching the user's term instead of printing them all. Then `xlsx.setCells` to make the change and read the answer from `result.touched[outputAddr]`. If the output address is missing from `touched`, the cell didn't recalculate — you likely have the wrong address.
3. **Report before and after** — always include the baseline value (read before the edit) and the new value from `touched`.

For sweeping multiple values (sensitivity tables), use `sweepInputs` instead — it runs all combinations in one call and returns structured before/after data.

Practical tips:

- Don't search for the output _after_ `setCells` — find it first so you know the exact address to check in `touched`.
- If `findCells` returns several hits for the same label, use context or `readRangeTsv` to disambiguate (e.g., "Net Income" may appear as both a label and a formula cell on different rows).
- After `setCells`, verify `result.errors` is empty. New errors mean the edit introduced or surfaced a calculation problem downstream.
- If the question names a specific metric (e.g., "loss ratio"), don't just search for that term — also check the sheet/row the question references, since models often have multiple versions of the same metric across sheets.

### Iterative / circular models

When a workbook has **iterative calculation** enabled, `setCells` recalculates
the edited cells and downstream dependents using the workbook's iterative
settings. If the model does not converge within `maxIterations`, the result
includes convergence errors.

Use `witan xlsx calc` when you want a standalone seeded or full-workbook
verification pass, or when you want CLI reporting of all calculation errors or
changed cells. `--show-touched` only changes output verbosity.

### calc — Full-workbook verification

`setCells` already recalculates the edited cells and all downstream dependents,
so you do **not** need `witan xlsx calc` after every normal edit. Use `calc`
mainly as a standalone verification/reporting command:

- See every formula error in one CLI call
- In `--verify` mode, see which cell values changed without modifying the file
- Run a workbook-wide verification pass outside your `exec` script

```bash
# Recalculate the workbook and print all formula errors, if any
witan xlsx calc model.xlsx

# Verification pass only — no file write, but prints changed addresses
witan xlsx calc model.xlsx --verify

# Verbose output — every touched cell with formula/value or error code
witan xlsx calc model.xlsx --show-touched

# Seed recalc from one or more ranges; downstream dependents still recalculate.
# Useful when you want to force a partial recalculation scope.
witan xlsx calc model.xlsx -r "Inputs!B1:B20" -r "Summary!A1:H10"
```

Flags:

- `-r, --range` (repeatable): sheet-qualified range to seed recalculation from
- `--verify`: no file write; exits `2` if errors exist or any computed value changed
- `--show-touched`: verbose; print every recalculated cell

Default output is concise:

- If there are no formula errors, it prints a one-line summary like `428 cells recalculated, 0 errors, 3 changed`
- If there are formula errors, it prints **all** of them, not just a count
- Without `--verify`, it still exits with code `2` if any formula errors are found
- With `--verify`, it also prints a sorted `Changed (N):` list of addresses
- `--show-touched` is the verbose mode when you need every recalculated cell

Example error output:

```text
2 errors:
  Summary!C18          =A18/B18                      #DIV/0!
  Revenue!F42          =VLOOKUP(A42,$A$2:$C$10,3,0) #N/A
```

Example verify output:

```text
428 cells recalculated, 0 errors, 3 changed

Changed (3):
  Inputs!B5
  Revenue!F42
  Summary!C18
```

`witan xlsx calc` exits with code `2` if any formula errors are found.
`witan xlsx calc --verify` also exits with code `2` if any computed values
changed. That makes `--verify` useful as a final audit step after fixing a
workbook.

### Response format

When `--json` is used, the full response envelope is returned:

**Success:**

```json
{
  "ok": true,
  "stdout": "...",
  "result": "<json>",
  "writes_detected": false,
  "accesses": [...]
}
```

`writes_detected` and `accesses` are only present when the script performs writes or tracked accesses — a pure-read script returns just `{ ok, stdout, result }`. With `--save` in files-backed mode, a successful write also adds `revision_id`.

**Failure:**

```json
{
  "ok": false,
  "stdout": "...",
  "error": { "type": "...", "code": "...", "message": "..." }
}
```

The `accesses` array documents all cell reads and writes with operation type and address.

## render — Visual Screenshot

Renders a sheet range as a PNG image, useful for inspecting layout, merged cells, formatting, and labels.

```bash
# Render a range to a temporary file (path printed to stdout)
witan xlsx render report.xlsx -r "Sheet1!A1:Z50"

# Render to a specific output path
witan xlsx render report.xlsx -r "'My Sheet'!B5:H20" -o snapshot.png

# Higher resolution (DPR 1-3, default auto)
witan xlsx render report.xlsx -r "Sheet1!A1:F10" --dpr 2

# Diff against a baseline — highlights changes in a new PNG
witan xlsx render report.xlsx -r "Sheet1!A1:F10" --diff before.png
```

- `--range`, `-r`: sheet-qualified range to render; required
- `--output`, `-o`: output path; defaults to a temporary file
- `--dpr` (default `auto`): device pixel ratio `1`-`3`
- `--format` (default `png`): `png` or `webp`
- `--diff`: compare against a baseline PNG and write a diff image

The `previewStyles` exec function (see Rendering in the API reference) provides the same capability from within a script.

## Error Guide

- `exactly one of --code, --script, --stdin, or --expr is required`: provide exactly one code source flag
- `--code, --script, --stdin, and --expr are mutually exclusive`: use only one code source flag per invocation
- `exec code must not be empty`: provide non-empty code
- `Import statements are not allowed`: no `import` in exec scripts; use the `xlsx` global
- `EXEC_SYNTAX_ERROR`: fix JavaScript syntax in your script
- `EXEC_RUNTIME_ERROR`: fix the runtime error; check the message for details
- `EXEC_RESULT_TOO_LARGE`: return less data; use `console.log()` for large output instead of return values
- `--timeout-ms must be > 0`: omit the flag or provide a positive value
- `invalid --input-json`: provide valid JSON
- `Sheet 'X' not found`: check the sheet name; use `listSheets` to enumerate
- `Shell quoting errors with sheet names`: use `--stdin <<'WITAN'`; it avoids shell quoting issues
- `findCells` returns empty: try synonym arrays, broader search, or check spelling
- `setCells` result missing expected output: the output cell may not be a dependent; trace the formula chain

### Full Type Definitions

````ts
type CellRef=string|{sheet:string;row:number;col:number|string}
type RangeRef=string|{sheet:string}|{sheet:string;from:{row?:number;col?:number|string};to:{row?:number;col?:number|string}}
type Visibility="visible"|"outsidePrintArea"|"collapsed"|"hidden"
function getWorkbookProperties(wb):Promise<{
	activeSheetIndex:number;
	defaultFont:{
		name:string;
		size:number;
	};
	iterativeCalculation:{
		enabled:boolean;
		maxIterations:number;
		maxChange:number;
	};
	metadata?:{
		author?:string;
		title?:string;
		subject?:string;
		company?:string;
		created?:string;
		modified?:string;
	};
	themeColors?:{
		dark1:string;
		light1:string;
		dark2:string;
		light2:string;
		accent1:string;
		accent2:string;
		accent3:string;
		accent4:string;
		accent5:string;
		accent6:string;
		hyperlink:string;
		followedHyperlink:string;
		majorFont?:string;
		minorFont?:string;
	};
}>;
function setWorkbookProperties(wb,properties:{
	activeSheetIndex?:number;
	defaultFont?:{
		name?:string;
		size?:number;
	};
	iterativeCalculation?:{
		enabled?:boolean;
		maxIterations?:number;
		maxChange?:number;
	};
	metadata?:{
		author?:string;
		title?:string;
		subject?:string;
		company?:string;
	};
	themeColors?:{
		dark1?:string;
		light1?:string;
		dark2?:string;
		light2?:string;
		accent1?:string;
		accent2?:string;
		accent3?:string;
		accent4?:string;
		accent5?:string;
		accent6?:string;
		hyperlink?:string;
		followedHyperlink?:string;
		majorFont?:string;
		minorFont?:string;
	};
}):Promise<void>;
function listSheets(wb):Promise<Array<{
	address:string;
	from:{row:number;col:number};
	to:{row:number;col:number};
	rows:number;
	cols:number;
	sheet:string;
	hidden?:boolean;
	printArea?:string;
	listObjects?:string[];
	dataTables?:string[];
	precedents?:string[];
	dependents?:string[];
}>>;
type TimeUnit="days"|"months"|"years"
/** Use exactly one of `text` or `ref`. */
interface ChartTextSource {text?:string;ref?:string}
interface ChartPositionAnchor {cell:string;xOffsetPts?:number;yOffsetPts?:number}
interface ChartPosIn {from:ChartPositionAnchor;to:ChartPositionAnchor}
interface ChartPos extends ChartPosIn {sheet:string}
interface ChartAxisSpec {
	title?:ChartTextSource;
	visible?:boolean;
	categoryType?:"category"|"date";
	min?:number;
	max?:number;
	majorUnit?:number;
	minorUnit?:number;
	baseTimeUnit?:TimeUnit;
	majorTimeUnit?:TimeUnit;
	minorTimeUnit?:TimeUnit;
	numberFormat?:string;
	reversed?:boolean;
	majorGridlines?:boolean;
	minorGridlines?:boolean;
	position?:"left"|"right"|"top"|"bottom";
}
interface ChartSpec {
	name:string;
	position:ChartPosIn;
	groups:{
		type:"column"|"bar"|"line"|"area"|"pie"|"doughnut"|"scatter"|"bubble";
		grouping?:"standard"|"stacked"|"percentStacked";
		axis?:"primary"|"secondary";
		gapWidth?:number;
		overlap?:number;
		varyColors?:boolean;
		smooth?:boolean;
		firstSliceAngle?:number;
		holeSize?:number;
		series:{
			name?:ChartTextSource;
			categories?:string;
			categoriesRefType?:"string"|"number";
			values?:string;
			xValues?:string;
			yValues?:string;
			bubbleSizes?:string;
			fillColor?:string;
			lineColor?:string;
			lineWidth?:number;
			lineDashStyle?:string;
			smooth?:boolean;
			marker?:{
				style?:"auto"|"none"|"circle"|"dash"|"diamond"|"dot"|"picture"|"plus"|"square"|"star"|"triangle"|"x";
				size?:number;
				fillColor?:string;
				borderColor?:string;
			};
			dataLabels?:{
				showValue?:boolean;
				showCategory?:boolean;
				showSeriesName?:boolean;
				showPercent?:boolean;
				position?:"bestFit"|"center"|"insideBase"|"insideEnd"|"outsideEnd"|"left"|"right"|"top"|"bottom";
			};
		}[];
	}[];
	title?:ChartTextSource&{overlay?:boolean};
	legend?:{visible?:boolean;position?:"left"|"right"|"top"|"bottom"|"topRight";overlay?:boolean};
	axes?:{category?:ChartAxisSpec;value?:ChartAxisSpec;secondaryCategory?:ChartAxisSpec;secondaryValue?:ChartAxisSpec};
	displayBlanksAs?:"gap"|"span"|"zero";
	styleId?:number;
}
function listCharts(wb,options?:{ sheet?:string }):Promise<Array<{
	sheet:string;
	name:string;
	type:string;
	groupCount:number;
	seriesCount:number;
	position:ChartPos;
}>>;
function getChart(wb,sheet:string,name:string):Promise<Omit<ChartSpec,"position">&{ position:ChartPos }>;
function addChart(wb,sheet:string,chart:ChartSpec):Promise<ChartSpec>;
function setChart(wb,sheet:string,name:string,chart:ChartSpec):Promise<ChartSpec>;
function deleteChart(wb,sheet:string,name:string):Promise<void>;
interface NameDef {name:string;range:string;scope:string|null}
function listDefinedNames(wb):Promise<NameDef[]>;
function addDefinedName(wb,name:string,range:string,scope?:string):Promise<NameDef>;
function deleteDefinedName(wb,name:string,scope?:string):Promise<NameDef>;
function addSheet(wb,name:string):Promise<string>;
function deleteSheet(wb,name:string):Promise<void>;
function renameSheet(wb,oldName:string,newName:string):Promise<void>;
interface Value {
	address:string;
	sheet:string;
	row:number;
	col:number;
	colLetter:string;
	value:string|number|boolean|null;
	formula?:string;
	type:"string"|"number"|"bool"|"date"|"error"|"blank";
	text:string;
	format?:string;
	numberType?:"currency"|"percent"|"fraction"|"exponential"|"date"|"text"|"number";
	visibility:Visibility;
	context?:string;
	note?:{author:string;text:string};
	link?:{type:"internal"|"external";target:string;tooltip?:string};
	thread?:{
		resolved:boolean;
		comments:{authorId:string;text:string;createdAt:string}[];
	};
}
function readCell(wb,cell:CellRef,opts?:{
	context?:number;
}):Promise<Value>;
function readRange(wb,range:RangeRef):Promise<Value[][]>;
function readColumn(wb,sheetName:string,col:number|string,opts?:{
	startRow?:number;
	endRow?:number;
}):Promise<Value[]>;
function readRow(wb,sheetName:string,row:number,opts?:{
	startCol?:number;
	endCol?:number;
}):Promise<Value[]>;
function readRangeTsv(wb,range:RangeRef,opts?:{
	includeEmpty?:boolean;
	includeFormulas?:boolean;
}):Promise<string>;
function readColumnTsv(wb,sheetName:string,col:number|string,opts?:{
	startRow?:number;
	endRow?:number;
	includeEmpty?:boolean;
	includeFormulas?:boolean;
}):Promise<string>;
function readRowTsv(wb,sheetName:string,row:number,opts?:{
	startCol?:number;
	endCol?:number;
	includeEmpty?:boolean;
	includeFormulas?:boolean;
}):Promise<string>;
declare class SearchResults<T> extends Array<T> {truncated?:boolean}
type Matcher=string|string[]|number|boolean|RegExp|RegExp[]
/** Cell search by scalar, string list, or regex; `formulas:true` matches formulas only. */
function findCells(wb,matcher:Matcher,opts?:{in?:RangeRef|string;context?:number;limit?:number;offset?:number;formulas?:boolean}):Promise<{
	type:"cell";
	address:string;
	value:any;
	text:string;
	formula?:string;
	row:number;
	col:number;
	colLetter:string;
	sheet:string;
	visibility:Visibility;
	context?:string;
	role:string;
}[]>;
function findRows(wb,matcher:Matcher,opts?:{in?:RangeRef|string;context?:number;limit?:number;offset?:number}):Promise<{
	type:"row";
	row:number;
	sheet:string;
	matchedAt:string;
	range:string;
	tsv:string;
	visibility:Visibility;
	context?:string;
}[]>;
/** Replace text in values or formulas; use regex boundaries for formula edits. */
function findAndReplace(wb,find:string|RegExp,replace:string,opts?:{
		in?:RangeRef|string;
		matchCase?:boolean;
		wholeCell?:boolean;
		inFormulas?:boolean;
		limit?:number;
	}):Promise<{
	replaced:number;
	cells:string[];
	errors:Diag[];
}>;
function describeSheet(wb,sheetName:string):Promise<{
	tables:Record<string,{address:string;headerRows:string;headerCols:string|null;tableName?:string}>;
	structure:string; // Compact ASCII structure map
}>;
function tableLookup(wb,args:{
	table:string;
	rowLabel:string|number|boolean;
	columnLabel:string|number|boolean;
}):Promise<{
	address:string;
	value:any;
	text:string;
	row:number;
	col:number;
	colLetter:string;
	sheet:string;
	visibility:Visibility;
	rowLabelFoundAt:string;
	rowLabelFound:string;
	columnLabelFoundAt:string;
	columnLabelFound:string;
}[]>;
interface Diag {code:string;detail?:string;address:string;formula?:string}
interface WriteResult {
	touched:Record<string,string>;
	changed:string[];
	errors:Diag[];
	invalidatedTiles:{
		sheet:string;
		tileRow:number;
		tileCol:number;
	}[];
	updatedSheets:{
		name:string;
		usedRange:{
			startRow:number;
			startCol:number;
			endRow:number;
			endCol:number;
		}|null;
		tileRowCount:number;
		tileColCount:number;
	}[];
}
function setCells(wb,cells:Array<{
	address:CellRef;
	value?:unknown;
	formula?:string;
	format?:string;
	/** Legacy note payload; `null` clears. */
	note?:{text?:string;author?:string}|null;
	/** Hyperlink payload; `null` clears. */
	link?:{url?:string;ref?:string;tooltip?:string}|null;
	/** Thread payload: append via `add`, set `resolved`, or delete. */
	thread?:{add?:Array<{author?:string;text:string}>;resolved?:boolean;delete?:boolean}|null;
}>):Promise<WriteResult>;
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
}>;
function scaleRange(wb,range:RangeRef,factor:number,opts?:{skipFormulas?:boolean}):Promise<WriteResult|null>;
function insertRowAfter(wb,sheetName:string,row:number,count?:number):Promise<void>;
function deleteRows(wb,sheetName:string,row:number,count?:number):Promise<void>;
function insertColumnAfter(wb,sheetName:string,column:number|string,count?:number):Promise<void>;
function deleteColumns(wb,sheetName:string,column:number|string,count?:number):Promise<void>;
function autoFitColumns(wb,sheetName:string,columns?:Array<number|string>,opts?:{
	minWidth?:number;
	maxWidth?:number;
	padding?:number;
}):Promise<Record<string,{width:number;previousWidth:number}>>;
function sortRange(wb,range:RangeRef,keys:Array<{
	col:number|string;
	order?:"asc"|"desc";
}>,opts?:{hasHeader?:boolean}):Promise<void>;
function copyRange(wb,source:RangeRef,destination:CellRef,opts?:{
	pasteType?:"all"|"values"|"formulas"|"formats";
}):Promise<{destination:string;cellsCopied:number;}>;
type StyleObj={
	fill?:{
		color?:string;
		pattern?:string;
		patternColor?:string;
		gradient?:{
			type:string;
			degree?:number;
			color1:string;
			color2:string;
			top?:number;
			bottom?:number;
			left?:number;
			right?:number;
		};
	};
	font?:{
		name?:string;
		size?:number;
		color?:string;
		bold?:boolean;
		italic?:boolean;
		strike?:boolean;
		underline?:string;
		verticalAlign?:string;
	};
	alignment?:{
		horizontal?:string;
		vertical?:string;
		rotation?:number;
		wrapText?:boolean;
		shrinkToFit?:boolean;
		indent?:number;
	};
	border?:{
		top?:{
			style:string;
			color:string;
		};
		bottom?:{
			style:string;
			color:string;
		};
		left?:{
			style:string;
			color:string;
		};
		right?:{
			style:string;
			color:string;
		};
		diagonal?:{
			style:string;
			color?:string;
			up?:boolean;
			down?:boolean;
		};
	};
	numberFormat?:string;
	centerContinuousSpan?:number;
	richText?:{
		text:string;
		style?:{
			name?:string;
			size?:number;
			color?:string;
			bold?:boolean;
			italic?:boolean;
			strike?:boolean;
			underline?:string;
			verticalAlign?:string;
		};
	}[];
};
function getStyle(wb,cell:CellRef):Promise<StyleObj>;
function setStyle(wb,target:CellRef|RangeRef,style:StyleObj):Promise<void>;
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
		showValue?:boolean;
		gradient?:boolean;
		border?:boolean;
		negativeBarColorSameAsPositive?:boolean;
		negativeBarBorderColorSameAsPositive?:boolean;
		axisPosition?:"automatic"|"middle"|"none";
		direction?:"context"|"leftToRight"|"rightToLeft";
		fillColor?:string;
		borderColor?:string;
		negativeFillColor?:string;
		negativeBorderColor?:string;
		axisColor?:string;
		lowValue?:CfDataBarThreshold;
		highValue?:CfDataBarThreshold;
	};
	timePeriod?:"today"|"yesterday"|"tomorrow"|"last7Days"|"thisWeek"|"lastWeek"|"nextWeek"|"thisMonth"|"lastMonth"|"nextMonth";
}
function getConditionalFormatting(wb,sheetName:string):Promise<Array<CfRuleShared&{
	index?:number;
	type:CfWritableRuleType|"iconSet";
}>>;
/** Add conditional formatting rules; `opts.clear` replaces existing rules; `iconSet` is read-only. */
function setConditionalFormatting(wb,sheetName:string,rules:Array<CfRuleShared&{
	type:CfWritableRuleType;
}>,opts?:{clear?:boolean}):Promise<void>;
function removeConditionalFormatting(wb,sheetName:string,indices:number[]):Promise<void>;
function setSheetProperties(wb,sheetName:string,properties:{
	visibility?:"visible"|"hidden"|"veryHidden";
	view?:{
		showGridLines?:boolean;
		zoomScale?:number;
	};
	outline?:{
		summaryRowsBelow?:boolean;
		summaryColumnsRight?:boolean;
		showSymbols?:boolean;
	};
	format?:{
		defaultRowHeight?:number;
		defaultColWidth?:number;
		font?:{
			name?:string;
			size?:number;
		};
	};
	merges?:string[];
}):Promise<void>;
function setRowProperties(wb,sheetName:string,fromRow:number,toRow:number,properties:{
	height?:number;
	hidden?:boolean;
	outlineLevel?:number;
	collapsed?:boolean;
}):Promise<void>;
function setColumnProperties(wb,sheetName:string,fromCol:number|string,toCol:number|string,properties:{
	width?:number;
	hidden?:boolean;
	outlineLevel?:number;
	collapsed?:boolean;
}):Promise<void>;
function getSheetProperties(wb,sheetName:string,filter?:{
	columns?:(number|string)[];
	rows?:number[];
}):Promise<{
	visibility:"visible"|"hidden"|"veryHidden";
	view:{
		showGridLines:boolean;
		zoomScale:number;
	};
	outline:{
		summaryRowsBelow:boolean;
		summaryColumnsRight:boolean;
		showSymbols:boolean;
	};
	format:{
		defaultRowHeight:number;
		defaultColWidth:number;
		font?:{
			name?:string;
			size?:number;
		}|null;
	};
	columns:Record<
		string,
		{
			col:string;
			width:number;
			hidden?:boolean;
			outlineLevel?:number;
			collapsed?:boolean;
		}
	>;
	rows:Record<
		number,
		{
			row:number;
			height:number;
			hidden?:boolean;
			outlineLevel?:number;
			collapsed?:boolean;
		}
	>;
	merges?:string[]|null;
}>;
interface Deps {cells:{address:string;depth:number;formula?:string;referenceType?:"direct"|"range"|"named"|"table"}[];warnings?:Diag[]}
function getCellPrecedents(wb,address:CellRef,depth?:number):Promise<Deps>;
function getCellDependents(wb,address:CellRef,depth?:number):Promise<Deps>;
function traceToInputs(wb,cell:CellRef):Promise<{
	address:string;
	referenceCount:number;
	text?:string;
	nearbyLabel?:string;
	context?:string;
}[]>;
function traceToOutputs(wb,cell:CellRef):Promise<{
	address:string;
	formula?:string;
	text?:string;
	visibility:Visibility;
	nearbyLabel?:string;
	context?:string;
}[]>;
interface FormulaEval {formula:string;value:number|string|boolean|null|unknown[][];error?:{code:string;detail?:string}}
function evaluateFormulas(wb,sheet:string,formulas:string[]):Promise<FormulaEval[]>;
function evaluateFormula(wb,sheet:string,formula:string):Promise<FormulaEval>;
function lint(wb,options?:{
	rangeAddresses?:string[];
	skipRuleIds?:string[];
	onlyRuleIds?:string[];
}):Promise<{
	diagnostics:{
		severity:"Info"|"Warning"|"Error";
		ruleId:string;
		message:string;
		location:string|null;
		visibility:Visibility|null;
	}[];
	total:number;
}>;
/** Generates a PNG and prints its path to stdout. */
function previewStyles(wb,range:RangeRef):Promise<void>;
````
