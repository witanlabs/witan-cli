# Witan MCP xlsx API reference

The complete reference for the Witan MCP server's xlsx tools — the `xlsx_exec` scripting API (`xlsx.*`), the `xlsx_calc`, `xlsx_lint`, and `xlsx_render` tools, and the `prepare_*` file plumbing, with parameters, result shapes, and full type definitions. **Read this before your first `xlsx_exec` call** — the API surface is large and not guessable from names alone.

This file is the *reference*. For the *playbook* — when to reach for what, the reading / what-if and authoring workflows, the quality bar, and the verification gate — see `SKILL.md`.

`xlsx_exec` runs sandboxed JavaScript against a workbook opened server-side. Your script gets two globals: `xlsx` (this API) and `wb` (the workbook handle, passed first to every call). Top-level `await` is supported; `import` is not.

## Quick Reference

Each example is the `code` string of one `xlsx_exec` call; the leading comment shows the other arguments.

```js
// Create a new workbook from scratch (.xlsx only)
// xlsx_exec { filename: "model.xlsx", save: true }
await xlsx.addSheet(wb, "Inputs")
await xlsx.setCells(wb, [{ address: "Inputs!A1", value: "Revenue" }])
return await xlsx.listSheets(wb)
```

```js
// Read from sheets with spaces, apostrophes, or parentheses — inner apostrophes are doubled (Excel convention)
// xlsx_exec { file_id }
const a = await xlsx.readCell(wb, "'Workers'' Compensation'!B50")
const b = await xlsx.readRangeTsv(wb, { sheet: "Reserve Summary (Net)", from: {row:1,col:1}, to: {row:10,col:5} })
return { a: a.value, b }
```

```js
// Multi-input sweep — compare several input values at once
// xlsx_exec { file_id }
const result = await xlsx.sweepInputs(wb, {
	inputs: [
		{ address: "Inputs!B5", values: [0.02, 0.04, 0.06] },
		{ address: "Inputs!B6", values: [0.08, 0.10, 0.12] },
	],
	outputs: ["Output!C30", "Output!C45"],
	mode: "cartesian",
	includeStats: true,
})
print(result.tsv)
return result.stats
```

```js
// Conditional formatting — add a highlight rule and a color scale
// xlsx_exec { file_id, save: true }
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
```

```js
// Chart authoring — add an embedded chart and verify the rendered result
// xlsx_exec { file_id, save: true }
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
```

```js
// Waterfall chart authoring with totals and connector lines
// xlsx_exec { file_id, save: true }
await xlsx.addChart(wb, "Sheet1", {
	name: "Bridge",
	position: { from: { cell: "F2" }, to: { cell: "N18" } },
	groups: [
		{
			type: "waterfall",
			series: [
				{
					name: { text: "Movement" },
					categories: "Sheet1!A2:A8",
					values: "Sheet1!B2:B8",
					totalIndexes: [0, 6],
					showConnectorLines: true,
					dataLabels: { showValue: true, position: "outsideEnd" }
				}
			]
		}
	],
	title: { text: "Revenue Bridge" },
	styleId: 395
})
await xlsx.previewStyles(wb, "Sheet1!F2:N18")
```

```js
// Add an image — pass the picture as a data URL via the `input` argument
// xlsx_exec { file_id, save: true, input: { logo: "data:image/png;base64,…" } }
await xlsx.addImage(wb, "Sheet1", {
	name: "Logo",
	position: { from: { cell: "A1" }, to: { cell: "D6" } },
	source: { base64: input.logo }
})
```

```js
// ListObject authoring — create an Excel table and read it back by table name
// xlsx_exec { file_id }
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
```

```js
// What-If Data Table authoring — create a visible one-variable table block
// xlsx_exec { file_id }
await xlsx.addDataTable(wb, "Sheet1", {
	type: "oneVariableColumn",
	ref: "E1:F4",
	columnInputCell: "H1",
	inputValues: [5, 10, 15],
	formulas: ["=H1*2"]
})
return await xlsx.getDataTable(wb, "Sheet1!E1:F4")
```

## xlsx_exec — Workbook Scripting

Runs JavaScript against a workbook opened server-side; scripts interact through the `xlsx` and `wb` globals.

### Arguments

- `file_id` *or* `filename` — exactly one. `file_id` opens an uploaded working copy; `filename` starts from a brand-new empty workbook (`.xlsx` only — add a sheet before writing) and names the file if saved. It always mints a new workbook; it never looks up an existing file by name.
- `revision_id` — optional, with `file_id` only; defaults to the current revision.
- `code` — the script (required). JSON-encoded, so sheet names need no shell-style escaping.
- `input` — JSON value bound to the `input` global; defaults to `{}`. Use it for parameters and base64 `data:` URLs (images).
- `save` — persist; defaults to `false`. Editing a `file_id` saves a new revision (only if the script wrote cells); a `filename` run saves a new file (always, even if empty).
- `org_id` — usually omit; auto-resolves for single-org users (see Errors).
- `locale` — e.g. `'en-GB'`; server default when omitted.
- `timeout_ms`, `max_output_chars` — sandbox limits, capped server-side.

### Runtime globals

- `xlsx` (`object`): curated API surface; all functions listed below
- `wb` (`WorkbookContext`): opened workbook handle; pass as first arg to all `xlsx.*` calls
- `input` (`any`): parsed value of the `input` argument (defaults to `{}`)
- `print` (`function`): output to stdout, like `console.log` but captured in the response

Top-level `await` is supported. No imports allowed (static or dynamic).

### Sheet name quoting

Sheet names with spaces, apostrophes, or parentheses follow Excel's quoting convention in A1 references: wrap in single quotes and double any inner apostrophes — `"'Workers'' Compensation'!B50"`.

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
- `describeSheet`: sheet structure with detected tables
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
- `getDataValidations`: inspect data validation rules (finding cells that violate them is a `lint` rule)

**Rendering**

- `previewStyles`: generate a PNG preview for a sheet range; image is auto-registered into the response

**Charts**

- `listCharts`: chart summaries for the workbook or a single sheet
- `getChart`: canonical spec for an existing chart
- `addChart`, `setChart`, `deleteChart`: create, replace, or remove embedded charts
- Supported specs include combo charts, secondary axes, stock charts, bubble charts, radar charts, waterfall charts, surface charts, histogram/pareto charts, funnel charts, box & whisker charts, chart/plot-area formatting, group/series data labels, linked number formats, and style IDs. Use `previewStyles` after authoring to inspect rendered placement and labels.

**Conditional Formatting**

- `getConditionalFormatting`: read all sheet rules; `iconSet` is read-only
- `setConditionalFormatting`: add writable rules; `opts.clear` replaces all rules
- `removeConditionalFormatting`: remove rules by index

**Images**

- `listImages`, `getImage`: inspect embedded images (metadata only, no pixel data)
- `addImage`, `setImage`, `deleteImage`: place, replace, or remove embedded images; `source.base64` takes raw base64 or a `data:` URL (pass it in via `input`)

**Writing (ephemeral)**

- Cells/ranges: `setCells`, `findAndReplace`, `scaleRange`, `copyRange`, `sortRange`
- Data validation: `setDataValidations`, `removeDataValidations`; `setCells` supports `opts.validationMode: "reject"` to reject invalid writes
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
- `setDataValidations`: `opts.clear` replaces all existing rules on the sheet

### The ephemeral write contract

By default (`save: false`), `xlsx_exec` **does not persist anything**. All write operations (`setCells`, `scaleRange`, inserts, deletes) take effect in the server-side session only, recalculate there, and are discarded on close. The `setCells` result's `touched` map contains the recalculated formatted text values — read answers from there.

This means:

- No risk of corrupting the working copy — each invocation starts clean from the current revision
- In create mode (`filename`), no file is minted at all — prototype freely, nothing accumulates
- For independent what-if tests, each call starts from the original state

To persist, pass `save: true`: a `file_id` run gains a new `revision_id` (only if the script wrote cells); a `filename` run mints a new `file_id` (always — an empty workbook is still a successful create). Either way the response's `output` bundle has the ids and a presigned `download_url` — GET it to write the result back over the user's local file.

### setCells result shape

```ts
{
  touched: Record<string, string>  // address → formatted text value
  changed: string[]                // addresses whose values changed
  errors: Diag[]                   // cells that errored after recalc
}
```

Read the output value from `result.touched["Sheet!Address"]`. Never compute the answer in JavaScript. `setCells` implicitly creates a sheet if `address` references one that does not yet exist.

### Response shape

**Success:**

```json
{
  "ok": true,
  "stdout": "...",
  "result": "<json>",
  "writes_detected": false,
  "accesses": [{ "operation": "read", "address": "Sheet1!A1" }],
  "file_id": "...", "revision_id": "...",
  "output": { "file_id": "...", "revision_id": "...", "filename": "...", "download_url": "...", "download_expires_at": "..." }
}
```

`output` appears only when `save: true` persisted something. Top-level `file_id`/`revision_id` are present on every `file_id` run, but on `filename` runs only once saved. `truncated: true` flags stdout cut at `max_output_chars`.

**Failure:**

```json
{
  "ok": false,
  "stdout": "...",
  "error": { "type": "syntax|runtime|timeout", "code": "...", "message": "..." }
}
```

### What-if / sensitivity workflow

For questions like "what happens to Y if X changes?", follow this sequence. **Steps 1 and 2 should be separate `xlsx_exec` calls** — review the output of step 1 before proceeding.

1. **Find the output cell (separate call)** — search for what the user is asking about (e.g., "net income", "reserve"). Use `xlsx.findCells` with `context: 2` and review the candidates. Labels and formula cells often share the same text — pick the formula cell, not the label. This step may take more than one attempt: try synonym arrays, read surrounding rows with `xlsx.readRangeTsv`, or check multiple sheets.
2. **Trace + run the what-if (second call)** — once you have the output address, call `xlsx.traceToInputs(wb, outputAddr)` to confirm the user's input actually drives it. Trace results can contain hundreds of cells — filter by `nearbyLabel` matching the user's term instead of printing them all. Then `xlsx.setCells` to make the change and read the answer from `result.touched[outputAddr]`. If the output address is missing from `touched`, the cell didn't recalculate — you likely have the wrong address.
3. **Report before and after** — always include the baseline value (read before the edit) and the new value from `touched`.

For sweeping multiple values (sensitivity tables), use `sweepInputs` instead — it runs all combinations in one call and returns structured before/after data.

Practical tips:

- Don't search for the output _after_ `setCells` — find it first so you know the exact address to check in `touched`.
- If `findCells` returns several hits for the same label, use context or `readRangeTsv` to disambiguate (e.g., "Net Income" may appear as both a label and a formula cell on different rows).
- After `setCells`, verify `result.errors` is empty. New errors mean the edit introduced or surfaced a calculation problem downstream.
- If the question names a specific metric (e.g., "loss ratio"), don't just search for that term — also check the sheet/row the question references, since models often have multiple versions of the same metric across sheets.

### Iterative / circular models

When a workbook has **iterative calculation** enabled, `setCells` recalculates the edited cells and downstream dependents using the workbook's iterative settings. If the model does not converge within `maxIterations`, the result includes convergence errors.

Use `xlsx_calc` when you want a standalone seeded or full-workbook verification pass with a report of all calculation errors or changed cells.

## Files in and out — prepare_upload, prepare_upload_revision, prepare_download

All three are two-step: the tool mints a short-lived presigned URL, you move the bytes yourself. Re-mint if a URL expires.

- **`prepare_upload { filename }`** → `{ upload_url, expires_at }`. POST the raw file bytes (not multipart) to `upload_url`; the JSON response has `id` (the new `file_id`) and `revision_id`. Accepts `.xlsx` and `.xls` (converted server-side); first upload only.
- **`prepare_upload_revision { file_id }`** → `{ upload_url, expires_at }`. POST updated bytes for a file you already uploaded — the response has the new `revision_id`; the filename is preserved. Use it to re-sync after the user edits their local file.
- **`prepare_download { file_id, revision_id? }`** → `{ download_url, filename, expires_at }`. GET to download (the response carries `Content-Disposition: attachment`). Usually unnecessary right after a save — mutating tools already return a `download_url` in their `output` bundle.

## xlsx_calc — Full-workbook verification

`setCells` already recalculates the edited cells and all downstream dependents, so you do **not** need `xlsx_calc` after every normal edit. Use it as a standalone verification/reporting pass:

- `verify: true` — **dry-run**: recalculates and reports without persisting. This is the verification-gate mode.
- Default (no `verify`) — recalculates **and saves a new revision**, returned in `output`. Only do this deliberately (e.g. to refresh a stale workbook).
- `addresses` — optional seed cells to recalc from (dependents propagate); omit for the whole workbook.

Response: `errors` (every formula error, with `address`, `code`, `formula`, `detail`), `changed` (addresses whose values changed), and `touched_count`. The gate is `errors` empty — fix and repeat until it is.

## xlsx_lint — Semantic formula checks

Reports logic problems that compute without error — double-counting from overlapping ranges, approximate-match lookups on unsorted data, mixed currencies or percent/non-percent in one expression, empty-ref coercion, references to non-anchor cells in a merged range, data-validation violations. Distinct from `xlsx_calc`, which catches hard formula errors (`#REF!`, `#DIV/0!`).

- `ranges` — optional A1 ranges to restrict the lint (e.g. `["Sheet1!A1:Z50"]`); omit for every sheet.
- `only_rules` / `skip_rules` — rule ids to include or exclude (mutually exclusive).

Response: `total` plus `diagnostics` (`severity` `Info`/`Warning`/`Error`, `ruleId`, `message`, `location`). Resolve or knowingly accept every `Warning`/`Error`. The same checks are available in-script via `xlsx.lint(wb, ...)`.

## xlsx_render — Visual screenshot

Renders a sheet range as an image, returned inline in the tool result — useful for inspecting layout, merged cells, formatting, charts, and labels.

- `range` — sheet-qualified A1 range, required (e.g. `"Sheet1!A1:Z50"`; quote sheet names as usual).
- `dpr` — device pixel ratio `1`–`3`; defaults to `2`.
- `format` — `png` (default) or `webp`.

The `previewStyles` exec function (see Rendering above) provides the same capability from within a script.

## Error Guide

- `EXEC_SYNTAX_ERROR`: fix JavaScript syntax in your script
- `EXEC_RUNTIME_ERROR`: fix the runtime error; check the message for details
- `EXEC_RESULT_TOO_LARGE`: return less data; use `print(...)` for large output instead of return values
- `Import statements are not allowed`: no `import` in exec scripts; use the `xlsx` global
- `Sheet 'X' not found`: check the sheet name; use `listSheets` to enumerate
- `ADDRESS_PARSE_ERROR` / `RANGE_PARSE_ERROR`: bad A1 reference; check sheet quoting
- org resolution failed with a candidate list: the user belongs to several orgs — ask which, then pass `org_id` on every call
- presigned URL rejected/expired: mint a fresh one with the `prepare_*` tool
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
type ChartDataLabelPosition="bestFit"|"center"|"insideBase"|"insideEnd"|"outsideEnd"|"left"|"right"|"top"|"bottom"
/** Use exactly one of `text` or `ref`. */
interface ChartTextSource {text?:string;ref?:string}
interface ChartPositionAnchor {cell:string;xOffsetPts?:number;yOffsetPts?:number}
interface ChartPosIn {from:ChartPositionAnchor;to:ChartPositionAnchor}
interface ChartPos extends ChartPosIn {sheet:string}
interface ChartFillFormat {noFill?:boolean;color?:string}
interface ChartLineFormat {noLine?:boolean;color?:string;weight?:number;lineStyle?:string}
interface ChartFontFormat {bold?:boolean;color?:string;italic?:boolean;name?:string;size?:number;underline?:string}
interface ChartDataLabelFormat {fill?:ChartFillFormat;border?:ChartLineFormat;font?:ChartFontFormat}
interface ChartPlotAreaSpec {format?:{fill?:ChartFillFormat;border?:ChartLineFormat}}
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
	numberFormatLinked?:boolean;
	reversed?:boolean;
	majorGridlines?:boolean;
	minorGridlines?:boolean;
	position?:"left"|"right"|"top"|"bottom";
}
interface ChartSpec {
	name:string;
	position:ChartPosIn;
	groups:{
		type:"column"|"bar"|"line"|"area"|"pie"|"doughnut"|"scatter"|"bubble"|"radar"|"surface"|"stockHLC"|"stockOHLC"|"waterfall"|"histogram"|"pareto"|"funnel"|"boxWhisker";
		scatterStyle?:"line"|"lineMarker"|"marker"|"smooth"|"smoothMarker"; /** scatter only */
		radarStyle?:"standard"|"marker"|"filled"; /** radar only */
		surfaceVariant?:"topView"|"topViewWireframe"; /** surface only */
		grouping?:"standard"|"stacked"|"percentStacked";
		axis?:"primary"|"secondary";
		gapWidth?:number;
		overlap?:number;
		varyColors?:boolean;
		smooth?:boolean;
		firstSliceAngle?:number;
		holeSize?:number;
		bubbleScale?:number; /** bubble only, 0-300 */
		showNegativeBubbles?:boolean; /** bubble only */
		sizeRepresents?:"area"|"width"; /** bubble only */
		dataLabels?:{
			showLegendKey?:boolean;
			showValue?:boolean;
			showCategory?:boolean;
			showSeriesName?:boolean;
			showPercent?:boolean;
			showBubbleSize?:boolean;
			showLeaderLines?:boolean;
			position?:ChartDataLabelPosition;
			numberFormat?:string;
			numberFormatLinked?:boolean;
			separator?:string;
			format?:ChartDataLabelFormat;
		};
		series:{
			name?:ChartTextSource;
			stockRole?:"volume"|"open"|"high"|"low"|"close"; /** stock charts only */
			categories?:string;
			categoriesRefType?:"string"|"number"|"multiLevelString";
			values?:string;
			xValues?:string;
			yValues?:string;
			bubbleSizes?:string; /** bubble only */
			fillColor?:string;
			lineColor?:string;
			lineWidth?:number;
			lineDashStyle?:string;
			smooth?:boolean;
			invertIfNegative?:boolean;
			totalIndexes?:number[]; /** waterfall only: zero-based subtotal/total point indexes */
			showConnectorLines?:boolean; /** waterfall only */
			binOptions?:{type?:"auto"|"binCount"|"binWidth"|"category";count?:number;width?:number;allowOverflow?:boolean;overflowValue?:number;allowUnderflow?:boolean;underflowValue?:number}; /** histogram/pareto only */
			quartileCalculation?:"exclusive"|"inclusive"; /** boxWhisker only */
			showInnerPoints?:boolean; /** boxWhisker only */
			showMeanLine?:boolean; /** boxWhisker only */
			showMeanMarker?:boolean; /** boxWhisker only */
			showOutlierPoints?:boolean; /** boxWhisker only */
			marker?:{
				style?:"auto"|"none"|"circle"|"dash"|"diamond"|"dot"|"picture"|"plus"|"square"|"star"|"triangle"|"x";
				size?:number;
				fillColor?:string;
				borderColor?:string;
			};
			dataLabels?:{
				showLegendKey?:boolean;
				showValue?:boolean;
				showCategory?:boolean;
				showSeriesName?:boolean;
				showPercent?:boolean;
				showBubbleSize?:boolean;
				showLeaderLines?:boolean;
				position?:ChartDataLabelPosition; /** for bubble charts only center/left/right/top/bottom */
				numberFormat?:string;
				numberFormatLinked?:boolean;
				separator?:string;
				format?:ChartDataLabelFormat;
			};
		}[];
	}[];
	title?:ChartTextSource&{overlay?:boolean};
	legend?:{visible?:boolean;position?:"left"|"right"|"top"|"bottom"|"topRight";overlay?:boolean};
	axes?:{category?:ChartAxisSpec;value?:ChartAxisSpec;secondaryCategory?:ChartAxisSpec;secondaryValue?:ChartAxisSpec};
	format?:ChartDataLabelFormat; /** chart-area fill/border/font */
	plotArea?:ChartPlotAreaSpec;
	displayBlanksAs?:"gap"|"span"|"zero";
	plotVisibleOnly?:boolean; /** rejected for waterfall charts */
	showDataLabelsOverMaximum?:boolean;
	roundedCorners?:boolean;
	styleId?:number; /** legacy styles 1-48, or modern catalog styles eg. 201,227,240,251,269,276. */
}
function listCharts(wb,options?:{ sheet?:string }):Promise<Array<{
	id?:number;
	sheet:string;
	name:string;
	type:string;
	groups:{type:string;axis?:string;seriesCount:number}[];
	groupCount:number;
	seriesCount:number;
	position:ChartPos;
}>>;
function getChart(wb,sheet:string,name:string):Promise<Omit<ChartSpec,"position">&{ position:ChartPos }>;
function previewChart(wb,sheet:string,chart:string|number,options?:{format?:"png"|"webp";dpr?:number;zoom?:number}):Promise<void>;
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
}>,opts?:{validationMode?:"ignore"|"reject"}):Promise<WriteResult>;
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
type DvOperator="Between"|"NotBetween"|"EqualTo"|"NotEqualTo"|"GreaterThan"|"LessThan"|"GreaterThanOrEqualTo"|"LessThanOrEqualTo";
type DvBasic={operator:DvOperator;formula1:string|number;formula2?:string|number|null};
type DvRule={
	wholeNumber?:DvBasic|null;
	decimal?:DvBasic|null;
	list?:{source:string;inCellDropDown?:boolean}|null;
	date?:DvBasic|null;
	time?:DvBasic|null;
	textLength?:DvBasic|null;
	custom?:{formula:string}|null;
};
type DvSpec={
	address:string;
	rule:DvRule;
	ignoreBlanks?:boolean;
	prompt?:{showPrompt?:boolean;title?:string|null;message?:string|null}|null;
	errorAlert?:{showAlert?:boolean;style?:"Stop"|"Warning"|"Information";title?:string|null;message?:string|null}|null;
};
function getDataValidations(wb,opts?:{sheet?:string;address?:string}):Promise<Array<DvSpec&{
	index:number;
	sheet:string;
	type:"None"|"WholeNumber"|"Decimal"|"List"|"Date"|"Time"|"TextLength"|"Custom";
}>>;
function setDataValidations(wb,sheetName:string,rules:DvSpec[],opts?:{clear?:boolean}):Promise<void>;
function removeDataValidations(wb,sheetName:string,target:{indices:number[]}|{address:string}):Promise<void>;
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
/** Renders a PNG of the range; the image is auto-registered into the response for inspection. */
function previewStyles(wb,range:RangeRef):Promise<void>;
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
/** Image reads/lists return metadata only — never pixel data. */
function listImages(wb,options?:{sheet?:string}):Promise<ImageInfo[]>;
function getImage(wb,sheet:string,selector:ImageSelector):Promise<ImageInfo>;
function addImage(wb,sheet:string,image:ImageSpec):Promise<ImageInfo>;
function setImage(wb,sheet:string,selector:ImageSelector,image:ImageUpdate):Promise<ImageInfo>;
function deleteImage(wb,sheet:string,selector:ImageSelector):Promise<void>;
````
