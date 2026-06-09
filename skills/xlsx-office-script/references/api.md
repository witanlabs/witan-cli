# Witan xlsx API reference — Office Scripts (ExcelScript)

The complete `witan xlsx` reference for the **Office Scripts (ExcelScript)** dialect — the `exec` scripting API plus the `calc`, `lint`, and `render` commands, with CLI flags, the script entry contract, a task → API map, and worked recipes. **Read this before your first `witan xlsx exec` call** — it covers the `// @office-script` harness, the engine's implemented subset, and the traps the ExcelScript docs won't warn you about.

This file is the *reference*. For the *playbook* — when to reach for what, the reading / what-if / authoring workflows, the quality bar, and the verification gate — see `SKILL.md`. The full ExcelScript type surface is `references/excelscript.d.ts` (grep it).

`witan xlsx exec` runs a sandboxed Office Script server-side. Your script is **`function main(workbook)`** behind a **`// @office-script`** first-line pragma. ExcelScript is **synchronous** — no `await`, no `context.sync()`. No `import`.

## Setup

The CLI supports `.xls`, `.xlsx`, and `.xlsm`; legacy `.xls` files are converted to `.xlsx` when needed. New workbook creation is `.xlsx` only.

## Quick Reference

```bash
# Create a new workbook from scratch (.xlsx only)
witan xlsx exec model.xlsx --create --save --stdin <<'JS'
// @office-script
function main(workbook) {
  const sheet = workbook.addWorksheet("Inputs");
  sheet.getRange("A1").setValue("Revenue");
  return workbook.getWorksheets().map(s => s.getName());   // ["Inputs"]
}
JS

# Read from sheets with spaces, apostrophes, or parentheses
witan xlsx exec model.xlsx --stdin <<'JS'
// @office-script
function main(workbook) {
  // getWorksheet takes the RAW name (single apostrophe, no quoting)
  const a = workbook.getWorksheet("Workers' Compensation").getRange("B50").getValue();
  // Only when a sheet name is embedded in an A1 STRING is the inner apostrophe doubled (Excel convention)
  const b = workbook.getWorksheet("Reserve Summary (Net)").getUsedRange().getValues();
  return { a, b };
}
JS

# Multi-input sweep — no sweepInputs helper, loop in one main(); recalc is automatic
witan xlsx exec model.xlsx --stdin <<'JS'
// @office-script
function main(workbook) {
  const inp = workbook.getWorksheet("Inputs"), out = workbook.getWorksheet("Output");
  const rows = [];
  for (const b5 of [0.02, 0.04, 0.06]) {
    for (const b6 of [0.08, 0.10, 0.12]) {
      inp.getRange("B5").setValue(b5);
      inp.getRange("B6").setValue(b6);
      rows.push({ b5, b6, c30: out.getRange("C30").getValue(), c45: out.getRange("C45").getValue() });
    }
  }
  return rows;   // cartesian = nested loops; parallel = one zipped loop
}
JS

# Conditional formatting — a highlight rule and a colour scale
witan xlsx exec model.xlsx --save --stdin <<'JS'
// @office-script
function main(workbook) {
  const sheet = workbook.getWorksheet("Sheet1");
  // Highlight A1:A100 where value > 100 → red fill
  const cf1 = sheet.getRange("A1:A100").addConditionalFormat(ExcelScript.ConditionalFormatType.cellValue);
  cf1.getCellValue().setRule({ formula1: "100", operator: ExcelScript.ConditionalCellValueOperator.greaterThan });
  cf1.getCellValue().getFormat().getFill().setColor("#FF0000");
  // Two-colour scale B1:B100 white(min) → red(max)
  const cf2 = sheet.getRange("B1:B100").addConditionalFormat(ExcelScript.ConditionalFormatType.colorScale);
  cf2.getColorScale().setCriteria({
    minimum: { type: ExcelScript.ConditionalFormatColorCriterionType.lowestValue, color: "#FFFFFF" },
    maximum: { type: ExcelScript.ConditionalFormatColorCriterionType.highestValue, color: "#FF0000" },
  });
}
JS

# Chart authoring — embedded column chart, then verify with `witan xlsx render`
witan xlsx exec model.xlsx --save --stdin <<'JS'
// @office-script
function main(workbook) {
  const sheet = workbook.getWorksheet("Sheet1");
  const chart = sheet.addChart(
    ExcelScript.ChartType.columnClustered,   // type
    sheet.getRange("A1:B9"),                  // source range
    ExcelScript.ChartSeriesBy.columns,        // columns | rows | auto — infers categories/series
  );
  chart.setName("Revenue");
  chart.getTitle().setText("Revenue");
  chart.getLegend().setPosition(ExcelScript.ChartLegendPosition.right);
  chart.setPosition(sheet.getRange("F2"), sheet.getRange("N18"));   // top-left + bottom-right cells (Ranges)
}
JS

# Waterfall chart with connector lines  (totals: NOT supported — see Not supported below)
witan xlsx exec model.xlsx --save --stdin <<'JS'
// @office-script
function main(workbook) {
  const sheet = workbook.getWorksheet("Sheet1");
  const chart = sheet.addChart(ExcelScript.ChartType.waterfall, sheet.getRange("A1:B8"), ExcelScript.ChartSeriesBy.columns);
  chart.setName("Bridge");
  chart.getTitle().setText("Revenue Bridge");
  chart.getSeries()[0].setShowConnectorLines(true);
  chart.setPosition(sheet.getRange("F2"), sheet.getRange("N18"));
}
JS

# Add an image from a local file (--input-file loads it as a data URL → input.logo)
witan xlsx exec model.xlsx --save --input-file logo=@./logo.png --stdin <<'JS'
// @office-script
function main(workbook, input) {
  const sheet = workbook.getWorksheet("Sheet1");
  const raw = String(input.logo).replace(/^data:[^,]+,/, "");   // addImage wants RAW base64, not the data: URL
  const shape = sheet.addImage(raw);                            // returns a Shape; takes ONLY a base64 string
  shape.setLeft(10); shape.setTop(10); shape.setWidth(120); shape.setHeight(80);   // POINTS, not cells
}
JS

# ListObject (Excel table) — create and read back by column name  (rename/totals: NOT supported — see below)
witan xlsx exec model.xlsx --stdin <<'JS'
// @office-script
function main(workbook) {
  const sheet = workbook.getWorksheet("Sheet1");
  sheet.getRange("A1:C1").setValues([["Region", "Sales", "DoubleSales"]]);
  sheet.getRange("A2:C3").setValues([["North", 10, null], ["South", 20, null]]);
  const table = sheet.addTable(sheet.getRange("A1:C3"), true);   // hasHeaders=true; auto-named "Table1"
  // Calculated column: per-cell A1 formulas. Structured refs (=[@Sales]*2) evaluate to #REF! — DON'T use them.
  table.getColumnByName("DoubleSales").getRangeBetweenHeaderAndTotal().setFormulas([["=B2*2"], ["=B3*2"]]);
  return {
    name: table.getName(),                                        // "Table1" (cannot be renamed)
    body: table.getRangeBetweenHeaderAndTotal().getValues(),      // [["North",10,20],["South",20,40]]
  };
}
JS

# Simple one-liner — use --code (NOT --expr, which rejects function main)
witan xlsx exec model.xlsx --create --code '// @office-script
function main(workbook){ workbook.addWorksheet("S"); return workbook.getWorksheets().map(s=>s.getName()); }'
```

### Not supported in Office Script (and the substitute)

- **What-If Data Tables** — no ExcelScript creation API exists. Writing `=TABLE(,H1)` does **not** throw but yields `#NAME?` (silent wrong result). Substitute: precompute with the sweep loop above and write a static labelled block of ordinary formulas — it won't live-recalc as a `TABLE()` object would.
- **Waterfall total/subtotal columns** — `setShowConnectorLines` works, but there is **no member to mark a point as a total** (no `setAsTotal`/`isTotal` on `ChartSeries`/`ChartPoint`). Start/End columns float instead of anchoring to the baseline.
- **Table rename & totals row** — `Table.setName` and `Table.setShowTotals` are **unimplemented** (`NotImplementedError`, see Error Guide). Tables keep their auto name (`Table1`, …); write a totals row as ordinary formulas in a cell *below* the table (plain `=SUM(B2:B3)`, never `=SUBTOTAL(109,Table1[Sales])` — structured refs break).
- **Structured table references** — `=[@Col]` / `Table1[Col]` evaluate to `#REF!` in the calc engine. Use plain A1 references everywhere.
- **Column widths** — `RangeFormat.setColumnWidth` is **unimplemented** (`NotImplementedError`). Substitute: `range.getFormat().autofitColumns()` (sizes to content). Row heights are fine — `setRowHeight` works.

## exec — Workbook Scripting

Runs an Office Script against a workbook opened server-side. If the workbook doesn't exist yet, use `--create` with a new `.xlsx` path.

- `--create --save` — produce a real workbook file on disk.
- `--create` without `--save` — prototype structure, test generation logic, inspect returned data, or validate formulas/layout without leaving a file behind.

### Invocation patterns

**Recommended: `--stdin` with a single-quoted heredoc** — safe for all sheet names, multi-line, and batches operations into one CLI call:

```bash
witan xlsx exec report.xlsx --stdin <<'JS'
// @office-script
function main(workbook) {
  const sheets = workbook.getWorksheets().map(s => s.getName());
  const a1 = workbook.getWorksheet("My Sheet").getRange("A1").getValue();
  return { sheets, a1 };
}
JS
```

The single-quoted delimiter (`<<'JS'`) prevents all shell expansion — apostrophes, parentheses, quotes, and globs in sheet names pass through verbatim.

**`--code`** takes the same script as an inline string (the pragma must still be the first line). **`--script file.ts`** runs a reusable file. **`--expr` does NOT work** — it is single-expression only and rejects `function main`; use `--code` or `--stdin` for the one-liner case.

Provide exactly one code source: `--stdin`, `--code`, or `--script`.

### Flags

- Code source: exactly one of `--stdin`, `--code`, or `--script` (not `--expr`).
- Inputs/control: `--input-json`, `--input-file key=@path`, `--timeout-ms`, `--max-output-chars`, `--stdin-timeout-ms`, `--locale`.
- Workbook lifecycle: `--create` creates a new `.xlsx` session; `--save` persists changes; `--json` prints the full response envelope.

### Script entry contract

- **`// @office-script` MUST be the first line.** It selects the ExcelScript dialect. Omit it and the script does **not** error — it runs in the other dialect, never calls `main`, and returns `null`. Silent wrong result. This is the #1 trap.
- **`function main(workbook)` is the entry point** — required. Return a JSON-serializable value; that becomes the CLI result.
- **`--input-json` / `--input-file` arrive as a second argument:** `function main(workbook, input)`. `--input-file logo=@./logo.png` sets `input.logo` to a `data:<mime>;base64,...` URL — strip the `data:` prefix before `addImage` (raw base64 only).
- **Synchronous** — no top-level `await`, no `context.sync()`. There are no `xlsx`/`wb` globals.

### The API surface

Everything hangs off the `workbook` argument; you build/get objects and mutate them. Full signatures: `references/excelscript.d.ts` (`rg -n "interface Worksheet|addChart|ConditionalFormatType|addTable|addImage" references/excelscript.d.ts`). Common tasks:

- **Sheets** — `workbook.getWorksheets()`, `getWorksheet(name)` (raw name, returns `undefined` if absent), `addWorksheet(name)`, `getActiveWorksheet()`. A `--create` workbook has **zero** sheets — `addWorksheet` first.
- **Read** — `sheet.getRange("A1:H20")` → `.getValues()` / `.getFormulas()` / `.getValue()` / `.getFormula()` / `.getText()`; `sheet.getUsedRange(valuesOnly?)` for the populated region. No `readRangeTsv` — build TSV in JS from `getValues()` if needed.
- **Find** — `range.find(text, { completeMatch, matchCase })` returns the cell or a **nullish value** if absent (no `*OrNullObject`). Exact/substring only — **not** fuzzy/synonym. There is **no** dependency tracing (`traceToInputs`/precedents), fuzzy `findCells`, or `tableLookup` — read formulas and walk them by hand.
- **Write** — `range.setValue(v)` / `setValues(2D)` / `setFormula(s)` / `setFormulas(2D)`. **`setFormula` on a multi-cell range writes the literal formula to every cell — no relative fill.** To fill a relative formula down, build per-cell strings: `setFormulas([["=A1*2"],["=A2*2"]])`.
- **Format** — `range.setNumberFormat(str)`, `range.merge(across?)`, and `range.getFormat()` → `getFont().setColor(hex)/setBold(bool)`, `getFill().setColor(hex)`, `setHorizontalAlignment("Left"|"Right"|"Center")`, `getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.continuous)/.setWeight(ExcelScript.BorderWeight.thin)/.setColor(hex)`.
- **Comments** — `workbook.addComment(address, "text")` (for sourcing notes).
- **Charts / Conditional formatting / Images / Tables** — imperative: create or `add*` an object, then configure the returned object (see Quick Reference). `ExcelScript.ChartType`: `columnClustered`, `barClustered`, `line`, `pie`, `xyscatter`, `area`, `doughnut`, `radar`, `histogram`, `pareto`, `boxwhisker`, `waterfall`, `funnel`, `treemap`, `sunburst`, `regionMap`.

### The ephemeral write contract

By default, `exec` **does not write workbook bytes back to disk** — all writes take effect in the server-side session only. So: no risk to the original file, no reset needed (each invocation starts clean), and each what-if `exec` starts from the original. Pass `--save` to persist. With `--create`, the same rule applies to the new session — no local file unless `--save` is also passed (`.xlsx` only).

### Reading results

There is **no `result.touched` map** (that is the other dialect). Read answers **in-script**: after `setValue`/`setFormula`, the engine recalculates synchronously, so `range.getValue()` returns the new value immediately. Return it from `main`. Never recompute the answer in JS.

### Response format

With `--json`, the full envelope is returned:

```json
{ "ok": true, "stdout": "...", "result": "<json>", "writes_detected": false, "accesses": [...] }
{ "ok": false, "stdout": "...", "error": { "type": "...", "code": "...", "message": "..." } }
```

`accesses` documents all cell reads/writes with operation type and address.

### What-if / sensitivity workflow

For "what happens to Y if X changes?", use **two separate `exec` calls** — review step 1 before editing.

1. **Find the output cell (separate call)** — search for the metric, read its `getFormula()`, confirm it's the formula cell (not the label). No fuzzy search — try the literal label, or `getValues()` and scan in JS.
2. **Change + read (second call)** — read the baseline `getValue()`, `setValue` on input X, then read Y's `getValue()` again (recalc is automatic). If Y didn't move, X doesn't drive it ⇒ wrong cell. There is **no `traceToInputs`** to confirm the link first and **no `sweepInputs`** — for many combinations, loop in one `main` (Quick Reference).
3. **Report baseline → new**, and re-read a few downstream cells or run `witan xlsx calc` to confirm nothing else broke.

### Iterative / circular models

`witan xlsx calc` runs a standalone full-workbook pass; if an iterative model doesn't converge, convergence errors show up there. Set iterative calculation through the workbook's application settings (see `excelscript.d.ts`) before relying on circular references.

## calc — Full-workbook verification

Dialect-independent — operates on the saved file. Use as a standalone verification/reporting command:

```bash
witan xlsx calc model.xlsx              # recalc; print all formula errors
witan xlsx calc model.xlsx --verify     # no write; also prints changed addresses
witan xlsx calc model.xlsx --show-touched
```

- No errors → one-line summary like `428 cells recalculated, 0 errors, 3 changed`.
- Any formula errors → prints **all** of them (address, formula, code) and exits `2`.
- `--verify` also exits `2` if any computed value changed — useful as a final audit after a fix.

```text
2 errors:
  Summary!C18          =A18/B18                      #DIV/0!
  Revenue!F42          =VLOOKUP(A42,$A$2:$C$10,3,0)  #N/A
```

## lint — Semantic formula checks

`witan xlsx lint` reports logic problems that compute without error — double-counting from overlapping ranges, approximate-match lookups on unsorted data, mixed currencies or percent/non-percent in one expression, empty-ref coercion, references to non-anchor cells in a merged range. Distinct from `calc`, which catches hard errors (`#REF!`, `#DIV/0!`).

```bash
witan xlsx lint model.xlsx                    # whole workbook
witan xlsx lint model.xlsx -r "Sheet1!A1:Z50" # limit to a range
witan xlsx lint model.xlsx --only-rule D001   # or --skip-rule D001
```

Exits `2` when any Error or Warning is reported; `--json` gives structured diagnostics.

## render — Visual Screenshot

Renders a sheet range as a PNG — inspect layout, merges, formatting, charts, and labels. Dialect-independent.

```bash
witan xlsx render report.xlsx -r "Sheet1!A1:Z50"              # → temp PNG path on stdout
witan xlsx render report.xlsx -r "'My Sheet'!B5:H20" -o out.png
witan xlsx render report.xlsx -r "Sheet1!A1:F10" --dpr 2      # DPR 1-3
witan xlsx render report.xlsx -r "Sheet1!A1:F10" --diff before.png
```

`-r/--range` (required, sheet-qualified), `-o/--output` (defaults to temp), `--dpr` (default auto), `--format` (`png`|`webp`), `--diff` (compare against a baseline). Always render after chart/image authoring to confirm the visual.

## Error Guide

- **Script returns `null` unexpectedly** — the `// @office-script` pragma is missing or not the first line; the script ran in the other dialect and never called `main`. Put the pragma on line 1.
- `EXEC_RUNTIME_ERROR: NotImplementedError: ExcelScript: <Member> is not implemented by xlsx-serve` — that member isn't wired up in the engine (e.g. `Table.setName`, `Table.setShowTotals`). **It aborts the run and is NOT catchable by `try/catch`** — pick a supported path. Cross-check `witan-alfred-2/XlsxServeCli/Exec/OfficeScript/` for what's implemented.
- `EXEC_RUNTIME_ERROR: Workbook has no worksheets` — a `--create` workbook starts empty; `addWorksheet(name)` before `getActiveWorksheet()`.
- `--expr is for single expressions; use --code` — Office Script needs `function main`; use `--code` or `--stdin`.
- `#REF!` in a calculated table column — structured refs (`=[@Col]`) aren't supported; use A1 references.
- `#NAME?` from a `=TABLE(...)` formula — What-If Data Tables have no ExcelScript path; precompute instead.
- `INVALID_ARG: image source.base64 is not a valid png image` — pass a real PNG (and strip the `data:` URL prefix to raw base64).
- `EXEC_SYNTAX_ERROR` — fix JavaScript/TypeScript syntax. `EXEC_RESULT_TOO_LARGE` — return less; use `console.log` for bulk output.
- `Sheet 'X' not found` / `getWorksheet` returns `undefined` — check the name with `workbook.getWorksheets()`.

## Type definitions

The complete ExcelScript surface is vendored at **`references/excelscript.d.ts`** (15k lines — the full Microsoft Office Scripts API). The engine implements it *where wired up*; an unimplemented member throws `NotImplementedError` (above). Grep the `.d.ts` for signatures; prototype ephemerally and let `calc`/`render` confirm.
