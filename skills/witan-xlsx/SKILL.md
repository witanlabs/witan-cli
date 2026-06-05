---
name: witan-xlsx
description: Read, explore, analyse, author, and edit Excel workbooks (.xls, .xlsx, .xlsm) with Witan's sandboxed JavaScript engine (`witan xlsx exec`), plus lint, recalculate, and render. Excel files are binary — cat/head and ordinary file tools cannot read them; this is the way to inspect or change them. Use when the user references a spreadsheet by name or path (even casually, e.g. "check report.xlsx", "what's in the model"), asks to read cells/rows/ranges, search for values, trace how a figure is calculated, run what-if or sensitivity scenarios, or build and modify sheets, formulas, formatting, charts, tables, and conditional formatting. Also use when you need to inspect a workbook yourself as part of a larger task, or to deliver a polished, correct workbook.
---

> **Running in Claude Cowork?** The `witan` CLI isn't preinstalled in the sandbox. Read [references/cowork-setup.md](references/cowork-setup.md) first for install, PATH and network-allowlist steps.
>
> **No `witan` on PATH?** Prefix commands with `npx` (e.g. `npx witan xlsx exec ...`).

## What this is

Witan drives Excel through **sandboxed JavaScript run server-side** against the workbook. You write a short script; the engine recalculates and returns results. Two globals are always in scope:

- `xlsx` — the API namespace; every function is `xlsx.<name>(...)` (full surface in [references/api.d.ts](references/api.d.ts))
- `wb` — an already-open workbook handle; pass it as the **first argument** to every `xlsx.*` call

Supported inputs: `.xls`, `.xlsx`, `.xlsm` (legacy `.xls` is converted as needed). New workbooks are `.xlsx` only.

Two principles do most of the work:

1. **The engine computes; you don't.** Read every answer from the recalculated workbook (`result.touched[...]`, `readCell`, `sweepInputs`). Never recompute a spreadsheet value in JavaScript — it defeats the point and drifts from Excel.
2. **Writes are ephemeral until you opt in.** `exec` changes a server-side session only; the file on disk is untouched unless you pass `--save`. So you can explore, prototype, and run what-ifs with zero risk.

## Quick start

```bash
# Inspect: what sheets exist, and a labelled TSV of a range
witan xlsx exec report.xlsx --stdin <<'WITAN'
const sheets = await xlsx.listSheets(wb)
const tsv = await xlsx.readRangeTsv(wb, "Summary!A1:F20")
return { sheets, tsv }
WITAN
```

```bash
# Author: create a small, well-formed model (formulas, number formats, styled header)
witan xlsx exec model.xlsx --create --save --stdin <<'WITAN'
await xlsx.addSheet(wb, "Model")
await xlsx.setCells(wb, [
  { address: "Model!A1", value: "Revenue model", format: "@" },
  { address: "Model!A3", value: "Units" },     { address: "Model!B3", value: 1200 },
  { address: "Model!A4", value: "Price ($)" }, { address: "Model!B4", value: 9.5,  format: "$#,##0.00" },
  { address: "Model!A5", value: "Revenue ($)"},{ address: "Model!B5", formula: "=B3*B4", format: "$#,##0" },
])
await xlsx.setStyle(wb, "Model!A1", { font: { bold: true, size: 14 } })
await xlsx.setStyle(wb, "Model!A3:A5", { font: { bold: true } })
await xlsx.autoFitColumns(wb, "Model")       // size columns so nothing is clipped
return await xlsx.readCell(wb, "Model!B5")   // read the recalculated result
WITAN
```

```bash
# What-if: how does an output move as an input changes? (no file written)
witan xlsx exec model.xlsx --stdin <<'WITAN'
return await xlsx.sweepInputs(wb, {
  inputs: [{ address: "Model!B4", values: [8, 9.5, 11] }],
  outputs: ["Model!B5"],
  includeStats: true,
})
WITAN
```

```bash
# Verify: prove there are no formula errors before you call it done
witan xlsx calc model.xlsx --verify      # exit 2 if any error or any value changed
witan xlsx lint model.xlsx               # semantic checks (mixed currencies, bad lookups, ...)
```

## Commands

| Command                       | Use it to                                                     | Writes file?           |
| ----------------------------- | ------------------------------------------------------------- | ---------------------- |
| `xlsx exec <file>`            | Run a script: read, search, trace, author, edit, what-if      | only with `--save`     |
| `xlsx calc <file>`            | Recalculate the whole workbook / report every formula error   | yes, unless `--verify` |
| `xlsx lint <file>`            | Run semantic quality checks (rules like D001/D008/D032)       | no                     |
| `xlsx render <file> -r RANGE` | Render a range to PNG/WebP to see real layout, merges, charts | writes an image        |

`exec` is the workhorse. `calc`/`lint`/`render` are for verification and visual inspection.

### exec invocation

Provide **exactly one** code source: `--stdin`, `--expr`, `--code`, or `--script`.

- **Prefer `--stdin` with a single-quoted heredoc** (`<<'WITAN'`). It suppresses all shell expansion, so sheet names with spaces, apostrophes, or parentheses pass through untouched (`"'Workers'' Comp'!B5"`), and it batches many operations into one round-trip — which matters for latency.
- `--expr 'xlsx.listSheets(wb)'` is fine for a single expression with no special characters.
- `--script file.js --input-json '{"rate":1.05}'` for reusable, parameterised scripts; the JSON arrives as the `input` global.

Useful flags: `--save` (persist), `--create` (new `.xlsx`), `--json` (full response envelope), `--locale`, `--timeout-ms`, `--max-output-chars`. Top-level `await` works; `import` does not.

### The result you read back

`setCells` (and other writes) return:

```
{ touched: { "Sheet!Addr": "formatted text" }, changed: [...], errors: [...] }
```

Read outputs from `touched`. If an address you expected is missing from `touched`, that cell didn't recalculate — you probably have the wrong address. After any write, check `errors` is empty.

## Build quality in, not on (read this before authoring)

A workbook is judged on **correctness, structure, and readability** — not just that cells got filled. The tooling above lets you verify all three; these principles tell you what to aim for. Full detail in **[references/authoring.md](references/authoring.md)**; financial models have extra conventions in **[references/financial-models.md](references/financial-models.md)** — read it whenever the task is a financial/3-statement/valuation model.

The load-bearing rules:

- **Derived values are formulas, never hardcoded numbers.** A total is `=SUM(B2:B9)`, not the sum you worked out. The sheet must recompute when inputs change.
- **No magic numbers in formulas.** Put assumptions (rates, growth, multiples) in their own labelled input cells and reference them: `=C7*(1+$C$3)`, not `=C7*1.08`.
- **Separate inputs → calculations → outputs**, and keep formulas consistent across a row/projection so every period works the same way.
- **Format for meaning.** Currencies with units in the header (`Revenue ($mm)`), thousands separators, sane decimals, real dates with date formats, negatives and zeros formatted deliberately. Numbers right-aligned, labels left-aligned.
- **Style with restraint.** One professional font, a styled header row, light borders, whitespace between sections, merges only where they help. Don't over-decorate.
- **When editing an existing workbook, match its house style.** Render it first; preserve its fonts, formats, and conventions; make the smallest change that does the job. Don't impose your own formatting.
- **Cite hardcoded inputs.** Attach the source to the cell as a note (`setCells` `note`) or hyperlink, so numbers are traceable.
- **Deliver zero formula errors.** Use the verification loop below — this is a standard, not a nice-to-have.

## Workflows

### Build a new workbook

1. **Plan** the sheets and the input/calc/output split before writing.
2. **Build** with batched `setCells` (values + formulas + number formats together); add styling, then tables, charts, and conditional formatting as needed.
3. **Format** per the principles above; `autoFitColumns` and cap widths so nothing is clipped or absurdly wide.
4. **Verify** (see below).
5. **Save** with `--save` only once it's correct.

### Edit an existing workbook

1. **Understand it first:** `listSheets`, `describeSheet`, `readRangeTsv`; `render` the area you'll touch.
2. **Preserve style** — read neighbouring cells' formats and match them.
3. **Change** the minimum needed; keep formulas consistent with the rows around them.
4. **Verify**, then `--save`.

### What-if / sensitivity

1. **Find the output cell first.** Search the metric with `findCells` (use `context`); the label and the formula cell often share text — pick the formula cell. Disambiguate with `readRangeTsv` if needed.
2. **Confirm the link:** `traceToInputs(wb, outputAddr)` should reach the user's input. Filter large traces by `nearbyLabel`.
3. **Run it:** `setCells` then read `touched[outputAddr]`, or `sweepInputs` for many combinations in one call. Always report baseline → new value.

## Verification loop (before you say "done")

Run these, fix what they surface, repeat until clean:

- `xlsx calc <file> --verify` — non-mutating; exit code **2** if any formula error exists or any value changed. The final audit after edits.
- `xlsx lint <file>` — semantic issues a recalc won't catch: mixed currencies (D008), percent/non-percent addition (D009), double-counting via overlapping ranges (D001), approximate-match lookups on unsorted ranges (D002), empty-cell coercion (D003), and more. Scope with `-r` or `--only-rule`.
- `xlsx render <file> -r RANGE` — render the result and actually look: clipped text, broken merges, charts covering data, misaligned numbers. Use `--diff baseline.png` for before/after.
- In scripts: `lint` and `evaluateFormula` (e.g. assert a balance sheet balances, or a check row sums to zero) for programmatic invariants.

## Finding things in the API

Don't guess function names — the full typed surface is in [references/api.d.ts](references/api.d.ts).

Highlights by purpose: **read** `readCell/readRange/readRangeTsv`, **search** `findCells/findRows/describeSheet/tableLookup`, **trace** `traceToInputs/traceToOutputs/getCellPrecedents`, **compute** `sweepInputs/evaluateFormula`, **write** `setCells/scaleRange/copyRange/sortRange`, **structure** `insert*/delete*/autoFitColumns`, **style** `setStyle/setSheetProperties/setRowProperties/setColumnProperties`, **objects** `addListObject/addDataTable`, **charts** `addChart/setChart/listCharts`, **conditional formatting** `setConditionalFormatting`.

## Error guide

- `exactly one of --code, --script, --stdin, or --expr is required` — pass one code source only.
- `Import statements are not allowed` — no `import`; use the `xlsx` global.
- `EXEC_RESULT_TOO_LARGE` — return less; use `console.log`/`print` for big output instead of `return`.
- `Sheet 'X' not found` — enumerate with `listSheets`; mind quoting of names with spaces.
- `ADDRESS_MISSING_SHEET` — every address needs a sheet qualifier (`"Sheet1!A1"`).
- Shell quoting trouble with sheet names — switch to `--stdin <<'WITAN'`.
- `setCells` output missing — the cell isn't a dependent of what you changed; trace the chain.
- More codes (file/sheet/address/data/style) are listed in [references/authoring.md](references/authoring.md).

## References

- [references/api.d.ts](references/api.d.ts) — authoritative typed API for `exec`. Grep it; don't guess.
- [references/authoring.md](references/authoring.md) — how to build high-quality workbooks: layout, number formats, formula discipline, tables, charts, conditional formatting, verification, editing existing files.
- [references/financial-models.md](references/financial-models.md) — conventions for financial models (colour coding, finance number formats, assumptions, 3-statement/DCF/LBO structure, sourcing). Read when the deliverable is a financial model.
- [references/cowork-setup.md](references/cowork-setup.md) — installing and running Witan inside Claude Cowork or other sandboxes.
