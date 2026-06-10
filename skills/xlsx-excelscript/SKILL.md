---
name: xlsx-excelscript
description: "Read, explore, understand, create, and modify Excel workbooks (.xls, .xlsx, .xlsm) by running Office Scripts (ExcelScript) JavaScript against the workbook via `witan xlsx exec`. You cannot read Excel files with cat, head, or normal file-reading tools — exec is the only way to open or inspect them. Trigger when you or the user need to open or look inside a workbook, find its sheets or where data lives, read cells/rows/columns/ranges, search for values or labels, trace how a figure is calculated, or run what-if scenarios; and equally when asked to create a new workbook or financial model from scratch, add or edit formulas, charts, formatting, tables, or data validation, or change an existing model without breaking its formulas or references. Trigger whenever a spreadsheet is referenced by name or path — even casually ('check report.xlsx', 'build me a model') — and when you need to inspect a workbook as part of a larger task."
---

> **Running in Claude Cowork?** The `witan` CLI isn't preinstalled — see [references/cowork-setup.md](references/cowork-setup.md) for install steps.

## Goal

Two jobs, one tool:

- **Read & Analyze** an existing workbook — find data, trace how figures are calculated, answer what-if questions — without guessing or corrupting it.
- **Produce & Edit** workbooks that are correct, polished, and idiomatic: formula-driven, sensibly formatted, and matched to the workbook's domain.

You are judged on correctness, layout, readability, and idiomatic style. **A workbook with formula errors is not finished.**

## The tool

`witan xlsx exec` runs sandboxed JavaScript against a workbook opened server-side; this skill drives it with the **Office Scripts (ExcelScript)** dialect. Siblings `calc`, `lint`, and `render` handle verification and previews (below). Invoke with `--stdin` and a single-quoted heredoc (safe for every sheet name, no shell escaping):

```bash
witan xlsx exec report.xlsx --stdin <<'JS'
// @office-script
function main(workbook) {
  const sheets = workbook.getWorksheets().map(s => s.getName());
  const used = workbook.getWorksheet("Summary").getUsedRange();
  return { sheets, address: used.getAddress(), values: used.getValues() };
}
JS
```

- **`// @office-script` MUST be the first line** — it selects the ExcelScript dialect. Omit it and the script doesn't error: it silently runs in the other dialect, never calls `main`, and returns `null`. Always lead with it.
- **`function main(workbook)` is the entry point** — required. The engine calls it with the workbook and returns whatever you return (must be JSON-serializable). With `--input-json`, the parsed value arrives as a second arg: `function main(workbook, input)`.
- **ExcelScript is synchronous** — `getValue()` after `setFormula` returns the **recalculated** value directly (no `await` needed); read results this way, don't recompute in JS. No `import`s.
- **`--create`** makes a new workbook (path need not exist; `.xlsx` only). It **starts empty** — call `workbook.addWorksheet(name)` before `getActiveWorksheet()`.
- **`--save`** persists changes. Without it every write is **ephemeral** — it applies in the server session, recalculates, then is discarded; so reads and what-ifs never risk the file, and each run starts clean from the original.

After reading this file, you MUST read **[references/api.md](references/api.md)** before your first `witan xlsx exec` call — the ExcelScript surface is large and not guessable.

`api.md` has the CLI flags, the `function main(workbook)` contract, and worked recipes you can grep — reading, what-if, conditional formatting, charts, images, tables — plus what ExcelScript can't do, then points on to the full type surface in **[references/excelscript.d.ts](references/excelscript.d.ts)**.

### Work efficiently (latency matters)

- Batch independent reads/writes into **one** `exec`; `main` takes a whole script, so prefer one rich call over many small ones.
- Don't re-read what you already pulled earlier in the task — reuse it.
- **Exception — what-if:** there, deliberately split into separate `exec` calls so you can review what you found before editing (see below).

## Quality floor

Applies to every workbook you create or change. **When editing an existing workbook, inspect or `render` it first and match its established conventions — the rules below are defaults for new or unspecified work, never a licence to restyle someone's model.**

- **Formulas, not values** — every derived number is a formula so the sheet stays live.
  - ✅ `sheet.getRange("B10").setFormula("=SUM(B2:B9)")`
  - ❌ summing in JS and writing `sheet.getRange("B10").setValue(4200)`
- **No magic numbers** — reference an input cell, don't bury a constant.
  - ✅ `=B5*(1+$B$6)`  ❌ `=B5*1.05`
- **Separate inputs from logic** — assumptions (rates, growth, multiples) live in their own labelled cells or an Inputs area, never inside formulas.
- **Right references** — absolute (`$B$6`) vs relative so fills and copies behave.
- **Format to the data type** — real dates with a date format (`yyyy-mm-dd`), not date-looking text; thousands separators for money/counts; sensible decimals on percentages; state units in the header when ambiguous (`Revenue ($000s)`).
- **Lay it out for a human** — styled header row; numbers right-aligned, labels left; sane column widths (cap autofit so nothing runs off-screen); modest row heights; whitespace between sections.
- **Stay valid** — give every Excel Table a unique, explicit name; prefix literal text starting with `=` using `'` so it isn't read as a formula.

Domain conventions go beyond this floor. For financial models, valuations, projections, or IB work, read **[references/domains/financial-modelling.md](references/domains/financial-modelling.md)**.

## Reading & understanding a workbook

- **Get the lay of the land** — `workbook.getWorksheets()` for the sheet list (`.map(s => s.getName())`); `sheet.getUsedRange()` for the populated region, then `.getValues()` to dump it cheaply (or `.getFormulas()` to see the logic). `getUsedRange(true)` trims to value-bearing cells.
- **Find things** — `range.find(text, { completeMatch, matchCase })` returns the matching cell, or a nullish value if absent. Matching is exact/substring, **not** fuzzy or synonym-aware — search the literal label, or pull `getValues()` and scan in JS. Tables come from `workbook.getTables()` (`table.getColumnByName(name)`).
- **Disambiguate** — a label and the formula cell beside it often share text. Read surrounding cells with `getValues()`/`getFormulas()`, then pick the **formula** cell (the one whose `getFormula()` starts with `=`), not the label.
- **Trace calculations** — ExcelScript has **no dependency-tracing API** (no precedents/dependents). Trace by hand: read the target cell's `getFormula()`, follow each cell reference it names, and repeat. Filter to the formulas you care about; don't dump whole sheets.

### What-if / sensitivity

For "what happens to Y if X changes?", use **two separate `exec` calls** so you verify before editing:

1. **Locate the output cell** (first call) — search for what's asked about, read its `getFormula()`, confirm you have the formula cell.
2. **Change + read** (second call) — read the baseline `getValue()`, `setValue` on input X, then read Y's `getValue()` again; recalc is automatic, so the new number is there immediately. If Y didn't move, X doesn't drive it ⇒ wrong cell.
3. **Report baseline → new**, and re-read a few downstream cells (or run `witan xlsx calc`) to confirm your edit broke nothing.

There is **no `sweepInputs` helper** — for many combinations, loop inside one `main`: set each input, read the outputs, collect the before/after rows, and return them. **Circular/iterative models:** a standalone full-workbook pass is `witan xlsx calc`; non-convergence shows up there as errors.

## Authoring a workbook

Build in this order — it minimises rework and round-trips:

1. **Plan the structure** — sheets, and where inputs vs calculations vs outputs live. Inputs first.
2. **Write content in bulk** — headers, data, and formulas in as few calls as possible. Prefer `setValues`/`setFormulas` with a 2-D array over a region to writing each cell. **Gotcha:** unlike desktop Excel, `setFormula` on a multi-cell range writes the *same literal formula* to every cell — it does **not** adjust references relatively (`B1:B3` ← `"=A1*2"` makes all three `=A1*2`). To fill a relative formula down, build the per-cell strings and pass them to `setFormulas` (`[["=A1*2"],["=A2*2"],["=A3*2"]]`).
3. **Format** — number formats (`setNumberFormat`), header styling (`getFormat().getFont()/getFill()`), widths/heights, alignment, borders (`getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop)`). Quality floor above; finance models also follow the domain reference.
4. **Add interactivity as needed** — data validation for categorical inputs, conditional formatting, then charts/tables.
5. **Verify** (below) before calling it done.

Prototype ephemerally (`--create` without `--save`) while iterating; add `--save` only once the structure is right.

## Verify before done — mandatory

A workbook with formula errors is not delivered. Before finishing any authoring or edit:

- **Recalculate and check errors** — `witan xlsx calc model.xlsx` must report `0 errors`. It prints every error with address and formula, and exits non-zero. Fix and repeat until clean.
- **Spot-check key ranges** — read formulas back (`getFormulas`) on totals, a few sample references, and edge rows.
- **Confirm layout** — `witan xlsx render` on the headline range; headers, merges, number formats, and charts look as intended.
- **Lint for logic errors** — `witan xlsx lint model.xlsx` flags what `calc` can't: double-counting, approximate-match lookups on unsorted data, mixed currencies/units. Review and resolve or knowingly accept each finding; exits 2 on any.

Then report what you built (sheets / ranges), and for what-ifs the baseline → new values.
