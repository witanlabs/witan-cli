---
name: xlsx-code-mode
description: "Read, explore, understand, create, and modify Excel workbooks (.xls, .xlsx, .xlsm). You cannot read Excel files with cat, head, or normal file-reading tools — running JavaScript against the workbook via `witan xlsx exec` is the only way to open or inspect them. Trigger when you or the user need to open or look inside a workbook, find its sheets or where data lives, read cells/rows/columns/ranges, search for values or labels, trace how a figure is calculated, or run what-if scenarios; and equally when asked to create a new workbook or financial model from scratch, add or edit formulas, charts, formatting, tables, or data validation, or change an existing model without breaking its formulas or references. Trigger whenever a spreadsheet is referenced by name or path — even casually ('check report.xlsx', 'build me a model') — and when you need to inspect a workbook as part of a larger task."
---

> **Running in Claude Cowork?** The `witan` CLI isn't preinstalled — see [references/cowork-setup.md](references/cowork-setup.md) for install steps.

## Goal

Two jobs, one tool:

- **Read & Analyze** an existing workbook — find data, trace how figures are calculated, answer what-if questions — without guessing or corrupting it.
- **Produce & Edit** workbooks that are correct, polished, and idiomatic: formula-driven, sensibly formatted, and matched to the workbook's domain.

You are judged on correctness, layout, readability, and idiomatic style. **A workbook with formula errors is not finished.**

## The tool

`witan` is a command-line tool driving a server-side spreadsheet engine. Its workhorse, **`witan xlsx exec`**, runs sandboxed JavaScript against a workbook opened server-side; siblings `calc`, `lint`, and `render` handle verification and previews (below). Invoke with `--stdin` and a single-quoted heredoc (safe for every sheet name, no shell escaping):

```bash
witan xlsx exec report.xlsx --stdin <<'WITAN'
const sheets = await xlsx.listSheets(wb)
const tsv = await xlsx.readRangeTsv(wb, { sheet: "Summary", from: {row:1,col:1}, to: {row:20,col:8} })
return { sheets, tsv }
WITAN
```

- **Globals** — `xlsx` (the API) and `wb` (the open workbook, passed first to every call). Top-level `await` works; no `import`s.
- **`--create`** makes a new workbook (path need not exist; `.xlsx` only).
- **`--save`** persists changes. Without it every write is **ephemeral** — it applies in the server session, recalculates, then is discarded; so reads and what-ifs never risk the file, and each run starts clean from the original.
- **Read answers from `result.touched["Sheet!A1"]`, never recompute them in JS** — after a write the recalculated value is there. Calculating it yourself defeats the engine.

After reading this file, you MUST read **[references/api.md](references/api.md)** before your first `witan xlsx exec` call — the function surface is large and not guessable.

`api.md` groups functions under headings you can grep for — Reading, Searching, Tracing, Computing, Validating, Rendering, Charts, Conditional Formatting, Images, Writing — each with full signatures.

### Work efficiently (latency matters)

- Batch independent reads/writes into **one** `exec`; the heredoc takes a whole script, so prefer one rich call over many small ones.
- Don't re-read what you already pulled earlier in the task — reuse it.
- **Exception — what-if:** there, deliberately split into separate `exec` calls so you can review what you found before editing (see below).

## Quality floor

Applies to every workbook you create or change. **When editing an existing workbook, inspect or `render` it first and match its established conventions — the rules below are defaults for new or unspecified work, never a licence to restyle someone's model.**

- **Formulas, not values** — every derived number is a formula so the sheet stays live.
  - ✅ `setCells(wb, [{ address: "B10", formula: "=SUM(B2:B9)" }])`
  - ❌ summing in JS and writing `{ address: "B10", value: 4200 }`
- **No magic numbers** — reference an input cell, don't bury a constant.
  - ✅ `=B5*(1+$B$6)`  ❌ `=B5*1.05`
- **Separate inputs from logic** — assumptions (rates, growth, multiples) live in their own labelled cells or an Inputs area, never inside formulas.
- **Right references** — absolute (`$B$6`) vs relative so fills and copies behave.
- **Format to the data type** — real dates with a date format (`yyyy-mm-dd`), not date-looking text; thousands separators for money/counts; sensible decimals on percentages; state units in the header when ambiguous (`Revenue ($000s)`).
- **Lay it out for a human** — styled header row; numbers right-aligned, labels left; sane column widths (cap autofit so nothing runs off-screen); modest row heights; whitespace between sections.
- **Stay valid** — give every Excel Table a unique, explicit name; prefix literal text starting with `=` using `'` so it isn't read as a formula.

Domain conventions go beyond this floor. For financial models, valuations, projections, or IB work, read **[references/domains/financial-modelling.md](references/domains/financial-modelling.md)**.

## Reading & understanding a workbook

- **Get the lay of the land** — `listSheets` (sheets, used ranges, cross-sheet deps), then `describeSheet` for a structure map, or `readRangeTsv` to dump a region cheaply.
- **Find things** — `findCells` / `findRows` are fuzzy and case-insensitive; pass synonym arrays or a `RegExp`. `tableLookup` reads a value by row + column label inside a detected table.
- **Disambiguate** — a label and the formula cell beside it often share text. Use `context` or read surrounding rows with `readRangeTsv`, then pick the **formula** cell, not the label.
- **Trace calculations** — `getCellPrecedents` / `getCellDependents` for local hops; `traceToInputs` / `traceToOutputs` for the full chain. Traces can be huge — filter by `nearbyLabel`, don't print them all.

### What-if / sensitivity

For "what happens to Y if X changes?", use **two separate `exec` calls** so you verify before editing:

1. **Locate the output cell** (first call) — search for what's asked about, review candidates, confirm you have the formula cell.
2. **Trace + change + read** (second call) — `traceToInputs(output)` to confirm X actually drives Y, then `setCells` to change X and read the new value from `result.touched[output]`. Missing from `touched` ⇒ it didn't recalc ⇒ wrong cell.
3. **Report baseline → new**, and check `result.errors` is empty (a new error means your edit broke something downstream).

For many values at once, use `sweepInputs` — all combinations in one call with structured before/after stats. **Circular/iterative models:** if iterative calc is enabled, `setCells` honours it and non-convergence shows up as errors; use `witan xlsx calc` for a standalone full-workbook pass.

## Authoring a workbook

Build in this order — it minimises rework and round-trips:

1. **Plan the structure** — sheets, and where inputs vs calculations vs outputs live. Inputs first.
2. **Write content in bulk** — headers, data, and formulas in as few `setCells` / `exec` calls as possible. Seed a formula once, then fill across the range rather than writing each cell.
3. **Format** — number formats, header styling, widths/heights, alignment, borders (Quality floor above; finance models also follow the domain reference).
4. **Add interactivity as needed** — data validation for categorical inputs, conditional formatting, then charts/tables.
5. **Verify** (below) before calling it done.

Prototype ephemerally (`--create` without `--save`) while iterating; add `--save` only once the structure is right.

## Verify before done — mandatory

A workbook with formula errors is not delivered. Before finishing any authoring or edit:

- **Recalculate and check errors** — `witan xlsx calc model.xlsx` must report `0 errors`. It prints every error with address and formula, and exits non-zero. Fix and repeat until clean.
- **Spot-check key ranges** — `readRangeTsv` (with formulas) on totals, a few sample references, and edge rows.
- **Confirm layout** — `witan xlsx render` or `previewStyles` on the headline range; headers, merges, number formats, and charts look as intended.
- **Lint for logic errors** — `witan xlsx lint model.xlsx` flags what `calc` can't: double-counting, approximate-match lookups on unsorted data, mixed currencies/units. Review and resolve or knowingly accept each finding; exits 2 on any.

Then report what you built (sheets / ranges), and for what-ifs the baseline → new values.
