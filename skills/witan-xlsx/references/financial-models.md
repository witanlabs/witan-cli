# Financial models

Read this whenever the deliverable is a financial model — a budget, forecast, three-statement model, DCF, LBO, valuation, or returns analysis. It layers finance conventions on top of [authoring.md](authoring.md); everything there still applies.

**Precedence:** explicit user instructions win, then the conventions of any template/workbook you're editing, then the defaults below. Never restyle someone's existing model into these conventions unless they ask.

These are the conventions practitioners expect. Following them makes a model legible to any analyst on sight; ignoring them makes even a correct model look amateur.

---

## 1. Cell colour conventions

The single most recognised convention in financial modelling is **colour-coding cells by what they are**, so a reader instantly knows what is safe to change. Apply via `setStyle` font `color` (and a fill for the highlight):

| Cell is…                                         | Colour                    | Hex       | Set with      |
| ------------------------------------------------ | ------------------------- | --------- | ------------- |
| A hardcoded input / assumption a user may change | Blue                      | `#0000FF` | `font.color`  |
| A formula / calculation                          | Black (default)           | `#000000` | leave default |
| A link to another sheet in the same workbook     | Green                     | `#008000` | `font.color`  |
| A link to another file                           | Red                       | `#FF0000` | `font.color`  |
| A key assumption that needs attention            | (any text) on yellow fill | `#FFFF00` | `fill.color`  |

The discipline that matters most: **inputs are blue, formulas are black.** If every blue cell is a lever and every black cell is derived, a reader can audit the model in seconds and knows exactly which cells to touch for a scenario.

```js
await xlsx.setStyle(wb, "Assumptions!B3", { font: { color: "#0000FF" } }); // an input
// calculations stay black/default; cross-sheet pulls go green:
await xlsx.setStyle(wb, "Summary!B10", { font: { color: "#008000" } });
```

---

## 2. Number formats for finance

In addition to the general cookbook in [authoring.md](authoring.md):

- **Negatives in parentheses, usually red:** `"#,##0;[Red](#,##0)"` → `(1,235)`. Standard for financial statements.
- **Zeros as a dash:** include a third section, e.g. `"#,##0;(#,##0);\"-\""`, so empty-looking zeros don't clutter. Apply to percentages too.
- **Multiples with an x:** `"0.0\"x\""` → `5.2x` for EV/EBITDA, P/E, leverage, returns.
- **Units in the header, scaled in the cells.** State `($mm)` or `($000)` in the column/section header and keep the cells as bare scaled numbers. Don't repeat `$` on every row.
- **Percentages to one decimal** by default (`"0.0%"`); basis points where precision matters.
- **Years are labels, not numbers.** A header of `2024 2025 2026` should be text so it never renders as `2,024`.
- Keep decimals consistent down a column, and right-align all of it.

---

## 3. Structure

- **Assumptions and drivers live apart from calculations** — ideally their own sheet (e.g. `Assumptions`), or a clearly fenced input block. Every growth rate, margin, multiple, and rate is a labelled input cell, referenced by formula elsewhere (see "no magic numbers" in [authoring.md](authoring.md)).
- **One period per column,** with an identical formula across the whole projection so every period is computed the same way. This is the backbone of a forecast; inconsistent period formulas are the classic error `lint`/`calc --verify` will help you catch.
- **A scenario switch** (e.g. an input cell selecting Base/Bull/Bear that drives `INDEX` (or `CHOOSE`) lookups) keeps cases in one model instead of duplicating sheets.
- **Inputs vs outputs are visually distinct** — colour coding (above) plus placement (inputs at top/left, outputs in a summary).

---

## 4. How statements are laid out

Financial statements follow a presentation grammar readers expect — match it:

- **Totals sum a contiguous block directly above (or beside) them** — `=SUM(B5:B11)` over the line items, so the range is obvious and auditable. Avoid totals that add scattered cells.
- **Distinguish a display total from one that feeds logic.** A `=SUM(...)` shown for the reader, with no downstream dependents, is presentation. The moment another formula consumes a total, give it its own explicit, labelled line rather than summing the block inline — so the dependency is visible and the range can't silently shift underneath it.
- **Rule a border above each total** spanning the full width of the figures (and the label column), and bold the total row. Use `setStyle` `border.top` (`thin` or `medium`).
- **Section headers** (e.g. "Income Statement", "Cash Flow") are merged or `centerContinuousSpan` banners, left-justified, dark fill with white bold text.
- **Period headers right-aligned** to sit over their right-aligned figures; **row labels left-justified**; **sub-metrics indented** beneath their parent (`alignment.indent`), e.g. a `% growth` line under a revenue line.

(Turning gridlines off on these sheets — see [authoring.md](authoring.md) — lets the borders carry the structure.)

---

## 5. Three-statement, DCF, and LBO specifics

- **Link the statements.** Net income flows IS → retained earnings on the BS and → the top of the CF; the CF's ending cash ties to the BS cash line. Build these as live links (green), not retyped numbers.
- **Build checks.** Add explicit check rows that must be zero — e.g. `=Assets-(Liabilities+Equity)` for the balance sheet, or a cash-tie check — and surface them. Assert them in-script with `evaluateFormula` and flag non-zero with conditional formatting.
- **Accumulate balances as opening + movements = closing,** with the opening linked to the prior period's closing — not a half-anchored `=SUM($B5:F5)` that creeps across the row. This structure is self-checking and sidesteps the range-anchor errors cumulative sums invite.
- **Handle deliberate circularity.** Interest-on-average-debt and revolver sweeps create intended circular references. Enable iterative calculation rather than fighting it:

  ```js
  await xlsx.setWorkbookProperties(wb, {
    iterativeCalculation: {
      enabled: true,
      maxIterations: 100,
      maxChange: 0.001,
    },
  });
  ```

  After enabling, `setCells`/`calc` converge using these settings; if a model won't converge you'll get convergence diagnostics — check your circular path.

- **Sensitivity tables.** For a clean grid a reader can see (e.g. value vs WACC and growth), author a What-If data table with `addDataTable`. For quick programmatic sweeps you read back yourself, use `sweepInputs` (cartesian for a full grid, with `includeStats`).
- **Use `XNPV`/`XIRR`, not `NPV`/`IRR`, for dated cash flows.** `NPV` assumes evenly-spaced periods and discounts the first cash flow by a full period — wrong for end-of-period or irregular dates. `XNPV` takes an explicit date column and discounts correctly. Reach for `NPV`/`IRR` only when the periods are genuinely uniform and you've handled the period-1 timing.

---

## 6. Sourcing every input

Finance numbers must be traceable. **Cite every hardcoded input** in a cell note (or a hyperlink) with enough detail to find it again — source, period, and where in the source:

```js
await xlsx.setCells(wb, [
  {
    address: "Assumptions!B7",
    value: 392000,
    format: "#,##0",
    note: { text: "FY2024 revenue, from the company's 10-K income statement." },
  },
]);
```

A model whose assumptions all carry sources is one a reviewer can actually sign off on.

---

## 7. Finance-specific verification

On top of the general loop in [authoring.md](authoring.md), pay attention to:

- **`lint` D008 (Error): mixed currencies** added or aggregated together — a real correctness bug in cross-region models.
- **`lint` D009 / D023:** percent or other semantic formats mixed into additive contexts.
- **`lint` D001:** double counting from overlapping `SUM` ranges — easy to introduce when statements share line items.
- **Balance and tie checks** via `evaluateFormula` (assert they're zero) before you save.
- **`calc --verify`** as the final gate: zero formula errors, nothing unexpectedly changed.
