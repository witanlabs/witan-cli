# Loan Amortization with Circular References

Build a 12-month loan amortization schedule in loan-amortization.xlsx.

## Parameters (cells H1:I3)

| Cell | Label | Value |
|------|-------|-------|
| H1/I1 | Annual Rate | 6% |
| H2/I2 | Monthly Rate | 0.5% |
| H3/I3 | Fee Rate | 1% |

## Schedule (columns A–F, rows 1–13)

Row 1 is headers. Rows 2–13 are months 1–12.

- **A**: Month number (1–12)
- **B**: Opening Balance (month 1 = $100,000; thereafter = prior month's Closing Balance)
- **C**: Interest = Opening Balance × Monthly Rate (reference $I$2)
- **D**: Payment — a fixed amount that fully amortizes the loan over 12 months,
  accounting for the servicing fee. Put the computed value as a number, not a formula.
  - Derivation: because `Closing = (Opening×(1+r) − Payment) / (1−f)`, the
    effective monthly multiplier is `R = (1+r)/(1−f)` and the payment is
    `P = Principal × (1−f) × R^n × (R−1) / (R^n − 1)` where n = 12.
- **E**: Closing Balance = Opening + Interest − Payment + Fee
  - Formula: `=B{row}+C{row}-D{row}+F{row}` (circular with F)
- **F**: Servicing Fee = Closing Balance × Fee Rate
  - Formula: `=E{row}*$I$3` (circular with E)

## Important

- The formulas in columns E and F create a **circular reference** — each row's
  closing balance depends on its fee, and the fee depends on the closing balance.
- Use **cell references** (e.g. `$I$2`, `$I$3`), not hardcoded values.
- **Iterative calculation** is already enabled in the workbook. After building
  the schedule, run `witan xlsx calc loan-amortization.xlsx --show-touched` to
  converge the circular references. Month 12's closing balance should be
  approximately $0.
