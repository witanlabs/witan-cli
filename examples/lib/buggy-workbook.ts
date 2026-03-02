import ExcelJS from 'exceljs';

/**
 * Generate a multi-region P&L workbook with 5 planted formula bugs
 * for demonstrating the xlsx-verify (lint) skill.
 *
 * Sheets:
 *   1. P&L        — income statement with North America, Europe, Asia-Pacific
 *   2. FX Rates   — currency conversion table (with duplicate EUR key)
 *   3. Headcount  — department headcount (clean filler)
 *
 * Planted bugs:
 *   D001 — P&L!F19  double-counts Net Income via SUM range overlap
 *   D008 — P&L!F2   adds USD + EUR without currency conversion
 *   D003 — P&L!D9   missing Asia-Pacific R&D value (empty cell in SUM)
 *   D007 — P&L!B22  VLOOKUP on duplicate "EUR" key in FX Rates
 *   D009 — P&L!B23  adds dollar amount (Net Income) to percentage (GP margin)
 */
export async function createBuggyWorkbook(filePath: string): Promise<void> {
  const wb = new ExcelJS.Workbook();

  // ── Sheet 1: P&L ──────────────────────────────────────────────────
  const pnl = wb.addWorksheet('P&L');

  // Column widths
  pnl.getColumn(1).width = 22; // labels
  for (let c = 2; c <= 7; c++) pnl.getColumn(c).width = 18;

  // Row 1 — headers
  pnl.getRow(1).values = [
    '',
    'North America',
    'Europe',
    'Asia-Pacific',
    '',
    'Group Total',
    '% of Revenue',
  ];
  pnl.getRow(1).font = { bold: true };

  // Currency formats
  const usdFmt = '$#,##0';
  const eurFmt = '[$EUR] #,##0';

  function setRowFormats(row: number): void {
    pnl.getCell(row, 2).numFmt = usdFmt; // B — North America (USD)
    pnl.getCell(row, 3).numFmt = eurFmt; // C — Europe (EUR)
    pnl.getCell(row, 4).numFmt = usdFmt; // D — Asia-Pacific (USD)
    pnl.getCell(row, 6).numFmt = usdFmt; // F — Group Total (USD)
    pnl.getCell(row, 7).numFmt = '0.0%'; // G — % of Revenue
  }

  // Row 2 — Revenue
  pnl.getCell('A2').value = 'Revenue';
  pnl.getCell('B2').value = 5200000;
  pnl.getCell('C2').value = 3800000;
  pnl.getCell('D2').value = 2100000;
  // BUG D008: SUM mixes USD + EUR without conversion
  pnl.getCell('F2').value = { formula: 'SUM(B2:D2)' };
  pnl.getCell('G2').value = 1; // 100%
  pnl.getCell('G2').numFmt = '0.0%';
  setRowFormats(2);

  // Row 3 — COGS
  pnl.getCell('A3').value = 'Cost of Goods Sold';
  pnl.getCell('B3').value = 2080000;
  pnl.getCell('C3').value = 1520000;
  pnl.getCell('D3').value = 840000;
  pnl.getCell('F3').value = { formula: 'SUM(B3:D3)' };
  pnl.getCell('G3').value = { formula: 'F3/F2' };
  setRowFormats(3);

  // Row 4 — Gross Profit
  pnl.getCell('A4').value = 'Gross Profit';
  pnl.getRow(4).font = { bold: true };
  pnl.getCell('B4').value = { formula: 'B2-B3' };
  pnl.getCell('C4').value = { formula: 'C2-C3' };
  pnl.getCell('D4').value = { formula: 'D2-D3' };
  pnl.getCell('F4').value = { formula: 'F2-F3' };
  pnl.getCell('G4').value = { formula: 'F4/F2' };
  setRowFormats(4);

  // Row 5 — blank
  // Row 6 — section header
  pnl.getCell('A6').value = 'Operating Expenses';
  pnl.getCell('A6').font = { bold: true, italic: true };

  // Rows 7-10 — OpEx line items
  const opexItems: [string, number, number, number | null][] = [
    ['Salaries & Benefits', 1200000, 950000, 580000],
    ['Marketing', 320000, 280000, 190000],
    ['Research & Development', 480000, 350000, null], // BUG D003: Asia-Pac R&D missing
    ['Rent & Facilities', 180000, 210000, 120000],
  ];

  for (let i = 0; i < opexItems.length; i++) {
    const row = 7 + i;
    const [label, na, eu, apac] = opexItems[i];
    pnl.getCell(`A${row}`).value = label;
    pnl.getCell(`B${row}`).value = na;
    pnl.getCell(`C${row}`).value = eu;
    if (apac !== null) pnl.getCell(`D${row}`).value = apac;
    // D9 is intentionally left empty (bug D003)
    pnl.getCell(`F${row}`).value = { formula: `SUM(B${row}:D${row})` };
    pnl.getCell(`G${row}`).value = { formula: `F${row}/F$2` };
    setRowFormats(row);
  }

  // Row 11 — Total OpEx
  pnl.getCell('A11').value = 'Total Operating Expenses';
  pnl.getRow(11).font = { bold: true };
  pnl.getCell('B11').value = { formula: 'SUM(B7:B10)' };
  pnl.getCell('C11').value = { formula: 'SUM(C7:C10)' };
  pnl.getCell('D11').value = { formula: 'SUM(D7:D10)' };
  pnl.getCell('F11').value = { formula: 'SUM(F7:F10)' };
  pnl.getCell('G11').value = { formula: 'F11/F$2' };
  setRowFormats(11);

  // Row 12 — blank
  // Row 13 — EBITDA
  pnl.getCell('A13').value = 'EBITDA';
  pnl.getRow(13).font = { bold: true };
  pnl.getCell('B13').value = { formula: 'B4-B11' };
  pnl.getCell('C13').value = { formula: 'C4-C11' };
  pnl.getCell('D13').value = { formula: 'D4-D11' };
  pnl.getCell('F13').value = { formula: 'F4-F11' };
  pnl.getCell('G13').value = { formula: 'F13/F$2' };
  setRowFormats(13);

  // Row 14 — D&A
  pnl.getCell('A14').value = 'Depreciation & Amortization';
  pnl.getCell('B14').value = 260000;
  pnl.getCell('C14').value = 190000;
  pnl.getCell('D14').value = 95000;
  pnl.getCell('F14').value = { formula: 'SUM(B14:D14)' };
  pnl.getCell('G14').value = { formula: 'F14/F$2' };
  setRowFormats(14);

  // Row 15 — EBIT
  pnl.getCell('A15').value = 'EBIT';
  pnl.getRow(15).font = { bold: true };
  pnl.getCell('B15').value = { formula: 'B13-B14' };
  pnl.getCell('C15').value = { formula: 'C13-C14' };
  pnl.getCell('D15').value = { formula: 'D13-D14' };
  pnl.getCell('F15').value = { formula: 'F13-F14' };
  pnl.getCell('G15').value = { formula: 'F15/F$2' };
  setRowFormats(15);

  // Row 16 — blank
  // Row 17 — Tax Rate
  pnl.getCell('A17').value = 'Tax Rate';
  pnl.getCell('B17').value = 0.21;
  pnl.getCell('C17').value = 0.25;
  pnl.getCell('D17').value = 0.18;
  pnl.getCell('B17').numFmt = '0%';
  pnl.getCell('C17').numFmt = '0%';
  pnl.getCell('D17').numFmt = '0%';

  // Row 18 — Tax
  pnl.getCell('A18').value = 'Tax';
  pnl.getCell('B18').value = { formula: 'B15*B17' };
  pnl.getCell('C18').value = { formula: 'C15*C17' };
  pnl.getCell('D18').value = { formula: 'D15*D17' };
  pnl.getCell('F18').value = { formula: 'SUM(B18:D18)' };
  pnl.getCell('G18').value = { formula: 'F18/F$2' };
  setRowFormats(18);

  // Row 19 — Net Income
  pnl.getCell('A19').value = 'Net Income';
  pnl.getRow(19).font = { bold: true };
  pnl.getCell('B19').value = { formula: 'B15-B18' };
  pnl.getCell('C19').value = { formula: 'C15-C18' };
  pnl.getCell('D19').value = { formula: 'D15-D18' };
  // BUG D001: double-counts Net Income — SUM(B13:D19) overlaps with SUM(B19:D19)
  pnl.getCell('F19').value = { formula: 'SUM(B13:D19)+SUM(B19:D19)' };
  pnl.getCell('G19').value = { formula: 'F19/F$2' };
  setRowFormats(19);

  // Row 20-21 — blank
  // Row 22 — EUR/USD rate lookup
  pnl.getCell('A22').value = 'EUR/USD Rate';
  // BUG D007: VLOOKUP hits duplicate "EUR" key in FX Rates
  pnl.getCell('B22').value = {
    formula: "VLOOKUP(\"EUR\",'FX Rates'!A1:B10,2,FALSE)",
  };
  pnl.getCell('C22').value = 'Europe Revenue (USD)';
  pnl.getCell('D22').value = { formula: 'C2*B22' };
  pnl.getCell('D22').numFmt = usdFmt;

  // Row 23 — NI + GP Margin (type confusion)
  pnl.getCell('A23').value = 'NI + GP Margin';
  // BUG D009: adds Net Income ($) to Gross Profit margin (%) — unit mismatch
  pnl.getCell('B23').value = { formula: 'F19+G4' };

  // ── Sheet 2: FX Rates ─────────────────────────────────────────────
  const fx = wb.addWorksheet('FX Rates');

  fx.getColumn(1).width = 12;
  fx.getColumn(2).width = 10;
  fx.getColumn(3).width = 14;

  fx.getRow(1).values = ['Currency', 'Rate', 'Updated'];
  fx.getRow(1).font = { bold: true };

  const fxData: [string, number, string][] = [
    ['GBP', 1.27, '2024-01-15'],
    ['EUR', 1.09, '2024-01-15'],
    ['USD', 1.0, '2024-01-15'],
    ['EUR', 1.08, '2023-12-31'], // duplicate EUR key (bug D007 trigger)
  ];

  for (let i = 0; i < fxData.length; i++) {
    const row = i + 2;
    const [currency, rate, updated] = fxData[i];
    fx.getCell(`A${row}`).value = currency;
    fx.getCell(`B${row}`).value = rate;
    fx.getCell(`C${row}`).value = new Date(updated);
    fx.getCell(`C${row}`).numFmt = 'yyyy-mm-dd';
  }

  // ── Sheet 3: Headcount ────────────────────────────────────────────
  const hc = wb.addWorksheet('Headcount');

  hc.getColumn(1).width = 22;
  hc.getColumn(2).width = 16;
  hc.getColumn(3).width = 16;
  hc.getColumn(4).width = 16;
  hc.getColumn(5).width = 12;

  hc.getRow(1).values = ['Department', 'North America', 'Europe', 'Asia-Pacific', 'Total'];
  hc.getRow(1).font = { bold: true };

  const departments: [string, number, number, number][] = [
    ['Engineering', 45, 32, 28],
    ['Sales', 30, 22, 15],
    ['Marketing', 12, 10, 8],
    ['Finance', 8, 6, 4],
    ['HR', 5, 4, 3],
  ];

  for (let i = 0; i < departments.length; i++) {
    const row = i + 2;
    const [dept, na, eu, apac] = departments[i];
    hc.getCell(`A${row}`).value = dept;
    hc.getCell(`B${row}`).value = na;
    hc.getCell(`C${row}`).value = eu;
    hc.getCell(`D${row}`).value = apac;
    hc.getCell(`E${row}`).value = { formula: `SUM(B${row}:D${row})` };
  }

  const totalRow = departments.length + 2;
  hc.getCell(`A${totalRow}`).value = 'Total';
  hc.getCell(`A${totalRow}`).font = { bold: true };
  for (const col of ['B', 'C', 'D', 'E']) {
    hc.getCell(`${col}${totalRow}`).value = {
      formula: `SUM(${col}2:${col}${totalRow - 1})`,
    };
    hc.getCell(`${col}${totalRow}`).font = { bold: true };
  }

  await wb.xlsx.writeFile(filePath);
}

export const BUGGY_FILENAME = 'regional-pnl.xlsx';
export const DEMO_PROMPT =
  'Audit this workbook for formula bugs, data-quality issues, and structural problems. ' +
  'Focus on the P&L sheet. Report each issue with its cell reference, what is wrong, ' +
  'and how to fix it.';
