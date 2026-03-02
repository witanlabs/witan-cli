import ExcelJS from 'exceljs';

/**
 * Generate a small quarterly-revenue workbook for demo purposes.
 *
 * Layout:
 *
 *              Q1       Q2       Q3       Q4      Total
 *  Widgets   120,000  145,000   98,000  175,000  =SUM(...)
 *  Gadgets    85,000   92,000  110,000   88,000  =SUM(...)
 *  Services   45,000   52,000   61,000   58,000  =SUM(...)
 *  Total     =SUM     =SUM     =SUM     =SUM     =SUM
 */
export async function createDemoWorkbook(filePath: string): Promise<void> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Revenue');

  // Headers
  ws.getRow(1).values = ['', 'Q1', 'Q2', 'Q3', 'Q4', 'Total'];
  ws.getRow(1).font = { bold: true };

  // Product rows
  const products: [string, number, number, number, number][] = [
    ['Widgets', 120000, 145000, 98000, 175000],
    ['Gadgets', 85000, 92000, 110000, 88000],
    ['Services', 45000, 52000, 61000, 58000],
  ];

  for (let i = 0; i < products.length; i++) {
    const row = i + 2;
    ws.getRow(row).values = products[i];
    ws.getCell(`F${row}`).value = { formula: `SUM(B${row}:E${row})` };
  }

  // Totals row
  const totalRow = products.length + 2;
  ws.getCell(`A${totalRow}`).value = 'Total';
  ws.getCell(`A${totalRow}`).font = { bold: true };
  for (const col of ['B', 'C', 'D', 'E', 'F']) {
    ws.getCell(`${col}${totalRow}`).value = {
      formula: `SUM(${col}2:${col}${totalRow - 1})`,
    };
  }

  // Number formatting
  for (let r = 2; r <= totalRow; r++) {
    for (let c = 2; c <= 6; c++) {
      ws.getRow(r).getCell(c).numFmt = '#,##0';
    }
  }

  // Column widths
  ws.getColumn(1).width = 12;
  for (let c = 2; c <= 6; c++) ws.getColumn(c).width = 14;

  await wb.xlsx.writeFile(filePath);
}

export const DEMO_FILENAME = 'quarterly-revenue.xlsx';
export const DEMO_QUESTION =
  'Which product line had the highest total annual revenue, and what was the amount?';
