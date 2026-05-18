#!/usr/bin/env npx tsx
/**
 * Demo of the Witan Node.js SDK.
 *
 * Usage:
 *   cd node && npx tsx examples/demo.ts
 */

import { Workbook } from '../src/index.js';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

async function main() {
  // Create a temp directory for our test file
  const tmpDir = await mkdtemp(join(tmpdir(), 'witan-demo-'));
  const testFile = join(tmpDir, 'test.xlsx');

  try {
    console.log('=== Creating new workbook ===');
    {
      await using wb = await Workbook.open(testFile, { create: true });

      // List sheets
      const sheets = await wb.listSheets();
      console.log('Initial sheets:', sheets.map(s => s.sheet));

      // Write some data
      console.log('\n=== Writing data ===');
      await wb.setCells([
        { address: 'Sheet1!A1', value: 'Name' },
        { address: 'Sheet1!B1', value: 'Age' },
        { address: 'Sheet1!C1', value: 'Score' },
        { address: 'Sheet1!A2', value: 'Alice' },
        { address: 'Sheet1!B2', value: 30 },
        { address: 'Sheet1!C2', value: 95.5 },
        { address: 'Sheet1!A3', value: 'Bob' },
        { address: 'Sheet1!B3', value: 25 },
        { address: 'Sheet1!C3', value: 87.0 },
        { address: 'Sheet1!A4', value: 'Charlie' },
        { address: 'Sheet1!B4', value: 35 },
        { address: 'Sheet1!C4', value: 92.0 },
      ]);
      console.log('Data written');

      // Add a formula
      await wb.setCells([
        { address: 'Sheet1!C5', value: null, formula: '=AVERAGE(C2:C4)' },
        { address: 'Sheet1!A5', value: 'Average:' },
      ]);
      console.log('Formula added');

      // Save
      const saved = await wb.save();
      console.log('Saved:', saved);
    }

    console.log('\n=== Reopening workbook ===');
    {
      await using wb = await Workbook.open(testFile);

      // Read range as TSV
      console.log('\nData as TSV:');
      const tsv = await wb.readRangeTsv('Sheet1!A1:C5');
      console.log(tsv);

      // Read specific cell
      const avgCell = await wb.readCell('Sheet1!C5');
      console.log('\nAverage cell:');
      console.log('  Value:', avgCell.value);
      console.log('  Formula:', avgCell.formula);
      console.log('  Text:', avgCell.text);

      // Find cells
      console.log('\n=== Search operations ===');
      const found = await wb.findCells('Alice');
      console.log('Found "Alice":', found.map(c => c.address));

      // Find with regex
      const numbers = await wb.findCells(/^\d+$/);
      console.log('Found numbers:', numbers.map(c => `${c.address}=${c.value}`));

      // Evaluate formula
      console.log('\n=== Formula evaluation ===');
      const result = await wb.evaluateFormula('Sheet1', '=SUM(B2:B4)');
      console.log('SUM(B2:B4) =', result.value);

      // Get cell precedents
      console.log('\n=== Dependency tracing ===');
      const precedents = await wb.getCellPrecedents('Sheet1!C5');
      console.log('C5 depends on:', precedents.cells.map(c => c.address));

      // Add a new sheet
      console.log('\n=== Sheet operations ===');
      await wb.addSheet('Summary');
      const sheets = await wb.listSheets();
      console.log('Sheets after adding:', sheets.map(s => s.sheet));

      // Copy data to new sheet
      await wb.copyRange('Sheet1!A1:C1', 'Summary!A1');

      // Read from new sheet
      const summaryTsv = await wb.readRangeTsv('Summary!A1:C1');
      console.log('Summary sheet header:', summaryTsv.trim());

      // Describe sheet
      console.log('\n=== Sheet description ===');
      const desc = await wb.describeSheet('Sheet1');
      console.log('Sheet1 description:', JSON.stringify(desc, null, 2).slice(0, 500) + '...');

      // Save changes
      await wb.save();
    }

    console.log('\n=== Done! ===');

  } finally {
    // Cleanup
    await rm(tmpDir, { recursive: true });
    console.log('\nCleaned up temp directory');
  }
}

main().catch(err => {
  console.error('Error:', err);
  process.exit(1);
});
