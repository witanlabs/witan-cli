#!/usr/bin/env npx tsx
/**
 * Demo of the Witan Node.js SDK for Google Sheets.
 *
 * Prerequisites (one-time CLI setup):
 *   witan auth login
 *   witan gsheets connect
 *
 * Usage:
 *   cd node && npx tsx examples/google-sheets-demo.ts gs://YOUR_SHEET_REF
 *   cd node && npx tsx examples/google-sheets-demo.ts https://docs.google.com/spreadsheets/d/ID/edit
 *   cd node && npx tsx examples/google-sheets-demo.ts --create --title "Witan demo"
 */

import { GoogleSheet, isGoogleAuthRequired, WitanRPCError } from '../src/index.js';

function parseArgs(): { ref?: string; create: boolean; title: string } {
  const args = process.argv.slice(2);
  let create = false;
  let title = 'Witan SDK demo';
  const positional: string[] = [];

  for (let i = 0; i < args.length; i++) {
    const arg = args[i]!;
    if (arg === '--create') {
      create = true;
      continue;
    }
    if (arg === '--title') {
      title = args[++i] ?? title;
      continue;
    }
    if (arg.startsWith('--title=')) {
      title = arg.slice('--title='.length);
      continue;
    }
    positional.push(arg);
  }

  return { ref: positional[0], create, title };
}

async function openSheet(args: ReturnType<typeof parseArgs>): Promise<GoogleSheet> {
  if (args.create) {
    return GoogleSheet.create(args.title);
  }
  if (!args.ref) {
    console.error('error: ref is required unless --create is set');
    process.exit(2);
  }
  return GoogleSheet.open(args.ref);
}

async function main() {
  const args = parseArgs();

  try {
    await using sheet = await openSheet(args);

    if (sheet.isCreate) {
      console.log(`=== Created spreadsheet: ${JSON.stringify(args.title)} ===`);
    } else {
      console.log(`=== Opened spreadsheet: ${args.ref} ===`);
    }

    const sheets = await sheet.listSheets();
    console.log(`Sheets: ${JSON.stringify(sheets.map((s) => s.sheet))}`);

    console.log('\n=== Writing sample data ===');
    await sheet.setCells([
      { address: 'Sheet1!A1', value: 'Name' },
      { address: 'Sheet1!B1', value: 'Score' },
      { address: 'Sheet1!A2', value: 'Alice' },
      { address: 'Sheet1!B2', value: 95 },
      { address: 'Sheet1!A3', value: 'Bob' },
      { address: 'Sheet1!B3', value: 87 },
      { address: 'Sheet1!A4', value: 'Average:' },
      { address: 'Sheet1!B4', value: null, formula: '=AVERAGE(B2:B3)' },
    ]);
    console.log('Changes applied (no save() — Google Sheets persists immediately)');

    console.log('\n=== Reading data ===');
    console.log(await sheet.readRangeTsv('Sheet1!A1:B4'));

    const avg = await sheet.readCell('Sheet1!B4');
    console.log(`\nAverage cell: value=${JSON.stringify(avg.value)} formula=${JSON.stringify(avg.formula ?? null)}`);

    console.log('\n=== Search ===');
    const found = await sheet.findCells('Alice');
    console.log(`Found Alice at: ${JSON.stringify(found.map((c) => c.address))}`);

    const numbers = await sheet.findCells(/^\d+$/);
    const numStrs = numbers.map((c) => `${c.address}=${c.value}`);
    console.log(`Numeric cells: ${JSON.stringify(numStrs)}`);

    console.log('\n=== Formula evaluation ===');
    console.log('Skipped: evaluateFormulas is not implemented for Google Sheets');
    console.log(`Read formula result from B4 instead: ${JSON.stringify(avg.value)}`);

    console.log('\n=== Sheet description (truncated) ===');
    const desc = await sheet.describeSheet('Sheet1');
    console.log(`${JSON.stringify(desc, null, 2).slice(0, 500)}...`);

    console.log('\n=== Done ===');
  } catch (err) {
    if (err instanceof WitanRPCError) {
      if (isGoogleAuthRequired(err)) {
        console.error(
          'Google authentication required. Run:\n' +
            '  witan auth login\n' +
            '  witan gsheets connect'
        );
      } else {
        console.error(`RPC error (${err.code}): ${err.message}`);
      }
      process.exit(1);
    }
    throw err;
  }
}

main().catch((err) => {
  console.error('Error:', err);
  process.exit(1);
});
