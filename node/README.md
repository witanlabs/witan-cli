# witan

Node.js SDK for the [Witan CLI](https://github.com/witanlabs/witan-cli) - read, write, and manipulate Excel files programmatically.

## Installation

```bash
npm install witan
```

This installs the SDK along with the appropriate platform-specific binary for your system.

## Quick Start

```typescript
import { Workbook } from 'witan';

// Open a workbook (uses await using for automatic cleanup)
{
  await using wb = await Workbook.open('report.xlsx');

  // List sheets
  const sheets = await wb.listSheets();
  console.log(sheets.map(s => s.sheet));

  // Read data
  const data = await wb.readRange('Sheet1!A1:D10');
  const tsv = await wb.readRangeTsv('Sheet1!A1:D10');

  // Write data
  await wb.setCells([
    { address: 'Sheet1!A1', value: 'Hello' },
    { address: 'Sheet1!B1', value: 42 },
  ]);

  // Save changes
  await wb.save();
} // Workbook automatically closed here
```

## Usage

### Opening Workbooks

```typescript
import { Workbook } from 'witan';

// Open existing file
const wb = await Workbook.open('report.xlsx');

// Create new file
const wb = await Workbook.open('new.xlsx', { create: true });

// With options
const wb = await Workbook.open('report.xlsx', {
  locale: 'en-US',
  requestTimeoutMs: 60_000, // 60 seconds
});
```

### Resource Management

The `Workbook` class implements `AsyncDisposable`, so you can use `await using` for automatic cleanup:

```typescript
// Recommended: await using (requires Node.js 22+ or TypeScript 5.2+)
{
  await using wb = await Workbook.open('report.xlsx');
  // ... work with workbook
} // Automatically closed

// Alternative: try/finally
const wb = await Workbook.open('report.xlsx');
try {
  // ... work with workbook
} finally {
  await wb.close();
}
```

### Reading Data

```typescript
// Read a single cell
const cell = await wb.readCell('Sheet1!A1');
console.log(cell.value, cell.formula);

// Read a range (returns 2D array)
const data = await wb.readRange('Sheet1!A1:C10');
for (const row of data) {
  console.log(row.map(cell => cell.value).join('\t'));
}

// Read as TSV string
const tsv = await wb.readRangeTsv('Sheet1!A1:C10');

// Read row/column
const row = await wb.readRow('Sheet1', 1);
const col = await wb.readColumn('Sheet1', 'A');
```

### Writing Data

```typescript
// Set cell values
await wb.setCells([
  { address: 'Sheet1!A1', value: 'Name' },
  { address: 'Sheet1!B1', value: 'Score' },
  { address: 'Sheet1!A2', value: 'Alice' },
  { address: 'Sheet1!B2', value: 95 },
]);

// Set with formula
await wb.setCells([
  { address: 'Sheet1!C2', value: null, formula: '=B2*1.1' },
]);

// Copy range
await wb.copyRange('Sheet1!A1:B2', 'Sheet1!D1');

// Scale numeric values
await wb.scaleRange('Sheet1!B:B', 1.1); // Increase by 10%

// Save changes
await wb.save();
```

### Searching

```typescript
// Find cells by value
const cells = await wb.findCells('revenue');

// Find with regex
const cells = await wb.findCells(/\d{3}-\d{4}/);

// Find in specific range
const cells = await wb.findCells('total', { in: 'Sheet1!A:Z' });

// Find rows
const rows = await wb.findRows('Alice');

// Find and replace
const result = await wb.findAndReplace('USD', 'EUR', {
  in: 'Prices!A1:D100',
});
console.log(`Replaced ${result.replaced} occurrences`);
```

### Sheet Operations

```typescript
// List sheets
const sheets = await wb.listSheets();
for (const sheet of sheets) {
  console.log(sheet.sheet, sheet.rows, sheet.cols);
}

// Add/rename/delete sheets
await wb.addSheet('NewSheet');
await wb.renameSheet('NewSheet', 'Data');
await wb.deleteSheet('OldSheet');

// Get/set sheet properties
const props = await wb.getSheetProperties('Sheet1');
await wb.setSheetProperties('Sheet1', { hidden: true });
```

### Row and Column Operations

```typescript
// Insert rows/columns
await wb.insertRowAfter('Sheet1', 5, 3);  // Insert 3 rows after row 5
await wb.insertColumnAfter('Sheet1', 'C', 2);  // Insert 2 columns after C

// Delete rows/columns
await wb.deleteRows('Sheet1', 5, 3);
await wb.deleteColumns('Sheet1', 'D', 2);

// Auto-fit
await wb.autoFitColumns('Sheet1');
await wb.autoFitRows('Sheet1');

// Sort
await wb.sortRange('Sheet1!A1:D100', [
  { column: 'B', descending: true },
]);
```

### Formula Operations

```typescript
// Evaluate formulas
const result = await wb.evaluateFormula('Sheet1', '=SUM(A1:A10)');
console.log(result.value);

// Trace dependencies
const precedents = await wb.getCellPrecedents('Sheet1!C5');
const dependents = await wb.getCellDependents('Sheet1!A1');

// Trace to inputs/outputs
const inputs = await wb.traceToInputs('Sheet1!D10');
const outputs = await wb.traceToOutputs('Sheet1!A1');
```

### Scenario Analysis

```typescript
// Sweep multiple input values
const result = await wb.sweepInputs(
  [
    { address: 'Sheet1!B1', values: [100, 200, 300] },
    { address: 'Sheet1!B2', values: [0.1, 0.2, 0.3] },
  ],
  ['Sheet1!C1', 'Sheet1!C2'],
  { mode: 'cartesian', includeStats: true }
);

console.log(result.tsv);
console.log(result.stats);
```

### Tables and Charts

```typescript
// List objects (tables)
const table = await wb.getListObject('SalesTable');
await wb.addListObject('Sheet1', {
  name: 'NewTable',
  ref: 'A1:D10',
  columns: [
    { name: 'Product' },
    { name: 'Region' },
    { name: 'Quarter' },
    { name: 'Sales' },
  ],
});

// Charts
const charts = await wb.listCharts();
const chart = await wb.getChart('Sheet1', 'Chart1');
await wb.addChart('Sheet1', {
  name: 'SalesChart',
  position: {
    from: { cell: 'F2' },
    to: { cell: 'N18' },
  },
  groups: [
    {
      type: 'bar',
      series: [
        {
          name: { text: 'Sales' },
          categories: 'Sheet1!A2:A10',
          values: 'Sheet1!B2:B10',
        },
      ],
    },
  ],
  title: { text: 'Sales' },
});
```

### Defined Names

```typescript
// List defined names
const names = await wb.listDefinedNames();

// Add/delete defined names
await wb.addDefinedName('TotalSales', 'Sheet1!B10');
await wb.deleteDefinedName('OldName');
```

### Styling

```typescript
// Set cell style
await wb.setStyle('Sheet1!A1', {
  font: { bold: true, size: 14 },
  fill: { color: '#FFFF00' },
});

// Get cell style
const style = await wb.getStyle('Sheet1!A1');

// Preview styles as image
const dataUrl = await wb.previewStyles('Sheet1!A1:D10');
```

### Linting

```typescript
// Run linting rules
const result = await wb.lint();
for (const diagnostic of result.diagnostics) {
  console.log(`${diagnostic.severity}: ${diagnostic.message} at ${diagnostic.location}`);
}

// Lint specific ranges
const result = await wb.lint({
  rangeAddresses: ['Sheet1!A1:Z100'],
  skipRuleIds: ['RULE_001'],
});
```

## CLI Usage

The package also provides a CLI:

```bash
# Via npx
npx witan xlsx calc report.xlsx
npx witan xlsx render report.xlsx "Summary!A1:F20"
npx witan xlsx lint report.xlsx

# If installed globally
npm install -g witan
witan xlsx calc report.xlsx
```

## Error Handling

```typescript
import {
  Workbook,
  WitanError,
  WitanProcessError,
  WitanRPCError,
  WitanTimeoutError
} from 'witan';

try {
  await using wb = await Workbook.open('report.xlsx');
  await wb.readRange('InvalidRange!!!');
} catch (err) {
  if (err instanceof WitanTimeoutError) {
    console.error('Request timed out:', err.message);
  } else if (err instanceof WitanRPCError) {
    console.error('RPC error:', err.code, err.message);
  } else if (err instanceof WitanProcessError) {
    console.error('Process error:', err.message);
    console.error('stderr:', err.stderrTail.join('\n'));
  } else if (err instanceof WitanError) {
    console.error('Witan error:', err.message);
  }
}
```

## Requirements

- Node.js 22.0.0 or later
- Supported platforms: macOS (arm64, x64), Linux (x64, arm64, musl), Windows (x64, arm64)

## License

Apache-2.0
