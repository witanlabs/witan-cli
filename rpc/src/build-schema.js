import { writeFileSync, mkdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));

/**
 * Helper to build a JSON Schema for a TypedDict-like object.
 * @param {string} description
 * @param {Record<string, any>} properties
 * @param {string[]} required
 * @param {string[]} [extendRefs]
 */
function obj(description, properties, required, extendRefs = []) {
  const base = {
    type: 'object',
    description,
    properties,
    required,
    additionalProperties: true,
  };
  if (extendRefs.length > 0) {
    return {
      allOf: [
        ...extendRefs.map((ref) => ({ $ref: `#/$defs/${ref}` })),
        base,
      ],
    };
  }
  return base;
}

/** Reference to another $def */
function ref(name) {
  return { $ref: `#/$defs/${name}` };
}

/** String enum */
function strEnum(values, description = '') {
  const s = { type: 'string', enum: values };
  if (description) s.description = description;
  return s;
}

/** Optional wrapper: T | null */
function nullable(schema) {
  return { anyOf: [schema, { type: 'null' }] };
}

/** Array of T */
function arr(schema, description = '') {
  const a = { type: 'array', items: schema };
  if (description) a.description = description;
  return a;
}

/** Record<string, T> */
function dict(valueSchema, description = '') {
  const d = { type: 'object', additionalProperties: valueSchema };
  if (description) d.description = description;
  return d;
}

/** Any value */
function any(description = '') {
  const a = {};
  if (description) a.description = description;
  return a;
}

// ============================================================================
// Schema definitions
// ============================================================================

const $defs = {};

// ---------------------------------------------------------------------------
// Primitive / literal aliases
// ---------------------------------------------------------------------------
$defs.Visibility = strEnum(['visible', 'outsidePrintArea', 'collapsed', 'hidden']);
$defs.SheetVisibility = strEnum(['visible', 'hidden', 'veryHidden']);
$defs.SetCellsValidationMode = strEnum(['ignore', 'reject']);
$defs.PasteType = strEnum(['all', 'values', 'formulas', 'formats']);
$defs.NumberType = strEnum(['currency', 'percent', 'fraction', 'exponential', 'date', 'text', 'number']);
$defs.ValueType = strEnum(['string', 'number', 'bool', 'date', 'error', 'blank']);
$defs.SweepMode = strEnum(['cartesian', 'parallel']);
$defs.LintSeverity = strEnum(['Info', 'Warning', 'Error']);
$defs.ReferenceType = strEnum(['direct', 'range', 'named', 'table']);
$defs.LinkType = strEnum(['internal', 'external']);

$defs.DataValidationOperator = strEnum([
  'Between', 'NotBetween', 'EqualTo', 'NotEqualTo',
  'GreaterThan', 'LessThan', 'GreaterThanOrEqualTo', 'LessThanOrEqualTo',
]);
$defs.DataValidationAlertStyle = strEnum(['Stop', 'Warning', 'Information']);
$defs.DataValidationStatus = strEnum(['Valid', 'Invalid', 'NoValidation', 'Mixed', 'Unknown']);

$defs.ChartType = strEnum([
  'column', 'bar', 'line', 'area', 'pie', 'doughnut',
  'scatter', 'bubble', 'radar', 'stockHLC', 'stockOHLC', 'waterfall',
]);
$defs.ChartGrouping = strEnum(['standard', 'stacked', 'percentStacked']);
$defs.ChartAxisBinding = strEnum(['primary', 'secondary']);
$defs.ChartStockRole = strEnum(['volume', 'open', 'high', 'low', 'close']);
$defs.ChartLegendPosition = strEnum(['left', 'right', 'top', 'bottom', 'topRight']);
$defs.ChartMarkerStyle = strEnum([
  'auto', 'none', 'circle', 'dash', 'diamond', 'dot',
  'picture', 'plus', 'square', 'star', 'triangle', 'x',
]);
$defs.ChartDataLabelPosition = strEnum([
  'bestFit', 'center', 'insideBase', 'insideEnd',
  'outsideEnd', 'left', 'right', 'top', 'bottom',
]);
$defs.ChartTimeUnit = strEnum(['days', 'months', 'years']);
$defs.ChartScatterStyle = strEnum(['line', 'lineMarker', 'marker', 'smooth', 'smoothMarker']);
$defs.ChartRadarStyle = strEnum(['standard', 'marker', 'filled']);
$defs.ChartCategoryAxisType = strEnum(['category', 'date']);

$defs.ThemeColorName = strEnum([
  'background1', 'light1', 'text1', 'dark1',
  'background2', 'light2', 'text2', 'dark2',
  'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
  'hyperlink', 'followedHyperlink',
]);

$defs.CfThresholdType = strEnum([
  'formula', 'max', 'min', 'num', 'percent', 'percentile', 'autoMin', 'autoMax',
]);
$defs.ConditionalFormattingRuleType = strEnum([
  'cellValue', 'containsText', 'notContainsText', 'beginsWith', 'endsWith',
  'containsBlanks', 'notContainsBlanks', 'containsErrors', 'notContainsErrors',
  'expression', 'timePeriod', 'top', 'bottom', 'aboveAverage', 'belowAverage',
  'duplicateValues', 'uniqueValues', 'twoColorScale', 'threeColorScale',
  'dataBar', 'iconSet',
]);
$defs.ConditionalFormattingOperator = strEnum([
  'equal', 'notEqual', 'greaterThan', 'greaterThanOrEqual',
  'lessThan', 'lessThanOrEqual', 'between', 'notBetween',
]);
$defs.ConditionalFormattingTimePeriod = strEnum([
  'today', 'yesterday', 'tomorrow', 'last7Days',
  'thisWeek', 'lastWeek', 'nextWeek', 'thisMonth', 'lastMonth', 'nextMonth',
]);

// ---------------------------------------------------------------------------
// Core data
// ---------------------------------------------------------------------------
$defs.CellCoordinates = obj('Cell coordinates', {
  sheet: { type: 'string' },
  row: { type: 'integer' },
  col: { anyOf: [{ type: 'integer' }, { type: 'string' }] },
}, ['sheet', 'row', 'col']);

$defs.Note = obj('Cell note', {
  author: { type: 'string' },
  text: { type: 'string' },
}, ['author', 'text']);

$defs.Link = obj('Hyperlink', {
  type: ref('LinkType'),
  target: { type: 'string' },
  tooltip: { type: 'string' },
}, ['type', 'target']);

$defs.ThreadComment = obj('Thread comment', {
  authorId: { type: 'string' },
  text: { type: 'string' },
  createdAt: { type: 'string' },
}, ['authorId', 'text', 'createdAt']);

$defs.ThreadInfo = obj('Thread info', {
  resolved: { type: 'boolean' },
  comments: arr(ref('ThreadComment')),
}, ['resolved', 'comments']);

$defs.Value = obj('Cell value returned by read operations', {
  address: { type: 'string' },
  sheet: { type: 'string' },
  row: { type: 'integer' },
  col: { type: 'integer' },
  colLetter: { type: 'string' },
  value: any('The raw cell value'),
  formula: { type: 'string' },
  type: ref('ValueType'),
  text: { type: 'string' },
  format: { type: 'string' },
  numberType: ref('NumberType'),
  visibility: ref('Visibility'),
  context: { type: 'string' },
  note: ref('Note'),
  link: ref('Link'),
  thread: ref('ThreadInfo'),
}, ['address', 'sheet', 'row', 'col', 'colLetter', 'value', 'type', 'text', 'visibility']);

$defs.Diagnostic = obj('Cell diagnostic / error', {
  code: { type: 'string' },
  address: { type: 'string' },
  detail: { type: 'string' },
  formula: { type: 'string' },
}, ['code', 'address']);

$defs.ViewRangeBounds = obj('Used range bounds', {
  startRow: { type: 'integer' },
  startCol: { type: 'integer' },
  endRow: { type: 'integer' },
  endCol: { type: 'integer' },
}, ['startRow', 'startCol', 'endRow', 'endCol']);

$defs.UpdatedSheetInfo = obj('Updated sheet info after mutation', {
  name: { type: 'string' },
  usedRange: { anyOf: [ref('ViewRangeBounds'), { type: 'null' }] },
  tileRowCount: { type: 'integer' },
  tileColCount: { type: 'integer' },
}, ['name', 'usedRange', 'tileRowCount', 'tileColCount']);

$defs.InvalidatedTile = obj('Invalidated tile', {
  sheet: { type: 'string' },
  tileRow: { type: 'integer' },
  tileCol: { type: 'integer' },
}, ['sheet', 'tileRow', 'tileCol']);

$defs.WriteResult = obj('Result of a write/mutation operation', {
  touched: dict({ type: 'string' }, 'Map of cell address -> stringified value'),
  changed: arr({ type: 'string' }),
  errors: arr(ref('Diagnostic')),
  invalidatedTiles: arr(ref('InvalidatedTile')),
  updatedSheets: arr(ref('UpdatedSheetInfo')),
}, ['touched', 'changed', 'errors', 'invalidatedTiles', 'updatedSheets']);

$defs.FindAndReplaceResult = obj('Find and replace result', {
  replaced: { type: 'integer' },
  cells: arr({ type: 'string' }),
  errors: arr(ref('Diagnostic')),
}, ['replaced', 'cells', 'errors']);

$defs.CopyRangeResult = obj('Copy range result', {
  destination: { type: 'string' },
  cellsCopied: { type: 'integer' },
}, ['destination', 'cellsCopied']);

$defs.AutoFitColumnResult = obj('Auto-fit column result', {
  width: { type: 'number' },
  previousWidth: { type: 'number' },
}, ['width', 'previousWidth']);

$defs.AutoFitRowResult = obj('Auto-fit row result', {
  height: { type: 'number' },
  previousHeight: { type: 'number' },
  hidden: { type: 'boolean' },
  previousHidden: { type: 'boolean' },
}, ['height', 'previousHeight', 'hidden', 'previousHidden']);

// ---------------------------------------------------------------------------
// Write helpers
// ---------------------------------------------------------------------------
$defs.CellWriteNote = obj('Note payload for cell writes', {
  text: { type: 'string' },
  author: { type: 'string' },
}, []);

$defs.CellWriteLink = obj('Link payload for cell writes', {
  url: { type: 'string' },
  ref: { type: 'string' },
  tooltip: { type: 'string' },
}, []);

$defs.CellWriteThreadComment = obj('Thread comment to add', {
  author: { type: 'string' },
  text: { type: 'string' },
}, ['text']);

$defs.CellWriteThread = obj('Thread payload for cell writes', {
  add: arr(ref('CellWriteThreadComment')),
  resolved: { type: 'boolean' },
  delete: { type: 'boolean' },
}, []);

$defs.CellWrite = obj('Cell write payload', {
  value: { anyOf: [{ type: 'string' }, { type: 'integer' }, { type: 'number' }, { type: 'boolean' }, { type: 'null' }] },
  formula: { type: 'string' },
  format: { type: 'string' },
  note: nullable(ref('CellWriteNote')),
  link: nullable(ref('CellWriteLink')),
  thread: nullable(ref('CellWriteThread')),
}, []);

$defs.CellAssignment = obj('Cell assignment (address + write payload)', {
  address: { type: 'string' },
  value: { anyOf: [{ type: 'string' }, { type: 'integer' }, { type: 'number' }, { type: 'boolean' }, { type: 'null' }] },
  formula: { type: 'string' },
  format: { type: 'string' },
  note: nullable(ref('CellWriteNote')),
  link: nullable(ref('CellWriteLink')),
  thread: nullable(ref('CellWriteThread')),
}, ['address']);

// ---------------------------------------------------------------------------
// Sheet & Workbook
// ---------------------------------------------------------------------------
$defs.SheetInfo = obj('Sheet metadata', {
  address: { type: 'string' },
  rows: { type: 'integer' },
  cols: { type: 'integer' },
  sheet: { type: 'string' },
  hidden: { type: 'boolean' },
  printArea: { type: 'string' },
  listObjects: arr({ type: 'string' }),
  dataTables: arr({ type: 'string' }),
  precedents: arr({ type: 'string' }),
  dependents: arr({ type: 'string' }),
}, ['address', 'rows', 'cols', 'sheet']);

$defs.FontOptions = obj('Font options', {
  name: nullable({ type: 'string' }),
  size: nullable({ type: 'number' }),
}, []);

$defs.WorkbookDefaultFont = obj('Workbook default font', {
  name: { type: 'string' },
  size: { type: 'number' },
}, ['name', 'size']);

$defs.ThemeColors = obj('Theme colors', {
  dark1: { type: 'string' },
  light1: { type: 'string' },
  dark2: { type: 'string' },
  light2: { type: 'string' },
  accent1: { type: 'string' },
  accent2: { type: 'string' },
  accent3: { type: 'string' },
  accent4: { type: 'string' },
  accent5: { type: 'string' },
  accent6: { type: 'string' },
  hyperlink: { type: 'string' },
  followedHyperlink: { type: 'string' },
  majorFont: nullable({ type: 'string' }),
  minorFont: nullable({ type: 'string' }),
}, ['dark1', 'light1', 'dark2', 'light2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hyperlink', 'followedHyperlink']);

$defs.WorkbookMetadata = obj('Workbook metadata', {
  author: nullable({ type: 'string' }),
  title: nullable({ type: 'string' }),
  subject: nullable({ type: 'string' }),
  company: nullable({ type: 'string' }),
  created: nullable({ type: 'string' }),
  modified: nullable({ type: 'string' }),
}, []);

$defs.IterativeCalculationSettings = obj('Iterative calculation settings', {
  enabled: { type: 'boolean' },
  maxIterations: { type: 'integer' },
  maxChange: { type: 'number' },
}, ['enabled', 'maxIterations', 'maxChange']);

$defs.WorkbookProperties = obj('Workbook properties', {
  activeSheetIndex: { type: 'integer' },
  defaultFont: ref('WorkbookDefaultFont'),
  metadata: nullable(ref('WorkbookMetadata')),
  themeColors: nullable(ref('ThemeColors')),
  iterativeCalculation: ref('IterativeCalculationSettings'),
}, ['activeSheetIndex', 'defaultFont', 'iterativeCalculation']);

$defs.WorkbookPropertiesUpdate = obj('Workbook properties update (partial)', {
  activeSheetIndex: { type: 'integer' },
  defaultFont: ref('FontOptions'),
  metadata: ref('WorkbookMetadata'),
  themeColors: ref('ThemeColors'),
  iterativeCalculation: ref('IterativeCalculationSettings'),
}, []);

$defs.SheetViewProperties = obj('Sheet view properties', {
  showGridLines: { type: 'boolean' },
  zoomScale: { type: 'integer' },
  freezeRows: { type: 'integer' },
  freezeColumns: { type: 'integer' },
}, ['showGridLines', 'zoomScale', 'freezeRows', 'freezeColumns']);

$defs.SheetOutlineProperties = obj('Sheet outline properties', {
  summaryRowsBelow: { type: 'boolean' },
  summaryColumnsRight: { type: 'boolean' },
  showSymbols: { type: 'boolean' },
}, ['summaryRowsBelow', 'summaryColumnsRight', 'showSymbols']);

$defs.SheetFormatProperties = obj('Sheet format properties', {
  defaultRowHeight: { type: 'number' },
  defaultColWidth: { type: 'number' },
  font: nullable(ref('FontOptions')),
}, ['defaultRowHeight', 'defaultColWidth', 'font']);

$defs.SheetViewPropertiesUpdate = obj('Sheet view properties update (partial)', {
  showGridLines: { type: 'boolean' },
  zoomScale: { type: 'integer' },
  freezeRows: { type: 'integer' },
  freezeColumns: { type: 'integer' },
}, []);

$defs.SheetOutlinePropertiesUpdate = obj('Sheet outline properties update (partial)', {
  summaryRowsBelow: { type: 'boolean' },
  summaryColumnsRight: { type: 'boolean' },
  showSymbols: { type: 'boolean' },
}, []);

$defs.SheetFormatPropertiesUpdate = obj('Sheet format properties update (partial)', {
  defaultRowHeight: { type: 'number' },
  defaultColWidth: { type: 'number' },
  font: nullable(ref('FontOptions')),
}, []);

$defs.RowProperties = obj('Row properties for updates', {
  height: { type: 'number' },
  hidden: { type: 'boolean' },
  outlineLevel: { type: 'integer' },
  collapsed: { type: 'boolean' },
}, []);

$defs.ColumnProperties = obj('Column properties for updates', {
  width: { type: 'number' },
  hidden: { type: 'boolean' },
  outlineLevel: { type: 'integer' },
  collapsed: { type: 'boolean' },
}, []);

$defs.RowDimension = obj('Row dimension (returned by getSheetProperties)', {
  row: { type: 'integer' },
  height: { type: 'number' },
  hidden: { type: 'boolean' },
  outlineLevel: { type: 'integer' },
  collapsed: { type: 'boolean' },
}, ['row']);

$defs.ColumnDimension = obj('Column dimension (returned by getSheetProperties)', {
  col: { type: 'string' },
  width: { type: 'number' },
  hidden: { type: 'boolean' },
  outlineLevel: { type: 'integer' },
  collapsed: { type: 'boolean' },
}, ['col']);

$defs.SheetProperties = obj('Sheet properties', {
  visibility: ref('SheetVisibility'),
  view: ref('SheetViewProperties'),
  outline: ref('SheetOutlineProperties'),
  format: ref('SheetFormatProperties'),
  columns: dict(ref('ColumnDimension')),
  rows: dict(ref('RowDimension')),
  merges: nullable(arr({ type: 'string' })),
}, ['visibility', 'view', 'outline', 'format', 'columns', 'rows']);

$defs.SheetPropertiesUpdate = obj('Sheet properties update (partial)', {
  visibility: ref('SheetVisibility'),
  view: ref('SheetViewPropertiesUpdate'),
  outline: ref('SheetOutlinePropertiesUpdate'),
  format: ref('SheetFormatPropertiesUpdate'),
  merges: arr({ type: 'string' }),
}, []);

// ---------------------------------------------------------------------------
// Styles
// ---------------------------------------------------------------------------
$defs.RichTextRun = obj('Rich text run', {
  text: { type: 'string' },
  style: ref('FontStyle'),
}, ['text']);

$defs.GradientFill = obj('Gradient fill', {
  type: { type: 'string' },
  color1: { type: 'string' },
  color2: { type: 'string' },
  degree: { type: 'number' },
  top: { type: 'number' },
  bottom: { type: 'number' },
  left: { type: 'number' },
  right: { type: 'number' },
}, ['type', 'color1', 'color2']);

$defs.FillStyle = obj('Fill style', {
  color: { type: 'string' },
  pattern: { type: 'string' },
  patternColor: { type: 'string' },
  gradient: ref('GradientFill'),
}, []);

$defs.FontStyle = obj('Font style', {
  name: { type: 'string' },
  size: { type: 'number' },
  color: { type: 'string' },
  bold: { type: 'boolean' },
  italic: { type: 'boolean' },
  strike: { type: 'boolean' },
  underline: { type: 'string' },
  verticalAlign: { type: 'string' },
}, []);

$defs.AlignmentStyle = obj('Alignment style', {
  horizontal: { type: 'string' },
  vertical: { type: 'string' },
  rotation: { type: 'integer' },
  wrapText: { type: 'boolean' },
  shrinkToFit: { type: 'boolean' },
  indent: { type: 'number' },
  readingOrder: strEnum(['context', 'leftToRight', 'rightToLeft']),
  autoIndent: { type: 'boolean' },
}, []);

$defs.ProtectionStyle = obj('Protection style', {
  locked: { type: 'boolean' },
  formulaHidden: { type: 'boolean' },
}, []);

$defs.BorderEdgeStyle = obj('Border edge style', {
  style: { type: 'string' },
  color: { type: 'string' },
  themeColor: ref('ThemeColorName'),
  tintAndShade: { type: 'number' },
}, []);

$defs.DiagonalBorderStyle = obj('Diagonal border style', {
  style: { type: 'string' },
  color: { type: 'string' },
  themeColor: ref('ThemeColorName'),
  tintAndShade: { type: 'number' },
  up: { type: 'boolean' },
  down: { type: 'boolean' },
}, []);

$defs.BorderStyle = obj('Border style', {
  top: ref('BorderEdgeStyle'),
  bottom: ref('BorderEdgeStyle'),
  left: ref('BorderEdgeStyle'),
  right: ref('BorderEdgeStyle'),
  diagonal: ref('DiagonalBorderStyle'),
}, []);

$defs.Style = obj('Cell style', {
  fill: nullable(ref('FillStyle')),
  font: nullable(ref('FontStyle')),
  alignment: nullable(ref('AlignmentStyle')),
  protection: nullable(ref('ProtectionStyle')),
  border: nullable(ref('BorderStyle')),
  numberFormat: nullable({ type: 'string' }),
  centerContinuousSpan: nullable({ type: 'integer' }),
  richText: nullable(arr(ref('RichTextRun'))),
}, []);

// ---------------------------------------------------------------------------
// List Objects & Data Tables
// ---------------------------------------------------------------------------
$defs.ListObjectColumn = obj('List object (table) column', {
  name: { type: 'string' },
  totalsRowFunction: nullable({ type: 'string' }),
  totalsRowLabel: nullable({ type: 'string' }),
  totalsRowFormula: nullable({ type: 'string' }),
  calculatedColumnFormula: nullable({ type: 'string' }),
}, ['name']);

$defs.ListObject = obj('List object (table)', {
  name: { type: 'string' },
  sheet: { type: 'string' },
  ref: { type: 'string' },
  showHeaderRow: { type: 'boolean' },
  showTotalsRow: { type: 'boolean' },
  showAutoFilter: { type: 'boolean' },
  tableStyleName: nullable({ type: 'string' }),
  showFirstColumn: { type: 'boolean' },
  showLastColumn: { type: 'boolean' },
  showRowStripes: { type: 'boolean' },
  showColumnStripes: { type: 'boolean' },
  headerRowRange: nullable({ type: 'string' }),
  dataRange: nullable({ type: 'string' }),
  totalsRowRange: nullable({ type: 'string' }),
  columns: arr(ref('ListObjectColumn')),
}, ['name', 'sheet', 'ref', 'showHeaderRow', 'showTotalsRow', 'showAutoFilter', 'tableStyleName', 'showFirstColumn', 'showLastColumn', 'showRowStripes', 'showColumnStripes', 'headerRowRange', 'dataRange', 'totalsRowRange', 'columns']);

$defs.ListObjectSpec = obj('List object spec for creation', {
  name: { type: 'string' },
  ref: { type: 'string' },
  showHeaderRow: { type: 'boolean' },
  showTotalsRow: { type: 'boolean' },
  showAutoFilter: { type: 'boolean' },
  tableStyleName: nullable({ type: 'string' }),
  showFirstColumn: { type: 'boolean' },
  showLastColumn: { type: 'boolean' },
  showRowStripes: { type: 'boolean' },
  showColumnStripes: { type: 'boolean' },
  columns: arr(ref('ListObjectColumn')),
  rows: nullable(arr(arr(ref('CellWrite')))),
}, ['name', 'ref', 'columns']);

$defs.ListObjectUpdate = obj('List object update (partial)', {
  ref: { type: 'string' },
  showHeaderRow: { type: 'boolean' },
  showTotalsRow: { type: 'boolean' },
  showAutoFilter: { type: 'boolean' },
  tableStyleName: nullable({ type: 'string' }),
  showFirstColumn: { type: 'boolean' },
  showLastColumn: { type: 'boolean' },
  showRowStripes: { type: 'boolean' },
  showColumnStripes: { type: 'boolean' },
  columns: arr(ref('ListObjectColumn')),
  rows: arr(arr(ref('CellWrite'))),
}, []);

$defs.ListObjectMutationResult = obj('List object mutation result', {
  listObject: ref('ListObject'),
}, ['listObject'], ['WriteResult']);

$defs.DataTableSourceFormula = obj('Data table source formula', {
  formula: { type: 'string' },
}, ['formula']);

$defs.OneVariableColumnDataTable = obj('One-variable column data table', {
  type: { const: 'oneVariableColumn' },
  sheet: { type: 'string' },
  ref: { type: 'string' },
  dataTableRange: { type: 'string' },
  rowInputCell: { type: 'null' },
  columnInputCell: { type: 'string' },
  inputValues: { type: 'null' },
  rowInputValues: { type: 'null' },
  columnInputValues: { type: 'null' },
  formulas: arr(ref('DataTableSourceFormula')),
  formula: { type: 'null' },
}, ['type', 'sheet', 'ref', 'dataTableRange', 'rowInputCell', 'columnInputCell', 'inputValues', 'rowInputValues', 'columnInputValues', 'formulas', 'formula']);

$defs.OneVariableRowDataTable = obj('One-variable row data table', {
  type: { const: 'oneVariableRow' },
  sheet: { type: 'string' },
  ref: { type: 'string' },
  dataTableRange: { type: 'string' },
  rowInputCell: { type: 'string' },
  columnInputCell: { type: 'null' },
  inputValues: { type: 'null' },
  rowInputValues: { type: 'null' },
  columnInputValues: { type: 'null' },
  formulas: arr(ref('DataTableSourceFormula')),
  formula: { type: 'null' },
}, ['type', 'sheet', 'ref', 'dataTableRange', 'rowInputCell', 'columnInputCell', 'inputValues', 'rowInputValues', 'columnInputValues', 'formulas', 'formula']);

$defs.TwoVariableDataTable = obj('Two-variable data table', {
  type: { const: 'twoVariable' },
  sheet: { type: 'string' },
  ref: { type: 'string' },
  dataTableRange: { type: 'string' },
  rowInputCell: { type: 'string' },
  columnInputCell: { type: 'string' },
  inputValues: { type: 'null' },
  rowInputValues: arr({ anyOf: [{ type: 'string' }, { type: 'integer' }, { type: 'number' }, { type: 'boolean' }, { type: 'null' }] }),
  columnInputValues: arr({ anyOf: [{ type: 'string' }, { type: 'integer' }, { type: 'number' }, { type: 'boolean' }, { type: 'null' }] }),
  formulas: { type: 'null' },
  formula: { type: 'string' },
}, ['type', 'sheet', 'ref', 'dataTableRange', 'rowInputCell', 'columnInputCell', 'inputValues', 'rowInputValues', 'columnInputValues', 'formulas', 'formula']);

$defs.DataTable = {
  anyOf: [ref('OneVariableColumnDataTable'), ref('OneVariableRowDataTable'), ref('TwoVariableDataTable')],
};

$defs.OneVariableColumnDataTableSpec = obj('One-variable column data table spec', {
  type: { const: 'oneVariableColumn' },
  ref: { type: 'string' },
  columnInputCell: { type: 'string' },
  inputValues: arr({ anyOf: [{ type: 'string' }, { type: 'integer' }, { type: 'number' }, { type: 'boolean' }, { type: 'null' }] }),
  formulas: arr({ type: 'string' }),
}, ['type', 'ref', 'columnInputCell', 'inputValues', 'formulas']);

$defs.OneVariableRowDataTableSpec = obj('One-variable row data table spec', {
  type: { const: 'oneVariableRow' },
  ref: { type: 'string' },
  rowInputCell: { type: 'string' },
  inputValues: arr({ anyOf: [{ type: 'string' }, { type: 'integer' }, { type: 'number' }, { type: 'boolean' }, { type: 'null' }] }),
  formulas: arr({ type: 'string' }),
}, ['type', 'ref', 'rowInputCell', 'inputValues', 'formulas']);

$defs.TwoVariableDataTableSpec = obj('Two-variable data table spec', {
  type: { const: 'twoVariable' },
  ref: { type: 'string' },
  rowInputCell: { type: 'string' },
  columnInputCell: { type: 'string' },
  rowInputValues: arr({ anyOf: [{ type: 'string' }, { type: 'integer' }, { type: 'number' }, { type: 'boolean' }, { type: 'null' }] }),
  columnInputValues: arr({ anyOf: [{ type: 'string' }, { type: 'integer' }, { type: 'number' }, { type: 'boolean' }, { type: 'null' }] }),
  formula: { type: 'string' },
}, ['type', 'ref', 'rowInputCell', 'columnInputCell', 'rowInputValues', 'columnInputValues', 'formula']);

$defs.DataTableSpec = {
  anyOf: [ref('OneVariableColumnDataTableSpec'), ref('OneVariableRowDataTableSpec'), ref('TwoVariableDataTableSpec')],
};

$defs.DataTableMutationResult = obj('Data table mutation result', {
  dataTable: ref('DataTable'),
}, ['dataTable'], ['WriteResult']);

// ---------------------------------------------------------------------------
// Charts
// ---------------------------------------------------------------------------
$defs.ChartTextSource = obj('Chart text source', {
  text: { type: 'string' },
  ref: { type: 'string' },
}, []);

$defs.ChartPositionAnchor = obj('Chart position anchor', {
  cell: { type: 'string' },
  xOffsetPts: { type: 'number' },
  yOffsetPts: { type: 'number' },
}, ['cell']);

$defs.ChartPositionInput = obj('Chart position input', {
  from: ref('ChartPositionAnchor'),
  to: ref('ChartPositionAnchor'),
}, ['from', 'to']);

$defs.ChartPosition = obj('Chart position (includes sheet)', {
  from: ref('ChartPositionAnchor'),
  to: ref('ChartPositionAnchor'),
  sheet: { type: 'string' },
}, ['from', 'to', 'sheet']);

$defs.ChartFillFormatSpec = obj('Chart fill format spec', {
  noFill: { type: 'boolean' },
  color: { type: 'string' },
}, []);

$defs.ChartBorderFormatSpec = obj('Chart border format spec', {
  noLine: { type: 'boolean' },
  color: { type: 'string' },
  weight: { type: 'number' },
  lineStyle: { type: 'string' },
}, []);

$defs.ChartLineFormatSpec = ref('ChartBorderFormatSpec');

$defs.ChartFontFormatSpec = obj('Chart font format spec', {
  bold: { type: 'boolean' },
  color: { type: 'string' },
  italic: { type: 'boolean' },
  name: { type: 'string' },
  size: { type: 'number' },
  underline: { type: 'string' },
}, []);

$defs.ChartDataLabelFormatSpec = obj('Chart data label format spec', {
  fill: ref('ChartFillFormatSpec'),
  border: ref('ChartBorderFormatSpec'),
  font: ref('ChartFontFormatSpec'),
}, []);

$defs.ChartAreaFormatSpec = ref('ChartDataLabelFormatSpec');

$defs.ChartPlotAreaFormatSpec = obj('Chart plot area format spec', {
  fill: ref('ChartFillFormatSpec'),
  border: ref('ChartBorderFormatSpec'),
}, []);

$defs.ChartPlotAreaSpec = obj('Chart plot area spec', {
  format: ref('ChartPlotAreaFormatSpec'),
}, []);

$defs.ChartSeriesFormatSpec = obj('Chart series format spec', {
  fill: ref('ChartFillFormatSpec'),
  line: ref('ChartLineFormatSpec'),
}, []);

$defs.ChartMarkerSpec = obj('Chart marker spec', {
  style: ref('ChartMarkerStyle'),
  size: { type: 'number' },
  fillColor: { type: 'string' },
  borderColor: { type: 'string' },
}, []);

$defs.ChartDataLabelsSpec = obj('Chart data labels spec', {
  showLegendKey: { type: 'boolean' },
  showValue: { type: 'boolean' },
  showCategory: { type: 'boolean' },
  showSeriesName: { type: 'boolean' },
  showPercent: { type: 'boolean' },
  showBubbleSize: { type: 'boolean' },
  showLeaderLines: { type: 'boolean' },
  position: ref('ChartDataLabelPosition'),
  numberFormat: { type: 'string' },
  numberFormatLinked: { type: 'boolean' },
  separator: { type: 'string' },
  format: ref('ChartDataLabelFormatSpec'),
}, []);

$defs.ChartSeriesSpec = obj('Chart series spec', {
  name: ref('ChartTextSource'),
  stockRole: ref('ChartStockRole'),
  categories: { type: 'string' },
  categoriesRefType: strEnum(['string', 'number', 'multiLevelString']),
  values: { type: 'string' },
  xValues: { type: 'string' },
  yValues: { type: 'string' },
  bubbleSizes: { type: 'string' },
  fillColor: { type: 'string' },
  lineColor: { type: 'string' },
  lineWidth: { type: 'number' },
  lineDashStyle: { type: 'string' },
  smooth: { type: 'boolean' },
  invertIfNegative: { type: 'boolean' },
  totalIndexes: arr({ type: 'integer' }),
  showConnectorLines: { type: 'boolean' },
  marker: ref('ChartMarkerSpec'),
  dataLabels: ref('ChartDataLabelsSpec'),
}, []);

$defs.ChartGroupSpec = obj('Chart group spec', {
  type: ref('ChartType'),
  scatterStyle: ref('ChartScatterStyle'),
  radarStyle: ref('ChartRadarStyle'),
  grouping: ref('ChartGrouping'),
  axis: ref('ChartAxisBinding'),
  gapWidth: { type: 'integer' },
  overlap: { type: 'integer' },
  varyColors: { type: 'boolean' },
  smooth: { type: 'boolean' },
  bubbleScale: { type: 'integer' },
  showNegativeBubbles: { type: 'boolean' },
  sizeRepresents: strEnum(['area', 'width']),
  firstSliceAngle: { type: 'integer' },
  holeSize: { type: 'integer' },
  dataLabels: ref('ChartDataLabelsSpec'),
  series: arr(ref('ChartSeriesSpec')),
}, ['type', 'series']);

$defs.ChartTitleSpec = obj('Chart title spec', {
  text: { type: 'string' },
  ref: { type: 'string' },
  overlay: { type: 'boolean' },
}, []);

$defs.ChartLegendSpec = obj('Chart legend spec', {
  visible: { type: 'boolean' },
  position: ref('ChartLegendPosition'),
  overlay: { type: 'boolean' },
}, []);

$defs.ChartAxisSpec = obj('Chart axis spec', {
  title: ref('ChartTextSource'),
  visible: { type: 'boolean' },
  categoryType: ref('ChartCategoryAxisType'),
  min: { type: 'number' },
  max: { type: 'number' },
  majorUnit: { type: 'number' },
  minorUnit: { type: 'number' },
  baseTimeUnit: ref('ChartTimeUnit'),
  majorTimeUnit: ref('ChartTimeUnit'),
  minorTimeUnit: ref('ChartTimeUnit'),
  numberFormat: { type: 'string' },
  numberFormatLinked: { type: 'boolean' },
  reversed: { type: 'boolean' },
  majorGridlines: { type: 'boolean' },
  minorGridlines: { type: 'boolean' },
  position: strEnum(['left', 'right', 'top', 'bottom']),
}, []);

$defs.ChartAxesSpec = obj('Chart axes spec', {
  category: ref('ChartAxisSpec'),
  value: ref('ChartAxisSpec'),
  secondaryCategory: ref('ChartAxisSpec'),
  secondaryValue: ref('ChartAxisSpec'),
}, []);

$defs.ChartSpec = obj('Chart spec', {
  id: { type: 'integer' },
  name: { type: 'string' },
  position: ref('ChartPositionInput'),
  groups: arr(ref('ChartGroupSpec')),
  title: ref('ChartTitleSpec'),
  legend: ref('ChartLegendSpec'),
  axes: ref('ChartAxesSpec'),
  format: ref('ChartAreaFormatSpec'),
  plotArea: ref('ChartPlotAreaSpec'),
  displayBlanksAs: strEnum(['gap', 'span', 'zero']),
  plotVisibleOnly: { type: 'boolean' },
  showDataLabelsOverMaximum: { type: 'boolean' },
  roundedCorners: { type: 'boolean' },
  styleId: { type: 'integer' },
}, ['name', 'position', 'groups']);

$defs.ChartInfo = obj('Chart info (returned by listCharts)', {
  id: { type: 'integer' },
  name: { type: 'string' },
  groups: arr(ref('ChartGroupSpec')),
  position: ref('ChartPosition'),
  title: ref('ChartTitleSpec'),
  legend: ref('ChartLegendSpec'),
  axes: ref('ChartAxesSpec'),
  format: ref('ChartAreaFormatSpec'),
  plotArea: ref('ChartPlotAreaSpec'),
  displayBlanksAs: strEnum(['gap', 'span', 'zero']),
  plotVisibleOnly: { type: 'boolean' },
  showDataLabelsOverMaximum: { type: 'boolean' },
  roundedCorners: { type: 'boolean' },
  styleId: { type: 'integer' },
}, ['name', 'groups', 'position']);

$defs.ChartGroupSummary = obj('Chart group summary', {
  type: { type: 'string' },
  axis: ref('ChartAxisBinding'),
  seriesCount: { type: 'integer' },
}, ['type', 'seriesCount']);

$defs.ChartSummary = obj('Chart summary', {
  id: { type: 'integer' },
  sheet: { type: 'string' },
  name: { type: 'string' },
  type: { type: 'string' },
  groups: arr(ref('ChartGroupSummary')),
  groupCount: { type: 'integer' },
  seriesCount: { type: 'integer' },
  position: ref('ChartPosition'),
}, ['sheet', 'name', 'type', 'groups', 'groupCount', 'seriesCount', 'position']);

// ---------------------------------------------------------------------------
// Conditional Formatting
// ---------------------------------------------------------------------------
$defs.CfColorScalePoint = obj('Conditional formatting color scale point', {
  type: ref('CfThresholdType'),
  value: { type: 'number' },
  formula: { type: 'string' },
  color: { type: 'string' },
}, ['type']);

$defs.CfDataBarThreshold = obj('Conditional formatting data bar threshold', {
  type: ref('CfThresholdType'),
  value: { type: 'number' },
  formula: { type: 'string' },
}, ['type']);

$defs.CfDataBarConfig = obj('Conditional formatting data bar config', {
  showValue: { type: 'boolean' },
  gradient: { type: 'boolean' },
  border: { type: 'boolean' },
  negativeBarColorSameAsPositive: { type: 'boolean' },
  negativeBarBorderColorSameAsPositive: { type: 'boolean' },
  axisPosition: strEnum(['automatic', 'middle', 'none']),
  direction: strEnum(['context', 'leftToRight', 'rightToLeft']),
  fillColor: { type: 'string' },
  borderColor: { type: 'string' },
  negativeFillColor: { type: 'string' },
  negativeBorderColor: { type: 'string' },
  axisColor: { type: 'string' },
  lowValue: ref('CfDataBarThreshold'),
  highValue: ref('CfDataBarThreshold'),
}, []);

$defs.ConditionalFormattingRule = obj('Conditional formatting rule', {
  address: { type: 'string' },
  type: ref('ConditionalFormattingRuleType'),
  index: { type: 'integer' },
  priority: { type: 'integer' },
  stopIfTrue: { type: 'boolean' },
  style: ref('Style'),
  operator: ref('ConditionalFormattingOperator'),
  formula: { type: 'string' },
  formula2: { type: 'string' },
  text: { type: 'string' },
  rank: { type: 'integer' },
  percent: { type: 'boolean' },
  equalAverage: { type: 'boolean' },
  stdDev: { type: 'integer' },
  lowValue: ref('CfColorScalePoint'),
  midValue: ref('CfColorScalePoint'),
  highValue: ref('CfColorScalePoint'),
  dataBar: ref('CfDataBarConfig'),
  iconSetStyle: { type: 'string' },
  timePeriod: ref('ConditionalFormattingTimePeriod'),
}, ['address', 'type']);

// ---------------------------------------------------------------------------
// Data Validation
// ---------------------------------------------------------------------------
$defs.BasicDataValidationRule = obj('Basic data validation rule', {
  operator: ref('DataValidationOperator'),
  formula1: { anyOf: [{ type: 'string' }, { type: 'integer' }, { type: 'number' }] },
  formula2: nullable({ anyOf: [{ type: 'string' }, { type: 'integer' }, { type: 'number' }] }),
}, ['operator', 'formula1']);

$defs.ListDataValidationRule = obj('List data validation rule', {
  source: { type: 'string' },
  inCellDropDown: { type: 'boolean' },
}, ['source']);

$defs.CustomDataValidationRule = obj('Custom data validation rule', {
  formula: { type: 'string' },
}, ['formula']);

$defs.DataValidationRulePayload = obj('Data validation rule payload', {
  wholeNumber: nullable(ref('BasicDataValidationRule')),
  decimal: nullable(ref('BasicDataValidationRule')),
  list: nullable(ref('ListDataValidationRule')),
  date: nullable(ref('BasicDataValidationRule')),
  time: nullable(ref('BasicDataValidationRule')),
  textLength: nullable(ref('BasicDataValidationRule')),
  custom: nullable(ref('CustomDataValidationRule')),
}, []);

$defs.DataValidationPrompt = obj('Data validation prompt', {
  showPrompt: { type: 'boolean' },
  title: nullable({ type: 'string' }),
  message: nullable({ type: 'string' }),
}, []);

$defs.DataValidationErrorAlert = obj('Data validation error alert', {
  showAlert: { type: 'boolean' },
  style: ref('DataValidationAlertStyle'),
  title: nullable({ type: 'string' }),
  message: nullable({ type: 'string' }),
}, []);

$defs.DataValidationSpec = obj('Data validation spec', {
  address: { type: 'string' },
  rule: ref('DataValidationRulePayload'),
  ignoreBlanks: { type: 'boolean' },
  prompt: nullable(ref('DataValidationPrompt')),
  errorAlert: nullable(ref('DataValidationErrorAlert')),
}, ['address', 'rule']);

$defs.DataValidationInfo = obj('Data validation info (returned by getDataValidations)', {
  address: { type: 'string' },
  rule: ref('DataValidationRulePayload'),
  index: { type: 'integer' },
  sheet: { type: 'string' },
  type: strEnum(['None', 'WholeNumber', 'Decimal', 'List', 'Date', 'Time', 'TextLength', 'Custom']),
  ignoreBlanks: { type: 'boolean' },
  prompt: ref('DataValidationPrompt'),
  errorAlert: ref('DataValidationErrorAlert'),
}, ['address', 'rule', 'index', 'sheet', 'type', 'ignoreBlanks', 'prompt', 'errorAlert']);

$defs.DataValidationDiagnostic = obj('Data validation diagnostic', {
  code: { type: 'string' },
  message: { type: 'string' },
  details: nullable(dict({ type: 'string' })),
}, ['code', 'message']);

$defs.DataValidationResult = obj('Data validation result', {
  status: ref('DataValidationStatus'),
  invalidCells: arr({ type: 'string' }),
  truncated: { type: 'boolean' },
  diagnostics: arr(ref('DataValidationDiagnostic')),
}, ['status', 'invalidCells', 'truncated', 'diagnostics']);

// ---------------------------------------------------------------------------
// Search & Formula & Dependency
// ---------------------------------------------------------------------------
$defs.SearchCell = obj('Search cell result', {
  type: { const: 'cell' },
  address: { type: 'string' },
  value: any('Cell value'),
  text: { type: 'string' },
  formula: { type: 'string' },
  row: { type: 'integer' },
  col: { type: 'integer' },
  colLetter: { type: 'string' },
  sheet: { type: 'string' },
  visibility: ref('Visibility'),
  context: { type: 'string' },
  role: { type: 'string' },
}, ['type', 'address', 'value', 'text', 'row', 'col', 'colLetter', 'sheet', 'visibility', 'role']);

$defs.SearchRow = obj('Search row result', {
  type: { const: 'row' },
  row: { type: 'integer' },
  sheet: { type: 'string' },
  matchedAt: { type: 'string' },
  range: { type: 'string' },
  tsv: { type: 'string' },
  visibility: ref('Visibility'),
  context: { type: 'string' },
}, ['type', 'row', 'sheet', 'matchedAt', 'range', 'tsv', 'visibility']);

$defs.TableLookupResult = obj('Table lookup result', {
  address: { type: 'string' },
  value: any('Cell value'),
  text: { type: 'string' },
  row: { type: 'integer' },
  col: { type: 'integer' },
  colLetter: { type: 'string' },
  sheet: { type: 'string' },
  visibility: ref('Visibility'),
  rowLabelFoundAt: { type: 'string' },
  rowLabelFound: { type: 'string' },
  columnLabelFoundAt: { type: 'string' },
  columnLabelFound: { type: 'string' },
}, ['address', 'value', 'text', 'row', 'col', 'colLetter', 'sheet', 'visibility', 'rowLabelFoundAt', 'rowLabelFound', 'columnLabelFoundAt', 'columnLabelFound']);

$defs.DependencyCell = obj('Dependency cell', {
  address: { type: 'string' },
  depth: { type: 'integer' },
  formula: { type: 'string' },
  referenceType: ref('ReferenceType'),
}, ['address', 'depth']);

$defs.DependencyResult = obj('Dependency result', {
  cells: arr(ref('DependencyCell')),
  warnings: arr(ref('Diagnostic')),
}, ['cells']);

$defs.TraceInput = obj('Trace input', {
  address: { type: 'string' },
  referenceCount: { type: 'integer' },
  text: { type: 'string' },
  nearbyLabel: { type: 'string' },
  context: { type: 'string' },
}, ['address', 'referenceCount']);

$defs.TraceOutput = obj('Trace output', {
  address: { type: 'string' },
  formula: { type: 'string' },
  text: { type: 'string' },
  visibility: ref('Visibility'),
  nearbyLabel: { type: 'string' },
  context: { type: 'string' },
}, ['address', 'visibility']);

$defs.FormulaResult = obj('Formula result', {
  formula: { type: 'string' },
  value: any('Formula value'),
  error: { type: 'object', additionalProperties: true },
}, ['formula', 'value']);

$defs.SweepInput = obj('Sweep input', {
  address: { type: 'string' },
  values: arr({ anyOf: [{ type: 'string' }, { type: 'integer' }, { type: 'number' }, { type: 'boolean' }, { type: 'null' }] }),
}, ['address', 'values']);

$defs.SweepEntry = obj('Sweep entry', {
  inputs: dict({ type: 'string' }),
  outputs: dict({ type: 'string' }),
  errors: arr(ref('Diagnostic')),
}, ['inputs', 'outputs', 'errors']);

$defs.OutputStats = obj('Output statistics', {
  min: { type: 'number' },
  max: { type: 'number' },
  mean: { type: 'number' },
  count: { type: 'integer' },
}, ['min', 'max', 'mean', 'count']);

$defs.SweepResult = obj('Sweep result', {
  tsv: { type: 'string' },
  sweeps: arr(ref('SweepEntry')),
  stats: dict(ref('OutputStats')),
  sweepCount: { type: 'integer' },
  inputCount: { type: 'integer' },
  outputCount: { type: 'integer' },
}, ['tsv', 'sweeps', 'sweepCount', 'inputCount', 'outputCount']);

// ---------------------------------------------------------------------------
// Defined Names, Lint, Misc
// ---------------------------------------------------------------------------
$defs.DefinedName = obj('Defined name', {
  name: { type: 'string' },
  range: { type: 'string' },
  scope: nullable({ type: 'string' }),
}, ['name', 'range', 'scope']);

$defs.LintDiagnostic = obj('Lint diagnostic', {
  severity: ref('LintSeverity'),
  ruleId: { type: 'string' },
  message: { type: 'string' },
  location: nullable({ type: 'string' }),
  visibility: nullable(ref('Visibility')),
}, ['severity', 'ruleId', 'message', 'location', 'visibility']);

$defs.LintResult = obj('Lint result', {
  diagnostics: arr(ref('LintDiagnostic')),
  total: { type: 'integer' },
}, ['diagnostics', 'total']);

$defs.DetectedTable = obj('Detected table', {
  address: { type: 'string' },
  headerRows: { type: 'string' },
  headerCols: nullable({ type: 'string' }),
  tableName: nullable({ type: 'string' }),
}, ['address', 'headerRows']);

$defs.SheetDescription = obj('Sheet description', {
  tables: dict(ref('DetectedTable')),
  structure: { type: 'string' },
}, ['tables', 'structure']);

// ---------------------------------------------------------------------------
// RPC Envelopes
// ---------------------------------------------------------------------------
$defs.RPCRequest = obj('RPC request envelope', {
  id: { type: 'string' },
  op: { type: 'string' },
  args: { type: 'object' },
}, ['id', 'op', 'args']);

$defs.RPCResponse = obj('RPC response envelope', {
  id: { type: 'string' },
  ok: { type: 'boolean' },
  result: {},
  code: { type: 'string' },
  message: { type: 'string' },
}, ['ok']);

// ---------------------------------------------------------------------------
// Scalar aliases (not generated as standalone types, but referenced)
// ---------------------------------------------------------------------------
$defs.ScalarCellValue = {
  anyOf: [{ type: 'string' }, { type: 'integer' }, { type: 'number' }, { type: 'boolean' }, { type: 'null' }],
  description: 'Scalar cell value',
};

// ============================================================================
// Assemble and write
// ============================================================================

const schema = {
  $schema: 'http://json-schema.org/draft-07/schema#',
  $id: 'witan-rpc',
  title: 'Witan RPC Types',
  description: 'JSON Schema for types used in the Witan NDJSON RPC protocol between SDKs and the Go CLI.',
  $defs,
};

mkdirSync(join(__dirname, '..', 'schemas'), { recursive: true });
writeFileSync(
  join(__dirname, '..', 'schemas', 'witan-rpc.json'),
  JSON.stringify(schema, null, 2),
);

console.log(`Wrote ${Object.keys($defs).length} type definitions to rpc/schemas/witan-rpc.json`);
