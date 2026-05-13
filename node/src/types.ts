// ============================================================================
// Type Aliases (matching Python)
// ============================================================================

export type JsonDict = Record<string, unknown>;
export type JsonMapping = Readonly<Record<string, unknown>>;
export type ScalarCellValue = string | number | boolean | null;
export type Visibility = 'visible' | 'outsidePrintArea' | 'collapsed' | 'hidden';
export type SheetVisibility = 'visible' | 'hidden' | 'veryHidden';
export type PasteType = 'all' | 'values' | 'formulas' | 'formats';
export type NumberType = 'currency' | 'percent' | 'fraction' | 'exponential' | 'date' | 'text' | 'number';
export type ValueType = 'string' | 'number' | 'bool' | 'date' | 'error' | 'blank';
export type SweepMode = 'cartesian' | 'parallel';
export type LintSeverity = 'Info' | 'Warning' | 'Error';
export type ReferenceType = 'direct' | 'range' | 'named' | 'table';
export type LinkType = 'internal' | 'external';

// Matcher types (Python distinguishes Matcher vs ReplaceMatcher)
export type Matcher = string | number | boolean | RegExp | (string | RegExp)[];
export type ReplaceMatcher = string | RegExp; // No array form for find_and_replace

// ============================================================================
// Option Interfaces (for method parameters)
// ============================================================================

export interface FindCellsOptions {
  /** Range to search within (e.g., "Sheet1!A:Z") */
  in?: string;
  /** Number of context rows/cols (default: 2) */
  context?: number;
  /** Maximum results (default: 20) */
  limit?: number;
  /** Skip first N results (default: 0) */
  offset?: number;
  /** Search in formulas instead of values */
  formulas?: boolean;
}

export interface FindRowsOptions {
  /** Range to search within */
  in?: string;
  /** Number of context rows */
  context?: number;
  /** Maximum results (default: 20) */
  limit?: number;
  /** Skip first N results (default: 0) */
  offset?: number;
}

export interface FindAndReplaceOptions {
  /** Range to search within */
  in?: string;
  /** Case-sensitive matching */
  matchCase?: boolean;
  /** Match entire cell content only */
  wholeCell?: boolean;
  /** Replace in formulas */
  inFormulas?: boolean;
  /** Maximum replacements */
  limit?: number;
}

export interface SweepOptions {
  /** Sweep mode: "cartesian" (all combinations) or "parallel" (zip) */
  mode?: SweepMode;
  /** Include min/max/mean statistics */
  includeStats?: boolean;
}

export interface LintOptions {
  /** Specific ranges to lint */
  rangeAddresses?: string[];
  /** Rule IDs to skip */
  skipRuleIds?: string[];
  /** Only run these rule IDs */
  onlyRuleIds?: string[];
}

export interface AutoFitColumnsOptions {
  /** Specific columns to fit (default: all) */
  columns?: (number | string)[];
  /** Minimum width */
  minWidth?: number;
  /** Maximum width */
  maxWidth?: number;
  /** Extra padding */
  padding?: number;
}

export interface AutoFitRowsOptions {
  /** Specific rows to fit (default: all) */
  rows?: number[];
  /** Minimum height */
  minHeight?: number;
  /** Maximum height */
  maxHeight?: number;
}

export interface SortKey {
  /** Column index or letter */
  column: number | string;
  /** Sort direction */
  descending?: boolean;
}

// ============================================================================
// Core Data Interfaces
// ============================================================================

export interface Note {
  author: string;
  text: string;
}

export interface Link {
  type: LinkType;
  target: string;
  tooltip?: string;
}

export interface ThreadComment {
  authorId: string;
  text: string;
  createdAt: string;
}

export interface ThreadInfo {
  resolved: boolean;
  comments: ThreadComment[];
}

export interface Value {
  address: string;
  sheet: string;
  row: number;
  col: number;
  colLetter: string;
  value: unknown;
  formula?: string;
  type: ValueType;
  text: string;
  format?: string;
  numberType?: NumberType;
  visibility: Visibility;
  context?: string;
  note?: Note;
  link?: Link;
  thread?: ThreadInfo;
}

export interface CellCoordinates {
  sheet: string;
  row: number;
  col: number | string;
}

export interface CellAssignment {
  address: string;
  value: ScalarCellValue;
  formula?: string;
}

// ============================================================================
// Sheet & Workbook Interfaces
// ============================================================================

export interface SheetInfo {
  address: string;
  rows: number;
  cols: number;
  sheet: string;
  hidden?: boolean;
  printArea?: string;
  listObjects?: string[];
  dataTables?: string[];
  precedents?: string[];
  dependents?: string[];
}

export interface UpdatedSheetInfo {
  name: string;
  usedRange: unknown;
  tileRowCount: number;
  tileColCount: number;
}

// Workbook/Sheet properties use JsonDict since schema is flexible
export type WorkbookProperties = JsonDict;
export type WorkbookPropertiesUpdate = JsonMapping;
export interface SheetProperties extends JsonDict {
  columns?: Record<string, unknown>;
  rows?: Record<string, unknown>;
  visibility?: SheetVisibility;
}
export type SheetPropertiesUpdate = JsonMapping;
export type RowProperties = JsonMapping;
export type ColumnProperties = JsonMapping;
export type Style = JsonMapping;

// ============================================================================
// Search Result Interfaces
// ============================================================================

export interface SearchCell {
  type: 'cell';
  address: string;
  value: unknown;
  text: string;
  formula?: string;
  row: number;
  col: number;
  colLetter: string;
  sheet: string;
  visibility: Visibility;
  context?: string;
  role: string;
}

export interface SearchRow {
  type: 'row';
  row: number;
  sheet: string;
  matchedAt: string;
  range: string;
  tsv: string;
  visibility: Visibility;
  context?: string;
}

export interface TableLookupResult {
  address: string;
  value: unknown;
  text: string;
  row: number;
  col: number;
  colLetter: string;
  sheet: string;
  visibility: Visibility;
  rowLabelFoundAt: string;
  rowLabelFound: string;
  columnLabelFoundAt: string;
  columnLabelFound: string;
}

// ============================================================================
// Write/Mutation Result Interfaces
// ============================================================================

export interface Diagnostic {
  code: string;
  address: string;
  detail?: string;
  formula?: string;
}

export interface WriteResult {
  touched: Record<string, string>;
  changed: string[];
  errors: Diagnostic[];
  invalidatedTiles: JsonDict[];
  updatedSheets: UpdatedSheetInfo[];
}

export interface FindAndReplaceResult {
  replaced: number;
  cells: string[];
  errors: Diagnostic[];
}

export interface CopyRangeResult {
  destination: string;
  cellsCopied: number;
}

export interface AutoFitColumnResult {
  width: number;
  previousWidth: number;
}

export interface AutoFitRowResult {
  height: number;
  previousHeight: number;
  hidden: boolean;
  previousHidden: boolean;
}

// ============================================================================
// Defined Names
// ============================================================================

export interface DefinedName {
  name: string;
  range: string;
  scope: string | null;
}

// ============================================================================
// List Objects (Tables) & Data Tables
// ============================================================================

export type ListObject = JsonDict;
export type ListObjectSpec = JsonMapping;
export type ListObjectUpdate = JsonMapping;

export interface ListObjectMutationResult extends WriteResult {
  listObject: JsonDict;
}

export type DataTable = JsonDict;
export type DataTableSpec = JsonMapping;

export interface DataTableMutationResult extends WriteResult {
  dataTable: JsonDict;
}

// ============================================================================
// Charts
// ============================================================================

export type ChartSpec = JsonMapping;
export type ChartInfo = JsonDict;
export type ChartSummary = JsonDict;

// ============================================================================
// Conditional Formatting
// ============================================================================

export type ConditionalFormattingRule = JsonMapping;

// ============================================================================
// Formula & Dependency Interfaces
// ============================================================================

export interface FormulaResult {
  formula: string;
  value: unknown;
  error?: JsonDict;
}

export interface DependencyCell {
  address: string;
  depth: number;
  formula?: string;
  referenceType?: ReferenceType;
}

export interface DependencyResult {
  cells: DependencyCell[];
  warnings?: Diagnostic[];
}

export interface TraceInput {
  address: string;
  referenceCount: number;
  text?: string;
  nearbyLabel?: string;
  context?: string;
}

export interface TraceOutput {
  address: string;
  formula?: string;
  text?: string;
  visibility: Visibility;
  nearbyLabel?: string;
  context?: string;
}

// ============================================================================
// Sweep/Scenario Interfaces
// ============================================================================

export interface SweepInput {
  address: string;
  values: ScalarCellValue[];
}

export interface SweepEntry {
  inputs: Record<string, string>;
  outputs: Record<string, string>;
  errors: Diagnostic[];
}

export interface OutputStats {
  min: number;
  max: number;
  mean: number;
  count: number;
}

export interface SweepResult {
  tsv: string;
  sweeps: SweepEntry[];
  stats?: Record<string, OutputStats>;
  sweepCount: number;
  inputCount: number;
  outputCount: number;
}

// ============================================================================
// Lint Interfaces
// ============================================================================

export interface LintDiagnostic {
  severity: LintSeverity;
  ruleId: string;
  message: string;
  location: string | null;
  visibility: Visibility | null;
}

export interface LintResult {
  diagnostics: LintDiagnostic[];
  total: number;
}

// ============================================================================
// Sheet Description (for describeSheet/describeSheets)
// ============================================================================

export type SheetDescription = JsonDict;
