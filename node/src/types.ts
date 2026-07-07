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

export interface CellWrite {
  value?: ScalarCellValue;
  formula?: string;
  format?: string;
  note?: { text?: string; author?: string } | null;
  link?: { url?: string; ref?: string; tooltip?: string } | null;
  thread?: {
    add?: Array<{ author?: string; text: string }>;
    resolved?: boolean;
    delete?: boolean;
  } | null;
}

export interface CellAssignment extends CellWrite {
  address: string;
}

export type SetCellsValidationMode = 'ignore' | 'reject';

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
  usedRange: {
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
  } | null;
  tileRowCount: number;
  tileColCount: number;
}

export interface FontOptions {
  name?: string | null;
  size?: number | null;
}

export interface WorkbookDefaultFont {
  name: string;
  size: number;
}

export interface ThemeColors {
  dark1: string;
  light1: string;
  dark2: string;
  light2: string;
  accent1: string;
  accent2: string;
  accent3: string;
  accent4: string;
  accent5: string;
  accent6: string;
  hyperlink: string;
  followedHyperlink: string;
  majorFont?: string | null;
  minorFont?: string | null;
}

export interface WorkbookMetadata {
  author?: string | null;
  title?: string | null;
  subject?: string | null;
  company?: string | null;
  created?: string | null;
  modified?: string | null;
}

export interface IterativeCalculationSettings {
  enabled: boolean;
  maxIterations: number;
  maxChange: number;
}

export interface WorkbookProperties {
  activeSheetIndex: number;
  defaultFont: WorkbookDefaultFont;
  metadata?: WorkbookMetadata | null;
  themeColors?: ThemeColors | null;
  iterativeCalculation: IterativeCalculationSettings;
}

export interface WorkbookPropertiesUpdate {
  activeSheetIndex?: number;
  defaultFont?: FontOptions;
  metadata?: Pick<WorkbookMetadata, 'author' | 'title' | 'subject' | 'company'>;
  themeColors?: Partial<ThemeColors>;
  iterativeCalculation?: Partial<IterativeCalculationSettings>;
}

export interface SheetViewProperties {
  showGridLines: boolean;
  zoomScale: number;
  freezeRows: number;
  freezeColumns: number;
}

export interface SheetOutlineProperties {
  summaryRowsBelow: boolean;
  summaryColumnsRight: boolean;
  showSymbols: boolean;
}

export interface SheetFormatProperties {
  defaultRowHeight: number;
  defaultColWidth: number;
  font?: FontOptions | null;
}

export interface RowProperties {
  height?: number;
  hidden?: boolean;
  outlineLevel?: number;
  collapsed?: boolean;
}

export interface ColumnProperties {
  width?: number;
  hidden?: boolean;
  outlineLevel?: number;
  collapsed?: boolean;
}

export interface RowDimension {
  row: number;
  height: number;
  hidden?: boolean;
  outlineLevel?: number;
  collapsed?: boolean;
}

export interface ColumnDimension {
  col: string;
  width: number;
  hidden?: boolean;
  outlineLevel?: number;
  collapsed?: boolean;
}

export interface SheetProperties {
  visibility: SheetVisibility;
  view: SheetViewProperties;
  outline: SheetOutlineProperties;
  format: SheetFormatProperties;
  columns: Record<string, ColumnDimension>;
  rows: Record<number, RowDimension>;
  merges?: string[] | null;
}

export interface SheetPropertiesUpdate {
  visibility?: SheetVisibility;
  view?: Partial<SheetViewProperties>;
  outline?: Partial<SheetOutlineProperties>;
  format?: Partial<SheetFormatProperties>;
  merges?: string[];
}

export interface RichTextRun {
  text: string;
  style?: FontStyle;
}

export type ThemeColorName =
  | 'background1'
  | 'light1'
  | 'text1'
  | 'dark1'
  | 'background2'
  | 'light2'
  | 'text2'
  | 'dark2'
  | 'accent1'
  | 'accent2'
  | 'accent3'
  | 'accent4'
  | 'accent5'
  | 'accent6'
  | 'hyperlink'
  | 'followedHyperlink';

export interface GradientFill {
  type: string;
  degree?: number;
  color1: string;
  color2: string;
  top?: number;
  bottom?: number;
  left?: number;
  right?: number;
}

export interface FillStyle {
  color?: string;
  pattern?: string;
  patternColor?: string;
  gradient?: GradientFill;
}

export interface FontStyle {
  name?: string;
  size?: number;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  strike?: boolean;
  underline?: string;
  verticalAlign?: string;
}

export interface AlignmentStyle {
  horizontal?: string;
  vertical?: string;
  rotation?: number;
  wrapText?: boolean;
  shrinkToFit?: boolean;
  indent?: number;
  readingOrder?: 'context' | 'leftToRight' | 'rightToLeft';
  autoIndent?: boolean;
}

export interface ProtectionStyle {
  locked?: boolean;
  formulaHidden?: boolean;
}

export interface BorderEdgeStyle {
  style?: string;
  color?: string;
  themeColor?: ThemeColorName;
  tintAndShade?: number;
}

export interface DiagonalBorderStyle extends BorderEdgeStyle {
  up?: boolean;
  down?: boolean;
}

export interface BorderStyle {
  top?: BorderEdgeStyle;
  bottom?: BorderEdgeStyle;
  left?: BorderEdgeStyle;
  right?: BorderEdgeStyle;
  diagonal?: DiagonalBorderStyle;
}

export interface Style {
  fill?: FillStyle | null;
  font?: FontStyle | null;
  alignment?: AlignmentStyle | null;
  protection?: ProtectionStyle | null;
  border?: BorderStyle | null;
  numberFormat?: string | null;
  centerContinuousSpan?: number | null;
  richText?: RichTextRun[] | null;
}

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

export interface InvalidatedTile {
  sheet: string;
  tileRow: number;
  tileCol: number;
}

export interface WriteResult {
  touched: Record<string, string>;
  changed: string[];
  errors: Diagnostic[];
  invalidatedTiles: InvalidatedTile[];
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

export interface ListObjectColumn {
  name: string;
  totalsRowFunction?: string | null;
  totalsRowLabel?: string | null;
  totalsRowFormula?: string | null;
  calculatedColumnFormula?: string | null;
}

export interface ListObject {
  name: string;
  sheet: string;
  ref: string;
  showHeaderRow: boolean;
  showTotalsRow: boolean;
  showAutoFilter: boolean;
  tableStyleName: string | null;
  showFirstColumn: boolean;
  showLastColumn: boolean;
  showRowStripes: boolean;
  showColumnStripes: boolean;
  headerRowRange: string | null;
  dataRange: string | null;
  totalsRowRange: string | null;
  columns: ListObjectColumn[];
}

export interface ListObjectSpec {
  name: string;
  ref: string;
  showHeaderRow?: boolean;
  showTotalsRow?: boolean;
  showAutoFilter?: boolean;
  tableStyleName?: string | null;
  showFirstColumn?: boolean;
  showLastColumn?: boolean;
  showRowStripes?: boolean;
  showColumnStripes?: boolean;
  columns: ListObjectColumn[];
  rows?: CellWrite[][];
}

export interface ListObjectUpdate {
  ref?: string;
  showHeaderRow?: boolean;
  showTotalsRow?: boolean;
  showAutoFilter?: boolean;
  tableStyleName?: string | null;
  showFirstColumn?: boolean;
  showLastColumn?: boolean;
  showRowStripes?: boolean;
  showColumnStripes?: boolean;
  columns?: ListObjectColumn[];
  rows?: CellWrite[][];
}

export interface ListObjectMutationResult extends WriteResult {
  listObject: ListObject;
}

export interface DataTableSourceFormula {
  formula: string;
}

export type DataTable =
  | {
      type: 'oneVariableColumn';
      sheet: string;
      ref: string;
      dataTableRange: string;
      rowInputCell: null;
      columnInputCell: string;
      inputValues: ScalarCellValue[];
      rowInputValues: null;
      columnInputValues: null;
      formulas: DataTableSourceFormula[];
      formula: null;
    }
  | {
      type: 'oneVariableRow';
      sheet: string;
      ref: string;
      dataTableRange: string;
      rowInputCell: string;
      columnInputCell: null;
      inputValues: ScalarCellValue[];
      rowInputValues: null;
      columnInputValues: null;
      formulas: DataTableSourceFormula[];
      formula: null;
    }
  | {
      type: 'twoVariable';
      sheet: string;
      ref: string;
      dataTableRange: string;
      rowInputCell: string;
      columnInputCell: string;
      inputValues: null;
      rowInputValues: ScalarCellValue[];
      columnInputValues: ScalarCellValue[];
      formulas: null;
      formula: string;
    };

export type DataTableSpec =
  | {
      type: 'oneVariableColumn';
      ref: string;
      columnInputCell: string;
      inputValues: ScalarCellValue[];
      formulas: string[];
    }
  | {
      type: 'oneVariableRow';
      ref: string;
      rowInputCell: string;
      inputValues: ScalarCellValue[];
      formulas: string[];
    }
  | {
      type: 'twoVariable';
      ref: string;
      rowInputCell: string;
      columnInputCell: string;
      rowInputValues: ScalarCellValue[];
      columnInputValues: ScalarCellValue[];
      formula: string;
    };

export interface DataTableMutationResult extends WriteResult {
  dataTable: DataTable;
}

// ============================================================================
// Charts
// ============================================================================

export type ChartType =
  | 'column'
  | 'bar'
  | 'line'
  | 'area'
  | 'pie'
  | 'doughnut'
  | 'scatter'
  | 'bubble'
  | 'radar'
  | 'surface'
  | 'stockHLC'
  | 'stockOHLC'
  | 'waterfall'
  | 'histogram'
  | 'pareto'
  | 'funnel'
  | 'boxWhisker';
export type ChartGrouping = 'standard' | 'stacked' | 'percentStacked';
export type ChartAxisBinding = 'primary' | 'secondary';
export type ChartStockRole = 'volume' | 'open' | 'high' | 'low' | 'close';
export type ChartBinType = 'auto' | 'binCount' | 'binWidth' | 'category';
export type ChartBoxWhiskerQuartileCalculation = 'exclusive' | 'inclusive';
export type ChartLegendPosition = 'left' | 'right' | 'top' | 'bottom' | 'topRight';
export type ChartMarkerStyle =
  | 'auto'
  | 'none'
  | 'circle'
  | 'dash'
  | 'diamond'
  | 'dot'
  | 'picture'
  | 'plus'
  | 'square'
  | 'star'
  | 'triangle'
  | 'x';
export type ChartDataLabelPosition =
  | 'bestFit'
  | 'center'
  | 'insideBase'
  | 'insideEnd'
  | 'outsideEnd'
  | 'left'
  | 'right'
  | 'top'
  | 'bottom';
export type ChartCategoryAxisType = 'category' | 'date';
export type ChartTimeUnit = 'days' | 'months' | 'years';
export type ChartScatterStyle = 'line' | 'lineMarker' | 'marker' | 'smooth' | 'smoothMarker';
export type ChartRadarStyle = 'standard' | 'marker' | 'filled';
export type ChartSurfaceVariant = 'topView' | 'topViewWireframe';
export type ChartSelector = string | number;
export type ChartPreviewFormat = 'png' | 'webp';

export interface ChartPreviewOptions {
  format?: ChartPreviewFormat;
  dpr?: number;
  zoom?: number;
}

export interface ChartTextSource {
  text?: string;
  ref?: string;
}

export interface ChartPositionAnchor {
  cell: string;
  xOffsetPts?: number;
  yOffsetPts?: number;
}

export interface ChartPositionInput {
  from: ChartPositionAnchor;
  to: ChartPositionAnchor;
}

export interface ChartPosition extends ChartPositionInput {
  sheet: string;
}

export interface ChartFillFormatSpec {
  noFill?: boolean;
  color?: string;
}

export interface ChartBorderFormatSpec {
  noLine?: boolean;
  color?: string;
  weight?: number;
  lineStyle?: string;
}

export type ChartLineFormatSpec = ChartBorderFormatSpec;

export interface ChartFontFormatSpec {
  bold?: boolean;
  color?: string;
  italic?: boolean;
  name?: string;
  size?: number;
  underline?: string;
}

export interface ChartDataLabelFormatSpec {
  fill?: ChartFillFormatSpec;
  border?: ChartBorderFormatSpec;
  font?: ChartFontFormatSpec;
}

export interface ChartAreaFormatSpec extends ChartDataLabelFormatSpec {}

export interface ChartPlotAreaFormatSpec {
  fill?: ChartFillFormatSpec;
  border?: ChartBorderFormatSpec;
}

export interface ChartPlotAreaSpec {
  format?: ChartPlotAreaFormatSpec;
}

export interface ChartSeriesFormatSpec {
  fill?: ChartFillFormatSpec;
  line?: ChartLineFormatSpec;
}

export interface ChartMarkerSpec {
  style?: ChartMarkerStyle;
  size?: number;
  fillColor?: string;
  borderColor?: string;
}

export interface ChartDataLabelsSpec {
  showLegendKey?: boolean;
  showValue?: boolean;
  showCategory?: boolean;
  showSeriesName?: boolean;
  showPercent?: boolean;
  showBubbleSize?: boolean;
  showLeaderLines?: boolean;
  position?: ChartDataLabelPosition;
  numberFormat?: string;
  numberFormatLinked?: boolean;
  separator?: string;
  format?: ChartDataLabelFormatSpec;
}

export interface ChartBinOptionsSpec {
  type?: ChartBinType;
  count?: number;
  width?: number;
  allowOverflow?: boolean;
  overflowValue?: number;
  allowUnderflow?: boolean;
  underflowValue?: number;
}

export interface ChartSeriesSpec {
  name?: ChartTextSource;
  stockRole?: ChartStockRole;
  categories?: string;
  categoriesRefType?: 'string' | 'number' | 'multiLevelString';
  values?: string;
  xValues?: string;
  yValues?: string;
  bubbleSizes?: string;
  fillColor?: string;
  lineColor?: string;
  lineWidth?: number;
  lineDashStyle?: string;
  smooth?: boolean;
  invertIfNegative?: boolean;
  totalIndexes?: number[];
  showConnectorLines?: boolean;
  binOptions?: ChartBinOptionsSpec;
  quartileCalculation?: ChartBoxWhiskerQuartileCalculation;
  showInnerPoints?: boolean;
  showMeanLine?: boolean;
  showMeanMarker?: boolean;
  showOutlierPoints?: boolean;
  marker?: ChartMarkerSpec;
  dataLabels?: ChartDataLabelsSpec;
}

export interface ChartGroupSpec {
  type: ChartType;
  scatterStyle?: ChartScatterStyle;
  radarStyle?: ChartRadarStyle;
  surfaceVariant?: ChartSurfaceVariant;
  grouping?: ChartGrouping;
  axis?: ChartAxisBinding;
  gapWidth?: number;
  overlap?: number;
  varyColors?: boolean;
  smooth?: boolean;
  bubbleScale?: number;
  showNegativeBubbles?: boolean;
  sizeRepresents?: 'area' | 'width';
  firstSliceAngle?: number;
  holeSize?: number;
  dataLabels?: ChartDataLabelsSpec;
  series: ChartSeriesSpec[];
}

export interface ChartTitleSpec extends ChartTextSource {
  overlay?: boolean;
}

export interface ChartLegendSpec {
  visible?: boolean;
  position?: ChartLegendPosition;
  overlay?: boolean;
}

export interface ChartAxisSpec {
  title?: ChartTextSource;
  visible?: boolean;
  categoryType?: ChartCategoryAxisType;
  min?: number;
  max?: number;
  majorUnit?: number;
  minorUnit?: number;
  baseTimeUnit?: ChartTimeUnit;
  majorTimeUnit?: ChartTimeUnit;
  minorTimeUnit?: ChartTimeUnit;
  numberFormat?: string;
  numberFormatLinked?: boolean;
  reversed?: boolean;
  majorGridlines?: boolean;
  minorGridlines?: boolean;
  position?: 'left' | 'right' | 'top' | 'bottom';
}

export interface ChartAxesSpec {
  category?: ChartAxisSpec;
  value?: ChartAxisSpec;
  secondaryCategory?: ChartAxisSpec;
  secondaryValue?: ChartAxisSpec;
}

export interface ChartSpec {
  id?: number;
  name: string;
  position: ChartPositionInput;
  groups: ChartGroupSpec[];
  title?: ChartTitleSpec;
  legend?: ChartLegendSpec;
  axes?: ChartAxesSpec;
  format?: ChartAreaFormatSpec;
  plotArea?: ChartPlotAreaSpec;
  displayBlanksAs?: 'gap' | 'span' | 'zero';
  plotVisibleOnly?: boolean;
  showDataLabelsOverMaximum?: boolean;
  roundedCorners?: boolean;
  styleId?: number;
}

export interface ChartInfo extends Omit<ChartSpec, 'position'> {
  position: ChartPosition;
}

export interface ChartGroupSummary {
  type: string;
  axis?: ChartAxisBinding;
  seriesCount: number;
}

export interface ChartSummary {
  id?: number;
  sheet: string;
  name: string;
  type: string;
  groups: ChartGroupSummary[];
  groupCount: number;
  seriesCount: number;
  position: ChartPosition;
}

// ============================================================================
// Images
// ============================================================================

export type ImageFormat = 'png' | 'jpeg';

export interface ImagePositionAnchor {
  cell: string;
  xOffsetPts?: number;
  yOffsetPts?: number;
}

export interface ImagePositionInput {
  from: ImagePositionAnchor;
  to: ImagePositionAnchor;
}

export interface ImagePosition extends ImagePositionInput {
  sheet?: string;
}

export interface ImageSource {
  base64: string;
}

export interface ImageSpec {
  name: string;
  position: ImagePositionInput;
  source: ImageSource;
  format?: ImageFormat;
  altText?: string | null;
  altTextTitle?: string | null;
  preserveAspectRatio?: boolean;
}

export interface ImageUpdate {
  name?: string;
  position?: ImagePositionInput;
  source?: ImageSource;
  format?: ImageFormat;
  altText?: string | null;
  altTextTitle?: string | null;
  preserveAspectRatio?: boolean;
}

export interface ImageInfo {
  id?: number;
  sheet: string;
  name: string;
  position: ImagePosition;
  format?: ImageFormat;
  widthPts?: number;
  heightPts?: number;
  naturalWidthPx?: number;
  naturalHeightPx?: number;
  altText?: string | null;
  altTextTitle?: string | null;
}

export interface ImageSelector {
  name?: string;
  id?: number;
}

// ============================================================================
// Conditional Formatting
// ============================================================================

export type CfThresholdType =
  | 'formula'
  | 'max'
  | 'min'
  | 'num'
  | 'percent'
  | 'percentile'
  | 'autoMin'
  | 'autoMax';

export interface CfColorScalePoint {
  type: CfThresholdType;
  value?: number;
  formula?: string;
  color?: string;
}

export interface CfDataBarThreshold {
  type: CfThresholdType;
  value?: number;
  formula?: string;
}

export interface CfDataBarConfig {
  showValue?: boolean;
  gradient?: boolean;
  border?: boolean;
  negativeBarColorSameAsPositive?: boolean;
  negativeBarBorderColorSameAsPositive?: boolean;
  axisPosition?: 'automatic' | 'middle' | 'none';
  direction?: 'context' | 'leftToRight' | 'rightToLeft';
  fillColor?: string;
  borderColor?: string;
  negativeFillColor?: string;
  negativeBorderColor?: string;
  axisColor?: string;
  lowValue?: CfDataBarThreshold;
  highValue?: CfDataBarThreshold;
}

export type ConditionalFormattingRuleType =
  | 'cellValue'
  | 'containsText'
  | 'notContainsText'
  | 'beginsWith'
  | 'endsWith'
  | 'containsBlanks'
  | 'notContainsBlanks'
  | 'containsErrors'
  | 'notContainsErrors'
  | 'expression'
  | 'timePeriod'
  | 'top'
  | 'bottom'
  | 'aboveAverage'
  | 'belowAverage'
  | 'duplicateValues'
  | 'uniqueValues'
  | 'twoColorScale'
  | 'threeColorScale'
  | 'dataBar'
  | 'iconSet';

export type ConditionalFormattingOperator =
  | 'equal'
  | 'notEqual'
  | 'greaterThan'
  | 'greaterThanOrEqual'
  | 'lessThan'
  | 'lessThanOrEqual'
  | 'between'
  | 'notBetween';

export type ConditionalFormattingTimePeriod =
  | 'today'
  | 'yesterday'
  | 'tomorrow'
  | 'last7Days'
  | 'thisWeek'
  | 'lastWeek'
  | 'nextWeek'
  | 'thisMonth'
  | 'lastMonth'
  | 'nextMonth';

export interface ConditionalFormattingRule {
  address: string;
  type: ConditionalFormattingRuleType;
  index?: number;
  priority?: number;
  stopIfTrue?: boolean;
  style?: Style;
  operator?: ConditionalFormattingOperator;
  formula?: string;
  formula2?: string;
  text?: string;
  rank?: number;
  percent?: boolean;
  equalAverage?: boolean;
  stdDev?: number;
  lowValue?: CfColorScalePoint;
  midValue?: CfColorScalePoint;
  highValue?: CfColorScalePoint;
  dataBar?: CfDataBarConfig;
  iconSetStyle?: string;
  timePeriod?: ConditionalFormattingTimePeriod;
}

// ============================================================================
// Data Validation
// ============================================================================

export type DataValidationOperator =
  | 'Between'
  | 'NotBetween'
  | 'EqualTo'
  | 'NotEqualTo'
  | 'GreaterThan'
  | 'LessThan'
  | 'GreaterThanOrEqualTo'
  | 'LessThanOrEqualTo';

export type DataValidationAlertStyle = 'Stop' | 'Warning' | 'Information';

export interface BasicDataValidationRule {
  operator: DataValidationOperator;
  formula1: string | number;
  formula2?: string | number | null;
}

export interface ListDataValidationRule {
  source: string;
  inCellDropDown?: boolean;
}

export interface CustomDataValidationRule {
  formula: string;
}

export interface DataValidationRulePayload {
  wholeNumber?: BasicDataValidationRule | null;
  decimal?: BasicDataValidationRule | null;
  list?: ListDataValidationRule | null;
  date?: BasicDataValidationRule | null;
  time?: BasicDataValidationRule | null;
  textLength?: BasicDataValidationRule | null;
  custom?: CustomDataValidationRule | null;
}

export interface DataValidationPrompt {
  showPrompt?: boolean;
  title?: string | null;
  message?: string | null;
}

export interface DataValidationErrorAlert {
  showAlert?: boolean;
  style?: DataValidationAlertStyle;
  title?: string | null;
  message?: string | null;
}

export interface DataValidationSpec {
  address: string;
  rule: DataValidationRulePayload;
  ignoreBlanks?: boolean;
  prompt?: DataValidationPrompt | null;
  errorAlert?: DataValidationErrorAlert | null;
}

export interface DataValidationInfo {
  address: string;
  rule: DataValidationRulePayload;
  index: number;
  sheet: string;
  type:
    | 'None'
    | 'WholeNumber'
    | 'Decimal'
    | 'List'
    | 'Date'
    | 'Time'
    | 'TextLength'
    | 'Custom';
  ignoreBlanks: boolean;
  prompt: Required<DataValidationPrompt>;
  errorAlert: Required<DataValidationErrorAlert>;
}

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

/** References the floating object (chart, image, shape) a diagnostic is about. */
export interface LintDiagnosticObject {
  /** Known kinds: "chart" | "image" | "shape" | "object". */
  kind: string;
  name: string | null;
}

export interface LintDiagnostic {
  severity: LintSeverity;
  ruleId: string;
  message: string;
  location: string | null;
  visibility: Visibility | null;
  /** Present only for diagnostics about a floating object (e.g. chart rules D100-D110). */
  object?: LintDiagnosticObject | null;
}

export interface LintResult {
  diagnostics: LintDiagnostic[];
  total: number;
}

// ============================================================================
// Sheet Description (for describeSheet/describeSheets)
// ============================================================================

export interface DetectedTable {
  address: string;
  headerRows: string;
  headerCols?: string | null;
  tableName?: string | null;
}

export interface SheetDescription {
  tables: Record<string, DetectedTable>;
  structure: string;
}
