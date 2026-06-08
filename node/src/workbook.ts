import { getBinaryPath } from './binary.js';
import { WitanProcessError } from './errors.js';
import { dropUndefined, serializeMatcher } from './helpers.js';
import { StdioRPCProcess } from './process.js';
import type {
  AutoFitColumnResult,
  AutoFitRowResult,
  CellAssignment,
  ChartInfo,
  ChartSpec,
  ChartSummary,
  ConditionalFormattingRule,
  CopyRangeResult,
  DataValidationInfo,
  DataValidationResult,
  DataValidationSpec,
  DataTable,
  DataTableMutationResult,
  DataTableSpec,
  DefinedName,
  DependencyResult,
  FindAndReplaceResult,
  FormulaResult,
  ImageInfo,
  ImageSelector,
  ImageSpec,
  ImageUpdate,
  LintResult,
  ListObject,
  ListObjectMutationResult,
  ListObjectSpec,
  ListObjectUpdate,
  Matcher,
  PasteType,
  ReplaceMatcher,
  SearchCell,
  SearchRow,
  SetCellsValidationMode,
  SheetInfo,
  SheetProperties,
  SheetPropertiesUpdate,
  RowProperties,
  ColumnProperties,
  SortKey,
  Style,
  SheetDescription,
  SweepInput,
  SweepMode,
  SweepResult,
  TableLookupResult,
  TraceInput,
  TraceOutput,
  Value,
  WorkbookProperties,
  WorkbookPropertiesUpdate,
  WriteResult,
} from './types.js';

/**
 * Options for opening a Workbook.
 */
export interface WorkbookOptions {
  /** Create the file if it doesn't exist */
  create?: boolean;
  /** Locale for number/date formatting */
  locale?: string;
  /** Hint for file format detection */
  hint?: string;
  /** Enable stateless mode */
  stateless?: boolean;
  /** API key for authentication */
  apiKey?: string;
  /** API URL override */
  apiUrl?: string;
  /** Custom binary path (overrides auto-detection) */
  binary?: string;
  /** Additional environment variables for the subprocess */
  env?: Record<string, string>;
  /** Request timeout in milliseconds (default: 90000). */
  requestTimeoutMs?: number;
}

/**
 * Async Workbook session backed by `witan xlsx rpc`.
 *
 * Use the static `Workbook.open()` factory to create instances.
 * Implements `AsyncDisposable` for use with `await using`.
 *
 * @example
 * ```typescript
 * // With await using (recommended)
 * {
 *   await using wb = await Workbook.open('report.xlsx');
 *   const sheets = await wb.listSheets();
 *   await wb.save();
 * } // Automatically closed
 *
 * // With try/finally
 * const wb = await Workbook.open('report.xlsx');
 * try {
 *   const sheets = await wb.listSheets();
 * } finally {
 *   await wb.close();
 * }
 * ```
 */
export class Workbook implements AsyncDisposable {
  private process: StdioRPCProcess;
  private requestId = 0;
  private closed = false;

  /**
   * Private constructor - use Workbook.open() instead.
   */
  private constructor(process: StdioRPCProcess) {
    this.process = process;
  }

  /**
   * Open a workbook and start the RPC subprocess.
   *
   * @param path - Path to the Excel file
   * @param options - Configuration options
   * @returns A ready-to-use Workbook instance
   * @throws {WitanProcessError} If the subprocess fails to start
   */
  static async open(path: string, options: WorkbookOptions = {}): Promise<Workbook> {
    const argv = Workbook.buildArgv(path, options);
    const process = new StdioRPCProcess(argv, {
      env: options.env,
      timeoutMs: options.requestTimeoutMs,
    });

    try {
      await process.waitReady();
    } catch (err) {
      process.terminate();
      throw err;
    }

    return new Workbook(process);
  }

  /**
   * Build command-line arguments for the witan subprocess.
   */
  private static buildArgv(path: string, options: WorkbookOptions): string[] {
    const binary = options.binary ?? getBinaryPath();
    const argv = [binary];

    if (options.apiKey !== undefined) {
      argv.push('--api-key', options.apiKey);
    }
    if (options.apiUrl !== undefined) {
      argv.push('--api-url', options.apiUrl);
    }
    if (options.stateless === true) {
      argv.push('--stateless');
    }

    argv.push('xlsx', 'rpc', path);

    if (options.create === true) {
      argv.push('--create');
    }
    if (options.hint !== undefined) {
      argv.push('--hint', options.hint);
    }
    if (options.locale !== undefined) {
      argv.push('--locale', options.locale);
    }

    return argv;
  }

  /**
   * Generate the next request ID.
   */
  private nextId(): string {
    return String(++this.requestId);
  }

  /**
   * Send an RPC request.
   */
  private async request(method: string, op: string, args: Record<string, unknown> = {}): Promise<unknown> {
    if (this.closed) {
      throw new WitanProcessError('Workbook is closed');
    }
    return this.process.request(method, op, args, this.nextId());
  }

  // ============================================================================
  // Lifecycle
  // ============================================================================

  /**
   * Close the workbook and terminate the subprocess.
   * Safe to call multiple times.
   */
  async close(): Promise<void> {
    if (this.closed) return;
    this.closed = true;
    await this.process.close();
  }

  /**
   * Implements AsyncDisposable for `await using` syntax.
   */
  async [Symbol.asyncDispose](): Promise<void> {
    await this.close();
  }

  /**
   * Whether the workbook has been closed.
   */
  get isClosed(): boolean {
    return this.closed;
  }

  // ============================================================================
  // Save
  // ============================================================================

  /**
   * Save the workbook to disk.
   * @returns true if the save was successful
   */
  async save(): Promise<boolean> {
    return (await this.request('save', 'save', {})) as boolean;
  }

  // ============================================================================
  // Workbook Properties
  // ============================================================================

  /**
   * Get workbook-level properties.
   */
  async getWorkbookProperties(): Promise<WorkbookProperties> {
    return (await this.request('getWorkbookProperties', 'getWorkbookProperties', {})) as WorkbookProperties;
  }

  /**
   * Set workbook-level properties.
   */
  async setWorkbookProperties(properties: WorkbookPropertiesUpdate): Promise<void> {
    await this.request('setWorkbookProperties', 'setWorkbookProperties', properties as Record<string, unknown>);
  }

  // ============================================================================
  // Sheet Operations
  // ============================================================================

  /**
   * List all sheets in the workbook.
   */
  async listSheets(): Promise<SheetInfo[]> {
    const result = (await this.request('listSheets', 'listSheets', {})) as { sheets?: SheetInfo[] };
    return result.sheets ?? [];
  }

  /**
   * Add a new sheet to the workbook.
   * @param name - Name for the new sheet
   * @returns The name of the created sheet
   */
  async addSheet(name: string): Promise<string> {
    await this.request('addSheet', 'addSheet', { name });
    return name;
  }

  /**
   * Delete a sheet from the workbook.
   * @param name - Name of the sheet to delete
   */
  async deleteSheet(name: string): Promise<void> {
    await this.request('deleteSheet', 'deleteSheet', { name });
  }

  /**
   * Rename a sheet.
   * @param oldName - Current name of the sheet
   * @param newName - New name for the sheet
   */
  async renameSheet(oldName: string, newName: string): Promise<void> {
    await this.request('renameSheet', 'renameSheet', { oldName, newName });
  }

  // ============================================================================
  // Sheet Properties
  // ============================================================================

  /**
   * Get properties of a sheet.
   * @param sheet - Sheet name
   * @param options - Filter options for columns/rows
   */
  async getSheetProperties(
    sheet: string,
    options: { columns?: (number | string)[]; rows?: number[] } = {}
  ): Promise<SheetProperties> {
    const filter = dropUndefined({
      columns: options.columns,
      rows: options.rows,
    });

    const args: Record<string, unknown> = { sheet };
    if (Object.keys(filter).length > 0) {
      args['filter'] = filter;
    }

    const result = (await this.request('getSheetProperties', 'getSheetProperties', args)) as SheetProperties;

    // Set defaults (matches Python's setdefault behavior)
    result.columns ??= {};
    result.rows ??= {};

    return result;
  }

  /**
   * Set properties of a sheet.
   * @param sheet - Sheet name
   * @param properties - Properties to set
   */
  async setSheetProperties(sheet: string, properties: SheetPropertiesUpdate): Promise<void> {
    await this.request('setSheetProperties', 'setSheetProperties', {
      sheet,
      properties,
    });
  }

  // ============================================================================
  // Reading Data
  // ============================================================================

  /**
   * Read a single cell value.
   * This is a composite operation that calls readRange and extracts the first cell.
   *
   * @param cell - Cell address (e.g., "Sheet1!A1" or "A1")
   * @param options - Read options
   * @returns The cell value
   */
  async readCell(cell: string, options: { context?: number } = {}): Promise<Value> {
    const data = (await this.request(
      'readCell',
      'readRange',
      dropUndefined({ address: cell, context: options.context })
    )) as Value[][];
    return data[0]![0]!;
  }

  /**
   * Read a range of cells.
   *
   * @param range - Range address (e.g., "Sheet1!A1:B10")
   * @returns 2D array of cell values
   */
  async readRange(range: string): Promise<Value[][]> {
    return (await this.request('readRange', 'readRange', { address: range })) as Value[][];
  }

  /**
   * Read a row of cells.
   *
   * @param sheet - Sheet name
   * @param row - Row number (1-based)
   * @param options - Read options
   * @returns Array of cell values
   */
  async readRow(
    sheet: string,
    row: number,
    options: { startCol?: number; endCol?: number } = {}
  ): Promise<Value[]> {
    return (await this.request(
      'readRow',
      'readRow',
      dropUndefined({
        sheet,
        row,
        startCol: options.startCol,
        endCol: options.endCol,
      })
    )) as Value[];
  }

  /**
   * Read a column of cells.
   *
   * @param sheet - Sheet name
   * @param col - Column number (1-based) or letter (e.g., "A")
   * @param options - Read options
   * @returns Array of cell values
   */
  async readColumn(
    sheet: string,
    col: number | string,
    options: { startRow?: number; endRow?: number } = {}
  ): Promise<Value[]> {
    return (await this.request(
      'readColumn',
      'readColumn',
      dropUndefined({
        sheet,
        col,
        startRow: options.startRow,
        endRow: options.endRow,
      })
    )) as Value[];
  }

  /**
   * Read a range as tab-separated values.
   *
   * @param range - Range address
   * @param options - Read options
   * @returns TSV string
   */
  async readRangeTsv(
    range: string,
    options: { includeEmpty?: boolean; includeFormulas?: boolean } = {}
  ): Promise<string> {
    return (await this.request(
      'readRangeTsv',
      'readRangeTsv',
      dropUndefined({
        address: range,
        includeEmpty: options.includeEmpty,
        includeFormulas: options.includeFormulas,
      })
    )) as string;
  }

  /**
   * Read a row as tab-separated values.
   *
   * @param sheet - Sheet name
   * @param row - Row number (1-based)
   * @param options - Read options
   * @returns TSV string
   */
  async readRowTsv(
    sheet: string,
    row: number,
    options: {
      startCol?: number;
      endCol?: number;
      includeEmpty?: boolean;
      includeFormulas?: boolean;
    } = {}
  ): Promise<string> {
    return (await this.request(
      'readRowTsv',
      'readRowTsv',
      dropUndefined({
        sheet,
        row,
        startCol: options.startCol,
        endCol: options.endCol,
        includeEmpty: options.includeEmpty,
        includeFormulas: options.includeFormulas,
      })
    )) as string;
  }

  /**
   * Read a column as tab-separated values.
   *
   * @param sheet - Sheet name
   * @param col - Column number (1-based) or letter
   * @param options - Read options
   * @returns TSV string
   */
  async readColumnTsv(
    sheet: string,
    col: number | string,
    options: {
      startRow?: number;
      endRow?: number;
      includeEmpty?: boolean;
      includeFormulas?: boolean;
    } = {}
  ): Promise<string> {
    return (await this.request(
      'readColumnTsv',
      'readColumnTsv',
      dropUndefined({
        sheet,
        col,
        startRow: options.startRow,
        endRow: options.endRow,
        includeEmpty: options.includeEmpty,
        includeFormulas: options.includeFormulas,
      })
    )) as string;
  }

  // ============================================================================
  // Writing Data
  // ============================================================================

  /**
   * Set values in multiple cells.
   *
   * @param cells - Array of cell assignments
   * @returns Write result with changed cells and any errors
   */
  async setCells(
    cells: CellAssignment[],
    options: { validationMode?: SetCellsValidationMode } = {}
  ): Promise<WriteResult> {
    return (await this.request(
      'setCells',
      'setCells',
      dropUndefined({ cells, validationMode: options.validationMode })
    )) as WriteResult;
  }

  /**
   * Set the style of a cell or range.
   *
   * @param target - Cell or range address
   * @param style - Style properties to apply
   */
  async setStyle(target: string, style: Style): Promise<void> {
    await this.request('setStyle', 'setStyle', { address: target, style });
  }

  /**
   * Copy a range to another location.
   *
   * @param source - Source range address
   * @param destination - Destination cell address
   * @param options - Copy options
   * @returns Copy result with destination and cell count
   */
  async copyRange(
    source: string,
    destination: string,
    options: { pasteType?: PasteType } = {}
  ): Promise<CopyRangeResult> {
    return (await this.request(
      'copyRange',
      'copyRange',
      dropUndefined({
        source,
        destination,
        pasteType: options.pasteType,
      })
    )) as CopyRangeResult;
  }

  /**
   * Scale numeric values in a range by a factor.
   * This is a composite operation that reads the range, filters numeric cells,
   * and calls setCells with the scaled values.
   *
   * @param range - Range address
   * @param factor - Multiplication factor
   * @param options - Scale options
   * @returns Write result, or null if no numeric cells were found
   */
  async scaleRange(
    range: string,
    factor: number,
    options: { skipFormulas?: boolean } = {}
  ): Promise<WriteResult | null> {
    const skipFormulas = options.skipFormulas ?? true;
    const data = await this.readRange(range);

    const assignments: CellAssignment[] = [];
    for (const row of data) {
      for (const cell of row) {
        const value = cell.value;
        const hasFormula = Boolean(cell.formula);

        // Only scale numeric values (not booleans)
        if (typeof value === 'number' && (!hasFormula || !skipFormulas)) {
          assignments.push({ address: cell.address, value: value * factor });
        }
      }
    }

    if (assignments.length === 0) {
      return null;
    }

    return this.setCells(assignments);
  }

  // ============================================================================
  // Search Operations
  // ============================================================================

  /**
   * Find cells matching a pattern.
   *
   * @param matcher - String, number, boolean, RegExp, or array of strings/RegExps
   * @param options - Search options
   * @returns Array of matching cells
   */
  async findCells(
    matcher: Matcher,
    options: {
      in?: string;
      context?: number;
      limit?: number;
      offset?: number;
      formulas?: boolean;
    } = {}
  ): Promise<SearchCell[]> {
    const result = (await this.request(
      'findCells',
      'findCells',
      dropUndefined({
        matcher: serializeMatcher(matcher),
        in: options.in,
        context: options.context ?? 2,
        limit: options.limit ?? 20,
        offset: options.offset ?? 0,
        formulas: options.formulas,
      })
    )) as { matches?: SearchCell[] };
    return result.matches ?? [];
  }

  /**
   * Find rows containing cells matching a pattern.
   *
   * @param matcher - String, number, boolean, RegExp, or array of strings/RegExps
   * @param options - Search options
   * @returns Array of matching rows
   */
  async findRows(
    matcher: Matcher,
    options: {
      in?: string;
      context?: number;
      limit?: number;
      offset?: number;
    } = {}
  ): Promise<SearchRow[]> {
    const result = (await this.request(
      'findRows',
      'findRows',
      dropUndefined({
        matcher: serializeMatcher(matcher),
        in: options.in,
        context: options.context,
        limit: options.limit ?? 20,
        offset: options.offset ?? 0,
      })
    )) as { matches?: SearchRow[] };
    return result.matches ?? [];
  }

  /**
   * Find and replace text in cells.
   *
   * @param find - String or RegExp to find
   * @param replace - Replacement string
   * @param options - Find and replace options
   * @returns Result with count of replacements made
   */
  async findAndReplace(
    find: ReplaceMatcher,
    replace: string,
    options: {
      in?: string;
      matchCase?: boolean;
      wholeCell?: boolean;
      inFormulas?: boolean;
      limit?: number;
    } = {}
  ): Promise<FindAndReplaceResult> {
    return (await this.request(
      'findAndReplace',
      'findAndReplace',
      dropUndefined({
        find: serializeMatcher(find),
        replace,
        in: options.in,
        matchCase: options.matchCase,
        wholeCell: options.wholeCell,
        inFormulas: options.inFormulas,
        limit: options.limit,
      })
    )) as FindAndReplaceResult;
  }

  // ============================================================================
  // Row/Column Structure Operations
  // ============================================================================

  /**
   * Insert rows after a specified row.
   *
   * @param sheet - Sheet name
   * @param row - Row number after which to insert (1-based)
   * @param count - Number of rows to insert (default: 1)
   */
  async insertRowAfter(sheet: string, row: number, count: number = 1): Promise<void> {
    await this.request('insertRowAfter', 'insertRowAfter', { sheet, row, count });
  }

  /**
   * Delete rows from a sheet.
   *
   * @param sheet - Sheet name
   * @param row - Starting row number (1-based)
   * @param count - Number of rows to delete (default: 1)
   */
  async deleteRows(sheet: string, row: number, count: number = 1): Promise<void> {
    await this.request('deleteRows', 'deleteRows', { sheet, row, count });
  }

  /**
   * Insert columns after a specified column.
   *
   * @param sheet - Sheet name
   * @param column - Column number (1-based) or letter after which to insert
   * @param count - Number of columns to insert (default: 1)
   */
  async insertColumnAfter(sheet: string, column: number | string, count: number = 1): Promise<void> {
    await this.request('insertColumnAfter', 'insertColumnAfter', { sheet, column, count });
  }

  /**
   * Delete columns from a sheet.
   *
   * @param sheet - Sheet name
   * @param column - Starting column number (1-based) or letter
   * @param count - Number of columns to delete (default: 1)
   */
  async deleteColumns(sheet: string, column: number | string, count: number = 1): Promise<void> {
    await this.request('deleteColumns', 'deleteColumns', { sheet, column, count });
  }

  // ============================================================================
  // Row/Column Properties
  // ============================================================================

  /**
   * Set properties for a range of rows.
   *
   * @param sheet - Sheet name
   * @param fromRow - Starting row number (1-based)
   * @param toRow - Ending row number (1-based)
   * @param properties - Properties to set (e.g., height, hidden)
   */
  async setRowProperties(
    sheet: string,
    fromRow: number,
    toRow: number,
    properties: RowProperties
  ): Promise<void> {
    await this.request('setRowProperties', 'setRowProperties', {
      sheet,
      fromRow,
      toRow,
      properties,
    });
  }

  /**
   * Set properties for a range of columns.
   *
   * @param sheet - Sheet name
   * @param fromCol - Starting column number (1-based) or letter
   * @param toCol - Ending column number (1-based) or letter
   * @param properties - Properties to set (e.g., width, hidden)
   */
  async setColumnProperties(
    sheet: string,
    fromCol: number | string,
    toCol: number | string,
    properties: ColumnProperties
  ): Promise<void> {
    await this.request('setColumnProperties', 'setColumnProperties', {
      sheet,
      fromCol,
      toCol,
      properties,
    });
  }

  /**
   * Auto-fit column widths to content.
   *
   * @param sheet - Sheet name
   * @param columns - Specific columns to fit (default: all)
   * @param options - Auto-fit options
   * @returns Map of column letters to width results
   */
  async autoFitColumns(
    sheet: string,
    columns?: (number | string)[],
    options: { minWidth?: number; maxWidth?: number; padding?: number } = {}
  ): Promise<Record<string, AutoFitColumnResult>> {
    const result = (await this.request(
      'autoFitColumns',
      'autoFitColumns',
      dropUndefined({
        sheet,
        columns,
        minWidth: options.minWidth,
        maxWidth: options.maxWidth,
        padding: options.padding,
      })
    )) as { columns?: Record<string, AutoFitColumnResult> };
    return result.columns ?? {};
  }

  /**
   * Auto-fit row heights to content.
   *
   * @param sheet - Sheet name
   * @param rows - Specific rows to fit (default: all)
   * @param options - Auto-fit options
   * @returns Map of row numbers to height results
   */
  async autoFitRows(
    sheet: string,
    rows?: number[],
    options: { minHeight?: number; maxHeight?: number } = {}
  ): Promise<Record<string, AutoFitRowResult>> {
    const result = (await this.request(
      'autoFitRows',
      'autoFitRows',
      dropUndefined({
        sheet,
        rows,
        minHeight: options.minHeight,
        maxHeight: options.maxHeight,
      })
    )) as { rows?: Record<string, AutoFitRowResult> };
    return result.rows ?? {};
  }

  /**
   * Sort a range by specified columns.
   *
   * @param range - Range address to sort
   * @param keys - Sort keys specifying columns and order
   * @param options - Sort options
   */
  async sortRange(
    range: string,
    keys: SortKey[],
    options: { hasHeader?: boolean } = {}
  ): Promise<void> {
    await this.request(
      'sortRange',
      'sortRange',
      dropUndefined({
        range,
        keys,
        hasHeader: options.hasHeader,
      })
    );
  }

  // ============================================================================
  // Defined Names
  // ============================================================================

  /**
   * List all defined names in the workbook.
   */
  async listDefinedNames(): Promise<DefinedName[]> {
    return (await this.request('listDefinedNames', 'listDefinedNames', {})) as DefinedName[];
  }

  /**
   * Add a defined name to the workbook.
   *
   * @param name - Name to define
   * @param range - Range address the name refers to
   * @param options - Options including optional scope
   * @returns The created defined name
   */
  async addDefinedName(
    name: string,
    range: string,
    options: { scope?: string } = {}
  ): Promise<DefinedName> {
    return (await this.request(
      'addDefinedName',
      'addDefinedName',
      dropUndefined({ name, range, scope: options.scope })
    )) as DefinedName;
  }

  /**
   * Delete a defined name from the workbook.
   *
   * @param name - Name to delete
   * @param options - Options including optional scope
   * @returns The deleted defined name
   */
  async deleteDefinedName(name: string, options: { scope?: string } = {}): Promise<DefinedName> {
    return (await this.request(
      'deleteDefinedName',
      'deleteDefinedName',
      dropUndefined({ name, scope: options.scope })
    )) as DefinedName;
  }

  // ============================================================================
  // List Objects (Tables)
  // ============================================================================

  /**
   * Get a list object (table) by name.
   *
   * @param name - Name of the list object
   * @returns The list object
   */
  async getListObject(name: string): Promise<ListObject> {
    return (await this.request('getListObject', 'getListObject', { name })) as ListObject;
  }

  /**
   * Add a list object (table) to a sheet.
   *
   * @param sheet - Sheet name
   * @param listObject - List object specification
   * @returns Mutation result with the created list object
   */
  async addListObject(sheet: string, listObject: ListObjectSpec): Promise<ListObjectMutationResult> {
    return (await this.request('addListObject', 'addListObject', {
      sheet,
      listObject,
    })) as ListObjectMutationResult;
  }

  /**
   * Update a list object (table).
   *
   * @param name - Name of the list object to update
   * @param listObject - Updated list object properties
   * @returns Mutation result with the updated list object
   */
  async setListObject(name: string, listObject: ListObjectUpdate): Promise<ListObjectMutationResult> {
    return (await this.request('setListObject', 'setListObject', {
      name,
      listObject,
    })) as ListObjectMutationResult;
  }

  /**
   * Delete a list object (table).
   *
   * @param name - Name of the list object to delete
   * @returns Write result
   */
  async deleteListObject(name: string): Promise<WriteResult> {
    return (await this.request('deleteListObject', 'deleteListObject', { name })) as WriteResult;
  }

  // ============================================================================
  // Data Tables
  // ============================================================================

  /**
   * Get a data table by address.
   *
   * @param address - Address of the data table
   * @returns The data table
   */
  async getDataTable(address: string): Promise<DataTable> {
    return (await this.request('getDataTable', 'getDataTable', { address })) as DataTable;
  }

  /**
   * Add a data table to a sheet.
   *
   * @param sheet - Sheet name
   * @param dataTable - Data table specification
   * @returns Mutation result with the created data table
   */
  async addDataTable(sheet: string, dataTable: DataTableSpec): Promise<DataTableMutationResult> {
    return (await this.request('addDataTable', 'addDataTable', {
      sheet,
      dataTable,
    })) as DataTableMutationResult;
  }

  /**
   * Delete a data table.
   *
   * @param address - Address of the data table to delete
   * @returns Write result
   */
  async deleteDataTable(address: string): Promise<WriteResult> {
    return (await this.request('deleteDataTable', 'deleteDataTable', { address })) as WriteResult;
  }

  // ============================================================================
  // Charts
  // ============================================================================

  /**
   * List charts in the workbook.
   *
   * @param options - Options to filter by sheet
   * @returns Array of chart summaries
   */
  async listCharts(options: { sheet?: string } = {}): Promise<ChartSummary[]> {
    const result = (await this.request(
      'listCharts',
      'listCharts',
      dropUndefined({ sheet: options.sheet })
    )) as { charts?: ChartSummary[] };
    return result.charts ?? [];
  }

  /**
   * Get a chart by name.
   *
   * @param sheet - Sheet containing the chart
   * @param name - Chart name
   * @returns The chart info
   */
  async getChart(sheet: string, name: string): Promise<ChartInfo> {
    const result = (await this.request('getChart', 'getChart', { sheet, name })) as {
      chart?: ChartInfo;
    };
    return result.chart ?? ({} as ChartInfo);
  }

  /**
   * Add a chart to a sheet.
   *
   * @param sheet - Sheet name
   * @param chart - Chart specification
   * @returns The created chart
   */
  async addChart(sheet: string, chart: ChartSpec): Promise<ChartSpec> {
    const result = (await this.request('addChart', 'addChart', { sheet, chart })) as {
      chart?: ChartSpec;
    };
    return result.chart ?? ({} as ChartSpec);
  }

  /**
   * Update a chart.
   *
   * @param sheet - Sheet containing the chart
   * @param name - Chart name
   * @param chart - Updated chart specification
   * @returns The updated chart
   */
  async setChart(sheet: string, name: string, chart: ChartSpec): Promise<ChartSpec> {
    const result = (await this.request('setChart', 'setChart', { sheet, name, chart })) as {
      chart?: ChartSpec;
    };
    return result.chart ?? ({} as ChartSpec);
  }

  /**
   * Delete a chart.
   *
   * @param sheet - Sheet containing the chart
   * @param name - Chart name
   */
  async deleteChart(sheet: string, name: string): Promise<void> {
    await this.request('deleteChart', 'deleteChart', { sheet, name });
  }

  // ============================================================================
  // Images
  // ============================================================================

  /**
   * List worksheet images in the workbook.
   *
   * @param options - Options to filter by sheet
   * @returns Array of image metadata
   */
  async listImages(options: { sheet?: string } = {}): Promise<ImageInfo[]> {
    const result = (await this.request(
      'listImages',
      'listImages',
      dropUndefined({ sheet: options.sheet })
    )) as { images?: ImageInfo[] };
    return result.images ?? [];
  }

  /**
   * Get worksheet image metadata by name or id.
   *
   * @param sheet - Sheet containing the image
   * @param selector - Image name or sheet-local id
   * @returns The image metadata
   */
  async getImage(sheet: string, selector: ImageSelector): Promise<ImageInfo> {
    const result = (await this.request(
      'getImage',
      'getImage',
      dropUndefined({ sheet, name: selector.name, id: selector.id })
    )) as { image?: ImageInfo };
    return result.image ?? ({} as ImageInfo);
  }

  /**
   * Add a PNG or JPEG image to a sheet.
   *
   * @param sheet - Sheet name
   * @param image - Image specification
   * @returns The created image metadata
   */
  async addImage(sheet: string, image: ImageSpec): Promise<ImageInfo> {
    const result = (await this.request('addImage', 'addImage', { sheet, image })) as {
      image?: ImageInfo;
    };
    return result.image ?? ({} as ImageInfo);
  }

  /**
   * Update worksheet image metadata, placement, or source bytes.
   *
   * @param sheet - Sheet containing the image
   * @param selector - Image name or sheet-local id
   * @param image - Image updates
   * @returns The updated image metadata
   */
  async setImage(sheet: string, selector: ImageSelector, image: ImageUpdate): Promise<ImageInfo> {
    const result = (await this.request(
      'setImage',
      'setImage',
      dropUndefined({ sheet, name: selector.name, id: selector.id, image })
    )) as { image?: ImageInfo };
    return result.image ?? ({} as ImageInfo);
  }

  /**
   * Delete a worksheet image by name or id.
   *
   * @param sheet - Sheet containing the image
   * @param selector - Image name or sheet-local id
   */
  async deleteImage(sheet: string, selector: ImageSelector): Promise<void> {
    await this.request(
      'deleteImage',
      'deleteImage',
      dropUndefined({ sheet, name: selector.name, id: selector.id })
    );
  }

  // ============================================================================
  // Conditional Formatting
  // ============================================================================

  /**
   * Get conditional formatting rules for a sheet.
   *
   * @param sheet - Sheet name
   * @returns Array of conditional formatting rules
   */
  async getConditionalFormatting(sheet: string): Promise<ConditionalFormattingRule[]> {
    const result = (await this.request('getConditionalFormatting', 'getConditionalFormatting', {
      sheet,
    })) as { rules?: ConditionalFormattingRule[] };
    return result.rules ?? [];
  }

  /**
   * Set conditional formatting rules for a sheet.
   *
   * @param sheet - Sheet name
   * @param rules - Array of conditional formatting rules
   * @param options - Options including whether to clear existing rules
   */
  async setConditionalFormatting(
    sheet: string,
    rules: ConditionalFormattingRule[],
    options: { clear?: boolean } = {}
  ): Promise<void> {
    await this.request(
      'setConditionalFormatting',
      'setConditionalFormatting',
      dropUndefined({ sheet, rules, clear: options.clear })
    );
  }

  /**
   * Remove conditional formatting rules by index.
   *
   * @param sheet - Sheet name
   * @param indices - Indices of rules to remove
   */
  async removeConditionalFormatting(sheet: string, indices: number[]): Promise<void> {
    await this.request('removeConditionalFormatting', 'removeConditionalFormatting', {
      sheet,
      indices,
    });
  }

  // ============================================================================
  // Data Validation
  // ============================================================================

  /**
   * Get data validation rules.
   *
   * @param options - Optional sheet and address filters
   * @returns Array of data validation rules
   */
  async getDataValidations(options: { sheet?: string; address?: string } = {}): Promise<DataValidationInfo[]> {
    const result = (await this.request(
      'getDataValidations',
      'getDataValidations',
      dropUndefined({ sheet: options.sheet, address: options.address })
    )) as { rules?: DataValidationInfo[] };
    return result.rules ?? [];
  }

  /**
   * Validate current cell values against their data validation rules.
   *
   * @param address - Range address to validate
   * @param options - Evaluation limits and unsupported-rule handling
   * @returns Validation status, invalid cell ranges, and diagnostics
   */
  async validateCells(
    address: string,
    options: {
      maxCellsToScan?: number;
      maxInvalidCells?: number;
      treatUnsupportedAsInvalid?: boolean;
    } = {}
  ): Promise<DataValidationResult> {
    return (await this.request(
      'validateCells',
      'validateCells',
      dropUndefined({
        address,
        maxCellsToScan: options.maxCellsToScan,
        maxInvalidCells: options.maxInvalidCells,
        treatUnsupportedAsInvalid: options.treatUnsupportedAsInvalid,
      })
    )) as DataValidationResult;
  }

  /**
   * Add data validation rules to a sheet.
   *
   * @param sheet - Sheet name
   * @param rules - Rules to add
   * @param options - Options including whether to clear existing rules
   */
  async setDataValidations(
    sheet: string,
    rules: DataValidationSpec[],
    options: { clear?: boolean } = {}
  ): Promise<void> {
    await this.request(
      'setDataValidations',
      'setDataValidations',
      dropUndefined({ sheet, rules, clear: options.clear })
    );
  }

  /**
   * Remove data validation rules by index or by range.
   *
   * @param sheet - Sheet name
   * @param target - Either rule indices or an address to clear
   */
  async removeDataValidations(sheet: string, target: { indices: number[] } | { address: string }): Promise<void> {
    await this.request('removeDataValidations', 'removeDataValidations', {
      sheet,
      ...target,
    });
  }

  // ============================================================================
  // Formula Operations
  // ============================================================================

  /**
   * Evaluate multiple formulas in a sheet context.
   *
   * @param sheet - Sheet name for formula context
   * @param formulas - Array of formula strings
   * @returns Array of formula results
   */
  async evaluateFormulas(sheet: string, formulas: string[]): Promise<FormulaResult[]> {
    return (await this.request('evaluateFormulas', 'evaluateFormulas', {
      sheet,
      formulas,
    })) as FormulaResult[];
  }

  /**
   * Evaluate a single formula in a sheet context.
   * This is a convenience wrapper around evaluateFormulas.
   *
   * @param sheet - Sheet name for formula context
   * @param formula - Formula string
   * @returns The formula result
   */
  async evaluateFormula(sheet: string, formula: string): Promise<FormulaResult> {
    const results = await this.evaluateFormulas(sheet, [formula]);
    return results[0]!;
  }

  /**
   * Get cells that a formula depends on (precedents).
   *
   * @param address - Cell address
   * @param depth - How many levels to trace (default: 1, use Infinity for all)
   * @returns Dependency result with cells and warnings
   */
  async getCellPrecedents(address: string, depth: number | typeof Infinity = 1): Promise<DependencyResult> {
    const rpcDepth = Number.isFinite(depth) ? Math.floor(depth) : -1;
    return (await this.request('getCellPrecedents', 'getCellPrecedents', {
      address,
      depth: rpcDepth,
    })) as DependencyResult;
  }

  /**
   * Get cells that depend on a cell (dependents).
   *
   * @param address - Cell address
   * @param depth - How many levels to trace (default: 1, use Infinity for all)
   * @returns Dependency result with cells and warnings
   */
  async getCellDependents(address: string, depth: number | typeof Infinity = 1): Promise<DependencyResult> {
    const rpcDepth = Number.isFinite(depth) ? Math.floor(depth) : -1;
    return (await this.request('getCellDependents', 'getCellDependents', {
      address,
      depth: rpcDepth,
    })) as DependencyResult;
  }

  /**
   * Trace a cell to its input sources.
   *
   * @param address - Cell address
   * @returns Array of input traces
   */
  async traceToInputs(address: string): Promise<TraceInput[]> {
    return (await this.request('traceToInputs', 'traceToInputs', { address })) as TraceInput[];
  }

  /**
   * Trace a cell to its output destinations.
   *
   * @param address - Cell address
   * @returns Array of output traces
   */
  async traceToOutputs(address: string): Promise<TraceOutput[]> {
    return (await this.request('traceToOutputs', 'traceToOutputs', { address })) as TraceOutput[];
  }

  // ============================================================================
  // Scenarios / Sweep
  // ============================================================================

  /**
   * Run a sweep over input combinations and capture outputs.
   *
   * @param inputs - Array of input specifications with addresses and values
   * @param outputs - Array of output cell addresses
   * @param options - Sweep options
   * @returns Sweep result with all combinations
   */
  async sweepInputs(
    inputs: SweepInput[],
    outputs: string[],
    options: { mode?: SweepMode; includeStats?: boolean } = {}
  ): Promise<SweepResult> {
    return (await this.request(
      'sweepInputs',
      'sweepInputs',
      dropUndefined({
        inputs,
        outputs,
        mode: options.mode,
        includeStats: options.includeStats,
      })
    )) as SweepResult;
  }

  /**
   * Alias for sweepInputs.
   */
  async scenarios(
    inputs: SweepInput[],
    outputs: string[],
    options: { mode?: SweepMode; includeStats?: boolean } = {}
  ): Promise<SweepResult> {
    return this.sweepInputs(inputs, outputs, options);
  }

  // ============================================================================
  // Utilities
  // ============================================================================

  /**
   * Get a description of a sheet's structure.
   *
   * @param sheet - Sheet name
   * @returns Sheet description
   */
  async describeSheet(sheet: string): Promise<SheetDescription> {
    return (await this.request('describeSheet', 'describeSheet', { sheet })) as SheetDescription;
  }

  /**
   * Get descriptions of all visible sheets.
   * This is a composite operation that calls listSheets and describeSheet.
   *
   * @returns Map of sheet names to descriptions
   */
  async describeSheets(): Promise<Record<string, SheetDescription>> {
    const sheets = await this.listSheets();
    const result: Record<string, SheetDescription> = {};
    for (const sheet of sheets) {
      if (!sheet.hidden) {
        result[sheet.sheet] = await this.describeSheet(sheet.sheet);
      }
    }
    return result;
  }

  /**
   * Look up a value in a table by row and column labels.
   *
   * @param table - Table name or address
   * @param rowLabel - Row label to find
   * @param columnLabel - Column label to find
   * @returns Array of lookup results
   */
  async tableLookup(
    table: string,
    rowLabel: string | number | boolean,
    columnLabel: string | number | boolean
  ): Promise<TableLookupResult[]> {
    return (await this.request('tableLookup', 'tableLookup', {
      table,
      rowLabel,
      columnLabel,
    })) as TableLookupResult[];
  }

  /**
   * Run linting rules on the workbook.
   *
   * @param options - Lint options
   * @returns Lint result with diagnostics
   */
  async lint(
    options: {
      rangeAddresses?: string[];
      skipRuleIds?: string[];
      onlyRuleIds?: string[];
    } = {}
  ): Promise<LintResult> {
    return (await this.request(
      'lint',
      'lint',
      dropUndefined({
        rangeAddresses: options.rangeAddresses,
        skipRuleIds: options.skipRuleIds,
        onlyRuleIds: options.onlyRuleIds,
      })
    )) as LintResult;
  }

  /**
   * Generate a preview image of cell styles.
   *
   * @param range - Range address to preview
   * @returns Data URL of the preview image
   */
  async previewStyles(range: string): Promise<string> {
    const result = (await this.request('previewStyles', 'previewStyles', { address: range })) as {
      contentType?: string;
      data?: string;
    };
    const contentType = result.contentType;
    const data = result.data;
    if (typeof contentType !== 'string' || typeof data !== 'string') {
      throw new WitanProcessError(`Invalid previewStyles result: ${JSON.stringify(result)}`);
    }
    return `data:${contentType};base64,${data}`;
  }

  /**
   * Reduce a list of addresses to minimal non-overlapping ranges.
   *
   * @param addresses - Array of cell or range addresses
   * @returns Array of reduced addresses
   */
  async reduceAddresses(addresses: string[]): Promise<string[]> {
    return (await this.request('reduceAddresses', 'reduceAddresses', { addresses })) as string[];
  }

  /**
   * Get the style of a cell.
   *
   * @param cell - Cell address
   * @returns Style properties
   */
  async getStyle(cell: string): Promise<Style> {
    return (await this.request('getStyle', 'getStyle', { address: cell })) as Style;
  }
}
