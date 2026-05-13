import { getBinaryPath } from './binary.js';
import { WitanProcessError } from './errors.js';
import { dropUndefined } from './helpers.js';
import { StdioRPCProcess } from './process.js';
import type {
  CellAssignment,
  CopyRangeResult,
  JsonMapping,
  PasteType,
  SheetInfo,
  SheetProperties,
  Style,
  Value,
  WorkbookProperties,
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
  /**
   * Request timeout in milliseconds (default: 90000).
   *
   * Note: The Python SDK uses seconds. When migrating from Python,
   * multiply your timeout value by 1000 (e.g., Python's `request_timeout=30`
   * becomes `requestTimeout: 30_000` in Node.js).
   */
  requestTimeout?: number;
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
      timeout: options.requestTimeout,
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
  async setWorkbookProperties(properties: JsonMapping): Promise<void> {
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
  async setSheetProperties(sheet: string, properties: JsonMapping): Promise<void> {
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
  async setCells(cells: CellAssignment[]): Promise<WriteResult> {
    return (await this.request('setCells', 'setCells', { cells })) as WriteResult;
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
}
