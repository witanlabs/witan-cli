import { getBinaryPath } from './binary.js';
import { WitanProcessError } from './errors.js';
import { dropUndefined } from './helpers.js';
import { StdioRPCProcess } from './process.js';
import type {
  SheetInfo,
  SheetProperties,
  WorkbookProperties,
  JsonMapping,
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
}
