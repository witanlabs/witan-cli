import { getBinaryPath } from './binary.js';
import { WitanProcessError } from './errors.js';
import { dropUndefined } from './helpers.js';
import { StdioRPCProcess } from './process.js';
import { SpreadsheetSessionBase } from './spreadsheet-base.js';
import type {
  DataValidationInfo,
  DataValidationResult,
  DataValidationSpec,
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
export class Workbook extends SpreadsheetSessionBase implements AsyncDisposable {
  private process: StdioRPCProcess;
  private requestId = 0;
  private closed = false;

  private constructor(process: StdioRPCProcess) {
    super();
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

  private nextId(): string {
    return String(++this.requestId);
  }

  protected override async request(
    method: string,
    op: string,
    args: Record<string, unknown> = {}
  ): Promise<unknown> {
    if (this.closed) {
      throw new WitanProcessError('Workbook is closed');
    }
    return this.process.request(method, op, args, this.nextId());
  }

  /**
   * Close the workbook and terminate the subprocess.
   * Safe to call multiple times.
   */
  async close(): Promise<void> {
    if (this.closed) return;
    this.closed = true;
    await this.process.close();
  }

  async [Symbol.asyncDispose](): Promise<void> {
    await this.close();
  }

  /** Whether the workbook has been closed. */
  get isClosed(): boolean {
    return this.closed;
  }

  /**
   * Save the workbook to disk.
   * @returns true if the save was successful
   */
  async save(): Promise<boolean> {
    return (await this.request('save', 'save', {})) as boolean;
  }

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
}
