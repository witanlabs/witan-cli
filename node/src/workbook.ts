import { getBinaryPath } from './binary.js';
import { WitanProcessError } from './errors.js';
import { dropUndefined } from './helpers.js';
import { StdioRPCProcess } from './process.js';
import { SpreadsheetSessionBase } from './spreadsheet-base.js';
import type {
  DataTable,
  DataTableMutationResult,
  DataTableSpec,
  DataValidationInfo,
  DataValidationSpec,
  FormulaResult,
  ImageInfo,
  ImageSelector,
  ImageSpec,
  ImageUpdate,
  ListObject,
  ListObjectMutationResult,
  ListObjectSpec,
  ListObjectUpdate,
  SweepInput,
  SweepMode,
  SweepResult,
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
  // The ops below are not supported by the Google Sheets backend
  // (NOT_IMPLEMENTED), so they live on Workbook rather than the shared base.
  // ============================================================================

  /** Get a list object (table) by name. */
  async getListObject(name: string): Promise<ListObject> {
    return (await this.request('getListObject', 'getListObject', { name })) as ListObject;
  }

  /** Add a list object (table) to a sheet. */
  async addListObject(sheet: string, listObject: ListObjectSpec): Promise<ListObjectMutationResult> {
    return (await this.request('addListObject', 'addListObject', {
      sheet,
      listObject,
    })) as ListObjectMutationResult;
  }

  /** Update a list object (table). */
  async setListObject(name: string, listObject: ListObjectUpdate): Promise<ListObjectMutationResult> {
    return (await this.request('setListObject', 'setListObject', {
      name,
      listObject,
    })) as ListObjectMutationResult;
  }

  /** Delete a list object (table). */
  async deleteListObject(name: string): Promise<WriteResult> {
    return (await this.request('deleteListObject', 'deleteListObject', { name })) as WriteResult;
  }

  /** Get a data table by address. */
  async getDataTable(address: string): Promise<DataTable> {
    return (await this.request('getDataTable', 'getDataTable', { address })) as DataTable;
  }

  /** Add a data table to a sheet. */
  async addDataTable(sheet: string, dataTable: DataTableSpec): Promise<DataTableMutationResult> {
    return (await this.request('addDataTable', 'addDataTable', {
      sheet,
      dataTable,
    })) as DataTableMutationResult;
  }

  /** Delete a data table. */
  async deleteDataTable(address: string): Promise<WriteResult> {
    return (await this.request('deleteDataTable', 'deleteDataTable', { address })) as WriteResult;
  }

  /** List worksheet images in the workbook. */
  async listImages(options: { sheet?: string } = {}): Promise<ImageInfo[]> {
    const result = (await this.request(
      'listImages',
      'listImages',
      dropUndefined({ sheet: options.sheet })
    )) as { images?: ImageInfo[] };
    return result.images ?? [];
  }

  /** Get worksheet image metadata by name or id. */
  async getImage(sheet: string, selector: ImageSelector): Promise<ImageInfo> {
    const result = (await this.request(
      'getImage',
      'getImage',
      dropUndefined({ sheet, name: selector.name, id: selector.id })
    )) as { image?: ImageInfo };
    return result.image ?? ({} as ImageInfo);
  }

  /** Add a PNG or JPEG image to a sheet. */
  async addImage(sheet: string, image: ImageSpec): Promise<ImageInfo> {
    const result = (await this.request('addImage', 'addImage', { sheet, image })) as {
      image?: ImageInfo;
    };
    return result.image ?? ({} as ImageInfo);
  }

  /** Update worksheet image metadata, placement, or source bytes. */
  async setImage(sheet: string, selector: ImageSelector, image: ImageUpdate): Promise<ImageInfo> {
    const result = (await this.request(
      'setImage',
      'setImage',
      dropUndefined({ sheet, name: selector.name, id: selector.id, image })
    )) as { image?: ImageInfo };
    return result.image ?? ({} as ImageInfo);
  }

  /** Delete a worksheet image by name or id. */
  async deleteImage(sheet: string, selector: ImageSelector): Promise<void> {
    await this.request(
      'deleteImage',
      'deleteImage',
      dropUndefined({ sheet, name: selector.name, id: selector.id })
    );
  }

  /** Run a sweep over input combinations and capture outputs. */
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

  /** Alias for sweepInputs. */
  async scenarios(
    inputs: SweepInput[],
    outputs: string[],
    options: { mode?: SweepMode; includeStats?: boolean } = {}
  ): Promise<SweepResult> {
    return this.sweepInputs(inputs, outputs, options);
  }

  /** Evaluate multiple formulas in a sheet context. */
  async evaluateFormulas(sheet: string, formulas: string[]): Promise<FormulaResult[]> {
    return (await this.request('evaluateFormulas', 'evaluateFormulas', {
      sheet,
      formulas,
    })) as FormulaResult[];
  }

  /** Evaluate a single formula in a sheet context. */
  async evaluateFormula(sheet: string, formula: string): Promise<FormulaResult> {
    const results = await this.evaluateFormulas(sheet, [formula]);
    return results[0]!;
  }
}
