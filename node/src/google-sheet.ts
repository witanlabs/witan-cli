import { getBinaryPath } from './binary.js';
import { WitanProcessError } from './errors.js';
import { StdioRPCProcess } from './process.js';
import { SpreadsheetSessionBase } from './spreadsheet-base.js';

/** Default request timeout for Google Sheets sessions (matches Python SDK). */
export const DEFAULT_GOOGLE_SHEET_REQUEST_TIMEOUT_MS = 180_000;

const CREATE_REFS = new Set(['', 'new', 'gs://new']);

/**
 * Options for opening a Google Sheet session.
 */
export interface GoogleSheetOptions {
  /** Create a new spreadsheet instead of opening an existing ref */
  create?: boolean;
  /** Title when creating a new spreadsheet */
  title?: string;
  /** Locale for number/date formatting */
  locale?: string;
  /** API URL override */
  apiUrl?: string;
  /** Custom binary path (overrides auto-detection) */
  binary?: string;
  /** Additional environment variables for the subprocess */
  env?: Record<string, string>;
  /** Request timeout in milliseconds (default: 180000). */
  requestTimeoutMs?: number;
}

type GoogleSheetRejectedOptions = {
  /** @deprecated Google Sheets requires user authentication via `witan auth login`. */
  apiKey?: string;
  /** @deprecated Not supported for Google Sheets. */
  stateless?: boolean;
  /** @deprecated Not supported for Google Sheets. */
  hint?: string;
};

/** Options for opening an existing Google Sheet. */
export type GoogleSheetOpenOptions = GoogleSheetOptions & GoogleSheetRejectedOptions;

/** Options for creating a new Google Sheet (excludes `create`, which is implicit). */
export type GoogleSheetCreateOptions = Omit<GoogleSheetOptions, 'create'> & GoogleSheetRejectedOptions;

function rejectGoogleSheetOptions(options: GoogleSheetOpenOptions | GoogleSheetCreateOptions): void {
  if (options.apiKey !== undefined) {
    throw new Error(
      "Google Sheets requires user authentication. Do not pass apiKey; run 'witan auth login' instead."
    );
  }
  if (options.stateless !== undefined) {
    throw new Error('GoogleSheet does not support stateless mode.');
  }
  if (options.hint !== undefined) {
    throw new Error('GoogleSheet does not support hint.');
  }
}

function resolveCreate(ref: string, create: boolean): boolean {
  if (create) {
    return true;
  }
  return CREATE_REFS.has(ref);
}

function validateRefCreate(ref: string, create: boolean): void {
  if (create) {
    if (!CREATE_REFS.has(ref)) {
      throw new Error("create requires ref 'new' or gs://new, or omit ref (use GoogleSheet.create())");
    }
    return;
  }
  if (!ref) {
    throw new Error('ref is required when opening an existing spreadsheet');
  }
}

function buildArgv(
  ref: string,
  options: GoogleSheetOptions,
  rpcCreate: boolean
): string[] {
  const binary = options.binary ?? getBinaryPath();
  const argv = [binary];

  if (options.apiUrl !== undefined) {
    argv.push('--api-url', options.apiUrl);
  }

  argv.push('gsheets', 'rpc');

  if (rpcCreate) {
    argv.push('--create');
    if (options.title !== undefined) {
      argv.push('--title', options.title);
    }
    if (ref === 'new' || ref === 'gs://new') {
      argv.push(ref);
    }
  } else {
    argv.push(ref);
  }

  if (options.locale !== undefined) {
    argv.push('--locale', options.locale);
  }

  return argv;
}

/**
 * Async Google Sheets session backed by `witan gsheets rpc`.
 *
 * Requires prior CLI setup: `witan auth login` and `witan gsheets connect`.
 * Changes persist immediately; there is no `save()` method.
 *
 * @example
 * ```typescript
 * {
 *   await using sheet = await GoogleSheet.open('gs://YOUR_SHEET_REF');
 *   const sheets = await sheet.listSheets();
 * }
 * ```
 */
export class GoogleSheet extends SpreadsheetSessionBase implements AsyncDisposable {
  private readonly rpcCreate: boolean;
  readonly requestTimeoutMs: number;
  private process: StdioRPCProcess;
  private requestId = 0;
  private closed = false;

  private constructor(process: StdioRPCProcess, rpcCreate: boolean, requestTimeoutMs: number) {
    super();
    this.process = process;
    this.rpcCreate = rpcCreate;
    this.requestTimeoutMs = requestTimeoutMs;
  }

  /**
   * Open an existing Google spreadsheet.
   *
   * @param ref - Spreadsheet ref (`gs://...` or Google Sheets URL)
   */
  static async open(ref: string, options: GoogleSheetOpenOptions = {}): Promise<GoogleSheet> {
    rejectGoogleSheetOptions(options);
    const rpcCreate = resolveCreate(ref, options.create ?? false);
    validateRefCreate(ref, rpcCreate);
    return GoogleSheet.start(ref, options, rpcCreate);
  }

  /**
   * Create a new Google spreadsheet and open an RPC session.
   */
  static async create(
    title?: string,
    options: GoogleSheetCreateOptions = {}
  ): Promise<GoogleSheet> {
    rejectGoogleSheetOptions(options);
    return GoogleSheet.start('', { ...options, title }, true);
  }

  private static async start(
    ref: string,
    options: GoogleSheetOptions,
    rpcCreate: boolean
  ): Promise<GoogleSheet> {
    const requestTimeoutMs = options.requestTimeoutMs ?? DEFAULT_GOOGLE_SHEET_REQUEST_TIMEOUT_MS;
    const argv = buildArgv(ref, options, rpcCreate);
    const process = new StdioRPCProcess(argv, {
      env: options.env,
      timeoutMs: requestTimeoutMs,
    });

    try {
      await process.waitReady();
    } catch (err) {
      process.terminate();
      throw err;
    }

    return new GoogleSheet(process, rpcCreate, requestTimeoutMs);
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
      throw new WitanProcessError('GoogleSheet is closed');
    }
    return this.process.request(method, op, args, this.nextId());
  }

  async close(): Promise<void> {
    if (this.closed) return;
    this.closed = true;
    await this.process.close();
  }

  async [Symbol.asyncDispose](): Promise<void> {
    await this.close();
  }

  get isClosed(): boolean {
    return this.closed;
  }

  /** True when this session was opened in create mode. */
  get isCreate(): boolean {
    return this.rpcCreate;
  }
}
