import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { getBinaryPath } from './binary.js';
import { WitanProcessError } from './errors.js';
import { StdioRPCProcess } from './process.js';
import { SpreadsheetSessionBase } from './spreadsheet-base.js';

const execFileAsync = promisify(execFile);

/** Default request timeout for Google Sheets sessions (matches Python SDK). */
export const DEFAULT_GOOGLE_SHEET_REQUEST_TIMEOUT_MS = 180_000;

/**
 * The authorize/status CLI commands cap their internal poll at 5 minutes;
 * allow a little headroom for the subprocess round trip.
 */
const AUTHORIZE_WAIT_TIMEOUT_MS = 360_000;

// Explicit "create a new sheet" sentinels. An empty ref is NOT included here:
// it must not silently infer create mode (e.g. an unset env var / config value).
const CREATE_REFS = new Set(['new', 'gs://new']);

/** Options for the one-shot authorize/status helper methods. */
export interface GoogleSheetAuthorizeOptions {
  /** API URL override */
  apiUrl?: string;
  /** Custom binary path (overrides auto-detection) */
  binary?: string;
  /** Additional environment variables for the subprocess */
  env?: Record<string, string>;
}

/** Result of {@link GoogleSheet.authorizeUrl}. */
export interface AuthorizeUrlResult {
  /** True when the sheet is already authorized (no picker needed). */
  authorized: boolean;
  /** Google Picker URL to open in a browser; absent when already authorized. */
  pickerUrl?: string;
  /** Seconds until the picker URL's state token expires. */
  expiresInSeconds?: number;
}

async function runCliJson(
  extra: string[],
  options: GoogleSheetAuthorizeOptions & { timeoutMs?: number } = {}
): Promise<Record<string, unknown>> {
  const binary = options.binary ?? getBinaryPath();
  const args: string[] = [];
  if (options.apiUrl !== undefined) {
    args.push('--api-url', options.apiUrl);
  }
  args.push(...extra);

  let stdout: string;
  try {
    const result = await execFileAsync(binary, args, {
      env: options.env ? { ...process.env, ...options.env } : process.env,
      timeout: options.timeoutMs ?? 60_000,
      maxBuffer: 10 * 1024 * 1024,
    });
    stdout = result.stdout;
  } catch (err) {
    const e = err as { stderr?: string };
    const stderrTail = (e.stderr ?? '')
      .trim()
      .split('\n')
      .filter(Boolean)
      .slice(-10);
    throw new WitanProcessError(`witan ${extra.join(' ')} failed`, stderrTail);
  }

  const trimmed = stdout.trim();
  if (!trimmed) {
    return {};
  }
  try {
    return JSON.parse(trimmed) as Record<string, unknown>;
  } catch {
    throw new WitanProcessError(`invalid JSON from witan ${extra.join(' ')}: ${trimmed}`);
  }
}

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

/**
 * Options for creating a new Google Sheet. Excludes `create` (implicit) and
 * `title` — the title is the positional argument to `create()`, so it has a
 * single source and can't be silently dropped by the positional/options merge.
 */
export type GoogleSheetCreateOptions = Omit<GoogleSheetOptions, 'create' | 'title'> & GoogleSheetRejectedOptions;

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
    // An empty ref is allowed only in create mode (GoogleSheet.create() omits
    // it); a non-empty ref must be an explicit "new" sentinel.
    if (ref && !CREATE_REFS.has(ref)) {
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

  /**
   * Begin per-file authorization for `ref`.
   *
   * Returns `{ authorized: true }` if the sheet is already authorized,
   * otherwise `{ authorized: false, pickerUrl, expiresInSeconds }`. Hand
   * `pickerUrl` to a human to open Google's file picker, then call
   * {@link GoogleSheet.waitUntilAuthorized} and retry `open`.
   */
  static async authorizeUrl(
    ref: string,
    options: GoogleSheetAuthorizeOptions = {}
  ): Promise<AuthorizeUrlResult> {
    const data = await runCliJson(['gsheets', 'authorize', ref, '--json'], options);
    return {
      authorized: Boolean(data['authorized']),
      pickerUrl: data['picker_url'] as string | undefined,
      expiresInSeconds: data['expires_in_seconds'] as number | undefined,
    };
  }

  /** Return whether `ref` is authorized for the app. */
  static async isAuthorized(
    ref: string,
    options: GoogleSheetAuthorizeOptions = {}
  ): Promise<boolean> {
    const data = await runCliJson(['gsheets', 'status', ref, '--json'], options);
    return Boolean(data['authorized']);
  }

  /** Block until `ref` is authorized. Resolves true; rejects on timeout. */
  static async waitUntilAuthorized(
    ref: string,
    options: GoogleSheetAuthorizeOptions & { timeoutMs?: number } = {}
  ): Promise<boolean> {
    const data = await runCliJson(['gsheets', 'status', ref, '--wait', '--json'], {
      timeoutMs: AUTHORIZE_WAIT_TIMEOUT_MS,
      ...options,
    });
    return Boolean(data['authorized']);
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
