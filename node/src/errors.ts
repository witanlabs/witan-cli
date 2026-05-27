/**
 * Base error class for all Witan SDK errors.
 */
export class WitanError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'WitanError';
  }
}

/**
 * Error thrown when the witan subprocess fails or exits unexpectedly.
 */
export class WitanProcessError extends WitanError {
  readonly stderrTail: string[];

  constructor(message: string, stderrTail: string[] = []) {
    const fullMessage = stderrTail.length
      ? `${message}\nstderr tail:\n${stderrTail.join('\n')}`
      : message;
    super(fullMessage);
    this.name = 'WitanProcessError';
    this.stderrTail = stderrTail;
  }
}

/**
 * Error thrown when an RPC request times out.
 */
export class WitanTimeoutError extends WitanProcessError {
  constructor(message: string, stderrTail: string[] = []) {
    super(message, stderrTail);
    this.name = 'WitanTimeoutError';
  }
}

/**
 * Error thrown when the RPC server returns an error response.
 */
export class WitanRPCError extends WitanError {
  readonly method: string;
  readonly op: string;
  readonly requestId: string;
  readonly code: string | null;
  readonly response: Record<string, unknown>;

  constructor(
    message: string,
    options: {
      method: string;
      op: string;
      requestId: string;
      code?: string | null;
      response?: Record<string, unknown>;
    }
  ) {
    const label = options.code ?? 'RPC_ERROR';
    super(`${label}: ${message}`);
    this.name = 'WitanRPCError';
    this.method = options.method;
    this.op = options.op;
    this.requestId = options.requestId;
    this.code = options.code ?? null;
    this.response = options.response ?? {};
  }
}

const GOOGLE_AUTH_REQUIRED_MARKERS = [
  // Error codes (when surfaced in a message / JSON).
  'google_auth_required',
  'google_sheets_not_connected',
  'google_sheets_scope_not_granted',
  // Every "needs connect/reconnect" CLI message ends with this remediation
  // (not-connected, expired/revoked, and sheet-op google_auth_required alike),
  // so match the command rather than each individual phrasing.
  'witan gsheets connect',
] as const;

const GOOGLE_AUTH_REQUIRED_CODES = [
  'google_auth_required',
  'google_sheets_not_connected',
  'google_sheets_scope_not_granted',
] as const;

function textIndicatesGoogleAuthRequired(text: string): boolean {
  return GOOGLE_AUTH_REQUIRED_MARKERS.some((marker) => text.includes(marker));
}

/**
 * Return true when `err` indicates the Google account must be connected or
 * re-authorized — i.e. the caller should run `witan gsheets connect`.
 *
 * Covers both an expired/revoked connection and the not-yet-connected case
 * (which surfaces from the authorize-sheet / status path, e.g. when
 * `GoogleSheet.authorizeUrl` is called before connecting).
 */
export function isGoogleAuthRequired(err: unknown): boolean {
  if (err instanceof WitanRPCError && err.code !== null && (GOOGLE_AUTH_REQUIRED_CODES as readonly string[]).includes(err.code)) {
    return true;
  }
  if (err instanceof WitanProcessError) {
    if (textIndicatesGoogleAuthRequired(err.message)) {
      return true;
    }
    return err.stderrTail.some(textIndicatesGoogleAuthRequired);
  }
  return false;
}

const NEEDS_FILE_AUTHORIZATION_MARKERS = [
  'needs_file_authorization',
  'must be authorized before Witan',
] as const;

function textIndicatesNeedsFileAuthorization(text: string): boolean {
  return NEEDS_FILE_AUTHORIZATION_MARKERS.some((marker) => text.includes(marker));
}

/**
 * Return true when `err` indicates the specific spreadsheet has not been
 * authorized for the app (drive.file scope). Recover by authorizing the sheet
 * (`GoogleSheet.authorizeUrl` + `GoogleSheet.waitUntilAuthorized`) and retrying.
 */
export function isNeedsFileAuthorization(err: unknown): boolean {
  if (err instanceof WitanRPCError && err.code === 'needs_file_authorization') {
    return true;
  }
  if (err instanceof WitanProcessError) {
    if (textIndicatesNeedsFileAuthorization(err.message)) {
      return true;
    }
    return err.stderrTail.some(textIndicatesNeedsFileAuthorization);
  }
  return false;
}
