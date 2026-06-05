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

/** Return true when `err` indicates Google Sheets authorization is required. */
export function isGoogleAuthRequired(err: unknown): boolean {
  return err instanceof WitanRPCError && err.code === 'google_auth_required';
}
