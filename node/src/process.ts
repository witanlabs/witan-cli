import { spawn, type ChildProcess } from 'node:child_process';
import { createInterface } from 'node:readline';
import { WitanProcessError, WitanRPCError, WitanTimeoutError } from './errors.js';

interface RPCRequest {
  id: string;
  op: string;
  args: Record<string, unknown>;
}

interface RPCResponse {
  id?: string;
  ok: boolean;
  result?: unknown;
  code?: string;
  message?: string;
}

export class StdioRPCProcess {
  private proc: ChildProcess;
  private responseQueue: Array<(line: string | null) => void> = [];
  private stderrBuffer: string[] = [];
  private closed = false;
  private readonly timeout: number;
  private readonly readyPromise: Promise<void>;
  private readyReject?: (err: Error) => void;

  // Mutex: serializes requests to prevent interleaved writes and response misattribution
  private requestChain: Promise<void> = Promise.resolve();

  constructor(argv: string[], options: { env?: Record<string, string>; timeout?: number } = {}) {
    this.timeout = options.timeout ?? 90_000;

    this.proc = spawn(argv[0]!, argv.slice(1), {
      stdio: ['pipe', 'pipe', 'pipe'],
      env: { ...process.env, ...options.env },
    });

    if (!this.proc.stdout || !this.proc.stdin || !this.proc.stderr) {
      this.terminate();
      throw new WitanProcessError('Failed to create stdio pipes');
    }

    // Read stdout line by line
    const rl = createInterface({ input: this.proc.stdout });
    rl.on('line', (line) => {
      const resolver = this.responseQueue.shift();
      if (resolver) resolver(line);
    });
    rl.on('close', () => {
      // Signal EOF to any waiting requests
      while (this.responseQueue.length) {
        const resolver = this.responseQueue.shift();
        if (resolver) resolver(null);
      }
    });

    // Capture stderr (keep last 50 lines)
    const stderrRl = createInterface({ input: this.proc.stderr });
    stderrRl.on('line', (line) => {
      this.stderrBuffer.push(line);
      if (this.stderrBuffer.length > 50) {
        this.stderrBuffer.shift();
      }
    });

    // Create a promise that resolves when subprocess is ready or rejects on early exit
    this.readyPromise = new Promise((resolve, reject) => {
      // Store reject so terminate() can reject if called during startup
      this.readyReject = reject;

      // Handle spawn errors (e.g., binary not found)
      this.proc.once('error', (err) => {
        this.closed = true;
        if (this.readyReject) {
          this.readyReject(new WitanProcessError(`witan subprocess failed to start: ${err.message}`, this.stderrBuffer));
          this.readyReject = undefined;
        }
      });

      // If process exits immediately, it's an error
      this.proc.once('exit', (code) => {
        if (!this.closed) {
          reject(new WitanProcessError(
            `witan subprocess exited during startup (exit=${code})`,
            this.stderrBuffer
          ));
        }
      });

      // Give the process a moment to fail, then assume it's ready
      // (The Go RPC process doesn't send a "ready" signal)
      setTimeout(() => {
        if (this.proc.exitCode === null && !this.closed) {
          this.readyReject = undefined; // No longer pending
          resolve();
        }
      }, 200);
    });
  }

  /**
   * Wait for the subprocess to be ready.
   * Throws if the subprocess exits during startup.
   */
  async waitReady(): Promise<void> {
    await this.readyPromise;
  }

  async request(method: string, op: string, args: Record<string, unknown>, requestId: string): Promise<unknown> {
    // Serialize requests to prevent interleaved writes and response misattribution.
    // This matches Python's asyncio.Lock() approach in _async_process.py.
    const resultPromise = this.requestChain.catch(() => {}).then(async () => {
      if (this.closed) {
        throw new WitanProcessError('witan subprocess is closed', this.stderrBuffer);
      }

      const payload: RPCRequest = { id: requestId, op, args };
      const line = JSON.stringify(payload) + '\n';

      // Write request
      await new Promise<void>((resolve, reject) => {
        this.proc.stdin!.write(line, (err) => {
          if (err) reject(new WitanProcessError(`Writing RPC request: ${err.message}`, this.stderrBuffer));
          else resolve();
        });
      });

      // Wait for response with timeout
      const responseLine = await new Promise<string | null>((resolve, reject) => {
        const timer = setTimeout(() => {
          this.terminate();
          reject(new WitanTimeoutError(
            `RPC timeout: ${method} (${op}) did not respond within ${this.timeout}ms`,
            this.stderrBuffer
          ));
        }, this.timeout);

        this.responseQueue.push((respLine) => {
          clearTimeout(timer);
          resolve(respLine);
        });
      });

      if (responseLine === null) {
        const code = this.proc.exitCode;
        throw new WitanProcessError(
          `witan subprocess exited before responding (exit=${code})`,
          this.stderrBuffer
        );
      }

      let response: RPCResponse;
      try {
        response = JSON.parse(responseLine) as RPCResponse;
      } catch {
        this.terminate();
        throw new WitanProcessError(`Invalid JSON RPC response: ${responseLine}`, this.stderrBuffer);
      }

      if (response.id && response.id !== requestId) {
        this.terminate();
        throw new WitanProcessError(
          `RPC response id mismatch: expected ${requestId}, got ${response.id}`,
          this.stderrBuffer
        );
      }

      if (response.ok === false) {
        throw new WitanRPCError(response.message ?? 'RPC request failed', {
          method,
          op,
          requestId,
          code: response.code,
          response: response as unknown as Record<string, unknown>,
        });
      }

      if (response.ok !== true) {
        this.terminate();
        throw new WitanProcessError(`Invalid RPC ok field: ${JSON.stringify(response)}`, this.stderrBuffer);
      }

      return response.result;
    });

    // Update chain for next request (don't await - just track for serialization)
    this.requestChain = resultPromise.then(() => {}, () => {});

    return resultPromise;
  }

  async close(): Promise<void> {
    if (this.closed) return;
    this.closed = true;

    this.proc.stdin?.end();

    await new Promise<void>((resolve) => {
      const timer = setTimeout(() => {
        this.terminate();
        resolve();
      }, 5000);

      this.proc.on('exit', () => {
        clearTimeout(timer);
        resolve();
      });
    });
  }

  terminate(): void {
    this.closed = true;

    // Reject readyPromise if still pending (prevents hang if terminate() called during startup)
    if (this.readyReject) {
      this.readyReject(new WitanProcessError('witan subprocess terminated during startup', this.stderrBuffer));
      this.readyReject = undefined;
    }

    if (this.proc.exitCode === null) {
      this.proc.kill('SIGTERM');
      setTimeout(() => {
        if (this.proc.exitCode === null) {
          this.proc.kill('SIGKILL');
        }
      }, 2000);
    }
  }

  get isClosed(): boolean {
    return this.closed;
  }
}
