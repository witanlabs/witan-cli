import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { join } from 'node:path';
import { writeFileSync, readFileSync, unlinkSync, existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { randomUUID } from 'node:crypto';
import { StdioRPCProcess } from '../src/process.js';
import { WitanProcessError, WitanRPCError, WitanTimeoutError } from '../src/errors.js';

const FAKE_RPC_PATH = join(__dirname, '../../python/tests/fake_witan_rpc.py');

function createProcess(mode = 'ok', options: { timeout?: number } = {}) {
  return new StdioRPCProcess(['python3', FAKE_RPC_PATH], {
    env: { WITAN_FAKE_MODE: mode },
    timeout: options.timeout,
  });
}

describe('StdioRPCProcess', () => {
  let proc: StdioRPCProcess | null = null;

  afterEach(async () => {
    if (proc && !proc.isClosed) {
      await proc.close();
    }
    proc = null;
  });

  describe('lifecycle', () => {
    it('starts and becomes ready', async () => {
      proc = createProcess();
      await proc.waitReady();
      expect(proc.isClosed).toBe(false);
    });

    it('closes gracefully', async () => {
      proc = createProcess();
      await proc.waitReady();
      await proc.close();
      expect(proc.isClosed).toBe(true);
    });

    it('terminate() forces shutdown', async () => {
      proc = createProcess();
      await proc.waitReady();
      proc.terminate();
      expect(proc.isClosed).toBe(true);
    });

    it('close() is idempotent', async () => {
      proc = createProcess();
      await proc.waitReady();
      await proc.close();
      await proc.close(); // Should not throw
      expect(proc.isClosed).toBe(true);
    });
  });

  describe('request/response', () => {
    it('sends request and receives response', async () => {
      proc = createProcess();
      await proc.waitReady();

      const result = await proc.request('test', 'listSheets', {}, 'req-1');
      expect(result).toEqual({
        sheets: [
          { sheet: 'Sheet1', address: 'Sheet1!A1:B2', rows: 2, cols: 2 },
          { sheet: 'Hidden', address: 'Hidden!A1:A1', rows: 1, cols: 1, hidden: true },
        ],
      });
    });

    it('handles UTF-8 responses', async () => {
      proc = createProcess('utf8');
      await proc.waitReady();

      const result = await proc.request('test', 'utf8', {}, 'req-1');
      expect(result).toEqual({ text: 'Café 📈 東京' });
    });

    it('passes arguments correctly', async () => {
      const requestsFile = join(tmpdir(), `witan-test-${randomUUID()}.json`);
      proc = new StdioRPCProcess(['python3', FAKE_RPC_PATH], {
        env: { WITAN_FAKE_MODE: 'ok', WITAN_FAKE_REQUESTS_FILE: requestsFile },
      });
      await proc.waitReady();

      await proc.request('test', 'readRange', { address: 'Sheet1!A1:B2' }, 'req-args');
      await proc.close();

      // Read logged requests
      const logged = readFileSync(requestsFile, 'utf-8')
        .trim()
        .split('\n')
        .map((line) => JSON.parse(line));

      expect(logged).toHaveLength(1);
      expect(logged[0]).toEqual({
        id: 'req-args',
        op: 'readRange',
        args: { address: 'Sheet1!A1:B2' },
      });

      // Cleanup
      unlinkSync(requestsFile);
    });
  });

  describe('error handling', () => {
    it('throws WitanRPCError on RPC error response', async () => {
      proc = createProcess('rpc-error');
      await proc.waitReady();

      await expect(proc.request('test', 'anyOp', {}, 'req-err')).rejects.toThrow(WitanRPCError);

      try {
        await proc.request('test', 'anyOp', {}, 'req-err2');
      } catch (err) {
        expect(err).toBeInstanceOf(WitanRPCError);
        const rpcErr = err as WitanRPCError;
        expect(rpcErr.code).toBe('BOOM');
        expect(rpcErr.message).toBe('BOOM: boom'); // Message includes code prefix
        expect(rpcErr.method).toBe('test');
        expect(rpcErr.op).toBe('anyOp');
      }
    });

    it('throws WitanProcessError on invalid JSON response', async () => {
      proc = createProcess('invalid-json');
      await proc.waitReady();

      await expect(proc.request('test', 'anyOp', {}, 'req-json')).rejects.toThrow(WitanProcessError);
      expect(proc.isClosed).toBe(true);
    });

    it('throws WitanProcessError on response ID mismatch', async () => {
      proc = createProcess('wrong-id');
      await proc.waitReady();

      await expect(proc.request('test', 'anyOp', {}, 'req-id')).rejects.toThrow(WitanProcessError);
      expect(proc.isClosed).toBe(true);
    });

    it('throws WitanProcessError when subprocess exits early', async () => {
      proc = createProcess('exit');
      await proc.waitReady();

      await expect(proc.request('test', 'anyOp', {}, 'req-exit')).rejects.toThrow(WitanProcessError);
    });

    it('throws WitanProcessError when closed', async () => {
      proc = createProcess();
      await proc.waitReady();
      await proc.close();

      await expect(proc.request('test', 'anyOp', {}, 'req-closed')).rejects.toThrow(WitanProcessError);
    });
  });

  describe('timeout', () => {
    it('throws WitanTimeoutError when response times out', async () => {
      proc = createProcess('hang', { timeout: 500 });
      await proc.waitReady();

      const start = Date.now();
      await expect(proc.request('test', 'anyOp', {}, 'req-timeout')).rejects.toThrow(
        WitanTimeoutError
      );
      const elapsed = Date.now() - start;

      expect(elapsed).toBeGreaterThanOrEqual(450);
      expect(elapsed).toBeLessThan(2000);
      expect(proc.isClosed).toBe(true);
    });
  });

  describe('request serialization', () => {
    it('serializes concurrent requests to prevent interleaving', async () => {
      const requestsFile = join(tmpdir(), `witan-test-${randomUUID()}.json`);
      proc = new StdioRPCProcess(['python3', FAKE_RPC_PATH], {
        env: { WITAN_FAKE_MODE: 'ok', WITAN_FAKE_REQUESTS_FILE: requestsFile },
      });
      await proc.waitReady();

      // Fire multiple requests concurrently
      const promises = [
        proc.request('test', 'listSheets', {}, 'req-a'),
        proc.request('test', 'readRange', { address: 'A1' }, 'req-b'),
        proc.request('test', 'listSheets', {}, 'req-c'),
      ];

      const results = await Promise.all(promises);
      expect(results).toHaveLength(3);

      await proc.close();

      // Verify requests were sent sequentially (not interleaved)
      const logged = readFileSync(requestsFile, 'utf-8')
        .trim()
        .split('\n')
        .map((line) => JSON.parse(line));

      expect(logged).toHaveLength(3);
      // Each request should be complete before the next one
      const ids = logged.map((r: { id: string }) => r.id);
      expect(ids).toEqual(['req-a', 'req-b', 'req-c']);

      unlinkSync(requestsFile);
    });

    it('continues after one request fails', async () => {
      // Create a process that will fail on first request then succeed
      const requestsFile = join(tmpdir(), `witan-test-${randomUUID()}.json`);

      // First, test that failed requests don't break the chain
      proc = createProcess('ok');
      await proc.waitReady();

      // First request succeeds
      const result1 = await proc.request('test', 'listSheets', {}, 'req-1');
      expect(result1).toBeDefined();

      // Close and create a new one to test error recovery
      await proc.close();

      // With rpc-error mode, requests fail but process continues
      proc = createProcess('rpc-error');
      await proc.waitReady();

      // First request fails
      await expect(proc.request('test', 'op1', {}, 'req-fail1')).rejects.toThrow(WitanRPCError);

      // Second request also fails (but the chain continues)
      await expect(proc.request('test', 'op2', {}, 'req-fail2')).rejects.toThrow(WitanRPCError);

      // Process is still open (rpc errors don't terminate)
      expect(proc.isClosed).toBe(false);
    });
  });

  describe('startup edge cases', () => {
    it('rejects waitReady if process fails to start', async () => {
      proc = new StdioRPCProcess(['nonexistent-binary-xyz'], { timeout: 1000 });
      await expect(proc.waitReady()).rejects.toThrow(WitanProcessError);
      expect(proc.isClosed).toBe(true);
    }, 5000); // Short timeout for this test

    it('terminate during startup rejects waitReady', async () => {
      proc = new StdioRPCProcess(['python3', FAKE_RPC_PATH], {
        env: { WITAN_FAKE_MODE: 'ok' },
      });

      // Immediately terminate before ready
      proc.terminate();

      await expect(proc.waitReady()).rejects.toThrow(WitanProcessError);
    });
  });
});
