import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));

/**
 * Returns environment variables for the fake witan RPC server.
 *
 * @param tmpDir - Temporary directory for test artifacts
 * @param mode - Fake server mode: 'ok', 'hang', 'rpc-error', 'invalid-json', 'wrong-id', 'exit'
 */
export function fakeEnv(tmpDir: string, mode = 'ok') {
  return {
    WITAN_FAKE_ARGV_FILE: join(tmpDir, 'argv.jsonl'),
    WITAN_FAKE_REQUESTS_FILE: join(tmpDir, 'requests.jsonl'),
    WITAN_FAKE_SAVE_FILE: join(tmpDir, 'saved.txt'),
    WITAN_FAKE_MODE: mode,
  };
}

/**
 * Returns the path to a wrapper script that invokes the Python fake server.
 * The wrapper is needed because Workbook.open() expects a single binary path.
 */
export function fakeBinaryPath(): string {
  return join(__dirname, 'fake-witan.sh');
}
