import { describe, it, expect } from 'vitest';
import { fakeEnv, fakeBinaryPath } from './helpers.js';
import { existsSync } from 'node:fs';

describe('test setup', () => {
  it('fakeEnv returns expected environment variables', () => {
    const env = fakeEnv('/tmp/test');
    expect(env.WITAN_FAKE_MODE).toBe('ok');
    expect(env.WITAN_FAKE_ARGV_FILE).toBe('/tmp/test/argv.jsonl');
    expect(env.WITAN_FAKE_REQUESTS_FILE).toBe('/tmp/test/requests.jsonl');
    expect(env.WITAN_FAKE_SAVE_FILE).toBe('/tmp/test/saved.txt');
  });

  it('fakeEnv accepts custom mode', () => {
    const env = fakeEnv('/tmp/test', 'hang');
    expect(env.WITAN_FAKE_MODE).toBe('hang');
  });

  it('fakeBinaryPath returns path to fake-witan.sh', () => {
    const path = fakeBinaryPath();
    expect(path).toContain('fake-witan.sh');
    expect(existsSync(path)).toBe(true);
  });
});
