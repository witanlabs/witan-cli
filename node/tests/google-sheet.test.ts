import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { mkdtemp, rm, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import {
  GoogleSheet,
  DEFAULT_GOOGLE_SHEET_REQUEST_TIMEOUT_MS,
} from '../src/google-sheet.js';
import {
  WitanRPCError,
  isGoogleAuthRequired,
} from '../src/errors.js';

const FAKE_WITAN_PATH = join(__dirname, 'fake-witan.sh');

function fakeEnv(tmpDir: string, mode = 'ok') {
  return {
    WITAN_FAKE_ARGV_FILE: join(tmpDir, 'argv.jsonl'),
    WITAN_FAKE_REQUESTS_FILE: join(tmpDir, 'requests.jsonl'),
    WITAN_FAKE_MODE: mode,
  };
}

async function readArgv(path: string): Promise<string[]> {
  const content = await readFile(path, 'utf-8');
  return JSON.parse(content.trim()) as string[];
}

async function readRequests(path: string): Promise<Array<{ op: string }>> {
  try {
    const content = await readFile(path, 'utf-8');
    return content
      .trim()
      .split('\n')
      .filter(Boolean)
      .map((line) => JSON.parse(line));
  } catch {
    return [];
  }
}

describe('GoogleSheet', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await mkdtemp(join(tmpdir(), 'witan-gsheet-test-'));
  });

  afterEach(async () => {
    await rm(tmpDir, { recursive: true });
  });

  it('open uses gsheets rpc', async () => {
    const env = fakeEnv(tmpDir);

    await using sheet = await GoogleSheet.open('gs://sheet-123', {
      binary: FAKE_WITAN_PATH,
      env,
    });

    const sheets = await sheet.listSheets();
    expect(sheets[0]?.sheet).toBe('Sheet1');

    const data = await sheet.readRange('Sheet1!A1:B2');
    expect(data[0]?.[0]?.value).toBe(2);

    const writeResult = await sheet.setCells([{ address: 'Sheet1!A1', value: 'done' }]);
    expect(writeResult.changed).toEqual(['Sheet1!A1']);

    expect('save' in sheet).toBe(false);

    const argv = await readArgv(env.WITAN_FAKE_ARGV_FILE);
    expect(argv).toEqual(['gsheets', 'rpc', 'gs://sheet-123']);

    const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
    expect(requests.map((request) => request.op)).toEqual(['listSheets', 'readRange', 'setCells']);
  });

  it('create uses single rpc with create flag', async () => {
    const env = fakeEnv(tmpDir);

    await using sheet = await GoogleSheet.create('Budget 2026', {
      binary: FAKE_WITAN_PATH,
      env,
    });

    expect(sheet.isCreate).toBe(true);
    expect((await sheet.readCell('Sheet1!A1')).value).toBe(2);

    const argv = await readArgv(env.WITAN_FAKE_ARGV_FILE);
    expect(argv).toEqual(['gsheets', 'rpc', '--create', '--title', 'Budget 2026']);

    const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
    expect(requests[0]?.op).toBe('readRange');
  });

  it('create via new ref', async () => {
    const env = fakeEnv(tmpDir);

    await using sheet = await GoogleSheet.open('new', {
      title: 'Q1',
      create: true,
      binary: FAKE_WITAN_PATH,
      env,
    });

    expect(sheet.isCreate).toBe(true);
    await sheet.readRange('Sheet1!A1:B2');

    const argv = await readArgv(env.WITAN_FAKE_ARGV_FILE);
    expect(argv).toEqual(['gsheets', 'rpc', '--create', '--title', 'Q1', 'new']);

    const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
    expect(requests.length).toBeGreaterThan(0);
  });

  it('uses default request timeout', async () => {
    const env = fakeEnv(tmpDir);
    const sheet = await GoogleSheet.open('gs://sheet-123', {
      binary: FAKE_WITAN_PATH,
      env,
    });

    try {
      expect(sheet.requestTimeoutMs).toBe(DEFAULT_GOOGLE_SHEET_REQUEST_TIMEOUT_MS);
    } finally {
      await sheet.close();
    }
  });

  it('rejects xlsx-only options', async () => {
    await expect(GoogleSheet.open('gs://id', { apiKey: 'secret' })).rejects.toThrow(/apiKey/);
    await expect(GoogleSheet.open('gs://id', { stateless: true })).rejects.toThrow(/stateless/);
    await expect(GoogleSheet.open('gs://id', { create: true })).rejects.toThrow(/create requires/);
    await expect(GoogleSheet.open('gs://id', { hint: 'Sheet1!A1' })).rejects.toThrow(/hint/);
    await expect(GoogleSheet.create(undefined, { apiKey: 'secret' })).rejects.toThrow(/apiKey/);
  });

  it('surfaces rpc error code', async () => {
    const env = fakeEnv(tmpDir, 'rpc-error');
    const sheet = await GoogleSheet.open('gs://sheet-123', {
      binary: FAKE_WITAN_PATH,
      env,
    });

    try {
      await expect(sheet.listSheets()).rejects.toSatisfy((err: unknown) => {
        expect(err).toBeInstanceOf(WitanRPCError);
        expect((err as WitanRPCError).code).toBe('BOOM');
        expect(isGoogleAuthRequired(err)).toBe(false);
        return true;
      });
    } finally {
      await sheet.close();
    }
  });
});
