import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { mkdtemp, rm, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Workbook } from '../src/workbook.js';
import { WitanProcessError, WitanRPCError, WitanTimeoutError } from '../src/errors.js';

const FAKE_WITAN_PATH = join(__dirname, 'fake-witan.sh');

function fakeEnv(tmpDir: string, mode = 'ok') {
  return {
    WITAN_FAKE_ARGV_FILE: join(tmpDir, 'argv.jsonl'),
    WITAN_FAKE_REQUESTS_FILE: join(tmpDir, 'requests.jsonl'),
    WITAN_FAKE_SAVE_FILE: join(tmpDir, 'saved.txt'),
    WITAN_FAKE_MODE: mode,
  };
}

async function readRequests(path: string): Promise<Array<{ id: string; op: string; args: Record<string, unknown> }>> {
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

describe('Workbook', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await mkdtemp(join(tmpdir(), 'witan-test-'));
  });

  afterEach(async () => {
    await rm(tmpDir, { recursive: true });
  });

  describe('lifecycle', () => {
    it('opens and closes via static factory', async () => {
      const env = fakeEnv(tmpDir);
      const wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      expect(wb.isClosed).toBe(false);
      await wb.close();
      expect(wb.isClosed).toBe(true);
    });

    it('works with await using syntax', async () => {
      const env = fakeEnv(tmpDir);
      let workbookRef: Workbook | null = null;

      {
        await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
          binary: FAKE_WITAN_PATH,
          env,
        });
        workbookRef = wb;
        expect(wb.isClosed).toBe(false);
      }

      // After block, workbook should be closed
      expect(workbookRef!.isClosed).toBe(true);
    });

    it('close is idempotent', async () => {
      const env = fakeEnv(tmpDir);
      const wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.close();
      await wb.close(); // Should not throw
      expect(wb.isClosed).toBe(true);
    });

    it('throws when calling methods after close', async () => {
      const env = fakeEnv(tmpDir);
      const wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.close();
      await expect(wb.listSheets()).rejects.toThrow(WitanProcessError);
    });

    it('passes create flag correctly', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'new.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
        create: true,
      });

      // Check argv was logged correctly
      const argvContent = await readFile(env.WITAN_FAKE_ARGV_FILE, 'utf-8');
      const argv = JSON.parse(argvContent);
      expect(argv).toContain('--create');
    });

    it('passes locale and hint flags', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
        locale: 'de-DE',
        hint: 'excel',
      });

      const argvContent = await readFile(env.WITAN_FAKE_ARGV_FILE, 'utf-8');
      const argv = JSON.parse(argvContent);
      expect(argv).toContain('--locale');
      expect(argv).toContain('de-DE');
      expect(argv).toContain('--hint');
      expect(argv).toContain('excel');
    });
  });

  describe('save', () => {
    it('saves the workbook', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const result = await wb.save();
      expect(result).toBe(true);

      // Verify save marker was written by fake server
      const saved = await readFile(env.WITAN_FAKE_SAVE_FILE, 'utf-8');
      expect(saved).toBe('saved\n');
    });
  });

  describe('sheet operations', () => {
    it('lists sheets', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const sheets = await wb.listSheets();
      expect(sheets).toHaveLength(2);
      expect(sheets[0].sheet).toBe('Sheet1');
      expect(sheets[1].sheet).toBe('Hidden');
      expect(sheets[1].hidden).toBe(true);
    });

    it('adds a sheet', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const name = await wb.addSheet('NewSheet');
      expect(name).toBe('NewSheet');

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const addReq = requests.find((r) => r.op === 'addSheet');
      expect(addReq).toBeDefined();
      expect(addReq!.args.name).toBe('NewSheet');
    });

    it('deletes a sheet', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.deleteSheet('Sheet1');

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const delReq = requests.find((r) => r.op === 'deleteSheet');
      expect(delReq).toBeDefined();
      expect(delReq!.args.name).toBe('Sheet1');
    });

    it('renames a sheet', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.renameSheet('Sheet1', 'RenamedSheet');

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const renReq = requests.find((r) => r.op === 'renameSheet');
      expect(renReq).toBeDefined();
      expect(renReq!.args.oldName).toBe('Sheet1');
      expect(renReq!.args.newName).toBe('RenamedSheet');
    });
  });

  describe('workbook properties', () => {
    it('gets workbook properties', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const props = await wb.getWorkbookProperties();
      expect(props.activeSheetIndex).toBe(0);
      expect(props.defaultFont).toEqual({ name: 'Calibri', size: 11 });
    });

    it('sets workbook properties', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.setWorkbookProperties({ activeSheetIndex: 1 });

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const setReq = requests.find((r) => r.op === 'setWorkbookProperties');
      expect(setReq).toBeDefined();
      expect(setReq!.args.activeSheetIndex).toBe(1);
    });
  });

  describe('sheet properties', () => {
    it('gets sheet properties with defaults', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const props = await wb.getSheetProperties('Sheet1');
      expect(props.visibility).toBe('visible');
      // Verify defaults are set
      expect(props.columns).toEqual({});
      expect(props.rows).toEqual({});
    });

    it('gets sheet properties with filter', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.getSheetProperties('Sheet1', { columns: ['A', 'B'], rows: [1, 2, 3] });

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const getReq = requests.find((r) => r.op === 'getSheetProperties');
      expect(getReq).toBeDefined();
      expect(getReq!.args.sheet).toBe('Sheet1');
      expect(getReq!.args.filter).toEqual({ columns: ['A', 'B'], rows: [1, 2, 3] });
    });

    it('sets sheet properties', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.setSheetProperties('Sheet1', { visibility: 'hidden' });

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const setReq = requests.find((r) => r.op === 'setSheetProperties');
      expect(setReq).toBeDefined();
      expect(setReq!.args.sheet).toBe('Sheet1');
      expect(setReq!.args.properties).toEqual({ visibility: 'hidden' });
    });
  });

  describe('error handling', () => {
    it('throws WitanRPCError on RPC error', async () => {
      const env = fakeEnv(tmpDir, 'rpc-error');
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await expect(wb.listSheets()).rejects.toThrow(WitanRPCError);
    });

    it('throws WitanTimeoutError on timeout', async () => {
      const env = fakeEnv(tmpDir, 'hang');
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
        requestTimeout: 500,
      });

      await expect(wb.listSheets()).rejects.toThrow(WitanTimeoutError);
    }, 10000);
  });

  describe('request ID generation', () => {
    it('generates incrementing request IDs', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.listSheets();
      await wb.save();
      await wb.listSheets();

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      expect(requests.map((r) => r.id)).toEqual(['1', '2', '3']);
    });
  });
});
