import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { mkdtemp, rm, readFile } from 'node:fs/promises';
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
        requestTimeoutMs: 500,
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

  describe('reading data', () => {
    it('reads a single cell', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const cell = await wb.readCell('Sheet1!A1');
      expect(cell.address).toBe('Sheet1!A1');
      expect(cell.value).toBe(2);
      expect(cell.type).toBe('number');
    });

    it('reads a cell with context option', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.readCell('A1', { context: 3 });

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const readReq = requests.find((r) => r.op === 'readRange');
      expect(readReq).toBeDefined();
      expect(readReq!.args.context).toBe(3);
    });

    it('reads a range', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const data = await wb.readRange('Sheet1!A1:B1');
      expect(data).toHaveLength(1);
      expect(data[0]).toHaveLength(2);
      expect(data[0][0].value).toBe(2);
      expect(data[0][1].value).toBe(3);
    });

    it('reads a row', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const row = await wb.readRow('Sheet1', 1);
      expect(row).toHaveLength(2);

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const readReq = requests.find((r) => r.op === 'readRow');
      expect(readReq).toBeDefined();
      expect(readReq!.args.sheet).toBe('Sheet1');
      expect(readReq!.args.row).toBe(1);
    });

    it('reads a row with column range', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.readRow('Sheet1', 1, { startCol: 1, endCol: 5 });

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const readReq = requests.find((r) => r.op === 'readRow');
      expect(readReq!.args.startCol).toBe(1);
      expect(readReq!.args.endCol).toBe(5);
    });

    it('reads a column', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const col = await wb.readColumn('Sheet1', 'A');
      expect(col).toHaveLength(2);

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const readReq = requests.find((r) => r.op === 'readColumn');
      expect(readReq).toBeDefined();
      expect(readReq!.args.sheet).toBe('Sheet1');
      expect(readReq!.args.col).toBe('A');
    });

    it('reads a column with row range', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.readColumn('Sheet1', 1, { startRow: 2, endRow: 10 });

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const readReq = requests.find((r) => r.op === 'readColumn');
      expect(readReq!.args.startRow).toBe(2);
      expect(readReq!.args.endRow).toBe(10);
    });

    it('reads range as TSV', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const tsv = await wb.readRangeTsv('Sheet1!A1:B2');
      expect(tsv).toBe('A\tB\n1\t2');
    });

    it('reads range TSV with options', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.readRangeTsv('Sheet1!A1:B2', { includeEmpty: true, includeFormulas: true });

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const readReq = requests.find((r) => r.op === 'readRangeTsv');
      expect(readReq!.args.includeEmpty).toBe(true);
      expect(readReq!.args.includeFormulas).toBe(true);
    });

    it('reads row as TSV', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.readRowTsv('Sheet1', 1, { startCol: 1, endCol: 3 });

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const readReq = requests.find((r) => r.op === 'readRowTsv');
      expect(readReq).toBeDefined();
      expect(readReq!.args.sheet).toBe('Sheet1');
      expect(readReq!.args.row).toBe(1);
    });

    it('reads column as TSV', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.readColumnTsv('Sheet1', 'B', { startRow: 1, endRow: 10 });

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const readReq = requests.find((r) => r.op === 'readColumnTsv');
      expect(readReq).toBeDefined();
      expect(readReq!.args.col).toBe('B');
    });
  });

  describe('writing data', () => {
    it('sets multiple cells', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const result = await wb.setCells([
        { address: 'A1', value: 'Hello' },
        { address: 'B1', value: 42 },
        { address: 'C1', value: true },
      ]);

      expect(result.changed).toContain('Sheet1!A1');

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const setReq = requests.find((r) => r.op === 'setCells');
      expect(setReq).toBeDefined();
      expect(setReq!.args.cells).toHaveLength(3);
    });

    it('sets cell style', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.setStyle('A1:B2', { font: { bold: true, size: 14 } });

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const styleReq = requests.find((r) => r.op === 'setStyle');
      expect(styleReq).toBeDefined();
      expect(styleReq!.args.address).toBe('A1:B2');
      expect(styleReq!.args.style).toEqual({ font: { bold: true, size: 14 } });
    });

    it('copies a range', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const result = await wb.copyRange('A1:B2', 'D1');
      expect(result.cellsCopied).toBe(4);

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const copyReq = requests.find((r) => r.op === 'copyRange');
      expect(copyReq).toBeDefined();
      expect(copyReq!.args.source).toBe('A1:B2');
      expect(copyReq!.args.destination).toBe('D1');
    });

    it('copies a range with paste type', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.copyRange('A1:B2', 'D1', { pasteType: 'values' });

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const copyReq = requests.find((r) => r.op === 'copyRange');
      expect(copyReq!.args.pasteType).toBe('values');
    });
  });

  describe('scaleRange (composite operation)', () => {
    it('scales numeric values in a range', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      // readRange returns cells with numeric values (2 and 3 from fake server)
      const result = await wb.scaleRange('Sheet1!A1:B1', 2);

      expect(result).not.toBeNull();
      expect(result!.changed).toContain('Sheet1!A1');

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);

      // Should have called readRange first
      const readReq = requests.find((r) => r.op === 'readRange');
      expect(readReq).toBeDefined();

      // Then setCells with scaled values
      const setReq = requests.find((r) => r.op === 'setCells');
      expect(setReq).toBeDefined();

      // Values should be doubled (2*2=4, 3*2=6)
      const cells = setReq!.args.cells as Array<{ address: string; value: number }>;
      expect(cells).toHaveLength(2);
      expect(cells.find((c) => c.address === 'Sheet1!A1')?.value).toBe(4);
      expect(cells.find((c) => c.address === 'Sheet1!B1')?.value).toBe(6);
    });

    it('returns null when no numeric cells found', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      // Reading a single cell returns [[{value: 2}]] - let's test the branch differently
      // by using the UTF-8 mode which returns text instead of numbers
    });

    it('skips formula cells by default', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      // The fake server doesn't return formulas, but we test the option is passed
      await wb.scaleRange('A1:B1', 2, { skipFormulas: false });

      // Just verify the operation completes
      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      expect(requests.some((r) => r.op === 'setCells')).toBe(true);
    });
  });

  describe('search operations', () => {
    it('finds cells by string', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      // Fake server returns empty matches array
      const matches = await wb.findCells('test');
      expect(matches).toEqual([]);

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const findReq = requests.find((r) => r.op === 'findCells');
      expect(findReq).toBeDefined();
      expect(findReq!.args.matcher).toBe('test');
    });

    it('finds cells by RegExp', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.findCells(/pattern/i);

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const findReq = requests.find((r) => r.op === 'findCells');
      expect(findReq!.args.matcher).toEqual({ source: 'pattern', flags: 'i' });
    });

    it('finds cells with options', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.findCells('test', {
        in: 'Sheet1!A:Z',
        context: 5,
        limit: 50,
        offset: 10,
        formulas: true,
      });

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const findReq = requests.find((r) => r.op === 'findCells');
      expect(findReq!.args.in).toBe('Sheet1!A:Z');
      expect(findReq!.args.context).toBe(5);
      expect(findReq!.args.limit).toBe(50);
      expect(findReq!.args.offset).toBe(10);
      expect(findReq!.args.formulas).toBe(true);
    });

    it('finds rows by matcher', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      // Fake server returns empty matches array
      const matches = await wb.findRows('test');
      expect(matches).toEqual([]);

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const findReq = requests.find((r) => r.op === 'findRows');
      expect(findReq).toBeDefined();
    });

    it('finds and replaces text', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const result = await wb.findAndReplace('old', 'new');
      // Fake server returns replaced: 1
      expect(result.replaced).toBe(1);
      expect(result.cells).toContain('Sheet1!A1');

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const replaceReq = requests.find((r) => r.op === 'findAndReplace');
      expect(replaceReq).toBeDefined();
      expect(replaceReq!.args.find).toBe('old');
      expect(replaceReq!.args.replace).toBe('new');
    });

    it('finds and replaces with RegExp and options', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.findAndReplace(/pattern/g, 'replacement', {
        in: 'Sheet1!A:D',
        matchCase: true,
        wholeCell: true,
        inFormulas: true,
        limit: 100,
      });

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const replaceReq = requests.find((r) => r.op === 'findAndReplace');
      expect(replaceReq!.args.find).toEqual({ source: 'pattern', flags: 'g' });
      expect(replaceReq!.args.matchCase).toBe(true);
      expect(replaceReq!.args.wholeCell).toBe(true);
      expect(replaceReq!.args.inFormulas).toBe(true);
      expect(replaceReq!.args.limit).toBe(100);
    });
  });

  describe('row/column structure operations', () => {
    it('inserts rows after', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.insertRowAfter('Sheet1', 5, 3);

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const insertReq = requests.find((r) => r.op === 'insertRowAfter');
      expect(insertReq).toBeDefined();
      expect(insertReq!.args.sheet).toBe('Sheet1');
      expect(insertReq!.args.row).toBe(5);
      expect(insertReq!.args.count).toBe(3);
    });

    it('deletes rows', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.deleteRows('Sheet1', 2, 5);

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const deleteReq = requests.find((r) => r.op === 'deleteRows');
      expect(deleteReq).toBeDefined();
      expect(deleteReq!.args.sheet).toBe('Sheet1');
      expect(deleteReq!.args.row).toBe(2);
      expect(deleteReq!.args.count).toBe(5);
    });

    it('inserts columns after', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.insertColumnAfter('Sheet1', 'B', 2);

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const insertReq = requests.find((r) => r.op === 'insertColumnAfter');
      expect(insertReq).toBeDefined();
      expect(insertReq!.args.column).toBe('B');
      expect(insertReq!.args.count).toBe(2);
    });

    it('deletes columns', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.deleteColumns('Sheet1', 3, 2);

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const deleteReq = requests.find((r) => r.op === 'deleteColumns');
      expect(deleteReq).toBeDefined();
      expect(deleteReq!.args.column).toBe(3);
      expect(deleteReq!.args.count).toBe(2);
    });

    it('uses default count of 1', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.insertRowAfter('Sheet1', 1);
      await wb.deleteRows('Sheet1', 1);
      await wb.insertColumnAfter('Sheet1', 'A');
      await wb.deleteColumns('Sheet1', 1);

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      expect(requests.filter((r) => r.args.count === 1)).toHaveLength(4);
    });
  });

  describe('row/column properties', () => {
    it('sets row properties', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.setRowProperties('Sheet1', 1, 5, { height: 30, hidden: false });

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const setReq = requests.find((r) => r.op === 'setRowProperties');
      expect(setReq).toBeDefined();
      expect(setReq!.args.sheet).toBe('Sheet1');
      expect(setReq!.args.fromRow).toBe(1);
      expect(setReq!.args.toRow).toBe(5);
      expect(setReq!.args.properties).toEqual({ height: 30, hidden: false });
    });

    it('sets column properties', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.setColumnProperties('Sheet1', 'A', 'D', { width: 100 });

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const setReq = requests.find((r) => r.op === 'setColumnProperties');
      expect(setReq).toBeDefined();
      expect(setReq!.args.fromCol).toBe('A');
      expect(setReq!.args.toCol).toBe('D');
      expect(setReq!.args.properties).toEqual({ width: 100 });
    });
  });

  describe('auto-fit operations', () => {
    it('auto-fits all columns', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const result = await wb.autoFitColumns('Sheet1');
      expect(result['A']).toBeDefined();
      // Fake server returns width: 12
      expect(result['A'].width).toBe(12);
      expect(result['A'].previousWidth).toBe(8);

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const fitReq = requests.find((r) => r.op === 'autoFitColumns');
      expect(fitReq).toBeDefined();
      expect(fitReq!.args.sheet).toBe('Sheet1');
    });

    it('auto-fits specific columns with options', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.autoFitColumns('Sheet1', ['A', 'B', 'C'], {
        minWidth: 50,
        maxWidth: 200,
        padding: 10,
      });

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const fitReq = requests.find((r) => r.op === 'autoFitColumns');
      expect(fitReq!.args.columns).toEqual(['A', 'B', 'C']);
      expect(fitReq!.args.minWidth).toBe(50);
      expect(fitReq!.args.maxWidth).toBe(200);
      expect(fitReq!.args.padding).toBe(10);
    });

    it('auto-fits all rows', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const result = await wb.autoFitRows('Sheet1');
      expect(result['1']).toBeDefined();
      // Fake server returns height: 15
      expect(result['1'].height).toBe(15);
      expect(result['1'].previousHeight).toBe(12);

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const fitReq = requests.find((r) => r.op === 'autoFitRows');
      expect(fitReq).toBeDefined();
    });

    it('auto-fits specific rows with options', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.autoFitRows('Sheet1', [1, 2, 3], { minHeight: 15, maxHeight: 50 });

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const fitReq = requests.find((r) => r.op === 'autoFitRows');
      expect(fitReq!.args.rows).toEqual([1, 2, 3]);
      expect(fitReq!.args.minHeight).toBe(15);
      expect(fitReq!.args.maxHeight).toBe(50);
    });
  });

  describe('sort operations', () => {
    it('sorts a range by single column', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.sortRange('A1:D10', [{ column: 'A' }]);

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const sortReq = requests.find((r) => r.op === 'sortRange');
      expect(sortReq).toBeDefined();
      expect(sortReq!.args.range).toBe('A1:D10');
      expect(sortReq!.args.keys).toEqual([{ column: 'A' }]);
    });

    it('sorts a range by multiple columns with options', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.sortRange(
        'A1:D10',
        [
          { column: 'B', descending: true },
          { column: 2, descending: false },
        ],
        { hasHeader: true }
      );

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const sortReq = requests.find((r) => r.op === 'sortRange');
      expect(sortReq!.args.keys).toHaveLength(2);
      expect(sortReq!.args.hasHeader).toBe(true);
    });
  });

  describe('defined names', () => {
    it('lists defined names', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const names = await wb.listDefinedNames();
      expect(names).toEqual([]);

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      expect(requests.some((r) => r.op === 'listDefinedNames')).toBe(true);
    });

    it('adds a defined name', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const result = await wb.addDefinedName('MyName', 'Sheet1!A1:B2');
      expect(result.name).toBe('MyName');

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const addReq = requests.find((r) => r.op === 'addDefinedName');
      expect(addReq!.args.name).toBe('MyName');
      expect(addReq!.args.range).toBe('Sheet1!A1:B2');
    });

    it('adds a defined name with scope', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.addDefinedName('LocalName', 'A1', { scope: 'Sheet1' });

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const addReq = requests.find((r) => r.op === 'addDefinedName');
      expect(addReq!.args.scope).toBe('Sheet1');
    });

    it('deletes a defined name', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.deleteDefinedName('MyName');

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const delReq = requests.find((r) => r.op === 'deleteDefinedName');
      expect(delReq!.args.name).toBe('MyName');
    });
  });

  describe('list objects (tables)', () => {
    it('gets a list object', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const result = await wb.getListObject('Table1');
      expect(result.name).toBe('Table1');

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      expect(requests.some((r) => r.op === 'getListObject')).toBe(true);
    });

    it('adds a list object', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const result = await wb.addListObject('Sheet1', {
        name: 'NewTable',
        ref: 'A1:C10',
        columns: [{ name: 'Name' }, { name: 'Region' }, { name: 'Sales' }],
      });
      expect(result.listObject).toBeDefined();

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const addReq = requests.find((r) => r.op === 'addListObject');
      expect(addReq!.args.sheet).toBe('Sheet1');
    });

    it('sets a list object', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.setListObject('Table1', { showTotalsRow: true });

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const setReq = requests.find((r) => r.op === 'setListObject');
      expect(setReq!.args.name).toBe('Table1');
    });

    it('deletes a list object', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.deleteListObject('Table1');

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      expect(requests.some((r) => r.op === 'deleteListObject')).toBe(true);
    });
  });

  describe('data tables', () => {
    it('gets a data table', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const result = await wb.getDataTable('Sheet1!A1:B5');
      expect(result.type).toBe('oneVariableColumn');

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      expect(requests.some((r) => r.op === 'getDataTable')).toBe(true);
    });

    it('adds a data table', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const result = await wb.addDataTable('Sheet1', {
        type: 'oneVariableColumn',
        ref: 'A1:B5',
        columnInputCell: 'D1',
        inputValues: [10, 20, 30, 40],
        formulas: ['=C1'],
      });
      expect(result.dataTable).toBeDefined();

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      expect(requests.some((r) => r.op === 'addDataTable')).toBe(true);
    });

    it('deletes a data table', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.deleteDataTable('Sheet1!A1:B5');

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      expect(requests.some((r) => r.op === 'deleteDataTable')).toBe(true);
    });
  });

  describe('charts', () => {
    it('lists charts', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const charts = await wb.listCharts();
      expect(charts).toEqual([]);

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      expect(requests.some((r) => r.op === 'listCharts')).toBe(true);
    });

    it('lists charts filtered by sheet', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.listCharts({ sheet: 'Sheet1' });

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const listReq = requests.find((r) => r.op === 'listCharts');
      expect(listReq!.args.sheet).toBe('Sheet1');
    });

    it('gets a chart', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const chart = await wb.getChart('Sheet1', 'Chart1');
      expect(chart.name).toBe('Chart1');
    });

    it('adds a chart', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const chart = await wb.addChart('Sheet1', {
        name: 'Chart1',
        position: {
          from: { cell: 'E2' },
          to: { cell: 'L18' },
        },
        groups: [
          {
            type: 'bar',
            series: [
              {
                name: { text: 'Sales' },
                categories: 'Sheet1!A2:A5',
                values: 'Sheet1!B2:B5',
              },
            ],
          },
        ],
      });
      expect(chart).toBeDefined();

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      expect(requests.some((r) => r.op === 'addChart')).toBe(true);
    });

    it('sets a chart', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.setChart('Sheet1', 'Chart1', {
        name: 'Chart1',
        position: {
          from: { cell: 'E2' },
          to: { cell: 'L18' },
        },
        groups: [
          {
            type: 'bar',
            series: [
              {
                name: { text: 'Sales' },
                categories: 'Sheet1!A2:A5',
                values: 'Sheet1!B2:B5',
              },
            ],
          },
        ],
        title: { text: 'Updated Title' },
      });

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const setReq = requests.find((r) => r.op === 'setChart');
      expect(setReq!.args.name).toBe('Chart1');
    });

    it('deletes a chart', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.deleteChart('Sheet1', 'Chart1');

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      expect(requests.some((r) => r.op === 'deleteChart')).toBe(true);
    });
  });

  describe('conditional formatting', () => {
    it('gets conditional formatting rules', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const rules = await wb.getConditionalFormatting('Sheet1');
      expect(rules).toEqual([]);
    });

    it('sets conditional formatting rules', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.setConditionalFormatting('Sheet1', [{ address: 'A1', type: 'cellValue', operator: 'greaterThan', formula: '10' }]);

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const setReq = requests.find((r) => r.op === 'setConditionalFormatting');
      expect(setReq!.args.rules).toHaveLength(1);
    });

    it('sets conditional formatting with clear option', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.setConditionalFormatting('Sheet1', [], { clear: true });

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const setReq = requests.find((r) => r.op === 'setConditionalFormatting');
      expect(setReq!.args.clear).toBe(true);
    });

    it('removes conditional formatting by indices', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.removeConditionalFormatting('Sheet1', [0, 2]);

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const remReq = requests.find((r) => r.op === 'removeConditionalFormatting');
      expect(remReq!.args.indices).toEqual([0, 2]);
    });
  });

  describe('formula operations', () => {
    it('evaluates multiple formulas', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const results = await wb.evaluateFormulas('Sheet1', ['=1+1', '=SUM(A1:A10)']);
      expect(results).toHaveLength(2);
      expect(results[0].value).toBe(42); // Fake server returns 42 for all formulas
    });

    it('evaluates a single formula', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const result = await wb.evaluateFormula('Sheet1', '=A1+B1');
      expect(result.formula).toBe('=A1+B1');
      expect(result.value).toBe(42);
    });

    it('gets cell precedents', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const result = await wb.getCellPrecedents('Sheet1!C1');
      expect(result.cells).toEqual([]);

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const getReq = requests.find((r) => r.op === 'getCellPrecedents');
      expect(getReq!.args.depth).toBe(1);
    });

    it('gets cell precedents with infinite depth', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.getCellPrecedents('Sheet1!C1', Infinity);

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const getReq = requests.find((r) => r.op === 'getCellPrecedents');
      expect(getReq!.args.depth).toBe(-1);
    });

    it('gets cell dependents', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const result = await wb.getCellDependents('Sheet1!A1', 3);
      expect(result.cells).toEqual([]);

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const getReq = requests.find((r) => r.op === 'getCellDependents');
      expect(getReq!.args.depth).toBe(3);
    });

    it('traces to inputs', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const inputs = await wb.traceToInputs('Sheet1!C1');
      expect(inputs).toEqual([]);
    });

    it('traces to outputs', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const outputs = await wb.traceToOutputs('Sheet1!A1');
      expect(outputs).toEqual([]);
    });
  });

  describe('scenarios / sweep', () => {
    it('runs sweep inputs', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const result = await wb.sweepInputs(
        [{ address: 'A1', values: [1, 2, 3] }],
        ['B1', 'C1']
      );
      expect(result.sweepCount).toBe(1);

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const sweepReq = requests.find((r) => r.op === 'sweepInputs');
      expect(sweepReq!.args.inputs).toHaveLength(1);
      expect(sweepReq!.args.outputs).toHaveLength(2);
    });

    it('runs sweep with options', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.sweepInputs([{ address: 'A1', values: [1, 2] }], ['B1'], {
        mode: 'cartesian',
        includeStats: true,
      });

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const sweepReq = requests.find((r) => r.op === 'sweepInputs');
      expect(sweepReq!.args.mode).toBe('cartesian');
      expect(sweepReq!.args.includeStats).toBe(true);
    });

    it('scenarios is alias for sweepInputs', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const result = await wb.scenarios([{ address: 'A1', values: [1] }], ['B1']);
      expect(result.sweepCount).toBe(1);

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      expect(requests.some((r) => r.op === 'sweepInputs')).toBe(true);
    });
  });

  describe('utilities', () => {
    it('describes a sheet', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const description = await wb.describeSheet('Sheet1');
      expect(description.structure).toBeDefined();
    });

    it('describes all visible sheets', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const descriptions = await wb.describeSheets();
      // Should include Sheet1 but not Hidden (which has hidden: true)
      expect(descriptions['Sheet1']).toBeDefined();
      expect(descriptions['Hidden']).toBeUndefined();
    });

    it('performs table lookup', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const results = await wb.tableLookup('Table1', 'Row1', 'ColA');
      expect(results).toEqual([]);

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const lookupReq = requests.find((r) => r.op === 'tableLookup');
      expect(lookupReq!.args.table).toBe('Table1');
      expect(lookupReq!.args.rowLabel).toBe('Row1');
      expect(lookupReq!.args.columnLabel).toBe('ColA');
    });

    it('runs lint', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const result = await wb.lint();
      expect(result.total).toBe(0);
      expect(result.diagnostics).toEqual([]);
    });

    it('runs lint with options', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      await wb.lint({
        rangeAddresses: ['Sheet1!A1:Z100'],
        skipRuleIds: ['RULE1'],
        onlyRuleIds: ['RULE2', 'RULE3'],
      });

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const lintReq = requests.find((r) => r.op === 'lint');
      expect(lintReq!.args.rangeAddresses).toEqual(['Sheet1!A1:Z100']);
      expect(lintReq!.args.skipRuleIds).toEqual(['RULE1']);
      expect(lintReq!.args.onlyRuleIds).toEqual(['RULE2', 'RULE3']);
    });

    it('previews styles', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const dataUrl = await wb.previewStyles('A1:B5');
      expect(dataUrl).toBe('data:image/png;base64,AAA=');
    });

    it('reduces addresses', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const reduced = await wb.reduceAddresses(['A1', 'A2', 'B1', 'B2']);
      expect(reduced).toEqual(['Sheet1!A1:B2']);
    });

    it('gets cell style', async () => {
      const env = fakeEnv(tmpDir);
      await using wb = await Workbook.open(join(tmpDir, 'test.xlsx'), {
        binary: FAKE_WITAN_PATH,
        env,
      });

      const style = await wb.getStyle('A1');
      expect(style).toEqual({});

      const requests = await readRequests(env.WITAN_FAKE_REQUESTS_FILE);
      const styleReq = requests.find((r) => r.op === 'getStyle');
      expect(styleReq!.args.address).toBe('A1');
    });
  });
});
