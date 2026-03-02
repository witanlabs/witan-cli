import { execSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import dotenv from 'dotenv';

/**
 * Load environment variables from `examples/.env` and ensure the witan CLI
 * binary is reachable on PATH.
 *
 * Call this once at the top of your entry-point (e.g. `qna.ts`) before
 * invoking any runner.
 */
export function loadEnv(): void {
  // 1. Load .env from the examples directory (one level up from lib/)
  const envPath = path.resolve(import.meta.dirname, '..', '.env');
  dotenv.config({ path: envPath, override: true });

  // 2. Verify the witan binary exists at the repo root (two levels up from lib/)
  const witanBinary = path.resolve(import.meta.dirname, '..', '..', 'witan');
  if (!fs.existsSync(witanBinary)) {
    console.error(
      `witan binary not found at ${witanBinary}\n` +
        'Run "make build" in the witan-cli root to compile it.',
    );
    process.exit(1);
  }

  // 3. Prepend the repo root to PATH so child processes can find witan
  const repoRoot = path.resolve(import.meta.dirname, '..', '..');
  process.env.PATH = `${repoRoot}${path.delimiter}${process.env.PATH ?? ''}`;
}


/**
 * Create a Python virtual environment inside `workDir` with openpyxl installed.
 *
 * The venv lives at `<workDir>/venv/` so the agent can use `./venv/bin/python`
 * without needing PATH manipulation (which gets reset by shell profile sourcing).
 */
export function setupPythonVenv(workDir: string): void {
  const venvDir = path.join(workDir, 'venv');
  if (fs.existsSync(venvDir)) return;

  console.log('Setting up Python venv with openpyxl...');
  execSync('python3 -m venv venv', { cwd: workDir, stdio: 'pipe' });
  execSync('./venv/bin/pip install -q openpyxl', { cwd: workDir, stdio: 'pipe' });
  console.log('Python venv ready.\n');
}

/**
 * Patch an xlsx file's workbook.xml to enable iterative calculation.
 *
 * ExcelJS does not write the `iterate`, `iterateCount`, or `iterateDelta`
 * attributes to `<calcPr>` even when `calcProperties` is set. This function
 * unzips the xlsx, patches the XML, and re-zips it.
 *
 * Uses Excel's default settings: 100 max iterations, 0.001 max change.
 */
export function enableIterativeCalc(xlsxPath: string): void {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlsx-patch-'));
  try {
    execSync(`unzip -o "${xlsxPath}" -d "${tmpDir}" > /dev/null`);

    const workbookXmlPath = path.join(tmpDir, 'xl/workbook.xml');
    let xml = fs.readFileSync(workbookXmlPath, 'utf-8');

    if (/<calcPr[^/]*\/>/.test(xml)) {
      xml = xml.replace(
        /<calcPr[^/]*\/>/,
        '<calcPr calcId="171027" iterate="1" iterateCount="100" iterateDelta="0.001"/>',
      );
    } else {
      xml = xml.replace(
        '</workbook>',
        '<calcPr calcId="171027" iterate="1" iterateCount="100" iterateDelta="0.001"/></workbook>',
      );
    }

    fs.writeFileSync(workbookXmlPath, xml);
    execSync(`cd "${tmpDir}" && zip -r "${xlsxPath}" . > /dev/null`);
  } finally {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
}
