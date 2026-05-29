#!/usr/bin/env node

import { mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { dirname, resolve, sep } from "node:path";

const [sourcePath, outputPath] = process.argv.slice(2);

if (!sourcePath || !outputPath) {
  console.error(
    "usage: generate-office-js-skill-reference.mjs <office.d.ts> <output-file>",
  );
  process.exit(1);
}

const resolvedOutputPath = resolve(outputPath);
if (
  resolvedOutputPath === sep ||
  !resolvedOutputPath.endsWith(
    `${sep}skills${sep}pptx-code-mode${sep}references${sep}office-js.d.ts`,
  )
) {
  console.error(
    "refusing to replace an unexpected output file; expected skills/pptx-code-mode/references/office-js.d.ts",
  );
  process.exit(1);
}

const source = readFileSync(sourcePath, "utf8");

rmSync(outputPath, { force: true });
mkdirSync(dirname(outputPath), { recursive: true });
writeFileSync(outputPath, source);
