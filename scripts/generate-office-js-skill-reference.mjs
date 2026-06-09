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

const allowedOutputSuffixes = [
  `${sep}skills${sep}pptx-code-mode${sep}references${sep}office-js.d.ts`,
  `${sep}skills${sep}xlsx-office-script${sep}references${sep}excelscript.d.ts`,
];

const resolvedOutputPath = resolve(outputPath);
if (
  resolvedOutputPath === sep ||
  !allowedOutputSuffixes.some((suffix) => resolvedOutputPath.endsWith(suffix))
) {
  console.error(
    `refusing to replace an unexpected output file; expected one of:\n${allowedOutputSuffixes
      .map((suffix) => `  ${suffix.replace(new RegExp(`^\\${sep}`), "")}`)
      .join("\n")}`,
  );
  process.exit(1);
}

const source = readFileSync(sourcePath, "utf8");

rmSync(outputPath, { force: true });
mkdirSync(dirname(outputPath), { recursive: true });
writeFileSync(outputPath, source);
