#!/usr/bin/env node
// CI gate: any change under skills/<name>/ must bump metadata.version in that
// skill's SKILL.md. See skills/README.md.
//
// Usage: node scripts/check-skill-versions.mjs <base-ref>
//   e.g. node scripts/check-skill-versions.mjs origin/main

import { execFileSync } from 'node:child_process';
import { pathToFileURL } from 'node:url';

function stripQuotes(value) {
  const match = value.match(/^"(.*)"$/) ?? value.match(/^'(.*)'$/);
  return match ? match[1] : value;
}

// Minimal frontmatter reader for the documented SKILL.md format: a YAML block
// mapping with metadata children at two-space indent. Anything it can't read
// comes back absent, which the gate treats as an error (fails closed).
export function parseFrontmatter(text) {
  const lines = text.split('\n');
  if (lines[0] !== '---') return null;
  const end = lines.indexOf('---', 1);
  if (end === -1) return null;

  const frontmatter = {};
  let inMetadata = false;
  for (const line of lines.slice(1, end)) {
    const topLevel = line.match(/^(\S[^:]*):/);
    if (topLevel) {
      inMetadata = topLevel[1] === 'metadata';
      frontmatter[topLevel[1]] = inMetadata ? {} : true;
      continue;
    }
    const child = inMetadata && line.match(/^  (\S[^:]*):\s*(.*)$/);
    if (child) {
      frontmatter.metadata[child[1]] = stripQuotes(child[2].trim());
    }
  }
  return frontmatter;
}

export function metadataVersion(frontmatter) {
  const version = frontmatter?.metadata?.version;
  return typeof version === 'string' ? version : null;
}

export function hasTopLevelVersion(frontmatter) {
  return frontmatter !== null && Object.hasOwn(frontmatter, 'version');
}

// Plain dotted numerics only (e.g. "1.0.0") — no prerelease or build suffixes.
export function isPlainVersion(version) {
  return /^\d+(\.\d+)*$/.test(version);
}

export function compareVersions(a, b) {
  const as = a.split('.').map(Number);
  const bs = b.split('.').map(Number);
  for (let i = 0; i < Math.max(as.length, bs.length); i++) {
    const diff = (as[i] ?? 0) - (bs[i] ?? 0);
    if (diff !== 0) return Math.sign(diff);
  }
  return 0;
}

function git(...args) {
  return execFileSync('git', args, { encoding: 'utf8' });
}

function frontmatterAt(rev, dir) {
  let text;
  try {
    text = git('show', `${rev}:${dir}/SKILL.md`);
  } catch {
    return null;
  }
  return parseFrontmatter(text);
}

function changedSkillDirs(base) {
  const changed = git('diff', '--name-only', base, 'HEAD', '--', 'skills/');
  const dirs = changed
    .split('\n')
    .filter((path) => path.split('/').length > 2)
    .map((path) => path.split('/').slice(0, 2).join('/'));
  return [...new Set(dirs)].sort();
}

function main() {
  const baseRef = process.argv[2];
  if (!baseRef) {
    console.error('usage: check-skill-versions.mjs <base-ref>');
    process.exit(2);
  }

  const base = git('merge-base', baseRef, 'HEAD').trim();
  let failed = false;
  const fail = (message) => {
    console.error(`error: ${message}`);
    failed = true;
  };

  for (const dir of changedSkillDirs(base)) {
    const head = frontmatterAt('HEAD', dir);
    if (head === null) continue; // skill deleted in this PR

    if (hasTopLevelVersion(head)) {
      fail(`${dir} SKILL.md has a top-level version key — the skills spec rejects it; use metadata.version`);
    }

    const newVersion = metadataVersion(head);
    const baseFm = frontmatterAt(base, dir);
    const oldVersion = metadataVersion(baseFm);

    if (newVersion === null) {
      fail(`${dir} changed but SKILL.md has no metadata.version`);
    } else if (!isPlainVersion(newVersion)) {
      fail(`${dir} metadata.version "${newVersion}" is not a plain dotted version like "1.0.0"`);
    } else if (oldVersion === null || !isPlainVersion(oldVersion)) {
      console.log(`${dir}: new at ${newVersion}`);
    } else if (compareVersions(newVersion, oldVersion) <= 0) {
      fail(`${dir} changed but metadata.version did not increase (${oldVersion} -> ${newVersion}) — bump it`);
    } else {
      console.log(`${dir}: ${oldVersion} -> ${newVersion}`);
    }
  }

  process.exit(failed ? 1 : 0);
}

if (import.meta.url === pathToFileURL(process.argv[1]).href) {
  main();
}
