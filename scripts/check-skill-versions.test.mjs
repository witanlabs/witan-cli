import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  parseFrontmatter,
  metadataVersion,
  hasTopLevelVersion,
  isPlainVersion,
  compareVersions,
} from './check-skill-versions.mjs';

const doc = (...lines) => ['---', ...lines, '---', '', '## Body', '---', 'not frontmatter'].join('\n');

test('reads metadata.version', () => {
  const fm = parseFrontmatter(doc('name: x', 'metadata:', '  version: "1.2.3"'));
  assert.equal(metadataVersion(fm), '1.2.3');
});

test('reads metadata.version among other metadata keys', () => {
  const fm = parseFrontmatter(doc('metadata:', '  internal: true', '  version: "2.0.0"'));
  assert.equal(metadataVersion(fm), '2.0.0');
});

test('flow-style metadata is unsupported and fails closed', () => {
  const fm = parseFrontmatter(doc('metadata: { version: "1.2.3" }'));
  assert.equal(metadataVersion(fm), null);
});

test('top-level version is not metadata.version', () => {
  const fm = parseFrontmatter(doc('name: x', 'version: "9.9.9"'));
  assert.equal(metadataVersion(fm), null);
  assert.equal(hasTopLevelVersion(fm), true);
});

test('nested metadata.release.version is not metadata.version', () => {
  const fm = parseFrontmatter(doc('metadata:', '  release:', '    version: "9.9.9"'));
  assert.equal(metadataVersion(fm), null);
});

test('version under a sibling top-level key is not metadata.version', () => {
  const fm = parseFrontmatter(doc('metadata:', '  internal: true', 'other:', '  version: "9.9.9"'));
  assert.equal(metadataVersion(fm), null);
  assert.equal(hasTopLevelVersion(fm), false);
});

test('unquoted version is read as written', () => {
  const fm = parseFrontmatter(doc('metadata:', '  version: 1.0.0'));
  assert.equal(metadataVersion(fm), '1.0.0');
});

test('missing or unreadable frontmatter yields no version', () => {
  assert.equal(parseFrontmatter('# no frontmatter'), null);
  assert.equal(parseFrontmatter('---\nunclosed: true'), null);
  assert.equal(metadataVersion(parseFrontmatter(doc('{not a mapping'))), null);
  assert.equal(metadataVersion(null), null);
  assert.equal(hasTopLevelVersion(null), false);
});

test('isPlainVersion accepts dotted numerics only', () => {
  assert.equal(isPlainVersion('1.0.0'), true);
  assert.equal(isPlainVersion('0.1'), true);
  assert.equal(isPlainVersion('1.0.0-rc.1'), false);
  assert.equal(isPlainVersion('v1.0.0'), false);
  assert.equal(isPlainVersion(''), false);
});

test('compareVersions orders numerically', () => {
  assert.equal(compareVersions('1.10.0', '1.9.0'), 1);
  assert.equal(compareVersions('1.0.0', '1.0.1'), -1);
  assert.equal(compareVersions('1.0.0', '1.0.0'), 0);
  assert.equal(compareVersions('1.0', '1.0.0'), 0);
  assert.equal(compareVersions('2.0', '1.9.9'), 1);
});
