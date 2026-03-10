import { describe, expect, test } from '@jest/globals';
import { existsSync, readFileSync, rmSync, writeFileSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';

import {
  cleanupExpiredFragments,
  commitReservedFragment,
  createFileBackedFragment,
  getFragment,
  reserveFragment
} from '../src/lib/fragment-store.js';

describe('Fragment Store', () => {
  test('commits and resolves reserved fragments', () => {
    const reserved = reserveFragment({
      prefix: 'wordfrag',
      app: 'word',
      kind: 'word_fragment',
      extension: 'docx',
      summary: { label: 'selection' }
    });

    writeFileSync(reserved.filePath, 'fragment', 'utf8');
    const payload = commitReservedFragment(reserved);
    const metadata = getFragment(payload.ref, 'word');

    expect(payload.app).toBe('word');
    expect(metadata.kind).toBe('word_fragment');
    expect(existsSync(metadata.filePath)).toBe(true);
  });

  test('creates file-backed fragments', () => {
    const sourcePath = join(tmpdir(), 'office_mcp_fragment_test.png');
    writeFileSync(sourcePath, 'png-data', 'utf8');

    const payload = createFileBackedFragment({
      prefix: 'wordimg',
      app: 'word',
      kind: 'image_file',
      sourcePath,
      summary: { label: 'image' }
    });

    const metadata = getFragment(payload.ref, 'word');
    expect(metadata.kind).toBe('image_file');
    expect(readFileSync(metadata.filePath, 'utf8')).toBe('png-data');

    rmSync(sourcePath, { force: true });
  });

  test('cleans up stale metadata with missing files', () => {
    const reserved = reserveFragment({
      prefix: 'excelfrag',
      app: 'excel',
      kind: 'excel_range',
      extension: 'xlsx',
      summary: { label: 'range' }
    });

    writeFileSync(reserved.filePath, 'xlsx-data', 'utf8');
    const payload = commitReservedFragment(reserved);
    rmSync(reserved.filePath, { force: true });

    cleanupExpiredFragments();

    expect(() => getFragment(payload.ref, 'excel')).toThrow(`Fragment ref not found: ${payload.ref}`);
  });
});
