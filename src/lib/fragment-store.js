import { randomUUID } from 'crypto';
import { copyFileSync, existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'fs';
import { tmpdir } from 'os';
import { basename, extname, join } from 'path';

import { ToolError } from './errors.js';

const STORE_DIR = join(tmpdir(), 'office-mcp-fragments');
const TTL_MS = 24 * 60 * 60 * 1000;

function ensureStoreDir() {
  mkdirSync(STORE_DIR, { recursive: true });
}

function metadataPathFor(ref) {
  return join(STORE_DIR, `${ref}.json`);
}

function removeFragmentFiles(metadata) {
  rmSync(metadataPathFor(metadata.ref), { force: true });
  if (metadata.filePath) {
    rmSync(metadata.filePath, { force: true });
  }
}

export function cleanupExpiredFragments() {
  ensureStoreDir();
  const now = Date.now();
  for (const entry of readdirSync(STORE_DIR)) {
    if (!entry.endsWith('.json')) {
      continue;
    }
    const metaPath = join(STORE_DIR, entry);
    try {
      const metadata = JSON.parse(readFileSync(metaPath, 'utf8'));
      if (!metadata.expiresAt || Date.parse(metadata.expiresAt) <= now || !metadata.filePath || !existsSync(metadata.filePath)) {
        removeFragmentFiles(metadata);
      }
    } catch {
      rmSync(metaPath, { force: true });
    }
  }
}

function buildMetadata({ ref, app, kind, format, filePath, summary }) {
  const createdAt = new Date().toISOString();
  const expiresAt = new Date(Date.now() + TTL_MS).toISOString();
  return { ref, app, kind, format, filePath, summary, createdAt, expiresAt };
}

export function reserveFragment({ prefix, app, kind, extension, summary }) {
  cleanupExpiredFragments();
  const ref = `${prefix}_${randomUUID().replace(/-/g, '')}`;
  const normalizedExtension = extension.replace(/^\./, '').toLowerCase();
  const filePath = join(STORE_DIR, `${ref}.${normalizedExtension}`);
  const metadata = buildMetadata({ ref, app, kind, format: normalizedExtension, filePath, summary });
  return {
    metadata,
    filePath,
    metadataPath: metadataPathFor(ref)
  };
}

export function commitReservedFragment(reserved) {
  writeFileSync(reserved.metadataPath, JSON.stringify(reserved.metadata, null, 2), 'utf8');
  return toFragmentPayload(reserved.metadata);
}

export function discardReservedFragment(reserved) {
  rmSync(reserved.metadataPath, { force: true });
  rmSync(reserved.filePath, { force: true });
}

export function createFileBackedFragment({ prefix, app, kind, sourcePath, summary }) {
  const extension = extname(sourcePath).slice(1) || 'bin';
  const reserved = reserveFragment({ prefix, app, kind, extension, summary });
  copyFileSync(sourcePath, reserved.filePath);
  return commitReservedFragment(reserved);
}

export function getFragment(ref, expectedApp = undefined) {
  cleanupExpiredFragments();
  const metaPath = metadataPathFor(ref);
  if (!existsSync(metaPath)) {
    throw new ToolError('NOT_FOUND', `Fragment ref not found: ${ref}`);
  }

  const metadata = JSON.parse(readFileSync(metaPath, 'utf8'));
  if (!metadata.filePath || !existsSync(metadata.filePath)) {
    removeFragmentFiles(metadata);
    throw new ToolError('NOT_FOUND', `Fragment ref not found: ${ref}`);
  }

  if (expectedApp && metadata.app !== expectedApp) {
    throw new ToolError('VALIDATION_ERROR', `ref must target ${expectedApp}`);
  }

  return metadata;
}

export function toFragmentPayload(metadata) {
  return {
    ref: metadata.ref,
    app: metadata.app,
    kind: metadata.kind,
    format: metadata.format,
    summary: metadata.summary,
    expiresAt: metadata.expiresAt
  };
}

export function buildSourceSummary(label, extra = undefined) {
  return extra ? { label, ...extra } : { label };
}

export function inferImageFormat(path) {
  return extname(path).slice(1).toLowerCase();
}

export function buildFileSummary(path) {
  return {
    label: basename(path),
    sourcePath: path
  };
}
