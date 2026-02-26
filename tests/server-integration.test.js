import { describe, expect, test } from '@jest/globals';
import { readFileSync } from 'fs';
import { createServer } from '../src/lib/server.js';
import { executeTool } from '../src/lib/tool-executor.js';
import { ALL_TOOLS } from '../src/lib/tool-registry.js';

const manifest = JSON.parse(readFileSync(new URL('../manifest.json', import.meta.url), 'utf8'));

function parsePayload(response) {
  return JSON.parse(response.content[0].text);
}

describe('MCP Server Integration', () => {
  test('registry and manifest define the same tools', () => {
    const registryNames = new Set(ALL_TOOLS.map(tool => tool.name));
    const manifestNames = new Set(manifest.tools.map(tool => tool.name));

    expect(registryNames.size).toBe(86);
    expect(manifestNames.size).toBe(86);
    expect([...registryNames].filter(name => !manifestNames.has(name))).toHaveLength(0);
    expect([...manifestNames].filter(name => !registryNames.has(name))).toHaveLength(0);
  });

  test('server can be created', () => {
    const server = createServer();
    expect(server).toBeDefined();
    expect(typeof server.connect).toBe('function');
    expect(typeof server.setRequestHandler).toBe('function');
  });

  test('executeTool returns success payload envelope', async () => {
    const response = await executeTool('demo_tool', {}, async () => 'Done');
    const payload = parsePayload(response);

    expect(response.isError).toBeUndefined();
    expect(payload.ok).toBe(true);
    expect(payload.message).toBe('Done');
  });

  test('executeTool converts thrown errors into error payload envelope', async () => {
    const response = await executeTool('demo_tool', {}, async () => {
      throw new Error('No workbook is open');
    });
    const payload = parsePayload(response);

    expect(response.isError).toBe(true);
    expect(payload.ok).toBe(false);
    expect(payload.error.code).toBe('NO_WORKBOOK_OPEN');
    expect(payload.error.message).toContain('Failed to demo_tool');
  });

  test('executeTool treats known error-like messages as error payload', async () => {
    const response = await executeTool('demo_tool', {}, async () => 'Text not found, no replacements made');
    const payload = parsePayload(response);

    expect(response.isError).toBe(true);
    expect(payload.ok).toBe(false);
    expect(payload.error.code).toBe('NOT_FOUND');
    expect(payload.error.message).toContain('not found');
  });
});
