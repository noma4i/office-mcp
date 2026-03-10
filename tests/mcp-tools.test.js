import { describe, expect, test } from '@jest/globals';
import { ALL_TOOLS, buildToolMap, getToolDefinitions, getToolHandler } from '../src/lib/tool-registry.js';

describe('MCP Office Tools Registry', () => {
  test('contains all declared tools', () => {
    expect(ALL_TOOLS).toHaveLength(93);
  });

  test('contains 56 Word tools and 37 Excel tools', () => {
    const wordTools = ALL_TOOLS.filter(tool => tool.name.startsWith('word_'));
    const excelTools = ALL_TOOLS.filter(tool => tool.name.startsWith('excel_'));

    expect(wordTools).toHaveLength(56);
    expect(excelTools).toHaveLength(37);
  });

  test('returns definitions with required fields', () => {
    const definitions = getToolDefinitions();
    expect(definitions).toHaveLength(93);

    for (const definition of definitions) {
      expect(definition).toHaveProperty('name');
      expect(definition).toHaveProperty('description');
      expect(definition).toHaveProperty('annotations');
      expect(definition).toHaveProperty('inputSchema');
    }
  });

  test('all tool names are unique', () => {
    const names = ALL_TOOLS.map(tool => tool.name);
    expect(new Set(names).size).toBe(names.length);
  });

  test('getToolHandler resolves handlers for known tools', () => {
    const handler = getToolHandler('word_create_document');
    expect(typeof handler).toBe('function');
  });

  test('getToolHandler throws for unknown tool', () => {
    expect(() => getToolHandler('unknown_tool')).toThrow('Unknown tool: unknown_tool');
  });

  test('buildToolMap rejects duplicate tool names', () => {
    expect(() => {
      buildToolMap([
        { name: 'duplicate', handler: () => {} },
        { name: 'duplicate', handler: () => {} }
      ]);
    }).toThrow('Duplicate tool registration: duplicate');
  });
});
