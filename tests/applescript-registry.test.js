import { describe, expect, jest, test } from '@jest/globals';
import { mkdirSync, writeFileSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';

import { compileAppleScript } from './helpers/applescript-compile.js';

let capturedScripts = {};

jest.unstable_mockModule('../src/lib/applescript/executor.js', () => ({
  runAppleScript: async script => {
    capturedScripts._last = script;
    return 'mocked';
  }
}));

const { ALL_TOOLS } = await import('../src/lib/tool-registry.js');

const WORD_DOC_PATH = '/tmp/office-mcp-test.docx';
const WORD_IMAGE_PATH = '/tmp/office-mcp-test.png';
const WORD_PDF_PATH = '/tmp/office-mcp-test.pdf';
const EXCEL_BOOK_PATH = '/tmp/office-mcp-test.xlsx';
const EXCEL_CSV_PATH = '/tmp/office-mcp-test.csv';
const WORD_FRAGMENT_REF = 'wordfrag_registry';
const WORD_IMAGE_REF = 'wordimg_registry';
const EXCEL_FRAGMENT_REF = 'excelfrag_registry';
const NON_SCRIPT_TOOLS = new Set([
  'word_create_image_ref',
  'word_capture_content_ref',
  'excel_capture_range_ref',
  'excel_insert_range_ref'
]);

function createFragmentFixture(ref, metadata) {
  const storeDir = join(tmpdir(), 'office-mcp-fragments');
  mkdirSync(storeDir, { recursive: true });
  if (metadata.filePath) {
    writeFileSync(metadata.filePath, 'fixture', 'utf8');
  }
  writeFileSync(join(storeDir, `${ref}.json`), JSON.stringify(metadata, null, 2), 'utf8');
}

function setupRefFixtures() {
  createFragmentFixture(WORD_FRAGMENT_REF, {
    ref: WORD_FRAGMENT_REF,
    app: 'word',
    kind: 'word_fragment',
    format: 'docx',
    filePath: join(tmpdir(), `${WORD_FRAGMENT_REF}.docx`),
    summary: { label: 'registry fixture' },
    createdAt: new Date().toISOString(),
    expiresAt: new Date(Date.now() + 60_000).toISOString()
  });

  createFragmentFixture(WORD_IMAGE_REF, {
    ref: WORD_IMAGE_REF,
    app: 'word',
    kind: 'image_file',
    format: 'png',
    filePath: WORD_IMAGE_PATH,
    summary: { label: 'registry image fixture' },
    createdAt: new Date().toISOString(),
    expiresAt: new Date(Date.now() + 60_000).toISOString()
  });

  createFragmentFixture(EXCEL_FRAGMENT_REF, {
    ref: EXCEL_FRAGMENT_REF,
    app: 'excel',
    kind: 'excel_range',
    format: 'xlsx',
    filePath: join(tmpdir(), `${EXCEL_FRAGMENT_REF}.xlsx`),
    summary: { label: 'registry excel fixture' },
    createdAt: new Date().toISOString(),
    expiresAt: new Date(Date.now() + 60_000).toISOString()
  });
}

function getSpecialArgs(toolName) {
  switch (toolName) {
    case 'word_format_text':
      return { bold: true };
    case 'word_resize_inline_shape':
      return { index: 1, width: 120 };
    case 'word_copy_content':
      return { scope: 'document' };
    case 'word_copy_story_content':
    case 'word_clear_story_content':
      return { scope: 'body' };
    case 'word_set_story_text':
      return { scope: 'body', text: 'Sample text' };
    case 'word_insert_content_ref':
      return { ref: WORD_IMAGE_REF };
    case 'word_create_image_ref':
      return { path: WORD_IMAGE_PATH };
    case 'excel_capture_range_ref':
    case 'excel_copy_range':
      return { range: 'A1:B3' };
    case 'excel_paste_range':
      return { targetCell: 'C3' };
    case 'excel_clear_worksheet':
      return { worksheet: 1 };
    case 'excel_set_range_values':
      return { range: 'A1:B2', values: '1\t2\n3\t4' };
    case 'excel_format_cells':
      return { range: 'A1:B3', bold: true };
    case 'excel_set_cell_color':
      return { range: 'A1:B3', color: [255, 0, 0] };
    case 'excel_sort_range':
      return { range: 'A1:B10', keyCell: 'B1' };
    case 'excel_autofit':
      return { range: 'A:C' };
    default:
      return null;
  }
}

function valueForProperty(toolName, key, property = {}) {
  if (Array.isArray(property.enum) && property.enum.length > 0) {
    return property.enum[0];
  }

  if (key === 'path') {
    if (toolName.includes('pdf')) return WORD_PDF_PATH;
    if (toolName.includes('excel_') || toolName.includes('workbook')) return EXCEL_BOOK_PATH;
    if (toolName.includes('csv')) return EXCEL_CSV_PATH;
    if (toolName.includes('image')) return WORD_IMAGE_PATH;
    return WORD_DOC_PATH;
  }

  if (key === 'ref') {
    if (toolName.startsWith('excel_')) return EXCEL_FRAGMENT_REF;
    if (toolName === 'word_insert_content_ref') return WORD_FRAGMENT_REF;
    return WORD_IMAGE_REF;
  }

  if (key === 'url') return 'https://example.com';
  if (key === 'displayText') return 'Example';
  if (key === 'searchText') return 'needle';
  if (key === 'find') return 'old';
  if (key === 'replace') return 'new';
  if (key === 'formula') return '=SUM(A1:A3)';
  if (key === 'format') return '#,##0.00';
  if (key === 'font') return 'Arial';
  if (key === 'styleName') return 'Heading 1';
  if (key === 'newName') return 'Renamed';
  if (key === 'name') return toolName.includes('sheet') ? 'Sheet1' : 'Bookmark1';
  if (key === 'nameOrIndex') return 1;
  if (key === 'worksheet') return 1;
  if (key === 'range') return toolName === 'excel_autofit' ? 'A:C' : 'A1:B3';
  if (key === 'cell' || key === 'targetCell' || key === 'keyCell') return 'A1';
  if (key === 'headerText') return 'Header';
  if (key === 'text') return 'Sample text';
  if (key === 'content') return 'Sample content';
  if (key === 'value') return 'Sample value';
  if (key === 'color') return [255, 0, 0];

  if (property.type === 'boolean') return true;
  if (property.type === 'integer') {
    if (key === 'rows' || key === 'columns' || key === 'count') return 2;
    if (key === 'section' || key === 'index' || key === 'row' || key === 'column') return 1;
    if (key === 'limit') return 5;
    return 1;
  }
  if (property.type === 'number') {
    if (key === 'size') return 12;
    if (key === 'width' || key === 'height') return 120;
    return 1;
  }
  if (property.type === 'string') return `${key}-value`;

  return undefined;
}

function buildArgs(tool) {
  const specialArgs = getSpecialArgs(tool.name);
  if (specialArgs) {
    return specialArgs;
  }

  const schema = tool.inputSchema || {};
  const properties = schema.properties || {};
  const args = {};

  for (const key of Object.keys(properties)) {
    const value = valueForProperty(tool.name, key, properties[key]);
    if (value !== undefined) {
      args[key] = value;
    }
  }

  return args;
}

async function captureScript(tool, args) {
  capturedScripts._last = null;
  try {
    await tool.handler(args);
  } catch {}
  return capturedScripts._last;
}

describe('AppleScript Registry Strict Syntax', () => {
  setupRefFixtures();

  test('every AppleScript-backed tool compiles from the registry', async () => {
    const failures = [];

    for (const tool of ALL_TOOLS) {
      const args = buildArgs(tool);
      const script = await captureScript(tool, args);

      if (!script) {
        if (NON_SCRIPT_TOOLS.has(tool.name)) {
          continue;
        }
        failures.push(`${tool.name}: script was not generated`);
        continue;
      }

      const result = compileAppleScript(script);
      if (!result.ok) {
        failures.push(`${tool.name}: ${result.error}`);
      }
    }

    expect(failures).toEqual([]);
  });
});
