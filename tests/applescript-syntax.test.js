import { describe, test, expect, beforeAll, jest } from '@jest/globals';
import { execSync } from 'child_process';
import { writeFileSync, unlinkSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';

let capturedScripts = {};

jest.unstable_mockModule('../src/lib/applescript/executor.js', () => ({
  runAppleScript: async script => {
    capturedScripts._last = script;
    return 'mocked';
  }
}));

const { textTools } = await import('../src/tools/text.js');
const { tableTools } = await import('../src/tools/tables.js');
const { paragraphTools } = await import('../src/tools/paragraphs.js');
const { bookmarkTools } = await import('../src/tools/bookmarks.js');
const { hyperlinkTools } = await import('../src/tools/hyperlinks.js');
const { navigationTools } = await import('../src/tools/navigation.js');
const { documentTools } = await import('../src/tools/documents.js');
const { imageTools } = await import('../src/tools/images.js');
const { processTemplate } = await import('../src/lib/applescript/template-engine.js');

function findTool(tools, name) {
  return tools.find(t => t.name === name);
}

function compileAppleScript(script) {
  const tmpFile = join(tmpdir(), `as_test_${Date.now()}_${Math.random().toString(36).slice(2, 8)}.applescript`);
  try {
    writeFileSync(tmpFile, script, 'utf8');
    execSync(`osacompile -o /dev/null ${tmpFile}`, {
      timeout: 10000,
      stdio: ['pipe', 'pipe', 'pipe']
    });
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.stderr?.toString() || err.message };
  } finally {
    try {
      unlinkSync(tmpFile);
    } catch {}
  }
}

async function captureScript(tool, args = {}) {
  capturedScripts._last = null;
  try {
    await tool.handler(args);
  } catch {}
  return capturedScripts._last;
}

describe('AppleScript Syntax Verification', () => {
  describe('Document Tools', () => {
    test('create_document compiles', async () => {
      const script = await captureScript(findTool(documentTools, 'create_document'), {});
      expect(script).toBeTruthy();
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('get_document_text compiles', async () => {
      const script = await captureScript(findTool(documentTools, 'get_document_text'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('get_document_info compiles', async () => {
      const script = await captureScript(findTool(documentTools, 'get_document_info'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('save_document compiles', async () => {
      const script = await captureScript(findTool(documentTools, 'save_document'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('close_document compiles', async () => {
      const script = await captureScript(findTool(documentTools, 'close_document'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('export_pdf compiles', async () => {
      const script = await captureScript(findTool(documentTools, 'export_pdf'), { path: '/tmp/test.pdf' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('open_document compiles', async () => {
      const script = await captureScript(findTool(documentTools, 'open_document'), { path: '/tmp/test.docx' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });
  });

  describe('Text Tools', () => {
    test('insert_text compiles', async () => {
      const script = await captureScript(findTool(textTools, 'insert_text'), { text: 'Hello' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('insert_text accepts newline', async () => {
      const script = await captureScript(findTool(textTools, 'insert_text'), { text: '\n' });
      expect(script).toBeTruthy();
      expect(script).toContain('type text selection text');
    });

    test('insert_text accepts whitespace', async () => {
      const script = await captureScript(findTool(textTools, 'insert_text'), { text: '  ' });
      expect(script).toBeTruthy();
    });

    test('replace_text compiles', async () => {
      const script = await captureScript(findTool(textTools, 'replace_text'), { find: 'old', replace: 'new' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('content of replacement of findObject');
    });

    test('format_text compiles (bold)', async () => {
      const script = await captureScript(findTool(textTools, 'format_text'), { bold: true });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('format_text compiles (font + size)', async () => {
      const script = await captureScript(findTool(textTools, 'format_text'), { font: 'Arial', size: 14 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });
  });

  describe('Table Tools', () => {
    test('create_table compiles with correct target', async () => {
      const script = await captureScript(findTool(tableTools, 'create_table'), { rows: 3, columns: 4 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('text object of selection');
      expect(script).not.toMatch(/make new table at selection(?! )/);
    });

    test('list_tables compiles', async () => {
      const script = await captureScript(findTool(tableTools, 'list_tables'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('get_table_cell compiles', async () => {
      const script = await captureScript(findTool(tableTools, 'get_table_cell'), { tableIndex: 1, row: 1, column: 1 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('set_table_cell compiles', async () => {
      const script = await captureScript(findTool(tableTools, 'set_table_cell'), { tableIndex: 1, row: 1, column: 1, text: 'Test' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('select_table_cell compiles', async () => {
      const script = await captureScript(findTool(tableTools, 'select_table_cell'), { tableIndex: 1, row: 1, column: 1 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('find_table_header compiles', async () => {
      const script = await captureScript(findTool(tableTools, 'find_table_header'), { tableIndex: 1, headerText: 'Name' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('add_table_row compiles with correct insert syntax', async () => {
      const script = await captureScript(findTool(tableTools, 'add_table_row'), { tableIndex: 1, afterRow: 2 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('insert rows selection position below');
      expect(script).toContain('select (text object of row');
    });

    test('add_table_row without afterRow adds at end', async () => {
      const script = await captureScript(findTool(tableTools, 'add_table_row'), { tableIndex: 1 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('count of rows of t');
    });

    test('delete_table_row compiles', async () => {
      const script = await captureScript(findTool(tableTools, 'delete_table_row'), { tableIndex: 1, row: 2 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('add_table_column compiles with correct insert syntax', async () => {
      const script = await captureScript(findTool(tableTools, 'add_table_column'), { tableIndex: 1, afterColumn: 2 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('insert columns selection position insert on the right');
    });

    test('delete_table_column compiles', async () => {
      const script = await captureScript(findTool(tableTools, 'delete_table_column'), { tableIndex: 1, column: 2 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });
  });

  describe('Paragraph Tools', () => {
    test('list_paragraphs compiles with name local', async () => {
      const script = await captureScript(findTool(paragraphTools, 'list_paragraphs'), { limit: 10 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('name local of style of p');
      expect(script).not.toContain('paragraph style');
    });

    test('goto_paragraph compiles', async () => {
      const script = await captureScript(findTool(paragraphTools, 'goto_paragraph'), { index: 1 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('set_paragraph_style compiles', async () => {
      const script = await captureScript(findTool(paragraphTools, 'set_paragraph_style'), { index: 1, styleName: 'Heading 1' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('set style of p to');
    });
  });

  describe('Bookmark Tools', () => {
    test('list_bookmarks compiles', async () => {
      const script = await captureScript(findTool(bookmarkTools, 'list_bookmarks'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('create_bookmark compiles with pipe-quoted bookmark range', async () => {
      const script = await captureScript(findTool(bookmarkTools, 'create_bookmark'), { name: 'TestBM' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('|bookmark range|');
      expect(script).not.toMatch(/(?<!\|)bookmark range(?!\|)/);
    });

    test('goto_bookmark compiles with text object', async () => {
      const script = await captureScript(findTool(bookmarkTools, 'goto_bookmark'), { name: 'TestBM' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('text object of b');
      expect(script).not.toContain('bookmark range of b');
    });

    test('delete_bookmark compiles', async () => {
      const script = await captureScript(findTool(bookmarkTools, 'delete_bookmark'), { name: 'TestBM' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });
  });

  describe('Hyperlink Tools', () => {
    test('list_hyperlinks uses hyperlink objects with try/catch for text to display', async () => {
      const script = await captureScript(findTool(hyperlinkTools, 'list_hyperlinks'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('hyperlink objects of d');
      expect(script).toContain('hyperlink object i of d');
      expect(script).toContain('text to display of h');
      expect(script).toContain('try');
      expect(script).toContain('end try');
      expect(script).toContain('(no text)');
    });

    test('create_hyperlink compiles with tell selection', async () => {
      const script = await captureScript(findTool(hyperlinkTools, 'create_hyperlink'), { url: 'https://example.com' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('tell selection');
      expect(script).toContain('make new hyperlink object at end');
      expect(script).toContain('|hyperlink address|');
    });

    test('create_hyperlink with displayText', async () => {
      const script = await captureScript(findTool(hyperlinkTools, 'create_hyperlink'), { url: 'https://example.com', displayText: 'Click' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('|text to display|');
    });
  });

  describe('Navigation Tools', () => {
    test('goto_start compiles', async () => {
      const script = await captureScript(findTool(navigationTools, 'goto_start'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('goto_end compiles', async () => {
      const script = await captureScript(findTool(navigationTools, 'goto_end'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('get_selection_info compiles', async () => {
      const script = await captureScript(findTool(navigationTools, 'get_selection_info'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('select_all compiles', async () => {
      const script = await captureScript(findTool(navigationTools, 'select_all'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('move_cursor_after_text compiles', async () => {
      const script = await captureScript(findTool(navigationTools, 'move_cursor_after_text'), { searchText: 'test' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });
  });

  describe('Image Tools', () => {
    test('insert_image compiles', async () => {
      const script = await captureScript(findTool(imageTools, 'insert_image'), { path: '/tmp/test.png' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('paste object selection');
    });

    test('insert_image with resize compiles', async () => {
      const script = await captureScript(findTool(imageTools, 'insert_image'), { path: '/tmp/test.png', width: 200, height: 100 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('set width of shp to');
      expect(script).toContain('set height of shp to');
    });

    test('list_inline_shapes compiles', async () => {
      const script = await captureScript(findTool(imageTools, 'list_inline_shapes'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('resize_inline_shape compiles', async () => {
      const script = await captureScript(findTool(imageTools, 'resize_inline_shape'), { index: 1, width: 200 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('resize_inline_shape rejects non-integer index', async () => {
      await expect(findTool(imageTools, 'resize_inline_shape').handler({ index: 1.5, width: 200 })).rejects.toThrow('index must be an integer');
    });
  });

  describe('Audit Fixes', () => {
    test('move_cursor_after_text does not set unused prevStart', async () => {
      const script = await captureScript(findTool(navigationTools, 'move_cursor_after_text'), { searchText: 'test' });
      expect(script).not.toContain('set prevStart to');
    });

    test('open_document annotation is not readOnlyHint: true', () => {
      const tool = findTool(documentTools, 'open_document');
      expect(tool.annotations.readOnlyHint).toBe(false);
    });
  });

  describe('Template Engine', () => {
    test('escapes regex metacharacters in keys', () => {
      const result = processTemplate('<<key.name>>', { 'key.name': 'value' });
      expect(result).toBe('"value"');
    });

    test('replaces string params with escaped quotes', () => {
      const result = processTemplate('<<NAME>>', { NAME: 'hello "world"' });
      expect(result).toBe('"hello \\"world\\""');
    });

    test('replaces boolean params', () => {
      expect(processTemplate('<<FLAG>>', { FLAG: true })).toBe('true');
      expect(processTemplate('<<FLAG>>', { FLAG: false })).toBe('false');
    });

    test('replaces number params', () => {
      expect(processTemplate('<<NUM>>', { NUM: 42 })).toBe('42');
    });

    test('throws on unsupported value types (object)', () => {
      expect(() => processTemplate('<<OBJ>>', { OBJ: { a: 1 } })).toThrow('Unsupported template value type');
    });

    test('throws on unsupported value types (array)', () => {
      expect(() => processTemplate('<<ARR>>', { ARR: [1, 2] })).toThrow('Unsupported template value type');
    });
  });
});
