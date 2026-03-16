import { describe, test, expect, jest } from '@jest/globals';
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

const { textTools } = await import('../src/tools/word-text.js');
const { tableTools } = await import('../src/tools/word-tables.js');
const { paragraphTools } = await import('../src/tools/word-paragraphs.js');
const { bookmarkTools } = await import('../src/tools/word-bookmarks.js');
const { hyperlinkTools } = await import('../src/tools/word-hyperlinks.js');
const { navigationTools } = await import('../src/tools/word-navigation.js');
const { documentTools } = await import('../src/tools/word-documents.js');
const { imageTools } = await import('../src/tools/word-images.js');
const { wordWorkflowTools } = await import('../src/tools/word-workflows.js');
const { headerFooterTools } = await import('../src/tools/word-headers-footers.js');
const { sectionTools } = await import('../src/tools/word-sections.js');
const { formattingReadTools } = await import('../src/tools/word-formatting-read.js');
const { clipboardTools } = await import('../src/tools/word-clipboard.js');
const { processTemplate } = await import('../src/lib/applescript/template-engine.js');
const { COMMON_SCRIPTS, toAppleScriptString, escapeForWordFind, escapeAppleScriptString, buildWordExecuteFind } = await import('../src/lib/applescript/helpers.js');
const { WORD_FIND_MODES, WORD_FIND_STRATEGIES, buildWordFindScript } = await import('../src/lib/applescript/word-find.js');
const { excelCellTools } = await import('../src/tools/excel-cells.js');

function findTool(tools, name) {
  return tools.find(t => t.name === name);
}

async function captureScript(tool, args = {}) {
  capturedScripts._last = null;
  try {
    await tool.handler(args);
  } catch {}
  return capturedScripts._last;
}

function createFragmentFixture(ref, metadata) {
  const storeDir = join(tmpdir(), 'office-mcp-fragments');
  mkdirSync(storeDir, { recursive: true });
  if (metadata.filePath) {
    writeFileSync(metadata.filePath, 'fixture', 'utf8');
  }
  writeFileSync(join(storeDir, `${ref}.json`), JSON.stringify(metadata, null, 2), 'utf8');
}

describe('AppleScript Syntax Verification', () => {
  describe('Document Tools', () => {
    test('create_document compiles', async () => {
      const script = await captureScript(findTool(documentTools, 'word_create_document'), {});
      expect(script).toBeTruthy();
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('get_document_text compiles', async () => {
      const script = await captureScript(findTool(documentTools, 'word_get_document_text'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('get_document_info compiles', async () => {
      const script = await captureScript(findTool(documentTools, 'word_get_document_info'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('save_document compiles', async () => {
      const script = await captureScript(findTool(documentTools, 'word_save_document'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('save_document rejects empty path', async () => {
      await expect(findTool(documentTools, 'word_save_document').handler({ path: '' })).rejects.toThrow('path cannot be empty');
    });

    test('close_document compiles', async () => {
      const script = await captureScript(findTool(documentTools, 'word_close_document'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('export_pdf compiles', async () => {
      const script = await captureScript(findTool(documentTools, 'word_export_pdf'), { path: '/tmp/test.pdf' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('open_document compiles', async () => {
      const script = await captureScript(findTool(documentTools, 'word_open_document'), { path: '/tmp/test.docx' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });
  });

  describe('Text Tools', () => {
    test('insert_text compiles', async () => {
      const script = await captureScript(findTool(textTools, 'word_insert_text'), { text: 'Hello' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('insert_text accepts newline', async () => {
      const script = await captureScript(findTool(textTools, 'word_insert_text'), { text: '\n' });
      expect(script).toBeTruthy();
      expect(script).toContain('type text selection text');
    });

    test('insert_text accepts whitespace', async () => {
      const script = await captureScript(findTool(textTools, 'word_insert_text'), { text: '  ' });
      expect(script).toBeTruthy();
    });

    test('insert_text with quotes and backslash compiles', async () => {
      const script = await captureScript(findTool(textTools, 'word_insert_text'), { text: 'say "hello" and back\\slash' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('replace_text compiles', async () => {
      const script = await captureScript(findTool(textTools, 'word_replace_text'), { find: 'old', replace: 'new' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('execute find findObject');
      expect(script).toContain('find text "old"');
      expect(script).toContain('replace with "new"');
      expect(script).not.toContain('set content of findObject');
      expect(script).not.toContain('content of replacement of findObject');
    });

    test('replace_text resets search to document start', async () => {
      const script = await captureScript(findTool(textTools, 'word_replace_text'), { find: 'old', replace: 'new' });
      expect(script).toContain('select (text object of activeDoc)');
      expect(script).toContain('set selection end of selection to selection start of selection');
    });

    test('format_text compiles (bold)', async () => {
      const script = await captureScript(findTool(textTools, 'word_format_text'), { bold: true });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('format_text compiles (font + size)', async () => {
      const script = await captureScript(findTool(textTools, 'word_format_text'), { font: 'Arial', size: 14 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });
  });

  describe('Table Tools', () => {
    test('create_table compiles with correct target', async () => {
      const script = await captureScript(findTool(tableTools, 'word_create_table'), { rows: 3, columns: 4 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('text object of selection');
      expect(script).not.toMatch(/make new table at selection(?! )/);
    });

    test('list_tables compiles', async () => {
      const script = await captureScript(findTool(tableTools, 'word_list_tables'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('get_table_cell compiles', async () => {
      const script = await captureScript(findTool(tableTools, 'word_get_table_cell'), { tableIndex: 1, row: 1, column: 1 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('if length of cellText is 1 then');
    });

    test('set_table_cell compiles', async () => {
      const script = await captureScript(findTool(tableTools, 'word_set_table_cell'), { tableIndex: 1, row: 1, column: 1, text: 'Test' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('select_table_cell compiles', async () => {
      const script = await captureScript(findTool(tableTools, 'word_select_table_cell'), { tableIndex: 1, row: 1, column: 1 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('find_table_header compiles', async () => {
      const script = await captureScript(findTool(tableTools, 'word_find_table_header'), { tableIndex: 1, headerText: 'Name' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('if length of cellText is 1 then');
    });

    test('find_table_header with quotes compiles', async () => {
      const script = await captureScript(findTool(tableTools, 'word_find_table_header'), { tableIndex: 1, headerText: 'Column "A"' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('try');
      expect(script).toContain('on error');
    });

    test('add_table_row compiles with correct insert syntax', async () => {
      const script = await captureScript(findTool(tableTools, 'word_add_table_row'), { tableIndex: 1, afterRow: 2 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('insert rows selection position below');
      expect(script).toContain('select (text object of row');
    });

    test('add_table_row without afterRow adds at end', async () => {
      const script = await captureScript(findTool(tableTools, 'word_add_table_row'), { tableIndex: 1 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('count of rows of t');
    });

    test('delete_table_row compiles', async () => {
      const script = await captureScript(findTool(tableTools, 'word_delete_table_row'), { tableIndex: 1, row: 2 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('add_table_column compiles with correct insert syntax', async () => {
      const script = await captureScript(findTool(tableTools, 'word_add_table_column'), { tableIndex: 1, afterColumn: 2 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('insert columns selection position insert on the right');
    });

    test('delete_table_column compiles', async () => {
      const script = await captureScript(findTool(tableTools, 'word_delete_table_column'), { tableIndex: 1, column: 2 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('get_table_cell has error handling for cell access', async () => {
      const script = await captureScript(findTool(tableTools, 'word_get_table_cell'), { tableIndex: 1, row: 1, column: 1 });
      expect(script).toContain('try');
      expect(script).toContain('on error');
      expect(script).toContain('Cell not found');
    });

    test('cleanCellMarkers helper protects single marker cells', () => {
      expect(COMMON_SCRIPTS.cleanCellMarkers).toContain('if length of cellText is 1 then');
      expect(COMMON_SCRIPTS.cleanCellMarkers).toContain('set cellText to ""');
      expect(COMMON_SCRIPTS.cleanCellMarkers).toContain('set lastCharCode to ASCII number of (character -1 of cellText)');
    });

    test('set_table_cell has error handling', async () => {
      const script = await captureScript(findTool(tableTools, 'word_set_table_cell'), { tableIndex: 1, row: 1, column: 1, text: 'Test' });
      expect(script).toContain('try');
      expect(script).toContain('on error');
    });

    test('delete_table_row has error handling', async () => {
      const script = await captureScript(findTool(tableTools, 'word_delete_table_row'), { tableIndex: 1, row: 1 });
      expect(script).toContain('try');
      expect(script).toContain('on error');
    });
  });

  describe('Paragraph Tools', () => {
    test('list_paragraphs compiles with name local', async () => {
      const script = await captureScript(findTool(paragraphTools, 'word_list_paragraphs'), { limit: 10 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('name local of style of p');
      expect(script).not.toContain('paragraph style');
      expect(script).not.toContain('on replaceText');
    });

    test('goto_paragraph compiles', async () => {
      const script = await captureScript(findTool(paragraphTools, 'word_goto_paragraph'), { index: 1 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('set_paragraph_style compiles', async () => {
      const script = await captureScript(findTool(paragraphTools, 'word_set_paragraph_style'), { index: 1, styleName: 'Heading 1' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('set style of p to');
    });
  });

  describe('Bookmark Tools', () => {
    test('list_bookmarks compiles', async () => {
      const script = await captureScript(findTool(bookmarkTools, 'word_list_bookmarks'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('create_bookmark compiles with pipe-quoted bookmark range', async () => {
      const script = await captureScript(findTool(bookmarkTools, 'word_create_bookmark'), { name: 'TestBM' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('|bookmark range|');
      expect(script).not.toMatch(/(?<!\|)bookmark range(?!\|)/);
    });

    test('goto_bookmark compiles with text object', async () => {
      const script = await captureScript(findTool(bookmarkTools, 'word_goto_bookmark'), { name: 'TestBM' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('text object of b');
      expect(script).not.toContain('bookmark range of b');
    });

    test('delete_bookmark compiles', async () => {
      const script = await captureScript(findTool(bookmarkTools, 'word_delete_bookmark'), { name: 'TestBM' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('goto_bookmark has try/catch', async () => {
      const script = await captureScript(findTool(bookmarkTools, 'word_goto_bookmark'), { name: 'TestBM' });
      expect(script).toContain('try');
      expect(script).toContain('on error');
      expect(script).toContain('Bookmark not found');
    });

    test('delete_bookmark has try/catch', async () => {
      const script = await captureScript(findTool(bookmarkTools, 'word_delete_bookmark'), { name: 'TestBM' });
      expect(script).toContain('try');
      expect(script).toContain('on error');
      expect(script).toContain('Bookmark not found');
    });
  });

  describe('Hyperlink Tools', () => {
    test('list_hyperlinks uses hyperlink objects with try/catch for text to display', async () => {
      const script = await captureScript(findTool(hyperlinkTools, 'word_list_hyperlinks'), {});
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
      const script = await captureScript(findTool(hyperlinkTools, 'word_create_hyperlink'), { url: 'https://example.com' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('tell selection');
      expect(script).toContain('make new hyperlink object at end');
      expect(script).toContain('|hyperlink address|');
    });

    test('create_hyperlink falls back to URL display text without selection', async () => {
      const script = await captureScript(findTool(hyperlinkTools, 'word_create_hyperlink'), { url: 'https://example.com' });
      expect(script).toContain('else if hasSelection then');
      expect(script).toContain('|text to display|:"https://example.com"');
    });

    test('create_hyperlink with displayText', async () => {
      const script = await captureScript(findTool(hyperlinkTools, 'word_create_hyperlink'), { url: 'https://example.com', displayText: 'Click' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('|text to display|');
    });
  });

  describe('Navigation Tools', () => {
    test('goto_start compiles', async () => {
      const script = await captureScript(findTool(navigationTools, 'word_goto_start'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('goto_end compiles', async () => {
      const script = await captureScript(findTool(navigationTools, 'word_goto_end'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('get_selection_info compiles', async () => {
      const script = await captureScript(findTool(navigationTools, 'word_get_selection_info'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('select_all compiles', async () => {
      const script = await captureScript(findTool(navigationTools, 'word_select_all'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('move_cursor_after_text compiles', async () => {
      const script = await captureScript(findTool(navigationTools, 'word_move_cursor_after_text'), { searchText: 'test' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('execute find findObject');
      expect(script).toContain('match forward true');
      expect(script).toContain('wrap find find stop');
      expect(script).not.toContain('set content of findObject');
    });

    test('get_selection_info has error handling', async () => {
      const script = await captureScript(findTool(navigationTools, 'word_get_selection_info'), {});
      expect(script).toContain('try');
      expect(script).toContain('on error');
    });

    test('move_cursor_after_text has error handling for find object', async () => {
      const script = await captureScript(findTool(navigationTools, 'word_move_cursor_after_text'), { searchText: 'test' });
      expect(script).toContain('try');
      expect(script).toContain('on error');
    });
  });

  describe('Image Tools', () => {
    test('insert_image compiles', async () => {
      const script = await captureScript(findTool(imageTools, 'word_insert_image'), { path: '/tmp/test.png' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('paste object selection');
    });

    test('create_image_ref rejects unsupported formats without AppleScript execution', async () => {
      await expect(findTool(imageTools, 'word_create_image_ref').handler({ path: '/tmp/test.gif' })).rejects.toThrow(
        'path must point to a PNG, JPEG, or TIFF image'
      );
    });

    test('insert_image with resize compiles', async () => {
      const script = await captureScript(findTool(imageTools, 'word_insert_image'), { path: '/tmp/test.png', width: 200, height: 100 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('set width of shp to');
      expect(script).toContain('set height of shp to');
    });

    test('list_inline_shapes compiles', async () => {
      const script = await captureScript(findTool(imageTools, 'word_list_inline_shapes'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('resize_inline_shape compiles', async () => {
      const script = await captureScript(findTool(imageTools, 'word_resize_inline_shape'), { index: 1, width: 200 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('resize_inline_shape rejects non-integer index', async () => {
      await expect(findTool(imageTools, 'word_resize_inline_shape').handler({ index: 1.5, width: 200 })).rejects.toThrow('index must be an integer');
    });
  });

  describe('Audit Fixes', () => {
    test('move_cursor_after_text does not set unused prevStart', async () => {
      const script = await captureScript(findTool(navigationTools, 'word_move_cursor_after_text'), { searchText: 'test' });
      expect(script).not.toContain('set prevStart to');
    });

    test('open_document annotation is not readOnlyHint: true', () => {
      const tool = findTool(documentTools, 'word_open_document');
      expect(tool.annotations.readOnlyHint).toBe(false);
    });

    test('word_replace_text captures find result and reports not found', async () => {
      const script = await captureScript(findTool(textTools, 'word_replace_text'), { find: 'old', replace: 'new' });
      expect(script).toContain('set findResult to');
      expect(script).toContain('if findResult then');
      expect(script).toContain('Text not found, no replacements made');
      expect(script).toContain('replace with "new"');
      expect(script).not.toContain('set content of findObject');
    });

    test('word_delete_text with text captures find result and reports not found', async () => {
      const script = await captureScript(findTool(textTools, 'word_delete_text'), { text: 'gone' });
      expect(script).toContain('set findResult to');
      expect(script).toContain('if findResult then');
      expect(script).toContain('Text not found, nothing deleted');
      expect(script).toContain('replace with ""');
      expect(script).not.toContain('set content of findObject');
    });

    test('word_create_hyperlink uses quoteAppleScriptString not JSON.stringify', async () => {
      const script = await captureScript(findTool(hyperlinkTools, 'word_create_hyperlink'), { url: 'https://example.com', displayText: 'say "hello"' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('|hyperlink address|:"https://example.com"');
      expect(script).toContain('|text to display|:"say \\"hello\\""');
    });

    test('word_insert_header_image rejects non-numeric width', async () => {
      await expect(findTool(headerFooterTools, 'word_insert_header_image').handler({ path: '/tmp/test.png', width: 'abc' })).rejects.toThrow('width must be a valid number');
    });

    test('word_insert_footer_image rejects non-numeric height', async () => {
      await expect(findTool(headerFooterTools, 'word_insert_footer_image').handler({ path: '/tmp/test.png', height: 'abc' })).rejects.toThrow('height must be a valid number');
    });
  });

  describe('Header/Footer Tools', () => {
    test('word_get_header_text compiles', async () => {
      const script = await captureScript(findTool(headerFooterTools, 'word_get_header_text'), {});
      expect(script).toBeTruthy();
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('word_set_header_text compiles', async () => {
      const script = await captureScript(findTool(headerFooterTools, 'word_set_header_text'), { text: 'Test Header' });
      expect(script).toBeTruthy();
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('word_get_footer_text compiles', async () => {
      const script = await captureScript(findTool(headerFooterTools, 'word_get_footer_text'), {});
      expect(script).toBeTruthy();
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('word_set_footer_text compiles', async () => {
      const script = await captureScript(findTool(headerFooterTools, 'word_set_footer_text'), { text: 'Test Footer' });
      expect(script).toBeTruthy();
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('word_insert_header_image compiles', async () => {
      const script = await captureScript(findTool(headerFooterTools, 'word_insert_header_image'), { path: '/tmp/test.png' });
      expect(script).toBeTruthy();
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('word_insert_footer_image compiles', async () => {
      const script = await captureScript(findTool(headerFooterTools, 'word_insert_footer_image'), { path: '/tmp/test.png' });
      expect(script).toBeTruthy();
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('word_insert_header_image with quotes in path compiles', async () => {
      const script = await captureScript(findTool(headerFooterTools, 'word_insert_header_image'), { path: '/tmp/my "image".png' });
      expect(script).toBeTruthy();
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('word_get_header_text has error handling for header access', async () => {
      const script = await captureScript(findTool(headerFooterTools, 'word_get_header_text'), {});
      expect(script).toContain('try');
      expect(script).toContain('on error');
      expect(script).toContain('Header not available');
    });

    test('word_get_footer_text has error handling for footer access', async () => {
      const script = await captureScript(findTool(headerFooterTools, 'word_get_footer_text'), {});
      expect(script).toContain('try');
      expect(script).toContain('on error');
      expect(script).toContain('Footer not available');
    });

    test('word_set_header_text has error handling', async () => {
      const script = await captureScript(findTool(headerFooterTools, 'word_set_header_text'), { text: 'Test' });
      expect(script).toContain('try');
      expect(script).toContain('on error');
    });

    test('word_set_footer_text has error handling', async () => {
      const script = await captureScript(findTool(headerFooterTools, 'word_set_footer_text'), { text: 'Test' });
      expect(script).toContain('try');
      expect(script).toContain('on error');
    });
  });

  describe('Section Tools', () => {
    test('word_list_sections compiles', async () => {
      const script = await captureScript(findTool(sectionTools, 'word_list_sections'), {});
      expect(script).toBeTruthy();
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('word_get_section_info compiles', async () => {
      const script = await captureScript(findTool(sectionTools, 'word_get_section_info'), { index: 1 });
      expect(script).toBeTruthy();
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('word_set_page_setup compiles', async () => {
      const script = await captureScript(findTool(sectionTools, 'word_set_page_setup'), { topMargin: 72 });
      expect(script).toBeTruthy();
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('word_insert_section_break compiles', async () => {
      const script = await captureScript(findTool(sectionTools, 'word_insert_section_break'), {});
      expect(script).toBeTruthy();
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('word_insert_section_break rejects invalid type', async () => {
      await expect(findTool(sectionTools, 'word_insert_section_break').handler({ type: 'weird' })).rejects.toThrow(
        'type must be one of: next_page, continuous, even_page, odd_page'
      );
    });

    test('word_get_section_info has error handling', async () => {
      const script = await captureScript(findTool(sectionTools, 'word_get_section_info'), { index: 1 });
      expect(script).toContain('try');
      expect(script).toContain('on error');
    });

    test('word_set_page_setup has error handling', async () => {
      const script = await captureScript(findTool(sectionTools, 'word_set_page_setup'), { topMargin: 72 });
      expect(script).toContain('try');
      expect(script).toContain('on error');
    });

    test('word_set_page_setup rejects invalid orientation', async () => {
      await expect(findTool(sectionTools, 'word_set_page_setup').handler({ orientation: 'sideways' })).rejects.toThrow(
        'orientation must be one of: portrait, landscape'
      );
    });
  });

  describe('Formatting Read Tools', () => {
    test('word_get_text_formatting compiles', async () => {
      const script = await captureScript(findTool(formattingReadTools, 'word_get_text_formatting'), {});
      expect(script).toBeTruthy();
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('word_get_text_formatting has error handling', async () => {
      const script = await captureScript(findTool(formattingReadTools, 'word_get_text_formatting'), {});
      expect(script).toContain('try');
      expect(script).toContain('on error');
    });

    test('word_get_paragraph_formatting compiles', async () => {
      const script = await captureScript(findTool(formattingReadTools, 'word_get_paragraph_formatting'), {});
      expect(script).toBeTruthy();
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('word_get_paragraph_formatting has error handling', async () => {
      const script = await captureScript(findTool(formattingReadTools, 'word_get_paragraph_formatting'), {});
      expect(script).toContain('try');
      expect(script).toContain('end try');
    });
  });

  describe('Delete Tools', () => {
    test('word_delete_text with text compiles', async () => {
      const script = await captureScript(findTool(textTools, 'word_delete_text'), { text: 'test' });
      expect(script).toBeTruthy();
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('execute find findObject');
      expect(script).toContain('find text "test"');
      expect(script).toContain('replace with ""');
      expect(script).not.toContain('set content of findObject');
    });

    test('word_delete_text with text resets search to document start', async () => {
      const script = await captureScript(findTool(textTools, 'word_delete_text'), { text: 'test' });
      expect(script).toContain('select (text object of activeDoc)');
      expect(script).toContain('set selection end of selection to selection start of selection');
    });

    test('word_delete_text without text compiles', async () => {
      const script = await captureScript(findTool(textTools, 'word_delete_text'), {});
      expect(script).toBeTruthy();
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('word_delete_paragraph compiles', async () => {
      const script = await captureScript(findTool(paragraphTools, 'word_delete_paragraph'), { index: 1 });
      expect(script).toBeTruthy();
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });
  });

  describe('Clipboard Tools', () => {
    test('word_copy_content without params compiles', async () => {
      const script = await captureScript(findTool(clipboardTools, 'word_copy_content'), {});
      expect(script).toBeTruthy();
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('copy object selection');
    });

    test('word_copy_content with paragraph range compiles', async () => {
      const script = await captureScript(findTool(clipboardTools, 'word_copy_content'), { startParagraph: 2, endParagraph: 5 });
      expect(script).toBeTruthy();
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('copy object selection');
      expect(script).toContain('set rStart to selection start of selection');
      expect(script).toContain('set rEnd to selection end of selection');
    });

    test('word_copy_content supports document scope', async () => {
      const script = await captureScript(findTool(clipboardTools, 'word_copy_content'), { scope: 'document' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('select (text object of d)');
    });

    test('word_copy_content supports inline_shape scope', async () => {
      const script = await captureScript(findTool(clipboardTools, 'word_copy_content'), { scope: 'inline_shape', inlineShapeIndex: 2 });
      expect(script).toContain('inline shape 2 of d');
      expect(script).toContain('copy object selection');
    });

    test('word_copy_content with single paragraph compiles', async () => {
      const script = await captureScript(findTool(clipboardTools, 'word_copy_content'), { startParagraph: 3 });
      expect(script).toBeTruthy();
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('word_copy_content rejects endParagraph without startParagraph', async () => {
      await expect(findTool(clipboardTools, 'word_copy_content').handler({ endParagraph: 5 })).rejects.toThrow('endParagraph requires startParagraph');
    });

    test('word_copy_content rejects endParagraph < startParagraph', async () => {
      await expect(findTool(clipboardTools, 'word_copy_content').handler({ startParagraph: 5, endParagraph: 2 })).rejects.toThrow('endParagraph must be >= startParagraph');
    });

    test('word_paste_content compiles', async () => {
      const script = await captureScript(findTool(clipboardTools, 'word_paste_content'), {});
      expect(script).toBeTruthy();
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('paste object selection');
    });

    test('word_capture_content_ref compiles', async () => {
      await expect(findTool(clipboardTools, 'word_capture_content_ref').handler({ scope: 'document' })).rejects.toMatchObject({
        code: 'NOT_SUPPORTED'
      });
    });

    test('word_insert_content_ref compiles for image refs', async () => {
      const ref = 'wordimg_demo';
      const filePath = '/tmp/test.png';
      createFragmentFixture(ref, {
        ref,
        app: 'word',
        kind: 'image_file',
        format: 'png',
        filePath,
        summary: { label: 'fixture' },
        createdAt: new Date().toISOString(),
        expiresAt: new Date(Date.now() + 60_000).toISOString()
      });

      const script = await captureScript(findTool(clipboardTools, 'word_insert_content_ref'), { ref: 'wordimg_demo' });
      expect(script).toBeTruthy();
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('paste object selection');
    });
  });

  describe('Workflow Tools', () => {
    test('word_copy_story_content compiles for header scope', async () => {
      const script = await captureScript(findTool(wordWorkflowTools, 'word_copy_story_content'), {
        scope: 'header',
        section: 1,
        type: 'primary'
      });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('set storyRange to text object of storyRef');
      expect(script).toContain('copy object selection');
    });

    test('word_clear_story_content compiles for body scope', async () => {
      const script = await captureScript(findTool(wordWorkflowTools, 'word_clear_story_content'), { scope: 'body' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('set content of storyRange to ""');
    });

    test('word_set_story_text compiles for footer scope', async () => {
      const script = await captureScript(findTool(wordWorkflowTools, 'word_set_story_text'), {
        scope: 'footer',
        section: 1,
        type: 'primary',
        text: 'Footer text'
      });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('set content of storyRange to');
      expect(script).toContain('Footer text');
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

    test('handles strings with newlines via toAppleScriptString', () => {
      const result = processTemplate('set x to <<TEXT>>', { TEXT: 'line1\nline2' });
      expect(result).toContain('& return &');
      expect(result).not.toContain('\\n');
    });
  });

  describe('Helper Functions', () => {
    test('escapeAppleScriptString escapes backslashes and quotes', () => {
      expect(escapeAppleScriptString('hello "world"')).toBe('hello \\"world\\"');
      expect(escapeAppleScriptString('back\\slash')).toBe('back\\\\slash');
    });

    test('toAppleScriptString handles single-line text', () => {
      expect(toAppleScriptString('hello')).toBe('"hello"');
    });

    test('toAppleScriptString handles multi-line text', () => {
      const result = toAppleScriptString('line1\nline2');
      expect(result).toBe('("line1" & return & "line2")');
    });

    test('toAppleScriptString handles multiple newlines', () => {
      const result = toAppleScriptString('a\nb\nc');
      expect(result).toBe('("a" & return & "b" & return & "c")');
    });

    test('toAppleScriptString handles \\r\\n', () => {
      const result = toAppleScriptString('a\r\nb');
      expect(result).toBe('("a" & return & "b")');
    });

    test('toAppleScriptString escapes quotes in multi-line text', () => {
      const result = toAppleScriptString('say "hi"\ngoodbye');
      expect(result).toBe('("say \\"hi\\"" & return & "goodbye")');
    });

    test('escapeForWordFind handles single-line text', () => {
      expect(escapeForWordFind('hello')).toBe('"hello"');
    });

    test('escapeForWordFind converts newlines to ^p', () => {
      expect(escapeForWordFind('line1\nline2')).toBe('"line1^pline2"');
    });

    test('escapeForWordFind converts \\r\\n to ^p', () => {
      expect(escapeForWordFind('line1\r\nline2')).toBe('"line1^pline2"');
    });

    test('escapeForWordFind converts \\r to ^p', () => {
      expect(escapeForWordFind('line1\rline2')).toBe('"line1^pline2"');
    });

    test('escapeForWordFind escapes quotes', () => {
      expect(escapeForWordFind('say "hi"')).toBe('"say \\"hi\\""');
    });

    test('buildWordExecuteFind builds direct execute-find command with parameters', () => {
      expect(
        buildWordExecuteFind('findObj', {
          findText: 'hello\nworld',
          replaceWith: 'new value',
          replace: 'replace all',
          matchForward: true,
          wrapFind: 'find stop'
        })
      ).toBe('execute find findObj find text "hello^pworld" match forward true wrap find find stop replace with "new value" replace replace all');
    });

    test('buildWordFindScript compiles direct strategy for replace', () => {
      const script = buildWordFindScript({
        strategy: WORD_FIND_STRATEGIES.DIRECT_EXECUTE_PARAMS,
        mode: WORD_FIND_MODES.REPLACE,
        findText: 'old value',
        replaceWith: 'new value'
      });
      expect(script).toContain('execute find findObject');
      expect(script).toContain('find text "old value"');
      expect(script).toContain('replace with "new value"');
      expect(script).not.toContain('set content of findObject');
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('buildWordFindScript compiles legacy strategy for replace', () => {
      const script = buildWordFindScript({
        strategy: WORD_FIND_STRATEGIES.LEGACY_FIND_OBJECT_CONTENT,
        mode: WORD_FIND_MODES.REPLACE,
        findText: 'old value',
        replaceWith: 'new value'
      });
      expect(script).toContain('set content of findObject to "old value"');
      expect(script).toContain('set content of replacement of findObject to "new value"');
      expect(script).toContain('set findResult to execute find findObject replace replace all');
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });
  });

  describe('Multiline Text in Tools', () => {
    test('word_delete_text with multiline text uses ^p', async () => {
      const script = await captureScript(findTool(textTools, 'word_delete_text'), { text: 'hello\nworld' });
      expect(script).toContain('hello^pworld');
      expect(script).not.toContain('\\n');
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('word_replace_text with multiline find uses ^p', async () => {
      const script = await captureScript(findTool(textTools, 'word_replace_text'), { find: 'old\ntext', replace: 'new\ntext' });
      expect(script).toContain('old^ptext');
      expect(script).toContain('new^ptext');
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('word_replace_text with long single-line find and replace uses execute find parameters', async () => {
      const longFind = 'automation incident details '.repeat(12).trim();
      const longReplace = '[describe incident impact] '.repeat(8).trim();
      const script = await captureScript(findTool(textTools, 'word_replace_text'), { find: longFind, replace: longReplace });
      expect(script).toContain(`find text "${longFind}"`);
      expect(script).toContain(`replace with "${longReplace}"`);
      expect(script).not.toContain('set content of findObject');
      expect(script).not.toContain('content of replacement of findObject');
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('word_insert_text with multiline text uses return concatenation', async () => {
      const script = await captureScript(findTool(textTools, 'word_insert_text'), { text: 'hello\nworld' });
      expect(script).toContain('& return &');
      expect(script).toContain('"hello"');
      expect(script).toContain('"world"');
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('word_move_cursor_after_text with multiline text uses ^p', async () => {
      const script = await captureScript(findTool(navigationTools, 'word_move_cursor_after_text'), { searchText: 'hello\nworld' });
      expect(script).toContain('hello^pworld');
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('word_set_table_cell with multiline text uses return', async () => {
      const script = await captureScript(findTool(tableTools, 'word_set_table_cell'), { tableIndex: 1, row: 1, column: 1, text: 'line1\nline2' });
      expect(script).toContain('& return &');
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('word_create_document with multiline content uses return', async () => {
      const script = await captureScript(findTool(documentTools, 'word_create_document'), { content: 'Hello\nWorld' });
      expect(script).toContain('& return &');
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('word_set_header_text with multiline text uses return', async () => {
      const script = await captureScript(findTool(headerFooterTools, 'word_set_header_text'), { text: 'Line 1\nLine 2' });
      expect(script).toContain('& return &');
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('word_set_footer_text with multiline text uses return', async () => {
      const script = await captureScript(findTool(headerFooterTools, 'word_set_footer_text'), { text: 'Footer\nLine 2' });
      expect(script).toContain('& return &');
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('excel_set_cell with multiline string uses return', async () => {
      const script = await captureScript(findTool(excelCellTools, 'excel_set_cell'), { cell: 'A1', value: 'line1\nline2' });
      expect(script).toContain('& return &');
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('excel_find_cell does not use Word-specific ^p markers', async () => {
      const script = await captureScript(findTool(excelCellTools, 'excel_find_cell'), { searchText: 'hello' });
      expect(script).not.toContain('^p');
      expect(script).toContain('try');
      expect(script).toContain('on error');
    });
  });
});
