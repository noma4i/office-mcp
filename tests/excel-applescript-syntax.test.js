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

const { excelWorkbookTools } = await import('../src/tools/excel-workbooks.js');
const { excelSheetTools } = await import('../src/tools/excel-sheets.js');
const { excelCellTools } = await import('../src/tools/excel-cells.js');
const { excelFormattingTools } = await import('../src/tools/excel-formatting.js');
const { excelRowColumnTools } = await import('../src/tools/excel-rows-columns.js');
const { excelDataTools } = await import('../src/tools/excel-data.js');
const { excelClipboardTools } = await import('../src/tools/excel-clipboard.js');

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

describe('Excel AppleScript Syntax Verification', () => {
  describe('Workbook Tools', () => {
    test('excel_create_workbook compiles', async () => {
      const script = await captureScript(findTool(excelWorkbookTools, 'excel_create_workbook'), {});
      expect(script).toBeTruthy();
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('excel_open_workbook compiles', async () => {
      const script = await captureScript(findTool(excelWorkbookTools, 'excel_open_workbook'), { path: '/tmp/test.xlsx' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('open workbook workbook file name');
    });

    test('excel_get_workbook_info compiles', async () => {
      const script = await captureScript(findTool(excelWorkbookTools, 'excel_get_workbook_info'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('excel_save_workbook compiles (no path)', async () => {
      const script = await captureScript(findTool(excelWorkbookTools, 'excel_save_workbook'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('excel_save_workbook compiles (with path) and saves workbook not sheet', async () => {
      const script = await captureScript(findTool(excelWorkbookTools, 'excel_save_workbook'), { path: '/tmp/test.xlsx' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('save workbook as wb filename');
      expect(script).not.toContain('save as ws');
      expect(script).not.toContain('set ws to active sheet');
    });

    test('excel_save_workbook with path uses try/finally for display alerts', async () => {
      const script = await captureScript(findTool(excelWorkbookTools, 'excel_save_workbook'), { path: '/tmp/test.xlsx' });
      expect(script).toContain('set display alerts to false');
      expect(script).toContain('on error errMsg');
      expect(script).toContain('set display alerts to true');
    });

    test('excel_save_workbook rejects empty path', async () => {
      await expect(findTool(excelWorkbookTools, 'excel_save_workbook').handler({ path: '' })).rejects.toThrow('path cannot be empty');
    });

    test('excel_close_workbook compiles', async () => {
      const script = await captureScript(findTool(excelWorkbookTools, 'excel_close_workbook'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('excel_list_workbooks compiles', async () => {
      const script = await captureScript(findTool(excelWorkbookTools, 'excel_list_workbooks'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });
  });

  describe('Sheet Tools', () => {
    test('excel_list_sheets compiles', async () => {
      const script = await captureScript(findTool(excelSheetTools, 'excel_list_sheets'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('excel_create_sheet compiles (no args)', async () => {
      const script = await captureScript(findTool(excelSheetTools, 'excel_create_sheet'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('make new worksheet');
    });

    test('excel_create_sheet compiles (with name)', async () => {
      const script = await captureScript(findTool(excelSheetTools, 'excel_create_sheet'), { name: 'DataSheet' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('excel_create_sheet uses dedicated after syntax', async () => {
      const script = await captureScript(findTool(excelSheetTools, 'excel_create_sheet'), { afterIndex: 1 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('make new worksheet after worksheet 1 of wb');
      expect(script).not.toContain('make new worksheet at after');
    });

    test('excel_delete_sheet compiles (by index)', async () => {
      const script = await captureScript(findTool(excelSheetTools, 'excel_delete_sheet'), { nameOrIndex: 1 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('excel_delete_sheet compiles (by name)', async () => {
      const script = await captureScript(findTool(excelSheetTools, 'excel_delete_sheet'), { nameOrIndex: 'Sheet1' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('excel_delete_sheet uses try/finally for display alerts', async () => {
      const script = await captureScript(findTool(excelSheetTools, 'excel_delete_sheet'), { nameOrIndex: 1 });
      expect(script).toContain('set display alerts to false');
      expect(script).toContain('on error errMsg');
      expect(script).toContain('set display alerts to true');
    });

    test('excel_rename_sheet compiles', async () => {
      const script = await captureScript(findTool(excelSheetTools, 'excel_rename_sheet'), { nameOrIndex: 1, newName: 'NewName' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('excel_activate_sheet compiles', async () => {
      const script = await captureScript(findTool(excelSheetTools, 'excel_activate_sheet'), { nameOrIndex: 'Sheet1' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('activate object');
    });

    test('excel_get_sheet_info compiles', async () => {
      const script = await captureScript(findTool(excelSheetTools, 'excel_get_sheet_info'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('used range');
    });

    test('excel_rename_sheet has error handling for sheet access', async () => {
      const script = await captureScript(findTool(excelSheetTools, 'excel_rename_sheet'), { nameOrIndex: 1, newName: 'Test' });
      expect(script).toContain('try');
      expect(script).toContain('on error');
      expect(script).toContain('Sheet not found');
    });

    test('excel_activate_sheet has error handling for sheet access', async () => {
      const script = await captureScript(findTool(excelSheetTools, 'excel_activate_sheet'), { nameOrIndex: 'Sheet1' });
      expect(script).toContain('try');
      expect(script).toContain('on error');
      expect(script).toContain('Sheet not found');
    });

    test('excel_activate_sheet keeps numeric string as sheet name', async () => {
      const script = await captureScript(findTool(excelSheetTools, 'excel_activate_sheet'), { nameOrIndex: '2024' });
      expect(script).toContain('worksheet "2024" of wb');
      expect(script).not.toContain('worksheet 2024 of wb');
    });
  });

  describe('Cell Tools', () => {
    test('excel_get_cell compiles', async () => {
      const script = await captureScript(findTool(excelCellTools, 'excel_get_cell'), { cell: 'A1' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('value of cell');
    });

    test('excel_set_cell compiles (string)', async () => {
      const script = await captureScript(findTool(excelCellTools, 'excel_set_cell'), { cell: 'A1', value: 'Hello' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('excel_set_cell compiles (number)', async () => {
      const script = await captureScript(findTool(excelCellTools, 'excel_set_cell'), { cell: 'B2', value: 42 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('excel_get_range compiles', async () => {
      const script = await captureScript(findTool(excelCellTools, 'excel_get_range'), { range: 'A1:B3' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('excel_set_cell_formula compiles', async () => {
      const script = await captureScript(findTool(excelCellTools, 'excel_set_cell_formula'), { cell: 'A1', formula: '=SUM(B1:B10)' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('set formula of cell');
      expect(script).not.toContain('formula value');
    });

    test('excel_set_cell_formula auto-prefixes = when missing', async () => {
      const script = await captureScript(findTool(excelCellTools, 'excel_set_cell_formula'), { cell: 'A1', formula: 'SUM(B1:B10)' });
      expect(script).toContain('=SUM(B1:B10)');
    });

    test('excel_set_cell_formula does not double-prefix =', async () => {
      const script = await captureScript(findTool(excelCellTools, 'excel_set_cell_formula'), { cell: 'A1', formula: '=SUM(B1:B10)' });
      expect(script).not.toContain('==SUM');
    });

    test('excel_clear_range compiles', async () => {
      const script = await captureScript(findTool(excelCellTools, 'excel_clear_range'), { range: 'A1:B3' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('clear contents');
    });

    test('excel_get_used_range compiles', async () => {
      const script = await captureScript(findTool(excelCellTools, 'excel_get_used_range'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('excel_find_cell compiles', async () => {
      const script = await captureScript(findTool(excelCellTools, 'excel_find_cell'), { searchText: 'test' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('find');
      expect(script).toContain('try');
      expect(script).toContain('on error');
      expect(script).toContain('end try');
    });

    test('excel_find_cell with quotes in searchText compiles', async () => {
      const script = await captureScript(findTool(excelCellTools, 'excel_find_cell'), { searchText: 'he said "hello"' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('he said \\"hello\\"');
    });

    test('excel_find_cell validates optional range before script generation', async () => {
      await expect(findTool(excelCellTools, 'excel_find_cell').handler({ searchText: 'x', range: 'bad range' })).rejects.toThrow(
        'range must be a valid Excel range reference'
      );
    });

    test('excel_get_cell has error handling', async () => {
      const script = await captureScript(findTool(excelCellTools, 'excel_get_cell'), { cell: 'A1' });
      expect(script).toContain('try');
      expect(script).toContain('on error');
    });

    test('excel_get_range has error handling for range access', async () => {
      const script = await captureScript(findTool(excelCellTools, 'excel_get_range'), { range: 'A1:B3' });
      expect(script).toContain('try');
      expect(script).toContain('on error');
      expect(script).toContain('Invalid range');
    });

    test('excel_set_cell_formula with quotes in formula compiles', async () => {
      const script = await captureScript(findTool(excelCellTools, 'excel_set_cell_formula'), { cell: 'A1', formula: '=IF(A1="yes","da","net")' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('excel_get_cell rejects invalid cell reference', async () => {
      await expect(findTool(excelCellTools, 'excel_get_cell').handler({ cell: 'A0' })).rejects.toThrow(
        'cell must be a valid A1 cell reference'
      );
    });
  });

  describe('Formatting Tools', () => {
    test('excel_format_cells compiles (bold)', async () => {
      const script = await captureScript(findTool(excelFormattingTools, 'excel_format_cells'), { range: 'A1', bold: true });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('bold of font object');
    });

    test('excel_format_cells compiles (font + size + color)', async () => {
      const script = await captureScript(findTool(excelFormattingTools, 'excel_format_cells'), { range: 'A1:B2', font: 'Arial', size: 14, fontColor: [255, 0, 0] });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('font size of font object');
    });

    test('excel_set_number_format compiles', async () => {
      const script = await captureScript(findTool(excelFormattingTools, 'excel_set_number_format'), { range: 'A1:A10', format: '#,##0.00' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('number format');
    });

    test('excel_set_number_format with quotes in format compiles', async () => {
      const script = await captureScript(findTool(excelFormattingTools, 'excel_set_number_format'), { range: 'A1', format: '0.00"kg"' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
    });

    test('excel_format_cells rejects invalid fontColor', async () => {
      await expect(findTool(excelFormattingTools, 'excel_format_cells').handler({ range: 'A1', fontColor: [256, 0, 0] })).rejects.toThrow(
        'fontColor must be an array of 3 numbers [R,G,B] with values 0-255'
      );
    });

    test('excel_format_cells rejects non-boolean bold', async () => {
      await expect(findTool(excelFormattingTools, 'excel_format_cells').handler({ range: 'A1', bold: 'true' })).rejects.toThrow(
        'bold must be a boolean'
      );
    });

    test('excel_set_cell_color rejects invalid color', async () => {
      await expect(findTool(excelFormattingTools, 'excel_set_cell_color').handler({ range: 'A1', color: [255, -1, 0] })).rejects.toThrow(
        'color must be an array of 3 numbers [R,G,B] with values 0-255'
      );
    });

    test('excel_set_cell_color compiles', async () => {
      const script = await captureScript(findTool(excelFormattingTools, 'excel_set_cell_color'), { range: 'A1', color: [255, 255, 0] });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('interior object');
    });

    test('excel_merge_cells compiles', async () => {
      const script = await captureScript(findTool(excelFormattingTools, 'excel_merge_cells'), { range: 'A1:B1' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('merge');
    });

    test('excel_autofit compiles', async () => {
      const script = await captureScript(findTool(excelFormattingTools, 'excel_autofit'), { range: 'A:C' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('autofit');
    });

    test('excel_format_cells has error handling for range and formatting', async () => {
      const script = await captureScript(findTool(excelFormattingTools, 'excel_format_cells'), { range: 'A1', bold: true });
      expect(script).toContain('try');
      expect(script).toContain('on error');
    });

    test('excel_set_number_format has error handling', async () => {
      const script = await captureScript(findTool(excelFormattingTools, 'excel_set_number_format'), { range: 'A1', format: '#,##0' });
      expect(script).toContain('try');
      expect(script).toContain('on error');
    });
  });

  describe('Row & Column Tools', () => {
    test('excel_insert_rows compiles', async () => {
      const script = await captureScript(findTool(excelRowColumnTools, 'excel_insert_rows'), { row: 3 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('insert into range');
      expect(script).toContain('shift shift down');
    });

    test('excel_insert_rows compiles (multiple)', async () => {
      const script = await captureScript(findTool(excelRowColumnTools, 'excel_insert_rows'), { row: 2, count: 3 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('2:4');
    });

    test('excel_delete_rows compiles', async () => {
      const script = await captureScript(findTool(excelRowColumnTools, 'excel_delete_rows'), { row: 2 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('delete range');
      expect(script).toContain('shift shift up');
    });

    test('excel_insert_columns compiles', async () => {
      const script = await captureScript(findTool(excelRowColumnTools, 'excel_insert_columns'), { column: 2 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('B:B');
      expect(script).toContain('shift shift to right');
    });

    test('excel_delete_columns compiles', async () => {
      const script = await captureScript(findTool(excelRowColumnTools, 'excel_delete_columns'), { column: 3 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('C:C');
      expect(script).toContain('shift shift to left');
    });

    test('excel_set_column_width compiles', async () => {
      const script = await captureScript(findTool(excelRowColumnTools, 'excel_set_column_width'), { column: 1, width: 25 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('column width');
    });

    test('excel_set_row_height compiles', async () => {
      const script = await captureScript(findTool(excelRowColumnTools, 'excel_set_row_height'), { row: 1, height: 30 });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('row height');
    });
  });

  describe('Data Tools', () => {
    test('excel_sort_range compiles', async () => {
      const script = await captureScript(findTool(excelDataTools, 'excel_sort_range'), { range: 'A1:B10', keyCell: 'B1' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('sort range');
      expect(script).toContain('key1');
      expect(script).toContain('order1');
    });

    test('excel_sort_range compiles (descending)', async () => {
      const script = await captureScript(findTool(excelDataTools, 'excel_sort_range'), { range: 'A1:C5', keyCell: 'A1', ascending: false, hasHeader: false });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('sort descending');
      expect(script).toContain('header header no');
    });

    test('excel_calculate compiles', async () => {
      const script = await captureScript(findTool(excelDataTools, 'excel_calculate'), {});
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('calculate');
    });

    test('excel_export_csv compiles and saves workbook not sheet', async () => {
      const script = await captureScript(findTool(excelDataTools, 'excel_export_csv'), { path: '/tmp/test.csv' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('CSV file format');
      expect(script).toContain('save workbook as wb filename');
      expect(script).not.toContain('save as ws');
      expect(script).not.toContain('set ws to active sheet');
    });

    test('excel_export_csv uses try/finally for display alerts', async () => {
      const script = await captureScript(findTool(excelDataTools, 'excel_export_csv'), { path: '/tmp/test.csv' });
      expect(script).toContain('set display alerts to false');
      expect(script).toContain('on error errMsg');
      expect(script).toContain('set display alerts to true');
    });

    test('excel_sort_range has error handling', async () => {
      const script = await captureScript(findTool(excelDataTools, 'excel_sort_range'), { range: 'A1:B10', keyCell: 'B1' });
      expect(script).toContain('try');
      expect(script).toContain('on error');
    });
  });

  describe('Clipboard Tools', () => {
    test('excel_copy_range compiles', async () => {
      const script = await captureScript(findTool(excelClipboardTools, 'excel_copy_range'), { range: 'A1:C5' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('System Events');
      expect(script).toContain('keystroke "c"');
    });

    test('excel_paste_range compiles', async () => {
      const script = await captureScript(findTool(excelClipboardTools, 'excel_paste_range'), { targetCell: 'D2' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('keystroke "v"');
    });

    test('excel_capture_range_ref compiles', async () => {
      const script = await captureScript(findTool(excelClipboardTools, 'excel_capture_range_ref'), { range: 'A1:B3' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('save workbook as fragWb filename');
    });

    test('excel_insert_range_ref compiles', async () => {
      const ref = 'excelfrag_demo';
      const filePath = join(tmpdir(), `${ref}.xlsx`);
      createFragmentFixture(ref, {
        ref,
        app: 'excel',
        kind: 'excel_range',
        format: 'xlsx',
        filePath,
        summary: { label: 'fixture' },
        createdAt: new Date().toISOString(),
        expiresAt: new Date(Date.now() + 60_000).toISOString()
      });

      const script = await captureScript(findTool(excelClipboardTools, 'excel_insert_range_ref'), { ref: 'excelfrag_demo', targetCell: 'B5' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('open workbook workbook file name');
      expect(script).toContain('keystroke "c"');
      expect(script).toContain('keystroke "v"');
    });
  });

  describe('Tool Count', () => {
    test('all 37 Excel tools are defined', () => {
      const total =
        excelWorkbookTools.length +
        excelSheetTools.length +
        excelCellTools.length +
        excelFormattingTools.length +
        excelRowColumnTools.length +
        excelDataTools.length +
        excelClipboardTools.length;
      expect(total).toBe(37);
    });

    test('all Excel tools have excel_ prefix', () => {
      const allTools = [
        ...excelWorkbookTools,
        ...excelSheetTools,
        ...excelCellTools,
        ...excelFormattingTools,
        ...excelRowColumnTools,
        ...excelDataTools,
        ...excelClipboardTools
      ];
      allTools.forEach(tool => {
        expect(tool.name).toMatch(/^excel_/);
      });
    });
  });
});
