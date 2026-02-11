import { describe, test, expect, jest } from '@jest/globals';
import { execSync } from 'child_process';
import { writeFileSync, unlinkSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';

let capturedScripts = {};

jest.unstable_mockModule('../src/lib/applescript/executor.js', () => ({
  runAppleScript: async (script) => {
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
    try { unlinkSync(tmpFile); } catch {}
  }
}

async function captureScript(tool, args = {}) {
  capturedScripts._last = null;
  try { await tool.handler(args); } catch {}
  return capturedScripts._last;
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

    test('excel_save_workbook compiles (with path)', async () => {
      const script = await captureScript(findTool(excelWorkbookTools, 'excel_save_workbook'), { path: '/tmp/test.xlsx' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('save as ws filename');
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

    test('excel_export_csv compiles', async () => {
      const script = await captureScript(findTool(excelDataTools, 'excel_export_csv'), { path: '/tmp/test.csv' });
      const result = compileAppleScript(script);
      expect(result.ok).toBe(true);
      expect(script).toContain('CSV file format');
    });
  });

  describe('Tool Count', () => {
    test('all 33 Excel tools are defined', () => {
      const total = excelWorkbookTools.length + excelSheetTools.length + excelCellTools.length + excelFormattingTools.length + excelRowColumnTools.length + excelDataTools.length;
      expect(total).toBe(33);
    });

    test('all Excel tools have excel_ prefix', () => {
      const allTools = [...excelWorkbookTools, ...excelSheetTools, ...excelCellTools, ...excelFormattingTools, ...excelRowColumnTools, ...excelDataTools];
      allTools.forEach(tool => {
        expect(tool.name).toMatch(/^excel_/);
      });
    });
  });
});
