import { describe, expect, test } from '@jest/globals';
import { wrapExcelScript, wrapWordScript } from '../src/lib/applescript/script-wrappers.js';

describe('AppleScript Wrappers', () => {
  test('wrapWordScript adds document guard by default', () => {
    const script = wrapWordScript('return "ok"');
    expect(script).toContain('tell application "Microsoft Word"');
    expect(script).toContain('No document is open');
    expect(script).toContain('return "ok"');
  });

  test('wrapWordScript can skip document guard and activate app', () => {
    const script = wrapWordScript('return "ok"', { activate: true, requireDocument: false });
    expect(script).toContain('tell application "Microsoft Word"');
    expect(script).toContain('activate');
    expect(script).not.toContain('No document is open');
  });

  test('wrapExcelScript sets workbook guard and active sheet', () => {
    const script = wrapExcelScript('return "ok"', { setActiveSheet: true });
    expect(script).toContain('tell application "Microsoft Excel"');
    expect(script).toContain('No workbook is open');
    expect(script).toContain('set ws to active sheet');
  });

  test('wrapExcelScript can include active workbook binding', () => {
    const script = wrapExcelScript('return "ok"', { setActiveWorkbook: true, requireWorkbook: false });
    expect(script).toContain('set wb to active workbook');
    expect(script).not.toContain('No workbook is open');
  });
});

