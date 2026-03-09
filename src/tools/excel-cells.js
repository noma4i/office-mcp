import { validateExcelCellReference, validateExcelRangeReference, validateString } from '../lib/validators.js';
import { runAppleScript } from '../lib/applescript/executor.js';
import { toAppleScriptString, quoteAppleScriptString } from '../lib/applescript/helpers.js';
import { wrapExcelScript } from '../lib/applescript/script-wrappers.js';

export const excelCellTools = [
  {
    name: 'excel_get_cell',
    description: 'Get the value of a cell in Excel (A1 notation)',
    annotations: { readOnlyHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        cell: {
          type: 'string',
          description: 'Cell reference in A1 notation (e.g., "A1", "B2")'
        }
      },
      required: ['cell']
    },
    async handler(args) {
      const cell = validateExcelCellReference(args.cell, 'cell');
      const script = wrapExcelScript(
        `
try
  set v to value of cell ${JSON.stringify(cell)} of ws
on error
  return "Cell ${cell} not accessible"
end try
if v is missing value then
  return ""
end if
return v as text
`,
        { setActiveSheet: true }
      );
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_set_cell',
    description: 'Set the value of a cell in the active Excel worksheet',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        cell: {
          type: 'string',
          description: 'Cell reference in A1 notation (e.g., "A1", "B2")'
        },
        value: {
          type: ['string', 'number'],
          description: 'Value to set in the cell'
        }
      },
      required: ['cell', 'value']
    },
    async handler(args) {
      const cell = validateExcelCellReference(args.cell, 'cell');
      if (args.value === undefined || args.value === null) {
        throw new Error('value is required');
      }
      const val = typeof args.value === 'number' ? args.value : toAppleScriptString(String(args.value));
      const script = wrapExcelScript(
        `
try
  set value of cell ${JSON.stringify(cell)} of ws to ${val}
on error errMsg
  return "Error setting cell ${cell}: " & errMsg
end try
return "Cell ${cell} set successfully"
`,
        { setActiveSheet: true }
      );
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_get_range',
    description: 'Get values from a range of cells in Excel as text',
    annotations: { readOnlyHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Range reference (e.g., "A1:B3", "A1:D10")'
        }
      },
      required: ['range']
    },
    async handler(args) {
      const range = validateExcelRangeReference(args.range, 'range');
      const script = wrapExcelScript(
        `
try
  set r to range ${JSON.stringify(range)} of ws
on error
  return "Invalid range: ${range}"
end try
set rc to count of rows of r
set cc to count of columns of r
set output to ""
repeat with i from 1 to rc
  set rowData to ""
  repeat with j from 1 to cc
    set v to value of cell j of row i of r
    if v is missing value then
      set cellText to ""
    else
      set cellText to v as text
    end if
    if j > 1 then set rowData to rowData & tab
    set rowData to rowData & cellText
  end repeat
  set output to output & rowData & linefeed
end repeat
return output
`,
        { setActiveSheet: true }
      );
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_set_cell_formula',
    description: 'Set a formula in an Excel cell',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        cell: {
          type: 'string',
          description: 'Cell reference in A1 notation'
        },
        formula: {
          type: 'string',
          description: 'Formula to set (e.g., "=SUM(A1:A10)", "=VLOOKUP(A1,B:C,2,FALSE)")'
        }
      },
      required: ['cell', 'formula']
    },
    async handler(args) {
      const cell = validateExcelCellReference(args.cell, 'cell');
      let formula = validateString(args.formula, 'formula', true);
      if (!formula.startsWith('=')) {
        formula = '=' + formula;
      }
      const script = wrapExcelScript(
        `
try
  set formula of cell ${JSON.stringify(cell)} of ws to ${JSON.stringify(formula)}
on error errMsg
  return "Error setting formula in ${cell}: " & errMsg
end try
return "Formula set in " & ${JSON.stringify(cell)} & ": " & ${JSON.stringify(formula)}
`,
        { setActiveSheet: true }
      );
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_clear_range',
    description: 'Clear contents of a range in the active Excel worksheet',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Range reference to clear (e.g., "A1:B3")'
        }
      },
      required: ['range']
    },
    async handler(args) {
      const range = validateExcelRangeReference(args.range, 'range');
      const script = wrapExcelScript(
        `
try
  clear contents range ${JSON.stringify(range)} of ws
on error
  return "Invalid range: ${range}"
end try
return "Range ${range} cleared"
`,
        { setActiveSheet: true }
      );
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_get_used_range',
    description: 'Get the used range address and dimensions of the active Excel sheet',
    annotations: { readOnlyHint: true },
    inputSchema: {
      type: 'object',
      properties: {}
    },
    async handler() {
      const script = wrapExcelScript(
        `
set ur to used range of ws
set addr to get address of ur
set rc to count of rows of ur
set cc to count of columns of ur
return "Used range: " & addr & linefeed & "Rows: " & rc & linefeed & "Columns: " & cc
`,
        { setActiveSheet: true }
      );
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_find_cell',
    description: 'Find a cell containing specific text in the active Excel worksheet',
    annotations: { readOnlyHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        searchText: {
          type: 'string',
          description: 'Text to search for'
        },
        range: {
          type: 'string',
          description: 'Range to search in (default: used range)'
        }
      },
      required: ['searchText']
    },
    async handler(args) {
      const searchText = validateString(args.searchText, 'searchText', true);
      const range = args.range !== undefined ? validateExcelRangeReference(args.range, 'range') : undefined;
      const rangeRef = range ? `range ${JSON.stringify(range)} of ws` : 'used range of ws';
      const script = wrapExcelScript(
        `
try
  set searchRange to ${rangeRef}
  set foundCell to find searchRange what ${quoteAppleScriptString(searchText)}
on error
  return "Not found"
end try
set addr to get address of foundCell
set v to value of foundCell
return "Found at " & addr & ": " & (v as text)
`,
        { setActiveSheet: true }
      );
      return await runAppleScript(script);
    }
  }
];
