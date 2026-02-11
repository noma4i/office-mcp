import { validateString } from '../lib/validators.js';
import { runAppleScript } from '../lib/applescript/executor.js';

export const excelCellTools = [
  {
    name: 'excel_get_cell',
    description: 'Get the value of a cell (A1 notation)',
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
      const cell = validateString(args.cell, 'cell', true);
      const script = `
        tell application "Microsoft Excel"
          if (count of workbooks) = 0 then
            return "No workbook is open"
          end if
          set ws to active sheet
          set v to value of cell ${JSON.stringify(cell)} of ws
          if v is missing value then
            return ""
          end if
          return v as text
        end tell
      `;
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_set_cell',
    description: 'Set the value of a cell',
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
      const cell = validateString(args.cell, 'cell', true);
      if (args.value === undefined || args.value === null) {
        throw new Error('value is required');
      }
      const val = typeof args.value === 'number' ? args.value : JSON.stringify(String(args.value));
      const script = `
        tell application "Microsoft Excel"
          if (count of workbooks) = 0 then
            return "No workbook is open"
          end if
          set ws to active sheet
          set value of cell ${JSON.stringify(cell)} of ws to ${val}
          return "Cell ${cell} set successfully"
        end tell
      `;
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_get_range',
    description: 'Get values from a range of cells as text',
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
      const range = validateString(args.range, 'range', true);
      const script = `
        tell application "Microsoft Excel"
          if (count of workbooks) = 0 then
            return "No workbook is open"
          end if
          set ws to active sheet
          set r to range ${JSON.stringify(range)} of ws
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
        end tell
      `;
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_set_cell_formula',
    description: 'Set a formula in a cell',
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
      const cell = validateString(args.cell, 'cell', true);
      const formula = validateString(args.formula, 'formula', true);
      const script = `
        tell application "Microsoft Excel"
          if (count of workbooks) = 0 then
            return "No workbook is open"
          end if
          set ws to active sheet
          set formula of cell ${JSON.stringify(cell)} of ws to ${JSON.stringify(formula)}
          return "Formula set in ${cell}: ${formula}"
        end tell
      `;
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_clear_range',
    description: 'Clear contents of a range',
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
      const range = validateString(args.range, 'range', true);
      const script = `
        tell application "Microsoft Excel"
          if (count of workbooks) = 0 then
            return "No workbook is open"
          end if
          set ws to active sheet
          clear contents range ${JSON.stringify(range)} of ws
          return "Range ${range} cleared"
        end tell
      `;
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_get_used_range',
    description: 'Get the used range address and dimensions of the active sheet',
    annotations: { readOnlyHint: true },
    inputSchema: {
      type: 'object',
      properties: {}
    },
    async handler() {
      const script = `
        tell application "Microsoft Excel"
          if (count of workbooks) = 0 then
            return "No workbook is open"
          end if
          set ws to active sheet
          set ur to used range of ws
          set addr to get address of ur
          set rc to count of rows of ur
          set cc to count of columns of ur
          return "Used range: " & addr & linefeed & "Rows: " & rc & linefeed & "Columns: " & cc
        end tell
      `;
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_find_cell',
    description: 'Find a cell containing specific text',
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
      const range = validateString(args.range, 'range', false);
      const rangeRef = range ? `range ${JSON.stringify(range)} of ws` : 'used range of ws';
      const script = `
        tell application "Microsoft Excel"
          if (count of workbooks) = 0 then
            return "No workbook is open"
          end if
          set ws to active sheet
          set searchRange to ${rangeRef}
          set foundCell to find searchRange what ${JSON.stringify(searchText)}
          if foundCell is missing value then
            return "Not found: ${searchText}"
          end if
          set addr to get address of foundCell
          set v to value of foundCell
          return "Found at " & addr & ": " & (v as text)
        end tell
      `;
      return await runAppleScript(script);
    }
  }
];
