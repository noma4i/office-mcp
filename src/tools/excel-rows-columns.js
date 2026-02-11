import { validateInteger, validateNumber } from '../lib/validators.js';
import { runAppleScript } from '../lib/applescript/executor.js';

function columnLetter(num) {
  let result = '';
  while (num > 0) {
    num--;
    result = String.fromCharCode(65 + (num % 26)) + result;
    num = Math.floor(num / 26);
  }
  return result;
}

export const excelRowColumnTools = [
  {
    name: 'excel_insert_rows',
    description: 'Insert rows at a specific position',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        row: {
          type: 'integer',
          description: 'Row number to insert before (1-based)'
        },
        count: {
          type: 'integer',
          description: 'Number of rows to insert (default: 1)',
          default: 1
        }
      },
      required: ['row']
    },
    async handler(args) {
      const row = validateInteger(args.row, 'row', 1);
      const count = validateInteger(args.count, 'count', 1) || 1;
      const endRow = row + count - 1;
      const script = `
        tell application "Microsoft Excel"
          if (count of workbooks) = 0 then
            return "No workbook is open"
          end if
          set ws to active sheet
          insert into range (range "${row}:${endRow}" of ws) shift shift down
          return "${count} row(s) inserted at row ${row}"
        end tell
      `;
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_delete_rows',
    description: 'Delete rows at a specific position',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        row: {
          type: 'integer',
          description: 'Row number to start deleting (1-based)'
        },
        count: {
          type: 'integer',
          description: 'Number of rows to delete (default: 1)',
          default: 1
        }
      },
      required: ['row']
    },
    async handler(args) {
      const row = validateInteger(args.row, 'row', 1);
      const count = validateInteger(args.count, 'count', 1) || 1;
      const endRow = row + count - 1;
      const script = `
        tell application "Microsoft Excel"
          if (count of workbooks) = 0 then
            return "No workbook is open"
          end if
          set ws to active sheet
          delete range (range "${row}:${endRow}" of ws) shift shift up
          return "${count} row(s) deleted starting at row ${row}"
        end tell
      `;
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_insert_columns',
    description: 'Insert columns at a specific position',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        column: {
          type: 'integer',
          description: 'Column number to insert before (1-based)'
        },
        count: {
          type: 'integer',
          description: 'Number of columns to insert (default: 1)',
          default: 1
        }
      },
      required: ['column']
    },
    async handler(args) {
      const column = validateInteger(args.column, 'column', 1);
      const count = validateInteger(args.count, 'count', 1) || 1;
      const startCol = columnLetter(column);
      const endCol = columnLetter(column + count - 1);
      const script = `
        tell application "Microsoft Excel"
          if (count of workbooks) = 0 then
            return "No workbook is open"
          end if
          set ws to active sheet
          insert into range (range "${startCol}:${endCol}" of ws) shift shift to right
          return "${count} column(s) inserted at column ${startCol}"
        end tell
      `;
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_delete_columns',
    description: 'Delete columns at a specific position',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        column: {
          type: 'integer',
          description: 'Column number to start deleting (1-based)'
        },
        count: {
          type: 'integer',
          description: 'Number of columns to delete (default: 1)',
          default: 1
        }
      },
      required: ['column']
    },
    async handler(args) {
      const column = validateInteger(args.column, 'column', 1);
      const count = validateInteger(args.count, 'count', 1) || 1;
      const startCol = columnLetter(column);
      const endCol = columnLetter(column + count - 1);
      const script = `
        tell application "Microsoft Excel"
          if (count of workbooks) = 0 then
            return "No workbook is open"
          end if
          set ws to active sheet
          delete range (range "${startCol}:${endCol}" of ws) shift shift to left
          return "${count} column(s) deleted starting at column ${startCol}"
        end tell
      `;
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_set_column_width',
    description: 'Set the width of a column',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        column: {
          type: 'integer',
          description: 'Column number (1-based)'
        },
        width: {
          type: 'number',
          description: 'Column width in character units'
        }
      },
      required: ['column', 'width']
    },
    async handler(args) {
      const column = validateInteger(args.column, 'column', 1);
      const width = validateNumber(args.width, 'width', 0, 255);
      const script = `
        tell application "Microsoft Excel"
          if (count of workbooks) = 0 then
            return "No workbook is open"
          end if
          set ws to active sheet
          set column width of column ${column} of ws to ${width}
          return "Column ${column} width set to ${width}"
        end tell
      `;
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_set_row_height',
    description: 'Set the height of a row',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        row: {
          type: 'integer',
          description: 'Row number (1-based)'
        },
        height: {
          type: 'number',
          description: 'Row height in points'
        }
      },
      required: ['row', 'height']
    },
    async handler(args) {
      const row = validateInteger(args.row, 'row', 1);
      const height = validateNumber(args.height, 'height', 0, 409);
      const script = `
        tell application "Microsoft Excel"
          if (count of workbooks) = 0 then
            return "No workbook is open"
          end if
          set ws to active sheet
          set row height of row ${row} of ws to ${height}
          return "Row ${row} height set to ${height}"
        end tell
      `;
      return await runAppleScript(script);
    }
  }
];
