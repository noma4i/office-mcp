import { validateInteger, validateNumber } from '../lib/validators.js';
import { runAppleScript } from '../lib/applescript/executor.js';
import { wrapExcelScript } from '../lib/applescript/script-wrappers.js';

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
    description: 'Insert rows at a specific position in the active Excel worksheet',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        row: { type: 'integer', description: 'Row number to insert before (1-based)' },
        count: { type: 'integer', description: 'Number of rows to insert (default: 1)', default: 1 }
      },
      required: ['row']
    },
    async handler(args) {
      const row = validateInteger(args.row, 'row', 1);
      const count = validateInteger(args.count, 'count', 1) || 1;
      const endRow = row + count - 1;
      const script = wrapExcelScript(`
set ws to active sheet
try
  insert into range (range "${row}:${endRow}" of ws) shift shift down
on error errMsg
  return "Error inserting rows: " & errMsg
end try
return "${count} row(s) inserted at row ${row}"
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'excel_delete_rows',
    description: 'Delete rows at a specific position in the active Excel worksheet',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        row: { type: 'integer', description: 'Row number to start deleting (1-based)' },
        count: { type: 'integer', description: 'Number of rows to delete (default: 1)', default: 1 }
      },
      required: ['row']
    },
    async handler(args) {
      const row = validateInteger(args.row, 'row', 1);
      const count = validateInteger(args.count, 'count', 1) || 1;
      const endRow = row + count - 1;
      const script = wrapExcelScript(`
set ws to active sheet
try
  delete range (range "${row}:${endRow}" of ws) shift shift up
on error errMsg
  return "Error deleting rows: " & errMsg
end try
return "${count} row(s) deleted starting at row ${row}"
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'excel_insert_columns',
    description: 'Insert columns at a specific position in the active Excel worksheet',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        column: { type: 'integer', description: 'Column number to insert before (1-based)' },
        count: { type: 'integer', description: 'Number of columns to insert (default: 1)', default: 1 }
      },
      required: ['column']
    },
    async handler(args) {
      const column = validateInteger(args.column, 'column', 1);
      const count = validateInteger(args.count, 'count', 1) || 1;
      const startCol = columnLetter(column);
      const endCol = columnLetter(column + count - 1);
      const script = wrapExcelScript(`
set ws to active sheet
try
  insert into range (range "${startCol}:${endCol}" of ws) shift shift to right
on error errMsg
  return "Error inserting columns: " & errMsg
end try
return "${count} column(s) inserted at column ${startCol}"
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'excel_delete_columns',
    description: 'Delete columns at a specific position in the active Excel worksheet',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        column: { type: 'integer', description: 'Column number to start deleting (1-based)' },
        count: { type: 'integer', description: 'Number of columns to delete (default: 1)', default: 1 }
      },
      required: ['column']
    },
    async handler(args) {
      const column = validateInteger(args.column, 'column', 1);
      const count = validateInteger(args.count, 'count', 1) || 1;
      const startCol = columnLetter(column);
      const endCol = columnLetter(column + count - 1);
      const script = wrapExcelScript(`
set ws to active sheet
try
  delete range (range "${startCol}:${endCol}" of ws) shift shift to left
on error errMsg
  return "Error deleting columns: " & errMsg
end try
return "${count} column(s) deleted starting at column ${startCol}"
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'excel_set_column_width',
    description: 'Set the width of a column in the active Excel worksheet',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        column: { type: 'integer', description: 'Column number (1-based)' },
        width: { type: 'number', description: 'Column width in character units' }
      },
      required: ['column', 'width']
    },
    async handler(args) {
      const column = validateInteger(args.column, 'column', 1);
      const width = validateNumber(args.width, 'width', 0, 255);
      const script = wrapExcelScript(`
set ws to active sheet
try
  set column width of column ${column} of ws to ${width}
on error errMsg
  return "Error setting column width: " & errMsg
end try
return "Column ${column} width set to ${width}"
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'excel_set_row_height',
    description: 'Set the height of a row in the active Excel worksheet',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        row: { type: 'integer', description: 'Row number (1-based)' },
        height: { type: 'number', description: 'Row height in points' }
      },
      required: ['row', 'height']
    },
    async handler(args) {
      const row = validateInteger(args.row, 'row', 1);
      const height = validateNumber(args.height, 'height', 0, 409);
      const script = wrapExcelScript(`
set ws to active sheet
try
  set row height of row ${row} of ws to ${height}
on error errMsg
  return "Error setting row height: " & errMsg
end try
return "Row ${row} height set to ${height}"
`);
      return await runAppleScript(script);
    }
  }
];

