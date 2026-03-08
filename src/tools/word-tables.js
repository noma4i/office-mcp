import { validateString, validateInteger } from '../lib/validators.js';
import { runAppleScript } from '../lib/applescript/executor.js';
import { toAppleScriptString, quoteAppleScriptString } from '../lib/applescript/helpers.js';
import { wrapWordScript } from '../lib/applescript/script-wrappers.js';

export const tableTools = [
  {
    name: 'word_list_tables',
    description: 'List all tables in the active Word document with their dimensions (rows x columns)',
    annotations: { readOnlyHint: true },
    inputSchema: { type: 'object', properties: {} },
    async handler() {
      const script = wrapWordScript(`
set activeDoc to active document
set tableCount to count of tables of activeDoc
if tableCount = 0 then
  return "No tables found in document"
end if
set tableInfo to ""
repeat with i from 1 to tableCount
  set t to table i of activeDoc
  set rowCount to count of rows of t
  set colCount to count of columns of t
  set tableInfo to tableInfo & "Table " & i & ": " & rowCount & " rows x " & colCount & " columns" & linefeed
end repeat
return tableInfo
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'word_get_table_cell',
    description: 'Get the text content of a specific table cell in Word',
    annotations: { readOnlyHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        tableIndex: { type: 'integer', description: 'Table index (1-based)' },
        row: { type: 'integer', description: 'Row number (1-based)' },
        column: { type: 'integer', description: 'Column number (1-based)' }
      },
      required: ['tableIndex', 'row', 'column']
    },
    async handler(args) {
      const tableIndex = validateInteger(args.tableIndex, 'tableIndex', 1);
      const row = validateInteger(args.row, 'row', 1);
      const column = validateInteger(args.column, 'column', 1);
      const script = wrapWordScript(`
set activeDoc to active document
set tableCount to count of tables of activeDoc
if ${tableIndex} > tableCount then
  return "Table index out of range. Document has " & tableCount & " tables."
end if
set t to table ${tableIndex} of activeDoc
try
  set c to cell ${column} of row ${row} of t
  set cellText to content of text object of c
on error
  return "Cell not found: row ${row}, column ${column} in table ${tableIndex}"
end try
if length of cellText > 0 then
  repeat while (length of cellText > 0) and ((ASCII number of (character -1 of cellText)) is in {7, 13})
    set cellText to text 1 thru -2 of cellText
  end repeat
end if
return cellText
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'word_set_table_cell',
    description: 'Set the text content of a specific table cell in Word',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        tableIndex: { type: 'integer', description: 'Table index (1-based)' },
        row: { type: 'integer', description: 'Row number (1-based)' },
        column: { type: 'integer', description: 'Column number (1-based)' },
        text: { type: 'string', description: 'Text to set in the cell' }
      },
      required: ['tableIndex', 'row', 'column', 'text']
    },
    async handler(args) {
      const tableIndex = validateInteger(args.tableIndex, 'tableIndex', 1);
      const row = validateInteger(args.row, 'row', 1);
      const column = validateInteger(args.column, 'column', 1);
      const text = validateString(args.text, 'text', true);
      const script = wrapWordScript(`
set activeDoc to active document
set tableCount to count of tables of activeDoc
if ${tableIndex} > tableCount then
  return "Table index out of range. Document has " & tableCount & " tables."
end if
set t to table ${tableIndex} of activeDoc
try
  set c to cell ${column} of row ${row} of t
  set content of text object of c to ${toAppleScriptString(text)}
on error
  return "Cell not found: row ${row}, column ${column} in table ${tableIndex}"
end try
return "Cell [" & ${row} & "," & ${column} & "] in table " & ${tableIndex} & " set successfully"
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'word_select_table_cell',
    description: 'Select a specific table cell and move cursor there in Word',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        tableIndex: { type: 'integer', description: 'Table index (1-based)' },
        row: { type: 'integer', description: 'Row number (1-based)' },
        column: { type: 'integer', description: 'Column number (1-based)' }
      },
      required: ['tableIndex', 'row', 'column']
    },
    async handler(args) {
      const tableIndex = validateInteger(args.tableIndex, 'tableIndex', 1);
      const row = validateInteger(args.row, 'row', 1);
      const column = validateInteger(args.column, 'column', 1);
      const script = wrapWordScript(`
set activeDoc to active document
set tableCount to count of tables of activeDoc
if ${tableIndex} > tableCount then
  return "Table index out of range. Document has " & tableCount & " tables."
end if
set t to table ${tableIndex} of activeDoc
try
  set c to cell ${column} of row ${row} of t
  select (text object of c)
on error
  return "Cell not found: row ${row}, column ${column} in table ${tableIndex}"
end try
return "Cursor moved to cell [" & ${row} & "," & ${column} & "] in table " & ${tableIndex}
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'word_find_table_header',
    description: 'Find a table column by header text in Word (searches in specified header row)',
    annotations: { readOnlyHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        tableIndex: { type: 'integer', description: 'Table index (1-based)' },
        headerText: { type: 'string', description: 'Text to search for in header' },
        headerRow: { type: 'integer', description: 'Row number to search in (default: 1)', default: 1 }
      },
      required: ['tableIndex', 'headerText']
    },
    async handler(args) {
      const tableIndex = validateInteger(args.tableIndex, 'tableIndex', 1);
      const headerText = validateString(args.headerText, 'headerText', true);
      const headerRow = validateInteger(args.headerRow, 'headerRow', 1) || 1;
      const script = wrapWordScript(`
set activeDoc to active document
set tableCount to count of tables of activeDoc
if ${tableIndex} > tableCount then
  return "Table index out of range. Document has " & tableCount & " tables."
end if
set t to table ${tableIndex} of activeDoc
set colCount to count of columns of t
set foundCol to 0
repeat with colIdx from 1 to colCount
  try
    set c to cell colIdx of row ${headerRow} of t
    set cellText to content of text object of c
  on error
    return "Header row ${headerRow} is out of range for table ${tableIndex}"
  end try
  if length of cellText > 0 then
    repeat while (length of cellText > 0) and ((ASCII number of (character -1 of cellText)) is in {7, 13})
      set cellText to text 1 thru -2 of cellText
    end repeat
  end if
  if cellText contains ${quoteAppleScriptString(headerText)} then
    set foundCol to colIdx
    exit repeat
  end if
end repeat
if foundCol = 0 then
  return "Header not found: " & ${quoteAppleScriptString(headerText)}
else
  return "Column " & foundCol & " contains header: " & ${quoteAppleScriptString(headerText)}
end if
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'word_create_table',
    description: 'Create a new table at the current cursor position in Word',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: { rows: { type: 'integer', description: 'Number of rows' }, columns: { type: 'integer', description: 'Number of columns' } },
      required: ['rows', 'columns']
    },
    async handler(args) {
      const rows = validateInteger(args.rows, 'rows', 1);
      const columns = validateInteger(args.columns, 'columns', 1);
      const script = wrapWordScript(`
make new table at text object of selection with properties {number of rows:${rows}, number of columns:${columns}}
return "Table created with ${rows} rows and ${columns} columns"
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'word_add_table_row',
    description: 'Add a new row to a table in Word',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        tableIndex: { type: 'integer', description: 'Table index (1-based)' },
        afterRow: { type: 'integer', description: 'Insert after this row (if not provided, adds at end)' }
      },
      required: ['tableIndex']
    },
    async handler(args) {
      const tableIndex = validateInteger(args.tableIndex, 'tableIndex', 1);
      const afterRow = validateInteger(args.afterRow, 'afterRow', 1);
      const targetRow = afterRow || 0;
      const script = wrapWordScript(`
set d to active document
try
  set t to table ${tableIndex} of d
on error
  return "Table ${tableIndex} not found"
end try
set targetRowNum to ${targetRow}
if targetRowNum = 0 then
  set targetRowNum to count of rows of t
end if
try
  select (text object of row targetRowNum of t)
  insert rows selection position below
on error
  return "Row " & targetRowNum & " not found in table ${tableIndex}"
end try
return "Row added after row " & targetRowNum & " in table ${tableIndex}"
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'word_delete_table_row',
    description: 'Delete a row from a table in Word',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        tableIndex: { type: 'integer', description: 'Table index (1-based)' },
        row: { type: 'integer', description: 'Row number to delete (1-based)' }
      },
      required: ['tableIndex', 'row']
    },
    async handler(args) {
      const tableIndex = validateInteger(args.tableIndex, 'tableIndex', 1);
      const row = validateInteger(args.row, 'row', 1);
      const script = wrapWordScript(`
set d to active document
try
  set t to table ${tableIndex} of d
  delete row ${row} of t
on error
  return "Row ${row} not found in table ${tableIndex}"
end try
return "Row ${row} deleted from table ${tableIndex}"
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'word_add_table_column',
    description: 'Add a new column to a table in Word',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        tableIndex: { type: 'integer', description: 'Table index (1-based)' },
        afterColumn: { type: 'integer', description: 'Insert after this column (if not provided, adds at end)' }
      },
      required: ['tableIndex']
    },
    async handler(args) {
      const tableIndex = validateInteger(args.tableIndex, 'tableIndex', 1);
      const afterColumn = validateInteger(args.afterColumn, 'afterColumn', 1);
      const targetCol = afterColumn || 0;
      const script = wrapWordScript(`
set d to active document
try
  set t to table ${tableIndex} of d
on error
  return "Table ${tableIndex} not found"
end try
set targetColNum to ${targetCol}
if targetColNum = 0 then
  set targetColNum to count of columns of t
end if
try
  select (text object of cell targetColNum of row 1 of t)
  insert columns selection position insert on the right
on error
  return "Column " & targetColNum & " not found in table ${tableIndex}"
end try
return "Column added after column " & targetColNum & " in table ${tableIndex}"
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'word_delete_table_column',
    description: 'Delete a column from a table in Word',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        tableIndex: { type: 'integer', description: 'Table index (1-based)' },
        column: { type: 'integer', description: 'Column number to delete (1-based)' }
      },
      required: ['tableIndex', 'column']
    },
    async handler(args) {
      const tableIndex = validateInteger(args.tableIndex, 'tableIndex', 1);
      const column = validateInteger(args.column, 'column', 1);
      const script = wrapWordScript(`
set d to active document
try
  set t to table ${tableIndex} of d
  delete column ${column} of t
on error
  return "Column ${column} not found in table ${tableIndex}"
end try
return "Column ${column} deleted from table ${tableIndex}"
`);
      return await runAppleScript(script);
    }
  }
];

