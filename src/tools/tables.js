import { validateString, validateInteger } from '../lib/validators.js';
import { runAppleScript } from '../lib/applescript/executor.js';

export const tableTools = [
  {
    name: "list_tables",
    description: "List all tables in the active document with their dimensions (rows x columns)",
    annotations: { readOnlyHint: true },
    inputSchema: {
      type: "object",
      properties: {},
    },
    async handler() {
      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
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
        end tell
      `;

      return await runAppleScript(script);
    }
  },

  {
    name: "get_table_cell",
    description: "Get the text content of a specific table cell",
    annotations: { readOnlyHint: true },
    inputSchema: {
      type: "object",
      properties: {
        tableIndex: {
          type: "integer",
          description: "Table index (1-based)",
        },
        row: {
          type: "integer",
          description: "Row number (1-based)",
        },
        column: {
          type: "integer",
          description: "Column number (1-based)",
        },
      },
      required: ["tableIndex", "row", "column"],
    },
    async handler(args) {
      const tableIndex = validateInteger(args.tableIndex, "tableIndex", 1);
      const row = validateInteger(args.row, "row", 1);
      const column = validateInteger(args.column, "column", 1);

      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set activeDoc to active document
          set tableCount to count of tables of activeDoc
          if ${tableIndex} > tableCount then
            return "Table index out of range. Document has " & tableCount & " tables."
          end if
          set t to table ${tableIndex} of activeDoc
          set c to cell ${column} of row ${row} of t
          set cellText to content of text object of c
          -- Remove trailing cell marker (ASCII 7 and 13)
          if length of cellText > 0 then
            repeat while (length of cellText > 0) and ((ASCII number of (character -1 of cellText)) is in {7, 13})
              set cellText to text 1 thru -2 of cellText
            end repeat
          end if
          return cellText
        end tell
      `;

      return await runAppleScript(script);
    }
  },

  {
    name: "set_table_cell",
    description: "Set the text content of a specific table cell",
    annotations: { destructiveHint: true },
    inputSchema: {
      type: "object",
      properties: {
        tableIndex: {
          type: "integer",
          description: "Table index (1-based)",
        },
        row: {
          type: "integer",
          description: "Row number (1-based)",
        },
        column: {
          type: "integer",
          description: "Column number (1-based)",
        },
        text: {
          type: "string",
          description: "Text to set in the cell",
        },
      },
      required: ["tableIndex", "row", "column", "text"],
    },
    async handler(args) {
      const tableIndex = validateInteger(args.tableIndex, "tableIndex", 1);
      const row = validateInteger(args.row, "row", 1);
      const column = validateInteger(args.column, "column", 1);
      const text = validateString(args.text, "text", true);

      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set activeDoc to active document
          set tableCount to count of tables of activeDoc
          if ${tableIndex} > tableCount then
            return "Table index out of range. Document has " & tableCount & " tables."
          end if
          set t to table ${tableIndex} of activeDoc
          set c to cell ${column} of row ${row} of t
          set content of text object of c to ${JSON.stringify(text)}
          return "Cell [" & ${row} & "," & ${column} & "] in table " & ${tableIndex} & " set successfully"
        end tell
      `;

      return await runAppleScript(script);
    }
  },

  {
    name: "select_table_cell",
    description: "Select a specific table cell and move cursor there",
    annotations: { destructiveHint: true },
    inputSchema: {
      type: "object",
      properties: {
        tableIndex: {
          type: "integer",
          description: "Table index (1-based)",
        },
        row: {
          type: "integer",
          description: "Row number (1-based)",
        },
        column: {
          type: "integer",
          description: "Column number (1-based)",
        },
      },
      required: ["tableIndex", "row", "column"],
    },
    async handler(args) {
      const tableIndex = validateInteger(args.tableIndex, "tableIndex", 1);
      const row = validateInteger(args.row, "row", 1);
      const column = validateInteger(args.column, "column", 1);

      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set activeDoc to active document
          set tableCount to count of tables of activeDoc
          if ${tableIndex} > tableCount then
            return "Table index out of range. Document has " & tableCount & " tables."
          end if
          set t to table ${tableIndex} of activeDoc
          set c to cell ${column} of row ${row} of t
          select (text object of c)
          return "Cursor moved to cell [" & ${row} & "," & ${column} & "] in table " & ${tableIndex}
        end tell
      `;

      return await runAppleScript(script);
    }
  },

  {
    name: "find_table_header",
    description: "Find a table column by header text (searches in specified header row)",
    annotations: { readOnlyHint: true },
    inputSchema: {
      type: "object",
      properties: {
        tableIndex: {
          type: "integer",
          description: "Table index (1-based)",
        },
        headerText: {
          type: "string",
          description: "Text to search for in header",
        },
        headerRow: {
          type: "integer",
          description: "Row number to search in (default: 1)",
          default: 1,
        },
      },
      required: ["tableIndex", "headerText"],
    },
    async handler(args) {
      const tableIndex = validateInteger(args.tableIndex, "tableIndex", 1);
      const headerText = validateString(args.headerText, "headerText", true);
      const headerRow = validateInteger(args.headerRow, "headerRow", 1) || 1;

      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set activeDoc to active document
          set tableCount to count of tables of activeDoc
          if ${tableIndex} > tableCount then
            return "Table index out of range. Document has " & tableCount & " tables."
          end if
          set t to table ${tableIndex} of activeDoc
          set colCount to count of columns of t
          set foundCol to 0
          repeat with colIdx from 1 to colCount
            set c to cell colIdx of row ${headerRow} of t
            set cellText to content of text object of c
            -- Remove trailing cell marker
            if length of cellText > 0 then
              repeat while (length of cellText > 0) and ((ASCII number of (character -1 of cellText)) is in {7, 13})
                set cellText to text 1 thru -2 of cellText
              end repeat
            end if
            if cellText contains ${JSON.stringify(headerText)} then
              set foundCol to colIdx
              exit repeat
            end if
          end repeat
          if foundCol = 0 then
            return "Header not found: " & ${JSON.stringify(headerText)}
          else
            return "Column " & foundCol & " contains header: " & ${JSON.stringify(headerText)}
          end if
        end tell
      `;

      return await runAppleScript(script);
    }
  },

  {
    name: "create_table",
    description: "Create a new table at the current cursor position",
    annotations: { destructiveHint: true },
    inputSchema: {
      type: "object",
      properties: {
        rows: {
          type: "integer",
          description: "Number of rows",
        },
        columns: {
          type: "integer",
          description: "Number of columns",
        },
      },
      required: ["rows", "columns"],
    },
    async handler(args) {
      const rows = validateInteger(args.rows, "rows", 1);
      const columns = validateInteger(args.columns, "columns", 1);

      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          make new table at selection with properties {number of rows:${rows}, number of columns:${columns}}
          return "Table created with ${rows} rows and ${columns} columns"
        end tell
      `;

      return await runAppleScript(script);
    }
  },

  {
    name: "add_table_row",
    description: "Add a new row to a table",
    annotations: { destructiveHint: true },
    inputSchema: {
      type: "object",
      properties: {
        tableIndex: {
          type: "integer",
          description: "Table index (1-based)",
        },
        afterRow: {
          type: "integer",
          description: "Insert after this row (if not provided, adds at end)",
        },
      },
      required: ["tableIndex"],
    },
    async handler(args) {
      const tableIndex = validateInteger(args.tableIndex, "tableIndex", 1);
      const afterRow = validateInteger(args.afterRow, "afterRow", 1);

      const script = afterRow
        ? `
          tell application "Microsoft Word"
            if (count of documents) = 0 then
              return "No document is open"
            end if
            set d to active document
            set t to table ${tableIndex} of d
            insert rows below row ${afterRow} of t
            return "Row added after row ${afterRow} in table ${tableIndex}"
          end tell
        `
        : `
          tell application "Microsoft Word"
            if (count of documents) = 0 then
              return "No document is open"
            end if
            set d to active document
            set t to table ${tableIndex} of d
            set rowCount to count of rows of t
            insert rows below row rowCount of t
            return "Row added at end of table ${tableIndex}"
          end tell
        `;

      return await runAppleScript(script);
    }
  },

  {
    name: "delete_table_row",
    description: "Delete a row from a table",
    annotations: { destructiveHint: true },
    inputSchema: {
      type: "object",
      properties: {
        tableIndex: {
          type: "integer",
          description: "Table index (1-based)",
        },
        row: {
          type: "integer",
          description: "Row number to delete (1-based)",
        },
      },
      required: ["tableIndex", "row"],
    },
    async handler(args) {
      const tableIndex = validateInteger(args.tableIndex, "tableIndex", 1);
      const row = validateInteger(args.row, "row", 1);

      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set d to active document
          set t to table ${tableIndex} of d
          delete row ${row} of t
          return "Row ${row} deleted from table ${tableIndex}"
        end tell
      `;

      return await runAppleScript(script);
    }
  },

  {
    name: "add_table_column",
    description: "Add a new column to a table",
    annotations: { destructiveHint: true },
    inputSchema: {
      type: "object",
      properties: {
        tableIndex: {
          type: "integer",
          description: "Table index (1-based)",
        },
        afterColumn: {
          type: "integer",
          description: "Insert after this column (if not provided, adds at end)",
        },
      },
      required: ["tableIndex"],
    },
    async handler(args) {
      const tableIndex = validateInteger(args.tableIndex, "tableIndex", 1);
      const afterColumn = validateInteger(args.afterColumn, "afterColumn", 1);

      const script = afterColumn
        ? `
          tell application "Microsoft Word"
            if (count of documents) = 0 then
              return "No document is open"
            end if
            set d to active document
            set t to table ${tableIndex} of d
            insert columns after column ${afterColumn} of t
            return "Column added after column ${afterColumn} in table ${tableIndex}"
          end tell
        `
        : `
          tell application "Microsoft Word"
            if (count of documents) = 0 then
              return "No document is open"
            end if
            set d to active document
            set t to table ${tableIndex} of d
            set colCount to count of columns of t
            insert columns after column colCount of t
            return "Column added at end of table ${tableIndex}"
          end tell
        `;

      return await runAppleScript(script);
    }
  },

  {
    name: "delete_table_column",
    description: "Delete a column from a table",
    annotations: { destructiveHint: true },
    inputSchema: {
      type: "object",
      properties: {
        tableIndex: {
          type: "integer",
          description: "Table index (1-based)",
        },
        column: {
          type: "integer",
          description: "Column number to delete (1-based)",
        },
      },
      required: ["tableIndex", "column"],
    },
    async handler(args) {
      const tableIndex = validateInteger(args.tableIndex, "tableIndex", 1);
      const column = validateInteger(args.column, "column", 1);

      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set d to active document
          set t to table ${tableIndex} of d
          delete column ${column} of t
          return "Column ${column} deleted from table ${tableIndex}"
        end tell
      `;

      return await runAppleScript(script);
    }
  }
];
