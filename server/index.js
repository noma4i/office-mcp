#!/usr/bin/env node

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { execFile } from "child_process";
import { promisify } from "util";

const execFileAsync = promisify(execFile);

// === Input Validation Helpers (Security Fix v0.4.0) ===
// Defense-in-depth against AppleScript injection

function validateString(value, name, required = true) {
  if (value === undefined || value === null) {
    if (required) {
      throw new Error(`${name} is required`);
    }
    return "";
  }
  if (typeof value !== "string") {
    throw new Error(`${name} must be a string`);
  }
  if (required && value.trim() === "") {
    throw new Error(`${name} cannot be empty`);
  }
  return value;
}

function validateBoolean(value, name, defaultValue = false) {
  if (value === undefined || value === null) {
    return defaultValue;
  }
  if (typeof value !== "boolean") {
    throw new Error(`${name} must be a boolean`);
  }
  return value;
}

function validateNumber(value, name, min = 0, max = Number.MAX_SAFE_INTEGER) {
  if (value === undefined || value === null) {
    return undefined;
  }
  const num = Number(value);
  if (!Number.isFinite(num)) {
    throw new Error(`${name} must be a valid number`);
  }
  if (num < min || num > max) {
    throw new Error(`${name} must be between ${min} and ${max}`);
  }
  return num;
}

function validateInteger(value, name, min = 0, max = Number.MAX_SAFE_INTEGER) {
  if (value === undefined || value === null) {
    return undefined;
  }
  const num = parseInt(value, 10);
  if (!Number.isInteger(num)) {
    throw new Error(`${name} must be an integer`);
  }
  if (num < min || num > max) {
    throw new Error(`${name} must be between ${min} and ${max}`);
  }
  return num;
}

function getErrorMessage(error) {
  if (error instanceof Error) return error.message;
  if (typeof error === "string") return error;
  return String(error);
}

async function runAppleScript(script) {
  try {
    const { stdout } = await execFileAsync("osascript", ["-e", script]);
    return stdout.trim();
  } catch (error) {
    throw new Error(`AppleScript error: ${getErrorMessage(error)}`);
  }
}

const server = new Server(
  {
    name: "Microsoft-Word-Server",
    version: "0.6.0",
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

server.setRequestHandler(ListToolsRequestSchema, async (request) => {
  return {
    tools: [
      {
        name: "create_document",
        description: "Create a new Word document with optional content",
        annotations: {
          destructiveHint: true,
        },
        inputSchema: {
          type: "object",
          properties: {
            content: {
              type: "string",
              description: "Optional initial content for the document",
            },
          },
        },
      },
      {
        name: "open_document",
        description: "Open an existing Word document",
        annotations: {
          readOnlyHint: true,
        },
        inputSchema: {
          type: "object",
          properties: {
            path: {
              type: "string",
              description: "Full path to the document to open",
            },
          },
          required: ["path"],
        },
      },
      {
        name: "get_document_text",
        description: "Get all text content from the active document",
        annotations: {
          readOnlyHint: true,
        },
        inputSchema: {
          type: "object",
          properties: {},
        },
      },
      {
        name: "insert_text",
        description: "Insert text at the current cursor position",
        annotations: {
          destructiveHint: true,
        },
        inputSchema: {
          type: "object",
          properties: {
            text: {
              type: "string",
              description: "Text to insert",
            },
          },
          required: ["text"],
        },
      },
      {
        name: "replace_text",
        description: "Find and replace text in the active document",
        annotations: {
          destructiveHint: true,
        },
        inputSchema: {
          type: "object",
          properties: {
            find: {
              type: "string",
              description: "Text to find",
            },
            replace: {
              type: "string",
              description: "Text to replace with",
            },
            all: {
              type: "boolean",
              description: "Replace all occurrences (default: true)",
              default: true,
            },
          },
          required: ["find", "replace"],
        },
      },
      {
        name: "format_text",
        description: "Apply formatting to selected text (bold, italic, underline, font, size, color)",
        annotations: {
          destructiveHint: true,
        },
        inputSchema: {
          type: "object",
          properties: {
            bold: {
              type: "boolean",
              description: "Make text bold",
            },
            italic: {
              type: "boolean",
              description: "Make text italic",
            },
            underline: {
              type: "boolean",
              description: "Underline text",
            },
            font: {
              type: "string",
              description: "Font name (e.g., 'Arial', 'Times New Roman')",
            },
            size: {
              type: "number",
              description: "Font size in points",
            },
          },
        },
      },
      {
        name: "save_document",
        description: "Save the active document",
        annotations: {
          destructiveHint: true,
        },
        inputSchema: {
          type: "object",
          properties: {
            path: {
              type: "string",
              description: "Optional path to save as (if not provided, saves to current location)",
            },
          },
        },
      },
      {
        name: "close_document",
        description: "Close the active document",
        annotations: {
          destructiveHint: true,
        },
        inputSchema: {
          type: "object",
          properties: {
            save: {
              type: "boolean",
              description: "Save before closing (default: true)",
              default: true,
            },
          },
        },
      },
      {
        name: "export_pdf",
        description: "Export the active document as PDF",
        annotations: {
          destructiveHint: true,
        },
        inputSchema: {
          type: "object",
          properties: {
            path: {
              type: "string",
              description: "Full path for the PDF file",
            },
          },
          required: ["path"],
        },
      },
      {
        name: "list_tables",
        description: "List all tables in the active document with their dimensions (rows x columns)",
        annotations: {
          readOnlyHint: true,
        },
        inputSchema: {
          type: "object",
          properties: {},
        },
      },
      {
        name: "get_table_cell",
        description: "Get the text content of a specific table cell",
        annotations: {
          readOnlyHint: true,
        },
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
      },
      {
        name: "set_table_cell",
        description: "Set the text content of a specific table cell",
        annotations: {
          destructiveHint: true,
        },
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
      },
      {
        name: "select_table_cell",
        description: "Move cursor to a specific table cell for subsequent operations",
        annotations: {
          destructiveHint: true,
        },
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
      },
      {
        name: "find_table_header",
        description: "Find the column index by header text in a table",
        annotations: {
          readOnlyHint: true,
        },
        inputSchema: {
          type: "object",
          properties: {
            tableIndex: {
              type: "integer",
              description: "Table index (1-based)",
            },
            headerText: {
              type: "string",
              description: "Text to search for in the header row",
            },
            headerRow: {
              type: "integer",
              description: "Row number to search for headers (default: 1)",
              default: 1,
            },
          },
          required: ["tableIndex", "headerText"],
        },
      },
      {
        name: "move_cursor_after_text",
        description: "Find text in document and move cursor right after it",
        annotations: {
          destructiveHint: true,
        },
        inputSchema: {
          type: "object",
          properties: {
            searchText: {
              type: "string",
              description: "Text to search for",
            },
            occurrence: {
              type: "integer",
              description: "Which occurrence to find (default: 1)",
              default: 1,
            },
          },
          required: ["searchText"],
        },
      },
      // === New tools v0.6.0 ===
      {
        name: "get_document_info",
        description: "Get document statistics (words, characters, paragraphs, pages)",
        annotations: {
          readOnlyHint: true,
        },
        inputSchema: {
          type: "object",
          properties: {},
        },
      },
      {
        name: "goto_start",
        description: "Move cursor to the beginning of the document",
        annotations: {
          destructiveHint: true,
        },
        inputSchema: {
          type: "object",
          properties: {},
        },
      },
      {
        name: "goto_end",
        description: "Move cursor to the end of the document",
        annotations: {
          destructiveHint: true,
        },
        inputSchema: {
          type: "object",
          properties: {},
        },
      },
      {
        name: "get_selection_info",
        description: "Get position and length of current selection",
        annotations: {
          readOnlyHint: true,
        },
        inputSchema: {
          type: "object",
          properties: {},
        },
      },
      {
        name: "select_all",
        description: "Select all content in the document",
        annotations: {
          destructiveHint: true,
        },
        inputSchema: {
          type: "object",
          properties: {},
        },
      },
      {
        name: "create_table",
        description: "Create a table at the current cursor position",
        annotations: {
          destructiveHint: true,
        },
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
      },
      {
        name: "add_table_row",
        description: "Add a row to a table",
        annotations: {
          destructiveHint: true,
        },
        inputSchema: {
          type: "object",
          properties: {
            tableIndex: {
              type: "integer",
              description: "Table index (1-based)",
            },
            afterRow: {
              type: "integer",
              description: "Insert after this row (1-based). If omitted, adds at end",
            },
          },
          required: ["tableIndex"],
        },
      },
      {
        name: "delete_table_row",
        description: "Delete a row from a table",
        annotations: {
          destructiveHint: true,
        },
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
      },
      {
        name: "add_table_column",
        description: "Add a column to a table",
        annotations: {
          destructiveHint: true,
        },
        inputSchema: {
          type: "object",
          properties: {
            tableIndex: {
              type: "integer",
              description: "Table index (1-based)",
            },
            afterColumn: {
              type: "integer",
              description: "Insert after this column (1-based). If omitted, adds at end",
            },
          },
          required: ["tableIndex"],
        },
      },
      {
        name: "delete_table_column",
        description: "Delete a column from a table",
        annotations: {
          destructiveHint: true,
        },
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
      },
      {
        name: "list_bookmarks",
        description: "List all bookmarks in the document",
        annotations: {
          readOnlyHint: true,
        },
        inputSchema: {
          type: "object",
          properties: {},
        },
      },
      {
        name: "create_bookmark",
        description: "Create a bookmark at the current selection",
        annotations: {
          destructiveHint: true,
        },
        inputSchema: {
          type: "object",
          properties: {
            name: {
              type: "string",
              description: "Bookmark name (no spaces allowed)",
            },
          },
          required: ["name"],
        },
      },
      {
        name: "goto_bookmark",
        description: "Go to a bookmark (select its content)",
        annotations: {
          destructiveHint: true,
        },
        inputSchema: {
          type: "object",
          properties: {
            name: {
              type: "string",
              description: "Bookmark name",
            },
          },
          required: ["name"],
        },
      },
      {
        name: "delete_bookmark",
        description: "Delete a bookmark",
        annotations: {
          destructiveHint: true,
        },
        inputSchema: {
          type: "object",
          properties: {
            name: {
              type: "string",
              description: "Bookmark name",
            },
          },
          required: ["name"],
        },
      },
      {
        name: "list_hyperlinks",
        description: "List all hyperlinks in the document",
        annotations: {
          readOnlyHint: true,
        },
        inputSchema: {
          type: "object",
          properties: {},
        },
      },
      {
        name: "create_hyperlink",
        description: "Create a hyperlink on the current selection",
        annotations: {
          destructiveHint: true,
        },
        inputSchema: {
          type: "object",
          properties: {
            url: {
              type: "string",
              description: "URL for the hyperlink",
            },
            displayText: {
              type: "string",
              description: "Optional display text (uses selection if omitted)",
            },
          },
          required: ["url"],
        },
      },
      {
        name: "list_paragraphs",
        description: "List paragraphs with their styles (first N)",
        annotations: {
          readOnlyHint: true,
        },
        inputSchema: {
          type: "object",
          properties: {
            limit: {
              type: "integer",
              description: "Maximum number of paragraphs to list (default: 50)",
              default: 50,
            },
          },
        },
      },
      {
        name: "goto_paragraph",
        description: "Go to a paragraph by index (select it)",
        annotations: {
          destructiveHint: true,
        },
        inputSchema: {
          type: "object",
          properties: {
            index: {
              type: "integer",
              description: "Paragraph index (1-based)",
            },
          },
          required: ["index"],
        },
      },
      {
        name: "set_paragraph_style",
        description: "Set the style of a paragraph",
        annotations: {
          destructiveHint: true,
        },
        inputSchema: {
          type: "object",
          properties: {
            index: {
              type: "integer",
              description: "Paragraph index (1-based)",
            },
            styleName: {
              type: "string",
              description: "Style name (e.g., 'Heading 1', 'Normal', 'Title')",
            },
          },
          required: ["index", "styleName"],
        },
      },
    ],
  };
});

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  switch (request.params.name) {
    case "create_document": {
      const content = validateString(request.params.arguments?.content, "content", false);

      // Build script conditionally to avoid any user input in AppleScript conditionals
      const script = content
        ? `
        tell application "Microsoft Word"
          activate
          set newDoc to make new document
          tell newDoc
            set content of text object to ${JSON.stringify(content)}
          end tell
          return "New document created successfully"
        end tell
      `
        : `
        tell application "Microsoft Word"
          activate
          set newDoc to make new document
          return "New document created successfully"
        end tell
      `;

      try {
        const result = await runAppleScript(script);
        return {
          content: [
            {
              type: "text",
              text: result,
            },
          ],
        };
      } catch (error) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to create document: ${getErrorMessage(error)}`,
            },
          ],
          isError: true,
        };
      }
    }

    case "open_document": {
      const path = validateString(request.params.arguments?.path, "path", true);

      const script = `
        tell application "Microsoft Word"
          activate
          open ${JSON.stringify(path)}
          return "Document opened successfully"
        end tell
      `;

      try {
        const result = await runAppleScript(script);
        return {
          content: [
            {
              type: "text",
              text: result,
            },
          ],
        };
      } catch (error) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to open document: ${getErrorMessage(error)}`,
            },
          ],
          isError: true,
        };
      }
    }

    case "get_document_text": {
      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set activeDoc to active document
          return content of text object of activeDoc as string
        end tell
      `;

      try {
        const result = await runAppleScript(script);
        return {
          content: [
            {
              type: "text",
              text: result,
            },
          ],
        };
      } catch (error) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to get document text: ${getErrorMessage(error)}`,
            },
          ],
          isError: true,
        };
      }
    }

    case "insert_text": {
      const text = validateString(request.params.arguments?.text, "text", true);

      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          type text selection text ${JSON.stringify(text)}
          return "Text inserted successfully"
        end tell
      `;

      try {
        const result = await runAppleScript(script);
        return {
          content: [
            {
              type: "text",
              text: result,
            },
          ],
        };
      } catch (error) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to insert text: ${getErrorMessage(error)}`,
            },
          ],
          isError: true,
        };
      }
    }

    case "replace_text": {
      const find = validateString(request.params.arguments?.find, "find", true);
      const replace = validateString(request.params.arguments?.replace, "replace", true);
      const all = validateBoolean(request.params.arguments?.all, "all", true);

      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set activeDoc to active document
          tell activeDoc
            set findObject to find object of selection
            clear formatting findObject
            set content of findObject to ${JSON.stringify(find)}
            set replacement to ${JSON.stringify(replace)}
            ${all ? 'execute find findObject replace replace all' : 'execute find findObject replace replace one'}
          end tell
          return "Text replaced successfully"
        end tell
      `;

      try {
        const result = await runAppleScript(script);
        return {
          content: [
            {
              type: "text",
              text: result,
            },
          ],
        };
      } catch (error) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to replace text: ${getErrorMessage(error)}`,
            },
          ],
          isError: true,
        };
      }
    }

    case "format_text": {
      // Validate all inputs before building AppleScript
      const bold = request.params.arguments?.bold !== undefined
        ? validateBoolean(request.params.arguments?.bold, "bold")
        : undefined;
      const italic = request.params.arguments?.italic !== undefined
        ? validateBoolean(request.params.arguments?.italic, "italic")
        : undefined;
      const underline = request.params.arguments?.underline !== undefined
        ? validateBoolean(request.params.arguments?.underline, "underline")
        : undefined;
      const font = request.params.arguments?.font
        ? validateString(request.params.arguments?.font, "font", false)
        : undefined;
      const size = request.params.arguments?.size !== undefined
        ? validateNumber(request.params.arguments?.size, "size", 1, 1000)
        : undefined;

      let formatCommands = [];
      if (bold !== undefined) {
        formatCommands.push(`set bold of font object of selection to ${bold}`);
      }
      if (italic !== undefined) {
        formatCommands.push(`set italic of font object of selection to ${italic}`);
      }
      if (underline !== undefined) {
        formatCommands.push(`set underline of font object of selection to ${underline ? 'underline single' : 'underline none'}`);
      }
      if (font) {
        formatCommands.push(`set name of font object of selection to ${JSON.stringify(font)}`);
      }
      if (size !== undefined) {
        formatCommands.push(`set font size of font object of selection to ${size}`);
      }

      if (formatCommands.length === 0) {
        throw new Error("At least one formatting option is required");
      }

      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          ${formatCommands.join('\n          ')}
          return "Formatting applied successfully"
        end tell
      `;

      try {
        const result = await runAppleScript(script);
        return {
          content: [
            {
              type: "text",
              text: result,
            },
          ],
        };
      } catch (error) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to format text: ${getErrorMessage(error)}`,
            },
          ],
          isError: true,
        };
      }
    }

    case "save_document": {
      const path = request.params.arguments?.path
        ? validateString(request.params.arguments?.path, "path", false)
        : undefined;

      let script;
      if (path) {
        script = `
          tell application "Microsoft Word"
            if (count of documents) = 0 then
              return "No document is open"
            end if
            set activeDoc to active document
            save as activeDoc file name ${JSON.stringify(path)}
            return "Document saved as " & ${JSON.stringify(path)}
          end tell
        `;
      } else {
        script = `
          tell application "Microsoft Word"
            if (count of documents) = 0 then
              return "No document is open"
            end if
            set activeDoc to active document
            save activeDoc
            return "Document saved successfully"
          end tell
        `;
      }

      try {
        const result = await runAppleScript(script);
        return {
          content: [
            {
              type: "text",
              text: result,
            },
          ],
        };
      } catch (error) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to save document: ${getErrorMessage(error)}`,
            },
          ],
          isError: true,
        };
      }
    }

    case "close_document": {
      const save = validateBoolean(request.params.arguments?.save, "save", true);

      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set activeDoc to active document
          close activeDoc ${save ? 'saving yes' : 'saving no'}
          return "Document closed successfully"
        end tell
      `;

      try {
        const result = await runAppleScript(script);
        return {
          content: [
            {
              type: "text",
              text: result,
            },
          ],
        };
      } catch (error) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to close document: ${getErrorMessage(error)}`,
            },
          ],
          isError: true,
        };
      }
    }

    case "export_pdf": {
      const path = validateString(request.params.arguments?.path, "path", true);

      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set activeDoc to active document
          save as activeDoc file name ${JSON.stringify(path)} file format format PDF
          return "Document exported as PDF to " & ${JSON.stringify(path)}
        end tell
      `;

      try {
        const result = await runAppleScript(script);
        return {
          content: [
            {
              type: "text",
              text: result,
            },
          ],
        };
      } catch (error) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to export PDF: ${getErrorMessage(error)}`,
            },
          ],
          isError: true,
        };
      }
    }

    case "list_tables": {
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

      try {
        const result = await runAppleScript(script);
        return {
          content: [
            {
              type: "text",
              text: result,
            },
          ],
        };
      } catch (error) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to list tables: ${getErrorMessage(error)}`,
            },
          ],
          isError: true,
        };
      }
    }

    case "get_table_cell": {
      const tableIndex = validateInteger(request.params.arguments?.tableIndex, "tableIndex", 1);
      const row = validateInteger(request.params.arguments?.row, "row", 1);
      const column = validateInteger(request.params.arguments?.column, "column", 1);

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

      try {
        const result = await runAppleScript(script);
        return {
          content: [
            {
              type: "text",
              text: result,
            },
          ],
        };
      } catch (error) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to get table cell: ${getErrorMessage(error)}`,
            },
          ],
          isError: true,
        };
      }
    }

    case "set_table_cell": {
      const tableIndex = validateInteger(request.params.arguments?.tableIndex, "tableIndex", 1);
      const row = validateInteger(request.params.arguments?.row, "row", 1);
      const column = validateInteger(request.params.arguments?.column, "column", 1);
      const text = validateString(request.params.arguments?.text, "text", true);

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

      try {
        const result = await runAppleScript(script);
        return {
          content: [
            {
              type: "text",
              text: result,
            },
          ],
        };
      } catch (error) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to set table cell: ${getErrorMessage(error)}`,
            },
          ],
          isError: true,
        };
      }
    }

    case "select_table_cell": {
      const tableIndex = validateInteger(request.params.arguments?.tableIndex, "tableIndex", 1);
      const row = validateInteger(request.params.arguments?.row, "row", 1);
      const column = validateInteger(request.params.arguments?.column, "column", 1);

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

      try {
        const result = await runAppleScript(script);
        return {
          content: [
            {
              type: "text",
              text: result,
            },
          ],
        };
      } catch (error) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to select table cell: ${getErrorMessage(error)}`,
            },
          ],
          isError: true,
        };
      }
    }

    case "find_table_header": {
      const tableIndex = validateInteger(request.params.arguments?.tableIndex, "tableIndex", 1);
      const headerText = validateString(request.params.arguments?.headerText, "headerText", true);
      const headerRow = validateInteger(request.params.arguments?.headerRow, "headerRow", 1) || 1;

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

      try {
        const result = await runAppleScript(script);
        return {
          content: [
            {
              type: "text",
              text: result,
            },
          ],
        };
      } catch (error) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to find table header: ${getErrorMessage(error)}`,
            },
          ],
          isError: true,
        };
      }
    }

    case "move_cursor_after_text": {
      const searchText = validateString(request.params.arguments?.searchText, "searchText", true);
      const occurrence = validateInteger(request.params.arguments?.occurrence, "occurrence", 1) || 1;

      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set activeDoc to active document

          -- Start from beginning of document
          select (text object of activeDoc)
          set selection end of selection to selection start of selection

          set findObj to find object of selection
          clear formatting findObj
          set content of findObj to ${JSON.stringify(searchText)}
          set wrap of findObj to find stop
          set forward of findObj to true

          set foundCount to 0
          repeat ${occurrence} times
            set prevStart to selection start of selection
            execute find findObj
            set selStart to selection start of selection
            set selEnd to selection end of selection
            if selStart is equal to selEnd then
              exit repeat
            end if
            set foundCount to foundCount + 1
            if foundCount < ${occurrence} then
              set selection end of selection to selEnd
              set selection start of selection to selEnd
            end if
          end repeat

          if foundCount < ${occurrence} then
            return "Text not found (or fewer than ${occurrence} occurrences): " & ${JSON.stringify(searchText)}
          end if

          -- Move cursor to end of found text
          set selection start of selection to selection end of selection
          return "Cursor moved after occurrence " & ${occurrence} & " of: " & ${JSON.stringify(searchText)}
        end tell
      `;

      try {
        const result = await runAppleScript(script);
        return {
          content: [
            {
              type: "text",
              text: result,
            },
          ],
        };
      } catch (error) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to move cursor: ${getErrorMessage(error)}`,
            },
          ],
          isError: true,
        };
      }
    }

    // === New tools v0.6.0 ===

    case "get_document_info": {
      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set d to active document
          set wordCount to compute statistics d statistic statistic words
          set charCount to compute statistics d statistic statistic characters
          set charWithSpaces to compute statistics d statistic statistic characters including spaces
          set paraCount to compute statistics d statistic statistic paragraphs
          set pageCount to compute statistics d statistic statistic pages
          return "Words: " & wordCount & linefeed & "Characters: " & charCount & linefeed & "Characters (with spaces): " & charWithSpaces & linefeed & "Paragraphs: " & paraCount & linefeed & "Pages: " & pageCount
        end tell
      `;

      try {
        const result = await runAppleScript(script);
        return {
          content: [{ type: "text", text: result }],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: `Failed to get document info: ${getErrorMessage(error)}` }],
          isError: true,
        };
      }
    }

    case "goto_start": {
      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set d to active document
          select (text object of d)
          set selection end of selection to selection start of selection
          return "Cursor moved to start of document"
        end tell
      `;

      try {
        const result = await runAppleScript(script);
        return {
          content: [{ type: "text", text: result }],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: `Failed to go to start: ${getErrorMessage(error)}` }],
          isError: true,
        };
      }
    }

    case "goto_end": {
      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set d to active document
          select (text object of d)
          set selection start of selection to selection end of selection
          return "Cursor moved to end of document"
        end tell
      `;

      try {
        const result = await runAppleScript(script);
        return {
          content: [{ type: "text", text: result }],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: `Failed to go to end: ${getErrorMessage(error)}` }],
          isError: true,
        };
      }
    }

    case "get_selection_info": {
      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set selStart to selection start of selection
          set selEnd to selection end of selection
          set selLength to selEnd - selStart
          return "Start: " & selStart & linefeed & "End: " & selEnd & linefeed & "Length: " & selLength
        end tell
      `;

      try {
        const result = await runAppleScript(script);
        return {
          content: [{ type: "text", text: result }],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: `Failed to get selection info: ${getErrorMessage(error)}` }],
          isError: true,
        };
      }
    }

    case "select_all": {
      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set d to active document
          select (text object of d)
          return "All content selected"
        end tell
      `;

      try {
        const result = await runAppleScript(script);
        return {
          content: [{ type: "text", text: result }],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: `Failed to select all: ${getErrorMessage(error)}` }],
          isError: true,
        };
      }
    }

    case "create_table": {
      const rows = validateInteger(request.params.arguments?.rows, "rows", 1);
      const columns = validateInteger(request.params.arguments?.columns, "columns", 1);

      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          make new table at selection with properties {number of rows:${rows}, number of columns:${columns}}
          return "Table created with ${rows} rows and ${columns} columns"
        end tell
      `;

      try {
        const result = await runAppleScript(script);
        return {
          content: [{ type: "text", text: result }],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: `Failed to create table: ${getErrorMessage(error)}` }],
          isError: true,
        };
      }
    }

    case "add_table_row": {
      const tableIndex = validateInteger(request.params.arguments?.tableIndex, "tableIndex", 1);
      const afterRow = validateInteger(request.params.arguments?.afterRow, "afterRow", 1);

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

      try {
        const result = await runAppleScript(script);
        return {
          content: [{ type: "text", text: result }],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: `Failed to add table row: ${getErrorMessage(error)}` }],
          isError: true,
        };
      }
    }

    case "delete_table_row": {
      const tableIndex = validateInteger(request.params.arguments?.tableIndex, "tableIndex", 1);
      const row = validateInteger(request.params.arguments?.row, "row", 1);

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

      try {
        const result = await runAppleScript(script);
        return {
          content: [{ type: "text", text: result }],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: `Failed to delete table row: ${getErrorMessage(error)}` }],
          isError: true,
        };
      }
    }

    case "add_table_column": {
      const tableIndex = validateInteger(request.params.arguments?.tableIndex, "tableIndex", 1);
      const afterColumn = validateInteger(request.params.arguments?.afterColumn, "afterColumn", 1);

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

      try {
        const result = await runAppleScript(script);
        return {
          content: [{ type: "text", text: result }],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: `Failed to add table column: ${getErrorMessage(error)}` }],
          isError: true,
        };
      }
    }

    case "delete_table_column": {
      const tableIndex = validateInteger(request.params.arguments?.tableIndex, "tableIndex", 1);
      const column = validateInteger(request.params.arguments?.column, "column", 1);

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

      try {
        const result = await runAppleScript(script);
        return {
          content: [{ type: "text", text: result }],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: `Failed to delete table column: ${getErrorMessage(error)}` }],
          isError: true,
        };
      }
    }

    case "list_bookmarks": {
      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set d to active document
          set bookmarkCount to count of bookmarks of d
          if bookmarkCount = 0 then
            return "No bookmarks found"
          end if
          set bookmarkList to ""
          repeat with i from 1 to bookmarkCount
            set b to bookmark i of d
            set bookmarkList to bookmarkList & i & ". " & (name of b) & linefeed
          end repeat
          return bookmarkList
        end tell
      `;

      try {
        const result = await runAppleScript(script);
        return {
          content: [{ type: "text", text: result }],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: `Failed to list bookmarks: ${getErrorMessage(error)}` }],
          isError: true,
        };
      }
    }

    case "create_bookmark": {
      const name = validateString(request.params.arguments?.name, "name", true);

      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set d to active document
          make new bookmark at d with properties {name:${JSON.stringify(name)}, bookmark range:selection}
          return "Bookmark created: " & ${JSON.stringify(name)}
        end tell
      `;

      try {
        const result = await runAppleScript(script);
        return {
          content: [{ type: "text", text: result }],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: `Failed to create bookmark: ${getErrorMessage(error)}` }],
          isError: true,
        };
      }
    }

    case "goto_bookmark": {
      const name = validateString(request.params.arguments?.name, "name", true);

      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set d to active document
          set b to bookmark ${JSON.stringify(name)} of d
          select (bookmark range of b)
          return "Jumped to bookmark: " & ${JSON.stringify(name)}
        end tell
      `;

      try {
        const result = await runAppleScript(script);
        return {
          content: [{ type: "text", text: result }],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: `Failed to go to bookmark: ${getErrorMessage(error)}` }],
          isError: true,
        };
      }
    }

    case "delete_bookmark": {
      const name = validateString(request.params.arguments?.name, "name", true);

      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set d to active document
          delete bookmark ${JSON.stringify(name)} of d
          return "Bookmark deleted: " & ${JSON.stringify(name)}
        end tell
      `;

      try {
        const result = await runAppleScript(script);
        return {
          content: [{ type: "text", text: result }],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: `Failed to delete bookmark: ${getErrorMessage(error)}` }],
          isError: true,
        };
      }
    }

    case "list_hyperlinks": {
      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set d to active document
          set linkCount to count of hyperlinks of d
          if linkCount = 0 then
            return "No hyperlinks found"
          end if
          set linkList to ""
          repeat with i from 1 to linkCount
            set h to hyperlink i of d
            set linkAddress to hyperlink address of h
            set linkText to ""
            try
              set linkText to content of text object of text range of h
            end try
            set linkList to linkList & i & ". " & linkText & " -> " & linkAddress & linefeed
          end repeat
          return linkList
        end tell
      `;

      try {
        const result = await runAppleScript(script);
        return {
          content: [{ type: "text", text: result }],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: `Failed to list hyperlinks: ${getErrorMessage(error)}` }],
          isError: true,
        };
      }
    }

    case "create_hyperlink": {
      const url = validateString(request.params.arguments?.url, "url", true);
      const displayText = validateString(request.params.arguments?.displayText, "displayText", false);

      const script = displayText
        ? `
          tell application "Microsoft Word"
            if (count of documents) = 0 then
              return "No document is open"
            end if
            set d to active document
            make new hyperlink at selection with properties {hyperlink address:${JSON.stringify(url)}, text to display:${JSON.stringify(displayText)}}
            return "Hyperlink created: " & ${JSON.stringify(url)}
          end tell
        `
        : `
          tell application "Microsoft Word"
            if (count of documents) = 0 then
              return "No document is open"
            end if
            set d to active document
            make new hyperlink at selection with properties {hyperlink address:${JSON.stringify(url)}}
            return "Hyperlink created: " & ${JSON.stringify(url)}
          end tell
        `;

      try {
        const result = await runAppleScript(script);
        return {
          content: [{ type: "text", text: result }],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: `Failed to create hyperlink: ${getErrorMessage(error)}` }],
          isError: true,
        };
      }
    }

    case "list_paragraphs": {
      const limit = validateInteger(request.params.arguments?.limit, "limit", 1) || 50;

      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set d to active document
          set paraCount to count of paragraphs of d
          set maxPara to ${limit}
          if paraCount < maxPara then
            set maxPara to paraCount
          end if
          set paraList to "Total paragraphs: " & paraCount & linefeed & linefeed
          repeat with i from 1 to maxPara
            set p to paragraph i of d
            set pStyle to name of paragraph style of p
            set pText to content of text object of p
            if length of pText > 50 then
              set pText to text 1 thru 50 of pText & "..."
            end if
            -- Remove line breaks for display
            set pText to my replaceText(pText, return, " ")
            set pText to my replaceText(pText, linefeed, " ")
            set paraList to paraList & i & ". [" & pStyle & "] " & pText & linefeed
          end repeat
          return paraList
        end tell

        on replaceText(theText, searchString, replacementString)
          set AppleScript's text item delimiters to searchString
          set theTextItems to text items of theText
          set AppleScript's text item delimiters to replacementString
          set theText to theTextItems as text
          set AppleScript's text item delimiters to ""
          return theText
        end replaceText
      `;

      try {
        const result = await runAppleScript(script);
        return {
          content: [{ type: "text", text: result }],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: `Failed to list paragraphs: ${getErrorMessage(error)}` }],
          isError: true,
        };
      }
    }

    case "goto_paragraph": {
      const index = validateInteger(request.params.arguments?.index, "index", 1);

      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set d to active document
          set paraCount to count of paragraphs of d
          if ${index} > paraCount then
            return "Paragraph index out of range. Document has " & paraCount & " paragraphs."
          end if
          set p to paragraph ${index} of d
          select (text object of p)
          return "Jumped to paragraph ${index}"
        end tell
      `;

      try {
        const result = await runAppleScript(script);
        return {
          content: [{ type: "text", text: result }],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: `Failed to go to paragraph: ${getErrorMessage(error)}` }],
          isError: true,
        };
      }
    }

    case "set_paragraph_style": {
      const index = validateInteger(request.params.arguments?.index, "index", 1);
      const styleName = validateString(request.params.arguments?.styleName, "styleName", true);

      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set d to active document
          set paraCount to count of paragraphs of d
          if ${index} > paraCount then
            return "Paragraph index out of range. Document has " & paraCount & " paragraphs."
          end if
          set p to paragraph ${index} of d
          set paragraph style of p to ${JSON.stringify(styleName)}
          return "Style " & ${JSON.stringify(styleName)} & " applied to paragraph ${index}"
        end tell
      `;

      try {
        const result = await runAppleScript(script);
        return {
          content: [{ type: "text", text: result }],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: `Failed to set paragraph style: ${getErrorMessage(error)}` }],
          isError: true,
        };
      }
    }

    default:
      throw new Error("Unknown tool");
  }
});

async function main() {
  try {
    const transport = new StdioServerTransport();
    await server.connect(transport);
  } catch (error) {
    console.error("Server error:", error);
    process.exit(1);
  }
}

main();
