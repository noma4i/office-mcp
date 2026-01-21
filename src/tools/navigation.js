import { validateString, validateInteger } from '../lib/validators.js';
import { runAppleScript } from '../lib/applescript/executor.js';

export const navigationTools = [
  {
    name: "goto_start",
    description: "Move cursor to the beginning of the document",
    annotations: { destructiveHint: true },
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
          set d to active document
          select (text object of d)
          set selection end of selection to selection start of selection
          return "Cursor moved to start of document"
        end tell
      `;

      return await runAppleScript(script);
    }
  },

  {
    name: "goto_end",
    description: "Move cursor to the end of the document",
    annotations: { destructiveHint: true },
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
          set d to active document
          select (text object of d)
          set selection start of selection to selection end of selection
          return "Cursor moved to end of document"
        end tell
      `;

      return await runAppleScript(script);
    }
  },

  {
    name: "get_selection_info",
    description: "Get position and length of current selection",
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
          set selStart to selection start of selection
          set selEnd to selection end of selection
          set selLength to selEnd - selStart
          return "Start: " & selStart & linefeed & "End: " & selEnd & linefeed & "Length: " & selLength
        end tell
      `;

      return await runAppleScript(script);
    }
  },

  {
    name: "select_all",
    description: "Select all content in the document",
    annotations: { destructiveHint: true },
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
          set d to active document
          select (text object of d)
          return "All content selected"
        end tell
      `;

      return await runAppleScript(script);
    }
  },

  {
    name: "move_cursor_after_text",
    description: "Find text and move cursor to the end of the specified occurrence",
    annotations: { destructiveHint: true },
    inputSchema: {
      type: "object",
      properties: {
        searchText: {
          type: "string",
          description: "Text to search for",
        },
        occurrence: {
          type: "integer",
          description: "Which occurrence to jump to (default: 1)",
          default: 1,
        },
      },
      required: ["searchText"],
    },
    async handler(args) {
      const searchText = validateString(args.searchText, "searchText", true);
      const occurrence = validateInteger(args.occurrence, "occurrence", 1) || 1;

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

      return await runAppleScript(script);
    }
  }
];
