import { validateString, validateInteger } from '../lib/validators.js';
import { runAppleScript } from '../lib/applescript/executor.js';
import { escapeForWordFind, quoteAppleScriptString } from '../lib/applescript/helpers.js';

export const navigationTools = [
  {
    name: 'word_goto_start',
    description: 'Move cursor to the beginning of the Word document',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {}
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
    name: 'word_goto_end',
    description: 'Move cursor to the end of the Word document',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {}
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
    name: 'word_get_selection_info',
    description: 'Get position and length of current selection in Word',
    annotations: { readOnlyHint: true },
    inputSchema: {
      type: 'object',
      properties: {}
    },
    async handler() {
      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          try
            set selStart to selection start of selection
            set selEnd to selection end of selection
          on error
            return "Cannot access selection position"
          end try
          set selLength to selEnd - selStart
          return "Start: " & selStart & linefeed & "End: " & selEnd & linefeed & "Length: " & selLength
        end tell
      `;

      return await runAppleScript(script);
    }
  },

  {
    name: 'word_select_all',
    description: 'Select all content in the Word document',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {}
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
    name: 'word_move_cursor_after_text',
    description: 'Find text and move cursor after the specified occurrence in Word',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        searchText: {
          type: 'string',
          description: 'Text to search for'
        },
        occurrence: {
          type: 'integer',
          description: 'Which occurrence to jump to (default: 1)',
          default: 1
        }
      },
      required: ['searchText']
    },
    async handler(args) {
      const searchText = validateString(args.searchText, 'searchText', true);
      const occurrence = validateInteger(args.occurrence, 'occurrence', 1) || 1;

      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set activeDoc to active document

          -- Start from beginning of document
          select (text object of activeDoc)
          set selection end of selection to selection start of selection

          try
            set findObj to find object of selection
          on error
            return "Cannot access find object. Make sure a document is active."
          end try
          clear formatting findObj
          set content of findObj to ${escapeForWordFind(searchText)}
          set wrap of findObj to find stop
          set forward of findObj to true

          set foundCount to 0
          repeat ${occurrence} times
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
            return "Text not found (or fewer than ${occurrence} occurrences): " & ${quoteAppleScriptString(searchText)}
          end if

          -- Move cursor to end of found text
          set selection start of selection to selection end of selection
          return "Cursor moved after occurrence " & ${occurrence} & " of: " & ${quoteAppleScriptString(searchText)}
        end tell
      `;

      return await runAppleScript(script);
    }
  }
];
