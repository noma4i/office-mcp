import { validateString, validateInteger, validateFindText } from '../lib/validators.js';
import { runAppleScript } from '../lib/applescript/executor.js';
import { WORD_FIND_MODES, runWordFindWithFallback } from '../lib/applescript/word-find.js';
import { wrapWordScript } from '../lib/applescript/script-wrappers.js';

export const navigationTools = [
  {
    name: 'word_goto_start',
    description: 'Move cursor to the beginning of the Word document',
    annotations: { destructiveHint: true },
    inputSchema: { type: 'object', properties: {} },
    async handler() {
      const script = wrapWordScript(`
set d to active document
select (text object of d)
set selection end of selection to selection start of selection
return "Cursor moved to start of document"
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'word_goto_end',
    description: 'Move cursor to the end of the Word document',
    annotations: { destructiveHint: true },
    inputSchema: { type: 'object', properties: {} },
    async handler() {
      const script = wrapWordScript(`
set d to active document
select (text object of d)
set selection start of selection to selection end of selection
return "Cursor moved to end of document"
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'word_get_selection_info',
    description: 'Get position and length of current selection in Word',
    annotations: { readOnlyHint: true },
    inputSchema: { type: 'object', properties: {} },
    async handler() {
      const script = wrapWordScript(`
try
  set selStart to selection start of selection
  set selEnd to selection end of selection
on error
  return "Cannot access selection position"
end try
set selLength to selEnd - selStart
return "Start: " & selStart & linefeed & "End: " & selEnd & linefeed & "Length: " & selLength
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'word_select_all',
    description: 'Select all content in the Word document',
    annotations: { destructiveHint: true },
    inputSchema: { type: 'object', properties: {} },
    async handler() {
      const script = wrapWordScript(`
set d to active document
select (text object of d)
return "All content selected"
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'word_move_cursor_after_text',
    description: 'Find text and move cursor after the specified occurrence in Word (max 255 chars)',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        searchText: { type: 'string', maxLength: 255, description: 'Text to search for (max 255 characters)' },
        occurrence: { type: 'integer', description: 'Which occurrence to jump to (default: 1)', default: 1 }
      },
      required: ['searchText']
    },
    async handler(args) {
      const searchText = validateFindText(args.searchText, 'searchText');
      const occurrence = validateInteger(args.occurrence, 'occurrence', 1) || 1;

      return await runWordFindWithFallback(
        {
          mode: WORD_FIND_MODES.MOVE_CURSOR_AFTER_TEXT,
          findText: searchText,
          occurrence
        },
        { executeAppleScript: runAppleScript }
      );
    }
  }
];
