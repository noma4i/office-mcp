import { validateString, validateBoolean, validateNumber } from '../lib/validators.js';
import { runAppleScript } from '../lib/applescript/executor.js';
import { toAppleScriptString } from '../lib/applescript/helpers.js';
import { WORD_FIND_MODES, runWordFindWithFallback } from '../lib/applescript/word-find.js';
import { wrapWordScript } from '../lib/applescript/script-wrappers.js';

export const textTools = [
  {
    name: 'word_insert_text',
    description: 'Insert text at the current cursor position in Microsoft Word',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        text: {
          type: 'string',
          description: 'Text to insert'
        }
      },
      required: ['text']
    },
    async handler(args) {
      const text = validateString(args.text, 'text', true);

      const script = wrapWordScript(`
type text selection text ${toAppleScriptString(text)}
return "Text inserted successfully"
`);

      return await runAppleScript(script);
    }
  },

  {
    name: 'word_replace_text',
    description: 'Find and replace text in the active Word document',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        find: {
          type: 'string',
          description: 'Text to find'
        },
        replace: {
          type: 'string',
          description: 'Text to replace with'
        },
        all: {
          type: 'boolean',
          description: 'Replace all occurrences (default: true)',
          default: true
        }
      },
      required: ['find']
    },
    async handler(args) {
      const find = validateString(args.find, 'find', true);
      const replace = args.replace !== undefined ? validateString(args.replace, 'replace', false) : '';
      const all = validateBoolean(args.all, 'all', true);

      return await runWordFindWithFallback(
        {
          mode: WORD_FIND_MODES.REPLACE,
          findText: find,
          replaceWith: replace,
          replaceAll: all
        },
        { executeAppleScript: runAppleScript }
      );
    }
  },

  {
    name: 'word_delete_text',
    description: 'Delete selected text or find and delete all occurrences of specific text in the active Word document',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        text: {
          type: 'string',
          description: 'Text to find and delete. If not provided, deletes the current selection.'
        }
      }
    },
    async handler(args) {
      if (args.text !== undefined) {
        const text = validateString(args.text, 'text', true);
        return await runWordFindWithFallback(
          {
            mode: WORD_FIND_MODES.DELETE_ALL,
            findText: text
          },
          { executeAppleScript: runAppleScript }
        );
      }

      const script = wrapWordScript(`
try
  delete (text object of selection)
on error
  return "No text selected to delete"
end try
return "Selected text deleted"
`);
      return await runAppleScript(script);
    }
  },

  {
    name: 'word_format_text',
    description: 'Apply formatting to the currently selected text in Word',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        bold: {
          type: 'boolean',
          description: 'Make text bold'
        },
        italic: {
          type: 'boolean',
          description: 'Make text italic'
        },
        underline: {
          type: 'boolean',
          description: 'Underline text'
        },
        font: {
          type: 'string',
          description: 'Font name'
        },
        size: {
          type: 'number',
          description: 'Font size'
        }
      }
    },
    async handler(args) {
      const bold = args.bold !== undefined ? validateBoolean(args.bold, 'bold') : undefined;
      const italic = args.italic !== undefined ? validateBoolean(args.italic, 'italic') : undefined;
      const underline = args.underline !== undefined ? validateBoolean(args.underline, 'underline') : undefined;
      const font = args.font ? validateString(args.font, 'font', false) : undefined;
      const size = args.size !== undefined ? validateNumber(args.size, 'size', 1, 1000) : undefined;

      const formatCommands = [];
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
        throw new Error('At least one formatting option is required');
      }

      const script = wrapWordScript(`
try
${formatCommands.join('\n')}
on error errMsg
  return "Error applying formatting: " & errMsg
end try
return "Formatting applied successfully"
`);

      return await runAppleScript(script);
    }
  }
];
