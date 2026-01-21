import { validateString, validateBoolean, validateNumber } from '../lib/validators.js';
import { runAppleScript } from '../lib/applescript/executor.js';

export const textTools = [
  {
    name: "insert_text",
    description: "Insert text at the current cursor position",
    annotations: { destructiveHint: true },
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
    async handler(args) {
      const text = validateString(args.text, "text", true);

      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          type text selection text ${JSON.stringify(text)}
          return "Text inserted successfully"
        end tell
      `;

      return await runAppleScript(script);
    }
  },

  {
    name: "replace_text",
    description: "Find and replace text in the active document",
    annotations: { destructiveHint: true },
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
    async handler(args) {
      const find = validateString(args.find, "find", true);
      const replace = validateString(args.replace, "replace", true);
      const all = validateBoolean(args.all, "all", true);

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

      return await runAppleScript(script);
    }
  },

  {
    name: "format_text",
    description: "Apply formatting to the currently selected text",
    annotations: { destructiveHint: true },
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
          description: "Font name",
        },
        size: {
          type: "number",
          description: "Font size",
        },
      },
    },
    async handler(args) {
      const bold = args.bold !== undefined
        ? validateBoolean(args.bold, "bold")
        : undefined;
      const italic = args.italic !== undefined
        ? validateBoolean(args.italic, "italic")
        : undefined;
      const underline = args.underline !== undefined
        ? validateBoolean(args.underline, "underline")
        : undefined;
      const font = args.font
        ? validateString(args.font, "font", false)
        : undefined;
      const size = args.size !== undefined
        ? validateNumber(args.size, "size", 1, 1000)
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

      return await runAppleScript(script);
    }
  }
];
