import { validateString, validateBoolean } from '../lib/validators.js';
import { runAppleScript } from '../lib/applescript/executor.js';

export const documentTools = [
  {
    name: 'word_create_document',
    description: 'Create a new Microsoft Word document with optional content',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        content: {
          type: 'string',
          description: 'Optional initial content for the document'
        }
      }
    },
    async handler(args) {
      const content = validateString(args.content, 'content', false);

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

      return await runAppleScript(script);
    }
  },

  {
    name: 'word_open_document',
    description: 'Open an existing Microsoft Word document',
    annotations: { readOnlyHint: false },
    inputSchema: {
      type: 'object',
      properties: {
        path: {
          type: 'string',
          description: 'Full path to the document to open'
        }
      },
      required: ['path']
    },
    async handler(args) {
      const path = validateString(args.path, 'path', true);

      const script = `
        tell application "Microsoft Word"
          activate
          open ${JSON.stringify(path)}
          return "Document opened successfully"
        end tell
      `;

      return await runAppleScript(script);
    }
  },

  {
    name: 'word_get_document_text',
    description: 'Get all text content from the active Word document',
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
          set activeDoc to active document
          return content of text object of activeDoc as string
        end tell
      `;

      return await runAppleScript(script);
    }
  },

  {
    name: 'word_get_document_info',
    description: 'Get Word document statistics (words, characters, paragraphs, pages)',
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
          set d to active document
          set wordCount to compute statistics d statistic statistic words
          set charCount to compute statistics d statistic statistic characters
          set charWithSpaces to compute statistics d statistic statistic characters with spaces
          set paraCount to compute statistics d statistic statistic paragraphs
          set pageCount to compute statistics d statistic statistic pages
          return "Words: " & wordCount & linefeed & "Characters: " & charCount & linefeed & "Characters (with spaces): " & charWithSpaces & linefeed & "Paragraphs: " & paraCount & linefeed & "Pages: " & pageCount
        end tell
      `;

      return await runAppleScript(script);
    }
  },

  {
    name: 'word_save_document',
    description: 'Save the active Word document',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        path: {
          type: 'string',
          description: 'Optional path to save as (if not provided, saves to current location)'
        }
      }
    },
    async handler(args) {
      const path = args.path ? validateString(args.path, 'path', false) : undefined;

      const script = path
        ? `
          tell application "Microsoft Word"
            if (count of documents) = 0 then
              return "No document is open"
            end if
            set activeDoc to active document
            save as activeDoc file name ${JSON.stringify(path)}
            return "Document saved as " & ${JSON.stringify(path)}
          end tell
        `
        : `
          tell application "Microsoft Word"
            if (count of documents) = 0 then
              return "No document is open"
            end if
            set activeDoc to active document
            save activeDoc
            return "Document saved successfully"
          end tell
        `;

      return await runAppleScript(script);
    }
  },

  {
    name: 'word_close_document',
    description: 'Close the active Word document',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        save: {
          type: 'boolean',
          description: 'Save before closing (default: true)',
          default: true
        }
      }
    },
    async handler(args) {
      const save = validateBoolean(args.save, 'save', true);

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

      return await runAppleScript(script);
    }
  },

  {
    name: 'word_export_pdf',
    description: 'Export the active Word document as PDF',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        path: {
          type: 'string',
          description: 'Full path for the PDF file'
        }
      },
      required: ['path']
    },
    async handler(args) {
      const path = validateString(args.path, 'path', true);

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

      return await runAppleScript(script);
    }
  }
];
