import { validateString } from '../lib/validators.js';
import { runAppleScript } from '../lib/applescript/executor.js';
import { quoteAppleScriptString } from '../lib/applescript/helpers.js';

export const hyperlinkTools = [
  {
    name: 'word_list_hyperlinks',
    description: 'List all hyperlinks in the active Word document',
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
          set linkCount to count of hyperlink objects of d
          if linkCount = 0 then
            return "No hyperlinks found"
          end if
          set linkList to ""
          repeat with i from 1 to linkCount
            set h to hyperlink object i of d
            set linkAddress to hyperlink address of h
            set linkText to ""
            try
              set linkText to text to display of h
            end try
            if linkText is "" then set linkText to "(no text)"
            set linkList to linkList & i & ". " & linkText & " -> " & linkAddress & linefeed
          end repeat
          return linkList
        end tell
      `;

      return await runAppleScript(script);
    }
  },

  {
    name: 'word_create_hyperlink',
    description: 'Create a hyperlink at the current selection or cursor position in Word',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        url: {
          type: 'string',
          description: 'URL for the hyperlink'
        },
        displayText: {
          type: 'string',
          description: 'Optional text to display (if not provided, uses current selection or URL)'
        }
      },
      required: ['url']
    },
    async handler(args) {
      const url = validateString(args.url, 'url', true);
      const displayText = validateString(args.displayText, 'displayText', false);

      const props = displayText
        ? `{|hyperlink address|:${quoteAppleScriptString(url)}, |text to display|:${quoteAppleScriptString(displayText)}, |text object|:theRange}`
        : `{|hyperlink address|:${quoteAppleScriptString(url)}, |text object|:theRange}`;

      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set theRange to text object of selection
          try
            tell selection
              make new hyperlink object at end with properties ${props}
            end tell
          on error errMsg
            return "Error creating hyperlink: " & errMsg
          end try
          return "Hyperlink created: " & ${quoteAppleScriptString(url)}
        end tell
      `;

      return await runAppleScript(script);
    }
  }
];
