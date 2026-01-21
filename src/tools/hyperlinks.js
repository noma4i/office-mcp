import { validateString } from '../lib/validators.js';
import { runAppleScript } from '../lib/applescript/executor.js';

export const hyperlinkTools = [
  {
    name: "list_hyperlinks",
    description: "List all hyperlinks in the active document",
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

      return await runAppleScript(script);
    }
  },

  {
    name: "create_hyperlink",
    description: "Create a hyperlink at the current selection or cursor position",
    annotations: { destructiveHint: true },
    inputSchema: {
      type: "object",
      properties: {
        url: {
          type: "string",
          description: "URL for the hyperlink",
        },
        displayText: {
          type: "string",
          description: "Optional text to display (if not provided, uses current selection or URL)",
        },
      },
      required: ["url"],
    },
    async handler(args) {
      const url = validateString(args.url, "url", true);
      const displayText = validateString(args.displayText, "displayText", false);

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

      return await runAppleScript(script);
    }
  }
];
