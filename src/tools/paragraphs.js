import { validateString, validateInteger } from '../lib/validators.js';
import { runAppleScript } from '../lib/applescript/executor.js';

export const paragraphTools = [
  {
    name: "list_paragraphs",
    description: "List paragraphs with their styles (limited to first N paragraphs)",
    annotations: { readOnlyHint: true },
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
    async handler(args) {
      const limit = validateInteger(args.limit, "limit", 1) || 50;

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

      return await runAppleScript(script);
    }
  },

  {
    name: "goto_paragraph",
    description: "Jump to a specific paragraph by index (1-based)",
    annotations: { destructiveHint: true },
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
    async handler(args) {
      const index = validateInteger(args.index, "index", 1);

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

      return await runAppleScript(script);
    }
  },

  {
    name: "set_paragraph_style",
    description: "Apply a paragraph style to a specific paragraph",
    annotations: { destructiveHint: true },
    inputSchema: {
      type: "object",
      properties: {
        index: {
          type: "integer",
          description: "Paragraph index (1-based)",
        },
        styleName: {
          type: "string",
          description: "Style name (e.g., 'Heading 1', 'Normal')",
        },
      },
      required: ["index", "styleName"],
    },
    async handler(args) {
      const index = validateInteger(args.index, "index", 1);
      const styleName = validateString(args.styleName, "styleName", true);

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

      return await runAppleScript(script);
    }
  }
];
