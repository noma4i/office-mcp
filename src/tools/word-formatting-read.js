import { runAppleScript } from '../lib/applescript/executor.js';

export const formattingReadTools = [
  {
    name: 'word_get_text_formatting',
    description: 'Get font formatting of the current selection in the active Word document (name, size, bold, italic, underline, color)',
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
          set f to font object of selection
          set fontName to name of f
          set fontSize to font size of f
          set isBold to bold of f
          set isItalic to italic of f
          set isUnderline to underline of f
          set fontColor to color of f
          set result to "Font: " & fontName & linefeed
          set result to result & "Size: " & fontSize & linefeed
          set result to result & "Bold: " & isBold & linefeed
          set result to result & "Italic: " & isItalic & linefeed
          set result to result & "Underline: " & isUnderline & linefeed
          set result to result & "Color: " & fontColor & linefeed
          return result
        end tell
      `;
      return await runAppleScript(script);
    }
  },

  {
    name: 'word_get_paragraph_formatting',
    description: 'Get paragraph formatting of the current selection in Word (style, alignment, spacing, indentation)',
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
          set pf to paragraph format of selection
          set pStyle to name local of style of selection
          set pAlign to paragraph alignment of pf
          set alignStr to "left"
          if pAlign is align paragraph center then
            set alignStr to "center"
          else if pAlign is align paragraph right then
            set alignStr to "right"
          else if pAlign is align paragraph justify then
            set alignStr to "justify"
          end if
          set spaceBefore to space before of pf
          set spaceAfter to space after of pf
          set li to paragraph format left indent of pf
          set ri to paragraph format right indent of pf
          set fli to paragraph format first line indent of pf
          set result to "Style: " & pStyle & linefeed
          set result to result & "Alignment: " & alignStr & linefeed
          set result to result & "Space before: " & spaceBefore & " pts" & linefeed
          set result to result & "Space after: " & spaceAfter & " pts" & linefeed
          set result to result & "Left indent: " & li & " pts" & linefeed
          set result to result & "Right indent: " & ri & " pts" & linefeed
          set result to result & "First line indent: " & fli & " pts" & linefeed
          return result
        end tell
      `;
      return await runAppleScript(script);
    }
  }
];
