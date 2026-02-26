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
          try
            set f to font object of selection
          on error
            return "No selection or font object not available"
          end try
          set fontName to "unknown"
          set fontSize to "unknown"
          set isBold to "unknown"
          set isItalic to "unknown"
          set isUnderline to "unknown"
          set fontColor to "unknown"
          try
            set fontName to name of f
          end try
          try
            set fontSize to font size of f
          end try
          try
            set isBold to bold of f
          end try
          try
            set isItalic to italic of f
          end try
          try
            set isUnderline to underline of f
          end try
          try
            set fontColor to color of f
          end try
          set output to "Font: " & fontName & linefeed
          set output to output & "Size: " & fontSize & linefeed
          set output to output & "Bold: " & isBold & linefeed
          set output to output & "Italic: " & isItalic & linefeed
          set output to output & "Underline: " & isUnderline & linefeed
          set output to output & "Color: " & fontColor & linefeed
          return output
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
          set pStyle to "unknown"
          set alignStr to "unknown"
          set spaceBefore to "unknown"
          set spaceAfter to "unknown"
          set li to "unknown"
          set ri to "unknown"
          set fli to "unknown"
          try
            set pStyle to name local of style of selection
          end try
          try
            set pf to paragraph format of selection
            set pAlign to paragraph alignment of pf
            set alignStr to "left"
            if pAlign is align paragraph center then
              set alignStr to "center"
            else if pAlign is align paragraph right then
              set alignStr to "right"
            else if pAlign is align paragraph justify then
              set alignStr to "justify"
            end if
          end try
          try
            set spaceBefore to space before of pf
          end try
          try
            set spaceAfter to space after of pf
          end try
          try
            set li to paragraph format left indent of pf
          end try
          try
            set ri to paragraph format right indent of pf
          end try
          try
            set fli to paragraph format first line indent of pf
          end try
          set output to "Style: " & pStyle & linefeed
          set output to output & "Alignment: " & alignStr & linefeed
          set output to output & "Space before: " & spaceBefore & " pts" & linefeed
          set output to output & "Space after: " & spaceAfter & " pts" & linefeed
          set output to output & "Left indent: " & li & " pts" & linefeed
          set output to output & "Right indent: " & ri & " pts" & linefeed
          set output to output & "First line indent: " & fli & " pts" & linefeed
          return output
        end tell
      `;
      return await runAppleScript(script);
    }
  }
];
