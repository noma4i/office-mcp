import { validateInteger } from '../lib/validators.js';
import { runAppleScript } from '../lib/applescript/executor.js';
import { wrapWordScript } from '../lib/applescript/script-wrappers.js';

export const clipboardTools = [
  {
    name: 'word_copy_content',
    description:
      'Copy content to the system clipboard preserving formatting. If startParagraph is specified, selects and copies that paragraph range. Without parameters, copies the current selection.',
    annotations: { readOnlyHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        startParagraph: { type: 'integer', description: 'First paragraph to copy (1-based). If omitted, copies current selection.' },
        endParagraph: { type: 'integer', description: 'Last paragraph to copy (1-based, default = startParagraph)' }
      }
    },
    async handler(args) {
      if (args.endParagraph !== undefined && args.startParagraph === undefined) {
        throw new Error('endParagraph requires startParagraph');
      }

      if (args.startParagraph !== undefined) {
        const startParagraph = validateInteger(args.startParagraph, 'startParagraph', 1);
        const endParagraph = args.endParagraph !== undefined ? validateInteger(args.endParagraph, 'endParagraph', 1) : startParagraph;
        if (endParagraph < startParagraph) {
          throw new Error('endParagraph must be >= startParagraph');
        }
        const script = wrapWordScript(`
set d to active document
set paraCount to count of paragraphs of d
if ${startParagraph} > paraCount then
  return "Start paragraph out of range. Document has " & paraCount & " paragraphs."
end if
if ${endParagraph} > paraCount then
  return "End paragraph out of range. Document has " & paraCount & " paragraphs."
end if
select (text object of paragraph ${startParagraph} of d)
set rStart to selection start of selection
select (text object of paragraph ${endParagraph} of d)
set rEnd to selection end of selection
set selection start of selection to rStart
set selection end of selection to rEnd
copy object selection
return "Copied paragraphs ${startParagraph} to ${endParagraph} to clipboard"
`);
        return await runAppleScript(script);
      }

      const script = wrapWordScript(`
try
  copy object selection
on error errMsg
  return "Error copying: " & errMsg
end try
return "Current selection copied to clipboard"
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'word_paste_content',
    description: 'Paste content from the system clipboard at the current cursor position preserving formatting.',
    annotations: { destructiveHint: true },
    inputSchema: { type: 'object', properties: {} },
    async handler() {
      const script = wrapWordScript(`
try
  paste object selection
on error errMsg
  return "Error pasting: " & errMsg
end try
return "Content pasted from clipboard"
`);
      return await runAppleScript(script);
    }
  }
];

