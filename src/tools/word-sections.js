import { validateInteger, validateNumber } from '../lib/validators.js';
import { runAppleScript } from '../lib/applescript/executor.js';
import { wrapWordScript } from '../lib/applescript/script-wrappers.js';

export const sectionTools = [
  {
    name: 'word_list_sections',
    description: 'List all sections in the active Word document with page setup info',
    annotations: { readOnlyHint: true },
    inputSchema: { type: 'object', properties: {} },
    async handler() {
      const script = wrapWordScript(`
set d to active document
set secCount to count of sections of d
set result to "Total sections: " & secCount & linefeed & linefeed
repeat with i from 1 to secCount
  try
    set s to section i of d
    set ps to page setup of s
    set orient to orientation of ps
    set orientStr to "portrait"
    if orient is orient landscape then
      set orientStr to "landscape"
    end if
    set lm to left margin of ps
    set rm to right margin of ps
    set tm to top margin of ps
    set bm to bottom margin of ps
    set diffFirst to different first page header footer of ps
    set result to result & "Section " & i & ": " & orientStr & ", margins: L=" & (round lm) & " R=" & (round rm) & " T=" & (round tm) & " B=" & (round bm) & " pts, different first page: " & diffFirst & linefeed
  on error
    set result to result & "Section " & i & ": (not accessible)" & linefeed
  end try
end repeat
return result
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'word_get_section_info',
    description: 'Get detailed section info (margins, orientation, header/footer settings) in Word',
    annotations: { readOnlyHint: true },
    inputSchema: { type: 'object', properties: { index: { type: 'integer', description: 'Section index (1-based)' } }, required: ['index'] },
    async handler(args) {
      const index = validateInteger(args.index, 'index', 1);
      const script = wrapWordScript(`
set d to active document
set secCount to count of sections of d
if ${index} > secCount then
  return "Section index out of range. Document has " & secCount & " sections."
end if
try
  set s to section ${index} of d
  set ps to page setup of s
on error
  return "Section ${index} not accessible"
end try
set orient to orientation of ps
set orientStr to "portrait"
if orient is orient landscape then
  set orientStr to "landscape"
end if
set lm to left margin of ps
set rm to right margin of ps
set tm to top margin of ps
set bm to bottom margin of ps
set pw to page width of ps
set ph to page height of ps
set diffFirst to different first page header footer of ps
set oddEven to odd and even pages header footer of ps
set hdrDist to header distance of ps
set ftrDist to footer distance of ps
set output to "Section ${index}:" & linefeed
set output to output & "  Orientation: " & orientStr & linefeed
set output to output & "  Page size: " & (round pw) & " x " & (round ph) & " pts" & linefeed
set output to output & "  Margins: left=" & (round lm) & " right=" & (round rm) & " top=" & (round tm) & " bottom=" & (round bm) & " pts" & linefeed
set output to output & "  Header distance: " & (round hdrDist) & " pts" & linefeed
set output to output & "  Footer distance: " & (round ftrDist) & " pts" & linefeed
set output to output & "  Different first page: " & diffFirst & linefeed
set output to output & "  Odd/even pages: " & oddEven & linefeed
return output
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'word_set_page_setup',
    description: 'Set page setup properties (margins, orientation) for a section in Word',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        index: { type: 'integer', description: 'Section index (1-based, default: 1)', default: 1 },
        topMargin: { type: 'number', description: 'Top margin in points' },
        bottomMargin: { type: 'number', description: 'Bottom margin in points' },
        leftMargin: { type: 'number', description: 'Left margin in points' },
        rightMargin: { type: 'number', description: 'Right margin in points' },
        orientation: { type: 'string', description: 'Page orientation: "portrait" or "landscape"', enum: ['portrait', 'landscape'] },
        differentFirstPage: { type: 'boolean', description: 'Enable different first page header/footer' }
      }
    },
    async handler(args) {
      const index = validateInteger(args.index, 'index', 1) || 1;
      const commands = [];
      if (args.topMargin !== undefined) commands.push(`set top margin of ps to ${validateNumber(args.topMargin, 'topMargin', 0, 1584)}`);
      if (args.bottomMargin !== undefined) commands.push(`set bottom margin of ps to ${validateNumber(args.bottomMargin, 'bottomMargin', 0, 1584)}`);
      if (args.leftMargin !== undefined) commands.push(`set left margin of ps to ${validateNumber(args.leftMargin, 'leftMargin', 0, 1584)}`);
      if (args.rightMargin !== undefined) commands.push(`set right margin of ps to ${validateNumber(args.rightMargin, 'rightMargin', 0, 1584)}`);
      if (args.orientation) commands.push(`set orientation of ps to ${args.orientation === 'landscape' ? 'orient landscape' : 'orient portrait'}`);
      if (args.differentFirstPage !== undefined) commands.push(`set different first page header footer of ps to ${args.differentFirstPage}`);
      if (commands.length === 0) throw new Error('At least one page setup property is required');

      const script = wrapWordScript(`
set d to active document
set secCount to count of sections of d
if ${index} > secCount then
  return "Section index out of range. Document has " & secCount & " sections."
end if
try
  set ps to page setup of section ${index} of d
${commands.join('\n')}
on error errMsg
  return "Error updating page setup: " & errMsg
end try
return "Page setup updated for section ${index}"
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'word_insert_section_break',
    description: 'Insert a section break at the current cursor position in Word',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        type: {
          type: 'string',
          description: 'Break type: "next_page" (default), "continuous", "even_page", "odd_page"',
          enum: ['next_page', 'continuous', 'even_page', 'odd_page'],
          default: 'next_page'
        }
      }
    },
    async handler(args) {
      const typeMap = {
        next_page: 'section break next page',
        continuous: 'section break continuous',
        even_page: 'section break even page',
        odd_page: 'section break odd page'
      };
      const breakType = typeMap[args.type] || typeMap.next_page;
      const script = wrapWordScript(`
set r to text object of selection
try
  insert break at r break type ${breakType}
on error errMsg
  return "Error inserting section break: " & errMsg
end try
return "Section break inserted (${args.type || 'next_page'})"
`);
      return await runAppleScript(script);
    }
  }
];

