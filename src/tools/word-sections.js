import { validateInteger, validateNumber, validateString } from '../lib/validators.js';
import { runAppleScript } from '../lib/applescript/executor.js';

export const sectionTools = [
  {
    name: 'word_list_sections',
    description: 'List all sections in the active Word document with page setup info',
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
          set secCount to count of sections of d
          set result to "Total sections: " & secCount & linefeed & linefeed
          repeat with i from 1 to secCount
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
          end repeat
          return result
        end tell
      `;
      return await runAppleScript(script);
    }
  },

  {
    name: 'word_get_section_info',
    description: 'Get detailed section info (margins, orientation, header/footer settings) in Word',
    annotations: { readOnlyHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        index: {
          type: 'integer',
          description: 'Section index (1-based)'
        }
      },
      required: ['index']
    },
    async handler(args) {
      const index = validateInteger(args.index, 'index', 1);

      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set d to active document
          set secCount to count of sections of d
          if ${index} > secCount then
            return "Section index out of range. Document has " & secCount & " sections."
          end if
          set s to section ${index} of d
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
          set pw to page width of ps
          set ph to page height of ps
          set diffFirst to different first page header footer of ps
          set oddEven to odd and even pages header footer of ps
          set hdrDist to header distance of ps
          set ftrDist to footer distance of ps
          set result to "Section ${index}:" & linefeed
          set result to result & "  Orientation: " & orientStr & linefeed
          set result to result & "  Page size: " & (round pw) & " x " & (round ph) & " pts" & linefeed
          set result to result & "  Margins: left=" & (round lm) & " right=" & (round rm) & " top=" & (round tm) & " bottom=" & (round bm) & " pts" & linefeed
          set result to result & "  Header distance: " & (round hdrDist) & " pts" & linefeed
          set result to result & "  Footer distance: " & (round ftrDist) & " pts" & linefeed
          set result to result & "  Different first page: " & diffFirst & linefeed
          set result to result & "  Odd/even pages: " & oddEven & linefeed
          return result
        end tell
      `;
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
        index: {
          type: 'integer',
          description: 'Section index (1-based, default: 1)',
          default: 1
        },
        topMargin: {
          type: 'number',
          description: 'Top margin in points'
        },
        bottomMargin: {
          type: 'number',
          description: 'Bottom margin in points'
        },
        leftMargin: {
          type: 'number',
          description: 'Left margin in points'
        },
        rightMargin: {
          type: 'number',
          description: 'Right margin in points'
        },
        orientation: {
          type: 'string',
          description: 'Page orientation: "portrait" or "landscape"',
          enum: ['portrait', 'landscape']
        },
        differentFirstPage: {
          type: 'boolean',
          description: 'Enable different first page header/footer'
        }
      }
    },
    async handler(args) {
      const index = validateInteger(args.index, 'index', 1) || 1;

      let commands = [];
      if (args.topMargin !== undefined) {
        const v = validateNumber(args.topMargin, 'topMargin', 0, 1584);
        commands.push(`set top margin of ps to ${v}`);
      }
      if (args.bottomMargin !== undefined) {
        const v = validateNumber(args.bottomMargin, 'bottomMargin', 0, 1584);
        commands.push(`set bottom margin of ps to ${v}`);
      }
      if (args.leftMargin !== undefined) {
        const v = validateNumber(args.leftMargin, 'leftMargin', 0, 1584);
        commands.push(`set left margin of ps to ${v}`);
      }
      if (args.rightMargin !== undefined) {
        const v = validateNumber(args.rightMargin, 'rightMargin', 0, 1584);
        commands.push(`set right margin of ps to ${v}`);
      }
      if (args.orientation) {
        const orientVal = args.orientation === 'landscape' ? 'orient landscape' : 'orient portrait';
        commands.push(`set orientation of ps to ${orientVal}`);
      }
      if (args.differentFirstPage !== undefined) {
        commands.push(`set different first page header footer of ps to ${args.differentFirstPage}`);
      }

      if (commands.length === 0) {
        throw new Error('At least one page setup property is required');
      }

      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set d to active document
          set secCount to count of sections of d
          if ${index} > secCount then
            return "Section index out of range. Document has " & secCount & " sections."
          end if
          set ps to page setup of section ${index} of d
          ${commands.join('\n          ')}
          return "Page setup updated for section ${index}"
        end tell
      `;
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
      const breakType = typeMap[args.type] || typeMap['next_page'];

      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set r to text object of selection
          insert break at r break type ${breakType}
          return "Section break inserted (${args.type || 'next_page'})"
        end tell
      `;
      return await runAppleScript(script);
    }
  }
];
