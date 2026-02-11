import { validateString, validateNumber, validateBoolean } from '../lib/validators.js';
import { runAppleScript } from '../lib/applescript/executor.js';

export const excelFormattingTools = [
  {
    name: 'excel_format_cells',
    description: 'Format cells in Excel (bold, italic, font, size, font color)',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Range to format (e.g., "A1", "A1:B3")'
        },
        bold: {
          type: 'boolean',
          description: 'Make text bold'
        },
        italic: {
          type: 'boolean',
          description: 'Make text italic'
        },
        font: {
          type: 'string',
          description: 'Font name (e.g., "Arial", "Helvetica")'
        },
        size: {
          type: 'number',
          description: 'Font size in points'
        },
        fontColor: {
          type: 'array',
          description: 'Font color as [R, G, B] (0-255 each)',
          items: { type: 'number' }
        }
      },
      required: ['range']
    },
    async handler(args) {
      const range = validateString(args.range, 'range', true);

      let formatCmds = [];
      if (args.bold !== undefined) {
        formatCmds.push(`set bold of font object of r to ${args.bold}`);
      }
      if (args.italic !== undefined) {
        formatCmds.push(`set italic of font object of r to ${args.italic}`);
      }
      if (args.font) {
        formatCmds.push(`set name of font object of r to ${JSON.stringify(args.font)}`);
      }
      if (args.size !== undefined) {
        const size = validateNumber(args.size, 'size', 1, 409);
        formatCmds.push(`set font size of font object of r to ${size}`);
      }
      if (args.fontColor && Array.isArray(args.fontColor) && args.fontColor.length === 3) {
        formatCmds.push(`set color of font object of r to {${args.fontColor.join(', ')}}`);
      }

      if (formatCmds.length === 0) {
        throw new Error('At least one formatting option is required');
      }

      const script = `
        tell application "Microsoft Excel"
          if (count of workbooks) = 0 then
            return "No workbook is open"
          end if
          set ws to active sheet
          set r to range ${JSON.stringify(range)} of ws
          ${formatCmds.join('\n          ')}
          return "Formatting applied to ${range}"
        end tell
      `;
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_set_number_format',
    description: 'Set the number format of a range in Excel',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Range to format (e.g., "A1:A10")'
        },
        format: {
          type: 'string',
          description: 'Number format string (e.g., "#,##0.00", "0%", "yyyy-mm-dd")'
        }
      },
      required: ['range', 'format']
    },
    async handler(args) {
      const range = validateString(args.range, 'range', true);
      const format = validateString(args.format, 'format', true);
      const script = `
        tell application "Microsoft Excel"
          if (count of workbooks) = 0 then
            return "No workbook is open"
          end if
          set ws to active sheet
          set number format of range ${JSON.stringify(range)} of ws to ${JSON.stringify(format)}
          return "Number format set for ${range}: ${format}"
        end tell
      `;
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_set_cell_color',
    description: 'Set the background color of a range in Excel',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Range to color (e.g., "A1", "A1:B3")'
        },
        color: {
          type: 'array',
          description: 'Background color as [R, G, B] (0-255 each)',
          items: { type: 'number' }
        }
      },
      required: ['range', 'color']
    },
    async handler(args) {
      const range = validateString(args.range, 'range', true);
      if (!args.color || !Array.isArray(args.color) || args.color.length !== 3) {
        throw new Error('color must be an array of 3 numbers [R, G, B]');
      }
      const script = `
        tell application "Microsoft Excel"
          if (count of workbooks) = 0 then
            return "No workbook is open"
          end if
          set ws to active sheet
          set color of interior object of range ${JSON.stringify(range)} of ws to {${args.color.join(', ')}}
          return "Background color set for ${range}"
        end tell
      `;
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_merge_cells',
    description: 'Merge a range of cells in the active Excel worksheet',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Range to merge (e.g., "A1:B1")'
        }
      },
      required: ['range']
    },
    async handler(args) {
      const range = validateString(args.range, 'range', true);
      const script = `
        tell application "Microsoft Excel"
          if (count of workbooks) = 0 then
            return "No workbook is open"
          end if
          set ws to active sheet
          merge (range ${JSON.stringify(range)} of ws)
          return "Cells merged: ${range}"
        end tell
      `;
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_autofit',
    description: 'Auto-fit column widths for a range in the active Excel worksheet',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Range or columns to autofit (e.g., "A:C", "A1:D10")'
        }
      },
      required: ['range']
    },
    async handler(args) {
      const range = validateString(args.range, 'range', true);
      const script = `
        tell application "Microsoft Excel"
          if (count of workbooks) = 0 then
            return "No workbook is open"
          end if
          set ws to active sheet
          set r to entire column of range ${JSON.stringify(range)} of ws
          autofit r
          return "Columns auto-fitted for ${range}"
        end tell
      `;
      return await runAppleScript(script);
    }
  }
];
