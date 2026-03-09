import { validateBoolean, validateExcelRangeReference, validateNumber, validateString } from '../lib/validators.js';
import { runAppleScript } from '../lib/applescript/executor.js';
import { wrapExcelScript } from '../lib/applescript/script-wrappers.js';

export const excelFormattingTools = [
  {
    name: 'excel_format_cells',
    description: 'Format cells in Excel (bold, italic, font, size, font color)',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        range: { type: 'string', description: 'Range to format (e.g., "A1", "A1:B3")' },
        bold: { type: 'boolean', description: 'Make text bold' },
        italic: { type: 'boolean', description: 'Make text italic' },
        font: { type: 'string', description: 'Font name (e.g., "Arial", "Helvetica")' },
        size: { type: 'number', description: 'Font size in points' },
        fontColor: {
          type: 'array',
          description: 'Font color as [R, G, B] (0-255 each)',
          items: { type: 'number' }
        }
      },
      required: ['range']
    },
    async handler(args) {
      const range = validateExcelRangeReference(args.range, 'range');
      const formatCmds = [];
      if (args.bold !== undefined) formatCmds.push(`set bold of font object of r to ${validateBoolean(args.bold, 'bold')}`);
      if (args.italic !== undefined) formatCmds.push(`set italic of font object of r to ${validateBoolean(args.italic, 'italic')}`);
      if (args.font) formatCmds.push(`set name of font object of r to ${JSON.stringify(args.font)}`);
      if (args.size !== undefined) {
        const size = validateNumber(args.size, 'size', 1, 409);
        formatCmds.push(`set font size of font object of r to ${size}`);
      }
      if (args.fontColor) {
        if (!Array.isArray(args.fontColor) || args.fontColor.length !== 3 || !args.fontColor.every(v => typeof v === 'number' && v >= 0 && v <= 255)) {
          throw new Error('fontColor must be an array of 3 numbers [R,G,B] with values 0-255');
        }
        formatCmds.push(`set color of font object of r to {${args.fontColor.join(', ')}}`);
      }
      if (formatCmds.length === 0) throw new Error('At least one formatting option is required');

      const script = wrapExcelScript(`
set ws to active sheet
try
  set r to range ${JSON.stringify(range)} of ws
on error
  return "Invalid range: ${range}"
end try
try
${formatCmds.join('\n')}
on error errMsg
  return "Error applying formatting: " & errMsg
end try
return "Formatting applied to ${range}"
`);
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
        range: { type: 'string', description: 'Range to format (e.g., "A1:A10")' },
        format: { type: 'string', description: 'Number format string (e.g., "#,##0.00", "0%", "yyyy-mm-dd")' }
      },
      required: ['range', 'format']
    },
    async handler(args) {
      const range = validateExcelRangeReference(args.range, 'range');
      const format = validateString(args.format, 'format', true);
      const script = wrapExcelScript(`
set ws to active sheet
try
  set number format of range ${JSON.stringify(range)} of ws to ${JSON.stringify(format)}
on error errMsg
  return "Error setting number format: " & errMsg
end try
return "Number format set for " & ${JSON.stringify(range)} & ": " & ${JSON.stringify(format)}
`);
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
        range: { type: 'string', description: 'Range to color (e.g., "A1", "A1:B3")' },
        color: { type: 'array', description: 'Background color as [R, G, B] (0-255 each)', items: { type: 'number' } }
      },
      required: ['range', 'color']
    },
    async handler(args) {
      const range = validateExcelRangeReference(args.range, 'range');
      if (!args.color || !Array.isArray(args.color) || args.color.length !== 3 || !args.color.every(v => typeof v === 'number' && v >= 0 && v <= 255)) {
        throw new Error('color must be an array of 3 numbers [R,G,B] with values 0-255');
      }
      const script = wrapExcelScript(`
set ws to active sheet
try
  set color of interior object of range ${JSON.stringify(range)} of ws to {${args.color.join(', ')}}
on error errMsg
  return "Error setting cell color: " & errMsg
end try
return "Background color set for ${range}"
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'excel_merge_cells',
    description: 'Merge a range of cells in the active Excel worksheet',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: { range: { type: 'string', description: 'Range to merge (e.g., "A1:B1")' } },
      required: ['range']
    },
    async handler(args) {
      const range = validateExcelRangeReference(args.range, 'range');
      const script = wrapExcelScript(`
set ws to active sheet
try
  merge (range ${JSON.stringify(range)} of ws)
on error errMsg
  return "Error merging cells: " & errMsg
end try
return "Cells merged: ${range}"
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'excel_autofit',
    description: 'Auto-fit column widths for a range in the active Excel worksheet',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: { range: { type: 'string', description: 'Range or columns to autofit (e.g., "A:C", "A1:D10")' } },
      required: ['range']
    },
    async handler(args) {
      const range = validateExcelRangeReference(args.range, 'range');
      const script = wrapExcelScript(`
set ws to active sheet
try
  set r to entire column of range ${JSON.stringify(range)} of ws
  autofit r
on error errMsg
  return "Error auto-fitting columns: " & errMsg
end try
return "Columns auto-fitted for ${range}"
`);
      return await runAppleScript(script);
    }
  }
];
