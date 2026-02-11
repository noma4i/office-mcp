import { validateString, validateBoolean } from '../lib/validators.js';
import { runAppleScript } from '../lib/applescript/executor.js';

export const excelDataTools = [
  {
    name: 'excel_sort_range',
    description: 'Sort a range by a key column in the active Excel worksheet',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Range to sort (e.g., "A1:D10")'
        },
        keyCell: {
          type: 'string',
          description: 'Key cell for sorting (e.g., "B1" to sort by column B)'
        },
        ascending: {
          type: 'boolean',
          description: 'Sort ascending (default: true)',
          default: true
        },
        hasHeader: {
          type: 'boolean',
          description: 'First row is header (default: true)',
          default: true
        }
      },
      required: ['range', 'keyCell']
    },
    async handler(args) {
      const range = validateString(args.range, 'range', true);
      const keyCell = validateString(args.keyCell, 'keyCell', true);
      const ascending = validateBoolean(args.ascending, 'ascending', true);
      const hasHeader = validateBoolean(args.hasHeader, 'hasHeader', true);
      const orderStr = ascending ? 'sort ascending' : 'sort descending';
      const headerStr = hasHeader ? 'header header yes' : 'header header no';

      const script = `
        tell application "Microsoft Excel"
          if (count of workbooks) = 0 then
            return "No workbook is open"
          end if
          set ws to active sheet
          sort range (range ${JSON.stringify(range)} of ws) key1 (range ${JSON.stringify(keyCell)} of ws) order1 ${orderStr} ${headerStr}
          return "Range ${range} sorted by ${keyCell} ${ascending ? 'ascending' : 'descending'}"
        end tell
      `;
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_calculate',
    description: 'Recalculate all formulas in all open Excel workbooks',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {}
    },
    async handler() {
      const script = `
        tell application "Microsoft Excel"
          if (count of workbooks) = 0 then
            return "No workbook is open"
          end if
          calculate
          return "All formulas recalculated"
        end tell
      `;
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_export_csv',
    description: 'Export the active Excel workbook as CSV',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        path: {
          type: 'string',
          description: 'Full path for the CSV file'
        }
      },
      required: ['path']
    },
    async handler(args) {
      const path = validateString(args.path, 'path', true);
      const script = `
        tell application "Microsoft Excel"
          if (count of workbooks) = 0 then
            return "No workbook is open"
          end if
          set display alerts to false
          set wb to active workbook
          set ws to active sheet of wb
          save as ws filename ${JSON.stringify(path)} file format CSV file format
          set display alerts to true
          return "Exported as CSV to " & ${JSON.stringify(path)}
        end tell
      `;
      return await runAppleScript(script);
    }
  }
];
