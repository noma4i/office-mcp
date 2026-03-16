import { validateExcelRangeReference, validateInteger, validateString } from '../lib/validators.js';
import { runAppleScript } from '../lib/applescript/executor.js';
import { quoteAppleScriptString, wrapExcelRowValues } from '../lib/applescript/helpers.js';
import { wrapExcelScript } from '../lib/applescript/script-wrappers.js';

function resolveWorksheetRef(value) {
  if (value === undefined || value === null) {
    return 'active sheet';
  }
  if (typeof value === 'number') {
    return `worksheet ${validateInteger(value, 'worksheet', 1)} of wb`;
  }
  return `worksheet ${quoteAppleScriptString(validateString(value, 'worksheet', true))} of wb`;
}

function parseTsv(values) {
  const normalized = validateString(values, 'values', true).replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  const rows = normalized.split('\n').map(row => row.split('\t'));
  const width = rows[0].length;
  if (rows.some(row => row.length !== width)) {
    throw new Error('values must be a rectangular TSV grid');
  }
  return rows;
}

function matrixToAppleScript(rows) {
  return `{${rows.map(row => wrapExcelRowValues(row)).join(', ')}}`;
}

export const excelWorkflowTools = [
  {
    name: 'excel_clear_worksheet',
    description: 'Clear the used range of the active Excel worksheet in place.',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        worksheet: { type: ['string', 'integer'], description: 'Worksheet name or 1-based index. Defaults to the active sheet.' }
      }
    },
    async handler(args) {
      const worksheetRef = resolveWorksheetRef(args.worksheet);
      const script = wrapExcelScript(`
set wb to active workbook
set ws to ${worksheetRef}
try
  set targetRange to used range of ws
  clear contents targetRange
on error errMsg
  return "Error clearing worksheet: " & errMsg
end try
return "Worksheet cleared"
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'excel_set_range_values',
    description: 'Set a rectangular Excel range from TSV text in the active workbook.',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        range: { type: 'string', description: 'Target range in A1 notation.' },
        values: { type: 'string', description: 'Rectangular TSV payload to write into the range.' },
        worksheet: { type: ['string', 'integer'], description: 'Worksheet name or 1-based index. Defaults to the active sheet.' }
      },
      required: ['range', 'values']
    },
    async handler(args) {
      const range = validateExcelRangeReference(args.range, 'range');
      const rows = parseTsv(args.values);
      const worksheetRef = resolveWorksheetRef(args.worksheet);
      const matrixLiteral = matrixToAppleScript(rows);
      const script = wrapExcelScript(`
set wb to active workbook
set ws to ${worksheetRef}
try
  set targetRange to range ${quoteAppleScriptString(range)} of ws
  set rangeValues to ${matrixLiteral}
  set value of targetRange to rangeValues
on error errMsg
  return "Error setting range ${range}: " & errMsg
end try
return "Range ${range} set successfully"
`);
      return await runAppleScript(script);
    }
  }
];
