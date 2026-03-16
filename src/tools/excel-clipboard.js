import { ToolError } from '../lib/errors.js';
import { runAppleScript } from '../lib/applescript/executor.js';
import { validateExcelCellReference, validateExcelRangeReference, validateInteger, validateString } from '../lib/validators.js';

function resolveWorksheetRef(value, workbookRef = 'wb') {
  if (value === undefined || value === null) {
    return `active sheet`;
  }
  if (typeof value === 'number') {
    return `worksheet ${validateInteger(value, 'worksheet', 1)} of ${workbookRef}`;
  }
  return `worksheet ${JSON.stringify(validateString(value, 'worksheet', true))} of ${workbookRef}`;
}

function buildExcelCopyRangeScript(range, worksheetRef) {
  return `
tell application "Microsoft Excel"
  if (count of workbooks) = 0 then
    return "No workbook is open"
  end if
  activate
  set wb to active workbook
  set ws to ${worksheetRef}
  try
    select range ${JSON.stringify(range)} of ws
  on error
    return "Invalid range: ${range}"
  end try
end tell
delay 0.2
tell application "System Events"
  keystroke "c" using command down
end tell
return "Range copied to clipboard"
`;
}

function buildExcelPasteRangeScript(targetCell, worksheetRef) {
  return `
tell application "Microsoft Excel"
  if (count of workbooks) = 0 then
    return "No workbook is open"
  end if
  activate
  set wb to active workbook
  set ws to ${worksheetRef}
  try
    select range ${JSON.stringify(targetCell)} of ws
  on error
    return "Cell ${targetCell} not accessible"
  end try
end tell
delay 0.2
tell application "System Events"
  keystroke "v" using command down
end tell
return "Range pasted from clipboard"
`;
}

export const excelClipboardTools = [
  {
    name: 'excel_copy_range',
    description: 'Copy an Excel range with formatting to the system clipboard.',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        range: { type: 'string', description: 'Range to copy (e.g., "A1:C5").' },
        worksheet: { type: ['string', 'integer'], description: 'Worksheet name or 1-based index. Defaults to the active sheet.' }
      },
      required: ['range']
    },
    async handler(args) {
      const range = validateExcelRangeReference(args.range, 'range');
      const worksheetRef = resolveWorksheetRef(args.worksheet);
      return await runAppleScript(buildExcelCopyRangeScript(range, worksheetRef));
    }
  },
  {
    name: 'excel_paste_range',
    description: 'Paste the system clipboard into the active Excel workbook at the target cell.',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        targetCell: { type: 'string', description: 'Top-left target cell in A1 notation.' },
        worksheet: { type: ['string', 'integer'], description: 'Worksheet name or 1-based index. Defaults to the active sheet.' }
      },
      required: ['targetCell']
    },
    async handler(args) {
      const targetCell = validateExcelCellReference(args.targetCell, 'targetCell');
      const worksheetRef = resolveWorksheetRef(args.worksheet);
      return await runAppleScript(buildExcelPasteRangeScript(targetCell, worksheetRef));
    }
  },
  {
    name: 'excel_capture_range_ref',
    description: 'Legacy tool disabled by in-place editing policy. Use excel_copy_range instead.',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        range: { type: 'string', description: 'Range to capture (e.g., "A1:C5").' },
        worksheet: { type: ['string', 'integer'], description: 'Worksheet name or 1-based index. Defaults to the active sheet.' }
      },
      required: ['range']
    },
    async handler(args) {
      validateExcelRangeReference(args.range, 'range');
      resolveWorksheetRef(args.worksheet);
      throw new ToolError('NOT_SUPPORTED', 'excel_capture_range_ref is disabled by in-place editing policy. Use excel_copy_range.');
    }
  },
  {
    name: 'excel_insert_range_ref',
    description: 'Legacy tool disabled by in-place editing policy. Use excel_paste_range instead.',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        ref: { type: 'string', description: 'Opaque ref returned by excel_capture_range_ref.' },
        targetCell: { type: 'string', description: 'Top-left target cell in A1 notation.' },
        worksheet: { type: ['string', 'integer'], description: 'Worksheet name or 1-based index. Defaults to the active sheet.' }
      },
      required: ['ref', 'targetCell']
    },
    async handler(args) {
      validateString(args.ref, 'ref', true);
      validateExcelCellReference(args.targetCell, 'targetCell');
      resolveWorksheetRef(args.worksheet);
      throw new ToolError('NOT_SUPPORTED', 'excel_insert_range_ref is disabled by in-place editing policy. Use excel_paste_range.');
    }
  }
];
