import { ToolError } from '../lib/errors.js';
import { commitReservedFragment, discardReservedFragment, getFragment, reserveFragment } from '../lib/fragment-store.js';
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

function buildExcelCaptureRangeRefScript(range, worksheetRef, fragmentPath) {
  return `
tell application "Microsoft Excel"
  if (count of workbooks) = 0 then
    return "No workbook is open"
  end if
  activate
  set sourceWb to active workbook
  set sourceWs to ${worksheetRef}
  try
    select range ${JSON.stringify(range)} of sourceWs
  on error
    return "Invalid range: ${range}"
  end try
end tell
delay 0.2
tell application "System Events"
  keystroke "c" using command down
end tell
delay 0.2
tell application "Microsoft Excel"
  activate
  set fragWb to make new workbook
  set fragWs to active sheet
  select range "A1" of fragWs
end tell
delay 0.2
tell application "System Events"
  keystroke "v" using command down
end tell
delay 0.2
tell application "Microsoft Excel"
  try
    save workbook as fragWb filename ${JSON.stringify(fragmentPath)}
    close fragWb saving no
  on error errMsg
    try
      close fragWb saving no
    end try
    return "Error creating range fragment: " & errMsg
  end try
end tell
return "Range captured to ref"
`;
}

function buildExcelInsertRangeRefScript(fragmentPath, targetCell, worksheetRef) {
  return `
tell application "Microsoft Excel"
  if (count of workbooks) = 0 then
    return "No workbook is open"
  end if
  activate
  set targetWb to active workbook
  set targetWs to ${worksheetRef}
end tell
tell application "Microsoft Excel"
  open workbook workbook file name ${JSON.stringify(fragmentPath)}
  set fragWb to active workbook
  set fragWs to active sheet
  set fragRange to used range of fragWs
  select fragRange
end tell
delay 0.2
tell application "System Events"
  keystroke "c" using command down
end tell
delay 0.2
tell application "Microsoft Excel"
  close fragWb saving no
  activate object targetWs
  try
    select range ${JSON.stringify(targetCell)} of targetWs
  on error
    return "Cell ${targetCell} not accessible"
  end try
end tell
delay 0.2
tell application "System Events"
  keystroke "v" using command down
end tell
return "Range inserted from ref"
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
    description: 'Capture an Excel range as a reusable ref backed by a temporary workbook.',
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
      const range = validateExcelRangeReference(args.range, 'range');
      const worksheetRef = resolveWorksheetRef(args.worksheet);
      const reserved = reserveFragment({
        prefix: 'excelfrag',
        app: 'excel',
        kind: 'excel_range',
        extension: 'xlsx',
        summary: { label: `range ${range}`, worksheet: args.worksheet ?? 'active' }
      });

      try {
        const result = await runAppleScript(buildExcelCaptureRangeRefScript(range, worksheetRef, reserved.filePath));
        if (result !== 'Range captured to ref') {
          throw new ToolError('OPERATION_ERROR', result);
        }
        return commitReservedFragment(reserved);
      } catch (error) {
        discardReservedFragment(reserved);
        throw error;
      }
    }
  },
  {
    name: 'excel_insert_range_ref',
    description: 'Insert a previously captured Excel range ref into the active workbook at the target cell.',
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
      const ref = validateString(args.ref, 'ref', true);
      const targetCell = validateExcelCellReference(args.targetCell, 'targetCell');
      const worksheetRef = resolveWorksheetRef(args.worksheet);
      const fragment = getFragment(ref, 'excel');
      if (fragment.kind !== 'excel_range') {
        throw new ToolError('VALIDATION_ERROR', `ref kind is not supported in Excel: ${fragment.kind}`);
      }

      const result = await runAppleScript(buildExcelInsertRangeRefScript(fragment.filePath, targetCell, worksheetRef));
      if (result !== 'Range inserted from ref') {
        throw new ToolError('OPERATION_ERROR', result);
      }
      return { inserted: true, ref: fragment.ref, kind: fragment.kind, targetCell };
    }
  }
];
