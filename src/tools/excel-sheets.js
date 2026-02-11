import { validateString, validateInteger } from '../lib/validators.js';
import { runAppleScript } from '../lib/applescript/executor.js';

function resolveSheet(nameOrIndex) {
  if (typeof nameOrIndex === 'number' || (typeof nameOrIndex === 'string' && /^\d+$/.test(nameOrIndex))) {
    const idx = parseInt(nameOrIndex, 10);
    return `worksheet ${idx} of wb`;
  }
  return `worksheet ${JSON.stringify(nameOrIndex)} of wb`;
}

export const excelSheetTools = [
  {
    name: 'excel_list_sheets',
    description: 'List all worksheets in the active Excel workbook',
    annotations: { readOnlyHint: true },
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
          set wb to active workbook
          set sheetCount to count of worksheets of wb
          if sheetCount = 0 then
            return "No worksheets found"
          end if
          set sheetList to ""
          repeat with i from 1 to sheetCount
            set ws to worksheet i of wb
            set wsName to name of ws
            set sheetList to sheetList & i & ". " & wsName & linefeed
          end repeat
          return sheetList
        end tell
      `;
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_create_sheet',
    description: 'Create a new worksheet in the active Excel workbook',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        name: {
          type: 'string',
          description: 'Name for the new sheet'
        },
        afterIndex: {
          type: 'integer',
          description: 'Insert after this sheet index (if not provided, adds at end)'
        }
      }
    },
    async handler(args) {
      const name = validateString(args.name, 'name', false);
      const afterIndex = args.afterIndex !== undefined ? validateInteger(args.afterIndex, 'afterIndex', 1) : undefined;

      let makeCmd = 'make new worksheet at end of wb';
      if (afterIndex !== undefined) {
        makeCmd = `make new worksheet at after worksheet ${afterIndex} of wb`;
      }

      const nameCmd = name ? `\n          set name of newSheet to ${JSON.stringify(name)}` : '';

      const script = `
        tell application "Microsoft Excel"
          if (count of workbooks) = 0 then
            return "No workbook is open"
          end if
          set wb to active workbook
          set newSheet to ${makeCmd}${nameCmd}
          return "Sheet created: " & name of newSheet
        end tell
      `;
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_delete_sheet',
    description: 'Delete a worksheet by name or index in Excel',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        nameOrIndex: {
          type: ['string', 'integer'],
          description: 'Sheet name (string) or index (integer, 1-based)'
        }
      },
      required: ['nameOrIndex']
    },
    async handler(args) {
      if (args.nameOrIndex === undefined || args.nameOrIndex === null) {
        throw new Error('nameOrIndex is required');
      }
      const sheetRef = resolveSheet(args.nameOrIndex);
      const script = `
        tell application "Microsoft Excel"
          if (count of workbooks) = 0 then
            return "No workbook is open"
          end if
          set wb to active workbook
          set display alerts to false
          delete ${sheetRef}
          set display alerts to true
          return "Sheet deleted successfully"
        end tell
      `;
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_rename_sheet',
    description: 'Rename a worksheet in the active Excel workbook',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        nameOrIndex: {
          type: ['string', 'integer'],
          description: 'Sheet name (string) or index (integer, 1-based)'
        },
        newName: {
          type: 'string',
          description: 'New name for the sheet'
        }
      },
      required: ['nameOrIndex', 'newName']
    },
    async handler(args) {
      if (args.nameOrIndex === undefined || args.nameOrIndex === null) {
        throw new Error('nameOrIndex is required');
      }
      const newName = validateString(args.newName, 'newName', true);
      const sheetRef = resolveSheet(args.nameOrIndex);
      const script = `
        tell application "Microsoft Excel"
          if (count of workbooks) = 0 then
            return "No workbook is open"
          end if
          set wb to active workbook
          set ws to ${sheetRef}
          set name of ws to ${JSON.stringify(newName)}
          return "Sheet renamed to " & ${JSON.stringify(newName)}
        end tell
      `;
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_activate_sheet',
    description: 'Activate (switch to) a worksheet by name or index in Excel',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        nameOrIndex: {
          type: ['string', 'integer'],
          description: 'Sheet name (string) or index (integer, 1-based)'
        }
      },
      required: ['nameOrIndex']
    },
    async handler(args) {
      if (args.nameOrIndex === undefined || args.nameOrIndex === null) {
        throw new Error('nameOrIndex is required');
      }
      const sheetRef = resolveSheet(args.nameOrIndex);
      const script = `
        tell application "Microsoft Excel"
          if (count of workbooks) = 0 then
            return "No workbook is open"
          end if
          set wb to active workbook
          set ws to ${sheetRef}
          activate object ws
          return "Activated sheet: " & name of ws
        end tell
      `;
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_get_sheet_info',
    description: 'Get worksheet info (used range address, row count, column count) in Excel',
    annotations: { readOnlyHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        nameOrIndex: {
          type: ['string', 'integer'],
          description: 'Sheet name or index (default: active sheet)'
        }
      }
    },
    async handler(args) {
      const sheetCmd = (args.nameOrIndex !== undefined && args.nameOrIndex !== null)
        ? `set ws to ${resolveSheet(args.nameOrIndex)}`
        : 'set ws to active sheet';
      const script = `
        tell application "Microsoft Excel"
          if (count of workbooks) = 0 then
            return "No workbook is open"
          end if
          set wb to active workbook
          ${sheetCmd}
          set wsName to name of ws
          set ur to used range of ws
          set addr to get address of ur
          set rc to count of rows of ur
          set cc to count of columns of ur
          return "Sheet: " & wsName & linefeed & "Used range: " & addr & linefeed & "Rows: " & rc & linefeed & "Columns: " & cc
        end tell
      `;
      return await runAppleScript(script);
    }
  }
];
