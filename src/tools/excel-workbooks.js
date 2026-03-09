import { validateString, validateBoolean } from '../lib/validators.js';
import { runAppleScript } from '../lib/applescript/executor.js';
import { wrapExcelScript } from '../lib/applescript/script-wrappers.js';

export const excelWorkbookTools = [
  {
    name: 'excel_create_workbook',
    description: 'Create a new workbook in Microsoft Excel',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {}
    },
    async handler() {
      const script = wrapExcelScript(
        `
activate
make new workbook
return "New workbook created successfully"
`,
        { requireWorkbook: false }
      );
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_open_workbook',
    description: 'Open an existing workbook in Microsoft Excel',
    annotations: { readOnlyHint: false },
    inputSchema: {
      type: 'object',
      properties: {
        path: {
          type: 'string',
          description: 'Full path to the workbook file'
        }
      },
      required: ['path']
    },
    async handler(args) {
      const path = validateString(args.path, 'path', true);
      const script = wrapExcelScript(
        `
activate
open workbook workbook file name ${JSON.stringify(path)}
return "Workbook opened successfully"
`,
        { requireWorkbook: false }
      );
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_get_workbook_info',
    description: 'Get workbook information (name, path, sheet count) in Microsoft Excel',
    annotations: { readOnlyHint: true },
    inputSchema: {
      type: 'object',
      properties: {}
    },
    async handler() {
      const script = wrapExcelScript(
        `
set wb to active workbook
set wbName to name of wb
set wbPath to full name of wb
set sheetCount to count of worksheets of wb
return "Name: " & wbName & linefeed & "Path: " & wbPath & linefeed & "Sheets: " & sheetCount
`,
        { setActiveWorkbook: false }
      );
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_save_workbook',
    description: 'Save the active workbook in Microsoft Excel',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        path: {
          type: 'string',
          description: 'Optional path to save as (if not provided, saves to current location)'
        }
      }
    },
    async handler(args) {
      const path = args.path !== undefined ? validateString(args.path, 'path', true) : undefined;
      const script = path
        ? wrapExcelScript(`
set wb to active workbook
set display alerts to false
try
  save workbook as wb filename ${JSON.stringify(path)}
on error errMsg
  set display alerts to true
  error errMsg
end try
set display alerts to true
return "Workbook saved as " & ${JSON.stringify(path)}
`)
        : wrapExcelScript(`
set wb to active workbook
save wb
return "Workbook saved successfully"
`);
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_close_workbook',
    description: 'Close the active workbook in Microsoft Excel',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        save: {
          type: 'boolean',
          description: 'Save before closing (default: true)',
          default: true
        }
      }
    },
    async handler(args) {
      const save = validateBoolean(args.save, 'save', true);
      const script = wrapExcelScript(`
set wb to active workbook
close wb ${save ? 'saving yes' : 'saving no'}
return "Workbook closed successfully"
`);
      return await runAppleScript(script);
    }
  },

  {
    name: 'excel_list_workbooks',
    description: 'List all open workbooks in Microsoft Excel',
    annotations: { readOnlyHint: true },
    inputSchema: {
      type: 'object',
      properties: {}
    },
    async handler() {
      const script = wrapExcelScript(
        `
set wbCount to count of workbooks
if wbCount = 0 then
  return "No workbooks are open"
end if
set wbList to ""
repeat with i from 1 to wbCount
  set wb to workbook i
  set wbName to name of wb
  set sheetCount to count of worksheets of wb
  set wbList to wbList & i & ". " & wbName & " (" & sheetCount & " sheets)" & linefeed
end repeat
return wbList
`,
        { requireWorkbook: false }
      );
      return await runAppleScript(script);
    }
  }
];
