export const COMMON_SCRIPTS = {
  checkDocumentOpen: `
if (count of documents) = 0 then
  return "No document is open"
end if`,

  getActiveDocument: `set activeDoc to active document`,

  cleanCellMarkers: `
if length of cellText > 0 then
  repeat while (length of cellText > 0) and ((ASCII number of (character -1 of cellText)) is in {7, 13})
    set cellText to text 1 thru -2 of cellText
  end repeat
end if`,

  getTable: tableIndex => `
set tableCount to count of tables of activeDoc
if ${tableIndex} > tableCount then
  return "Table index out of range. Document has " & tableCount & " tables."
end if
set t to table ${tableIndex} of activeDoc`,

  collapseToStart: `set selection end of selection to selection start of selection`,

  collapseToEnd: `set selection start of selection to selection end of selection`,

  checkWorkbookOpen: `
if (count of workbooks) = 0 then
  return "No workbook is open"
end if`,

  getActiveWorkbook: `set wb to active workbook`,

  getActiveSheet: `set ws to active sheet`,

  getSheetByNameOrIndex: (nameOrIndex, isString) => (isString ? `set ws to worksheet ${nameOrIndex} of wb` : `set ws to worksheet ${nameOrIndex} of wb`)
};

export function escapeAppleScriptString(str) {
  return str.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
}

export function quoteAppleScriptString(str) {
  return `"${escapeAppleScriptString(str)}"`;
}
