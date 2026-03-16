export const COMMON_SCRIPTS = {
  checkDocumentOpen: `
if (count of documents) = 0 then
  return "No document is open"
end if`,

  getActiveDocument: `set activeDoc to active document`,

  cleanCellMarkers: `
if length of cellText > 0 then
  repeat while length of cellText > 0
    set lastCharCode to ASCII number of (character -1 of cellText)
    if lastCharCode is not in {7, 13} then
      exit repeat
    end if
    if length of cellText is 1 then
      set cellText to ""
    else
      set cellText to text 1 thru -2 of cellText
    end if
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

  getSheetByNameOrIndex: (nameOrIndex, isString) => (isString ? `set ws to worksheet "${escapeAppleScriptString(nameOrIndex)}" of wb` : `set ws to worksheet ${nameOrIndex} of wb`)
};

export function escapeAppleScriptString(str) {
  return str.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
}

export function quoteAppleScriptString(str) {
  return `"${escapeAppleScriptString(str)}"`;
}

export function toAppleScriptString(str) {
  const parts = str.split(/\r\n|\n|\r/);
  const escaped = parts.map(p => `"${escapeAppleScriptString(p)}"`);
  return escaped.length === 1 ? escaped[0] : '(' + escaped.join(' & return & ') + ')';
}

export function escapeForWordFind(str) {
  const escaped = escapeAppleScriptString(str).replace(/\r\n/g, '^p').replace(/\n/g, '^p').replace(/\r/g, '^p');
  return `"${escaped}"`;
}

export function buildWordExecuteFind(findObjectName, {
  findText,
  replaceWith,
  replace,
  matchForward,
  wrapFind
} = {}) {
  const parts = [`execute find ${findObjectName}`];

  if (findText !== undefined) {
    parts.push(`find text ${escapeForWordFind(findText)}`);
  }
  if (matchForward !== undefined) {
    parts.push(`match forward ${matchForward}`);
  }
  if (wrapFind !== undefined) {
    parts.push(`wrap find ${wrapFind}`);
  }
  if (replaceWith !== undefined) {
    parts.push(`replace with ${escapeForWordFind(replaceWith)}`);
  }
  if (replace !== undefined) {
    parts.push(`replace ${replace}`);
  }

  return parts.join(' ');
}

export function wrapExcelRowValues(values) {
  return `{${values.map(value => quoteAppleScriptString(String(value))).join(', ')}}`;
}
