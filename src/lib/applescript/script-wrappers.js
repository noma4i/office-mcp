function indentBlock(block, spaces = 2) {
  return block
    .split('\n')
    .map(line => (line.length === 0 ? line : `${' '.repeat(spaces)}${line}`))
    .join('\n');
}

export function wrapWordScript(body, { activate = false, requireDocument = true } = {}) {
  const lines = ['tell application "Microsoft Word"'];
  if (activate) {
    lines.push('  activate');
  }
  if (requireDocument) {
    lines.push('  if (count of documents) = 0 then');
    lines.push('    return "No document is open"');
    lines.push('  end if');
  }
  lines.push(indentBlock(body.trim(), 2));
  lines.push('end tell');
  return `\n${lines.join('\n')}\n`;
}

export function wrapExcelScript(body, { requireWorkbook = true, setActiveWorkbook = false, setActiveSheet = false } = {}) {
  const lines = ['tell application "Microsoft Excel"'];
  if (requireWorkbook) {
    lines.push('  if (count of workbooks) = 0 then');
    lines.push('    return "No workbook is open"');
    lines.push('  end if');
  }
  if (setActiveWorkbook) {
    lines.push('  set wb to active workbook');
  }
  if (setActiveSheet) {
    lines.push('  set ws to active sheet');
  }
  lines.push(indentBlock(body.trim(), 2));
  lines.push('end tell');
  return `\n${lines.join('\n')}\n`;
}

