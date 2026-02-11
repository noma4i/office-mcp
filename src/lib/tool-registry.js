import { documentTools } from '../tools/documents.js';
import { textTools } from '../tools/text.js';
import { tableTools } from '../tools/tables.js';
import { bookmarkTools } from '../tools/bookmarks.js';
import { hyperlinkTools } from '../tools/hyperlinks.js';
import { paragraphTools } from '../tools/paragraphs.js';
import { navigationTools } from '../tools/navigation.js';
import { imageTools } from '../tools/images.js';
import { headerFooterTools } from '../tools/headers-footers.js';
import { sectionTools } from '../tools/sections.js';
import { formattingReadTools } from '../tools/formatting-read.js';
import { excelWorkbookTools } from '../tools/excel-workbooks.js';
import { excelSheetTools } from '../tools/excel-sheets.js';
import { excelCellTools } from '../tools/excel-cells.js';
import { excelFormattingTools } from '../tools/excel-formatting.js';
import { excelRowColumnTools } from '../tools/excel-rows-columns.js';
import { excelDataTools } from '../tools/excel-data.js';

export const ALL_TOOLS = [
  ...documentTools,
  ...textTools,
  ...tableTools,
  ...bookmarkTools,
  ...hyperlinkTools,
  ...paragraphTools,
  ...navigationTools,
  ...imageTools,
  ...headerFooterTools,
  ...sectionTools,
  ...formattingReadTools,
  ...excelWorkbookTools,
  ...excelSheetTools,
  ...excelCellTools,
  ...excelFormattingTools,
  ...excelRowColumnTools,
  ...excelDataTools
];

export function getToolDefinitions() {
  return ALL_TOOLS.map(tool => ({
    name: tool.name,
    description: tool.description,
    annotations: tool.annotations,
    inputSchema: tool.inputSchema
  }));
}

export function getToolHandler(toolName) {
  const tool = ALL_TOOLS.find(t => t.name === toolName);
  if (!tool) {
    throw new Error(`Unknown tool: ${toolName}`);
  }
  return tool.handler;
}
