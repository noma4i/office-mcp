import { documentTools } from '../tools/word-documents.js';
import { textTools } from '../tools/word-text.js';
import { tableTools } from '../tools/word-tables.js';
import { bookmarkTools } from '../tools/word-bookmarks.js';
import { hyperlinkTools } from '../tools/word-hyperlinks.js';
import { paragraphTools } from '../tools/word-paragraphs.js';
import { navigationTools } from '../tools/word-navigation.js';
import { imageTools } from '../tools/word-images.js';
import { clipboardTools } from '../tools/word-clipboard.js';
import { wordWorkflowTools } from '../tools/word-workflows.js';
import { headerFooterTools } from '../tools/word-headers-footers.js';
import { sectionTools } from '../tools/word-sections.js';
import { formattingReadTools } from '../tools/word-formatting-read.js';
import { excelWorkbookTools } from '../tools/excel-workbooks.js';
import { excelSheetTools } from '../tools/excel-sheets.js';
import { excelCellTools } from '../tools/excel-cells.js';
import { excelFormattingTools } from '../tools/excel-formatting.js';
import { excelRowColumnTools } from '../tools/excel-rows-columns.js';
import { excelDataTools } from '../tools/excel-data.js';
import { excelClipboardTools } from '../tools/excel-clipboard.js';
import { excelWorkflowTools } from '../tools/excel-workflows.js';

export const ALL_TOOLS = [
  ...documentTools,
  ...textTools,
  ...tableTools,
  ...bookmarkTools,
  ...hyperlinkTools,
  ...paragraphTools,
  ...navigationTools,
  ...imageTools,
  ...clipboardTools,
  ...wordWorkflowTools,
  ...headerFooterTools,
  ...sectionTools,
  ...formattingReadTools,
  ...excelWorkbookTools,
  ...excelSheetTools,
  ...excelCellTools,
  ...excelFormattingTools,
  ...excelRowColumnTools,
  ...excelDataTools,
  ...excelClipboardTools,
  ...excelWorkflowTools
];

export function buildToolMap(tools) {
  const map = new Map();
  for (const tool of tools) {
    if (map.has(tool.name)) {
      throw new Error(`Duplicate tool registration: ${tool.name}`);
    }
    map.set(tool.name, tool);
  }
  return map;
}

const TOOL_MAP = buildToolMap(ALL_TOOLS);

export function getToolDefinitions() {
  return ALL_TOOLS.map(tool => ({
    name: tool.name,
    description: tool.description,
    annotations: tool.annotations,
    inputSchema: tool.inputSchema
  }));
}

export function getToolHandler(toolName) {
  const tool = TOOL_MAP.get(toolName);
  if (!tool) {
    throw new Error(`Unknown tool: ${toolName}`);
  }
  return tool.handler;
}
