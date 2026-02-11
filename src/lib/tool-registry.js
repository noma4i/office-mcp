import { documentTools } from '../tools/documents.js';
import { textTools } from '../tools/text.js';
import { tableTools } from '../tools/tables.js';
import { bookmarkTools } from '../tools/bookmarks.js';
import { hyperlinkTools } from '../tools/hyperlinks.js';
import { paragraphTools } from '../tools/paragraphs.js';
import { navigationTools } from '../tools/navigation.js';
import { imageTools } from '../tools/images.js';

export const ALL_TOOLS = [...documentTools, ...textTools, ...tableTools, ...bookmarkTools, ...hyperlinkTools, ...paragraphTools, ...navigationTools, ...imageTools];

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
