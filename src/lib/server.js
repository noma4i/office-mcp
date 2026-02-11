import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { CallToolRequestSchema, ListToolsRequestSchema } from '@modelcontextprotocol/sdk/types.js';

import { getToolDefinitions, getToolHandler } from './tool-registry.js';
import { executeTool } from './tool-executor.js';

export function createServer() {
  const server = new Server(
    {
      name: 'Microsoft-Office-Server',
      version: '0.8.0'
    },
    {
      capabilities: {
        tools: {}
      }
    }
  );

  server.setRequestHandler(ListToolsRequestSchema, async () => {
    return {
      tools: getToolDefinitions()
    };
  });

  server.setRequestHandler(CallToolRequestSchema, async request => {
    const toolName = request.params.name;
    const args = request.params.arguments || {};

    const handler = getToolHandler(toolName);
    return await executeTool(toolName, args, async () => {
      return await handler(args);
    });
  });

  return server;
}

export async function startServer() {
  const server = createServer();
  const transport = new StdioServerTransport();
  await server.connect(transport);
}
