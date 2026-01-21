import { getErrorMessage } from "./validators.js";

export async function executeTool(toolName, args, handler) {
  try {
    const result = await handler(args);
    return {
      content: [{ type: "text", text: result }]
    };
  } catch (error) {
    return {
      content: [{ type: "text", text: `Failed to ${toolName}: ${getErrorMessage(error)}` }],
      isError: true
    };
  }
}
