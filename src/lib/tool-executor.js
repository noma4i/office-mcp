import { getErrorMessage } from "./validators.js";
import { inferErrorCode, isLikelyErrorMessage } from './errors.js';

export async function executeTool(toolName, args, handler) {
  try {
    const result = await handler(args);
    if (isLikelyErrorMessage(result)) {
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              ok: false,
              error: {
                code: inferErrorCode(result),
                message: result
              }
            })
          }
        ],
        isError: true
      };
    }

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            ok: true,
            message: typeof result === 'string' ? result : 'Operation completed successfully',
            data: typeof result === 'string' ? undefined : result
          })
        }
      ]
    };
  } catch (error) {
    const message = getErrorMessage(error);
    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            ok: false,
            error: {
              code: inferErrorCode(message),
              message: `Failed to ${toolName}: ${message}`
            }
          })
        }
      ],
      isError: true
    };
  }
}
