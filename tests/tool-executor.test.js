import { describe, expect, test } from '@jest/globals';
import { executeTool } from '../src/lib/tool-executor.js';
import { ToolError, inferErrorCode, isLikelyErrorMessage } from '../src/lib/errors.js';

function parseResult(response) {
  return JSON.parse(response.content[0].text);
}

describe('Tool Executor', () => {
  test('classifies known error messages', () => {
    expect(isLikelyErrorMessage('No document is open')).toBe(true);
    expect(isLikelyErrorMessage('Text not found')).toBe(true);
    expect(isLikelyErrorMessage('Operation completed')).toBe(false);
  });

  test('infers stable error codes', () => {
    expect(inferErrorCode('No document is open')).toBe('NO_DOCUMENT_OPEN');
    expect(inferErrorCode('No workbook is open')).toBe('NO_WORKBOOK_OPEN');
    expect(inferErrorCode('AppleScript error: timeout')).toBe('APPSCRIPT_ERROR');
  });

  test('returns ok payload for success', async () => {
    const response = await executeTool('demo_tool', {}, async () => 'Done');
    const payload = parseResult(response);
    expect(payload.ok).toBe(true);
    expect(payload.message).toBe('Done');
  });

  test('returns error payload when handler throws', async () => {
    const response = await executeTool('demo_tool', {}, async () => {
      throw new Error('find is required');
    });
    const payload = parseResult(response);
    expect(response.isError).toBe(true);
    expect(payload.ok).toBe(false);
    expect(payload.error.code).toBe('VALIDATION_ERROR');
  });

  test('keeps success payload for arbitrary string content', async () => {
    const response = await executeTool('demo_tool', {}, async () => 'This document says: Text not found...');
    const payload = parseResult(response);
    expect(response.isError).toBeUndefined();
    expect(payload.ok).toBe(true);
    expect(payload.message).toBe('This document says: Text not found...');
  });

  test('serializes ToolError code and details', async () => {
    const response = await executeTool('demo_tool', {}, async () => {
      throw new ToolError('NOT_FOUND', 'Bookmark not found', { bookmark: 'intro' });
    });
    const payload = parseResult(response);
    expect(response.isError).toBe(true);
    expect(payload.ok).toBe(false);
    expect(payload.error.code).toBe('NOT_FOUND');
    expect(payload.error.details).toEqual({ bookmark: 'intro' });
  });
});
