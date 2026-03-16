import { describe, expect, test } from '@jest/globals';

import { WORD_FIND_MODES, runWordFindWithFallback } from '../src/lib/applescript/word-find.js';

function compatibilityError() {
  return new Error('AppleScript error: Microsoft Word got an error: find id 1 of selection doesn’t understand the “execute find” message. (-1708)');
}

describe('Word find orchestration', () => {
  test('returns direct strategy result without fallback', async () => {
    const calls = [];
    const executeAppleScript = async script => {
      calls.push(script);
      return 'Text replaced successfully';
    };

    const result = await runWordFindWithFallback(
      {
        mode: WORD_FIND_MODES.REPLACE,
        findText: 'old',
        replaceWith: 'new'
      },
      { executeAppleScript }
    );

    expect(result).toBe('Text replaced successfully');
    expect(calls).toHaveLength(1);
    expect(calls[0]).toContain('execute find findObject');
    expect(calls[0]).not.toContain('set content of findObject');
  });

  test('retries with legacy strategy on execute-find compatibility error', async () => {
    const calls = [];
    const executeAppleScript = async script => {
      calls.push(script);
      if (calls.length === 1) {
        throw compatibilityError();
      }
      return 'Text replaced successfully';
    };

    const result = await runWordFindWithFallback(
      {
        mode: WORD_FIND_MODES.REPLACE,
        findText: 'old',
        replaceWith: 'new'
      },
      { executeAppleScript }
    );

    expect(result).toBe('Text replaced successfully');
    expect(calls).toHaveLength(2);
    expect(calls[0]).toContain('execute find findObject');
    expect(calls[0]).not.toContain('set content of findObject');
    expect(calls[1]).toContain('set content of findObject to "old"');
    expect(calls[1]).toContain('set content of replacement of findObject to "new"');
  });

  test('returns fallback not-found result after compatibility retry', async () => {
    const calls = [];
    const executeAppleScript = async script => {
      calls.push(script);
      if (calls.length === 1) {
        throw compatibilityError();
      }
      return 'Text not found, nothing deleted';
    };

    const result = await runWordFindWithFallback(
      {
        mode: WORD_FIND_MODES.DELETE_ALL,
        findText: 'missing'
      },
      { executeAppleScript }
    );

    expect(result).toBe('Text not found, nothing deleted');
    expect(calls).toHaveLength(2);
    expect(calls[1]).toContain('set content of findObject to "missing"');
    expect(calls[1]).toContain('set content of replacement of findObject to ""');
  });

  test('does not fallback on non-compatibility error', async () => {
    const calls = [];
    const executeAppleScript = async script => {
      calls.push(script);
      throw new Error('AppleScript error: No document is open');
    };

    await expect(
      runWordFindWithFallback(
        {
          mode: WORD_FIND_MODES.REPLACE,
          findText: 'old',
          replaceWith: 'new'
        },
        { executeAppleScript }
      )
    ).rejects.toThrow('No document is open');

    expect(calls).toHaveLength(1);
  });

  test('uses the same fallback flow for move_cursor_after_text', async () => {
    const calls = [];
    const executeAppleScript = async script => {
      calls.push(script);
      if (calls.length === 1) {
        throw compatibilityError();
      }
      return 'Cursor moved after occurrence 2 of: "needle"';
    };

    const result = await runWordFindWithFallback(
      {
        mode: WORD_FIND_MODES.MOVE_CURSOR_AFTER_TEXT,
        findText: 'needle',
        occurrence: 2
      },
      { executeAppleScript }
    );

    expect(result).toBe('Cursor moved after occurrence 2 of: "needle"');
    expect(calls).toHaveLength(2);
    expect(calls[0]).toContain('execute find findObject find text "needle"');
    expect(calls[1]).toContain('set content of findObject to "needle"');
    expect(calls[1]).toContain('set wrap of findObject to find stop');
    expect(calls[1]).toContain('set forward of findObject to true');
  });

  test('combines direct and fallback errors when both strategies fail', async () => {
    const calls = [];
    const executeAppleScript = async script => {
      calls.push(script);
      if (calls.length === 1) {
        throw compatibilityError();
      }
      throw new Error('AppleScript error: legacy strategy also failed');
    };

    await expect(
      runWordFindWithFallback(
        {
          mode: WORD_FIND_MODES.REPLACE,
          findText: 'old',
          replaceWith: 'new'
        },
        { executeAppleScript }
      )
    ).rejects.toThrow('Word find failed for both strategies.');

    expect(calls).toHaveLength(2);
  });
});
