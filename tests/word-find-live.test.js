import { afterEach, describe, expect, jest, test } from '@jest/globals';

import { documentTools } from '../src/tools/word-documents.js';
import { navigationTools } from '../src/tools/word-navigation.js';
import { textTools } from '../src/tools/word-text.js';

const describeWordLive = process.env.WORD_LIVE_TESTS === '1' ? describe : describe.skip;

jest.setTimeout(120000);

function findTool(tools, name) {
  return tools.find(tool => tool.name === name);
}

function uniqueToken(label) {
  return `${label}_${Date.now()}_${Math.random().toString(36).slice(2, 10)}`;
}

const createDocument = findTool(documentTools, 'word_create_document').handler;
const getDocumentText = findTool(documentTools, 'word_get_document_text').handler;
const closeDocument = findTool(documentTools, 'word_close_document').handler;
const insertText = findTool(textTools, 'word_insert_text').handler;
const replaceText = findTool(textTools, 'word_replace_text').handler;
const deleteText = findTool(textTools, 'word_delete_text').handler;
const moveCursorAfterText = findTool(navigationTools, 'word_move_cursor_after_text').handler;

async function createDocumentWith(content) {
  const result = await createDocument({ content });
  expect(String(result)).toContain('New document created successfully');
}

async function readDocumentText() {
  return String(await getDocumentText({}));
}

describeWordLive('Word find live smoke', () => {
  afterEach(async () => {
    try {
      await closeDocument({ save: false });
    } catch {}
  });

  test('word_replace_text replaces short text in a real document', async () => {
    const marker = uniqueToken('short_find');
    const replacement = '[short_replaced]';

    await createDocumentWith(`prefix ${marker} suffix`);

    const result = await replaceText({ find: marker, replace: replacement });
    expect(String(result)).toContain('Text replaced successfully');

    const content = await readDocumentText();
    expect(content).toContain(replacement);
    expect(content).not.toContain(marker);
  });

  test('word_replace_text replaces long single-line text in a real document', async () => {
    const marker = uniqueToken('long_find');
    const longFind = `An incorrect automation configuration change ${marker} caused cascading availability degradation, delayed automation processing, unintended notifications, and incorrect module record updates across multiple customer flows.`;
    const replacement = `[describe_incident_${marker}]`;

    await createDocumentWith(`prefix ${longFind} suffix`);

    const result = await replaceText({ find: longFind, replace: replacement });
    expect(String(result)).toContain('Text replaced successfully');

    const content = await readDocumentText();
    expect(content).toContain(replacement);
    expect(content).not.toContain(longFind);
  });

  test('word_replace_text replaces placeholder text with unicode dash and quotes in a real document', async () => {
    const marker = uniqueToken('placeholder_root_cause');
    const placeholder = `[DESCRIBE THE ROOT CAUSE — WHAT TRIGGERED IT AND WHY ${marker}]`;
    const replacement = 'The API response included a "Link" header containing pagination URLs.';

    await createDocumentWith(`prefix ${placeholder} suffix`);

    const result = await replaceText({ find: placeholder, replace: replacement });
    expect(String(result)).toContain('Text replaced successfully');

    const content = await readDocumentText();
    expect(content).toContain(replacement);
    expect(content).not.toContain(placeholder);
  });

  test('word_replace_text replaces placeholder text with long API endpoint replacement in a real document', async () => {
    const marker = uniqueToken('placeholder_recovery');
    const placeholder = `[RECOVERY ACTION 3, e.g. Completed full audit and ${marker}]`;
    const replacement = 'Deployed alternative POST endpoint (POST /api/sub_form_completions/search) allowing filters to be passed in request body, eliminating URL length constraints.';

    await createDocumentWith(`prefix ${placeholder} suffix`);

    const result = await replaceText({ find: placeholder, replace: replacement });
    expect(String(result)).toContain('Text replaced successfully');

    const content = await readDocumentText();
    expect(content).toContain(replacement);
    expect(content).not.toContain(placeholder);
  });

  test('word_delete_text deletes long single-line text in a real document', async () => {
    const marker = uniqueToken('long_delete');
    const longFind = `All customers on server group ${marker} experienced a range of symptoms, from no visible impact to full login and automation disruption, with a smaller subset affected by incorrect data updates.`;

    await createDocumentWith(`before ${longFind} after`);

    const result = await deleteText({ text: longFind });
    expect(String(result)).toContain('Text deleted successfully');

    const content = await readDocumentText();
    expect(content).toContain('before');
    expect(content).toContain('after');
    expect(content).not.toContain(longFind);
  });

  test('word_move_cursor_after_text can advance after long text in a real document', async () => {
    const marker = uniqueToken('cursor_move');
    const longFind = `Investigation summary ${marker} identified a long single-line incident description that previously failed when passed through the Word find-object content setter path.`;
    const suffix = `[after_${marker}]`;
    const cursorMarker = `[cursor_${marker}]`;

    await createDocumentWith(`${longFind}${suffix}`);

    const moveResult = await moveCursorAfterText({ searchText: longFind });
    expect(String(moveResult)).toContain('Cursor moved after occurrence 1');

    const insertResult = await insertText({ text: cursorMarker });
    expect(String(insertResult)).toContain('Text inserted successfully');

    const content = await readDocumentText();
    expect(content).toContain(`${longFind}${cursorMarker}${suffix}`);
  });
});
