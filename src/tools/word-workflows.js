import { validateEnum, validateInteger, validateString } from '../lib/validators.js';
import { runAppleScript } from '../lib/applescript/executor.js';
import { toAppleScriptString } from '../lib/applescript/helpers.js';
import { wrapWordScript } from '../lib/applescript/script-wrappers.js';

function headerFooterIndex(type) {
  switch (type) {
    case 'first_page':
      return 'header footer first page';
    case 'even_pages':
      return 'header footer even pages';
    default:
      return 'header footer primary';
  }
}

function resolveStoryArgs(args, { requireText = false } = {}) {
  const scope = validateEnum(args.scope, 'scope', ['body', 'header', 'footer'], 'body');
  const text = requireText ? validateString(args.text, 'text', true) : undefined;
  if (scope === 'body') {
    if (args.section !== undefined || args.type !== undefined) {
      throw new Error('section/type cannot be used with body scope');
    }
    return text === undefined ? { scope } : { scope, text };
  }

  const section = validateInteger(args.section, 'section', 1) || 1;
  const type = validateEnum(args.type, 'type', ['primary', 'first_page', 'even_pages'], 'primary');
  const resolved = { scope, section, type, hfIndex: headerFooterIndex(type) };
  if (text !== undefined) resolved.text = text;

  return resolved;
}

function buildWordStoryRangeScript(target) {
  if (target.scope === 'body') {
    return `
set d to active document
set storyRange to text object of d`;
  }

  const storyKind = target.scope === 'header' ? 'header' : 'footer';
  return `
set d to active document
set secCount to count of sections of d
if ${target.section} > secCount then
  return "Section index out of range. Document has " & secCount & " sections."
end if
try
  set storyRef to get ${storyKind} of section ${target.section} of d index ${target.hfIndex}
  set storyRange to text object of storyRef
on error
  return "${storyKind === 'header' ? 'Header' : 'Footer'} not available for this section/type"
end try`;
}

export const wordWorkflowTools = [
  {
    name: 'word_copy_story_content',
    description: 'Copy body, header, or footer content from the active Word document to the system clipboard.',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        scope: { type: 'string', enum: ['body', 'header', 'footer'], description: 'Story to copy (default: body).' },
        section: { type: 'integer', description: 'Section number for header/footer scopes (default: 1).' },
        type: {
          type: 'string',
          enum: ['primary', 'first_page', 'even_pages'],
          description: 'Header/footer type for header/footer scopes (default: primary).'
        }
      }
    },
    async handler(args) {
      const target = resolveStoryArgs(args);
      const script = wrapWordScript(`
${buildWordStoryRangeScript(target)}
select storyRange
try
  copy object selection
on error errMsg
  return "Error copying story content: " & errMsg
end try
return "Story content copied to clipboard"
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'word_clear_story_content',
    description: 'Clear body, header, or footer content in the active Word document without creating temporary documents.',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        scope: { type: 'string', enum: ['body', 'header', 'footer'], description: 'Story to clear (default: body).' },
        section: { type: 'integer', description: 'Section number for header/footer scopes (default: 1).' },
        type: {
          type: 'string',
          enum: ['primary', 'first_page', 'even_pages'],
          description: 'Header/footer type for header/footer scopes (default: primary).'
        }
      }
    },
    async handler(args) {
      const target = resolveStoryArgs(args);
      const script = wrapWordScript(`
${buildWordStoryRangeScript(target)}
set content of storyRange to ""
return "Story content cleared"
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'word_set_story_text',
    description: 'Replace body, header, or footer text in the active Word document in place.',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        scope: { type: 'string', enum: ['body', 'header', 'footer'], description: 'Story to replace (default: body).' },
        text: { type: 'string', description: 'Text to write into the target story.' },
        section: { type: 'integer', description: 'Section number for header/footer scopes (default: 1).' },
        type: {
          type: 'string',
          enum: ['primary', 'first_page', 'even_pages'],
          description: 'Header/footer type for header/footer scopes (default: primary).'
        }
      },
      required: ['text']
    },
    async handler(args) {
      const target = resolveStoryArgs(args, { requireText: true });
      const script = wrapWordScript(`
${buildWordStoryRangeScript(target)}
set content of storyRange to ${toAppleScriptString(target.text)}
return "Story text replaced"
`);
      return await runAppleScript(script);
    }
  }
];
