import { ToolError } from '../lib/errors.js';
import { buildSourceSummary, getFragment } from '../lib/fragment-store.js';
import { runAppleScript } from '../lib/applescript/executor.js';
import { wrapWordScript } from '../lib/applescript/script-wrappers.js';
import { validateEnum, validateInteger, validateNumber } from '../lib/validators.js';
import { buildWordInsertImageScript } from './word-images.js';

function resolveCopySelection(args) {
  const explicitScope = validateEnum(args.scope, 'scope', ['selection', 'document', 'paragraph_range', 'inline_shape']);
  const hasParagraphRange = args.startParagraph !== undefined || args.endParagraph !== undefined;
  const hasInlineShape = args.inlineShapeIndex !== undefined;

  let scope = explicitScope;
  if (!scope) {
    if (hasParagraphRange) {
      scope = 'paragraph_range';
    } else if (hasInlineShape) {
      scope = 'inline_shape';
    } else {
      scope = 'selection';
    }
  }

  if (scope === 'paragraph_range') {
    if (args.startParagraph === undefined) {
      throw new Error('startParagraph is required for paragraph_range');
    }
    if (hasInlineShape) {
      throw new Error('inlineShapeIndex cannot be used with paragraph_range');
    }
    const startParagraph = validateInteger(args.startParagraph, 'startParagraph', 1);
    const endParagraph = args.endParagraph !== undefined ? validateInteger(args.endParagraph, 'endParagraph', 1) : startParagraph;
    if (endParagraph < startParagraph) {
      throw new Error('endParagraph must be >= startParagraph');
    }
    return {
      scope,
      startParagraph,
      endParagraph,
      summary: buildSourceSummary(`paragraphs ${startParagraph}-${endParagraph}`, { scope })
    };
  }

  if (scope === 'inline_shape') {
    if (args.inlineShapeIndex === undefined) {
      throw new Error('inlineShapeIndex is required for inline_shape');
    }
    if (hasParagraphRange) {
      throw new Error('startParagraph/endParagraph cannot be used with inline_shape');
    }
    const inlineShapeIndex = validateInteger(args.inlineShapeIndex, 'inlineShapeIndex', 1);
    return {
      scope,
      inlineShapeIndex,
      summary: buildSourceSummary(`inline shape ${inlineShapeIndex}`, { scope })
    };
  }

  if (hasParagraphRange || hasInlineShape) {
    throw new Error(`startParagraph/endParagraph/inlineShapeIndex cannot be used with ${scope}`);
  }

  return {
    scope,
    summary: buildSourceSummary(scope === 'document' ? 'full document' : 'current selection', { scope })
  };
}

function buildWordSelectionScript(selection) {
  if (selection.scope === 'document') {
    return `
set d to active document
select (text object of d)`;
  }

  if (selection.scope === 'paragraph_range') {
    return `
set d to active document
set paraCount to count of paragraphs of d
if ${selection.startParagraph} > paraCount then
  return "Start paragraph out of range. Document has " & paraCount & " paragraphs."
end if
if ${selection.endParagraph} > paraCount then
  return "End paragraph out of range. Document has " & paraCount & " paragraphs."
end if
select (text object of paragraph ${selection.startParagraph} of d)
set rStart to selection start of selection
select (text object of paragraph ${selection.endParagraph} of d)
set rEnd to selection end of selection
set selection start of selection to rStart
set selection end of selection to rEnd`;
  }

  if (selection.scope === 'inline_shape') {
    return `
set d to active document
set shapeCount to count of inline shapes of d
if ${selection.inlineShapeIndex} > shapeCount then
  return "Shape index out of range. Document has " & shapeCount & " inline shapes."
end if
try
  select (inline shape ${selection.inlineShapeIndex} of d)
on error errMsg
  return "Error selecting inline shape ${selection.inlineShapeIndex}: " & errMsg
end try`;
  }

  return '';
}

function buildWordCopyScript(selection) {
  const selectionScript = buildWordSelectionScript(selection);
  return wrapWordScript(`
${selectionScript}
try
  copy object selection
on error errMsg
  return "Error copying: " & errMsg
end try
return "Content copied to clipboard"
`);
}

export const clipboardTools = [
  {
    name: 'word_copy_content',
    description:
      'Copy Word content with formatting to the system clipboard. Supports selection, full document, paragraph range, or inline shape.',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        scope: {
          type: 'string',
          enum: ['selection', 'document', 'paragraph_range', 'inline_shape'],
          description: 'What to copy. Defaults to selection, or paragraph_range/inline_shape if those selectors are provided.'
        },
        startParagraph: { type: 'integer', description: 'First paragraph to copy (1-based). Required for paragraph_range.' },
        endParagraph: { type: 'integer', description: 'Last paragraph to copy (1-based, default = startParagraph).' },
        inlineShapeIndex: { type: 'integer', description: 'Inline shape index to copy (1-based). Required for inline_shape.' }
      }
    },
    async handler(args) {
      if (args.endParagraph !== undefined && args.startParagraph === undefined && args.scope !== 'paragraph_range') {
        throw new Error('endParagraph requires startParagraph');
      }
      const selection = resolveCopySelection(args);
      return await runAppleScript(buildWordCopyScript(selection));
    }
  },
  {
    name: 'word_capture_content_ref',
    description: 'Legacy tool disabled by in-place editing policy. Use word_copy_content or word_copy_story_content instead.',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        scope: {
          type: 'string',
          enum: ['selection', 'document', 'paragraph_range', 'inline_shape'],
          description: 'What to capture. Defaults to selection, or paragraph_range/inline_shape if those selectors are provided.'
        },
        startParagraph: { type: 'integer', description: 'First paragraph to capture (1-based). Required for paragraph_range.' },
        endParagraph: { type: 'integer', description: 'Last paragraph to capture (1-based, default = startParagraph).' },
        inlineShapeIndex: { type: 'integer', description: 'Inline shape index to capture (1-based). Required for inline_shape.' }
      }
    },
    async handler(args) {
      resolveCopySelection(args);
      throw new ToolError('NOT_SUPPORTED', 'word_capture_content_ref is disabled by in-place editing policy. Use word_copy_content or word_copy_story_content.');
    }
  },
  {
    name: 'word_insert_content_ref',
    description: 'Insert a local image ref at the current selection in the active document. Native Word fragment refs are disabled by in-place editing policy.',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        ref: { type: 'string', description: 'Opaque ref returned by word_create_image_ref.' },
        width: { type: 'number', description: 'Optional width for image refs in points.' },
        height: { type: 'number', description: 'Optional height for image refs in points.' }
      },
      required: ['ref']
    },
    async handler(args) {
      const ref = args.ref;
      if (typeof ref !== 'string' || ref.length === 0) {
        throw new Error('ref is required');
      }

      const fragment = getFragment(ref, 'word');
      if (fragment.kind === 'image_file') {
        const width = args.width !== undefined ? validateNumber(args.width, 'width', 1, 10000) : undefined;
        const height = args.height !== undefined ? validateNumber(args.height, 'height', 1, 10000) : undefined;
        const result = await runAppleScript(buildWordInsertImageScript(fragment.filePath, { width, height }));
        if (!result.startsWith('Image inserted successfully')) {
          throw new ToolError('OPERATION_ERROR', result);
        }
        return { inserted: true, ref: fragment.ref, kind: fragment.kind };
      }

      if (fragment.kind !== 'word_fragment') {
        throw new ToolError('VALIDATION_ERROR', `ref kind is not supported in Word: ${fragment.kind}`);
      }
      throw new ToolError('NOT_SUPPORTED', 'Native Word fragment refs are disabled by in-place editing policy. Use word_paste_content or word_copy_story_content.');
    }
  },
  {
    name: 'word_paste_content',
    description: 'Paste content from the system clipboard at the current cursor position preserving formatting.',
    annotations: { destructiveHint: true },
    inputSchema: { type: 'object', properties: {} },
    async handler() {
      const script = wrapWordScript(`
try
  paste object selection
on error errMsg
  return "Error pasting: " & errMsg
end try
return "Content pasted from clipboard"
`);
      return await runAppleScript(script);
    }
  }
];
