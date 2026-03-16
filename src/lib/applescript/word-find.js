import { getErrorMessage } from '../validators.js';
import { runAppleScript } from './executor.js';
import { buildWordExecuteFind, escapeForWordFind, quoteAppleScriptString } from './helpers.js';
import { wrapWordScript } from './script-wrappers.js';

export const WORD_FIND_MODES = Object.freeze({
  REPLACE: 'replace',
  DELETE_ALL: 'delete_all',
  MOVE_CURSOR_AFTER_TEXT: 'move_cursor_after_text'
});

export const WORD_FIND_STRATEGIES = Object.freeze({
  DIRECT_EXECUTE_PARAMS: 'direct_execute_params',
  LEGACY_FIND_OBJECT_CONTENT: 'legacy_find_object_content'
});

function normalizeWordFindSpec({
  mode,
  findText,
  replaceWith = '',
  occurrence = 1,
  replaceAll = true
}) {
  if (!Object.values(WORD_FIND_MODES).includes(mode)) {
    throw new Error(`Unsupported Word find mode: ${mode}`);
  }

  return {
    mode,
    findText,
    replaceWith,
    occurrence,
    replaceAll
  };
}

function buildWordFindSetup() {
  return `
set activeDoc to active document
select (text object of activeDoc)
set selection end of selection to selection start of selection
try
  set findObject to find object of selection
on error
  return "Cannot access find object. Make sure a document is active."
end try
clear formatting findObject
`.trim();
}

function buildMoveCursorBody(findCommand, occurrence, findText) {
  return `
set foundCount to 0
repeat ${occurrence} times
  set findResult to ${findCommand}
  if findResult is not true then
    exit repeat
  end if
  set foundCount to foundCount + 1
  if foundCount < ${occurrence} then
    set selection start of selection to selection end of selection
    set selection end of selection to selection start of selection
  end if
end repeat
if foundCount < ${occurrence} then
  return "Text not found (or fewer than ${occurrence} occurrences): " & ${quoteAppleScriptString(findText)}
end if
set selection start of selection to selection end of selection
return "Cursor moved after occurrence " & ${occurrence} & " of: " & ${quoteAppleScriptString(findText)}
`.trim();
}

function buildReplaceBody(strategy, findText, replaceWith, replaceAll) {
  const replace = replaceAll ? 'replace all' : 'replace one';

  if (strategy === WORD_FIND_STRATEGIES.DIRECT_EXECUTE_PARAMS) {
    return `
set findResult to ${buildWordExecuteFind('findObject', {
  findText,
  replaceWith,
  replace
})}
if findResult then
  return "Text replaced successfully"
else
  return "Text not found, no replacements made"
end if
`.trim();
  }

  if (strategy === WORD_FIND_STRATEGIES.LEGACY_FIND_OBJECT_CONTENT) {
    return `
set content of findObject to ${escapeForWordFind(findText)}
set content of replacement of findObject to ${escapeForWordFind(replaceWith)}
set findResult to execute find findObject replace ${replace}
if findResult then
  return "Text replaced successfully"
else
  return "Text not found, no replacements made"
end if
`.trim();
  }

  throw new Error(`Unsupported Word find strategy: ${strategy}`);
}

function buildDeleteAllBody(strategy, findText) {
  if (strategy === WORD_FIND_STRATEGIES.DIRECT_EXECUTE_PARAMS) {
    return `
set findResult to ${buildWordExecuteFind('findObject', {
  findText,
  replaceWith: '',
  replace: 'replace all'
})}
if findResult then
  return "Text deleted successfully"
else
  return "Text not found, nothing deleted"
end if
`.trim();
  }

  if (strategy === WORD_FIND_STRATEGIES.LEGACY_FIND_OBJECT_CONTENT) {
    return `
set content of findObject to ${escapeForWordFind(findText)}
set content of replacement of findObject to ""
set findResult to execute find findObject replace replace all
if findResult then
  return "Text deleted successfully"
else
  return "Text not found, nothing deleted"
end if
`.trim();
  }

  throw new Error(`Unsupported Word find strategy: ${strategy}`);
}

function buildMoveCursorAfterTextBody(strategy, findText, occurrence) {
  if (strategy === WORD_FIND_STRATEGIES.DIRECT_EXECUTE_PARAMS) {
    return buildMoveCursorBody(
      buildWordExecuteFind('findObject', {
        findText,
        matchForward: true,
        wrapFind: 'find stop'
      }),
      occurrence,
      findText
    );
  }

  if (strategy === WORD_FIND_STRATEGIES.LEGACY_FIND_OBJECT_CONTENT) {
    return `
set content of findObject to ${escapeForWordFind(findText)}
set wrap of findObject to find stop
set forward of findObject to true
${buildMoveCursorBody('execute find findObject', occurrence, findText)}
`.trim();
  }

  throw new Error(`Unsupported Word find strategy: ${strategy}`);
}

function buildWordFindModeBody(strategy, spec) {
  switch (spec.mode) {
    case WORD_FIND_MODES.REPLACE:
      return buildReplaceBody(strategy, spec.findText, spec.replaceWith, spec.replaceAll);
    case WORD_FIND_MODES.DELETE_ALL:
      return buildDeleteAllBody(strategy, spec.findText);
    case WORD_FIND_MODES.MOVE_CURSOR_AFTER_TEXT:
      return buildMoveCursorAfterTextBody(strategy, spec.findText, spec.occurrence);
    default:
      throw new Error(`Unsupported Word find mode: ${spec.mode}`);
  }
}

export function buildWordFindScript({ strategy = WORD_FIND_STRATEGIES.DIRECT_EXECUTE_PARAMS, ...spec }) {
  const normalized = normalizeWordFindSpec(spec);
  const body = `${buildWordFindSetup()}
${buildWordFindModeBody(strategy, normalized)}`;
  return wrapWordScript(body);
}

export function isWordFindCompatibilityError(error) {
  const message = getErrorMessage(error);
  const hasExecuteFind = /execute find/i.test(message);
  const hasCompatibilitySignal = /-1708/.test(message) || /doesn['’]t understand/i.test(message) || /does not understand/i.test(message);
  return hasExecuteFind && hasCompatibilitySignal;
}

export async function runWordFindWithFallback(spec, { executeAppleScript = runAppleScript } = {}) {
  const directStrategy = WORD_FIND_STRATEGIES.DIRECT_EXECUTE_PARAMS;
  const fallbackStrategy = WORD_FIND_STRATEGIES.LEGACY_FIND_OBJECT_CONTENT;
  const directScript = buildWordFindScript({ ...spec, strategy: directStrategy });

  try {
    return await executeAppleScript(directScript);
  } catch (directError) {
    if (!isWordFindCompatibilityError(directError)) {
      throw directError;
    }

    const fallbackScript = buildWordFindScript({ ...spec, strategy: fallbackStrategy });

    try {
      return await executeAppleScript(fallbackScript);
    } catch (fallbackError) {
      throw new Error([
        'Word find failed for both strategies.',
        `${directStrategy}: ${getErrorMessage(directError)}`,
        `${fallbackStrategy}: ${getErrorMessage(fallbackError)}`
      ].join('\n'));
    }
  }
}
