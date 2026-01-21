import { COMMON_SCRIPTS } from './helpers.js';

export function processTemplate(template, params = {}) {
  let script = template;

  script = script.replace(/<<CHECK_DOCUMENT>>/g, COMMON_SCRIPTS.checkDocumentOpen);
  script = script.replace(/<<GET_ACTIVE_DOC>>/g, COMMON_SCRIPTS.getActiveDocument);
  script = script.replace(/<<CLEAN_CELL>>/g, COMMON_SCRIPTS.cleanCellMarkers);
  script = script.replace(/<<COLLAPSE_TO_START>>/g, COMMON_SCRIPTS.collapseToStart);
  script = script.replace(/<<COLLAPSE_TO_END>>/g, COMMON_SCRIPTS.collapseToEnd);

  for (const [key, value] of Object.entries(params)) {
    const placeholder = new RegExp(`<<${key}>>`, 'g');
    let replacementValue;

    if (typeof value === 'string') {
      replacementValue = value.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
      replacementValue = `"${replacementValue}"`;
    } else if (typeof value === 'boolean') {
      replacementValue = value ? 'true' : 'false';
    } else if (typeof value === 'number') {
      replacementValue = value.toString();
    } else {
      replacementValue = String(value);
    }

    script = script.replace(placeholder, replacementValue);
  }

  return script;
}
