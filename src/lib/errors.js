const ERROR_PATTERNS = [
  { code: 'NO_DOCUMENT_OPEN', pattern: /^No document is open$/i },
  { code: 'NO_WORKBOOK_OPEN', pattern: /^No workbook is open$/i },
  { code: 'NOT_FOUND', pattern: /\bnot found\b/i },
  { code: 'OUT_OF_RANGE', pattern: /\bout of range\b/i },
  { code: 'VALIDATION_ERROR', pattern: /\b(is required|must be|cannot be)\b/i },
  { code: 'APPSCRIPT_ERROR', pattern: /^AppleScript error:/i },
  { code: 'OPERATION_ERROR', pattern: /^(Error|Cannot)\b/i }
];

export class ToolError extends Error {
  constructor(code, message, details = undefined) {
    super(message);
    this.name = 'ToolError';
    this.code = code;
    this.details = details;
  }
}

export function inferErrorCode(message) {
  for (const rule of ERROR_PATTERNS) {
    if (rule.pattern.test(message)) {
      return rule.code;
    }
  }
  return 'OPERATION_ERROR';
}

export function isLikelyErrorMessage(value) {
  if (typeof value !== 'string' || value.length === 0) {
    return false;
  }
  return ERROR_PATTERNS.some(rule => rule.pattern.test(value));
}

