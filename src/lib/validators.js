export function validateString(value, name, required = true) {
  if (value === undefined || value === null) {
    if (required) {
      throw new Error(`${name} is required`);
    }
    return '';
  }
  if (typeof value !== 'string') {
    throw new Error(`${name} must be a string`);
  }
  if (required && value.length === 0) {
    throw new Error(`${name} cannot be empty`);
  }
  return value;
}

export function validateBoolean(value, name, defaultValue = false) {
  if (value === undefined || value === null) {
    return defaultValue;
  }
  if (typeof value !== 'boolean') {
    throw new Error(`${name} must be a boolean`);
  }
  return value;
}

export function validateEnum(value, name, allowedValues, defaultValue = undefined) {
  if (value === undefined || value === null) {
    return defaultValue;
  }
  const str = validateString(value, name, true);
  if (!allowedValues.includes(str)) {
    throw new Error(`${name} must be one of: ${allowedValues.join(', ')}`);
  }
  return str;
}

export function validateNumber(value, name, min = 0, max = Number.MAX_SAFE_INTEGER) {
  if (value === undefined || value === null) {
    return undefined;
  }
  if (typeof value === 'boolean') {
    throw new Error(`${name} must be a valid number`);
  }
  const num = Number(value);
  if (!Number.isFinite(num)) {
    throw new Error(`${name} must be a valid number`);
  }
  if (num < min || num > max) {
    throw new Error(`${name} must be between ${min} and ${max}`);
  }
  return num;
}

export function validateInteger(value, name, min = 0, max = Number.MAX_SAFE_INTEGER) {
  if (value === undefined || value === null) {
    return undefined;
  }
  if (typeof value === 'boolean') {
    throw new Error(`${name} must be an integer`);
  }
  const num = Number(value);
  if (!Number.isInteger(num)) {
    throw new Error(`${name} must be an integer`);
  }
  if (num < min || num > max) {
    throw new Error(`${name} must be between ${min} and ${max}`);
  }
  return num;
}

export function validateExcelCellReference(value, name = 'cell') {
  const cell = validateString(value, name, true).toUpperCase();
  if (!/^[A-Z]{1,3}[1-9][0-9]*$/.test(cell)) {
    throw new Error(`${name} must be a valid A1 cell reference`);
  }
  return cell;
}

export function validateExcelRangeReference(value, name = 'range') {
  const range = validateString(value, name, true).toUpperCase();
  const isCell = /^[A-Z]{1,3}[1-9][0-9]*$/.test(range);
  const isCellRange = /^[A-Z]{1,3}[1-9][0-9]*:[A-Z]{1,3}[1-9][0-9]*$/.test(range);
  const isColumnRange = /^[A-Z]{1,3}:[A-Z]{1,3}$/.test(range);
  const isRowRange = /^[1-9][0-9]*:[1-9][0-9]*$/.test(range);
  if (!isCell && !isCellRange && !isColumnRange && !isRowRange) {
    throw new Error(`${name} must be a valid Excel range reference`);
  }
  return range;
}

export function getErrorMessage(error) {
  if (error instanceof Error) return error.message;
  if (typeof error === 'string') return error;
  return String(error);
}
