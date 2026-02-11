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

export function validateNumber(value, name, min = 0, max = Number.MAX_SAFE_INTEGER) {
  if (value === undefined || value === null) {
    return undefined;
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
  const num = Number(value);
  if (!Number.isInteger(num)) {
    throw new Error(`${name} must be an integer`);
  }
  if (num < min || num > max) {
    throw new Error(`${name} must be between ${min} and ${max}`);
  }
  return num;
}

export function getErrorMessage(error) {
  if (error instanceof Error) return error.message;
  if (typeof error === 'string') return error;
  return String(error);
}
