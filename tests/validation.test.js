import { describe, test, expect } from '@jest/globals';

describe('Validation Functions', () => {
  describe('validateString', () => {
    test('should validate required string', () => {
      const validateString = (value, name, required = true) => {
        if (value === undefined || value === null) {
          if (required) {
            throw new Error(`${name} is required`);
          }
          return "";
        }
        if (typeof value !== "string") {
          throw new Error(`${name} must be a string`);
        }
        if (required && value.trim() === "") {
          throw new Error(`${name} cannot be empty`);
        }
        return value;
      };

      expect(validateString('test', 'field')).toBe('test');
      expect(() => validateString(null, 'field')).toThrow('field is required');
      expect(() => validateString(undefined, 'field')).toThrow('field is required');
      expect(() => validateString(123, 'field')).toThrow('field must be a string');
      expect(() => validateString('  ', 'field')).toThrow('field cannot be empty');
      expect(validateString(null, 'field', false)).toBe('');
    });
  });

  describe('validateBoolean', () => {
    test('should validate boolean with default', () => {
      const validateBoolean = (value, name, defaultValue = false) => {
        if (value === undefined || value === null) {
          return defaultValue;
        }
        if (typeof value !== "boolean") {
          throw new Error(`${name} must be a boolean`);
        }
        return value;
      };

      expect(validateBoolean(true, 'field')).toBe(true);
      expect(validateBoolean(false, 'field')).toBe(false);
      expect(validateBoolean(null, 'field')).toBe(false);
      expect(validateBoolean(undefined, 'field', true)).toBe(true);
      expect(() => validateBoolean('true', 'field')).toThrow('field must be a boolean');
      expect(() => validateBoolean(1, 'field')).toThrow('field must be a boolean');
    });
  });

  describe('validateNumber', () => {
    test('should validate number with min/max', () => {
      const validateNumber = (value, name, min = 0, max = Number.MAX_SAFE_INTEGER) => {
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
      };

      expect(validateNumber(5, 'field')).toBe(5);
      expect(validateNumber(10, 'field', 0, 20)).toBe(10);
      expect(validateNumber(null, 'field')).toBeUndefined();
      expect(() => validateNumber(50, 'field', 0, 20)).toThrow('field must be between 0 and 20');
      expect(() => validateNumber(-5, 'field', 0, 20)).toThrow('field must be between 0 and 20');
      expect(() => validateNumber(NaN, 'field')).toThrow('field must be a valid number');
      expect(() => validateNumber(Infinity, 'field')).toThrow('field must be a valid number');
    });
  });

  describe('validateInteger', () => {
    test('should validate integer with min/max', () => {
      const validateInteger = (value, name, min = 0, max = Number.MAX_SAFE_INTEGER) => {
        if (value === undefined || value === null) {
          return undefined;
        }
        const num = parseInt(value, 10);
        if (!Number.isInteger(num)) {
          throw new Error(`${name} must be an integer`);
        }
        if (num < min || num > max) {
          throw new Error(`${name} must be between ${min} and ${max}`);
        }
        return num;
      };

      expect(validateInteger(5, 'field')).toBe(5);
      expect(validateInteger('10', 'field')).toBe(10);
      expect(validateInteger(null, 'field')).toBeUndefined();
      expect(() => validateInteger(50, 'field', 0, 20)).toThrow('field must be between 0 and 20');
      expect(() => validateInteger(-5, 'field', 1, 20)).toThrow('field must be between 1 and 20');
    });
  });

  describe('getErrorMessage', () => {
    test('should extract error message', () => {
      const getErrorMessage = (error) => {
        if (error instanceof Error) return error.message;
        if (typeof error === "string") return error;
        return String(error);
      };

      expect(getErrorMessage(new Error('test error'))).toBe('test error');
      expect(getErrorMessage('string error')).toBe('string error');
      expect(getErrorMessage(123)).toBe('123');
      expect(getErrorMessage({ toString: () => 'object error' })).toBe('object error');
    });
  });
});
