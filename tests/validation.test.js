import { describe, test, expect } from '@jest/globals';
import { getErrorMessage, validateBoolean, validateInteger, validateNumber, validateString } from '../src/lib/validators.js';

describe('Validation Functions', () => {
  describe('validateString', () => {
    test('validates required string', () => {
      expect(validateString('test', 'field')).toBe('test');
      expect(() => validateString(null, 'field')).toThrow('field is required');
      expect(() => validateString(undefined, 'field')).toThrow('field is required');
      expect(() => validateString(123, 'field')).toThrow('field must be a string');
      expect(() => validateString('', 'field')).toThrow('field cannot be empty');
      expect(validateString(null, 'field', false)).toBe('');
    });
  });

  describe('validateBoolean', () => {
    test('validates boolean with default', () => {
      expect(validateBoolean(true, 'field')).toBe(true);
      expect(validateBoolean(false, 'field')).toBe(false);
      expect(validateBoolean(null, 'field')).toBe(false);
      expect(validateBoolean(undefined, 'field', true)).toBe(true);
      expect(() => validateBoolean('true', 'field')).toThrow('field must be a boolean');
      expect(() => validateBoolean(1, 'field')).toThrow('field must be a boolean');
    });
  });

  describe('validateNumber', () => {
    test('validates number with min/max', () => {
      expect(validateNumber(5, 'field')).toBe(5);
      expect(validateNumber('10', 'field', 0, 20)).toBe(10);
      expect(validateNumber(null, 'field')).toBeUndefined();
      expect(() => validateNumber(false, 'field')).toThrow('field must be a valid number');
      expect(() => validateNumber(50, 'field', 0, 20)).toThrow('field must be between 0 and 20');
      expect(() => validateNumber(-5, 'field', 0, 20)).toThrow('field must be between 0 and 20');
      expect(() => validateNumber(NaN, 'field')).toThrow('field must be a valid number');
      expect(() => validateNumber(Infinity, 'field')).toThrow('field must be a valid number');
    });
  });

  describe('validateInteger', () => {
    test('validates integer with min/max', () => {
      expect(validateInteger(5, 'field')).toBe(5);
      expect(validateInteger('10', 'field')).toBe(10);
      expect(validateInteger(null, 'field')).toBeUndefined();
      expect(() => validateInteger(true, 'field')).toThrow('field must be an integer');
      expect(() => validateInteger(10.2, 'field')).toThrow('field must be an integer');
      expect(() => validateInteger(50, 'field', 0, 20)).toThrow('field must be between 0 and 20');
      expect(() => validateInteger(-5, 'field', 1, 20)).toThrow('field must be between 1 and 20');
    });
  });

  describe('getErrorMessage', () => {
    test('extracts error message', () => {
      expect(getErrorMessage(new Error('test error'))).toBe('test error');
      expect(getErrorMessage('string error')).toBe('string error');
      expect(getErrorMessage(123)).toBe('123');
      expect(getErrorMessage({ toString: () => 'object error' })).toBe('object error');
    });
  });
});
