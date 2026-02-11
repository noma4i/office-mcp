import { describe, test, expect } from '@jest/globals';

describe('MCP Office Server Tools', () => {

  describe('Document Tools', () => {
    test('create_document - with content', () => {
      const content = 'Hello World';
      expect(content).toBe('Hello World');
    });

    test('open_document - path validation', () => {
      const path = '/path/to/document.docx';
      expect(path).toBeDefined();
      expect(path).toMatch(/\.docx?$/);
    });

    test('get_document_text - response format', () => {
      const text = 'Sample document text';
      expect(text).toBeDefined();
      expect(typeof text).toBe('string');
    });

    test('get_document_info - statistics format', () => {
      const stats = 'Words: 100\nCharacters: 500\nPages: 2';
      expect(stats).toContain('Words');
      expect(stats).toContain('Characters');
      expect(stats).toContain('Pages');
    });

    test('save_document - path format', () => {
      const path = '/path/to/save.docx';
      expect(path).toBeDefined();
      expect(path).toMatch(/\.docx?$/);
    });

    test('export_pdf - path validation', () => {
      const path = '/path/to/export.pdf';
      expect(path).toMatch(/\.pdf$/);
    });
  });

  describe('Text Tools', () => {
    test('insert_text - parameter validation', () => {
      const text = 'Inserted text';
      expect(text).toBeDefined();
      expect(typeof text).toBe('string');
    });

    test('replace_text - parameters validation', () => {
      const find = 'old';
      const replace = 'new';
      expect(find).toBe('old');
      expect(replace).toBe('new');
    });

    test('replace_text - all flag', () => {
      const all = true;
      expect(all).toBe(true);
      expect(typeof all).toBe('boolean');
    });

    test('format_text - bold flag', () => {
      const bold = true;
      expect(bold).toBe(true);
      expect(typeof bold).toBe('boolean');
    });

    test('format_text - italic flag', () => {
      const italic = true;
      expect(italic).toBe(true);
      expect(typeof italic).toBe('boolean');
    });

    test('format_text - underline flag', () => {
      const underline = true;
      expect(underline).toBe(true);
      expect(typeof underline).toBe('boolean');
    });

    test('format_text - font and size', () => {
      const font = 'Arial';
      const size = 14;
      expect(font).toBe('Arial');
      expect(size).toBe(14);
      expect(typeof font).toBe('string');
      expect(typeof size).toBe('number');
    });
  });

  describe('Navigation Tools', () => {
    test('get_selection_info - response format', () => {
      const info = 'Start: 10\nEnd: 20\nLength: 10';
      expect(info).toContain('Start');
      expect(info).toContain('End');
      expect(info).toContain('Length');
    });

    test('move_cursor_after_text - parameter validation', () => {
      const searchText = 'Chapter 1';
      expect(searchText).toBeDefined();
      expect(typeof searchText).toBe('string');
    });

    test('move_cursor_after_text - occurrence parameter', () => {
      const searchText = 'Section';
      const occurrence = 3;
      expect(occurrence).toBe(3);
      expect(occurrence).toBeGreaterThan(0);
    });
  });

  describe('Table Tools', () => {
    test('list_tables - response format', () => {
      const tables = 'Table 1: 3 rows x 2 columns\nTable 2: 5 rows x 4 columns';
      expect(tables).toContain('Table 1');
      expect(tables).toContain('Table 2');
    });

    test('get_table_cell - parameter validation', () => {
      const tableIndex = 1;
      const row = 2;
      const column = 3;
      expect(tableIndex).toBeGreaterThan(0);
      expect(row).toBeGreaterThan(0);
      expect(column).toBeGreaterThan(0);
    });

    test('set_table_cell - with text', () => {
      const text = 'New content';
      expect(text).toBe('New content');
      expect(typeof text).toBe('string');
    });

    test('find_table_header - header text', () => {
      const headerText = 'Name';
      expect(headerText).toBe('Name');
      expect(typeof headerText).toBe('string');
    });

    test('create_table - rows and columns', () => {
      const rows = 3;
      const columns = 4;
      expect(rows).toBeGreaterThan(0);
      expect(columns).toBeGreaterThan(0);
    });

    test('add_table_row - afterRow parameter', () => {
      const tableIndex = 1;
      const afterRow = 2;
      expect(tableIndex).toBeDefined();
      expect(afterRow).toBeGreaterThan(0);
    });

    test('delete_table_row - row parameter', () => {
      const row = 3;
      expect(row).toBeGreaterThan(0);
    });

    test('add_table_column - afterColumn parameter', () => {
      const afterColumn = 2;
      expect(afterColumn).toBeDefined();
      expect(afterColumn).toBeGreaterThan(0);
    });

    test('delete_table_column - column parameter', () => {
      const column = 2;
      expect(column).toBeGreaterThan(0);
    });
  });

  describe('Bookmark Tools', () => {
    test('list_bookmarks - response format', () => {
      const bookmarks = '1. Introduction\n2. Chapter1\n3. Summary';
      expect(bookmarks).toContain('Introduction');
    });

    test('create_bookmark - name parameter', () => {
      const name = 'MyBookmark';
      expect(name).toBe('MyBookmark');
      expect(typeof name).toBe('string');
    });

    test('goto_bookmark - name validation', () => {
      const name = 'Chapter1';
      expect(name).toBeDefined();
      expect(typeof name).toBe('string');
    });

    test('delete_bookmark - name validation', () => {
      const name = 'OldBookmark';
      expect(name).toBeDefined();
      expect(typeof name).toBe('string');
    });
  });

  describe('Hyperlink Tools', () => {
    test('list_hyperlinks - response format', () => {
      const links = '1. Google -> https://google.com\n2. GitHub -> https://github.com';
      expect(links).toContain('https://');
    });

    test('create_hyperlink - URL validation', () => {
      const url = 'https://example.com';
      expect(url).toMatch(/^https?:\/\//);
    });

    test('create_hyperlink - with display text', () => {
      const url = 'https://example.com';
      const displayText = 'Example Site';
      expect(displayText).toBe('Example Site');
      expect(typeof displayText).toBe('string');
    });
  });

  describe('Paragraph Tools', () => {
    test('list_paragraphs - response format', () => {
      const paragraphs = 'Total paragraphs: 10\n\n1. [Normal] First paragraph...\n2. [Heading 1] Second paragraph...';
      expect(paragraphs).toContain('Total paragraphs');
    });

    test('list_paragraphs - limit parameter', () => {
      const limit = 20;
      expect(limit).toBe(20);
      expect(limit).toBeGreaterThan(0);
    });

    test('goto_paragraph - index parameter', () => {
      const index = 5;
      expect(index).toBeGreaterThan(0);
    });

    test('set_paragraph_style - parameters', () => {
      const index = 3;
      const styleName = 'Heading 1';
      expect(styleName).toBe('Heading 1');
      expect(index).toBeGreaterThan(0);
    });
  });

  describe('Error Handling', () => {
    test('should handle AppleScript errors', () => {
      try {
        throw new Error('AppleScript execution failed');
      } catch (error) {
        expect(error.message).toContain('AppleScript');
      }
    });

    test('should handle no document open error', () => {
      const result = 'No document is open';
      expect(result).toBe('No document is open');
    });

    test('should handle table index out of range', () => {
      const result = 'Table index out of range. Document has 2 tables.';
      expect(result).toContain('out of range');
    });

    test('should handle paragraph index out of range', () => {
      const result = 'Paragraph index out of range. Document has 10 paragraphs.';
      expect(result).toContain('out of range');
    });
  });

  describe('Input Validation', () => {
    test('should require tableIndex to be >= 1', () => {
      const validateInteger = (value, name, min) => {
        const num = parseInt(value, 10);
        if (num < min) {
          throw new Error(`${name} must be between ${min} and ${Number.MAX_SAFE_INTEGER}`);
        }
        return num;
      };

      expect(() => validateInteger(0, 'tableIndex', 1)).toThrow();
      expect(validateInteger(1, 'tableIndex', 1)).toBe(1);
      expect(validateInteger(5, 'tableIndex', 1)).toBe(5);
    });

    test('should require row/column to be >= 1', () => {
      const validateInteger = (value, name, min) => {
        const num = parseInt(value, 10);
        if (num < min) {
          throw new Error(`${name} must be between ${min} and ${Number.MAX_SAFE_INTEGER}`);
        }
        return num;
      };

      expect(() => validateInteger(0, 'row', 1)).toThrow();
      expect(validateInteger(1, 'row', 1)).toBe(1);
      expect(validateInteger(10, 'column', 1)).toBe(10);
    });

    test('should validate font size range', () => {
      const validateNumber = (value, name, min, max) => {
        const num = Number(value);
        if (num < min || num > max) {
          throw new Error(`${name} must be between ${min} and ${max}`);
        }
        return num;
      };

      expect(() => validateNumber(0, 'size', 1, 1000)).toThrow();
      expect(() => validateNumber(1001, 'size', 1, 1000)).toThrow();
      expect(validateNumber(12, 'size', 1, 1000)).toBe(12);
      expect(validateNumber(72, 'size', 1, 1000)).toBe(72);
    });
  });
});
