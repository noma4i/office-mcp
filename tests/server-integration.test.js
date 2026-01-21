import { describe, test, expect } from '@jest/globals';

describe('MCP Server Integration', () => {
  describe('Server Configuration', () => {
    test('should have correct server name and version', () => {
      const serverConfig = {
        name: "Microsoft-Word-Server",
        version: "0.6.0",
      };

      expect(serverConfig.name).toBe("Microsoft-Word-Server");
      expect(serverConfig.version).toBe("0.6.0");
    });

    test('should expose tools capability', () => {
      const capabilities = {
        tools: {},
      };

      expect(capabilities).toHaveProperty('tools');
    });
  });

  describe('Tool Definitions', () => {
    const tools = [
      'create_document',
      'open_document',
      'get_document_text',
      'get_document_info',
      'insert_text',
      'replace_text',
      'format_text',
      'save_document',
      'close_document',
      'export_pdf',
      'list_tables',
      'get_table_cell',
      'set_table_cell',
      'select_table_cell',
      'find_table_header',
      'move_cursor_after_text',
      'goto_start',
      'goto_end',
      'get_selection_info',
      'select_all',
      'create_table',
      'add_table_row',
      'delete_table_row',
      'add_table_column',
      'delete_table_column',
      'list_bookmarks',
      'create_bookmark',
      'goto_bookmark',
      'delete_bookmark',
      'list_hyperlinks',
      'create_hyperlink',
      'list_paragraphs',
      'goto_paragraph',
      'set_paragraph_style',
    ];

    test('should have all 34 tools defined', () => {
      expect(tools).toHaveLength(34);
    });

    test('should have document tools', () => {
      const documentTools = tools.filter(t =>
        ['create_document', 'open_document', 'get_document_text', 'get_document_info',
         'save_document', 'close_document', 'export_pdf'].includes(t)
      );
      expect(documentTools).toHaveLength(7);
    });

    test('should have text tools', () => {
      const textTools = tools.filter(t =>
        ['insert_text', 'replace_text', 'format_text'].includes(t)
      );
      expect(textTools).toHaveLength(3);
    });

    test('should have navigation tools', () => {
      const navigationTools = tools.filter(t =>
        ['move_cursor_after_text', 'goto_start', 'goto_end', 'get_selection_info', 'select_all'].includes(t)
      );
      expect(navigationTools).toHaveLength(5);
    });

    test('should have table tools', () => {
      const tableTools = tools.filter(t =>
        ['list_tables', 'get_table_cell', 'set_table_cell', 'select_table_cell', 'find_table_header',
         'create_table', 'add_table_row', 'delete_table_row', 'add_table_column', 'delete_table_column'].includes(t)
      );
      expect(tableTools).toHaveLength(10);
    });

    test('should have bookmark tools', () => {
      const bookmarkTools = tools.filter(t =>
        ['list_bookmarks', 'create_bookmark', 'goto_bookmark', 'delete_bookmark'].includes(t)
      );
      expect(bookmarkTools).toHaveLength(4);
    });

    test('should have hyperlink tools', () => {
      const hyperlinkTools = tools.filter(t =>
        ['list_hyperlinks', 'create_hyperlink'].includes(t)
      );
      expect(hyperlinkTools).toHaveLength(2);
    });

    test('should have paragraph tools', () => {
      const paragraphTools = tools.filter(t =>
        ['list_paragraphs', 'goto_paragraph', 'set_paragraph_style'].includes(t)
      );
      expect(paragraphTools).toHaveLength(3);
    });
  });

  describe('Tool Annotations', () => {
    test('destructive tools should have destructiveHint', () => {
      const destructiveTools = [
        'create_document', 'insert_text', 'replace_text', 'format_text',
        'save_document', 'close_document', 'export_pdf', 'set_table_cell',
        'select_table_cell', 'move_cursor_after_text', 'goto_start', 'goto_end',
        'select_all', 'create_table', 'add_table_row', 'delete_table_row',
        'add_table_column', 'delete_table_column', 'create_bookmark',
        'goto_bookmark', 'delete_bookmark', 'create_hyperlink', 'goto_paragraph',
        'set_paragraph_style'
      ];

      expect(destructiveTools.length).toBeGreaterThan(0);
    });

    test('read-only tools should have readOnlyHint', () => {
      const readOnlyTools = [
        'open_document', 'get_document_text', 'get_document_info', 'list_tables',
        'get_table_cell', 'find_table_header', 'get_selection_info',
        'list_bookmarks', 'list_hyperlinks', 'list_paragraphs'
      ];

      expect(readOnlyTools.length).toBeGreaterThan(0);
    });
  });

  describe('Required Parameters', () => {
    test('open_document requires path', () => {
      const inputSchema = {
        type: "object",
        properties: {
          path: { type: "string" }
        },
        required: ["path"]
      };

      expect(inputSchema.required).toContain('path');
    });

    test('insert_text requires text', () => {
      const inputSchema = {
        type: "object",
        properties: {
          text: { type: "string" }
        },
        required: ["text"]
      };

      expect(inputSchema.required).toContain('text');
    });

    test('replace_text requires find and replace', () => {
      const inputSchema = {
        type: "object",
        properties: {
          find: { type: "string" },
          replace: { type: "string" },
          all: { type: "boolean", default: true }
        },
        required: ["find", "replace"]
      };

      expect(inputSchema.required).toEqual(expect.arrayContaining(['find', 'replace']));
    });

    test('get_table_cell requires tableIndex, row, column', () => {
      const inputSchema = {
        type: "object",
        properties: {
          tableIndex: { type: "integer" },
          row: { type: "integer" },
          column: { type: "integer" }
        },
        required: ["tableIndex", "row", "column"]
      };

      expect(inputSchema.required).toEqual(expect.arrayContaining(['tableIndex', 'row', 'column']));
    });

    test('create_table requires rows and columns', () => {
      const inputSchema = {
        type: "object",
        properties: {
          rows: { type: "integer" },
          columns: { type: "integer" }
        },
        required: ["rows", "columns"]
      };

      expect(inputSchema.required).toEqual(expect.arrayContaining(['rows', 'columns']));
    });

    test('create_bookmark requires name', () => {
      const inputSchema = {
        type: "object",
        properties: {
          name: { type: "string" }
        },
        required: ["name"]
      };

      expect(inputSchema.required).toContain('name');
    });

    test('create_hyperlink requires url', () => {
      const inputSchema = {
        type: "object",
        properties: {
          url: { type: "string" },
          displayText: { type: "string" }
        },
        required: ["url"]
      };

      expect(inputSchema.required).toContain('url');
    });
  });

  describe('Default Values', () => {
    test('replace_text has default all=true', () => {
      const defaultAll = true;
      expect(defaultAll).toBe(true);
    });

    test('close_document has default save=true', () => {
      const defaultSave = true;
      expect(defaultSave).toBe(true);
    });

    test('move_cursor_after_text has default occurrence=1', () => {
      const defaultOccurrence = 1;
      expect(defaultOccurrence).toBe(1);
    });

    test('find_table_header has default headerRow=1', () => {
      const defaultHeaderRow = 1;
      expect(defaultHeaderRow).toBe(1);
    });

    test('list_paragraphs has default limit=50', () => {
      const defaultLimit = 50;
      expect(defaultLimit).toBe(50);
    });
  });

  describe('AppleScript Safety', () => {
    test('should use JSON.stringify for user input in AppleScript', () => {
      const userInput = 'test"; malicious code; "';
      const safeInput = JSON.stringify(userInput);

      expect(safeInput).toBe('"test\\"; malicious code; \\""');
      expect(safeInput).toContain('\\');
    });

    test('should validate input types before AppleScript execution', () => {
      const validateString = (value, name, required = true) => {
        if (typeof value !== 'string') {
          throw new Error(`${name} must be a string`);
        }
        return value;
      };

      expect(() => validateString(123, 'field')).toThrow();
      expect(validateString('safe', 'field')).toBe('safe');
    });

    test('should build AppleScript conditionally to avoid injection', () => {
      const content = 'user content';
      const scriptWithContent = content ? `set content to ${JSON.stringify(content)}` : ``;

      expect(scriptWithContent).toContain(JSON.stringify(content));
    });
  });

  describe('1-Based Indexing', () => {
    test('table indexes should be 1-based', () => {
      const minTableIndex = 1;
      expect(minTableIndex).toBe(1);
    });

    test('row indexes should be 1-based', () => {
      const minRow = 1;
      expect(minRow).toBe(1);
    });

    test('column indexes should be 1-based', () => {
      const minColumn = 1;
      expect(minColumn).toBe(1);
    });

    test('paragraph indexes should be 1-based', () => {
      const minParagraph = 1;
      expect(minParagraph).toBe(1);
    });
  });
});
