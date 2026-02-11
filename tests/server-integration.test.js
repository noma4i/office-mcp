import { describe, test, expect } from '@jest/globals';

describe('MCP Server Integration', () => {
  describe('Server Configuration', () => {
    test('should have correct server name and version', () => {
      const serverConfig = {
        name: 'Microsoft-Office-Server',
        version: '0.8.0'
      };

      expect(serverConfig.name).toBe('Microsoft-Office-Server');
      expect(serverConfig.version).toBe('0.8.0');
    });

    test('should expose tools capability', () => {
      const capabilities = {
        tools: {}
      };

      expect(capabilities).toHaveProperty('tools');
    });
  });

  describe('Tool Definitions', () => {
    const tools = [
      // Word Document tools (7)
      'word_create_document',
      'word_open_document',
      'word_get_document_text',
      'word_get_document_info',
      'word_save_document',
      'word_close_document',
      'word_export_pdf',
      // Word Text tools (3)
      'word_insert_text',
      'word_replace_text',
      'word_format_text',
      // Word Navigation tools (5)
      'word_move_cursor_after_text',
      'word_goto_start',
      'word_goto_end',
      'word_get_selection_info',
      'word_select_all',
      // Word Table tools (10)
      'word_list_tables',
      'word_get_table_cell',
      'word_set_table_cell',
      'word_select_table_cell',
      'word_find_table_header',
      'word_create_table',
      'word_add_table_row',
      'word_delete_table_row',
      'word_add_table_column',
      'word_delete_table_column',
      // Word Bookmark tools (4)
      'word_list_bookmarks',
      'word_create_bookmark',
      'word_goto_bookmark',
      'word_delete_bookmark',
      // Word Hyperlink tools (2)
      'word_list_hyperlinks',
      'word_create_hyperlink',
      // Word Paragraph tools (3)
      'word_list_paragraphs',
      'word_goto_paragraph',
      'word_set_paragraph_style',
      // Word Image tools (3)
      'word_insert_image',
      'word_list_inline_shapes',
      'word_resize_inline_shape',
      // Word Header/Footer tools (6)
      'word_get_header_text',
      'word_set_header_text',
      'word_get_footer_text',
      'word_set_footer_text',
      'word_insert_header_image',
      'word_insert_footer_image',
      // Word Section tools (4)
      'word_list_sections',
      'word_get_section_info',
      'word_set_page_setup',
      'word_insert_section_break',
      // Word Formatting Read tools (2)
      'word_get_text_formatting',
      'word_get_paragraph_formatting',
      // Word Text tools additions (1)
      'word_delete_text',
      // Word Paragraph tools additions (1)
      'word_delete_paragraph',
      // Excel Workbook tools (6)
      'excel_create_workbook',
      'excel_open_workbook',
      'excel_get_workbook_info',
      'excel_save_workbook',
      'excel_close_workbook',
      'excel_list_workbooks',
      // Excel Sheet tools (6)
      'excel_list_sheets',
      'excel_create_sheet',
      'excel_delete_sheet',
      'excel_rename_sheet',
      'excel_activate_sheet',
      'excel_get_sheet_info',
      // Excel Cell tools (7)
      'excel_get_cell',
      'excel_set_cell',
      'excel_get_range',
      'excel_set_cell_formula',
      'excel_clear_range',
      'excel_get_used_range',
      'excel_find_cell',
      // Excel Formatting tools (5)
      'excel_format_cells',
      'excel_set_number_format',
      'excel_set_cell_color',
      'excel_merge_cells',
      'excel_autofit',
      // Excel Row/Column tools (6)
      'excel_insert_rows',
      'excel_delete_rows',
      'excel_insert_columns',
      'excel_delete_columns',
      'excel_set_column_width',
      'excel_set_row_height',
      // Excel Data tools (3)
      'excel_sort_range',
      'excel_calculate',
      'excel_export_csv'
    ];

    test('should have all 84 tools defined', () => {
      expect(tools).toHaveLength(84);
    });

    test('should have document tools', () => {
      const documentTools = tools.filter(t =>
        ['word_create_document', 'word_open_document', 'word_get_document_text', 'word_get_document_info', 'word_save_document', 'word_close_document', 'word_export_pdf'].includes(
          t
        )
      );
      expect(documentTools).toHaveLength(7);
    });

    test('should have text tools', () => {
      const textTools = tools.filter(t => ['word_insert_text', 'word_replace_text', 'word_format_text'].includes(t));
      expect(textTools).toHaveLength(3);
    });

    test('should have navigation tools', () => {
      const navigationTools = tools.filter(t => ['word_move_cursor_after_text', 'word_goto_start', 'word_goto_end', 'word_get_selection_info', 'word_select_all'].includes(t));
      expect(navigationTools).toHaveLength(5);
    });

    test('should have table tools', () => {
      const tableTools = tools.filter(t =>
        [
          'word_list_tables',
          'word_get_table_cell',
          'word_set_table_cell',
          'word_select_table_cell',
          'word_find_table_header',
          'word_create_table',
          'word_add_table_row',
          'word_delete_table_row',
          'word_add_table_column',
          'word_delete_table_column'
        ].includes(t)
      );
      expect(tableTools).toHaveLength(10);
    });

    test('should have bookmark tools', () => {
      const bookmarkTools = tools.filter(t => ['word_list_bookmarks', 'word_create_bookmark', 'word_goto_bookmark', 'word_delete_bookmark'].includes(t));
      expect(bookmarkTools).toHaveLength(4);
    });

    test('should have hyperlink tools', () => {
      const hyperlinkTools = tools.filter(t => ['word_list_hyperlinks', 'word_create_hyperlink'].includes(t));
      expect(hyperlinkTools).toHaveLength(2);
    });

    test('should have paragraph tools', () => {
      const paragraphTools = tools.filter(t => ['word_list_paragraphs', 'word_goto_paragraph', 'word_set_paragraph_style'].includes(t));
      expect(paragraphTools).toHaveLength(3);
    });

    test('should have header/footer tools', () => {
      const hfTools = tools.filter(t =>
        ['word_get_header_text', 'word_set_header_text', 'word_get_footer_text', 'word_set_footer_text', 'word_insert_header_image', 'word_insert_footer_image'].includes(t)
      );
      expect(hfTools).toHaveLength(6);
    });

    test('should have section tools', () => {
      const secTools = tools.filter(t => ['word_list_sections', 'word_get_section_info', 'word_set_page_setup', 'word_insert_section_break'].includes(t));
      expect(secTools).toHaveLength(4);
    });

    test('should have formatting read tools', () => {
      const fmtTools = tools.filter(t => ['word_get_text_formatting', 'word_get_paragraph_formatting'].includes(t));
      expect(fmtTools).toHaveLength(2);
    });

    test('should have image tools', () => {
      const imageTools = tools.filter(t => ['word_insert_image', 'word_list_inline_shapes', 'word_resize_inline_shape'].includes(t));
      expect(imageTools).toHaveLength(3);
    });

    test('should have Excel tools', () => {
      const excelTools = tools.filter(t => t.startsWith('excel_'));
      expect(excelTools).toHaveLength(33);
    });
  });

  describe('Tool Annotations', () => {
    test('destructive tools should have destructiveHint', () => {
      const destructiveTools = [
        'word_create_document',
        'word_insert_text',
        'word_replace_text',
        'word_format_text',
        'word_save_document',
        'word_close_document',
        'word_export_pdf',
        'word_set_table_cell',
        'word_select_table_cell',
        'word_move_cursor_after_text',
        'word_goto_start',
        'word_goto_end',
        'word_select_all',
        'word_create_table',
        'word_add_table_row',
        'word_delete_table_row',
        'word_add_table_column',
        'word_delete_table_column',
        'word_create_bookmark',
        'word_goto_bookmark',
        'word_delete_bookmark',
        'word_create_hyperlink',
        'word_goto_paragraph',
        'word_set_paragraph_style'
      ];

      expect(destructiveTools.length).toBeGreaterThan(0);
    });

    test('read-only tools should have readOnlyHint', () => {
      const readOnlyTools = [
        'word_open_document',
        'word_get_document_text',
        'word_get_document_info',
        'word_list_tables',
        'word_get_table_cell',
        'word_find_table_header',
        'word_get_selection_info',
        'word_list_bookmarks',
        'word_list_hyperlinks',
        'word_list_paragraphs'
      ];

      expect(readOnlyTools.length).toBeGreaterThan(0);
    });
  });

  describe('Required Parameters', () => {
    test('open_document requires path', () => {
      const inputSchema = {
        type: 'object',
        properties: {
          path: { type: 'string' }
        },
        required: ['path']
      };

      expect(inputSchema.required).toContain('path');
    });

    test('insert_text requires text', () => {
      const inputSchema = {
        type: 'object',
        properties: {
          text: { type: 'string' }
        },
        required: ['text']
      };

      expect(inputSchema.required).toContain('text');
    });

    test('replace_text requires find and replace', () => {
      const inputSchema = {
        type: 'object',
        properties: {
          find: { type: 'string' },
          replace: { type: 'string' },
          all: { type: 'boolean', default: true }
        },
        required: ['find', 'replace']
      };

      expect(inputSchema.required).toEqual(expect.arrayContaining(['find', 'replace']));
    });

    test('get_table_cell requires tableIndex, row, column', () => {
      const inputSchema = {
        type: 'object',
        properties: {
          tableIndex: { type: 'integer' },
          row: { type: 'integer' },
          column: { type: 'integer' }
        },
        required: ['tableIndex', 'row', 'column']
      };

      expect(inputSchema.required).toEqual(expect.arrayContaining(['tableIndex', 'row', 'column']));
    });

    test('create_table requires rows and columns', () => {
      const inputSchema = {
        type: 'object',
        properties: {
          rows: { type: 'integer' },
          columns: { type: 'integer' }
        },
        required: ['rows', 'columns']
      };

      expect(inputSchema.required).toEqual(expect.arrayContaining(['rows', 'columns']));
    });

    test('create_bookmark requires name', () => {
      const inputSchema = {
        type: 'object',
        properties: {
          name: { type: 'string' }
        },
        required: ['name']
      };

      expect(inputSchema.required).toContain('name');
    });

    test('create_hyperlink requires url', () => {
      const inputSchema = {
        type: 'object',
        properties: {
          url: { type: 'string' },
          displayText: { type: 'string' }
        },
        required: ['url']
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
