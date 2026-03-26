# Office MCP Server

An MCP server that lets Claude control Microsoft Word and Excel on macOS via AppleScript.

**98 tools** - 59 for Word, 39 for Excel - covering documents, text, tables, formatting, navigation, clipboard, headers/footers, sections, images, and more.

## Requirements

- macOS with Microsoft Word and/or Excel installed
- Node.js >= 16
- Yarn

## Quick Start

```bash
yarn install
yarn build
yarn start
```

### Claude Desktop Configuration

Add to your Claude Desktop config:

```json
{
  "mcpServers": {
    "office": {
      "command": "node",
      "args": ["/absolute/path/to/office_mcp/dist/index.js"]
    }
  }
}
```

### MCPB Bundle

```bash
yarn build && npx @anthropic-ai/mcpb pack . office-mcp.mcpb
```

One-click install for Claude Desktop.

## Word Tools (59)

| Category            | Tools                                                | Examples                                           |
| ------------------- | ---------------------------------------------------- | -------------------------------------------------- |
| Documents (7)       | create, open, save, close, export PDF, get text/info | `word_create_document`, `word_export_pdf`          |
| Text (4)            | insert, replace, delete, format                      | `word_insert_text`, `word_replace_text`            |
| Navigation (5)      | cursor movement, selection, goto start/end           | `word_move_cursor_after_text`, `word_goto_start`   |
| Tables (10)         | CRUD rows/columns/cells, find headers                | `word_create_table`, `word_set_table_cell`         |
| Paragraphs (4)      | list, goto, style, delete                            | `word_list_paragraphs`, `word_set_paragraph_style` |
| Bookmarks (4)       | list, create, goto, delete                           | `word_create_bookmark`, `word_goto_bookmark`       |
| Hyperlinks (2)      | list, create                                         | `word_list_hyperlinks`, `word_create_hyperlink`    |
| Images (4)          | insert, create ref, list shapes, resize              | `word_insert_image`, `word_resize_inline_shape`    |
| Clipboard (4)       | copy, paste, capture ref, insert ref                 | `word_copy_content`, `word_paste_content`          |
| Workflows (3)       | copy/clear/set story content                         | `word_copy_story_content`, `word_set_story_text`   |
| Headers/Footers (6) | get/set text, insert images                          | `word_get_header_text`, `word_insert_header_image` |
| Sections (4)        | list, info, page setup, breaks                       | `word_list_sections`, `word_set_page_setup`        |
| Formatting Read (2) | text formatting, paragraph formatting                | `word_get_text_formatting`                         |

## Excel Tools (39)

| Category         | Tools                                        | Examples                                          |
| ---------------- | -------------------------------------------- | ------------------------------------------------- |
| Workbooks (6)    | create, open, save, close, info, list        | `excel_create_workbook`, `excel_list_workbooks`   |
| Sheets (6)       | list, create, delete, rename, activate, info | `excel_create_sheet`, `excel_rename_sheet`        |
| Cells (7)        | get/set value, range, formula, clear, find   | `excel_get_cell`, `excel_set_cell_formula`        |
| Formatting (5)   | font, number format, color, merge, autofit   | `excel_format_cells`, `excel_set_cell_color`      |
| Rows/Columns (6) | insert/delete rows/columns, width, height    | `excel_insert_rows`, `excel_set_column_width`     |
| Data (3)         | sort, calculate, export CSV                  | `excel_sort_range`, `excel_export_csv`            |
| Clipboard (4)    | copy, paste, capture ref, insert ref         | `excel_copy_range`, `excel_paste_range`           |
| Workflows (2)    | clear worksheet, set range values            | `excel_clear_worksheet`, `excel_set_range_values` |

## License

MIT - see [LICENSE](LICENSE)
