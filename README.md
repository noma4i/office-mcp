# Word MCP Server

MCP server for controlling Microsoft Word via AppleScript on macOS.

## Overview

This MCP (Model Context Protocol) server provides Claude with the ability to interact with Microsoft Word on macOS through AppleScript automation. It enables document creation, editing, formatting, and management directly from Claude conversations.

## Features

- **34 tools** for comprehensive Word automation
- **Document operations**: Create, open, save, close, export to PDF
- **Text manipulation**: Insert, replace, format text with rich formatting options
- **Navigation**: Move cursor, select text, jump to bookmarks
- **Table management**: Create, modify, and query tables
- **Bookmarks**: Create, navigate, and manage bookmarks
- **Hyperlinks**: Create and list hyperlinks
- **Paragraphs**: List, navigate, and style paragraphs
- **AppleScript automation**: Direct control of Word through macOS
- **macOS only**: Requires macOS and Microsoft Word

## Installation

```bash
npm install
```

### MCP Configuration

Add to your Claude Desktop configuration file:

**macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "word": {
      "command": "node",
      "args": ["/path/to/word_mcp/server/index.js"]
    }
  }
}
```

## Tools

### Documents (7 tools)
- `create_document` - Create new document with optional content
- `open_document` - Open existing document
- `get_document_text` - Retrieve all text from document
- `get_document_info` - Get statistics (words, characters, pages)
- `save_document` - Save document to specified path
- `close_document` - Close document with optional save
- `export_pdf` - Export document as PDF

### Text (3 tools)
- `insert_text` - Insert text at cursor position
- `replace_text` - Find and replace text
- `format_text` - Apply formatting (bold, italic, underline, font, size)

### Navigation (5 tools)
- `move_cursor_after_text` - Move cursor after found text
- `goto_start` - Jump to document start
- `goto_end` - Jump to document end
- `get_selection_info` - Get selection position and length
- `select_all` - Select entire document

### Tables (10 tools)
- `list_tables` - List all tables with dimensions
- `get_table_cell` - Get text from specific cell
- `set_table_cell` - Set text in specific cell
- `select_table_cell` - Move cursor to cell
- `find_table_header` - Find column by header text
- `create_table` - Create table at cursor
- `add_table_row` - Add row to table
- `delete_table_row` - Delete row from table
- `add_table_column` - Add column to table
- `delete_table_column` - Delete column from table

### Bookmarks (4 tools)
- `list_bookmarks` - List all bookmarks
- `create_bookmark` - Create bookmark at selection
- `goto_bookmark` - Navigate to bookmark
- `delete_bookmark` - Delete bookmark

### Hyperlinks (2 tools)
- `list_hyperlinks` - List all hyperlinks
- `create_hyperlink` - Create hyperlink at selection

### Paragraphs (3 tools)
- `list_paragraphs` - List paragraphs with styles
- `goto_paragraph` - Navigate to paragraph by index
- `set_paragraph_style` - Apply style to paragraph

## Requirements

- **macOS** (AppleScript only available on macOS)
- **Microsoft Word** installed and accessible
- **Node.js** >= 16.0.0

## Indexing

All indexes are **1-based** (tableIndex=1, row=1, column=1, index=1).

## Testing

The project includes comprehensive test coverage with Jest.

### Running Tests

```bash
npm test              # Run all tests
npm run test:watch    # Watch mode
npm run test:coverage # With coverage report
```

### Test Coverage

- **3 test suites** with **79 tests**
- **738 lines** of test code
- Tests cover:
  - Input validation functions
  - All 34 MCP tools
  - Server configuration and integration
  - Error handling scenarios
  - AppleScript safety

## License

MIT - see [LICENSE](LICENSE) file

## Credits

**Current Maintainer**: [noma4i](https://github.com/noma4i)

**Original Author**: [Anthropic](https://www.anthropic.com)

This project is based on Anthropic's Word extension and has been modified and extended with additional features.
