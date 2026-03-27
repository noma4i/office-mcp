# Changelog

## v0.1.1

- Fix word_replace_text -1708 fallback not loaded in dist/
- Fix CI: Node 22 only, yarn install, immutable installs
- Add MCPB artifact to release workflow

## v0.1.0

Initial public release.

- 98 tools: 59 for Word, 39 for Excel
- Word: documents, text, tables, bookmarks, hyperlinks, paragraphs, navigation, images, clipboard, workflows, headers/footers, sections, formatting read
- Excel: workbooks, sheets, cells, formatting, rows/columns, data, clipboard, workflows
- Rich object clipboard - copy/paste with full formatting
- Fragment store for temporary content refs
- Word Find orchestration with compatibility fallback
- AppleScript syntax validation via osacompile
- MCPB packaging for Claude Desktop
