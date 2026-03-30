# Changelog

## v0.1.2

- Word Find/Replace: enforce 255-character limit with clear error message and alternative (paragraph tools)
- Tool schemas expose maxLength so agents see the limit before calling

## v0.1.1

- word_replace_text/word_delete_text: -1708 fallback (legacy find object content strategy) was missing from dist/ build
- CI: drop Node 18, test on Node 22 only, fix yarn install with immutable installs disabled
- Release workflow: add MCPB artifact upload

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
