# Office MCP Server

MCP-сервер для автоматизации Microsoft Word и Microsoft Excel через AppleScript на macOS.

## Что поддерживается

- 86 инструментов: 53 для Word и 33 для Excel
- Операции с документами и книгами: create/open/save/close/export
- Текст, таблицы, закладки, навигация, headers/footers, секции в Word
- Листы, ячейки, формулы, форматирование, сортировка и экспорт CSV в Excel
- Проверка AppleScript-синтаксиса тестами

## Требования

- macOS
- Microsoft Word и Microsoft Excel
- Node.js >= 16
- Yarn

## Установка

```bash
yarn install
```

## Запуск

```bash
yarn build
yarn start
```

Для разработки:

```bash
yarn build:watch
```

## MCP-конфигурация (Claude Desktop)

```json
{
  "mcpServers": {
    "office": {
      "command": "node",
      "args": ["/absolute/path/to/word_mcp/dist/index.js"]
    }
  }
}
```

## Формат ответов инструментов

Каждый tool call возвращает JSON в `content[0].text`.

- Успех:
  - `{"ok": true, "message": "...", "data": ...}`
- Ошибка:
  - `{"ok": false, "error": {"code": "...", "message": "...", "details": ...}}`

Примеры кодов ошибок: `NO_DOCUMENT_OPEN`, `NO_WORKBOOK_OPEN`, `NOT_FOUND`, `OUT_OF_RANGE`, `VALIDATION_ERROR`, `APPSCRIPT_ERROR`.

## Тесты

```bash
yarn test
yarn test:coverage
```

## Сборка MCPB

```bash
yarn build && npx @anthropic-ai/mcpb pack . office-mcp.mcpb
```

## Структура

- `src/` — исходники сервера и инструментов
- `dist/` — собранная копия `src/`
- `tests/` — unit/integration/syntax-тесты
- `manifest.json` — описание MCP-пакета

