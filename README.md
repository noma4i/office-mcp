# Office MCP Server

MCP-сервер для автоматизации Microsoft Word и Microsoft Excel через AppleScript на macOS.

## Что поддерживается

- 93 инструмента: 56 для Word и 37 для Excel
- Операции с документами и книгами: create/open/save/close/export
- Текст, таблицы, закладки, навигация, headers/footers, секции, rich copy/paste и `ref`-фрагменты в Word
- Листы, ячейки, формулы, форматирование, сортировка, clipboard/range refs и экспорт CSV в Excel
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
  - `{"ok": true, "message": "..."}` для строковых результатов
  - `{"ok": true, "message": "Operation completed successfully", "data": {...}}` для нестроковых результатов
- Ошибка:
  - `{"ok": false, "error": {"code": "...", "message": "...", "details": ...}}`

Примеры кодов ошибок: `NO_DOCUMENT_OPEN`, `NO_WORKBOOK_OPEN`, `NOT_FOUND`, `OUT_OF_RANGE`, `VALIDATION_ERROR`, `APPSCRIPT_ERROR`, `UNKNOWN_TOOL`.

Rich-content инструменты используют временные `ref`-хендлы:

- `word_capture_content_ref` и `excel_capture_range_ref` сохраняют native Office-фрагменты во временное хранилище и возвращают `data.ref`
- `word_create_image_ref` создаёт `ref` для локального изображения
- `word_insert_content_ref` и `excel_insert_range_ref` вставляют ранее сохранённые `ref`
- `ref` является opaque-идентификатором и истекает автоматически

## Тесты

```bash
yarn test
yarn test:coverage
yarn test:applescript:strict
```

`test:applescript:strict` запускает строгую компиляцию AppleScript через `osacompile`.

## Сборка MCPB

```bash
yarn build && npx @anthropic-ai/mcpb pack . office-mcp.mcpb
```

## Структура

- `src/` — исходники сервера и инструментов
- `dist/` — собранная копия `src/`
- `tests/` — unit/integration/syntax-тесты
- `manifest.json` — описание MCP-пакета
