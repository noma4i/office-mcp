# Office MCP Server

MCP-сервер для автоматизации Microsoft Word и Microsoft Excel через AppleScript на macOS.

## Что поддерживается

- 98 инструментов: 59 для Word и 39 для Excel
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

- `word_capture_content_ref` и `excel_capture_range_ref` переведены в legacy/disabled под in-place policy
- `word_create_image_ref` создаёт `ref` для локального изображения
- `word_insert_content_ref` продолжает работать для image refs; native Word/Excel fragment refs не являются основным workflow path
- `ref` является opaque-идентификатором и истекает автоматически

## Тесты

```bash
yarn test
yarn test:coverage
yarn test:applescript:strict
yarn test:word-find:live
```

`test:applescript:strict` запускает строгую компиляцию AppleScript через `osacompile`.

`test:word-find:live` запускает opt-in runtime smoke suite против реального Microsoft Word для Word Find/Replace/Delete/Move сценариев, включая placeholder-like replace кейсы с Unicode punctuation и длинными API replacement строками. Этот набор нужно запускать в локальной GUI-сессии; он не входит в обычный `yarn test`.

Word Find runtime теперь проходит через общий orchestration layer: primary path использует direct `execute find ... find text ...`, а compatibility fallback на legacy `set content of find object` включается только после runtime-ошибки `execute find` в Microsoft Word.

Дополнительно этот набор теперь включает registry-level проверку: тест проходит по всему `ALL_TOOLS`, генерирует минимально валидный AppleScript для каждого AppleScript-backed инструмента и компилирует его в strict-режиме.

## Сборка MCPB

```bash
yarn build && npx @anthropic-ai/mcpb pack . office-mcp.mcpb
```

## Структура

- `src/` — исходники сервера и инструментов
- `dist/` — собранная копия `src/`
- `tests/` — unit/integration/syntax-тесты
- `manifest.json` — описание MCP-пакета
