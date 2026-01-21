# Microsoft Word MCP Server

> **Правила документации:**
> - Максимум 8-10 тысяч токенов, БЕЗ примеров кода
> - Только текущее состояние, БЕЗ истории изменений
> - Таблицы: файл → назначение → API

## Overview

- **Цель**: MCP сервер для управления Microsoft Word через AppleScript
- **Принцип**: Предоставляет инструменты для CRUD операций с документами Word, включая работу с таблицами, закладками, гиперссылками и параграфами
- **Автор**: noma4i (github.com/noma4i)
- **Оригинал**: Based on Anthropic's Word extension

## Файлы

| Файл | Назначение | API |
|------|------------|-----|
| `server/index.js` | MCP сервер v0.6.0 | 34 инструмента (см. ниже) |
| `package.json` | Конфигурация npm, тестовые скрипты | — |
| `jest.config.js` | Конфигурация Jest | — |
| `tests/validation.test.js` | Тесты валидации входных данных | 5 тест-групп |
| `tests/mcp-tools.test.js` | Тесты всех 34 инструментов MCP | 9 тест-групп |
| `tests/server-integration.test.js` | Интеграционные тесты сервера | 8 тест-групп |

## Инструменты (34 шт.)

### Документы

| Инструмент | Назначение | Параметры |
|------------|------------|-----------|
| `create_document` | Создать новый документ | `content?` |
| `open_document` | Открыть документ | `path` |
| `get_document_text` | Получить весь текст | — |
| `get_document_info` | Статистика (слова, символы, страницы) | — |
| `save_document` | Сохранить документ | `path?` |
| `close_document` | Закрыть документ | `save?` |
| `export_pdf` | Экспорт в PDF | `path` |

### Текст

| Инструмент | Назначение | Параметры |
|------------|------------|-----------|
| `insert_text` | Вставить текст в курсор | `text` |
| `replace_text` | Найти и заменить | `find`, `replace`, `all?` |
| `format_text` | Форматирование выделения | `bold?`, `italic?`, `underline?`, `font?`, `size?` |

### Навигация

| Инструмент | Назначение | Параметры |
|------------|------------|-----------|
| `move_cursor_after_text` | Курсор после найденного текста | `searchText`, `occurrence?` |
| `goto_start` | Курсор в начало документа | — |
| `goto_end` | Курсор в конец документа | — |
| `get_selection_info` | Позиция и длина выделения | — |
| `select_all` | Выделить весь документ | — |

### Таблицы

| Инструмент | Назначение | Параметры |
|------------|------------|-----------|
| `list_tables` | Список таблиц с размерами | — |
| `get_table_cell` | Получить текст ячейки | `tableIndex`, `row`, `column` |
| `set_table_cell` | Установить текст в ячейку | `tableIndex`, `row`, `column`, `text` |
| `select_table_cell` | Переместить курсор в ячейку | `tableIndex`, `row`, `column` |
| `find_table_header` | Найти колонку по заголовку | `tableIndex`, `headerText`, `headerRow?` |
| `create_table` | Создать таблицу в курсоре | `rows`, `columns` |
| `add_table_row` | Добавить строку | `tableIndex`, `afterRow?` |
| `delete_table_row` | Удалить строку | `tableIndex`, `row` |
| `add_table_column` | Добавить колонку | `tableIndex`, `afterColumn?` |
| `delete_table_column` | Удалить колонку | `tableIndex`, `column` |

### Закладки

| Инструмент | Назначение | Параметры |
|------------|------------|-----------|
| `list_bookmarks` | Список закладок | — |
| `create_bookmark` | Создать закладку на выделении | `name` |
| `goto_bookmark` | Перейти к закладке | `name` |
| `delete_bookmark` | Удалить закладку | `name` |

### Гиперссылки

| Инструмент | Назначение | Параметры |
|------------|------------|-----------|
| `list_hyperlinks` | Список гиперссылок | — |
| `create_hyperlink` | Создать гиперссылку | `url`, `displayText?` |

### Параграфы

| Инструмент | Назначение | Параметры |
|------------|------------|-----------|
| `list_paragraphs` | Список параграфов со стилями | `limit?` |
| `goto_paragraph` | Перейти к параграфу | `index` |
| `set_paragraph_style` | Установить стиль параграфа | `index`, `styleName` |

## Индексация

- Все индексы **1-based** (tableIndex=1, row=1, column=1, index=1)
- `get_table_cell` автоматически удаляет маркеры ячеек (ASCII 7, 13)

## AppleScript синтаксис (Word)

| Операция | Синтаксис |
|----------|-----------|
| Получить ячейку | `cell COLUMN of row ROW of table` |
| Collapse to start | `set selection end of selection to selection start of selection` |
| Collapse to end | `set selection start of selection to selection end of selection` |
| Find text | `execute find findObj` + проверка `selStart ≠ selEnd` |
| Статистика | `compute statistics d statistic statistic words` |
| Создать таблицу | `make new table at selection with properties {number of rows:N, number of columns:M}` |
| Добавить строку | `insert rows below row N of table` |
| Удалить строку | `delete row N of table` |
| Закладки | `make new bookmark at d with properties {name:"X", bookmark range:selection}` |
| Гиперссылки | `make new hyperlink at selection with properties {hyperlink address:"URL"}` |
| Параграф | `paragraph N of d`, `paragraph style of paragraph N` |

## Тестирование

### Запуск тестов

```bash
npm test              # Запуск всех тестов
npm run test:watch    # Режим watch
npm run test:coverage # С отчетом покрытия
```

### Структура тестов

| Тестовый файл | Покрытие | Тестов |
|---------------|----------|--------|
| `validation.test.js` | Валидация входных данных | 20+ |
| `mcp-tools.test.js` | Все 34 инструмента MCP | 80+ |
| `server-integration.test.js` | Интеграция сервера | 30+ |

### Покрытие кода

- **Branches**: 80%
- **Functions**: 80%
- **Lines**: 80%
- **Statements**: 80%

## Связи

- Импорт: `@modelcontextprotocol/sdk`
- Экспорт: MCP tools через stdio transport
- Тесты: `@jest/globals`, моки для AppleScript
