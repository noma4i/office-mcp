# Microsoft Office MCP Server

> **Правила документации:**
>
> - Максимум 8-10 тысяч токенов, БЕЗ примеров кода
> - Только текущее состояние, БЕЗ истории изменений
> - Таблицы: файл → назначение → API

## Overview

- **Цель**: MCP сервер для управления Microsoft Word и Excel через AppleScript
- **Принцип**: Модульная архитектура с разделением инструментов по категориям
- **Версия**: 0.8.0
- **Автор**: noma4i (github.com/noma4i)
- **Пакет**: `office-mcp`

## Архитектура

### Структура проекта

```
/
├── src/
│   ├── index.js              # Точка входа
│   ├── lib/
│   │   ├── server.js         # MCP Server setup (Microsoft-Office-Server)
│   │   ├── tool-registry.js  # Регистрация 70 инструментов (37 Word + 33 Excel)
│   │   ├── tool-executor.js  # Обработка CallTool requests
│   │   ├── validators.js     # Функции валидации
│   │   └── applescript/
│   │       ├── executor.js   # Выполнение AppleScript (таймаут 30с)
│   │       ├── helpers.js    # Общие AppleScript фрагменты (Word + Excel)
│   │       └── template-engine.js # Шаблонизатор
│   └── tools/
│       ├── documents.js        # Word: 7 инструментов
│       ├── text.js             # Word: 3 инструмента
│       ├── tables.js           # Word: 10 инструментов
│       ├── bookmarks.js        # Word: 4 инструмента
│       ├── hyperlinks.js       # Word: 2 инструмента
│       ├── paragraphs.js       # Word: 3 инструмента
│       ├── navigation.js       # Word: 5 инструментов
│       ├── images.js           # Word: 3 инструмента
│       ├── excel-workbooks.js  # Excel: 6 инструментов
│       ├── excel-sheets.js     # Excel: 6 инструментов
│       ├── excel-cells.js      # Excel: 7 инструментов
│       ├── excel-formatting.js # Excel: 5 инструментов
│       ├── excel-rows-columns.js # Excel: 6 инструментов
│       └── excel-data.js       # Excel: 3 инструмента
├── dist/                     # Собранные файлы (копия src/)
├── scripts/build.js          # Скрипт сборки
├── tests/                    # Тесты
├── package.json              # office-mcp v0.8.0
└── yarn.lock
```

### Основные модули

| Модуль                                   | Назначение                                     | API                                                     |
| ---------------------------------------- | ---------------------------------------------- | ------------------------------------------------------- |
| `src/index.js`                           | Точка входа                                    | `main()`                                                |
| `src/lib/server.js`                      | MCP Server v0.8.0                              | `createServer()`, `startServer()`                       |
| `src/lib/tool-registry.js`               | Регистрация 70 инструментов                    | `ALL_TOOLS`, `getToolDefinitions()`, `getToolHandler()` |
| `src/lib/tool-executor.js`               | Обработчик инструментов                        | `executeTool()`                                         |
| `src/lib/validators.js`                  | Валидация (`validateInteger` через `Number()`) | 5 функций                                               |
| `src/lib/applescript/executor.js`        | Выполнение AppleScript (таймаут 30с)           | `runAppleScript()`                                      |
| `src/lib/applescript/helpers.js`         | Фрагменты Word + Excel                         | `COMMON_SCRIPTS`                                        |
| `src/lib/applescript/template-engine.js` | Шаблоны (regex-safe, type-safe)                | `processTemplate()`                                     |

### Нейминг инструментов

- **Word**: префикс `word_` (`word_create_document`, `word_insert_text`, `word_list_tables`)
- **Excel**: префикс `excel_` (`excel_create_workbook`, `excel_set_cell`, `excel_sort_range`)

## Инструменты Word (37 шт.)

### Документы (7)

| Инструмент               | Назначение       | Параметры  |
| ------------------------ | ---------------- | ---------- |
| `word_create_document`   | Создать документ | `content?` |
| `word_open_document`     | Открыть документ | `path`     |
| `word_get_document_text` | Получить текст   | —          |
| `word_get_document_info` | Статистика       | —          |
| `word_save_document`     | Сохранить        | `path?`    |
| `word_close_document`    | Закрыть          | `save?`    |
| `word_export_pdf`        | Экспорт в PDF    | `path`     |

### Текст (3)

| Инструмент          | Назначение       | Параметры                                          |
| ------------------- | ---------------- | -------------------------------------------------- |
| `word_insert_text`  | Вставить текст   | `text`                                             |
| `word_replace_text` | Найти и заменить | `find`, `replace`, `all?`                          |
| `word_format_text`  | Форматирование   | `bold?`, `italic?`, `underline?`, `font?`, `size?` |

### Навигация (5)

| Инструмент                    | Назначение          | Параметры                   |
| ----------------------------- | ------------------- | --------------------------- |
| `word_move_cursor_after_text` | Курсор после текста | `searchText`, `occurrence?` |
| `word_goto_start`             | В начало            | —                           |
| `word_goto_end`               | В конец             | —                           |
| `word_get_selection_info`     | Позиция выделения   | —                           |
| `word_select_all`             | Выделить всё        | —                           |

### Таблицы (10)

| Инструмент                 | Назначение        | Параметры                                |
| -------------------------- | ----------------- | ---------------------------------------- |
| `word_list_tables`         | Список таблиц     | —                                        |
| `word_get_table_cell`      | Значение ячейки   | `tableIndex`, `row`, `column`            |
| `word_set_table_cell`      | Установить ячейку | `tableIndex`, `row`, `column`, `text`    |
| `word_select_table_cell`   | Курсор в ячейку   | `tableIndex`, `row`, `column`            |
| `word_find_table_header`   | Найти заголовок   | `tableIndex`, `headerText`, `headerRow?` |
| `word_create_table`        | Создать таблицу   | `rows`, `columns`                        |
| `word_add_table_row`       | Добавить строку   | `tableIndex`, `afterRow?`                |
| `word_delete_table_row`    | Удалить строку    | `tableIndex`, `row`                      |
| `word_add_table_column`    | Добавить колонку  | `tableIndex`, `afterColumn?`             |
| `word_delete_table_column` | Удалить колонку   | `tableIndex`, `column`                   |

### Закладки (4), Гиперссылки (2), Параграфы (3), Изображения (3)

| Инструмент                 | Назначение                             | Параметры                                        |
| -------------------------- | -------------------------------------- | ------------------------------------------------ |
| `word_list_bookmarks`      | Список закладок                        | —                                                |
| `word_create_bookmark`     | Создать закладку                       | `name`                                           |
| `word_goto_bookmark`       | Перейти                                | `name`                                           |
| `word_delete_bookmark`     | Удалить                                | `name`                                           |
| `word_list_hyperlinks`     | Список (try/catch для text to display) | —                                                |
| `word_create_hyperlink`    | Создать                                | `url`, `displayText?`                            |
| `word_list_paragraphs`     | Список со стилями                      | `limit?`                                         |
| `word_goto_paragraph`      | Перейти                                | `index`                                          |
| `word_set_paragraph_style` | Установить стиль                       | `index`, `styleName`                             |
| `word_insert_image`        | Вставить через clipboard               | `path`, `width?`, `height?`                      |
| `word_list_inline_shapes`  | Список shapes                          | —                                                |
| `word_resize_inline_shape` | Изменить размер                        | `index`, `width?`, `height?`, `lockAspectRatio?` |

## Инструменты Excel (33 шт.)

### Workbooks (6)

| Инструмент                | Назначение           | Параметры |
| ------------------------- | -------------------- | --------- |
| `excel_create_workbook`   | Создать книгу        | —         |
| `excel_open_workbook`     | Открыть книгу        | `path`    |
| `excel_get_workbook_info` | Имя, путь, листы     | —         |
| `excel_save_workbook`     | Сохранить            | `path?`   |
| `excel_close_workbook`    | Закрыть              | `save?`   |
| `excel_list_workbooks`    | Список открытых книг | —         |

### Sheets (6)

| Инструмент             | Назначение          | Параметры                |
| ---------------------- | ------------------- | ------------------------ |
| `excel_list_sheets`    | Список листов       | —                        |
| `excel_create_sheet`   | Создать лист        | `name?`, `afterIndex?`   |
| `excel_delete_sheet`   | Удалить лист        | `nameOrIndex`            |
| `excel_rename_sheet`   | Переименовать       | `nameOrIndex`, `newName` |
| `excel_activate_sheet` | Переключиться       | `nameOrIndex`            |
| `excel_get_sheet_info` | Used range, размеры | `nameOrIndex?`           |

### Cells (7)

| Инструмент               | Назначение               | Параметры              |
| ------------------------ | ------------------------ | ---------------------- |
| `excel_get_cell`         | Значение ячейки          | `cell` (A1-нотация)    |
| `excel_set_cell`         | Установить значение      | `cell`, `value`        |
| `excel_get_range`        | Значения диапазона (TSV) | `range`                |
| `excel_set_cell_formula` | Установить формулу       | `cell`, `formula`      |
| `excel_clear_range`      | Очистить диапазон        | `range`                |
| `excel_get_used_range`   | Адрес и размеры          | —                      |
| `excel_find_cell`        | Найти текст              | `searchText`, `range?` |

### Formatting (5)

| Инструмент                | Назначение          | Параметры                                                   |
| ------------------------- | ------------------- | ----------------------------------------------------------- |
| `excel_format_cells`      | Шрифт, размер, цвет | `range`, `bold?`, `italic?`, `font?`, `size?`, `fontColor?` |
| `excel_set_number_format` | Числовой формат     | `range`, `format`                                           |
| `excel_set_cell_color`    | Цвет фона           | `range`, `color` [R,G,B]                                    |
| `excel_merge_cells`       | Объединить          | `range`                                                     |
| `excel_autofit`           | Авто-ширина колонок | `range`                                                     |

### Rows & Columns (6)

| Инструмент               | Назначение       | Параметры          |
| ------------------------ | ---------------- | ------------------ |
| `excel_insert_rows`      | Вставить строки  | `row`, `count?`    |
| `excel_delete_rows`      | Удалить строки   | `row`, `count?`    |
| `excel_insert_columns`   | Вставить колонки | `column`, `count?` |
| `excel_delete_columns`   | Удалить колонки  | `column`, `count?` |
| `excel_set_column_width` | Ширина колонки   | `column`, `width`  |
| `excel_set_row_height`   | Высота строки    | `row`, `height`    |

### Data (3)

| Инструмент         | Назначение      | Параметры                                      |
| ------------------ | --------------- | ---------------------------------------------- |
| `excel_sort_range` | Сортировка      | `range`, `keyCell`, `ascending?`, `hasHeader?` |
| `excel_calculate`  | Пересчёт формул | —                                              |
| `excel_export_csv` | Экспорт в CSV   | `path`                                         |

## Индексация

- **Word**: все индексы 1-based (tableIndex, row, column, index)
- **Excel**: ячейки в A1-нотации, листы по имени или 1-based индексу, строки/колонки 1-based

## AppleScript синтаксис

### Word

| Операция       | Синтаксис                                                                         |
| -------------- | --------------------------------------------------------------------------------- |
| Ячейка таблицы | `cell COLUMN of row ROW of table`                                                 |
| Collapse       | `set selection end/start of selection to selection start/end of selection`        |
| Find           | `execute find findObj`                                                            |
| Закладки       | `make new bookmark at d with properties {name:"X", \|bookmark range\|:selection}` |
| Гиперссылки    | `hyperlink objects of d`, `text to display of h` (в try/catch)                    |
| Изображение    | clipboard + `paste object selection`                                              |

### Excel

| Операция        | Синтаксис                                                         |
| --------------- | ----------------------------------------------------------------- |
| Ячейка          | `value of cell "A1" of ws`, `set value of cell "A1" of ws to V`   |
| Формула         | `set formula of cell "A1" of ws to "=SUM()"` (НЕ `formula value`) |
| Диапазон        | `range "A1:B2" of ws`                                             |
| Used range      | `used range of ws`, `get address of used range`                   |
| Лист            | `worksheet N of wb`, `worksheet "Name" of wb`                     |
| Создать лист    | `make new worksheet at end of wb`                                 |
| Сохранить       | `save workbook as wb filename PATH`                               |
| Открыть         | `open workbook workbook file name PATH`                           |
| Шрифт           | `bold of font object of cell`, `font size of font object of cell` |
| Цвет фона       | `color of interior object of cell to {R,G,B}`                     |
| Вставить строку | `insert into range (range "N:N" of ws) shift shift down`          |
| Удалить строку  | `delete range (range "N:N" of ws) shift shift up`                 |
| Сортировка      | `sort range ... key1 ... order1 sort ascending header header yes` |
| Autofit         | `set r to entire column of range "A:C" of ws` → `autofit r`       |
| Поиск           | `find searchRange what "text"`                                    |

## Тестирование

```bash
yarn test              # Все тесты
yarn test:watch        # Watch mode
yarn test:coverage     # С покрытием
```

| Тестовый файл                            | Покрытие                                   | Тестов |
| ---------------------------------------- | ------------------------------------------ | ------ |
| `tests/validation.test.js`               | Валидация                                  | 20+    |
| `tests/mcp-tools.test.js`                | Word 37 инструментов                       | 80+    |
| `tests/server-integration.test.js`       | Интеграция                                 | 30+    |
| `tests/applescript-syntax.test.js`       | Word AppleScript + аудит + template-engine | 52     |
| `tests/excel-applescript-syntax.test.js` | Excel AppleScript (все 33 инструмента)     | 42     |

## Связи

- Импорт: `@modelcontextprotocol/sdk`
- Экспорт: MCP tools через stdio transport
- Тесты: `@jest/globals`, моки для AppleScript
- Сборка: Yarn v4+, Node.js v16+
