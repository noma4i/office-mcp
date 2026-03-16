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
│   │   ├── tool-registry.js  # Регистрация 98 инструментов (59 Word + 39 Excel)
│   │   ├── tool-executor.js  # Обработка CallTool requests
│   │   ├── validators.js     # Функции валидации
│   │   ├── fragment-store.js # Временное хранилище rich-content refs
│   │   └── applescript/
│   │       ├── executor.js   # Выполнение AppleScript (таймаут 30с)
│   │       ├── helpers.js    # Общие AppleScript фрагменты (Word + Excel)
│   │       ├── word-find.js  # Word Find orchestration + compatibility fallback
│   │       └── template-engine.js # Шаблонизатор
│   └── tools/
│       ├── word-documents.js        # Word: 7 инструментов
│       ├── word-text.js             # Word: 4 инструмента (insert, replace, delete, format)
│       ├── word-tables.js           # Word: 10 инструментов
│       ├── word-bookmarks.js        # Word: 4 инструмента
│       ├── word-hyperlinks.js       # Word: 2 инструмента
│       ├── word-paragraphs.js       # Word: 4 инструмента (list, goto, style, delete)
│       ├── word-navigation.js       # Word: 5 инструментов
│       ├── word-images.js           # Word: 4 инструмента
│       ├── word-clipboard.js        # Word: 4 инструмента (copy, capture ref, insert ref, paste)
│       ├── word-workflows.js        # Word: 3 инструмента (copy/clear/set story content)
│       ├── word-headers-footers.js  # Word: 6 инструментов (get/set header/footer, insert images)
│       ├── word-sections.js         # Word: 4 инструмента (list, info, page setup, break)
│       ├── word-formatting-read.js  # Word: 2 инструмента (text formatting, paragraph formatting)
│       ├── excel-workbooks.js  # Excel: 6 инструментов
│       ├── excel-sheets.js     # Excel: 6 инструментов
│       ├── excel-cells.js      # Excel: 7 инструментов
│       ├── excel-formatting.js # Excel: 5 инструментов
│       ├── excel-rows-columns.js # Excel: 6 инструментов
│       ├── excel-data.js       # Excel: 3 инструмента
│       ├── excel-clipboard.js  # Excel: 4 инструмента (copy, paste, capture ref, insert ref)
│       └── excel-workflows.js  # Excel: 2 инструмента (clear worksheet, set range values)
├── dist/                     # Собранные файлы (копия src/)
├── scripts/build.js          # Скрипт сборки
├── tests/                    # Тесты
├── package.json              # office-mcp v0.8.0
└── yarn.lock
```

### Основные модули

| Модуль                                   | Назначение                                     | API                                                              |
| ---------------------------------------- | ---------------------------------------------- | ---------------------------------------------------------------- |
| `src/index.js`                           | Точка входа                                    | `main()`                                                         |
| `src/lib/server.js`                      | MCP Server v0.8.0                              | `createServer()`, `startServer()`                                |
| `src/lib/tool-registry.js`               | Регистрация 98 инструментов                    | `ALL_TOOLS`, `getToolDefinitions()`, `getToolHandler()`          |
| `src/lib/tool-executor.js`               | Обработчик инструментов + MCP envelope ошибок  | `executeTool()`                                                  |
| `src/lib/validators.js`                  | Валидация строк, чисел, enum и Excel refs      | 8 функций                                                        |
| `src/lib/fragment-store.js`              | Временные rich-content refs с TTL              | `reserveFragment()`, `commitReservedFragment()`, `getFragment()` |
| `src/lib/applescript/executor.js`        | Выполнение AppleScript (таймаут 30с)           | `runAppleScript()`                                               |
| `src/lib/applescript/helpers.js`         | Фрагменты Word + Excel + экранирование строк   | `COMMON_SCRIPTS`, `toAppleScriptString()`, `escapeForWordFind()`, `buildWordExecuteFind()` |
| `src/lib/applescript/word-find.js`       | Единый Word Find runner + direct/legacy fallback | `buildWordFindScript()`, `runWordFindWithFallback()`           |
| `src/lib/applescript/template-engine.js` | Шаблоны (regex-safe, type-safe)                | `processTemplate()`                                              |

### Нейминг инструментов

- **Word**: префикс `word_` (`word_create_document`, `word_insert_text`, `word_list_tables`)
- **Excel**: префикс `excel_` (`excel_create_workbook`, `excel_set_cell`, `excel_sort_range`)

## Инструменты Word (59 шт.)

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

### Текст (4)

| Инструмент          | Назначение                                                       | Параметры                                          |
| ------------------- | ---------------------------------------------------------------- | -------------------------------------------------- |
| `word_insert_text`  | Вставить текст                                                   | `text`                                             |
| `word_replace_text` | Найти и заменить (возвращает "not found" если не найдено)        | `find`, `replace?` (default ""), `all?`            |
| `word_delete_text`  | Удалить текст/выделение (возвращает "not found" если не найдено) | `text?` (если указан — найти и удалить все)        |
| `word_format_text`  | Форматирование                                                   | `bold?`, `italic?`, `underline?`, `font?`, `size?` |

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

### Закладки (4), Гиперссылки (2), Параграфы (4), Изображения (4)

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
| `word_delete_paragraph`    | Удалить параграф                       | `index`                                          |
| `word_insert_image`        | Вставить через clipboard               | `path`, `width?`, `height?`                      |
| `word_create_image_ref`    | Создать `ref` для локального изображения | `path`                                         |
| `word_list_inline_shapes`  | Список shapes                          | —                                                |
| `word_resize_inline_shape` | Изменить размер                        | `index`, `width?`, `height?`, `lockAspectRatio?` |

### Clipboard (4)

| Инструмент                  | Назначение                                   | Параметры                                                           |
| --------------------------- | -------------------------------------------- | ------------------------------------------------------------------- |
| `word_copy_content`         | Копировать в clipboard с форматами           | `scope?`, `startParagraph?`, `endParagraph?`, `inlineShapeIndex?`   |
| `word_capture_content_ref`  | Legacy disabled под in-place policy          | `scope?`, `startParagraph?`, `endParagraph?`, `inlineShapeIndex?`   |
| `word_insert_content_ref`   | Вставить image ref                           | `ref`, `width?`, `height?`                                           |
| `word_paste_content`        | Вставить из clipboard с форматами            | —                                                                   |

### Workflows (3)

| Инструмент                 | Назначение                                 | Параметры                           |
| -------------------------- | ------------------------------------------ | ----------------------------------- |
| `word_copy_story_content`  | Копировать body/header/footer в clipboard  | `scope?`, `section?`, `type?`       |
| `word_clear_story_content` | Очистить body/header/footer in-place       | `scope?`, `section?`, `type?`       |
| `word_set_story_text`      | Заменить текст body/header/footer in-place | `scope?`, `text`, `section?`, `type?` |

### Headers/Footers (6)

| Инструмент                 | Назначение                    | Параметры                                                       |
| -------------------------- | ----------------------------- | --------------------------------------------------------------- |
| `word_get_header_text`     | Получить текст header         | `section?` (default 1), `type?` (primary/first_page/even_pages) |
| `word_set_header_text`     | Установить текст header       | `text`, `section?`, `type?`                                     |
| `word_get_footer_text`     | Получить текст footer         | `section?`, `type?`                                             |
| `word_set_footer_text`     | Установить текст footer       | `text`, `section?`, `type?`                                     |
| `word_insert_header_image` | Вставить изображение в header | `path`, `section?`, `type?`, `width?`, `height?`                |
| `word_insert_footer_image` | Вставить изображение в footer | `path`, `section?`, `type?`, `width?`, `height?`                |

### Секции (4)

| Инструмент                  | Назначение                     | Параметры                                                                                                     |
| --------------------------- | ------------------------------ | ------------------------------------------------------------------------------------------------------------- |
| `word_list_sections`        | Список секций с page setup     | —                                                                                                             |
| `word_get_section_info`     | Детальная инфо о секции        | `index`                                                                                                       |
| `word_set_page_setup`       | Установить margins/orientation | `index?`, `topMargin?`, `bottomMargin?`, `leftMargin?`, `rightMargin?`, `orientation?`, `differentFirstPage?` |
| `word_insert_section_break` | Вставить разрыв секции         | `type?` (next_page/continuous/even_page/odd_page)                                                             |

### Чтение форматирования (2)

| Инструмент                      | Назначение                                  | Параметры |
| ------------------------------- | ------------------------------------------- | --------- |
| `word_get_text_formatting`      | Шрифт, размер, bold, italic, цвет выделения | —         |
| `word_get_paragraph_formatting` | Стиль, выравнивание, отступы, интервалы     | —         |

## Инструменты Excel (39 шт.)

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

| Инструмент               | Назначение                           | Параметры              |
| ------------------------ | ------------------------------------ | ---------------------- |
| `excel_get_cell`         | Значение ячейки                      | `cell` (A1-нотация)    |
| `excel_set_cell`         | Установить значение                  | `cell`, `value`        |
| `excel_get_range`        | Значения диапазона (TSV)             | `range`                |
| `excel_set_cell_formula` | Установить формулу (auto-prefix `=`) | `cell`, `formula`      |
| `excel_clear_range`      | Очистить диапазон                    | `range`                |
| `excel_get_used_range`   | Адрес и размеры                      | —                      |
| `excel_find_cell`        | Найти текст                          | `searchText`, `range?` |

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
| `excel_export_csv` | Экспорт листа в CSV | `path`, `worksheet?`                        |

### Clipboard (4)

| Инструмент                | Назначение                                  | Параметры                    |
| ------------------------- | ------------------------------------------- | ---------------------------- |
| `excel_copy_range`        | Копировать диапазон с форматами в clipboard | `range`, `worksheet?`        |
| `excel_paste_range`       | Вставить clipboard в target cell            | `targetCell`, `worksheet?`   |
| `excel_capture_range_ref` | Legacy disabled под in-place policy         | `range`, `worksheet?`        |
| `excel_insert_range_ref`  | Legacy disabled под in-place policy         | `ref`, `targetCell`, `worksheet?` |

### Workflows (2)

| Инструмент               | Назначение                              | Параметры                      |
| ------------------------ | --------------------------------------- | ------------------------------ |
| `excel_clear_worksheet`  | Очистить used range активного листа     | `worksheet?`                   |
| `excel_set_range_values` | Записать TSV-матрицу в диапазон in-place | `range`, `values`, `worksheet?` |
| `excel_export_csv` | Экспорт в CSV   | `path`                                         |

## Индексация

- **Word**: все индексы 1-based (tableIndex, row, column, index, section)
- **Excel**: ячейки в A1-нотации, листы по имени или 1-based индексу, строки/колонки 1-based

## Экранирование строк для AppleScript

**Правило:** `JSON.stringify` НЕ использовать для пользовательского текста. Хелперы в `src/lib/applescript/helpers.js`:

| Функция                        | Назначение                             | Пример вывода                  |
| ------------------------------ | -------------------------------------- | ------------------------------ |
| `toAppleScriptString(str)`     | Контент (insert, set cell, set header) | `("line1" & return & "line2")` |
| `escapeForWordFind(str)`       | Word Find/Replace, поиск               | `"line1^pline2"`               |
| `buildWordExecuteFind(...)`    | Прямой `execute find` с параметрами    | `execute find findObject ...`  |
| `escapeAppleScriptString(str)` | Экранирование `\` и `"` (базовый)      | `say \"hi\"`                   |
| `quoteAppleScriptString(str)`  | Базовый + оборачивание в кавычки       | `"say \"hi\""`                 |

**Контекст использования:**

- `toAppleScriptString` — newlines → `& return &` конкатенация. Для: `word_insert_text`, `word_create_document`, `word_set_table_cell`, `word_set_header/footer_text`, `excel_set_cell`
- `escapeForWordFind` — newlines → `^p` (Word paragraph mark). Для: `word_delete_text`, `word_replace_text`, `word_move_cursor_after_text` (ТОЛЬКО Word Find API)
- `buildWordExecuteFind` — генерирует primary-path вызов `execute find ... find text ... replace with ...`
- `src/lib/applescript/word-find.js` — единая точка выполнения Word Find. Direct `execute find ... find text ...` остаётся primary path; legacy `set content of find object` допустим только как compatibility fallback после runtime-dispatch ошибки (`-1708` / `doesn't understand execute find`)
- `quoteAppleScriptString` — экранирование `\` и `"` + оборачивание в кавычки. Для: `word_create_hyperlink` (URL, displayText), `word_find_table_header` (contains), `excel_find_cell` (what), `word_goto/delete_bookmark` (name), `word_move_cursor_after_text` (return строки), `excel_create/rename_sheet` (name), return-строки с пользовательским текстом
- `escapeAppleScriptString` — для вставки в уже кавычеченные строки. Для: `word_insert_header/footer_image` (hfsPath внутри `{file name:"..."}`)
- `JSON.stringify` — допустимо ТОЛЬКО для: путей (`open`, `save as`), cell refs (A1), range refs (A1:B3), font names, style names, формул (`set formula`)

## AppleScript синтаксис

### Обработка ошибок (ОБЯЗАТЕЛЬНО)

**Правило:** ВСЕ обращения к объектам Word/Excel, которые могут не существовать, ДОЛЖНЫ быть обёрнуты в `try/on error`:

- `font object of selection`, `paragraph format of selection` — selection может быть пустым
- `get header/footer of section` — header/footer type может быть недоступен
- `cell N of row M of table` — ячейка может не существовать
- `find object of selection` — может упасть без документа
- `name local of style of` — стиль может быть недоступен
- `worksheet N of wb`, `range "X" of ws` — лист/range может не существовать
- `inline shape N of d` — shape может не существовать

Паттерн: инициализировать fallback-значение → try → set → end try (или try → on error → return error msg → end try)

### Word

| Операция            | Синтаксис                                                                                   |
| ------------------- | ------------------------------------------------------------------------------------------- |
| Ячейка таблицы      | `cell COLUMN of row ROW of table`                                                           |
| Collapse            | `set selection end/start of selection to selection start/end of selection`                  |
| Find                | `find object of selection` (НЕ внутри `tell activeDoc`), primary path: `execute find ... find text ... replace with ...`; compatibility fallback на `set content of find object` допустим только внутри `src/lib/applescript/word-find.js` после runtime-dispatch ошибки |
| Закладки            | `make new bookmark at d with properties {name:"X", \|bookmark range\|:selection}`           |
| Гиперссылки         | `hyperlink objects of d`, `text to display of h` (в try/catch)                              |
| Copy/Paste          | `copy object selection`, `paste object selection`                                           |
| Изображение         | clipboard + `paste object selection`                                                        |
| Delete paragraph    | `select (text object of paragraph N of d)` → `delete (text object of selection)`            |
| Header              | `get header of section N of d index header footer primary`                                  |
| Footer              | `get footer of section N of d index header footer primary`                                  |
| Header/Footer текст | `content of text object of refHeader`                                                       |
| Section break       | `insert break at r break type section break next page`                                      |
| Page setup          | `page setup of section N of d` → margins, orientation                                       |
| Paragraph format    | `paragraph format left indent of pf` (НЕ `left indent of pf`)                               |

### Excel

| Операция        | Синтаксис                                                                              |
| --------------- | -------------------------------------------------------------------------------------- |
| Ячейка          | `value of cell "A1" of ws`, `set value of cell "A1" of ws to V`                        |
| Формула         | `set formula of cell "A1" of ws to "=SUM()"` (НЕ `formula value`)                      |
| Диапазон        | `range "A1:B2" of ws`                                                                  |
| Used range      | `used range of ws`, `get address of used range`                                        |
| Лист            | `worksheet N of wb`, `worksheet "Name" of wb`                                          |
| Создать лист    | `make new worksheet at end of wb`                                                      |
| Сохранить       | `save workbook as wb filename PATH`                                                    |
| Открыть         | `open workbook workbook file name PATH`                                                |
| Шрифт           | `bold of font object of cell`, `font size of font object of cell`                      |
| Цвет фона       | `color of interior object of cell to {R,G,B}`                                          |
| Вставить строку | `insert into range (range "N:N" of ws) shift shift down`                               |
| Удалить строку  | `delete range (range "N:N" of ws) shift shift up`                                      |
| Сортировка      | `sort range ... key1 ... order1 sort ascending header header yes`                      |
| Autofit         | `set r to entire column of range "A:C" of ws` → `autofit r`                            |
| Поиск           | `find searchRange what "text"` (в `try/on error` — выбрасывает ошибку если не найдено) |
| Display alerts  | try/finally: `set display alerts to false` → try → on error → restore → end try        |
| Clipboard copy  | `select range "A1:B2" of ws` + `System Events` `keystroke "c" using command down`      |
| Clipboard paste | `select range "C1" of ws` + `System Events` `keystroke "v" using command down`         |

## Тестирование

```bash
yarn test              # Все тесты
yarn test:watch        # Watch mode
yarn test:coverage     # С покрытием
yarn test:applescript:strict # Strict compile через osacompile
yarn test:word-find:live     # Opt-in runtime smoke для Word Find
```

| Тестовый файл                            | Покрытие                                                                                  | Тестов |
| ---------------------------------------- | ----------------------------------------------------------------------------------------- | ------ |
| `tests/validation.test.js`               | Валидация                                                                                 | 20+    |
| `tests/mcp-tools.test.js`                | Word инструменты                                                                          | 80+    |
| `tests/server-integration.test.js`       | Интеграция (98 инструментов)                                                               | 40+    |
| `tests/applescript-syntax.test.js`       | Word AppleScript + headers/sections/formatting + multiline + спецсимволы + error handling | 110+   |
| `tests/applescript-registry.test.js`     | Registry-level strict compile для всех AppleScript-backed tools через `ALL_TOOLS`         | 1      |
| `tests/excel-applescript-syntax.test.js` | Excel AppleScript (все 39 инструментов) + спецсимволы + валидация RGB + error handling     | 60+    |
| `tests/fragment-store.test.js`           | `ref`-хранилище, TTL cleanup, file-backed fragments                                      | 5+     |
| `tests/tool-executor.test.js`            | MCP envelope, коды ошибок, details                                                       | 5+     |
| `tests/applescript-wrappers.test.js`     | Word/Excel wrappers и guards                                                              | 4+     |
| `tests/word-find-orchestration.test.js`  | Word Find retry orchestration, compatibility fallback, combined errors                    | 6      |
| `tests/word-find-live.test.js`           | Opt-in runtime smoke против реального Microsoft Word для short/long/placeholder replace, delete, move | 6      |

- `yarn test:applescript:strict` проверяет синтаксис через `osacompile`, но не подтверждает runtime-поведение Word Find API
- `tests/word-find-orchestration.test.js` обязателен для любых изменений в `src/lib/applescript/word-find.js` и Word Find инструментах
- `yarn test:word-find:live` запускать только в локальной GUI-сессии с доступным Microsoft Word; suite должен покрывать short/long/placeholder replace, delete и move

## Сборка MCPB

```bash
yarn build && npx @anthropic-ai/mcpb pack . office-mcp.mcpb
```

- `.mcpbignore` — исключает src/, tests/, scripts/ и прочее из бандла
- Бандл содержит: `manifest.json`, `dist/`, `node_modules/`, `icon.png`
- Формат: ZIP-архив с расширением `.mcpb` для one-click установки в Claude Desktop

## Связи

- Импорт: `@modelcontextprotocol/sdk`
- Экспорт: MCP tools через stdio transport
- Тесты: `@jest/globals`, моки для AppleScript
- Сборка: Yarn v4+, Node.js v16+
