# Microsoft Word MCP Server

> **Правила документации:**
>
> - Максимум 8-10 тысяч токенов, БЕЗ примеров кода
> - Только текущее состояние, БЕЗ истории изменений
> - Таблицы: файл → назначение → API

## Overview

- **Цель**: MCP сервер для управления Microsoft Word через AppleScript
- **Принцип**: Модульная архитектура с разделением инструментов по категориям
- **Версия**: 0.7.0 (модульный рефакторинг)
- **Автор**: noma4i (github.com/noma4i)
- **Оригинал**: Based on Anthropic's Word extension

## Архитектура (v0.7.0)

### Структура проекта

```
/
├── src/                      # Исходники
│   ├── index.js              # Точка входа (15 строк)
│   ├── lib/
│   │   ├── server.js         # MCP Server setup
│   │   ├── tool-registry.js  # Регистрация инструментов
│   │   ├── tool-executor.js  # Обработка CallTool requests
│   │   ├── validators.js     # Функции валидации
│   │   └── applescript/
│   │       ├── executor.js   # Выполнение AppleScript
│   │       ├── helpers.js    # Общие AppleScript функции
│   │       └── template-engine.js # Шаблонизатор
│   └── tools/
│       ├── documents.js      # 7 инструментов для документов
│       ├── text.js           # 3 инструмента для текста
│       ├── tables.js         # 10 инструментов для таблиц
│       ├── bookmarks.js      # 4 инструмента для закладок
│       ├── hyperlinks.js     # 2 инструмента для гиперссылок
│       ├── paragraphs.js     # 3 инструмента для параграфов
│       ├── navigation.js     # 5 инструментов для навигации
│       └── images.js         # 3 инструмента для изображений
├── dist/                     # Собранные файлы (копия src/)
├── scripts/
│   └── build.js              # Скрипт сборки
├── server/
│   └── index.old.js          # Старый монолитный файл (backup)
├── tests/                    # Тесты
├── .yarnrc.yml               # Yarn config (nodeLinker: node-modules)
├── package.json              # Yarn + build скрипты
└── yarn.lock                 # Yarn lockfile
```

### Основные модули

| Модуль                                   | Назначение                                                                                 | API                                                     |
| ---------------------------------------- | ------------------------------------------------------------------------------------------ | ------------------------------------------------------- |
| `src/index.js`                           | Точка входа, запуск сервера                                                                | `main()`                                                |
| `src/lib/server.js`                      | MCP Server setup v0.7.0                                                                    | `createServer()`, `startServer()`                       |
| `src/lib/tool-registry.js`               | Регистрация всех 37 инструментов                                                           | `ALL_TOOLS`, `getToolDefinitions()`, `getToolHandler()` |
| `src/lib/tool-executor.js`               | Единый обработчик инструментов                                                             | `executeTool()`                                         |
| `src/lib/validators.js`                  | Валидация входных данных (`validateInteger` строго проверяет целые числа через `Number()`) | 5 функций валидации                                     |
| `src/lib/applescript/executor.js`        | Выполнение AppleScript (таймаут 30с)                                                       | `runAppleScript()`                                      |
| `src/lib/applescript/helpers.js`         | Переиспользуемые фрагменты                                                                 | `COMMON_SCRIPTS`                                        |
| `src/lib/applescript/template-engine.js` | Подстановка параметров (regex-safe ключи, type-safe значения)                              | `processTemplate()`                                     |

### Модули инструментов (37 инструментов)

| Модуль     | Инструменты     | Файл                      |
| ---------- | --------------- | ------------------------- |
| Documents  | 7 инструментов  | `src/tools/documents.js`  |
| Text       | 3 инструмента   | `src/tools/text.js`       |
| Navigation | 5 инструментов  | `src/tools/navigation.js` |
| Bookmarks  | 4 инструмента   | `src/tools/bookmarks.js`  |
| Hyperlinks | 2 инструмента   | `src/tools/hyperlinks.js` |
| Paragraphs | 3 инструмента   | `src/tools/paragraphs.js` |
| Images     | 3 инструмента   | `src/tools/images.js`     |
| Tables     | 10 инструментов | `src/tools/tables.js`     |

## Сборка и запуск

### Команды Yarn

| Команда              | Назначение                        |
| -------------------- | --------------------------------- |
| `yarn install`       | Установка зависимостей            |
| `yarn build`         | Сборка (копирование src/ → dist/) |
| `yarn build:watch`   | Сборка в режиме watch             |
| `yarn dev`           | Сборка + запуск                   |
| `yarn start`         | Запуск из dist/                   |
| `yarn test`          | Запуск всех тестов                |
| `yarn test:watch`    | Тесты в режиме watch              |
| `yarn test:coverage` | Тесты с покрытием                 |
| `yarn clean`         | Удалить dist/                     |

### Система сборки

- **Простая**: Копирование файлов из `src/` в `dist/`
- **Без транспиляции**: Код уже в ES modules
- **Скрипт**: `scripts/build.js`
- **Entry point**: `dist/index.js`

## Добавление нового инструмента

1. Добавить объект `{name, description, annotations, inputSchema, handler}` в массив в `src/tools/[category].js`
2. Использовать валидаторы из `../lib/validators.js` и `runAppleScript()` из `../lib/applescript/executor.js`
3. Автоматическая регистрация через `tool-registry.js` — дополнительных действий не требуется

## Инструменты (37 шт.)

### Документы

| Инструмент          | Назначение                            | Параметры  |
| ------------------- | ------------------------------------- | ---------- |
| `create_document`   | Создать новый документ                | `content?` |
| `open_document`     | Открыть документ                      | `path`     |
| `get_document_text` | Получить весь текст                   | —          |
| `get_document_info` | Статистика (слова, символы, страницы) | —          |
| `save_document`     | Сохранить документ                    | `path?`    |
| `close_document`    | Закрыть документ                      | `save?`    |
| `export_pdf`        | Экспорт в PDF                         | `path`     |

### Текст

| Инструмент     | Назначение               | Параметры                                          |
| -------------- | ------------------------ | -------------------------------------------------- |
| `insert_text`  | Вставить текст в курсор  | `text`                                             |
| `replace_text` | Найти и заменить         | `find`, `replace`, `all?`                          |
| `format_text`  | Форматирование выделения | `bold?`, `italic?`, `underline?`, `font?`, `size?` |

### Навигация

| Инструмент               | Назначение                     | Параметры                   |
| ------------------------ | ------------------------------ | --------------------------- |
| `move_cursor_after_text` | Курсор после найденного текста | `searchText`, `occurrence?` |
| `goto_start`             | Курсор в начало документа      | —                           |
| `goto_end`               | Курсор в конец документа       | —                           |
| `get_selection_info`     | Позиция и длина выделения      | —                           |
| `select_all`             | Выделить весь документ         | —                           |

### Таблицы

| Инструмент            | Назначение                  | Параметры                                |
| --------------------- | --------------------------- | ---------------------------------------- |
| `list_tables`         | Список таблиц с размерами   | —                                        |
| `get_table_cell`      | Получить текст ячейки       | `tableIndex`, `row`, `column`            |
| `set_table_cell`      | Установить текст в ячейку   | `tableIndex`, `row`, `column`, `text`    |
| `select_table_cell`   | Переместить курсор в ячейку | `tableIndex`, `row`, `column`            |
| `find_table_header`   | Найти колонку по заголовку  | `tableIndex`, `headerText`, `headerRow?` |
| `create_table`        | Создать таблицу в курсоре   | `rows`, `columns`                        |
| `add_table_row`       | Добавить строку             | `tableIndex`, `afterRow?`                |
| `delete_table_row`    | Удалить строку              | `tableIndex`, `row`                      |
| `add_table_column`    | Добавить колонку            | `tableIndex`, `afterColumn?`             |
| `delete_table_column` | Удалить колонку             | `tableIndex`, `column`                   |

### Закладки

| Инструмент        | Назначение                    | Параметры |
| ----------------- | ----------------------------- | --------- |
| `list_bookmarks`  | Список закладок               | —         |
| `create_bookmark` | Создать закладку на выделении | `name`    |
| `goto_bookmark`   | Перейти к закладке            | `name`    |
| `delete_bookmark` | Удалить закладку              | `name`    |

### Гиперссылки

| Инструмент         | Назначение          | Параметры             |
| ------------------ | ------------------- | --------------------- |
| `list_hyperlinks`  | Список гиперссылок  | —                     |
| `create_hyperlink` | Создать гиперссылку | `url`, `displayText?` |

### Параграфы

| Инструмент            | Назначение                   | Параметры            |
| --------------------- | ---------------------------- | -------------------- |
| `list_paragraphs`     | Список параграфов со стилями | `limit?`             |
| `goto_paragraph`      | Перейти к параграфу          | `index`              |
| `set_paragraph_style` | Установить стиль параграфа   | `index`, `styleName` |

### Изображения

| Инструмент            | Назначение                               | Параметры                                        |
| --------------------- | ---------------------------------------- | ------------------------------------------------ |
| `insert_image`        | Вставить изображение через clipboard     | `path`, `width?`, `height?`                      |
| `list_inline_shapes`  | Список inline shapes (картинки, объекты) | —                                                |
| `resize_inline_shape` | Изменить размер inline shape             | `index`, `width?`, `height?`, `lockAspectRatio?` |

## Индексация

- Все индексы **1-based** (tableIndex=1, row=1, column=1, index=1)
- `get_table_cell` автоматически удаляет маркеры ячеек (ASCII 7, 13)

## AppleScript синтаксис (Word)

| Операция            | Синтаксис                                                                                                |
| ------------------- | -------------------------------------------------------------------------------------------------------- |
| Получить ячейку     | `cell COLUMN of row ROW of table`                                                                        |
| Collapse to start   | `set selection end of selection to selection start of selection`                                         |
| Collapse to end     | `set selection start of selection to selection end of selection`                                         |
| Find text           | `execute find findObj` + проверка `selStart ≠ selEnd`                                                    |
| Статистика          | `compute statistics d statistic statistic words`                                                         |
| Создать таблицу     | `make new table at text object of selection with properties {number of rows:N, number of columns:M}`     |
| Добавить строку     | `select (text object of row N of t)` → `insert rows selection position below`                            |
| Добавить колонку    | `select (text object of cell N of row 1 of t)` → `insert columns selection position insert on the right` |
| Удалить строку      | `delete row N of table`                                                                                  |
| Закладки создать    | `make new bookmark at d with properties {name:"X", \|bookmark range\|:selection}`                        |
| Закладки перейти    | `select (text object of b)`                                                                              |
| Гиперссылки создать | `tell selection` → `make new hyperlink object at end with properties {\|hyperlink address\|:"URL"}`      |
| Гиперссылки список  | `hyperlink objects of d`, `hyperlink object i of d`, `text to display of h` (в try/catch)                |
| Параграф стиль      | `name local of style of p`, `set style of p to "Heading 1"`                                              |
| Изображение         | clipboard + `paste object selection`, `inline shape N of d`, `width/height of shp`                       |

## Тестирование

### Запуск тестов

```bash
yarn test              # Запуск всех тестов
yarn test:watch        # Режим watch
yarn test:coverage     # С отчетом покрытия
```

### Структура тестов

| Тестовый файл                      | Покрытие                                               | Тестов |
| ---------------------------------- | ------------------------------------------------------ | ------ |
| `tests/validation.test.js`         | Валидация входных данных                               | 20+    |
| `tests/mcp-tools.test.js`          | Все 37 инструментов MCP                                | 80+    |
| `tests/server-integration.test.js` | Интеграция сервера                                     | 30+    |
| `tests/applescript-syntax.test.js` | Компиляция AppleScript + аудит-фиксы + template-engine | 52     |

### Покрытие кода

- **Branches**: 80%
- **Functions**: 80%
- **Lines**: 80%
- **Statements**: 80%

## Связи

- Импорт: `@modelcontextprotocol/sdk`
- Экспорт: MCP tools через stdio transport
- Тесты: `@jest/globals`, моки для AppleScript
- Сборка: Yarn v4+, Node.js v16+
