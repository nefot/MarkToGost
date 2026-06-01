# MarkToGost Архитектура после рефакторинга

## Структура проекта

```
MarkToGost/
├── main.py                  # Точка входа: process_md_file(), create_document()
├── config.py                # Глобальные настройки: DocumentSettings, CaptionSettings
│
├── parser/                  # Парсинг Markdown в блоки
│   ├── __init__.py
│   ├── blocks.py            # Dataclasses: TextBlock, HeadingBlock, FormulaBlock, ...
│   ├── markdown_parser.py   # MarkdownParser class
│   └── metadata.py          # extract_metadata()
│
├── renderer/                # Рендеринг блоков в DOCX
│   ├── __init__.py
│   ├── document_renderer.py # DocumentRenderer class (диспетчер)
│   ├── text.py              # _render_text_block, _render_heading_block
│   ├── formula.py           # _render_formula_block
│   ├── image.py             # _render_image_block
│   ├── table.py             # _render_table_block
│   ├── list_.py             # _render_list_block
│   └── code.py              # _render_code_block
│
├── utils/                   # Общие вспомогательные функции
│   ├── __init__.py          # Экспорт всех функций
│   ├── formatting.py        # set_run_font, apply_italic_formatting, set_paragraph_formatting
│   ├── xml_helpers.py       # set_table_borders, add_page_number_centered
│   ├── toc.py               # add_toc, reset_heading_styles, get_heading_level_from_number
│   └── document_helpers.py  # split_md_table_row, normalize_table_caption, compute_image_width_cm
│
└── tests/                   # Тесты
    ├── __init__.py
    ├── test_formatting.py         # Тесты utils/formatting.py (17 тестов)
    ├── test_document_helpers.py   # Тесты utils/document_helpers.py (26 тестов)
    ├── test_parser.py             # Тесты parser/ (если нужны)
    ├── test_metadata.py           # Тесты metadata extraction
    └── ...
```

## Принцип импортов (безопасность от циклов)

**Иерархия зависимостей:**

```
config.py  (нет зависимостей)
    ↓
utils/  (знает только о config)
    ↓
parser/  (знает о config, utils)
    ↓
renderer/  (знает о config, utils, parser)
    ↓
main.py  (знает обо всём)
```

**Правило:** Модуль не должен импортировать ничего "выше" себя в структуре.

## Основные функции

### main.py (Точка входа)

```python
# Создание документа из Markdown
doc = create_document(md_text)
doc.save("output.docx")

# Обработка файла
process_md_file("input.md")
```

### parser/markdown_parser.py

```python
parser = MarkdownParser(md_text)
blocks = parser.parse()
# → List[BaseBlock]
```

### renderer/document_renderer.py

```python
renderer = DocumentRenderer(doc, toc_entries)
for block in blocks:
    renderer.render_block(block)
```

### utils/

**formatting.py:**
- `apply_italic_formatting(text)` — разбор `_курсива_`
- `set_run_font(run, size_pt, bold, italic)` — установка шрифта
- `set_paragraph_formatting(paragraph, **kwargs)` — форматирование абзаца

**xml_helpers.py:**
- `set_table_borders(table)` — границы таблицы
- `set_repeat_table_header(row)` — повторение заголовка таблицы
- `add_page_number_centered(document)` — номера страниц

**toc.py:**
- `add_toc(document)` — вставка оглавления
- `reset_heading_styles(doc)` — сброс стилей заголовков
- `get_heading_level_from_number(text)` — определение уровня по нумерации (1. → 1, 1.1 → 2, ...)

**document_helpers.py:**
- `split_md_table_row(line)` — разбор строки таблицы
- `is_md_table_row(line)` — проверка строки таблицы
- `is_md_table_separator(line)` — проверка разделителя (| --- | --- |)
- `normalize_table_caption(text)` — извлечение названия таблицы
- `replace_image_refs(text, refs)` — замена @img_id на рис. N
- `compute_image_width_cm(doc, fraction)` — вычисление ширины изображения
- `render_formula_with_pandoc(latex)` — конвертация TeX → OMML (через pandoc)

## Тесты

### Запуск всех тестов
```bash
pytest MarkToGost/tests/ -v
```

### Текущее покрытие
- **test_formatting.py**: 17 тестов (apply_italic_formatting, set_run_font)
- **test_document_helpers.py**: 26 тестов (split_md_table_row, is_md_table_separator, normalize_table_caption, replace_image_refs)
- **Итого**: 43 теста ✅

### Как добавить тесты для parser/

```python
# tests/test_parser.py
from MarkToGost.parser.markdown_parser import MarkdownParser
from MarkToGost.parser.blocks import FormulaBlock

def test_parse_formula_block():
    md = "$$\nF = ma\n$$\n\nгде $F$ — сила."
    blocks = MarkdownParser(md).parse()
    assert isinstance(blocks[0], FormulaBlock)
    assert blocks[0].latex == "F = ma"
    assert "где" in blocks[0].explanation
```

## Миграция со старой структуры

Старый файл `md_to_gost_block.py` содержал монолитный код. После рефакторинга:

1. ✅ **config.py** — содержит `DocumentSettings` (был в md_to_gost_block.py)
2. ✅ **parser/** — вразбрасывает MarkdownParser (был в md_to_gost_block.py)
3. ✅ **renderer/** — вразбрасывает DocumentRenderer (был в document_renderer.py)
4. ✅ **utils/** — содержит все вспомогательные функции (были в md_to_gost_block.py)
5. ✅ **main.py** — точка входа с `create_document()` и `process_md_file()`

**Что сохранилось:** Все функции, классы и поведение. **Что изменилось:** Организация кода.

## Как использовать

### Как пользователь (пакетная обработка)

```bash
cd MarkToGost
python -m main
# Обработает все .md файлы в текущей директории
```

### Как разработчик (программный API)

```python
from MarkToGost.main import create_document

with open("input.md", "r") as f:
    md_text = f.read()

doc = create_document(md_text)
doc.save("output.docx")
```

### Как расширять

1. **Новый элемент Markdown?** → Добавить класс в `parser/blocks.py`
2. **Новой метод рендеринга?** → Добавить функцию в `renderer/` (например, `renderer/new_block.py`)
3. **Новые утилиты?** → Добавить функцию в соответствующий модуль `utils/`
4. **Тестирование?** → Добавить тесты в `tests/test_*.py`

## Запуск локально

```bash
# Установка зависимостей
pip install python-docx pytest

# Запуск тестов
pytest MarkToGost/tests/ -v

# Обработка Markdown файлов
python -m MarkToGost.main

# Создание docs
python -m MarkToGost.main < input.md > output.docx
```

## Примечания

- **Никаких циклических импортов** — каждый модуль знает только о "младших" слоях
- **Легко тестировать** — каждая функция изолирована, нет глобального состояния
- **Легко расширять** — добавьте новый рендерер в `renderer/`, новый парсер в `parser/`
- **ГОСТ 7.32-2001** — соблюдаются все требования по форматированию

