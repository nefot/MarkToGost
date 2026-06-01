# 📚 Руководство по добавлению новых функций в MarkToGost

## 🏗️ Архитектура напоминание

```
config.py (настройки)
    ↓
utils/ (базовые функции - не знают о других слоях)
    ↓
parser/ (парсинг Markdown → блоки)
    ↓
renderer/ (рендеринг блоков → DOCX)
    ↓
main.py (точка входа)
```

**Правило:** Каждый слой может импортировать только слои НИЖЕ себя. Никаких циклических импортов!

---

## 📝 Сценарий 1: Добавить новый элемент Markdown

Допустим, вы хотите добавить **цитаты** (`> текст`).

### Шаг 1️⃣: Определить блок в `parser/blocks.py`

```python
# MarkToGost/parser/blocks.py

from dataclasses import dataclass
from typing import Optional, List

@dataclass
class QuoteBlock(BaseBlock):
    """Блок цитаты"""
    text: str
    author: Optional[str] = None  # Кто это сказал
    is_emphasized: bool = False   # Выделенная цитата
```

### Шаг 2️⃣: Добавить парсинг в `parser/markdown_parser.py`

```python
# MarkToGost/parser/markdown_parser.py

class MarkdownParser:
    
    def _is_quote_block(self, line: str) -> bool:
        """Проверка начала цитаты (> или >> для подцитат)"""
        return line.startswith("> ") or line.startswith(">> ")
    
    def _parse_quote(self) -> QuoteBlock:
        """Парсинг блока цитаты"""
        line = self.lines[self.index].strip()
        
        # Определяем уровень цитаты
        depth = 0
        while line.startswith("> "):
            depth += 1
            line = line[2:].strip()
        
        # Собираем весь текст цитаты
        quote_lines = [line]
        self.index += 1
        
        while self.index < len(self.lines):
            peek = self.lines[self.index].strip()
            if not peek.startswith("> "):
                break
            quote_lines.append(peek[2:].strip())
            self.index += 1
        
        # Ищем автора (опционально: --- автор)
        text = " ".join(quote_lines)
        author = None
        
        if " --- " in text:
            text, author = text.rsplit(" --- ", 1)
        
        return QuoteBlock(
            text=text.strip(),
            author=author.strip() if author else None,
            is_emphasized=(depth == 2)  # >> это выделенная цитата
        )
    
    def parse(self) -> List[BaseBlock]:
        """Основной метод парсинга"""
        blocks = []
        
        while self.index < len(self.lines):
            line = self.lines[self.index].strip()
            
            if not line:
                self.index += 1
                continue
            
            # ... другие проверки ...
            elif self._is_quote_block(line):  # ← ДОБАВИТЬ ЭТУ СТРОКУ
                blocks.append(self._parse_quote())
            # ...
        
        return blocks
```

### Шаг 3️⃣: Добавить рендерер в `renderer/`

Создаем новый файл `renderer/quote.py`:

```python
# MarkToGost/renderer/quote.py
"""Рендеринг цитат"""

from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

from MarkToGost.config import DocumentSettings
from MarkToGost.utils.formatting import set_run_font, apply_italic_formatting


def render_quote_block(doc, block, set_paragraph_formatting):
    """
    Рендеринг цитаты в документ
    
    Формат:
        ╭─────────────────────╮
        │  "Текст цитаты"     │
        │          — Автор    │
        ╰─────────────────────╯
    """
    # Цвет для выделения цитаты (серый)
    quote_color = RGBColor(100, 100, 100)
    
    # Параграф с текстом цитаты
    p = doc.add_paragraph()
    set_paragraph_formatting(
        p,
        align=WD_ALIGN_PARAGRAPH.LEFT,
        left_indent=Cm(1.0),  # Отступ слева
        first_line_indent=Cm(0),
        line_spacing=DocumentSettings.LINE_SPACING
    )
    
    # Добавляем кавычку
    run = p.add_run('"')
    set_run_font(run, size_pt=DocumentSettings.FONT_SIZE_PT, italic=True)
    run.font.color.rgb = quote_color
    
    # Текст цитаты (с поддержкой курсива _текст_)
    for part_text, is_italic in apply_italic_formatting(block.text):
        run = p.add_run(part_text)
        set_run_font(
            run,
            size_pt=DocumentSettings.FONT_SIZE_PT,
            italic=is_italic or block.is_emphasized
        )
        if block.is_emphasized:
            run.font.bold = True
        run.font.color.rgb = quote_color
    
    # Закрывающая кавычка
    run = p.add_run('"')
    set_run_font(run, size_pt=DocumentSettings.FONT_SIZE_PT, italic=True)
    run.font.color.rgb = quote_color
    
    # Автор (если есть)
    if block.author:
        run = p.add_run(f"\n— {block.author}")
        set_run_font(run, size_pt=DocumentSettings.FONT_SIZE_PT, italic=True)
        run.font.color.rgb = quote_color
    
    doc.add_paragraph()  # Отступ после цитаты
```

### Шаг 4️⃣: Интегрировать в `renderer/document_renderer.py`

```python
# MarkToGost/renderer/document_renderer.py

from MarkToGost.renderer.quote import render_quote_block
from MarkToGost.parser.blocks import QuoteBlock  # ← ДОБАВИТЬ

class DocumentRenderer:
    
    def render_block(self, block: BaseBlock):
        """Рендеринг блока"""
        if isinstance(block, Section):
            self._render_section_block(block)
        elif isinstance(block, QuoteBlock):  # ← ДОБАВИТЬ
            self._render_quote_block(block)
        # ... остальные блоки ...
    
    def _render_quote_block(self, block):
        """Делегируем в render_quote_block"""
        render_quote_block(self.doc, block, set_paragraph_formatting)
        self._mark_content()
```

### Шаг 5️⃣: Написать тесты `tests/test_quote.py`

```python
# MarkToGost/tests/test_quote.py

import pytest
from MarkToGost.parser.markdown_parser import MarkdownParser
from MarkToGost.parser.blocks import QuoteBlock


class TestParseQuote:
    """Тесты парсинга цитат"""
    
    def test_simple_quote(self):
        """Простая цитата"""
        md = "> Еще одна цитата"
        blocks = MarkdownParser(md).parse()
        assert len(blocks) == 1
        assert isinstance(blocks[0], QuoteBlock)
        assert blocks[0].text == "Еще одна цитата"
        assert blocks[0].author is None
    
    def test_quote_with_author(self):
        """Цитата с автором"""
        md = "> Жизнь — это сон — Философ"
        blocks = MarkdownParser(md).parse()
        assert isinstance(blocks[0], QuoteBlock)
        assert blocks[0].text == "Жизнь — это сон"
        assert blocks[0].author == "Философ"
    
    def test_emphasized_quote(self):
        """Выделенная цитата (>>)"""
        md = ">> Очень важно!"
        blocks = MarkdownParser(md).parse()
        assert isinstance(blocks[0], QuoteBlock)
        assert blocks[0].is_emphasized is True
    
    def test_multiline_quote(self):
        """Многострочная цитата"""
        md = "> Текст первой строки\n> Текст второй строки"
        blocks = MarkdownParser(md).parse()
        assert isinstance(blocks[0], QuoteBlock)
        assert "первой" in blocks[0].text
        assert "второй" in blocks[0].text
```

### ✅ Готово! Теперь цитаты работают:

```bash
# Тестируем
pytest MarkToGost/tests/test_quote.py -v

# Используем
echo "> Знание — сила — Фрэнсис Бэкон" > input/test.md
python -m MarkToGost.main test.md
```

---

## 🛠️ Сценарий 2: Добавить вспомогательную функцию

Допустим, вы хотите функцию для **проверки палиндромов в тексте**.

### Шаг 1️⃣: Добавить функцию в `utils/`

Создаем файл `utils/text_helpers.py`:

```python
# MarkToGost/utils/text_helpers.py
"""Вспомогательные функции для работы с текстом"""

import re
from typing import List, Tuple


def find_palindromes(text: str) -> List[Tuple[str, int]]:
    """
    Находит все палиндромы в тексте
    
    Args:
        text: Исходный текст
        
    Returns:
        Список (палиндром, позиция) в тексте
        
    Examples:
        >>> find_palindromes("Это завод ада заводима")
        [("завод ада завод", 7)]
    """
    palindromes = []
    words = re.findall(r'\b\w+\b', text.lower())
    
    for i, word in enumerate(words):
        if word == word[::-1] and len(word) > 2:
            pos = text.lower().find(word)
            palindromes.append((word, pos))
    
    return palindromes


def truncate_text(text: str, max_length: int = 100, suffix: str = "...") -> str:
    """
    Обрезает текст до max_length и добавляет суффикс
    
    Args:
        text: Исходный текст
        max_length: Максимальная длина
        suffix: Что добавляются в конец
        
    Returns:
        Обрезанный текст
    """
    if len(text) <= max_length:
        return text
    return text[:max_length - len(suffix)] + suffix
```

### Шаг 2️⃣: Экспортировать из `utils/__init__.py`

```python
# MarkToGost/utils/__init__.py

from MarkToGost.utils.text_helpers import (
    find_palindromes,
    truncate_text,
)

__all__ = [
    # ... существующие ...
    "find_palindromes",
    "truncate_text",
]
```

### Шаг 3️⃣: Написать тесты

```python
# MarkToGost/tests/test_text_helpers.py

import pytest
from MarkToGost.utils.text_helpers import find_palindromes, truncate_text


class TestFindPalindromes:
    
    def test_find_palindromes(self):
        result = find_palindromes("Это завод ада заводима")
        assert ("ада", 11) in result
    
    def test_no_palindromes(self):
        result = find_palindromes("Hello world example text")
        assert len([p for p in result if len(p[0]) > 2]) == 0


class TestTruncateText:
    
    def test_short_text_unchanged(self):
        text = "Short"
        assert truncate_text(text, 100) == text
    
    def test_long_text_truncated(self):
        text = "X" * 100
        result = truncate_text(text, 50)
        assert len(result) == 53  # 50 - 3 (len(suffix)) + 3
        assert result.endswith("...")
```

### ✅ Использование:

```python
from MarkToGost.utils import find_palindromes, truncate_text

# В любом месте проекта:
palindromes = find_palindromes("Ночь. Улица. Фонарь. Аптека.")
short = truncate_text("Очень длинный текст...", max_length=10)
```

---

## 🎨 Сценарий 3: Добавить новый параметр конфигурации

Вы хотите добавить **поддержку темных тем**.

### Шаг 1️⃣: Обновить `config.py`

```python
# MarkToGost/config.py

class DocumentSettings:
    """Настройки документа"""
    FONT_NAME = "Times New Roman"
    FONT_SIZE_PT = 14
    # ... существующие ...


class ThemeSettings:
    """Настройки темы документа"""
    # Светлая тема (по умолчанию)
    THEME_LIGHT = {
        "text_color": (0, 0, 0),      # Чёрный
        "background": (255, 255, 255), # Белый
        "quote_color": (100, 100, 100), # Серый
        "code_background": (240, 240, 240), # Светло-серый
    }
    
    # Тёмная тема
    THEME_DARK = {
        "text_color": (220, 220, 220),  # Светло-серый
        "background": (30, 30, 30),     # Тёмный
        "quote_color": (150, 150, 150), # Светлый серый
        "code_background": (50, 50, 50), # Тёмно-серый
    }
    
    # Текущая тема (можно менять)
    ACTIVE_THEME = THEME_LIGHT
```

### Шаг 2️⃣: Использовать в утилитах

```python
# MarkToGost/utils/theme_helpers.py
"""Управление темами документа"""

from MarkToGost.config import ThemeSettings


def get_text_color():
    """Получить цвет текста текущей темы"""
    color_tuple = ThemeSettings.ACTIVE_THEME["text_color"]
    from docx.shared import RGBColor
    return RGBColor(*color_tuple)


def get_background_color():
    """Получить цвет фона текущей темы"""
    color_tuple = ThemeSettings.ACTIVE_THEME["background"]
    from docx.shared import RGBColor
    return RGBColor(*color_tuple)
```

### Шаг 3️⃣: Использовать в рендерере

```python
# MarkToGost/renderer/document_renderer.py

from MarkToGost.utils.theme_helpers import get_text_color

class DocumentRenderer:
    
    def _render_text_block(self, block):
        # ...
        for part_text, is_italic in apply_italic_formatting(text):
            run = p.add_run(part_text)
            run.font.color.rgb = get_text_color()  # ← Используем тему
```

### ✅ Изменить тему:

```python
# main.py
from MarkToGost.config import ThemeSettings

# До cоздания документа:
ThemeSettings.ACTIVE_THEME = ThemeSettings.THEME_DARK

doc = create_document(md_text)
```

---

## 📊 Сценарий 4: Добавить статистику документа

Вы хотите собирать **статистику: количество слов, абзацев, таблиц**.

### Шаг 1️⃣: Создать класс статистики

```python
# MarkToGost/utils/statistics.py
"""Сбор статистики документа"""

from dataclasses import dataclass
from typing import List
from MarkToGost.parser.blocks import BaseBlock


@dataclass
class DocumentStatistics:
    """Статистика документа"""
    total_blocks: int = 0
    text_blocks: int = 0
    heading_blocks: int = 0
    code_blocks: int = 0
    table_blocks: int = 0
    formula_blocks: int = 0
    image_blocks: int = 0
    total_words: int = 0
    total_characters: int = 0


def collect_statistics(blocks: List[BaseBlock]) -> DocumentStatistics:
    """Собирает статистику из списка блоков"""
    from MarkToGost.parser.blocks import (
        TextBlock, HeadingBlock, CodeBlock, TableBlock,
        FormulaBlock, ImageBlock
    )
    
    stats = DocumentStatistics()
    
    for block in blocks:
        stats.total_blocks += 1
        
        if isinstance(block, TextBlock):
            stats.text_blocks += 1
            stats.total_words += len(block.text.split())
            stats.total_characters += len(block.text)
        elif isinstance(block, HeadingBlock):
            stats.heading_blocks += 1
        elif isinstance(block, CodeBlock):
            stats.code_blocks += 1
        elif isinstance(block, TableBlock):
            stats.table_blocks += 1
        elif isinstance(block, FormulaBlock):
            stats.formula_blocks += 1
        elif isinstance(block, ImageBlock):
            stats.image_blocks += 1
    
    return stats
```

### Шаг 2️⃣: Использовать в main.py

```python
# MarkToGost/main.py

from MarkToGost.utils.statistics import collect_statistics


def create_document(md_text: str) -> Document:
    # ... существующий код ...
    
    # Парсинг
    parser = MarkdownParser(md_text_for_parsing)
    blocks = parser.parse()
    
    # Собираем статистику
    stats = collect_statistics(blocks)
    print(f"📊 Статистика:")
    print(f"   Блоков: {stats.total_blocks}")
    print(f"   Слов: {stats.total_words}")
    print(f"   Таблиц: {stats.table_blocks}")
    print(f"   Формул: {stats.formula_blocks}")
    
    # Рендеринг
    renderer = DocumentRenderer(doc, toc_entries)
    # ...
```

---

## 📋 Checklist для добавления новой функции

Используйте этот чеклист для любой новой функции:

- [ ] **Новый блок?**
  - [ ] Добавить класс в `parser/blocks.py`
  - [ ] Добавить парсинг в `parser/markdown_parser.py`
  - [ ] Создать `renderer/new_block.py` с функцией рендеринга
  - [ ] Интегрировать в `renderer/document_renderer.py`
  - [ ] Написать тесты в `tests/test_new_block.py`

- [ ] **Вспомогательная функция?**
  - [ ] Создать файл в `utils/`
  - [ ] Экспортировать из `utils/__init__.py`
  - [ ] Написать тесты в `tests/test_*.py`
  - [ ] Добавить примеры использования

- [ ] **Новый параметр конфигурации?**
  - [ ] Добавить в `config.py`
  - [ ] Использовать через `from MarkToGost.config import *`
  - [ ] Написать тесты для новой опции

- [ ] **CLI команда?**
  - [ ] Добавить аргумент в `argparse` в `main.py`
  - [ ] Реализовать функцию обработки
  - [ ] Добавить справку `--help`

---

## 🔍 Проверка архитектуры

Перед коммитом проверьте:

```bash
# 1. Все тесты должны проходить
pytest MarkToGost/tests/ -v

# 2. Нет циклических импортов
python -c "import MarkToGost.main"

# 3. Код должен работать
python -m MarkToGost.main --help

# 4. На конкретном файле
python -m MarkToGost.main test.md
```

---

## 🚀 Быстрый старт для расширения

```bash
# 1. Создать ветку
git checkout -b feature/new-element

# 2. Добавить блок в parser/blocks.py
# 3. Добавить парсинг в parser/markdown_parser.py
# 4. Создать renderer/new_element.py
# 5. Интегрировать в renderer/document_renderer.py
# 6. Написать тесты в tests/
# 7. Проверить

pytest MarkToGost/tests/ -v
python -m MarkToGost.main

# 8. Коммитить
git add .
git commit -m "Add support for new-element"
```

---

## 💡 Советы

1. **Следите за импортами!** Используйте инструмент для проверки циклов
   ```bash
   python -c "import MarkToGost.main; print('✅ No circular imports')"
   ```

2. **Документируйте код** с docstrings и примерами
   ```python
   def my_function(param: str) -> str:
       """Здесь описание
       
       Args:
           param: Описание параметра
           
       Returns:
           Описание результата
           
       Examples:
           >>> my_function("test")
           "result"
       """
   ```

3. **Писать тесты ПЕРЕД кодом** (TDD)
   - Сначала пишите тест, который падает
   - Потом пишите код, который его проходит

4. **Используйте type hints**
   ```python
   from typing import List, Optional
   
   def process(items: List[str], prefix: Optional[str] = None) -> dict:
       ...
   ```

5. **Одна ответственность на функцию**
   - Функция должна делать ОДНО
   - Если функция длинная > 50 строк, разбейте её

6. **Избегайте глобального состояния**
   - Передавайте параметры в функции
   - Не изменяйте глобальные переменные

---

Удачи в разработке! 🚀

