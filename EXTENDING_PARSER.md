# 🔧 Гайд по расширению MarkdownParser

## 📋 Философия кода

`MarkdownParser` следует **одному паттерну** для каждого типа блока:

```
1. _is_<type>()        → Проверка: это блок этого типа?
2. _parse_<type>()     → Парсинг: преобразовать в объект блока
3. render_block()      → Использование: вызвать парсер в основном цикле
```

**Пример для TextBlock:**
```python
# 1. Проверка
def _is_text_block(self, line: str) -> bool:
    return not line.startswith("#") and not line.startswith("![") ...

# 2. Парсинг
def _parse_text_block(self) -> TextBlock:
    buffer = []
    while ... : buffer.append(line)
    return TextBlock(text=" ".join(buffer))

# 3. Использование в parse()
elif self._is_text_block(line):
    blocks.append(self._parse_text_block())
```

---

## 🆕 Как добавить НОВЫЙ тип блока

Допустим, вы хотите добавить **Блоки цитат** (`> текст`).

### Шаг 1️⃣: Определить dataclass в `parser/blocks.py`

```python
# MarkToGost/parser/blocks.py

from dataclasses import dataclass
from typing import Optional

@dataclass
class QuoteBlock(BaseBlock):
    """Блок цитаты"""
    text: str
    author: Optional[str] = None
    is_emphasized: bool = False
```

### Шаг 2️⃣: Добавить методы в `MarkdownParser`

```python
# MarkToGost/parser/markdown_parser.py

class MarkdownParser:
    # ...
    
    # 1️⃣ Проверка: это цитата?
    def _is_quote_block(self, line: str) -> bool:
        """Цитата начинается с > или >>"""
        return line.startswith("> ") or line.startswith(">> ")
    
    # 2️⃣ Парсинг: преобразовать в блок
    def _parse_quote_block(self) -> QuoteBlock:
        """Собрать все строки цитаты"""
        line = self.lines[self.index].strip()
        
        # Определяем уровень (> = 1, >> = 2)
        depth = 0
        while line.startswith("> "):
            depth += 1
            line = line[2:].strip()
        
        # Собираем текст цитаты
        quote_lines = [line]
        self.index += 1
        
        while self.index < len(self.lines):
            peek = self.lines[self.index].strip()
            if not peek.startswith("> "):
                break
            quote_lines.append(peek[2:].strip())
            self.index += 1
        
        # Парсим автора (опционально: --- автор)
        full_text = " ".join(quote_lines)
        author = None
        if " --- " in full_text:
            full_text, author = full_text.rsplit(" --- ", 1)
        
        return QuoteBlock(
            text=full_text.strip(),
            author=author.strip() if author else None,
            is_emphasized=(depth == 2)
        )
```

### Шаг 3️⃣: Зарегистрировать в `parse()`

```python
# MarkToGost/parser/markdown_parser.py

def parse(self) -> List[BaseBlock]:
    blocks = []
    
    while self.index < len(self.lines):
        line = self.lines[self.index].strip()
        
        if not line:
            self.index += 1
            continue
        
        # ... существующие проверки ...
        
        elif self._is_quote_block(line):  # ← ДОБАВИТЬ ЭТУ СТРОКУ
            blocks.append(self._parse_quote_block())
        
        # ... остальное ...
```

### Шаг 4️⃣: Добавить рендерер в `renderer/quote.py`

```python
# MarkToGost/renderer/quote.py

def render_quote_block(doc, block, mark_content_cb):
    """Рендеринг цитаты"""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import RGBColor, Cm
    from MarkToGost.config import DocumentSettings
    from MarkToGost.utils.formatting import set_run_font, set_paragraph_formatting
    
    p = doc.add_paragraph()
    set_paragraph_formatting(
        p,
        align=WD_ALIGN_PARAGRAPH.LEFT,
        left_indent=Cm(1.0),
        first_line_indent=Cm(0)
    )
    
    quote_color = RGBColor(100, 100, 100)
    
    # Дозируем кавычку
    run = p.add_run('"')
    set_run_font(run, size_pt=DocumentSettings.FONT_SIZE_PT, italic=True)
    run.font.color.rgb = quote_color
    
    # Текст
    run = p.add_run(block.text)
    set_run_font(run, size_pt=DocumentSettings.FONT_SIZE_PT, italic=block.is_emphasized or False)
    run.font.color.rgb = quote_color
    
    # Закрывающая кавычка
    run = p.add_run('"')
    set_run_font(run, size_pt=DocumentSettings.FONT_SIZE_PT, italic=True)
    run.font.color.rgb = quote_color
    
    # Автор
    if block.author:
        run = p.add_run(f"\n— {block.author}")
        set_run_font(run, size_pt=DocumentSettings.FONT_SIZE_PT, italic=True)
        run.font.color.rgb = quote_color
    
    doc.add_paragraph()
    mark_content_cb()
```

### Шаг 5️⃣: Зарегистрировать в `DocumentRenderer`

```python
# MarkToGost/renderer/document_renderer.py

from MarkToGost.renderer.quote import render_quote_block
from MarkToGost.parser.blocks import QuoteBlock

class DocumentRenderer:
    def render_block(self, block: BaseBlock):
        """Рендеринг блока"""
        # ...
        elif isinstance(block, QuoteBlock):  # ← ДОБАВИТЬ
            render_quote_block(self.doc, block, self._mark_content)
```

### Шаг 6️⃣: Написать тесты

```python
# MarkToGost/tests/test_quote.py

import pytest
from MarkToGost.parser.markdown_parser import MarkdownParser
from MarkToGost.parser.blocks import QuoteBlock

def test_parse_simple_quote():
    md = "> Это цитата"
    blocks = MarkdownParser(md).parse()
    assert len(blocks) == 1
    assert isinstance(blocks[0], QuoteBlock)
    assert blocks[0].text == "Это цитата"

def test_parse_quote_with_author():
    md = "> Жизнь прекрасна — Философ"
    blocks = MarkdownParser(md).parse()
    assert blocks[0].text == "Жизнь прекрасна"
    assert blocks[0].author == "Философ"

def test_parse_emphasized_quote():
    md = ">> Очень важно!"
    blocks = MarkdownParser(md).parse()
    assert blocks[0].is_emphasized == True
```

### ✅ Готово! Теперь цитаты работают:

```markdown
> Обычная цитата

>> Выделенная цитата — Автор

Текст после цитаты
```

---

## 🔧 Как ИСПРАВЛЯТЬ существующие блоки

### Проблема 1: Неверный парсинг TextBlock

**Текущее поведение:** Текст обрезается или не собирается правильно.

**Найти:** `_parse_text_block()`

```python
def _parse_text_block(self) -> TextBlock:
    """Парсинг блока обычного текста"""
    buffer = []
    # ... цикл сбора строк ...
    start_index = self.index
    
    while self.index < len(self.lines) and iterations < max_iterations:
        # ← ЗДЕСЬ ПРОБЛЕМА?
```

**Исправить:** Добавить проверку на новые условия остановки

```python
def _parse_text_block(self) -> TextBlock:
    buffer = []
    
    while self.index < len(self.lines):
        line = self.lines[self.index].strip()
        
        # Условия остановки - добавьте сюда новые типы блоков
        if not line or \
           line.startswith("#") or \
           line.startswith("![") or \
           line.startswith("> ") or \  # ← ДОБАВИЛИ проверку цитат
           line.startswith("- ") or \
           self._is_table_start() or \
           self._is_list_start():
            break
        
        buffer.append(line)
        self.index += 1
    
    return TextBlock(text=" ".join(buffer))
```

---

### Проблема 2: TableBlock неверно собирает строки

**Найти:** `_parse_table()`

```python
def _parse_table(self) -> TableBlock:
    header = self.lines[self.index].strip()
    separator = self.lines[self.index + 1].strip()
    self.index += 2
    
    table_lines = [header]
    
    # ← ЗДЕСЬ МОЖЕТ БЫТЬ ПРОБЛЕМА
    while self.index < len(self.lines):
        candidate = self.lines[self.index].strip()
        if candidate and is_md_table_row(candidate):
            table_lines.append(candidate)
            self.index += 1
        else:
            break
    
    return TableBlock(rows=table_lines, caption=...)
```

**Исправить:** Добавить отладку

```python
def _parse_table(self) -> TableBlock:
    header = self.lines[self.index].strip()
    separator = self.lines[self.index + 1].strip()
    self.index += 2
    
    # Проверим что это действительно таблица
    if not is_md_table_row(header) or not is_md_table_separator(separator):
        # Не таблица - вернуть как текст
        self.index -= 2
        return self._parse_text_block()
    
    table_lines = [header]
    
    while self.index < len(self.lines):
        candidate = self.lines[self.index].strip()
        
        # Пустая строка после таблицы
        if not candidate:
            break
        
        # Это таблица?
        if is_md_table_row(candidate):
            table_lines.append(candidate)
            self.index += 1
        else:
            # Не таблица - конец таблицы
            break
    
    return TableBlock(rows=table_lines, caption=...)
```

---

### Проблема 3: FormulaBlock не собирает многострочные формулы

**Текущее:**
```python
def _parse_formula_block(self):
    line = self.lines[self.index].strip()
    line = line[2:].strip()  # Убираем $$
    
    buffer = []
    while True:
        if "$$" in line:
            # ← ПРОБЛЕМА: Не обрабатывает \n правильно
            before, _, _ = line.partition("$$")
            buffer.append(before)
            break
        else:
            buffer.append(line)
            self.index += 1
            if self.index >= len(self.lines):
                break
            line = self.lines[self.index].strip()
    
    formula_text = "\\n".join(buffer).strip()
```

**Исправить:**
```python
def _parse_formula_block(self):
    self.index += 1  # Пропускаем строку с $$
    
    buffer = []
    while self.index < len(self.lines):
        line = self.lines[self.index]
        
        if line.strip() == "$$":
            # Конец формулы
            self.index += 1
            break
        
        # Добавляем строку (без strip, чтобы сохранить отступы)
        buffer.append(line.rstrip())
        self.index += 1
    
    formula_text = "\n".join(buffer).strip()
    
    # Собираем пояснения...
```

---

## 📊 Структура добавления: Чек-лист

Используйте этот чек-лист при добавлении нового типа блока:

### 1. Parser (`parser/blocks.py`)
- [ ] Добавить `@dataclass class <Type>Block(BaseBlock)`
- [ ] Добавить поля с типами
- [ ] Добавить docstring

### 2. Converter (`parser/markdown_parser.py`)
- [ ] Добавить `def _is_<type>(self, line: str) -> bool`
- [ ] Добавить `def _parse_<type>(self) -> <Type>Block`
- [ ] Зарегистрировать в `parse()` методе

### 3. Renderer (`renderer/<type>.py`)
- [ ] Создать новый файл `renderer/<type>.py`
- [ ] Написать `def render_<type>_block(doc, block, mark_content_cb)`
- [ ] Импортировать в `document_renderer.py`
- [ ] Добавить в диспетчер `render_block()`

### 4. Tests (`tests/`)
- [ ] Создать `tests/test_<type>.py`
- [ ] Написать тесты парсинга
- [ ] Написать тесты рендеринга

### 5. Documentation
- [ ] Создать пример в `examples/blocks/0N_<Type>Block.md`
- [ ] Обновить `README.md`

---

## 🎨 Примеры расширений

### Добавить поддержку `~~зачёркивания~~`

**В `_parse_text_block()`:**
```python
def apply_strikethrough(text):
    return text.replace("~~", "<s>").replace("~~", "</s>")
```

**В рендере:**
```python
for part_text, is_italic in apply_italic_formatting(text):
    run = p.add_run(part_text)
    if "~~" in text:
        run.font.strikethrough = True
```

### Добавить поддержку `[сноска]`

```python
# parser/blocks.py
@dataclass
class FootnoteBlock(BaseBlock):
    text: str
    footnote_id: str
    footnote_text: str

# parser/markdown_parser.py
def _is_footnote(self, line: str) -> bool:
    return re.match(r'^\[\^[^\]]+\]:', line)
```

### Добавить поддержку `> блокировок` (divs)

```python
# parser/blocks.py
@dataclass
class BlockBlock(BaseBlock):  # Блокировка (div>
    text: str
    block_type: str  # "info", "warning", "danger"

# parser/markdown_parser.py
def _is_block(self, line: str) -> bool:
    return line.startswith("> [!") or line.startswith("> [warning]")
```

---

## 🧪 Тестирование расширений

```bash
# Запустить тесты для нового типа
pytest MarkToGost/tests/test_quote.py -v

# Запустить все тесты
pytest MarkToGost/tests/ -v

# Проверить парсер вручную
python test_blocks.py 09_QuoteBlock
```

---

## 🚀 Порядок работы

Если вы хотите:

1. **Добавить новый тип блока**
   - Следуйте чек-листу выше
   - Начните с шага 1 (dataclass)
   - Тестируйте на каждом шаге

2. **Исправить существующий блок**
   - Сначала напишите тест, который должен падать
   - Потом исправьте код в парсере
   - Убедитесь что тест проходит

3. **Оптимизировать парсер**
   - Добавьте граничные случаи
   - Проверьте на больших файлах
   - Напишите спецификацию поведения

---

**Готов помочь с конкретным расширением? Скажите какой тип блока нужен! 🚀**

