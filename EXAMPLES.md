#  Примеры: От идеи к коду

## Пример 1: Добавляем поддержку **жирного текста** в Markdown

### Что есть сейчас

```
_курсив_        ✅ Поддерживается
*звездочка*     ✗ Нет
__жирный__      ✗ Нет
~~зачеркнуто~~  ✗ Нет
```

### Хотим добавить: `__жирный текст__`

---

## Шаг 1️⃣: Запустить тест (TDD)

```python
# tests/test_bold_formatting.py

def test_apply_bold_formatting_simple():
    """Тест: простой жирный текст"""
    text = "до __жирный__ после"
    result = apply_bold_formatting(text)
    
    assert result == [
        ("до ", False),
        ("жирный", True),  # ← жирный
        (" после", False),
    ]
```

### Запускаем (падает ❌)

```bash
pytest tests/test_bold_formatting.py
# FAILED: apply_bold_formatting not found
```

---

## Шаг 2️⃣: Добавить функцию в utils

### Расширяем `utils/formatting.py`

```python
def apply_bold_formatting(text: str) -> List[Tuple[str, bool]]:
    """
    Разбивает текст на части по маркерам __жирный__.
    Возвращает список кортежей (текст, is_bold).
    """
    if not text or '__' not in text:
        return [(text, False)]
    
    result = []
    pattern = re.compile(r'__(.+?)__')
    last_end = 0
    
    for match in pattern.finditer(text):
        if match.start() > last_end:
            result.append((text[last_end:match.start()], False))
        result.append((match.group(1), True))
        last_end = match.end()
    
    if last_end < len(text):
        result.append((text[last_end:], False))
    
    return result
```

### Обновляем `utils/__init__.py`

```python
from MarkToGost.utils.formatting import (
    apply_italic_formatting,
    apply_bold_formatting,  # ← ДОБАВИЛИ
    set_run_font,
    set_paragraph_formatting,
)

__all__ = [
    # ...
    "apply_bold_formatting",  # ← ДОБАВИЛИ
]
```

### Тест проходит ✅

```bash
pytest tests/test_bold_formatting.py
# PASSED
```

---

## Шаг 3️⃣: Использовать функцию в рендерере

### В `renderer/document_renderer.py`

```python
from MarkToGost.utils.formatting import (
    apply_italic_formatting,
    apply_bold_formatting,  # ← ИМПОРТИРОВАЛИ
)

class DocumentRenderer:
    
    def _render_text_block(self, block: TextBlock):
        """Рендеринг текста с поддержкой __жирного__"""
        p = self.doc.add_paragraph()
        
        # Очищаем параграф
        p.clear()
        
        # Разбираем текст на части
        for part_text, is_bold in apply_bold_formatting(block.text):
            run = p.add_run(part_text)
            set_run_font(
                run,
                size_pt=DocumentSettings.FONT_SIZE_PT,
                bold=is_bold
            )
```

---

##  Готово! Тестируем в реальности

### Markdown файл

```markdown
# Заголовок

Это обычный текст с __жирным текстом__ внутри.

И еще один абзац с __несколькими__ __жирными__ словами.
```

### Запускаем

```bash
echo "Текст с __жирным__" > input/test.md
python -m MarkToGost.main test.md
```

### Результат

Открываем `output/test.docx` → видим **________жирный текст________** 

---

## Пример 2: Добавляем новый стиль абзаца (**Примечание**)

### Что хотим

```
:::note
Это важное примечание для читателя!
:::
```

### 1️⃣ Добавляем блок

```python
# parser/blocks.py

@dataclass
class NoteBlock(BaseBlock):
    """Блок примечания"""
    text: str
    note_type: str = "note"  # note, warning, info, error
```

### 2️⃣ Добавляем парсинг

```python
# parser/markdown_parser.py

def _is_note_block(self, line: str) -> bool:
    return line.startswith(":::note") or \
           line.startswith(":::warning") or \
           line.startswith(":::info")

def _parse_note_block(self) -> NoteBlock:
    line = self.lines[self.index].strip()
    
    # Определяем тип
    if "warning" in line:
        note_type = "warning"
    elif "info" in line:
        note_type = "info"
    else:
        note_type = "note"
    
    # Собираем текст
    content_lines = []
    self.index += 1
    
    while self.index < len(self.lines):
        line = self.lines[self.index]
        if line.strip() == ":::":
            break
        content_lines.append(line)
        self.index += 1
    
    self.index += 1  # Пропускаем закрывающий :::
    
    return NoteBlock(
        text="\n".join(content_lines).strip(),
        note_type=note_type
    )

# В parse():
elif self._is_note_block(line):
    blocks.append(self._parse_note_block())
```

### 3️⃣ Добавляем рендерер

```python
# renderer/note.py

def render_note_block(doc, block, set_paragraph_formatting, set_run_font):
    """Рендеринг примечания с цветными границами"""
    from docx.shared import RGBColor
    
    # Цвета для разных типов
    colors = {
        "note": RGBColor(0, 102, 204),      # Синий
        "warning": RGBColor(255, 153, 0),   # Оранжевый
        "info": RGBColor(51, 153, 102),     # Зелёный
        "error": RGBColor(204, 0, 0),       # Красный
    }
    
    color = colors.get(block.note_type, colors["note"])
    
    # Параграф с левой границей
    p = doc.add_paragraph()
    set_paragraph_formatting(
        p,
        left_indent=Cm(0.5),
        first_line_indent=Cm(0),
    )
    
    # Заголовок примечания
    run = p.add_run(f"[{block.note_type.upper()}] ")
    run.font.bold = True
    run.font.color.rgb = color
    
    # Текст
    run = p.add_run(block.text)
    run.font.color.rgb = color
    
    # Отступ
    doc.add_paragraph()
```

### 4️⃣ Интегрируем

```python
# renderer/document_renderer.py

from MarkToGost.renderer.note import render_note_block
from MarkToGost.parser.blocks import NoteBlock

class DocumentRenderer:
    def render_block(self, block):
        # ...
        elif isinstance(block, NoteBlock):
            self._render_note_block(block)
    
    def _render_note_block(self, block):
        render_note_block(
            self.doc, block,
            set_paragraph_formatting,
            set_run_font
        )
        self._mark_content()
```

### Результат

```markdown
:::note
Это примечание будет **синего цвета**
:::

:::warning
Это предупреждение будет **оранжевого цвета**
:::
```

---

## Пример 3: Добавляем вспомогательную функцию

### Задача: Подсчет читаемости текста (Flesch Reading Score)

### 1️⃣ Функция в utils

```python
# utils/readability.py

import re
from typing import Dict


def count_syllables(word: str) -> int:
    """Примерный подсчет слогов в слове"""
    word = word.lower()
    vowels = "aeiouy"
    syllable_count = 0
    previous_was_vowel = False
    
    for char in word:
        is_vowel = char in vowels
        if is_vowel and not previous_was_vowel:
            syllable_count += 1
        previous_was_vowel = is_vowel
    
    # Корректировки
    if word.endswith("e"):
        syllable_count -= 1
    if word.endswith("le"):
        syllable_count += 1
    
    return max(1, syllable_count)


def calculate_flesch_score(text: str) -> Dict[str, float]:
    """
    Расчет индекса читаемости Flesch
    
    Результат:
    - 90-100: Очень легко читать (начальная школа)
    - 80-90: Легко читать
    - 70-80: Довольно легко читать
    - 60-70: Стандартный уровень (средняя школа)
    - 50-60: Довольно сложно читать
    - 30-50: Сложно читать (высшее образование)
    - 0-30: Очень сложно читать
    """
    sentences = len(re.split(r'[.!?]+', text)) - 1
    words = len(text.split())
    syllables = sum(count_syllables(w) for w in text.split())
    
    if words == 0 or sentences == 0:
        return {"score": 0, "level": "Unknown"}
    
    # Формула Флеша
    score = 206.835 - 1.015 * (words / sentences) - 84.6 * (syllables / words)
    score = max(0, min(100, score))  # Клиппируем от 0 до 100
    
    # Определяем уровень
    if score >= 90:
        level = "Очень легко"
    elif score >= 80:
        level = "Легко"
    elif score >= 70:
        level = "Довольно легко"
    elif score >= 60:
        level = "Стандартный"
    elif score >= 50:
        level = "Довольно сложно"
    elif score >= 30:
        level = "Сложно"
    else:
        level = "Очень сложно"
    
    return {
        "score": round(score, 1),
        "level": level,
        "words": words,
        "sentences": sentences,
        "syllables": syllables,
    }
```

### 2️⃣ Экспортируем

```python
# utils/__init__.py

from MarkToGost.utils.readability import (
    calculate_flesch_score,
    count_syllables,
)

__all__ = [
    # ...
    "calculate_flesch_score",
    "count_syllables",
]
```

### 3️⃣ Тесты

```python
# tests/test_readability.py

from MarkToGost.utils.readability import calculate_flesch_score


def test_simple_text():
    """Простой текст легко читать"""
    text = "The cat sat on the mat."
    result = calculate_flesch_score(text)
    assert result["score"] > 70


def test_complex_text():
    """Сложный текст сложнo читать"""
    text = "The proliferation of multifaceted methodologies " \
           "necessitates comprehensive analysis."
    result = calculate_flesch_score(text)
    assert result["score"] < 50
```

### 4️⃣ Используем в main

```python
# main.py

from MarkToGost.utils import calculate_flesch_score

def create_document(md_text: str):
    # ...
    
    # Анализируем читаемость
    readability = calculate_flesch_score(md_text)
    print(f" Читаемость: {readability['level']}")
    print(f"   Индекс: {readability['score']}")
```

### Результат

```bash
$ python -m MarkToGost.main myfile.md

 Читаемость: Стандартный
   Индекс: 62.3
✅ input\myfile.md → output\myfile.docx
```

---

##  Визуальная схема: Как всё связано

```
input/myfile.md
    ↓
    ├─ Читается файл
    ├─ extract_metadata() ← из utils
    ├─ MarkdownParser.parse() ← из parser
    │  └─ find_paragraphs() ← вспомогательные функции parser
    │  └─ find_tables() ← вспомогательные функции parser
    │  └─ find_formulas() ← вспомогательные функции parser
    │
    ├─ MarkdownParser возвращает List[BaseBlock]
    │  ├─ TextBlock
    │  ├─ HeadingBlock
    │  ├─ TableBlock
    │  ├─ ImageBlock
    │  └─ FormulaBlock
    │
    ├─ DocumentRenderer.render_block()
    │  ├─ _render_text_block() ← использует apply_italic_formatting() из utils
    │  ├─ _render_table_block() ← использует split_md_table_row() из utils
    │  ├─ _render_image_block() ← использует compute_image_width_cm() из utils
    │  ├─ _render_heading_block() ← использует get_heading_level_from_number() из utils
    │  └─ _render_formula_block() ← использует formula_renderer
    │
    ├─ Document.save()
    │
    ↓
output/myfile.docx
    (готовый документ по ГОСТ)
```

---

**Теперь вы понимаете, как расширять проект! **
