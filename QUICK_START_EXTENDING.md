# ⚡ БЫСТРАЯ ШПАРГАЛКА: Как расширять MarkdownParser

## 🎯 Хотите добавить новый тип блока?

**Пример: добавить поддержку цитат `> текст`**

### Шаг 1: Блок (data class)
```python
# MarkToGost/parser/blocks.py
@dataclass
class QuoteBlock(BaseBlock):
    text: str
    author: Optional[str] = None
```

### Шаг 2: Парсер (два метода)
```python
# MarkToGost/parser/markdown_parser.py

# Проверка: это цитата?
def _is_quote(self, line: str) -> bool:
    return line.startswith("> ")

# Парсинг: преобразовать в блок
def _parse_quote(self) -> QuoteBlock:
    text = self.lines[self.index][2:].strip()  # убираем "> "
    self.index += 1
    return QuoteBlock(text=text)
```

### Шаг 3: Регистрация в parse()
```python
# В методе parse():
elif self._is_quote(line):
    blocks.append(self._parse_quote())
```

### Шаг 4: Рендер
```python
# MarkToGost/renderer/quote.py
def render_quote_block(doc, block, mark_content_cb):
    p = doc.add_paragraph()
    run = p.add_run(f'"{block.text}"')
    if block.author:
        run = p.add_run(f'\n— {block.author}')
```

### Шаг 5: Интеграция рендера
```python
# MarkToGost/renderer/document_renderer.py
from MarkToGost.renderer.quote import render_quote_block

# В render_block():
elif isinstance(block, QuoteBlock):
    render_quote_block(self.doc, block, self._mark_content)
```

### Шаг 6: Тесты
```python
# MarkToGost/tests/test_quote.py
def test_parse_quote():
    md = "> Это цитата"
    blocks = MarkdownParser(md).parse()
    assert isinstance(blocks[0], QuoteBlock)
    assert blocks[0].text == "Это цитата"
```

---

## 🔧 Хотите исправить ошибку?

**Проблема: текст обрезается**

1. **Найти метод:** `_parse_text_block()`
2. **Добавить проверку:**
```python
while self.index < len(self.lines):
    line = self.lines[self.index].strip()
    
    # Добавить условия остановки
    if not line or line.startswith("#"):
        break
    
    buffer.append(line)
    self.index += 1
```

3. **Тестировать:**
```bash
pytest MarkToGost/tests/ -v
python test_blocks.py all
```

---

## 📋 Чек-лист: 5 простых шагов

- [ ] **1. Dataclass** в `parser/blocks.py`
- [ ] **2. Методы** в `parser/markdown_parser.py` (_is_X, _parse_X)
- [ ] **3. Регистрация** в методе `parse()`
- [ ] **4. Рендер** в `renderer/X.py`
- [ ] **5. Тесты** в `tests/test_X.py`

---

## ⚠️ Частые ошибки

| Ошибка | Решение |
|--------|---------|
| "Индекс за границами" | Всегда проверяйте `if self.index >= len(self.lines)` |
| "Бесконечный цикл" | ВСЕГДА делайте `self.index += 1` |
| "Блок не распознаётся" | Проверьте регистрацию в `parse()` методе |
| "Текст дублируется" | Добавляйте в буфер ИЛИ в блоки, не оба |

---

## 🚀 Команды для быстрого старта

```bash
# Проверить один блок
python test_blocks.py 01_TextBlock

# Проверить все блоки
python test_blocks.py all

# Запустить все тесты
pytest MarkToGost/tests/ -v

# Обработать файлы
python -m MarkToGost.main
```

---

## 💬 Где найти ответ?

| Вопрос | Документ |
|--------|----------|
| "Как добавить новый блок?" | **EXTENDING_PARSER.md** |
| "Как исправить ошибку?" | **PARSER_DEBUGGING.md** |
| "Что такое архитектура?" | **ARCHITECTURE.md** |
| "Как писать тесты?" | **TESTING.md** |
| "Карта документации?" | **DOCUMENTATION_MAP.md** |

---

## 📚 Полный пример: Добавить поддержку `~~зачёркивания~~`

### 1. Блок нужен?
Нет! Используем существующий `TextBlock`.

### 2. Парсер нужен?
Нет! Markdown уже поддерживает `~~текст~~`.

### 3. Рендер нужен?
ДА! Добавить в `text.py`:

```python
def render_text_block(doc, block: TextBlock, image_refs, mark_content_cb):
    # ...
    
    # Обработка зачёркивания
    text = block.text.replace("~~", "<s>").replace("~~", "</s>")
    
    # В рендере применить strikethrough
    for part in parts:
        run = p.add_run(part)
        if "<s>" in part:
            run.font.strikethrough = True
```

### 4. Тест
```python
def test_strikethrough():
    md = "Текст со ~~зачёркиванием~~"
    blocks = MarkdownParser(md).parse()
    assert isinstance(blocks[0], TextBlock)
    assert "~~" in blocks[0].text
```

---

## ✨ Итого: Расширять MarkdownParser легко!

1. **Простое изменение?** (например, зачёркивание)
   → Отредактируйте рендер, добавьте тест

2. **Новый тип блока?** (например, цитаты)
   → Следуйте 5-шаговому чек-листу выше

3. **Ошибка в парсере?**
   → Используйте шпаргалку "Частые ошибки"

**Всё просто! Начните с шага 1! 🚀**

