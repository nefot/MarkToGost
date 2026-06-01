# 🐛 Шпаргалка: Ошибки и исправления

## ❌ Частые ошибки при работе с парсером

### 1. "Текст обрезается посередине"

**Признак:** `TextBlock` содержит только первую строку

**Причина:** Условие остановки в `_parse_text_block()` срабатывает слишком рано

**Решение:**
```python
# ❌ НЕПРАВИЛЬНО - прерывается на первой пустой строке
while self.index < len(self.lines):
    line = self.lines[self.index].strip()
    if not line:  # ← ЗДЕСЬ ОШИБКА
        break

# ✅ ПРАВИЛЬНО - позволяет пустые строки внутри абзаца
while self.index < len(self.lines):
    line = self.lines[self.index].strip()
    
    # Пустая строка между абзацами
    if not line:
        self.index += 1
        # Проверяем что следующая строка не является новым блоком
        if self.index < len(self.lines):
            peek = self.lines[self.index].strip()
            if peek.startswith("#") or peek.startswith("!["):
                break
        continue
```

---

### 2. "Таблица не распознаётся"

**Признак:** Таблица парсится как текст

**Причина:** `_is_table_start()` неправильно проверяет разделитель

**Решение:**
```python
# ❌ НЕПРАВИЛЬНО
def _is_table_start(self) -> bool:
    if self.index + 1 >= len(self.lines):
        return False
    return is_md_table_row(self.lines[self.index])

# ✅ ПРАВИЛЬНО - проверяем оба: строку AND разделитель
def _is_table_start(self) -> bool:
    if self.index + 1 >= len(self.lines):
        return False
    current = self.lines[self.index].strip()
    next_line = self.lines[self.index + 1].strip()
    
    # ОБЯЗАТЕЛЬНО оба условия
    return is_md_table_row(current) and is_md_table_separator(next_line)
```

---

### 3. "Формула не собирается многострочной"

**Признак:** `FormulaBlock.latex` содержит только одну строку

**Причина:** Неверное использование `split()` или `strip()`

**Решение:**
```python
# ❌ НЕПРАВИЛЬНО - теряет строки
line = line.strip()
if "$$" in line:
    before, _, _ = line.partition("$$")
    buffer.append(before.strip())  # ← strip теряет выравнивание
    break

# ✅ ПРАВИЛЬНО - сохраняет структуру
self.index += 1
while self.index < len(self.lines):
    line = self.lines[self.index]
    
    if line.strip() == "$$":
        self.index += 1
        break
    
    # Сохраняем оригинальную строку без strip!
    buffer.append(line.rstrip("\n"))
    self.index += 1

formula_text = "\n".join(buffer)
```

---

### 4. "Список прерывается на первом элементе"

**Признак:** `ListBlock.items` содержит только одну строку

**Причина:** Неверная регулярка для проверки элемента

**Решение:**
```python
# ❌ НЕПРАВИЛЬНО
def _parse_list(self) -> ListBlock:
    items = []
    first_line = self.lines[self.index].strip()
    ordered = bool(re.match(r'^\s*\d+[.)]\s+', first_line))
    
    while self.index < len(self.lines):
        line = self.lines[self.index].strip()
        if not line:  # ← Прерывается на пустой строке
            break
        
        m = re.match(r'^\d+[.)]\s+(.*)', line)  # ← Не учитывает неюорядоченные

# ✅ ПРАВИЛЬНО
def _parse_list(self) -> ListBlock:
    items = []
    first_line = self.lines[self.index].strip()
    ordered = bool(re.match(r'^\d+[.)]\s+', first_line))
    
    while self.index < len(self.lines):
        line = self.lines[self.index].strip()
        
        if not line:
            # Пропускаем пустые строки внутри списка
            self.index += 1
            continue
        
        # Проверяем тип элемента
        if ordered:
            m = re.match(r'^\d+[.)]\s+(.*)', line)
        else:
            m = re.match(r'^[-–—*+]\s+(.*)', line)  # ← Все типы маркеров
        
        if m:
            items.append(m.group(1))
            self.index += 1
        else:
            break  # Конец списка
    
    return ListBlock(items=items, ordered=ordered)
```

---

### 5. "Секция с фигурными скобками не распознаётся"

**Признак:** `Section` пуст или неполный

**Причина:** Неверное подсчитывание скобок

**Решение:**
```python
# ❌ НЕПРАВИЛЬНО - не проверяет закрывающей скобки
def _parse_section_with_braces(self, ...):
    self.index += 1
    blocks = []
    
    while self.index < len(self.lines):
        line = self.lines[self.index].strip()
        # ← Забыли проверку на "}"
        
        blocks.append(...)
        self.index += 1

# ✅ ПРАВИЛЬНО
def _parse_section_with_braces(self, ...):
    self.index += 1
    blocks = []
    brace_count = 1  # Одна скобка открыта в строке 1
    
    while self.index < len(self.lines) and brace_count > 0:
        line = self.lines[self.index].strip()
        
        # Проверяем закрывающую скобку
        if line == "}" or line.endswith("}"):
            brace_count -= 1
            if brace_count == 0:
                self.index += 1
                break
        
        # Парсим содержимое
        blocks.append(...)
        self.index += 1
```

---

## 📌 Правила отладки

### Правило 1: Защита от бесконечных циклов

```python
# ❌ ОПАСНЫЙ КОД
while self.index < len(self.lines):
    if some_condition:
        blocks.append(...)
        # Забыли self.index += 1
    # ← БЕСКОНЕЧНЫЙ ЦИКЛ!

# ✅ БЕЗОПАСНЫЙ КОД
max_iterations = len(self.lines)
iterations = 0

while self.index < len(self.lines) and iterations < max_iterations:
    iterations += 1
    # ...
    if some_condition:
        blocks.append(...)
        self.index += 1  # ← ВСЕГДА!
```

### Правило 2: Проверяй граничные случаи

```python
# ✅ Добавить проверки:
if self.index >= len(self.lines):
    break  # Конец файла

if self.index + 1 >= len(self.lines):
    break  # Нет следующей строки

if not self.lines[self.index]:
    continue  # Пустая строка
```

### Правило 3: Используй отладку

```python
# В методе парсинга добавьте:
DEBUG = True  # Toggle для отладки

if DEBUG:
    print(f"[{self.__class__.__name__}] index={self.index}, line={line[:50]}...")
    print(f"  → Распознано как: {type(block).__name__}")

# В конце добавьте ассерт:
assert self.index > start_index, "Index не сдвинулся! Бесконечный цикл"
```

---

## 🔍 Как найти ошибку

### Шаг 1: Найти строку в Markdown

```markdown
# Это заголовок

Этот текст не должен пропадать.

- Пункт списка
```

### Шаг 2: Написать тест

```python
def test_my_issue():
    md = """# Это заголовок

Этот текст не должен пропадать.

- Пункт списка"""
    
    blocks = MarkdownParser(md).parse()
    
    # Проверяем результат
    assert len(blocks) == 3
    assert isinstance(blocks[0], HeadingBlock)
    assert isinstance(blocks[1], TextBlock)
    assert isinstance(blocks[2], ListBlock)
    
    assert blocks[1].text == "Этот текст не должен пропадать."
```

### Шаг 3: Запустить тест

```bash
pytest tests/test_my_issue.py -v
```

### Шаг 4: Добавить отладку

```python
class MarkdownParser:
    def parse(self) -> List[BaseBlock]:
        blocks = []
        DEBUG = True
        
        while self.index < len(self.lines):
            line = self.lines[self.index].strip()
            
            if DEBUG and line:
                print(f"Line {self.index}: {line[:50]}")
            
            if not line:
                self.index += 1
                continue
            
            # ... парсинг ...
        
        if DEBUG:
            print(f"Collected {len(blocks)} blocks")
            for i, block in enumerate(blocks):
                print(f"  {i}: {type(block).__name__}")
```

### Шаг 5: Исправить код

```python
# Добавить условие остановки для TextBlock
elif self._is_text_block(line):
    blocks.append(self._parse_text_block())
```

---

## 📋 Чек-лист исправления

- [ ] Написать тест, что падает ❌
- [ ] Добавить отладку (print)
- [ ] Найти точку разлома в коде
- [ ] Исправить логику
- [ ] Проверить тест проходит ✅
- [ ] Добавить граничные случаи
- [ ] Запустить все тесты: `pytest`
- [ ] Обновить документацию

---

## 🚀 Быстрые исправления

### "Индекс выходит за границы"
```python
# Всегда проверяйте
if self.index >= len(self.lines):
    break
```

### "Бесконечный цикл"
```python
# Всегда передвигайте индекс
self.index += 1
```

### "Данные дублируются"
```python
# Не добавляйте дважды
blocks.append(block)
self.index += 1  # ОДИН раз!
```

### "NoneType ошибка"
```python
# Проверяйте перед использованием
if block and block.text:
    ...
```

---

**Используйте этот гайд при отладке! 🔧**

