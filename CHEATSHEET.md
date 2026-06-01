# 🚀 Шпаргалка разработчика

## Структура слоев (иерархия импортов)

```
┌──────────────────────────────┐
│      main.py                 │  ← Точка входа, импортирует ВСЁ
└──────────────────────────────┘
            ↑
┌──────────────────────────────┐
│    renderer/                 │  ← Рендеринг блоков
│ document_renderer.py,        │
│ text.py, table.py, ...       │
└──────────────────────────────┘
            ↑
┌──────────────────────────────┐
│    parser/                   │  ← Парсинг MD → блоки
│ markdown_parser.py,          │
│ blocks.py, metadata.py       │
└──────────────────────────────┘
            ↑
┌──────────────────────────────┐
│    utils/                    │  ← Базовые полезные функции
│ formatting.py, xml_helpers   │
│ toc.py, document_helpers.py  │
└──────────────────────────────┘
            ↑
┌──────────────────────────────┐
│    config.py                 │  ← Настройки (нет импортов)
└──────────────────────────────┘
```

**☝️ Правило:** Каждый слой импортирует ТОЛЬКО слои НИЖЕ себя!

---

## 🎯 Добавить новый элемент Markdown

### 1. Блок (`parser/blocks.py`)
```python
@dataclass
class MyNewBlock(BaseBlock):
    """Описание"""
    field1: str
    field2: Optional[int] = None
```

### 2. Парсер (`parser/markdown_parser.py`)
```python
def _is_my_block(self, line: str) -> bool:
    return line.startswith("@@@")

def _parse_my_block(self) -> MyNewBlock:
    # Собрать данные
    self.index += 1
    return MyNewBlock(...)

# В parse():
elif self._is_my_block(line):
    blocks.append(self._parse_my_block())
```

### 3. Рендерер (`renderer/my_element.py`)
```python
def render_my_block(doc, block, set_paragraph_formatting):
    p = doc.add_paragraph()
    # Рендерим в doc
    set_paragraph_formatting(p, ...)
```

### 4. Интеграция (`renderer/document_renderer.py`)
```python
# В импортах:
from MarkToGost.renderer.my_element import render_my_block
from MarkToGost.parser.blocks import MyNewBlock

# В render_block():
elif isinstance(block, MyNewBlock):
    self._render_my_block(block)

# Добавить метод:
def _render_my_block(self, block):
    render_my_block(self.doc, block, set_paragraph_formatting)
    self._mark_content()
```

### 5. Тесты (`tests/test_my_element.py`)
```python
from MarkToGost.parser.markdown_parser import MarkdownParser
from MarkToGost.parser.blocks import MyNewBlock

def test_parse_my_element():
    md = "@@@\nТест"
    blocks = MarkdownParser(md).parse()
    assert isinstance(blocks[0], MyNewBlock)
```

---

## 🔧 Добавить вспомогательную функцию

### 1. Создать `utils/my_helpers.py`
```python
def my_helper_function(param: str) -> str:
    """Описание"""
    return result
```

### 2. Экспортировать `utils/__init__.py`
```python
from MarkToGost.utils.my_helpers import my_helper_function

__all__ = [
    # ... существующие ...
    "my_helper_function",
]
```

### 3. Написать тест `tests/test_my_helpers.py`
```python
def test_my_helper():
    result = my_helper_function("input")
    assert result == "expected"
```

### 4. Использовать где угодно
```python
from MarkToGost.utils import my_helper_function
result = my_helper_function("test")
```

---

## ⚙️ Добавить параметр конфигурации

### 1. В `config.py`
```python
class MySettings:
    MY_PARAM = "default_value"
    MY_NUMBER = 42
```

### 2. Используется через импорт
```python
from MarkToGost.config import MySettings

if MySettings.MY_PARAM == "something":
    ...
```

### 3. Тест `tests/test_config.py`
```python
from MarkToGost.config import MySettings

def test_my_setting():
    assert MySettings.MY_PARAM == "default_value"
```

---

## 📝 Типовой паттерн класса

```python
from dataclasses import dataclass
from typing import Optional, List

@dataclass
class MyClass:
    """Описание класса
    
    Attributes:
        required_field: Обязательное поле
        optional_field: Опциональное поле
    """
    required_field: str
    optional_field: Optional[str] = None
    list_field: List[str] = None
    
    def __post_init__(self):
        """Инициализация после создания"""
        if self.list_field is None:
            self.list_field = []
    
    def get_summary(self) -> str:
        """Получить краткую информацию"""
        return f"{self.required_field}: {self.optional_field}"
```

---

## 🧪 Написать хороший тест

```python
import pytest
from my_module import my_function

class TestMyFunction:
    """Тесты для my_function"""
    
    def test_normal_case(self):
        """Обычный случай"""
        result = my_function("input")
        assert result == "expected"
    
    def test_empty_input(self):
        """Пустой ввод"""
        result = my_function("")
        assert result == ""
    
    def test_invalid_input_raises_error(self):
        """Неправильный ввод вызывает ошибку"""
        with pytest.raises(ValueError):
            my_function(None)
    
    def test_large_input(self):
        """Большой ввод"""
        large_input = "x" * 10000
        result = my_function(large_input)
        assert len(result) > 0
```

---

## 🔍 Команды для проверки

```bash
# Все тесты
pytest MarkToGost/tests/ -v

# Конкретный тест
pytest MarkToGost/tests/test_formatting.py::TestApplyItalicFormatting::test_plain_text -v

# С покрытием
pytest MarkToGost/tests/ --cov=MarkToGost

# С фильтром по имени
pytest MarkToGost/tests/ -k "formatting" -v

# Обработка файла
python -m MarkToGost.main test.md

# Обработка всех файлов
python -m MarkToGost.main

# Проверка кода
python -m py_compile MarkToGost/**/*.py
```

---

## 📁 Быстрый поиск где что находится

| Что нужно | Где искать |
|-----------|-----------|
| Константы документа | `MarkToGost/config.py` |
| Элементы Markdown | `MarkToGost/parser/blocks.py` |
| Парсер | `MarkToGost/parser/markdown_parser.py` |
| Вывод в DOCX | `MarkToGost/renderer/document_renderer.py` |
| Работа с шрифтом | `MarkToGost/utils/formatting.py` |
| Работа с XML | `MarkToGost/utils/xml_helpers.py` |
| Точка входа | `MarkToGost/main.py` |
| Тесты | `MarkToGost/tests/test_*.py` |

---

## 🚨 Частые ошибки и как их избежать

### ❌ Циклический импорт
```python
# ПЛОХО
# utils/a.py
from MarkToGost.renderer.b import something

# renderer/b.py
from MarkToGost.utils.a import something
```

✅ РЕШЕНИЕ: Следите за иерархией! Вам поможет:
```bash
python -c "import MarkToGost.main"
```

### ❌ Не документирован код
```python
# ПЛОХО
def f(x):
    return x * 2

# ХОРОШО
def double(value: int) -> int:
    """Удваивает значение
    
    Args:
        value: Число для удвоения
        
    Returns:
        Удвоенное число
    """
    return value * 2
```

### ❌ Нет типов
```python
# ПЛОХО
def process(data):
    ...

# ХОРОШО
from typing import List, Dict, Optional

def process(data: List[Dict[str, str]]) -> Optional[str]:
    ...
```

### ❌ Функция делает слишком много
```python
# ПЛОХО
def process_and_save_and_log(doc, path):
    # 50 строк кода...

# ХОРОШО
def process(doc) -> dict:
    return result

def save(data: dict, path: str) -> bool:
    return success

def log_operation(msg: str) -> None:
    print(msg)
```

---

## 💾 Git Workflow

```bash
# Создать ветку
git checkout -b feature/new-feature

# Разработка + тесты
pytest MarkToGost/tests/ -v

# Коммитить
git add MarkToGost/
git commit -m "Add support for new-feature"

# Push
git push origin feature/new-feature

# Pull Request на GitHub
```

---

## 📞 Помощь при проблемах

```
ошибка                     решение
─────────────────────────────────────────────────────────────
ImportError                Проверить иерархию импортов
circular import            Переместить импорт ниже в слой stack
AttributeError             Убедиться, что объект создан правильно
TypeError                  Проверить типы аргументов
FileNotFoundError          Убедиться, что input/ и output/ существуют
```

---

**Удачи в разработке! 🚀**

