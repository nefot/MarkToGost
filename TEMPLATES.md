#  Шаблоны для расширения

Используйте эти шаблоны как отправную точку для добавления новых функций.

---

## Шаблон 1️⃣: Новый элемент Markdown (полный цикл)

### 1. Блок данных (`parser/blocks.py`)

```python
from dataclasses import dataclass
from typing import Optional, List

@dataclass
class YourNewBlock(BaseBlock):
    """Описание вашего блока
    
    Это может быть:
    - Цитаты
    - Инфобоксы
    - Адмониции (примечания)
    - Таблицы стилей
    - Галереи изображений
    - И многое другое
    """
    content: str
    metadata: Optional[dict] = None
    is_active: bool = True
```

### 2. Парсер (`parser/markdown_parser.py`)

```python
class MarkdownParser:
    
    def _is_your_element(self, line: str) -> bool:
        """Проверка маркера вашего элемента"""
        # Примеры маркеров:
        # > для цитат
        # ::: для адмониций
        # @@ для адаптивных элементов
        return line.startswith(":::YOUR_MARKER")
    
    def _parse_your_element(self) -> YourNewBlock:
        """Парсинг вашего элемента"""
        line = self.lines[self.index].strip()
        content_lines = []
        
        # Собираем контент до конца маркера
        self.index += 1
        while self.index < len(self.lines):
            line = self.lines[self.index]
            
            # Стоп-условия
            if self._is_section_start(line) or line.startswith("#"):
                break
            
            content_lines.append(line)
            self.index += 1
        
        content = "\n".join(content_lines).strip()
        
        return YourNewBlock(
            content=content,
            metadata={"type": "your_element"},
            is_active=True
        )
    
    def parse(self) -> List[BaseBlock]:
        """Не забудьте добавить проверку в основной цикл!"""
        blocks = []
        
        while self.index < len(self.lines):
            line = self.lines[self.index].strip()
            
            if not line:
                self.index += 1
                continue
            
            # ... другие проверки ...
            elif self._is_your_element(line):  # ← ДОБАВИТЬ
                blocks.append(self._parse_your_element())
            # ...
        
        return blocks
```

### 3. Рендерер (`renderer/your_new_element.py`)

```python
"""Рендеринг вашего элемента в DOCX"""

from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

from MarkToGost.config import DocumentSettings
from MarkToGost.utils.formatting import set_run_font, set_paragraph_formatting


def render_your_new_block(doc, block, set_paragraph_formatting):
    """
    Рендеринг вашего элемента
    
    Args:
        doc: Document объект python-docx
        block: YourNewBlock instance
        set_paragraph_formatting: Функция для форматирования
    """
    # Добавить параграф
    p = doc.add_paragraph()
    
    # Форматирование
    set_paragraph_formatting(
        p,
        align=WD_ALIGN_PARAGRAPH.LEFT,
        left_indent=Cm(0.5),  # Отступ слева
        first_line_indent=Cm(0),
        line_spacing=DocumentSettings.LINE_SPACING
    )
    
    # Добавить содержимое
    run = p.add_run(block.content)
    set_run_font(
        run,
        size_pt=DocumentSettings.FONT_SIZE_PT,
        bold=False,
        italic=False
    )
    
    # Отступ после элемента
    doc.add_paragraph()
```

### 4. Интеграция (`renderer/document_renderer.py`)

```python
# В импортах добавить:
from MarkToGost.renderer.your_new_element import render_your_new_block
from MarkToGost.parser.blocks import YourNewBlock

# В класс DocumentRenderer добавить:
def render_block(self, block: BaseBlock):
    """Рендеринг блока"""
    if isinstance(block, Section):
        self._render_section_block(block)
    elif isinstance(block, YourNewBlock):  # ← ДОБАВИТЬ
        self._render_your_new_block(block)
    # ... остальные блоки ...

def _render_your_new_block(self, block):
    """Делегируем в render_your_new_block"""
    render_your_new_block(self.doc, block, set_paragraph_formatting)
    self._mark_content()
```

### 5. Тесты (`tests/test_your_new_element.py`)

```python
"""Тесты для вашего элемента"""

import pytest
from MarkToGost.parser.markdown_parser import MarkdownParser
from MarkToGost.parser.blocks import YourNewBlock


class TestParseYourElement:
    """Тесты парсинга"""
    
    def test_simple_element(self):
        """Простой элемент"""
        md = ":::YOUR_MARKER\nТестовый контент"
        blocks = MarkdownParser(md).parse()
        assert len(blocks) == 1
        assert isinstance(blocks[0], YourNewBlock)
        assert "Тестовый" in blocks[0].content
    
    def test_multiline_content(self):
        """Многострочный контент"""
        md = ":::YOUR_MARKER\nСтрока 1\nСтрока 2\nСтрока 3"
        blocks = MarkdownParser(md).parse()
        assert isinstance(blocks[0], YourNewBlock)
        assert len(blocks[0].content.split("\n")) == 3
    
    def test_empty_element(self):
        """Пустой элемент"""
        md = ":::YOUR_MARKER\n"
        blocks = MarkdownParser(md).parse()
        assert len(blocks) == 1
        # Может быть пусто или непусто - зависит от вашей логики


class TestRenderYourElement:
    """Тесты рендеринга"""
    
    def test_renders_to_paragraph(self):
        """Элемент рендерится в параграф"""
        from docx import Document
        from MarkToGost.renderer.your_new_element import render_your_new_block
        
        doc = Document()
        block = YourNewBlock(content="Test content")
        render_your_new_block(doc, block, lambda *args, **kwargs: None)
        
        # Проверяем что параграф добавлен
        assert len(doc.paragraphs) > 0
        assert "Test content" in doc.paragraphs[0].text
```

---

## Шаблон 2️⃣: Новая утилита (простая)

### `utils/my_utils.py`

```python
"""Утилиты для работы с (вашей функцией)"""

from typing import List, Dict, Optional, Tuple
import re


def my_utility_function(
    input_data: str,
    param1: Optional[str] = None,
    param2: int = 10
) -> Dict[str, any]:
    """
    Краткое описание что делает функция
    
    Длинное описание если нужно.
    Может быть несколько абзацев.
    
    Args:
        input_data: Входные данные
        param1: Опциональный параметр (default: None)
        param2: Числовой параметр (default: 10)
        
    Returns:
        Словарь с результатами
        
    Raises:
        ValueError: Если input_data пусто
        TypeError: Если param2 не число
        
    Examples:
        >>> result = my_utility_function("test")
        >>> result['status']
        'success'
        
        >>> result = my_utility_function("test", param2=20)
        >>> result['count']
        20
    """
    # Валидация входов
    if not input_data:
        raise ValueError("input_data не может быть пустым")
    
    if not isinstance(param2, int):
        raise TypeError("param2 должен быть числом")
    
    # Основная логика
    result = {
        "status": "success",
        "input": input_data,
        "param1": param1,
        "param2": param2,
        "output": input_data.upper()  # Пример
    }
    
    return result


def helper_function(data: List[str]) -> int:
    """Вспомогательная функция"""
    return len(data)
```

### `tests/test_my_utils.py`

```python
"""Тесты для my_utils"""

import pytest
from MarkToGost.utils.my_utils import my_utility_function, helper_function


class TestMyUtilityFunction:
    
    def test_success_case(self):
        """Успешный случай"""
        result = my_utility_function("test")
        assert result['status'] == 'success'
        assert result['output'] == 'TEST'
    
    def test_with_parameters(self):
        """С параметрами"""
        result = my_utility_function("test", param1="custom", param2=20)
        assert result['param1'] == "custom"
        assert result['param2'] == 20
    
    def test_empty_input_raises_error(self):
        """Пустой ввод вызывает ошибку"""
        with pytest.raises(ValueError, match="не может быть пустым"):
            my_utility_function("")
    
    def test_invalid_param_type(self):
        """Неправильный тип параметра"""
        with pytest.raises(TypeError, match="должен быть числом"):
            my_utility_function("test", param2="not_number")
    
    def test_none_param1(self):
        """param1 может быть None"""
        result = my_utility_function("test", param1=None)
        assert result['param1'] is None


class TestHelperFunction:
    
    def test_counts_items(self):
        """Считает элементы"""
        assert helper_function(["a", "b", "c"]) == 3
    
    def test_empty_list(self):
        """Пустой список"""
        assert helper_function([]) == 0
```

---

## Шаблон 3️⃣: Конфигурация (настройка)

### `config.py` (добавить)

```python
class MyFeatureSettings:
    """Настройки для (вашей функции)"""
    
    # Основные параметры
    ENABLED = True
    FEATURE_NAME = "MyFeature"
    VERSION = "1.0.0"
    
    # Рабочие параметры
    MAX_SIZE = 1000
    TIMEOUT_SECONDS = 30
    RETRY_COUNT = 3
    
    # Стили/внешний вид
    COLOR_PRIMARY = (0, 0, 0)
    COLOR_SECONDARY = (100, 100, 100)
    FONT_SIZE = 12
    
    # Пути и директории
    TEMP_DIR = "/tmp/marktosgost"
    CACHE_SIZE_MB = 100
    
    @classmethod
    def get_config_dict(cls) -> dict:
        """Получить всю конфигурацию как словарь"""
        return {
            k: v for k, v in vars(cls).items()
            if not k.startswith('_') and k.isupper()
        }
```

### Использование

```python
from MarkToGost.config import MyFeatureSettings

# Используем в коде
if MyFeatureSettings.ENABLED:
    max_size = MyFeatureSettings.MAX_SIZE
    # ...
```

---

## Шаблон 4️⃣: Класс с состоянием

### Пример: Счетчик/Статистика

```python
from dataclasses import dataclass, field
from typing import Dict, List


@dataclass
class DocumentStatistics:
    """Статистика документа (пример со счетчиком)"""
    
    # Счетчики
    total_words: int = 0
    total_chars: int = 0
    total_lines: int = 0
    
    # Коллекции
    unique_words: set = field(default_factory=set)
    word_frequencies: Dict[str, int] = field(default_factory=dict)
    
    # Флаги
    is_complete: bool = False
    
    def add_word(self, word: str) -> None:
        """Добавить слово в статистику"""
        self.total_words += 1
        self.unique_words.add(word.lower())
        self.word_frequencies[word.lower()] = \
            self.word_frequencies.get(word.lower(), 0) + 1
    
    def get_summary(self) -> Dict[str, any]:
        """Получить краткую статистику"""
        return {
            "total_words": self.total_words,
            "unique_words": len(self.unique_words),
            "avg_frequency": self.total_words / max(1, len(self.unique_words)),
            "is_complete": self.is_complete
        }
    
    def reset(self) -> None:
        """Сбросить статистику"""
        self.total_words = 0
        self.total_chars = 0
        self.total_lines = 0
        self.unique_words.clear()
        self.word_frequencies.clear()
        self.is_complete = False
```

### Использование

```python
stats = DocumentStatistics()
stats.add_word("hello")
stats.add_word("world")
stats.add_word("hello")
print(stats.get_summary())
# {
#   'total_words': 3,
#   'unique_words': 2,
#   'avg_frequency': 1.5,
#   'is_complete': False
# }
```

---

## Чеклист использования шаблонов

- [ ] Скопировали шаблон для вашего случая
- [ ] Заменили ALL_CAPS_NAMES на ваши имена
- [ ] Добавили импорты в нужные `__init__.py`
- [ ] Написали как минимум 3 теста
- [ ] Проверили что `pytest` проходит
- [ ] Проверили что нет `ImportError`
- [ ] Добавили docstrings со следующими: `Args`, `Returns`, `Examples`
- [ ] Использовали type hints (`str`, `int`, `Optional[str]` и т.д.)
- [ ] Структурировали код (функции < 50 строк)

---

**Готовые шаблоны помогут вам быстро начать! **
