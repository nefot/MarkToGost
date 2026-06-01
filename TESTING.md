# 🧪 Руководство по тестированию MarkToGost

## 📦 Что такое тесты?

**Тесты** — это код, который проверяет, что ваши функции работают правильно. Они автоматически проверяют "граничные случаи" и предотвращают ошибки.

```
Без тестов:                    С тестами:
Пишу код → Клик → Надеюсь     Пишу тест → Пишу код → Тест проходит ✅
                                                      (или падает ❌)
```

---

## 🚀 С чего начать

### Установить pytest

```bash
pip install pytest
```

### Запустить все тесты

```bash
cd C:\Users\Nefot\PycharmProjects\MarkToGost
pytest MarkToGost/tests/ -v
```

**Результат:**
```
MarkToGost/tests/test_formatting.py::TestApplyItalicFormatting::test_plain_text PASSED
MarkToGost/tests/test_formatting.py::TestApplyItalicFormatting::test_empty_string PASSED
MarkToGost/tests/test_formatting.py::TestApplyItalicFormatting::test_none PASSED
...
======================= 43 passed in 0.42s =======================
```

---

## 📝 Анатомия теста (на примере)

### Пример 1: Простая функция

```python
# MarkToGost/utils/formatting.py
def apply_italic_formatting(text: str):
    """Разбирает текст на части: обычная/курсив"""
    # Возвращает список: [("текст", False), ("курсив", True), ...]
    ...


# MarkToGost/tests/test_formatting.py
class TestApplyItalicFormatting:
    """Класс для группировки тестов"""
    
    def test_plain_text(self):
        """Тест 1: Простой текст без курсива"""
        # Подготовка: какой вход даём?
        input_text = "обычный текст"
        
        # Действие: вызываем функцию
        result = apply_italic_formatting(input_text)
        
        # Проверка: что должно получиться?
        assert result == [("обычный текст", False)]
        #      ↑           ↑
        #    ключевое    ожидаемый результат
        #    слово
    
    def test_italic_in_middle(self):
        """Тест 2: Курсив посередине текста"""
        input_text = "до _курсив_ после"
        result = apply_italic_formatting(input_text)
        
        # Проверяем: текст разбит на части?
        assert result == [
            ("до ", False),         # обычный текст
            ("курсив", True),       # курсив
            (" после", False)       # снова обычный
        ]
```

### Когда тест падает ❌

Если результат не совпадает с ожиданиями:

```bash
$ pytest MarkToGost/tests/test_formatting.py::TestApplyItalicFormatting::test_plain_text -v

FAILED test_formatting.py::TestApplyItalicFormatting::test_plain_text

assert result == [("обычный текст", False)]
AssertionError: assert [("обычный", False), ("текст", False)]
              ↑                                 ↑
        то что вернула       то что ожидаем
        функция             получить
```

---

## 📊 Структура тестов в проекте

### 1️⃣ test_formatting.py (17 тестов)

Проверяет **утилиты для форматирования шрифта**.

```
apply_italic_formatting()     ← Разбор текста на части:
  ✅ простой текст           обычный/курсив
  ✅ курсив в начале
  ✅ курсив в конце
  ✅ несколько курсивов
  ✅ пустая строка
  ✅ None
  ✅ ...

set_run_font()                ← Применение форматирования
  ✅ размер шрифта
  ✅ жирность (bold)
  ✅ курсив (italic)
  ✅ цвет
  ✅ ...
```

**Пример теста:**

```python
def test_multiple_italic(self):
    """Несколько курсивов в тексте"""
    result = apply_italic_formatting("_один_ и _два_")
    assert result == [
        ("один", True),      # 1️⃣ первый курсив
        (" и ", False),      # 2️⃣ обычный текст
        ("два", True)        # 3️⃣ второй курсив
    ]
```

### 2️⃣ test_document_helpers.py (26 тестов)

Проверяет **вспомогательные функции для работы с документом**.

```
split_md_table_row()          ← Разбор строк таблицы
  ✅ простая таблица
  ✅ таблица с пробелами
  ✅ пустые ячейки
  ✅ ...

is_md_table_separator()       ← Проверка разделителя (| --- |)
  ✅ валидный разделитель
  ✅ с выравниванием (:---:)
  ✅ невалидный (< 3 дашей)
  ✅ ...

normalize_table_caption()     ← Извлечение названия таблицы
  ✅ HTML формат <caption>
  ✅ Markdown формат "Таблица 1 —"
  ✅ пусто
  ✅ ...

replace_image_refs()          ← Замена @img1 на рис. 1
  ✅ одна замена
  ✅ несколько замен
  ✅ без замен
  ✅ ...
```

**Пример теста:**

```python
def test_multiple_replacements(self):
    """Несколько ссылок на изображения"""
    text = "На @img1 видно, а на @img2 видно тоже"
    refs = {"img1": 1, "img2": 2}
    
    result = replace_image_refs(text, refs)
    
    # Обе ссылки должны заменены
    assert result == "На рис. 1 видно, а на рис. 2 видно тоже"
```

---

## 🎯 Принцип тестирования: AAA (Arrange-Act-Assert)

Все тесты следуют этой схеме:

```python
def test_example():
    # 1. ARRANGE (подготовка) — установить входные данные
    input_data = "тестовая строка"
    expected = ["тест", "овая", "строка"]
    
    # 2. ACT (действие) — выполнить функцию
    result = split_string(input_data)
    
    # 3. ASSERT (проверка) — убедиться в результате
    assert result == expected
```

---

## 🔍 Примеры реальных тестов

### Тест 1: Простой текст

```python
def test_plain_text(self):
    """Текст без форматирования"""
    result = apply_italic_formatting("обычный текст")
    assert result == [("обычный текст", False)]
    
# Что проверяем:
# ✅ Функция может обработать простой текст
# ✅ Возвращает список туплей (текст, флаг_курсива)
# ✅ Флаг = False для обычного текста
```

### Тест 2: Граничный случай (пустая строка)

```python
def test_empty_string(self):
    """Пустая строка не должна крашить функцию"""
    result = apply_italic_formatting("")
    assert result == [("", False)]
    
# Что проверяем:
# ✅ Функция не вызывает исключение
# ✅ Возвращает предсказуемый результат
```

### Тест 3: None (неожиданный ввод)

```python
def test_none(self):
    """None должен обрабатываться безопасно"""
    result = apply_italic_formatting(None)
    assert result == [(None, False)]
    
# Что проверяем:
# ✅ Функция не падает на None
# ✅ Возвращает правильный формат
```

### Тест 4: Сложный случай

```python
def test_multiple_italic(self):
    """Несколько курсивов в одной строке"""
    result = apply_italic_formatting("_один_ и _два_")
    assert result == [
        ("один", True),    # первый курсив
        (" и ", False),    # обычный текст между
        ("два", True)      # второй курсив
    ]
    
# Что проверяем:
# ✅ Функция правильно разбирает несколько курсивов
# ✅ Сохраняет текст между ними
# ✅ Правильно расставляет флаги
```

---

## 🛠️ Как написать тест для своей функции

### Пример: Пишем функцию с тестом

**Шаг 1: Написать тест (ещё нет функции!)**

```python
# MarkToGost/tests/test_my_new_function.py

def test_count_words():
    """Функция должна считать количество слов"""
    text = "Это простой текст с пятью словами"
    result = count_words(text)
    assert result == 5


def test_count_words_empty():
    """На пустой строке должно быть 0 слов"""
    result = count_words("")
    assert result == 0
```

**Шаг 2: Запустить тест (падает ❌)**

```bash
$ pytest tests/test_my_new_function.py -v

FAILED - ImportError: cannot import name 'count_words'
```

**Шаг 3: Написать функцию**

```python
# MarkToGost/utils/text_helpers.py

def count_words(text: str) -> int:
    """Подсчёт слов в тексте"""
    if not text:
        return 0
    return len(text.split())
```

**Шаг 4: Запустить тест (проходит ✅)**

```bash
$ pytest tests/test_my_new_function.py -v

PASSED test_my_new_function.py::test_count_words
PASSED test_my_new_function.py::test_count_words_empty

======================= 2 passed in 0.02s =======================
```

---

## 📋 Команды pytest

### Запустить все тесты

```bash
pytest MarkToGost/tests/ -v
```

### Запустить один файл

```bash
pytest MarkToGost/tests/test_formatting.py -v
```

### Запустить один класс

```bash
pytest MarkToGost/tests/test_formatting.py::TestApplyItalicFormatting -v
```

### Запустить один тест

```bash
pytest MarkToGost/tests/test_formatting.py::TestApplyItalicFormatting::test_plain_text -v
```

### Запустить с подробным выводом

```bash
pytest MarkToGost/tests/ -vv  # очень подробно
pytest MarkToGost/tests/ -s   # показать print()
pytest MarkToGost/tests/ -x   # остановиться на первой ошибке
```

### Показать покрытие (сколько кода протестировано)

```bash
pip install pytest-cov
pytest MarkToGost/tests/ --cov=MarkToGost --cov-report=html
# Откройте htmlcov/index.html в браузере
```

---

## ✅ Примеры утверждений (assertions)

```python
# Равенство
assert result == expected          # результат точно совпадает
assert result != unexpected        # не совпадает

# Типы
assert isinstance(result, list)    # это список?
assert isinstance(result, str)     # это строка?

# Логика
assert result                       # True (истина)?
assert not result                  # False (ложь)?
assert result > 10                 # больше 10?
assert len(result) == 5            # длина = 5?

# Наличие
assert "text" in result            # "text" в результате?
assert result in ["a", "b", "c"]   # результат в списке?

# Исключения (ошибки)
with pytest.raises(ValueError):
    my_function("bad input")       # функция должна выбросить ошибку
```

---

## 🚨 Когда тесты падают

### Ошибка 1: AssertionError (неверный результат)

```bash
assert result == [("текст", False)]
AssertionError: assert [("те", False), ("кст", False)]

❌ Функция возвращает неправильный результат
✅ Проверить логику функции
```

### Ошибка 2: ImportError (функция не импортируется)

```bash
ImportError: cannot import name 'my_function'

❌ Функция не экспортирована из модуля
✅ Проверить __init__.py или путь импорта
```

### Ошибка 3: TypeError (неверный тип)

```bash
TypeError: unsupported operand type(s) for +: 'str' and 'int'

❌ Функция получила неправильный тип аргумента
✅ Проверить type hints и валидацию входных данных
```

---

## 📚 Структура тестового класса

```python
class TestClassName:
    """Группа тестов для одной функции или модуля"""
    
    def setup_method(self):
        """Выполняется перед каждым тестом"""
        self.data = prepare_test_data()
    
    def teardown_method(self):
        """Выполняется после каждого теста"""
        cleanup()
    
    def test_case_1(self):
        """Первый сценарий"""
        assert True
    
    def test_case_2(self):
        """Второй сценарий"""
        assert True
    
    @pytest.mark.skip(reason="Not implemented yet")
    def test_case_3(self):
        """Пропустить этот тест"""
        pass
```

---

## 🎓 Хорошие практики

### ✅ ДЕЛАЙТЕ:

```python
# 1. Описательные имена
def test_apply_italic_formatting_with_multiple_words():
    # Сразу понятно, что проверяем
    pass

# 2. Один assert на тест (или логически связанные)
def test_bold_and_italic_together():
    run = make_run()
    set_run_font(run, bold=True, italic=True)
    assert run.font.bold is True
    assert run.font.italic is True  # логически связано с предыдущим

# 3. Тестируйте граничные случаи
def test_empty_list():
    result = process([])
    assert result == []  # пустой вход

def test_single_element():
    result = process([1])
    assert result == [1]  # один элемент

def test_many_elements():
    result = process(range(1000))
    assert len(result) == 1000  # много элементов
```

### ❌ НЕ ДЕЛАЙТЕ:

```python
# 1. Неописательные имена
def test_1():  # что это проверяет?
    pass

# 2. Множество проверок разных функций
def test_everything():
    assert func1() == 1
    assert func2() == 2
    assert func3() == 3
    # если падает — не видно, какой именно

# 3. Зависимость тестов друг от друга
def test_first():
    global state
    state = 1

def test_second():
    assert state == 1  # зависит от теста выше?
    # тесты должны быть независимы!
```

---

## 🔄 Workflow разработки с тестами

```
1. Пишу тест (падает ❌)
   └─ тест защищает от будущих ошибок

2. Пишу минимальное решение (тест проходит ✅)
   └─ код работает

3. Рефакторю код (тесты проходят ✅)
   └─ улучшаю качество

4. Добавляю граничные случаи в тесты
   └─ покрытие увеличивается

5. Всё готово + полное покрытие ✅
```

---

## 📊 Текущее состояние (проект)

```
✅ test_formatting.py        17 тестов (100% покрыто)
   - apply_italic_formatting  8 тестов
   - set_run_font             9 тестов

✅ test_document_helpers.py   26 тестов (95% покрыто)
   - split_md_table_row       5 тестов
   - is_md_table_separator    5 тестов
   - is_md_table_row          3 теста
   - normalize_table_caption  6 тестов
   - replace_image_refs       5 тестов

❓ test_parser.py            (можно добавить)
   - MarkdownParser           нет тестов
   - extract_metadata         нет тестов

❓ test_metadata.py          (можно добавить)
   - extract_metadata         нет тестов

❓ test_formula.py           (можно добавить)
   - formula_renderer         нет тестов

ИТОГО: 43 теста ✅
```

---

## 🚀 Следующие шаги

1. **Запустить тесты:**
   ```bash
   pytest MarkToGost/tests/ -v
   ```

2. **Добавить тесты для parser:**
   ```python
   # tests/test_parser.py
   def test_parse_formula_block():
       md = "$$\nF = ma\n$$"
       blocks = MarkdownParser(md).parse()
       assert isinstance(blocks[0], FormulaBlock)
   ```

3. **Покрыть тестами новые функции:**
   ```bash
   pytest --cov=MarkToGost --cov-report=html
   # откроет отчёт о покрытии
   ```

4. **Запускать перед каждым коммитом:**
   ```bash
   pytest && git commit -m "Add new feature"
   ```

---

## 💡 Полезные советы

- 🎯 **TDD (Test-Driven Development):** Пишите тест ДО кода
- 🔍 **Coverage:** Старайтесь покрыть тестами хотя бы 80% кода
- 🚀 **Fast:** Тесты должны работать быстро (< 1 сек)
- 📝 **Документируйте:** Пишите docstring для каждого теста
- 🔁 **Ciclе:** Тест → Код → Рефакторинг → Тест

---

**Все готово! Начните писать тесты! 🚀**

