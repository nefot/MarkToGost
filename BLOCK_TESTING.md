# 🎓 Руководство по тестированию отдельных блоков

## 📋 Быстрый старт

Каждый тип блока Markdown теперь можно проверять **отдельно** с помощью утилиты `test_blocks.py`.

### Показать все доступные блоки

```bash
python test_blocks.py --list
```

**Результат:**
```
📚 Доступные блоки для проверки:

1. 01_TextBlock      (Текстовые абзацы)
2. 02_HeadingBlock   (Заголовки)
3. 03_ListBlock      (Списки)
4. 04_CodeBlock      (Код)
5. 05_TableBlock     (Таблицы)
6. 06_FormulaBlock   (Формулы)
7. 07_ImageBlock     (Изображения)
8. 08_Section        (Секции)
```

---

## 🔍 Тестирование одного блока

### Проверить конкретный блок

```bash
# Проверить TextBlock
python test_blocks.py 01_TextBlock

# Проверить TableBlock
python test_blocks.py 05_TableBlock

# Проверить FormulaBlock
python test_blocks.py 06_FormulaBlock
```

**Результат:**
```
🔍 Проверка блока: 01_TextBlock

✅ 01_TextBlock
   📄 C:\Users\Nefot\PycharmProjects\MarkToGost\output\blocks\01_TextBlock.docx
   📊 Размер: 35.9 KB
```

**Что происходит:**
1. Читается файл `examples/blocks/01_TextBlock.md`
2. Парсится как Markdown
3. Создается DOCX документ
4. Сохраняется в `output/blocks/01_TextBlock.docx`

---

## 🚀 Тестирование всех блоков

### Проверить все блоки сразу

```bash
python test_blocks.py all
```

**Результат:**
```
🔍 Проверка всех блоков:

✅ 01_TextBlock      Размер: 35.9 KB
✅ 02_HeadingBlock   Размер: 35.8 KB
✅ 03_ListBlock      Размер: 35.9 KB
✅ 04_CodeBlock      Размер: 35.7 KB
✅ 05_TableBlock     Размер: 36.4 KB
✅ 06_FormulaBlock   Размер: 36.4 KB
✅ 07_ImageBlock     Размер: 35.7 KB
✅ 08_Section        Размер: 35.9 KB

📊 Результаты: 8/8 блоков ✅
🎉 Все блоки обработаны успешно!
```

---

## 📂 Структура файлов

```
examples/blocks/
├── 01_TextBlock.md          → output/blocks/01_TextBlock.docx
├── 02_HeadingBlock.md       → output/blocks/02_HeadingBlock.docx
├── 03_ListBlock.md          → output/blocks/03_ListBlock.docx
├── 04_CodeBlock.md          → output/blocks/04_CodeBlock.docx
├── 05_TableBlock.md         → output/blocks/05_TableBlock.docx
├── 06_FormulaBlock.md       → output/blocks/06_FormulaBlock.docx
├── 07_ImageBlock.md         → output/blocks/07_ImageBlock.docx
├── 08_Section.md            → output/blocks/08_Section.docx
└── README.md                (Полная документация)
```

---

## 📝 Описание каждого блока

### 1️⃣ TextBlock - `01_TextBlock.md`

**Проверяемые свойства:**
- ✅ Текст выравнен по ширине (justify)
- ✅ Отступ первой строки (1.25см)
- ✅ Межстрочный интервал (1.5)
- ✅ Поддержка курсива (_текст_)

**Команда:**
```bash
python test_blocks.py 01_TextBlock
```

---

### 2️⃣ HeadingBlock - `02_HeadingBlock.md`

**Проверяемые свойства:**
- ✅ Заголовки уровней 1-5
- ✅ H1 центрирован, остальные слева
- ✅ Жирный текст
- ✅ Отсутствие CAPS форматирования
- ✅ Курсив в заголовках

**Команда:**
```bash
python test_blocks.py 02_HeadingBlock
```

---

### 3️⃣ ListBlock - `03_ListBlock.md`

**Проверяемые свойства:**
- ✅ Упорядоченные списки (с номерами)
- ✅ Неупорядоченные списки (маркеры "–")
- ✅ Выступ маркеров/номеров влево
- ✅ Курсив в элементах списка

**Команда:**
```bash
python test_blocks.py 03_ListBlock
```

---

### 4️⃣ CodeBlock - `04_CodeBlock.md`

**Проверяемые свойства:**
- ✅ Моноширинный шрифт (Courier New)
- ✅ Форматирование для Python
- ✅ Форматирование для JavaScript
- ✅ Сохранение отступов

**Команда:**
```bash
python test_blocks.py 04_CodeBlock
```

---

### 5️⃣ TableBlock - `05_TableBlock.md`

**Проверяемые свойства:**
- ✅ Правильное количество колонок
- ✅ Заголовок центрирован и жирный
- ✅ Разбиение по страницам для больших таблиц
- ✅ Повторение заголовка на каждой странице
- ✅ Подпись "Таблица N — описание"
- ✅ Курсив в ячейках

**Команда:**
```bash
python test_blocks.py 05_TableBlock
```

---

### 6️⃣ FormulaBlock - `06_FormulaBlock.md`

**Проверяемые свойства:**
- ✅ Блочные формулы ($$...$$)
- ✅ Встроенные формулы ($...$)
- ✅ Нумерация формул
- ✅ Пояснения (где...)
- ✅ Преобразование в OMML

**Команда:**
```bash
python test_blocks.py 06_FormulaBlock
```

---

### 7️⃣ ImageBlock - `07_ImageBlock.md`

**Проверяемые свойства:**
- ✅ Вставка изображения
- ✅ Центрирование
- ✅ Нумерация рисунков
- ✅ Подпись под изображением
- ✅ Масштабирование (70% ширины)

**Команда:**
```bash
python test_blocks.py 07_ImageBlock
```

---

### 8️⃣ Section - `08_Section.md`

**Проверяемые свойства:**
- ✅ Начало новой страницы
- ✅ Нумерация разделов (1., 1.1, и т.д.)
- ✅ Вложенные разделы
- ✅ Разделы в оглавлении

**Команда:**
```bash
python test_blocks.py 08_Section
```

---

## 🛠️ Опции команды

```bash
# Тихий режим (без подробного вывода)
python test_blocks.py 01_TextBlock --quiet
python test_blocks.py all -q

# Показать справку
python test_blocks.py --help
```

---

## 📖 Как модифицировать примеры

### 1. Откройте файл блока

```bash
# Открыть в редакторе
notepad examples/blocks/01_TextBlock.md
```

### 2. Измените содержимое

```markdown
___ФИО___: Новое имя
___Группа___: 54321

# TextBlock - Новый заголовок

Новый текст с изменениями...
```

### 3. Переобработайте

```bash
python test_blocks.py 01_TextBlock
```

### 4. Откройте результат

```bash
# В Windows - откроется в Word/Libre Office
start output\blocks\01_TextBlock.docx

# В Linux/Mac
open output/blocks/01_TextBlock.docx
```

---

## 🧪 Экспресс-проверка перед коммитом

```bash
# Запустить все тесты (44 unit-теста)
pytest MarkToGost/tests/ -v

# Проверить все блоки (8 типов)
python test_blocks.py all

# Если всё ✅ - готово к коммиту!
git add .
git commit -m "Update: all blocks tested"
```

---

## 🐛 Отладка проблем

### Блок не обработался?

```bash
# Проверьте синтаксис
python test_blocks.py 01_TextBlock

# Если ошибка - смотрите вывод:
#   ImportError - проблема с импортом
#   FileNotFoundError - файл не найден
#   ValueError - неверный синтаксис
```

### Проверить файл блока

```python
# test_blocks_debug.py
from pathlib import Path

block_file = Path("examples/blocks/01_TextBlock.md")
print(f"Файл существует: {block_file.exists()}")
print(f"Размер: {block_file.stat().st_size} байт")

with open(block_file, "r", encoding="utf-8") as f:
    content = f.read()
    print(f"Содержимое ({len(content)} символов):")
    print(content[:200])
```

---

## 📊 Статистика файлов

| Блок | Размер | Содержимое |
|------|--------|-----------|
| 01_TextBlock | 35.9 KB | Текст + курсив |
| 02_HeadingBlock | 35.8 KB | Заголовки 1-5 уровней |
| 03_ListBlock | 35.9 KB | Упорядоч. + неупорядоч. списки |
| 04_CodeBlock | 35.7 KB | Python + JavaScript |
| 05_TableBlock | 36.4 KB | 2 таблицы |
| 06_FormulaBlock | 36.4 KB | Формулы LaTeX |
| 07_ImageBlock | 35.7 KB | 2 изображения |
| 08_Section | 35.9 KB | Секции с вложенностью |

---

## ✅ Результаты последний проверки

```
✅ 01_TextBlock      Размер: 35.9 KB
✅ 02_HeadingBlock   Размер: 35.8 KB
✅ 03_ListBlock      Размер: 35.9 KB
✅ 04_CodeBlock      Размер: 35.7 KB
✅ 05_TableBlock     Размер: 36.4 KB
✅ 06_FormulaBlock   Размер: 36.4 KB
✅ 07_ImageBlock     Размер: 35.7 KB
✅ 08_Section        Размер: 35.9 KB

📊 8/8 блоков ✅
🎉 Все работает!
```

---

## 🚀 Следующие шаги

1. **Добавить новый тип блока?**
   - Создайте файл `examples/blocks/09_NewBlock.md`
   - Запустите `python test_blocks.py 09_NewBlock`
   - Результат в `output/blocks/09_NewBlock.docx`

2. **Проверить отдельный файл вручную?**
   - Скопируйте в `input/`
   - Запустите `python -m MarkToGost.main <имя>`
   - Откройте в `output/`

3. **Расширить функциональность?**
   - Смотрите `CONTRIBUTING.md`
   - Добавляйте функции в отдельные модули
   - Пишите тесты

---

**Все готово к тестированию отдельных блоков! 🎉**

