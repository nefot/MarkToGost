# MarkToGost — Структура и использование

##  Структура папок

```
MarkToGost/
├── input/                   #  Входные Markdown файлы
│   ├── Kursovaya.md
│   ├── test.md
│   └── философия.md
│
├── output/                  #  Результаты конвертации (DOCX файлы)
│
├── MarkToGost/             # Основной пакет
│   ├── main.py             # Точка входа
│   ├── config.py           # Настройки
│   ├── parser/
│   ├── renderer/
│   ├── utils/
│   └── tests/
│
└── ARCHITECTURE.md         # Документация архитектуры
```

##  Использование

### 1️⃣ Обработать ВСЕ файлы из папки `input/`

```bash
python -m MarkToGost.main
```

**Результат:** Все DOCX файлы появятся в папке `output/`

**Вывод:**
```
 Папка входных файлов: C:\...\MarkToGost\input
 Папка для сохранения: C:\...\MarkToGost\output
 Найдено файлов: 3

✅ input\Kursovaya.md → output\Kursovaya.docx
✅ input\test.md → output\test.docx
✅ input\философия.md → output\философия.docx

 Обработано: 3/3 файлов
```

### 2️⃣ Обработать конкретный файл

```bash
python -m MarkToGost.main test.md
```

**Результат:** Конвертируется только `input/test.md` → `output/test.docx`

### 3️⃣ Обработать файл с кастомным выходным именем

```bash
python -m MarkToGost.main Kursovaya.md --output Моя_курсовая.docx
```

**Результат:** `input/Kursovaya.md` → `output/Моя_курсовая.docx`

### 4️⃣ Справка

```bash
python -m MarkToGost.main --help
```

##  Рабочий процесс

1. **Поместите** ваши Markdown файлы в папку **`input/`**
   ```bash
   # Например:
   cp my_document.md input/
   ```

2. **Запустите конвертер**
   ```bash
   python -m MarkToGost.main
   # или для конкретного файла:
   python -m MarkToGost.main my_document.md
   ```

3. **Получите результаты** в папке **`output/`**
   ```bash
   ls output/
   # my_document.docx  ← готовый документ по ГОСТ
   ```

##  Запуск тестов

```bash
# Все тесты
pytest MarkToGost/tests/ -v

# Только конкретный тест
pytest MarkToGost/tests/test_formatting.py::TestApplyItalicFormatting -v
```

##  Примечания

- ✅ **Входные файлы** остаются в `input/` без изменений
- ✅ **Выходные файлы** создаются в `output/` с номерами страниц
- ✅ **Поддерживаются** все элементы Markdown: таблицы, формулы, изображения, списки, коды
- ✅ **ГОСТ 7.32-2001** — все требования соблюдены (шрифт, отступы, нумерация и т.д.)

##  Программный API

Если вы хотите использовать в своем коде:

```python
from MarkToGost.main import create_document

# Читаем Markdown
with open("input/my_file.md", "r", encoding="utf-8") as f:
    md_text = f.read()

# Создаем документ
doc = create_document(md_text)

# Сохраняем
doc.save("output/my_file.docx")

print("✅ Готово!")
```

## ❓ Часто задаваемые вопросы

**Q: Как изменить шрифт?**
A: Отредактируйте `MarkToGost/config.py`, поле `FONT_NAME`

**Q: Как изменить поля страницы?**
A: Отредактируйте `MarkToGost/config.py`, поля `*_MARGIN_CM`

**Q: Как добавить новый элемент Markdown?**
A: См. `ARCHITECTURE.md` — раздел "Как расширять"

**Q: Может ли быть ошибка при конвертации?**
A: Если файл поврежден, вы получите чуть подробную ошибку в консоли. Проверьте синтаксис Markdown

**Q: Что если файл слишком большой?**
A: Должно работать, но может медленнее. Если есть проблемы — разделите на несколько файлов

