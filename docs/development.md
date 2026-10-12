# Разработка

## Окружение

```bash
python -m venv .venv
.venv\Scripts\activate          # Windows; в Linux/macOS — source .venv/bin/activate
pip install -r requirements-dev.txt
```

## Проверка изменений

```bash
python -m pytest                       # модульные тесты (MarkToGost/tests)
python tools/check_blocks.py all       # сборка всех примеров examples/blocks → output/blocks
python tools/check_blocks.py --list    # список примеров
python tools/check_blocks.py 05_TableBlock
```

Тесты проверяют разбор Markdown и шапки, профили, структуру, вставные листы, форматирование,
таблицы и сборку документа без Word. Примеры из `examples/blocks` нужно открыть в Word и
посмотреть глазами — особенно после изменений в рендерерах.

Для итоговой проверки вёрстки соберите реальный документ с доводкой в Word:

```bash
python -m MarkToGost.main --new курсовая проверка.md
python -m MarkToGost.main проверка.md --pdf
```

## Профили и титульные листы

- Профили — `MarkToGost/profiles/*.yaml`, формат описан в
  [document-types.md](document-types.md#свой-тип-документа). Загрузчик и слияние профилей —
  `MarkToGost/profiles/__init__.py`.
- Новый параметр оформления: атрибут в `config.DocumentSettings`, его русское имя в
  `profiles.FORMAT_KEYS`, значение в `profiles/gost.yaml` и строка в таблице параметров
  в `docs/document-types.md`. Если у параметра есть флаг командной строки — добавьте его в
  `main.CLI_TO_SETTINGS`.
- Стандартные шаблоны листов создаёт `tools/make_title_templates.py`: только недостающие,
  `--force` — все заново (ручные правки в Word при этом теряются). Новый стандартный лист —
  функция в этом скрипте, запись в `TEMPLATES` и описание в разделе `листы` профиля.

## Как добавить новый вид блока

Пример — цитата, строки которой начинаются с `> `.

1. Блок в `parser/blocks.py`:

   ```python
   @dataclass
   class QuoteBlock(BaseBlock):
       text: str
   ```

2. Распознавание и разбор в `parser/markdown_parser.py`:

   ```python
   def _is_quote(self, line: str) -> bool:
       return line.startswith("> ")

   def _parse_quote(self) -> QuoteBlock:
       lines = []
       while self.index < len(self.lines) and self.lines[self.index].strip().startswith("> "):
           lines.append(self.lines[self.index].strip()[2:])
           self.index += 1
       return QuoteBlock(text=" ".join(lines))
   ```

   Проверку нужно добавить **в оба места**: в `parse()` (блоки верхнего уровня) и в
   `_parse_section_with_braces()` (блоки внутри разделов) — до ветки обычного текста.
   Если блок может прервать абзац, добавьте условие остановки и в `_parse_text_block()`.

3. Рендерер `renderer/quote.py`:

   ```python
   def render_quote_block(renderer, block: QuoteBlock):
       p = renderer.doc.add_paragraph()
       for part_text, is_italic in apply_italic_formatting(block.text):
           set_run_font(p.add_run(part_text), size_pt=DocumentSettings.FONT_SIZE_PT, italic=True)
       renderer._mark_content()
   ```

   Используйте `apply_italic_formatting`, чтобы в новом блоке работали курсив и `\_`,
   и вызывайте `renderer._mark_content()` после добавления содержимого: так применяется
   отложенный разрыв страницы.

4. Регистрация в словаре `_renderers` в `renderer/document_renderer.py`.

5. Тесты в `MarkToGost/tests/` и пример `examples/blocks/NN_QuoteBlock.md`.

## Соглашения

- Всё, что отличается между видами документов, — в профилях, а не в коде: пакет остаётся
  универсальным. Вспомогательные сборщики конкретного документа живут рядом с этим документом.
- Параметры оформления читаются из `DocumentSettings`; профиль и шапка меняют их только на
  время сборки одного документа (`main.settings_scope`).
- Рендерер не вставляет разрывы страниц пустыми абзацами — только `renderer._safe_page_break()`
  или свойства абзаца (`page_break_before`, `keep_with_next`).
- Текст, который видит читатель документа и пользователь в консоли, — на русском; комментарии
  в коде объясняют причину, а не пересказывают код.
- Изменения, влияющие на вёрстку, проверяйте на реальном документе в Word, а не только тестами.
