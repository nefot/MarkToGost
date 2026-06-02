import re
from typing import List
from MarkToGost.parser.blocks import *
from MarkToGost.utils.document_helpers import (
    is_md_table_row, is_md_table_separator, split_md_table_row, normalize_table_caption
)
from MarkToGost.utils.html_table_parser import parse_html_table
from MarkToGost.utils.toc import get_heading_level_from_number
# Импорт в начале файла
from html.parser import HTMLParser
from MarkToGost.parser.blocks import HtmlTableBlock, HtmlTableRow, HtmlTableCell


class MarkdownParser:
    """Парсер Markdown в блоки"""

    def __init__(self, md_text: str):
        self.lines = md_text.splitlines()
        self.index = 0

    def parse(self) -> List[BaseBlock]:
        """Основной метод парсинга"""
        blocks = []

        while self.index < len(self.lines):
            line = self.lines[self.index].strip()

            if not line:
                self.index += 1
                continue

            if self._is_section_start(line):
                blocks.append(self._parse_section())
            elif self._is_formula_block(line):
                blocks.append(self._parse_formula_block())
            elif self._is_heading(line):
                blocks.append(self._parse_heading(line))
            elif self._is_image(line):
                blocks.append(self._parse_image(line))
            elif self._is_table_start():
                blocks.append(self._parse_table())
            elif self._is_list_start():
                blocks.append(self._parse_list())
            elif self._is_code_block_start(line):
                blocks.append(self._parse_code_block())
            elif self._is_html_table_start(line):
                blocks.append(self._parse_html_table())

            else:
                blocks.append(self._parse_text_block())

        return blocks

    def _is_heading(self, line: str) -> bool:
        return line.startswith("#")

    def _is_section_start(self, line: str) -> bool:
        """Проверка начала раздела: [//]: # (ID), [//]: ## (ID), [//]: ### (ID), [//]: #### (ID) и т.д."""
        return bool(re.match(r'^\[//\]:\s*#{1,6}\s*\([^)]+\)', line))

    def _is_formula_block(self, line):
        """Проверка начала блока формулы ($$)"""
        return line.startswith("$$")

    def _parse_formula_block(self):
        line = self.lines[self.index].strip()
        line = line[2:].strip()  # Убираем начальные $$

        buffer = []
        while True:
            if "$$" in line:
                before, _, _ = line.partition("$$")
                before = before.strip()
                if before:
                    buffer.append(before)
                break
            else:
                if line:
                    buffer.append(line)
                self.index += 1
                if self.index >= len(self.lines):
                    break
                line = self.lines[self.index].strip()

        self.index += 1  # Пропускаем строку с $$

        formula_text = "\\n".join(buffer).strip()

        # Собираем строки пояснения "где ..."
        explanation_lines = []
        while self.index < len(self.lines):
            peek = self.lines[self.index].strip()

            if not peek:
                # Пустая строка — смотрим вперёд
                if self.index + 1 < len(self.lines):
                    next_line = self.lines[self.index + 1].strip()
                    if next_line.startswith('$') or next_line.lower().startswith('где'):
                        self.index += 1
                        continue
                break

            # Стоп-условия
            if (peek.startswith('#') or peek.startswith('$$') or
                    peek.startswith('![') or peek.startswith('[//]:')):
                break

            explanation_lines.append(peek)
            self.index += 1

        explanation = "\\n".join(explanation_lines).strip() if explanation_lines else None

        return FormulaBlock(latex=formula_text, explanation=explanation)

    def _extract_section_info(self, line: str) -> tuple[Optional[str], int, bool]:
        """Извлечение ID раздела, уровня заголовка и параметров из строки [//]: #### (ID)[params]
        Возвращает (section_id, heading_level, add_page_breaks)

        Примеры:
        [//]: # (ID) -> level=4, add_page_breaks=True
        [//]: ### (ID)[new_page=false] -> level=3, add_page_breaks=False
        [//]: #### (ID)[new_page=false] -> level=4, add_page_breaks=False
        """
        # Regex для: [//]: ### (ID)[опциональные параметры]
        # Используем более безопасный regex без catastrophic backtracking
        match = re.match(r'^\[//\]:\s*(#{1,6})\s*\(([^)]*)\)(?:\[([^\]]*)\])?', line)
        if match:
            hashes = match.group(1)
            section_id = match.group(2).strip() if match.group(2) else ""
            params_str = match.group(3) or ""

            if not section_id:
                return None, 4, True

            heading_level = len(hashes)
            # Для обратной совместимости: если используется [//]: # (старый синтаксис),
            # то по умолчанию используем уровень 4 (####)
            if heading_level == 1:
                heading_level = 4

            # Парсим параметры [new_page=false]
            add_page_breaks = True
            if params_str:
                # Проверяем флаг new_page=false
                params_lower = params_str.lower()
                if "new_page=false" in params_lower or "new_page=no" in params_lower:
                    add_page_breaks = False

            return section_id, heading_level, add_page_breaks
        return None, 4, True  # По умолчанию уровень 4, с page breaks

    def _is_section_with_content(self, line: str) -> bool:
        """Проверка, содержит ли строка открывающую скобку раздела: [//]: #### (ID) {"""
        return bool(re.match(r'^\[//\]:\s*#{1,6}\s*\([^)]+\)\s*\{', line))

    def _parse_section(self) -> Section:
        """Парсинг раздела с блоками или без"""
        current_line = self.lines[self.index].strip()
        section_id, heading_level, add_page_breaks = self._extract_section_info(current_line)

        if not section_id:
            self.index += 1
            return Section(section_id="", blocks=[], heading_level=4, add_page_breaks=True)

        has_content = self._is_section_with_content(current_line)

        if has_content:
            # Раздел с блоками: [//]: #### (ID)[params] { ... }
            return self._parse_section_with_braces(section_id, heading_level, add_page_breaks)
        else:
            # Раздел без блоков: [//]: #### (ID)[params]
            self.index += 1
            return Section(section_id=section_id, blocks=[], heading_level=heading_level,
                           add_page_breaks=add_page_breaks)

    def _parse_section_with_braces(self, section_id: str, heading_level: int, add_page_breaks: bool) -> Section:
        """Парсинг раздела с фигурными скобками [//]: #### (ID)[params] { ... }"""
        self.index += 1  # Пропускаем строку открытия раздела

        section_blocks = []
        brace_count = 1  # Одна скобка уже открылась
        max_iterations = len(self.lines) * 2  # Защита от бесконечного цикла
        iterations = 0

        # Проверяем, есть ли закрывающая скобка на той же строке
        current_line = self.lines[self.index - 1].strip()
        if current_line.endswith("}"):
            # Раздел пустой или содержит только скобки на одной строке
            brace_count -= 1
            # Пропускаем пустые линии после закрытия
            while self.index < len(self.lines):
                line = self.lines[self.index].strip()
                if not line or line == "}":
                    self.index += 1
                else:
                    break
            return Section(section_id=section_id, blocks=section_blocks, heading_level=heading_level,
                           add_page_breaks=add_page_breaks)

        # Собираем блоки до закрывающей скобки
        while self.index < len(self.lines) and brace_count > 0 and iterations < max_iterations:
            iterations += 1
            line = self.lines[self.index].strip()

            if not line:
                self.index += 1
                continue

            # Проверяем закрывающую скобку
            if line == "}" or line.endswith("}"):
                brace_count -= 1
                if brace_count == 0:
                    self.index += 1
                    break
                else:
                    self.index += 1
                    continue

            # Проверяем открывающую скобку (вложенные разделы)
            if self._is_section_with_content(line):
                # Вложенный раздел - пропускаем его целиком
                self.index += 1
                continue

            # Парсим блоки внутри раздела
            if self._is_heading(line):
                section_blocks.append(self._parse_heading(line))
            elif self._is_formula_block(line):
                section_blocks.append(self._parse_formula_block())
            elif self._is_image(line):
                section_blocks.append(self._parse_image(line))
            elif self._is_table_start():
                section_blocks.append(self._parse_table())
            elif self._is_list_start():
                section_blocks.append(self._parse_list())
            elif self._is_code_block_start(line):
                section_blocks.append(self._parse_code_block())
            elif not self._is_section_start(line):  # Не начинаем новый раздел внутри раздела
                section_blocks.append(self._parse_text_block())
            else:
                self.index += 1

        # Если не нашли закрывающую скобку, просто продолжаем со следующей строки
        if brace_count > 0:
            pass  # Раздел был незавершён, но мы всё равно продолжаем

        return Section(section_id=section_id, blocks=section_blocks, heading_level=heading_level,
                       add_page_breaks=add_page_breaks)

    def _is_image(self, line: str) -> bool:
        return line.startswith("![")

    def _is_table_start(self) -> bool:
        """Проверка начала таблицы"""
        if self.index + 1 >= len(self.lines):
            return False

        current = self.lines[self.index].strip()
        next_line = self.lines[self.index + 1].strip()

        return is_md_table_row(current) and is_md_table_separator(next_line)

    def _is_list_start(self) -> bool:
        """Проверка начала списка"""
        line = self.lines[self.index].strip()
        return re.match(r'^\s*[-–—*+]\s+', line) or re.match(r'^\s*\d+[.)]\s+', line)

    def _is_code_block_start(self, line: str) -> bool:
        """Проверка начала блока кода"""
        return line.startswith("```")

    def _parse_heading(self, line: str) -> HeadingBlock:
        text = line.lstrip('#').strip()
        level = get_heading_level_from_number(text)
        self.index += 1
        return HeadingBlock(text=text, level=level)

    def _parse_image(self, line: str) -> ImageBlock:
        # Поддержка формата: ![caption](path){#id}
        match = re.match(r'!\[(.*?)\]\((.*?)\)(?:\{#([^}]+)\})?', line)
        caption = match.group(1) if match else ""
        path = match.group(2) if match else ""
        img_id = match.group(3) if match and match.group(3) else None
        self.index += 1
        return ImageBlock(path=path, caption=caption, img_id=img_id)

    def _parse_table(self) -> TableBlock:
        table_caption = None

        # Проверяем caption перед таблицей
        if self.index > 0:
            prev_line = self.lines[self.index - 1].strip()
            prev_caption = normalize_table_caption(prev_line)
            if prev_caption:
                table_caption = prev_caption

        header = self.lines[self.index].strip()
        separator = self.lines[self.index + 1].strip()
        self.index += 2

        table_lines = [header]

        # Собираем все строки таблицы
        while self.index < len(self.lines):
            candidate = self.lines[self.index].strip()
            if candidate and is_md_table_row(candidate):
                table_lines.append(candidate)
                self.index += 1
            else:
                break

        return TableBlock(rows=table_lines, caption=table_caption)

    def _parse_list(self) -> ListBlock:
        """Парсинг списка"""
        items = []
        first_line = self.lines[self.index].strip()
        ordered = bool(re.match(r'^\s*\d+[.)]\s+', first_line))

        while self.index < len(self.lines):
            line = self.lines[self.index].strip()
            if not line:
                break

            if ordered:
                m = re.match(r'^\d+[.)]\s+(.*)', line)
                if m:
                    items.append(m.group(1))
                    self.index += 1
                else:
                    break
            else:
                m = re.match(r'^\s*[-–—*+]\s+(.*)', line)
                if m:
                    items.append(m.group(1))
                    self.index += 1
                else:
                    break

        return ListBlock(items=items, ordered=ordered)

    def _parse_code_block(self) -> CodeBlock:
        """Парсинг блока кода"""
        self.index += 1  # Пропускаем строку с ```

        code_lines = []
        while self.index < len(self.lines):
            line = self.lines[self.index]
            if line.strip() == "```":
                break
            code_lines.append(line)
            self.index += 1

        self.index += 1  # Пропускаем строку с ```

        code = "\n".join(code_lines).strip()
        language = ""

        # Попытка извлечь язык из первой строки кода
        if code_lines and ":" in code_lines[0]:
            parts = code_lines[0].split(":", 1)
            language = parts[0].strip()
            code = parts[1].strip() if len(parts) > 1 else ""

        return CodeBlock(code=code, language=language)

    def _parse_text_block(self) -> TextBlock:
        """Парсинг блока обычного текста"""
        buffer = []
        max_iterations = len(self.lines)  # Защита от бесконечного цикла
        iterations = 0

        start_index = self.index  # Сохраняем начальный индекс

        while self.index < len(self.lines) and iterations < max_iterations:
            iterations += 1
            line = self.lines[self.index].strip()

            # Проверяем условия остановки
            if not line or line.startswith("#") or line.startswith("!["):
                break
            if line == "}" or line.endswith("}"):  # Закрывающая скобка раздела
                break
            if line.startswith("$$"):  # Формула
                break
            if self._is_table_start() or self._is_list_start():
                break
            if self._is_section_start(line):
                break

            buffer.append(line)
            self.index += 1

        # Если никакие из проверок не прошли и индекс не изменился, вынуждаем движение вперёд
        if self.index == start_index and self.index < len(self.lines):
            # Странная строка, которая не соответствует никаким условиям
            buffer.append(self.lines[self.index].strip())
            self.index += 1

        return TextBlock(text=" ".join(buffer))

    def _is_html_table_start(self, line: str) -> bool:
        return line.strip().lower().startswith('<table')

    def _parse_html_table(self) -> HtmlTableBlock:
        html_lines = []
        while self.index < len(self.lines):
            l = self.lines[self.index]
            html_lines.append(l)
            self.index += 1
            if '</table>' in l.lower():
                break
        return parse_html_table('\n'.join(html_lines))
# За пределами класса MarkdownParser (в том же файле):
