# -*- coding: utf-8 -*-
"""
Пакетный конвертер Markdown -> DOCX по ГОСТ 7.32-2001 (новая версия)

Полный функционал из md_to_gost_docx_batch.py, но с улучшенной архитектурой:
- Модульная структура с разделением ответственности
- Легко читаемый и поддерживаемый код
- Полная поддержка всех элементов Markdown
- Корректные подписи к рисункам и таблицам
- Нумерация элементов
- Правильное форматирование по ГОСТ
"""

import os
import re
from pathlib import Path
from dataclasses import dataclass
from typing import List, Optional, Dict, Any

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import subprocess
import tempfile
from docx import Document
from copy import deepcopy
def reset_heading_styles(doc: Document):
    """Сбрасывает форматирование стилей Heading 2-9 до нужного вида"""
    for level in range(2, 10):
        style_name = f'Heading {level}'
        try:
            style = doc.styles[style_name]
        except KeyError:
            continue

        # Шрифт
        style.font.name = DocumentSettings.FONT_NAME
        style.font.size = Pt(DocumentSettings.FONT_SIZE_PT)
        style.font.bold = True
        style.font.italic = False
        style.font.all_caps = False
        style.font.small_caps = False
        style.font.color.rgb = RGBColor(0, 0, 0)

        # Абзац
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        style.paragraph_format.first_line_indent = None
        style.paragraph_format.left_indent = Cm(0)
        style.paragraph_format.space_before = Pt(12)
        style.paragraph_format.space_after = Pt(6)
        style.paragraph_format.line_spacing = DocumentSettings.LINE_SPACING

        # Сброс caps через XML (на случай если python-docx не перебивает)
        rPr = style.element.get_or_add_rPr()
        for tag in ('w:caps', 'w:smallCaps'):
            el = rPr.find(qn(tag))
            if el is not None:
                rPr.remove(el)
        # Явно выставляем w:caps val=0
        caps_el = OxmlElement('w:caps')
        caps_el.set(qn('w:val'), '0')
        rPr.append(caps_el)
def convert_inline_math(text: str) -> str:
    """
    Конвертирует inline math $...$ в читаемый текст.
    Убирает знаки доллара, оставляя содержимое.

    Примеры:
        $F$ → F
        $m_1$ → m₁  (нижний индекс через юникод)
        $6.674 \\times 10^{-11}$ → 6.674×10⁻¹¹
        $E_k$ → Eₖ
    """
    if not text:
        return text

    # Словари для конвертации индексов в юникод
    SUBSCRIPT_MAP = str.maketrans("0123456789aeinoruvxhklmnpst", "₀₁₂₃₄₅₆₇₈₉ₐₑᵢₙₒᵣᵤᵥₓₕₖₗₘₙₚₛₜ")
    SUPERSCRIPT_MAP = str.maketrans("0123456789+-=()ni", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿⁱ")

    def process_math(m: str) -> str:
        """Обрабатывает содержимое между знаками $"""
        # Замены LaTeX команд
        m = re.sub(r'\s*\\times\s*', '×', m)
        m = m.replace(r'\cdot', '·')
        m = m.replace(r'\pm', '±')
        m = m.replace(r'\leq', '≤')
        m = m.replace(r'\geq', '≥')
        m = m.replace(r'\neq', '≠')
        m = m.replace(r'\alpha', 'α')
        m = m.replace(r'\beta', 'β')
        m = m.replace(r'\gamma', 'γ')
        m = m.replace(r'\delta', 'δ')
        m = m.replace(r'\Delta', 'Δ')
        m = m.replace(r'\sigma', 'σ')
        m = m.replace(r'\rho', 'ρ')
        m = m.replace(r'\mu', 'μ')
        m = m.replace(r'\lambda', 'λ')
        m = m.replace(r'\omega', 'ω')
        m = m.replace(r'\Omega', 'Ω')
        m = m.replace(r'\pi', 'π')
        m = m.replace(r'\infty', '∞')
        m = m.replace(r'\frac', '')

        # Нижние индексы _x или _{xxx}
        def replace_sub(match):
            content = match.group(1) or match.group(2)
            return content.translate(SUBSCRIPT_MAP)

        m = re.sub(r'_\{([^}]+)\}|_([^{}\s])', replace_sub, m)

        # Верхние индексы ^x или ^{xxx}
        def replace_sup(match):
            content = match.group(1) or match.group(2)
            return content.translate(SUPERSCRIPT_MAP)

        m = re.sub(r'\^\{([^}]+)\}|\^([^{}\s])', replace_sup, m)

        # Убираем оставшиеся фигурные скобки
        m = m.replace('{', '').replace('}', '')

        return m.strip()

    # Заменяем $...$
    result = re.sub(r'\$([^$]+)\$', lambda match: process_math(match.group(1)), text)
    return result

# ================================
# КОНФИГУРАЦИЯ И НАСТРОЙКИ
# ================================
def render_formula_with_pandoc(latex: str) -> Document:
    """Генерирует DOCX с формулой через Pandoc или текстовый резервный вариант"""

    with tempfile.TemporaryDirectory() as tmpdir:
        md_path = os.path.join(tmpdir, "formula.md")
        docx_path = os.path.join(tmpdir, "formula.docx")

        # ВАЖНО: display math
        md_content = f"$${latex}$$"
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(md_content)

        try:
            subprocess.run([
                "pandoc",
                md_path,
                "-o",
                docx_path
            ], check=True, capture_output=True, timeout=10)

            return Document(docx_path)
        except (FileNotFoundError, subprocess.TimeoutExpired, subprocess.CalledProcessError):
            # Если pandoc не установлен или произошла ошибка, создаем документ с текстом формулы
            doc = Document()
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(f"${latex}$")
            run.font.size = Pt(DocumentSettings.FONT_SIZE_PT)
            return doc


def append_docx_content(src_doc: Document, dst_doc: Document):
    """Копирует содержимое одного docx в другой"""
    for element in src_doc.element.body:
        dst_doc.element.body.append(deepcopy(element))

from docx.oxml import parse_xml

class DocumentSettings:
    """Настройки документа"""
    FONT_NAME = "Times New Roman"
    FONT_SIZE_PT = 14
    LINE_SPACING = 1.5
    FIRST_LINE_INDENT_CM = 1.25
    # Размеры страницы A4
    PAGE_WIDTH_CM = 21.0
    PAGE_HEIGHT_CM = 29.7

    # Поля по ГОСТ 7.32-2001
    LEFT_MARGIN_CM = 3.0  # не менее 30 мм
    RIGHT_MARGIN_CM = 1.0  # не менее 10 мм
    TOP_MARGIN_CM = 2.0  # не менее 20 мм
    BOTTOM_MARGIN_CM = 2.0  # не менее 20 мм

    # Настройки изображений
    IMAGE_WIDTH_FRACTION = 0.70

    # Настройки подписей
    CAPTION_FONT_SIZE_PT = 12
    CAPTION_ITALIC = True

    # Настройки таблиц
    TABLE_FONT_SIZE_PT = 12


class CaptionSettings:
    """Настройки подписей к элементам"""
    FONT_SIZE = Pt(12)
    ITALIC = True


class TextSettings:
    """Настройки текста"""
    FONT_NAME = "Times New Roman"
    FONT_SIZE = Pt(14)
    LINE_SPACING = 1.5
    FIRST_LINE_INDENT = Cm(1.25)


# ================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ================================

def set_run_font(run, size_pt=14, bold=False, italic=False):
    """Установка шрифта для run"""
    run.font.name = DocumentSettings.FONT_NAME
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor(0, 0, 0)


def set_paragraph_formatting(paragraph, *,
                             align=None,
                             first_line_indent=None,
                             left_indent=None,
                             line_spacing=1.25,
                             space_before=0,
                             space_after=0):
    """Установка форматирования абзаца"""
    if align is not None:
        paragraph.alignment = align

    pf = paragraph.paragraph_format
    if first_line_indent is not None:
        pf.first_line_indent = first_line_indent
    if left_indent is not None:
        pf.left_indent = left_indent
    pf.line_spacing = line_spacing
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)


def apply_italic_formatting(text):
    """
    Обработка текста с курсивом через _текст_
    Возвращает список кортежей (текст, is_italic)
    """
    if not text or '_' not in text:
        return [(text, False)]

    result = []
    pattern = re.compile(r'_(.*?)_')
    last_end = 0

    for match in pattern.finditer(text):
        if match.start() > last_end:
            result.append((text[last_end:match.start()], False))
        result.append((match.group(1), True))
        last_end = match.end()

    if last_end < len(text):
        result.append((text[last_end:], False))

    return result


def set_table_borders(table):
    """Установка границ таблицы через XML"""
    tbl = table._tbl
    tblPr = tbl.tblPr

    # Удаляем старые границы
    for child in list(tblPr):
        if child.tag.split('}')[-1] == 'tblBorders':
            tblPr.remove(child)

    tblBorders = OxmlElement('w:tblBorders')
    for border_type in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        elm = OxmlElement(f'w:{border_type}')
        elm.set(qn('w:val'), 'single')
        elm.set(qn('w:sz'), '8')
        elm.set(qn('w:space'), '0')
        elm.set(qn('w:color'), '000000')
        tblBorders.append(elm)

    tblPr.append(tblBorders)


def set_repeat_table_header(row):
    """Установка повторения заголовка таблицы на каждой странице"""
    trPr = row._tr.get_or_add_trPr()
    tblHeader = OxmlElement('w:tblHeader')
    tblHeader.set(qn('w:val'), "true")
    trPr.append(tblHeader)


def add_page_number_centered(document):
    """Добавление номеров страниц в нижний колонтитул"""
    for section in document.sections:
        footer = section.footer
        for p in footer.paragraphs:
            p.clear()
        p = footer.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        fld = OxmlElement('w:fldSimple')
        fld.set(qn('w:instr'), 'PAGE')
        run = OxmlElement('w:r')
        run.append(fld)
        p._p.append(run)

        for r in p.runs:
            r.font.name = DocumentSettings.FONT_NAME
            r.font.size = Pt(12)
            r.font.color.rgb = RGBColor(0, 0, 0)


def compute_image_width_cm(doc, fraction=0.70):
    """Вычисление ширины изображения"""
    sec = doc.sections[0]
    try:
        page_width_cm = sec.page_width.cm
    except:
        page_width_cm = float(sec.page_width) / 360000.0

    left = sec.left_margin.cm
    right = sec.right_margin.cm
    avail = page_width_cm - left - right
    return Cm(avail * fraction)


def replace_image_refs(text: str, image_refs: Dict[str, int]) -> str:
    """Замена ссылок на изображения @img_id на их номера"""
    if not text or not image_refs:
        return text

    result = text
    for img_id, fig_num in image_refs.items():
        result = result.replace(f"@{img_id}", f"рис. {fig_num}")

    return result


# ================================
# РАБОТА С ТАБЛИЦАМИ
# ================================

def split_md_table_row(line):
    """Разбор строки markdown-таблицы"""
    s = line.strip()
    if s.startswith('|'):
        s = s[1:]
    if s.endswith('|'):
        s = s[:-1]

    cells = []
    buf = []
    escape = False

    for ch in s:
        if escape:
            buf.append(ch)
            escape = False
            continue
        if ch == '\\':
            escape = True
            continue
        if ch == '|':
            cells.append(''.join(buf).strip())
            buf = []
            continue
        buf.append(ch)

    cells.append(''.join(buf).strip())
    return cells


def is_md_table_separator(line):
    """Проверка разделителя таблицы"""
    if '|' not in line:
        return False

    parts = split_md_table_row(line)
    if len(parts) < 2:
        return False

    for part in parts:
        p = part.replace(' ', '')
        if not re.fullmatch(r':?-{3,}:?', p):
            return False
    return True


def is_md_table_row(line):
    """Проверка строки таблицы"""
    s = line.strip()
    return bool(s) and ('|' in s)


CAPTION_HTML_RE = re.compile(r'^\s*<caption>\s*(.*?)\s*</caption>\s*$', re.I)
CAPTION_MD_RE = re.compile(r'^\s*(?:Таблица|Table)\s*\d*\s*[—:-]\s*(.+?)\s*$', re.I)
CAPTION_PLN_RE = re.compile(r'^\s*(?:Название таблицы|Caption)\s*[:\-]\s*(.+?)\s*$', re.I)


def normalize_table_caption(text):
    """Извлечение названия таблицы"""
    if not text:
        return None

    text = re.sub(r'\s+', ' ', text).strip().rstrip('.')

    for regex in [CAPTION_HTML_RE, CAPTION_MD_RE, CAPTION_PLN_RE]:
        m = regex.match(text)
        if m:
            title = m.group(1).strip().rstrip('.')
            return title or None

    return None


def format_table_cell(cell, *, is_header=False, size_pt=12):
    """Форматирование ячейки таблицы"""
    cell.text = cell.text

    for p in cell.paragraphs:
        p.style = 'Normal'  # Изменено с None на 'Normal'
        p.paragraph_format.first_line_indent = None
        p.paragraph_format.left_indent = Cm(0)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0

        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if is_header else WD_ALIGN_PARAGRAPH.LEFT

        if not p.runs:
            p.add_run("")

        # Применяем курсив
        p.clear()
        for part_text, is_italic in apply_italic_formatting(cell.text):
            run = p.add_run(part_text)
            set_run_font(run, size_pt=size_pt, bold=is_header, italic=is_italic)

    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


# ================================
# БЛОКИ ДОКУМЕНТА
# ================================

@dataclass
class BaseBlock:
    """Базовый класс блока"""
    pass


@dataclass
class TextBlock(BaseBlock):
    """Блок обычного текста"""
    text: str


@dataclass
class HeadingBlock(BaseBlock):
    """Блок заголовка"""
    text: str
    level: int


@dataclass
@dataclass
class ImageBlock(BaseBlock):
    """Блок изображения"""
    path: str
    caption: str
    img_id: Optional[str] = None


@dataclass
class ListBlock(BaseBlock):
    """Блок списка"""
    items: List[str]
    ordered: bool = False


@dataclass
class TableBlock(BaseBlock):
    """Блок таблицы"""
    rows: List[str]
    caption: str = None


@dataclass
class CodeBlock(BaseBlock):
    """Блок кода"""
    code: str
    language: str = ""

@dataclass
class FormulaBlock(BaseBlock):
    """Блок формулы LaTeX (Pandoc + ГОСТ)"""
    latex: str
    formula_id: Optional[str] = None
    explanation: Optional[str] = None
    number: Optional[str] = None
@dataclass
class Section(BaseBlock):
    """Раздел документа с уникальным идентификатором и блоками"""
    section_id: str  # Уникальный идентификатор раздела (e.g. "Литература")
    blocks: List[BaseBlock]  # Блоки внутри раздела
    heading_level: int = 4  # Уровень заголовка раздела (2-6), по умолчанию 4 (####)
    add_page_breaks: bool = True  # Добавлять page breaks перед/после раздела (default=True)


# ================================
# ПАРСЕР MARKDOWN
# ================================

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


# ================================
# РЕНДЕРЕР ДОКУМЕНТА
# ================================

class DocumentRenderer:
    """Рендерер блоков в DOCX"""

    def __init__(self, doc: Document, toc_entries: List[tuple] = None, use_headings: bool = True):
        self.doc = doc
        self.table_counter = 1
        self.formula_counter = 1  # Счетчик формул
        self.figure_counter = 0
        self.image_map = {}
        self.image_refs = {}  # Для хранения ссылок на изображения по их ID
        self.formula_refs = {}  # Для хранения ссылок на формулы по их ID
        self.first_h1_seen = False
        self.first_h4_seen = False
        self.chapter_num = 0
        self.section_num = 0
        self.use_headings = use_headings
        self.subsection_num = 0
        self.toc_entries = toc_entries or []
        self._last_was_break = False  # Флаг для защиты от двойных разрывов

    def _is_document_start(self) -> bool:
        """Проверка начала документа"""
        return len(self.doc.paragraphs) == 0

    def _safe_page_break(self):
        """Не допускает двойных пустых страниц"""
        if not self._last_was_break:
            self.doc.add_page_break()
            self._last_was_break = True

    def _mark_content(self):
        """Сброс флага page break после добавления контента"""
        self._last_was_break = False

    def render_block(self, block: BaseBlock):
        """Рендеринг блока"""
        if isinstance(block, Section):
            self._render_section_block(block)
        elif isinstance(block, TextBlock):
            self._render_text_block(block)
        elif isinstance(block, HeadingBlock):
            self._render_heading_block(block)
        elif isinstance(block, ImageBlock):
            self._render_image_block(block)
        elif isinstance(block, ListBlock):
            self._render_list_block(block)
        elif isinstance(block, TableBlock):
            self._render_table_block(block)
        elif isinstance(block, CodeBlock):
            self._render_code_block(block)
        elif isinstance(block, FormulaBlock):
            self._render_formula_block(block)

    def _render_section_block(self, section: Section):
        """Рендеринг раздела с его блоками
        Строгое правило: раздел ВСЕГДА начинается и заканчивается разрывом (если включено)
        """
        if not section.add_page_breaks:
            # Если page breaks отключены, просто рендерим содержимое
            if section.section_id:
                if section.section_id == "ОГЛАВЛЕНИЕ":
                    # Обычный текст, НЕ Heading (чтобы не попало в TOC)
                    p = self.doc.add_paragraph()
                    # Явный сброс форматирования стиля
                    p.style.font.all_caps = False
                    p.style.font.bold = False
                    p.paragraph_format.alignment = None  # сбросим до явной установки ниже

                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

                    run = p.add_run(section.section_id)
                    set_run_font(run, size_pt=DocumentSettings.FONT_SIZE_PT, bold=True)

                    self._mark_content()
                else:
                    level = get_heading_level_from_number(section.section_id)
                    heading_block = HeadingBlock(text=section.section_id, level=level)
                    self._render_heading_block(heading_block)

            for block in section.blocks:
                self.render_block(block)
            return

        # 🔹 ВСЕГДА начинаем с новой страницы (кроме самого начала документа)
        if not self._is_document_start():
            self._safe_page_break()

        # 🔹 Заголовок раздела
        if section.section_id:
            if section.section_id == "ОГЛАВЛЕНИЕ":
                # Обычный текст, НЕ Heading (чтобы не попало в TOC)
                p = self.doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_paragraph_formatting(p, space_before=0, space_after=6)

                run = p.add_run(section.section_id)
                set_run_font(run, size_pt=DocumentSettings.FONT_SIZE_PT, bold=True)

                self._mark_content()
            else:
                level = get_heading_level_from_number(section.section_id)
                heading_block = HeadingBlock(text=section.section_id, level=level)
                self._render_heading_block(heading_block)

        # 🔹 Специальная обработка для оглавления
        if section.section_id == "ОГЛАВЛЕНИЕ":
            add_toc(self.doc)
            self._mark_content()

        # 🔹 Контент раздела
        for block in section.blocks:
            self.render_block(block)

        # 🔹 ВСЕГДА завершаем раздел разрывом страницы
        self._safe_page_break()

    def _render_formula_block(self, block: FormulaBlock):
        # --- Нумерация ---
        if not block.number:
            block.number = str(self.formula_counter)
            self.formula_counter += 1

        if block.formula_id:
            self.formula_refs[block.formula_id] = block.number

        # --- Получаем формулу через Pandoc ---
        formula_doc = render_formula_with_pandoc(block.latex)

        # --- Создаём таблицу 1x2 ---
        table = self.doc.add_table(rows=1, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False

        # ширины колонок
        sec = self.doc.sections[0]
        usable_width = sec.page_width.cm - sec.left_margin.cm - sec.right_margin.cm

        table.columns[0].width = Cm(usable_width * 0.85)
        table.columns[1].width = Cm(usable_width * 0.15)

        # --- Левая ячейка (формула) ---
        left_cell = table.cell(0, 0)
        left_cell.text = ""

        # вставляем формулу внутрь ячейки
        for el in formula_doc.element.body:
            left_cell._element.append(deepcopy(el))

        for p in left_cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # --- Правая ячейка (номер) ---
        right_cell = table.cell(0, 1)
        right_cell.text = ""

        p = right_cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        run = p.add_run(f"({block.number})")
        set_run_font(run, size_pt=DocumentSettings.FONT_SIZE_PT)

        # --- Убираем границы таблицы ---
        tbl = table._tbl
        tblPr = tbl.tblPr

        borders = OxmlElement('w:tblBorders')
        for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = OxmlElement(f'w:{edge}')
            el.set(qn('w:val'), 'nil')
            borders.append(el)

        tblPr.append(borders)

        # --- Пояснение ---
        if block.explanation:
            lines = block.explanation.strip().split('\\n')
            for raw_line in lines:
                raw_line = raw_line.strip()
                if not raw_line:
                    continue
                line_text = convert_inline_math(raw_line)
                exp_p = self.doc.add_paragraph()
                set_paragraph_formatting(
                    exp_p,
                    align=WD_ALIGN_PARAGRAPH.LEFT,
                    first_line_indent=Cm(DocumentSettings.FIRST_LINE_INDENT_CM),
                    line_spacing=DocumentSettings.LINE_SPACING,
                    space_before=0,
                    space_after=0
                )
                run = exp_p.add_run(line_text)
                set_run_font(run, size_pt=DocumentSettings.FONT_SIZE_PT)

        self.doc.add_paragraph()
        self._mark_content()
    def _render_formula(self, latex: str):
        """Старый метод для совместимости (упрощенный рендеринг)"""
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        try:
            omml = latex_to_omml(latex)
            p._p.append(omml)
        except:
            run = p.add_run(f"${latex}$")
            set_run_font(run, size_pt=DocumentSettings.FONT_SIZE_PT)

    def _render_text_block(self, block: TextBlock):
        """Рендеринг текстового блока с заменой ссылок на изображения"""
        # Заменяем ссылки @img_id на номера
        text = replace_image_refs(block.text, self.image_refs)

        p = self.doc.add_paragraph(text)
        set_paragraph_formatting(
            p,
            align=WD_ALIGN_PARAGRAPH.JUSTIFY,
            first_line_indent=Cm(DocumentSettings.FIRST_LINE_INDENT_CM),
            line_spacing=DocumentSettings.LINE_SPACING
        )

        # Применяем курсив
        p.clear()
        for part_text, is_italic in apply_italic_formatting(text):
            run = p.add_run(part_text)
            set_run_font(run, size_pt=DocumentSettings.FONT_SIZE_PT, italic=is_italic)

        self._mark_content()

    def _render_heading_block(self, block: HeadingBlock):
        if not self.use_headings:
            # Рендерим как обычный жирный текст без стиля Heading
            p = self.doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if block.level > 1 else WD_ALIGN_PARAGRAPH.CENTER
            set_paragraph_formatting(p, space_before=12, space_after=6, line_spacing=DocumentSettings.LINE_SPACING)
            text = block.text.upper() if block.level == 1 else block.text
            for part_text, is_italic in apply_italic_formatting(text):
                run = p.add_run(part_text)
                set_run_font(run, size_pt=DocumentSettings.FONT_SIZE_PT, bold=True, italic=is_italic)
            self._mark_content()
            return
        # ... остальной код метода без изменений
        level = min(block.level, 9)
        style_name = f'Heading {level}'
        p = self.doc.add_paragraph(style=style_name)

        if level == 1:
            set_paragraph_formatting(p, space_before=0, space_after=6)
        else:
            set_paragraph_formatting(p, space_before=12, space_after=6)

        p.paragraph_format.outline_level = level - 1

        # Принудительно задаём выравнивание через XML
        pPr = p._p.get_or_add_pPr()
        for old in pPr.findall(qn('w:jc')):
            pPr.remove(old)
        jc = OxmlElement('w:jc')
        jc.set(qn('w:val'), 'center' if level == 1 else 'left')
        pPr.append(jc)

        # Текст: uppercase только для H1
        text = block.text.upper() if level == 1 else block.text

        p.clear()
        for part_text, is_italic in apply_italic_formatting(text):
            run = p.add_run(part_text)
            set_run_font(run, size_pt=DocumentSettings.FONT_SIZE_PT, bold=True, italic=is_italic)
            # Принудительно отключаем caps через XML
            rPr = run._r.get_or_add_rPr()
            for old in rPr.findall(qn('w:caps')):
                rPr.remove(old)
            for old in rPr.findall(qn('w:smallCaps')):
                rPr.remove(old)
            caps = OxmlElement('w:caps')
            caps.set(qn('w:val'), '0')
            rPr.append(caps)

        self._mark_content()

    def _render_image_block(self, block: ImageBlock):
        """Рендеринг изображения"""
        img_path = block.path
        if not os.path.isabs(img_path):
            img_path = os.path.join(BASE_DIR, block.path)

        if os.path.exists(img_path):
            # Изображение
            pic_para = self.doc.add_paragraph()
            pic_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

            img_width_cm = compute_image_width_cm(self.doc, DocumentSettings.IMAGE_WIDTH_FRACTION)
            try:
                pic_para.add_run().add_picture(img_path, width=img_width_cm)
            except Exception:
                pic_para.add_run(f"[Ошибка вставки: {os.path.basename(img_path)}]")

            # Получаем номер изображения
            # Если img_id указан и был собран на первом проходе, используем сохраненный номер
            if block.img_id and block.img_id in self.image_refs:
                fig_num = self.image_refs[block.img_id]
            else:
                # Иначе увеличиваем счетчик
                self.figure_counter += 1
                fig_num = self.figure_counter
                # Сохраняем в image_refs, если есть ID
                if block.img_id:
                    self.image_refs[block.img_id] = fig_num

            fname = os.path.basename(block.path)
            caption_text = block.caption or ""

            if caption_text:
                caption_full = f"Рисунок {fig_num} - {caption_text.strip().capitalize()}"
            else:
                caption_full = f"Рисунок {fig_num}"

            cap_para = self.doc.add_paragraph(caption_full)
            cap_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_paragraph_formatting(cap_para, line_spacing=DocumentSettings.LINE_SPACING)

            # Применяем курсив
            cap_para.clear()
            for part_text, is_italic in apply_italic_formatting(caption_full):
                run = cap_para.add_run(part_text)
                set_run_font(run, size_pt=DocumentSettings.CAPTION_FONT_SIZE_PT,
                             italic=DocumentSettings.CAPTION_ITALIC or is_italic)

            self.image_map[fname] = str(fig_num)
            self._mark_content()
        else:
            p = self.doc.add_paragraph(f"[Изображение не найдено: {os.path.basename(img_path)}]")
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            self._mark_content()

    def _render_list_block(self, block: ListBlock):
        """Рендеринг списка"""
        for i, item in enumerate(block.items, 1):
            if block.ordered:
                text = f"{i}. {item}"
            else:
                text = f"– {item}"

            p = self.doc.add_paragraph()

            if block.ordered:
                set_paragraph_formatting(
                    p,
                    align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                    first_line_indent=Cm(DocumentSettings.FIRST_LINE_INDENT_CM),
                    line_spacing=DocumentSettings.LINE_SPACING
                )

                # Номер
                run = p.add_run(f"{i}. ")
                set_run_font(run, size_pt=DocumentSettings.FONT_SIZE_PT)

                # Текст с курсивом
                for part_text, is_italic in apply_italic_formatting(item):
                    run = p.add_run(part_text)
                    set_run_font(run, size_pt=DocumentSettings.FONT_SIZE_PT, italic=is_italic)

            else:
                set_paragraph_formatting(
                    p,
                    align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                    left_indent=Cm(DocumentSettings.FIRST_LINE_INDENT_CM * 2),
                    first_line_indent=Cm(-0.35),
                    line_spacing=DocumentSettings.LINE_SPACING,
                    space_before=0,
                    space_after=0 if i < len(block.items) else 6
                )

                # Маркер
                run = p.add_run("– ")
                set_run_font(run, size_pt=DocumentSettings.FONT_SIZE_PT)

                # Текст с курсивом
                for part_text, is_italic in apply_italic_formatting(item):
                    run = p.add_run(part_text)
                    set_run_font(run, size_pt=DocumentSettings.FONT_SIZE_PT, italic=is_italic)

        self._mark_content()

    def _render_code_block(self, block: CodeBlock):
        """Рендеринг блока кода"""
        code = block.code.strip()
        language = block.language.strip().lower()

        if language == "python":
            # Специфическое форматирование для Python
            code = re.sub(r'^\s*def\s+(\w+)\s*\(', r'    def \1(', code, flags=re.MULTILINE)
            code = re.sub(r'^\s*class\s+(\w+)\s*\(', r'    class \1(', code, flags=re.MULTILINE)
            code = re.sub(r'^\s*#', r'    #', code, flags=re.MULTILINE)
            code = re.sub(r'^\s*print\s*\(', r'    print(', code, flags=re.MULTILINE)
            code = re.sub(r'^\s*return\s+', r'        return ', code, flags=re.MULTILINE)
            code = re.sub(r'^\s*if\s+', r'    if ', code, flags=re.MULTILINE)
            code = re.sub(r'^\s*else\s+', r'    else ', code, flags=re.MULTILINE)
            code = re.sub(r'^\s*elif\s+', r'    elif ', code, flags=re.MULTILINE)
            code = re.sub(r'^\s*for\s+', r'    for ', code, flags=re.MULTILINE)
            code = re.sub(r'^\s*while\s+', r'    while ', code, flags=re.MULTILINE)
            code = re.sub(r'^\s*try\s+', r'    try:', code, flags=re.MULTILINE)
            code = re.sub(r'^\s*except\s+', r'    except:', code, flags=re.MULTILINE)
            code = re.sub(r'^\s*finally\s+', r'    finally:', code, flags=re.MULTILINE)
            code = re.sub(r'^\s*with\s+', r'    with ', code, flags=re.MULTILINE)
            code = re.sub(r'^\s*as\s+', r'    as ', code, flags=re.MULTILINE)
            code = re.sub(r'^\s*await\s+', r'    await ', code, flags=re.MULTILINE)
            code = re.sub(r'^\s*async\s+', r'    async ', code, flags=re.MULTILINE)

        # Общие правила для всех языков
        code = re.sub(r'^\s*//\s*', r'    // ', code, flags=re.MULTILINE)
        code = re.sub(r'^\s*#\s*', r'    # ', code, flags=re.MULTILINE)
        code = re.sub(r'^\s*;\s*', r'    ;', code, flags=re.MULTILINE)
        code = re.sub(r'^\s*{\s*', r'    {', code, flags=re.MULTILINE)
        code = re.sub(r'^\s*}\s*', r'    }', code, flags=re.MULTILINE)
        code = re.sub(r'^\s*\(\s*', r'    (', code, flags=re.MULTILINE)
        code = re.sub(r'^\s*\)\s*', r'    )', code, flags=re.MULTILINE)

        # Удаление пустых строк
        code = re.sub(r'^\s*\n', '', code, flags=re.MULTILINE)

        # Добавление блока кода
        p = self.doc.add_paragraph()
        p.add_run(code).font.name = 'Courier New'
        p.paragraph_format.left_indent = Cm(0.5)
        p.paragraph_format.right_indent = Cm(0.5)
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(6)

        self._mark_content()

    def _estimate_row_height_cm(self, row_data: list, col_width_cm: float, size_pt: int = 12) -> float:
        LINE_HEIGHT_CM = size_pt * 0.0353  # pt -> cm
        CELL_PADDING_CM = 0.15  # внутренние отступы ячейки
        SAFETY_FACTOR = 1.05  # минимальный запас

        # Кириллица шире латиницы — используем 0.025 вместо 0.022/0.030
        CHAR_WIDTH_CM = size_pt * 0.025

        max_lines = 1
        for cell_text in row_data:
            if not cell_text:
                continue
            chars_per_line = max(1, int(col_width_cm / CHAR_WIDTH_CM))
            lines = 0
            for paragraph in str(cell_text).split('\n'):
                lines += max(1, -(-len(paragraph) // chars_per_line))  # ceiling division
            max_lines = max(max_lines, lines)

        return max_lines * LINE_HEIGHT_CM * DocumentSettings.LINE_SPACING * SAFETY_FACTOR + CELL_PADDING_CM
    def _render_table_block(self, block: TableBlock):
        if not block.rows:
            return

        parsed_rows = [split_md_table_row(row) for row in block.rows]
        max_cols = max(len(r) for r in parsed_rows) if parsed_rows else 0
        if max_cols <= 0:
            return

        for row in parsed_rows:
            if len(row) < max_cols:
                row.extend([""] * (max_cols - len(row)))

        # Вычисляем доступную высоту страницы
        sec = self.doc.sections[0]
        usable_width_cm = sec.page_width.cm - sec.left_margin.cm - sec.right_margin.cm
        PAGE_MARGIN_CM = 0.8  # только колонтитул с номером страницы
        usable_height_cm = sec.page_height.cm - sec.top_margin.cm - sec.bottom_margin.cm - PAGE_MARGIN_CM
        col_width_cm = usable_width_cm / max_cols

        CAPTION_HEIGHT_CM = 0.8  # высота строки подписи "Таблица N"
        HEADER_HEIGHT_CM = self._estimate_row_height_cm(
            parsed_rows[0], col_width_cm, DocumentSettings.TABLE_FONT_SIZE_PT
        )

        # Разбиваем на chunks по реальной высоте
        chunks = []
        current_chunk = []
        # Первый chunk: вычитаем высоту подписи + заголовка
        current_height = CAPTION_HEIGHT_CM + HEADER_HEIGHT_CM
        header_row = parsed_rows[0]

        for row_data in parsed_rows:
            row_h = self._estimate_row_height_cm(
                row_data, col_width_cm, DocumentSettings.TABLE_FONT_SIZE_PT
            )

            if current_chunk and current_height + row_h > usable_height_cm:
                chunks.append(current_chunk)
                current_chunk = [header_row]  # повторяем заголовок в каждом chunk
                # Следующий chunk: вычитаем высоту подписи "Продолжение" + заголовка
                current_height = CAPTION_HEIGHT_CM + HEADER_HEIGHT_CM + row_h
            else:
                current_height += row_h
                current_chunk.append(row_data)

        if current_chunk:
            chunks.append(current_chunk)

        col_width = Cm(col_width_cm)

        for chunk_index, chunk in enumerate(chunks):
            # --- Подпись ---
            cap = self.doc.add_paragraph()
            cap.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            set_paragraph_formatting(
                cap,
                first_line_indent=None,
                left_indent=Cm(0),
                space_before=6,
                space_after=0,
                line_spacing=DocumentSettings.LINE_SPACING
            )

            if chunk_index == 0:
                caption_full = (
                    f"Таблица {self.table_counter} — {block.caption}"
                    if block.caption
                    else f"Таблица {self.table_counter}"
                )
            else:
                caption_full = f"Продолжение таблицы {self.table_counter}"

            cap.clear()
            for part_text, is_italic in apply_italic_formatting(caption_full):
                run = cap.add_run(part_text)
                set_run_font(run, size_pt=DocumentSettings.FONT_SIZE_PT, bold=False, italic=is_italic)

            # --- Таблица ---
            table = self.doc.add_table(rows=len(chunk), cols=max_cols)
            try:
                table.style = 'Table Grid'
            except Exception:
                pass
            table.alignment = WD_TABLE_ALIGNMENT.LEFT
            table.autofit = False

            for col_idx in range(max_cols):
                for row in table.rows:
                    row.cells[col_idx].width = col_width

            # Повтор заголовка средствами Word (на случай если наша оценка немного ошиблась)
            if chunk:
                set_repeat_table_header(table.rows[0])

            # --- Заполнение ячеек ---
            for r_idx, row_data in enumerate(chunk):
                row = table.rows[r_idx]

                # Запрет разрыва строки посередине
                trPr = row._tr.get_or_add_trPr()
                cantSplit = OxmlElement('w:cantSplit')
                cantSplit.set(qn('w:val'), '1')
                trPr.append(cantSplit)

                for c_idx, value in enumerate(row_data):
                    cell = row.cells[c_idx]
                    cell.text = ""
                    for p in cell.paragraphs:
                        p.clear()

                    p = cell.paragraphs[0]
                    p.style = None
                    p.paragraph_format.first_line_indent = Pt(0)
                    p.paragraph_format.left_indent = Pt(0)
                    p.paragraph_format.right_indent = Pt(0)
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    p.paragraph_format.line_spacing = DocumentSettings.LINE_SPACING
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if r_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT

                    for part_text, is_italic in apply_italic_formatting(value if value is not None else ""):
                        run = p.add_run(part_text)
                        set_run_font(
                            run,
                            size_pt=DocumentSettings.TABLE_FONT_SIZE_PT,
                            bold=r_idx == 0,
                            italic=is_italic
                        )

                    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

            set_table_borders(table)

            if chunk_index < len(chunks) - 1:
                self.doc.add_page_break()

        self.table_counter += 1
        self._mark_content()
def add_toc(document):
    """Добавление автоматического оглавления (TOC)"""
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_paragraph_formatting(paragraph, space_before=0, space_after=6)

    run = paragraph.add_run()

    # Начало поля
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'begin')
    run._r.append(fldChar)

    # Инструкция TOC
    instrText = OxmlElement('w:instrText')
    instrText.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    instrText.text = 'TOC \\o "1-4" \\h \\z \\u'
    run._r.append(instrText)

    # Разделитель
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'separate')
    run._r.append(fldChar)

    # Конец поля
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar)


# ================================
# ОСНОВНАЯ ЛОГИКА
# ================================

def get_heading_level_from_number(text: str) -> int:
    """
    Определяет уровень заголовка по нумерации:
    1. → 1
    1.1 → 2
    1.1.1 → 3
    """
    match = re.match(r'^(\d+(?:\.\d+)*)', text.strip())
    if not match:
        return 1  # если нет номера — считаем верхним уровнем
    
    number = match.group(1)
    level = number.count('.') + 1
    # Защита: Word поддерживает максимум Heading 1-9
    return min(level, 9)


def get_toc_level(text: str) -> int:
    """Определение уровня оглавления на основе нумерации текста заголовка"""
    return get_heading_level_from_number(text)


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IMAGES_DIR = BASE_DIR


def extract_metadata(md_text: str) -> Dict[str, Any]:
    """Извлечение метаданных из текста"""
    metadata = {
        'is_table': False,
        'teacher': '',
        'fio': '',
        'group': '',
        'image_map': {}
    }

    # Извлечение флагов


    prefix = re.escape('[//]:')
    m_headings = re.search(prefix + r"\s*#\s*use_headings\s*=\s*(true|false)", md_text, re.I)
    if m_headings:
        metadata['use_headings'] = m_headings.group(1).lower() != 'false'
    else:
        metadata['use_headings'] = True
    m_numerate = re.search(prefix + r"\s*#\s*numerate\s*=\s*(true|false)", md_text, re.I)
    if m_numerate:
        metadata['numerate'] = m_numerate.group(1).lower() != 'false'
    else:
        metadata['numerate'] = True  # по умолчанию нумерация включена
    m_flag = re.search(prefix + r"\s*#\s*is_table\s*=\s*(true|false)", md_text, re.I)
    if m_flag:
        metadata['is_table'] = m_flag.group(1).lower() == 'true'

    m_teacher = re.search(prefix + r"\s*#\s*teacher\s*=\s*\"?([^\"\n]+)\"?", md_text, re.I)
    if m_teacher:
        metadata['teacher'] = m_teacher.group(1).strip()

    # Извлечение ФИО и группы
    fio_match = re.search(r'__?ФИО__?[:\-]?\s*(.*)', md_text)
    group_match = re.search(r'__?Группа__?[:\-]?\s*(.*)', md_text)

    metadata['fio'] = fio_match.group(1).strip() if fio_match else ""
    metadata['group'] = group_match.group(1).strip() if group_match else ""

    # Карта изображений
    img_pattern = re.compile(r'!\[([^]]*)\]\(([^)\s]+)(?:\s"([^\"]*)")?\)')
    for m in img_pattern.finditer(md_text):
        alt = m.group(1).strip()
        path = m.group(2).strip()
        title = (m.group(3) or "").strip()
        key = os.path.basename(path)
        caption = alt or title or ""
        metadata['image_map'][key] = caption

    return metadata


def create_document(md_text: str) -> Document:
    """Создание документа из Markdown"""

    # Извлечение метаданных
    metadata = extract_metadata(md_text)

    # Парсинг ПЕРЕД очисткой (чтобы сохранить разделы)
    # Очищаем только служебные метаданные, но не разделы
    md_text_for_parsing = md_text
    md_text_for_parsing = re.sub(r'__?ФИО__?[:\-]?.*\n?', '', md_text_for_parsing)
    md_text_for_parsing = re.sub(r'__?Группа__?[:\-]?.*\n?', '', md_text_for_parsing)
    # Удаляем только служебные комментарии (с =), но НЕ разделы (без =)
    md_text_for_parsing = re.sub(r'\[//\]:\s*#\s*\w+\s*=.*\n?', '', md_text_for_parsing, flags=re.I)

    # Создание документа
    doc = Document()

    # Настройка страницы
    for s in doc.sections:
        s.page_height = Cm(DocumentSettings.PAGE_HEIGHT_CM)
        s.page_width = Cm(DocumentSettings.PAGE_WIDTH_CM)
        s.left_margin = Cm(DocumentSettings.LEFT_MARGIN_CM)
        s.right_margin = Cm(DocumentSettings.RIGHT_MARGIN_CM)
        s.top_margin = Cm(DocumentSettings.TOP_MARGIN_CM)
        s.bottom_margin = Cm(DocumentSettings.BOTTOM_MARGIN_CM)

    # Стиль Normal
    normal = doc.styles['Normal']
    normal.font.name = DocumentSettings.FONT_NAME
    normal.font.size = Pt(DocumentSettings.FONT_SIZE_PT)
    normal.paragraph_format.first_line_indent = Cm(DocumentSettings.FIRST_LINE_INDENT_CM)
    normal.paragraph_format.line_spacing = DocumentSettings.LINE_SPACING
    reset_heading_styles(doc)

    # Блок ФИО (если не таблица)
    if not metadata['is_table'] and (metadata['fio'] or metadata['group']):
        group_full = f"090302-{metadata['group']}" if metadata['group'] else ""
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run1 = p.add_run(f"Выполнил студент: {metadata['fio']}\n")
        set_run_font(run1, size_pt=DocumentSettings.FONT_SIZE_PT, bold=True)
        run2 = p.add_run(f"Группа: {group_full}")
        set_run_font(run2, size_pt=DocumentSettings.FONT_SIZE_PT, bold=True)

    # Парсинг и сбор заголовков для оглавления
    parser = MarkdownParser(md_text_for_parsing)
    blocks = parser.parse()

    toc_entries = []

    def collect_headings(block):
        if isinstance(block, Section):
            for b in block.blocks:
                collect_headings(b)
        elif isinstance(block, HeadingBlock):
            toc_level = get_toc_level(block.text)
            toc_entries.append((toc_level, block.text))

    for block in blocks:
        collect_headings(block)

    # Рендеринг
    renderer = DocumentRenderer(doc, toc_entries, use_headings=metadata['use_headings'])
    for block in blocks:
        renderer.render_block(block)

    # Таблица подписей (если is_table)
    if metadata['is_table']:
        doc.add_page_break()

        # Вычисление места для таблицы
        sec = doc.sections[-1]
        page_height_cm = float(sec.page_height) / 360000.0
        top_cm = float(sec.top_margin) / 360000.0
        bottom_cm = float(sec.bottom_margin) / 360000.0
        usable_height_cm = page_height_cm - top_cm - bottom_cm

        total_cm = (
                               DocumentSettings.PAGE_WIDTH_CM - DocumentSettings.LEFT_MARGIN_CM - DocumentSettings.RIGHT_MARGIN_CM) * 0.75
        col_w_cm = total_cm / 2.0

        # Простая таблица подписей
        table = doc.add_table(rows=2, cols=2)
        try:
            table.style = 'Table Grid'
        except Exception:
            pass
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False

        for col in table.columns:
            for cell in col.cells:
                cell.width = Cm(col_w_cm)

        table.cell(0, 0).text = 'Выполнил:'
        table.cell(0, 1).text = metadata['fio']
        table.cell(1, 0).text = 'Проверил:'
        table.cell(1, 1).text = metadata['teacher']

        for row_idx, row in enumerate(table.rows):
            if row_idx == 0:
                set_repeat_table_header(row)
            for cell in row.cells:
                for p in cell.paragraphs:
                    p.style = doc.styles['Normal']
                    p.paragraph_format.first_line_indent = None
                    p.paragraph_format.left_indent = Cm(0)
                    p.paragraph_format.line_spacing = 1.5
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    for run in p.runs:
                        set_run_font(run, size_pt=12)

        set_table_borders(table)

    # Номера страниц
    # Номера страниц
    if metadata['numerate']:
        add_page_number_centered(doc)

    return doc


def process_md_file(input_path: str):
    """Обработка одного файла"""
    filename = os.path.basename(input_path)
    name_wo_ext = os.path.splitext(filename)[0]
    output_path = os.path.join(BASE_DIR, f"{name_wo_ext}.docx")

    with open(input_path, "r", encoding="utf-8") as f:
        md_text = f.read()

    doc = create_document(md_text)
    doc.save(output_path)
    print(f"✅ {filename} → {output_path}")


if __name__ == "__main__":
    # Поиск всех MD файлов
    all_md = [
        f for f in os.listdir(BASE_DIR)
        if f.lower().endswith(".md")
    ]

    if not all_md:
        print("⚠️ Не найдено ни одного Markdown-файла.")
    else:
        print(f"🔍 Найдено файлов: {len(all_md)}")
        for file in all_md:
            process_md_file(os.path.join(BASE_DIR, file))
        print("🎉 Все файлы обработаны.")



