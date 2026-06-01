# utils/toc.py
"""Функции для работы с оглавлением и стилями заголовков"""

import re
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor

from MarkToGost.config import DocumentSettings


def add_toc(document):
    """Добавление автоматического оглавления (TOC)"""
    from MarkToGost.utils.formatting import set_paragraph_formatting
    
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


def reset_heading_styles(doc):
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
        style.paragraph_format.left_indent = None
        style.paragraph_format.space_before = Pt(12)
        style.paragraph_format.space_after = Pt(6)
        style.paragraph_format.line_spacing = DocumentSettings.LINE_SPACING

        # Сброс caps через XML
        rPr = style.element.get_or_add_rPr()
        for tag in ('w:caps', 'w:smallCaps'):
            el = rPr.find(qn(tag))
            if el is not None:
                rPr.remove(el)
        # Явно выставляем w:caps val=0
        caps_el = OxmlElement('w:caps')
        caps_el.set(qn('w:val'), '0')
        rPr.append(caps_el)


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

