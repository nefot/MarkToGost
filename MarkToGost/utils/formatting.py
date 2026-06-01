# utils/formatting.py

import re
from typing import List, Tuple

from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

from MarkToGost.config import DocumentSettings


def apply_italic_formatting(text: str) -> List[Tuple[str, bool]]:
    """
    Разбивает текст на части по маркерам _курсив_.
    Возвращает список кортежей (текст, is_italic).

    Примеры:
        "обычный"         -> [("обычный", False)]
        "_курсив_"        -> [("курсив", True)]
        "до _курсив_ после" -> [("до ", False), ("курсив", True), (" после", False)]
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


def set_run_font(run, size_pt: int = 14, bold: bool = False, italic: bool = False):
    """
    Устанавливает шрифт, размер, начертание и цвет для run.
    Всегда Times New Roman, чёрный цвет.
    """
    run.font.name = DocumentSettings.FONT_NAME
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor(0, 0, 0)


def set_paragraph_formatting(paragraph, *,
                              align=None,
                              first_line_indent=None,
                              left_indent=None,
                              line_spacing: float = 1.25,
                              space_before: float = 0,
                              space_after: float = 0):
    """
    Устанавливает форматирование абзаца.

    Параметры:
        align               — WD_ALIGN_PARAGRAPH.*
        first_line_indent   — Cm(...) или None
        left_indent         — Cm(...) или None
        line_spacing        — множитель межстрочного интервала
        space_before        — отступ до абзаца в pt
        space_after         — отступ после абзаца в pt
    """
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