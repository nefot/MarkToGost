# -*- coding: utf-8 -*-
"""
ПАТЧ для md_to_gost_docx_batch.py
===================================

Шаг 1: Добавь импорты в начало файла (после остальных импортов):

    from formula_renderer import add_formula_paragraph, extract_omml

Шаг 2: Удали функции render_formula_with_pandoc и append_docx_content —
        они больше не нужны.

Шаг 3: Замени метод _render_formula_block в классе DocumentRenderer
        на версию ниже.

Шаг 4: Удали метод _render_formula (старый, с комментарием "для совместимости").

Больше ничего менять не нужно.
"""
from docx.enum.text import WD_ALIGN_PARAGRAPH

from md_to_gost_block import convert_inline_math, set_paragraph_formatting, DocumentSettings
from formula_renderer import add_formula_paragraph, extract_omml


# ============================================================
# ЗАМЕНА ДЛЯ DocumentRenderer._render_formula_block
# ============================================================


