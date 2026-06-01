# utils/__init__.py
"""Вспомогательные функции для работы с документами и форматированием"""

from MarkToGost.utils.formatting import (
    apply_italic_formatting,
    set_run_font,
    set_paragraph_formatting,
)

from MarkToGost.utils.xml_helpers import (
    set_table_borders,
    set_repeat_table_header,
    add_page_number_centered,
)

from MarkToGost.utils.toc import (
    add_toc,
    reset_heading_styles,
    get_heading_level_from_number,
    get_toc_level,
)

from MarkToGost.utils.document_helpers import (
    compute_image_width_cm,
    replace_image_refs,
    split_md_table_row,
    is_md_table_separator,
    is_md_table_row,
    normalize_table_caption,
    append_docx_content,
    render_formula_with_pandoc,
)

__all__ = [
    # formatting
    "apply_italic_formatting",
    "set_run_font",
    "set_paragraph_formatting",
    # xml_helpers
    "set_table_borders",
    "set_repeat_table_header",
    "add_page_number_centered",
    # toc
    "add_toc",
    "reset_heading_styles",
    "get_heading_level_from_number",
    "get_toc_level",
    # document_helpers
    "compute_image_width_cm",
    "replace_image_refs",
    "split_md_table_row",
    "is_md_table_separator",
    "is_md_table_row",
    "normalize_table_caption",
    "append_docx_content",
    "render_formula_with_pandoc",
]

