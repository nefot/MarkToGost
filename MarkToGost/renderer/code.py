"""Рендеринг блоков кода"""

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm

from MarkToGost.config import DocumentSettings
from MarkToGost.parser.blocks import CodeBlock


def render_code_block(renderer, block: CodeBlock):
    """Рендеринг блока кода: моноширинный шрифт, текст и отступы сохраняются без изменений"""
    p = renderer.doc.add_paragraph()
    run = p.add_run(block.code)
    run.font.name = 'Courier New'
    if DocumentSettings.CODE_FONT_SIZE_PT:
        run.font.size = Pt(DocumentSettings.CODE_FONT_SIZE_PT)
    if DocumentSettings.CODE_LINE_SPACING:
        p.paragraph_format.line_spacing = DocumentSettings.CODE_LINE_SPACING
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.left_indent = Cm(0.5)
    p.paragraph_format.right_indent = Cm(0.5)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)

    renderer._mark_content()
