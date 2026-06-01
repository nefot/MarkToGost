"""Рендеринг блоков кода"""

import re
from docx.shared import Pt, Cm

from MarkToGost.config import DocumentSettings
from MarkToGost.parser.blocks import CodeBlock


def render_code_block(renderer, block: CodeBlock):
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
    p = renderer.doc.add_paragraph()
    p.add_run(code).font.name = 'Courier New'
    p.paragraph_format.left_indent = Cm(0.5)
    p.paragraph_format.right_indent = Cm(0.5)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)

    renderer._mark_content()

