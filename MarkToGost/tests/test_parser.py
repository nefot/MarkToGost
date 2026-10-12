# tests/test_parser.py

from MarkToGost.parser.blocks import (
    CodeBlock, HeadingBlock, ImageBlock, ListBlock, Section, TableBlock, TextBlock
)
from MarkToGost.parser.markdown_parser import MarkdownParser
from MarkToGost.utils.formatting import apply_italic_formatting


def parse(md: str):
    return MarkdownParser(md).parse()


# ================================
# Абзацы, заголовки, списки
# ================================

class TestBasicBlocks:

    def test_paragraph_lines_are_joined(self):
        blocks = parse("первая строка\nвторая строка")
        assert blocks == [TextBlock(text="первая строка вторая строка")]

    def test_blank_line_separates_paragraphs(self):
        blocks = parse("абзац 1\n\nабзац 2")
        assert [b.text for b in blocks] == ["абзац 1", "абзац 2"]

    def test_heading_level_from_number(self):
        blocks = parse("### 2.3 Связи между сущностями")
        assert blocks == [HeadingBlock(text="2.3 Связи между сущностями", level=2)]

    def test_appendix_subsection_level(self):
        assert parse("## А.1 Форма заказа")[0].level == 2
        assert parse("### B.2.1 Пункт")[0].level == 3
        assert parse("# ПРИЛОЖЕНИЕ А. ИСХОДНЫЕ ДОКУМЕНТЫ")[0].level == 1

    def test_unordered_list(self):
        blocks = parse("- один\n- два")
        assert blocks == [ListBlock(items=["один", "два"], ordered=False)]

    def test_ordered_list(self):
        blocks = parse("1. один\n2. два")
        assert blocks == [ListBlock(items=["один", "два"], ordered=True)]


# ================================
# Таблицы и код
# ================================

class TestTablesAndCode:

    def test_table_with_caption(self):
        md = "Таблица — Материалы\n| Код | Название |\n|---|---|\n| 1 | PLA |"
        blocks = parse(md)
        assert len(blocks) == 1
        assert isinstance(blocks[0], TableBlock)
        assert blocks[0].caption == "Материалы"
        assert blocks[0].rows == ["| Код | Название |", "| 1 | PLA |"]

    def test_code_block_keeps_lines(self):
        blocks = parse("```\nSELECT 1;\nSELECT 2;\n```")
        assert blocks == [CodeBlock(code="SELECT 1;\nSELECT 2;", language="")]

    def test_code_block_language_from_fence(self):
        block = parse("```sql\nSELECT 1;\n```")[0]
        assert block.language == "sql"

    def test_colon_in_first_line_does_not_drop_code(self):
        block = parse("```\n-- Запрос: заказы\nSELECT *\nFROM print_order;\n```")[0]
        assert block.code == "-- Запрос: заказы\nSELECT *\nFROM print_order;"

    def test_code_indentation_and_blank_lines_preserved(self):
        block = parse("```\ndef f():\n\n        return 1\n```")[0]
        assert block.code == "def f():\n\n        return 1"


# ================================
# Рисунки
# ================================

class TestImages:

    def test_plain_image(self):
        blocks = parse("![Схема](img/a.png)")
        assert blocks == [ImageBlock(path="img/a.png", caption="Схема")]

    def test_image_id(self):
        block = parse("![Схема](img/a.png){#scheme}")[0]
        assert block.img_id == "scheme"
        assert block.landscape is False

    def test_landscape_image(self):
        block = parse("![Схема](img/a.drawio#2){landscape}")[0]
        assert block.path == "img/a.drawio#2"
        assert block.landscape is True

    def test_id_and_landscape_together(self):
        block = parse("![Схема](img/a.png){#scheme landscape}")[0]
        assert block.img_id == "scheme"
        assert block.landscape is True

    def test_image_ending_with_brace_does_not_close_section(self):
        md = "[//]: ### (1. РАЗДЕЛ){\n![Схема](img/a.png){landscape}\nтекст после рисунка\n}"
        blocks = parse(md)
        assert len(blocks) == 1
        section = blocks[0]
        assert isinstance(section, Section)
        assert [type(b) for b in section.blocks] == [ImageBlock, TextBlock]


# ================================
# Разделы
# ================================

class TestSections:

    def test_section_with_content(self):
        blocks = parse("[//]: ### (ВВЕДЕНИЕ){\nТекст введения\n}")
        assert len(blocks) == 1
        section = blocks[0]
        assert section.section_id == "ВВЕДЕНИЕ"
        assert section.heading_level == 3
        assert section.add_page_breaks is True
        assert section.blocks == [TextBlock(text="Текст введения")]

    def test_html_table_inside_section(self):
        md = "[//]: ### (1. РАЗДЕЛ){\n<table>\n<tr><td>A</td></tr>\n</table>\n}"
        section = parse(md)[0]
        assert [type(b).__name__ for b in section.blocks] == ["HtmlTableBlock"]

    def test_section_without_page_break(self):
        section = parse("[//]: ### (1. РАЗДЕЛ)[new_page=false]")[0]
        assert section.add_page_breaks is False

    def test_section_without_page_break_with_content(self):
        blocks = parse("[//]: ### (1.2 Подраздел)[new_page=false]{\nтекст\n}")
        assert len(blocks) == 1
        assert blocks[0].add_page_breaks is False
        assert blocks[0].blocks == [TextBlock(text="текст")]


# ================================
# Экранирование подчёркивания
# ================================

class TestEscapedUnderscore:

    def test_escaped_underscore_is_literal(self):
        assert apply_italic_formatting(r"таблица print\_order") == [("таблица print_order", False)]

    def test_escaped_underscore_inside_italic_text(self):
        assert apply_italic_formatting(r"_a\_b_ и c") == [("a_b", True), (" и c", False)]
