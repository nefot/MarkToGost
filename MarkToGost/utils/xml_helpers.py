# utils/xml_helpers.py
"""Вспомогательные функции для работы с XML-элементами DOCX"""

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor

from MarkToGost.config import DocumentSettings


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

