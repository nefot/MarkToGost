# -*- coding: utf-8 -*-
"""
Модуль рендеринга формул LaTeX -> OMML (нативный Word)

Принцип работы:
  1. Pandoc конвертирует LaTeX -> DOCX с нативным OMML
  2. Мы извлекаем XML-узел <m:oMathPara> из этого DOCX
  3. Вставляем его напрямую в параграф целевого документа

Результат: формулы редактируемы в редакторе формул Word,
           корректно отображаются без шрифтов/картинок.
"""

import os
import subprocess
import tempfile
import zipfile
from copy import deepcopy
from typing import Optional, List

from lxml import etree
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# XML namespace для OMML
_OMML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


# ---------------------------------------------------------------------------
# Публичный API
# ---------------------------------------------------------------------------

def extract_omml(latex: str) -> Optional[List]:
    """
    Конвертирует LaTeX-формулу в список OMML XML-элементов через pandoc.

    Возвращает список элементов для вставки в p._p,
    либо None если pandoc недоступен или произошла ошибка.

    Поддерживает любой валидный LaTeX: дроби, суммы, интегралы,
    матрицы, греческие символы и т.д.
    """
    try:
        return _pandoc_latex_to_omml(latex)
    except Exception:
        return None


def add_formula_paragraph(doc: Document, latex: str,
                           number: Optional[str] = None,
                           font_size_pt: int = 14,
                           indent_cm: float = 3.0) -> bool:
    """
    Добавляет формулу в документ в формате ГОСТ:
      [отступ]  <формула>  [(номер)]

    Если pandoc недоступен — добавляет текстовый fallback $latex$.

    Возвращает True если формула вставлена как OMML, False если fallback.
    """
    omml_nodes = extract_omml(latex)

    if omml_nodes:
        _insert_omml_table(doc, omml_nodes, number, font_size_pt)
        return True
    else:
        _insert_formula_fallback(doc, latex, number, font_size_pt)
        return False


def add_inline_formula(paragraph, latex: str) -> bool:
    """
    Вставляет inline-формулу в существующий параграф.

    Возвращает True если успешно, False если fallback (текст).
    """
    omml_nodes = extract_omml(latex)
    if omml_nodes:
        for node in omml_nodes:
            paragraph._p.append(deepcopy(node))
        return True
    else:
        run = paragraph.add_run(f"${latex}$")
        run.font.size = Pt(12)
        return False


# ---------------------------------------------------------------------------
# Внутренние функции
# ---------------------------------------------------------------------------

def _pandoc_latex_to_omml(latex: str) -> Optional[List]:
    """Запускает pandoc и извлекает OMML из результирующего DOCX."""
    with tempfile.TemporaryDirectory() as tmpdir:
        md_path = os.path.join(tmpdir, "formula.md")
        docx_path = os.path.join(tmpdir, "formula.docx")

        # Оборачиваем в display math ($$...$$)
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(f"$$\n{latex}\n$$\n")

        result = subprocess.run(
            ["pandoc", md_path, "-o", docx_path],
            capture_output=True,
            timeout=15
        )

        if result.returncode != 0 or not os.path.exists(docx_path):
            return None

        with zipfile.ZipFile(docx_path) as z:
            xml_bytes = z.read("word/document.xml")

    root = etree.fromstring(xml_bytes)

    # Ищем oMathPara — блочная формула (предпочтительно)
    nodes = root.findall(f".//{{{_OMML_NS}}}oMathPara")
    if nodes:
        return [deepcopy(nodes[0])]

    # Fallback: просто oMath (inline внутри pandoc)
    nodes = root.findall(f".//{{{_OMML_NS}}}oMath")
    if nodes:
        return [deepcopy(n) for n in nodes]

    return None


def _insert_omml_table(doc: Document, omml_nodes: List,
                        number: Optional[str],
                        font_size_pt: int):
    """
    Вставляет формулу в таблицу без границ:
      | <OMML по центру> | (N) по правому краю |
    """
    sec = doc.sections[0]
    usable_w = sec.page_width.cm - sec.left_margin.cm - sec.right_margin.cm

    table = doc.add_table(rows=1, cols=2)
    table.autofit = False

    # Ширины: 85% формула, 15% номер
    w_formula = Cm(usable_w * 0.85)
    w_number = Cm(usable_w * 0.15)
    table.columns[0].width = w_formula
    table.columns[1].width = w_number

    # --- Левая ячейка: формула ---
    left = table.cell(0, 0)
    left_p = left.paragraphs[0]
    left_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    left_p.paragraph_format.space_before = Pt(0)
    left_p.paragraph_format.space_after = Pt(0)

    for node in omml_nodes:
        left_p._p.append(deepcopy(node))

    # --- Правая ячейка: номер ---
    right = table.cell(0, 1)
    right_p = right.paragraphs[0]
    right_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    right_p.paragraph_format.space_before = Pt(0)
    right_p.paragraph_format.space_after = Pt(0)

    if number:
        run = right_p.add_run(f"({number})")
        run.font.name = "Times New Roman"
        run.font.size = Pt(font_size_pt)

    # --- Убираем все границы ---
    _remove_table_borders(table)


def _remove_table_borders(table):
    """Убирает все границы таблицы через XML."""
    tbl = table._tbl
    tblPr = tbl.tblPr

    # Удаляем старые границы
    for child in list(tblPr):
        if child.tag.split("}")[-1] == "tblBorders":
            tblPr.remove(child)

    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "nil")
        borders.append(el)
    tblPr.append(borders)


def _insert_formula_fallback(doc: Document, latex: str,
                              number: Optional[str],
                              font_size_pt: int):
    """Текстовый fallback если pandoc недоступен."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    text = f"${latex}$"
    if number:
        text += f"    ({number})"
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(font_size_pt)
