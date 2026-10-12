from typing import List, Dict, Callable
import os

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from MarkToGost.parser.blocks import (
    BaseBlock, Section, TextBlock, HeadingBlock, ImageBlock, ListBlock,
    TableBlock, CodeBlock, FormulaBlock, SheetBlock
)
# Импортируем функции рендеринга из отдельных модулей
from MarkToGost.renderer.text import render_text_block, render_heading_block
from MarkToGost.renderer.formula import render_formula_block
from MarkToGost.renderer.image import render_image_block, close_landscape
from MarkToGost.renderer.sheets import render_sheet_block
from MarkToGost.config import DocumentSettings
from MarkToGost.utils.numbering import Numbering
from MarkToGost.renderer.list_ import render_list_block
from MarkToGost.renderer.code import render_code_block
from MarkToGost.renderer.table import render_table_block
from MarkToGost.renderer.section import render_section_block

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

from MarkToGost.parser.blocks import HtmlTableBlock
from MarkToGost.renderer.html_table import render_html_table_block

class DocumentRenderer:
    """Рендерер блоков в DOCX — главный диспетчер"""

    def __init__(self, doc: Document, toc_entries: List[tuple] = None, use_headings: bool = True):
        self.doc = doc
        # Номера рисунков, таблиц и формул (сквозные или по разделам — DocumentSettings.OBJECT_NUMBERING)
        self.numbering = Numbering(DocumentSettings.OBJECT_NUMBERING)
        self.image_map = {}
        self.image_refs = {}  # Для хранения ссылок на изображения по их ID
        self.formula_refs = {}  # Для хранения ссылок на формулы по их ID
        self.use_headings = use_headings
        self.toc_entries = toc_entries or []
        self._last_was_break = False  # Флаг для защиты от двойных разрывов
        
        # Регистр функций рендеринга - единый интерфейс (renderer, block) -> None
        self._renderers: Dict[type, Callable] = {
            Section: render_section_block,
            TextBlock: render_text_block,
            HeadingBlock: render_heading_block,
            ImageBlock: render_image_block,
            ListBlock: render_list_block,
            TableBlock: render_table_block,
            CodeBlock: render_code_block,
            FormulaBlock: render_formula_block,
            HtmlTableBlock: render_html_table_block,
            SheetBlock: render_sheet_block,
        }
        # Для вставных листов: значения полей, папка документа и стандартных шаблонов, замечания
        self.fields = {}
        self.base_dir = None
        self.titles_dir = ""
        self.warnings = []

    def _is_document_start(self) -> bool:
        """Проверка начала документа"""
        return len(self.doc.paragraphs) == 0

    def _safe_page_break(self):
        """Не допускает двойных пустых страниц.

        Разрыв не вставляется сразу: следующий за ним абзац получает свойство «с новой страницы».
        Отдельный абзац с разрывом давал пустую страницу, если предыдущая была заполнена до конца,
        а разрыв в конце абзаца растягивал по ширине его последнюю строку.
        """
        if self._last_was_break:
            return
        self._apply_pending_break()  # предыдущий разрыв не должен потеряться
        body = self.doc.element.body
        self._break_marker = body.sectPr.getprevious() if body.sectPr is not None else None
        self._pending_break = True
        self._last_was_break = True

    def _apply_pending_break(self):
        if not getattr(self, "_pending_break", False):
            return
        body = self.doc.element.body
        nxt = self._break_marker.getnext() if self._break_marker is not None else body[0]
        if nxt is None or nxt is body.sectPr:
            return  # после разрыва пока ничего нет
        self._pending_break = False
        if nxt.tag == qn('w:p'):
            pPr = nxt.get_or_add_pPr()
            el = OxmlElement('w:pageBreakBefore')
            pPr.insert(0, el)
        else:
            # Перед таблицей — отдельный абзац с разрывом
            p = OxmlElement('w:p')
            r = OxmlElement('w:r')
            br = OxmlElement('w:br')
            br.set(qn('w:type'), 'page')
            r.append(br)
            p.append(r)
            nxt.addprevious(p)

    def _mark_content(self):
        """Вызывается после добавления содержимого: применяет отложенный разрыв, сбрасывает флаг"""
        self._apply_pending_break()
        self._last_was_break = False

    def render_block(self, block: BaseBlock):
        """🔹 ОБЩАЯ ЛОГИКА для всех блоков (Template Method)
        
        Все блоки проходят через эту единую точку:
        1. Проверка типа блока
        2. Вызов функции рендеринга через регистр
        3. Единая сигнатура для всех (renderer, block)
        """
        block_type = type(block)
        self._apply_pending_break()
        if not (isinstance(block, ImageBlock) and block.landscape):
            close_landscape(self)
        if block_type in self._renderers:
            render_func = self._renderers[block_type]
            render_func(self, block)
            self._apply_pending_break()
        else:
            raise ValueError(f"Unknown block type: {block_type}")

