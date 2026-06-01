from typing import List, Dict, Callable
import os

from docx import Document

from MarkToGost.parser.blocks import (
    BaseBlock, Section, TextBlock, HeadingBlock, ImageBlock, ListBlock,
    TableBlock, CodeBlock, FormulaBlock
)
# Импортируем функции рендеринга из отдельных модулей
from MarkToGost.renderer.text import render_text_block, render_heading_block
from MarkToGost.renderer.formula import render_formula_block
from MarkToGost.renderer.image import render_image_block
from MarkToGost.renderer.list_ import render_list_block
from MarkToGost.renderer.code import render_code_block
from MarkToGost.renderer.table import render_table_block
from MarkToGost.renderer.Section import render_section_block

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


class DocumentRenderer:
    """Рендерер блоков в DOCX — главный диспетчер"""

    def __init__(self, doc: Document, toc_entries: List[tuple] = None, use_headings: bool = True):
        self.doc = doc
        self.table_counter = 1
        self.formula_counter = 1  # Счетчик формул
        self.figure_counter = 0
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
        }

    def _is_document_start(self) -> bool:
        """Проверка начала документа"""
        return len(self.doc.paragraphs) == 0

    def _safe_page_break(self):
        """Не допускает двойных пустых страниц"""
        if not self._last_was_break:
            self.doc.add_page_break()
            self._last_was_break = True

    def _mark_content(self):
        """Сброс флага page break после добавления контента"""
        self._last_was_break = False

    def render_block(self, block: BaseBlock):
        """🔹 ОБЩАЯ ЛОГИКА для всех блоков (Template Method)
        
        Все блоки проходят через эту единую точку:
        1. Проверка типа блока
        2. Вызов функции рендеринга через регистр
        3. Единая сигнатура для всех (renderer, block)
        """
        block_type = type(block)
        if block_type in self._renderers:
            render_func = self._renderers[block_type]
            render_func(self, block)
        else:
            raise ValueError(f"Unknown block type: {block_type}")

