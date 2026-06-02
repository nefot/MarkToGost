# utils/html_table_parser.py
"""Парсер HTML-таблиц"""

from html.parser import HTMLParser
from typing import List, Optional
from MarkToGost.parser.blocks import HtmlTableBlock, HtmlTableRow, HtmlTableCell, CellAlign


class HtmlTableParser(HTMLParser):

    def __init__(self):
        super().__init__()
        self.rows: List[HtmlTableRow] = []
        self.transparent: bool = False
        self._current_row: Optional[List[HtmlTableCell]] = None
        self._current_cell: Optional[HtmlTableCell] = None
        self._in_bold = False

    def handle_starttag(self, tag, attrs):
        attrs_dict = dict(attrs)
        if tag == 'table':
            # <table class="transparent"> или <table transparent>
            classes = attrs_dict.get('class', '')
            self.transparent = (
                    'transparent' in classes.split()
                    or 'transparent' in attrs_dict
            )
        elif tag == 'tr':
            self._current_row = []
        elif tag == 'td':
            attrs_dict = dict(attrs)
            # Атрибуты без значения (bold, italic, underline) приходят как (name, name) или (name, None)
            attr_names = {a[0].lower() for a in attrs}

            colspan = int(attrs_dict.get('colspan', 1))
            rowspan = int(attrs_dict.get('rowspan', 1))
            align_str = attrs_dict.get('align', 'left').lower()
            align = {
                'left': CellAlign.LEFT,
                'center': CellAlign.CENTER,
                'right': CellAlign.RIGHT,
            }.get(align_str, CellAlign.LEFT)

            formula = attrs_dict.get('formula', None)

            self._current_cell = HtmlTableCell(
                text='',
                colspan=colspan,
                rowspan=rowspan,
                bold='bold' in attr_names,
                italic='italic' in attr_names,
                underline='underline' in attr_names,
                align=align,
                formula=formula,
            )
        elif tag == 'b':
            self._in_bold = True

    def handle_endtag(self, tag):
        if tag == 'tr':
            if self._current_row is not None:
                self.rows.append(HtmlTableRow(cells=self._current_row))
            self._current_row = None
        elif tag == 'td':
            if self._current_row is not None and self._current_cell is not None:
                self._current_cell.text = self._current_cell.text.strip()
                if self._in_bold:
                    self._current_cell.bold = True
                self._current_row.append(self._current_cell)
            self._current_cell = None
        elif tag == 'b':
            self._in_bold = False

    def handle_data(self, data):
        if self._current_cell is not None:
            self._current_cell.text += data


def parse_html_table(raw_html: str) -> HtmlTableBlock:
    parser = HtmlTableParser()
    parser.feed(raw_html)
    return HtmlTableBlock(rows=parser.rows, transparent=parser.transparent)
