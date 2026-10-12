"""Рендеринг таблиц"""
import re

from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from MarkToGost.formula_renderer import add_inline_formula
from MarkToGost.config import DocumentSettings
from MarkToGost.parser.blocks import TableBlock
from MarkToGost.utils.formatting import apply_italic_formatting, set_run_font, set_paragraph_formatting
from MarkToGost.utils.xml_helpers import set_table_borders, set_repeat_table_header
from MarkToGost.utils.document_helpers import split_md_table_row


MIN_TABLE_FONT_PT = 10


def _min_widths_cm(rows: list, size_pt: float) -> list:
    """Минимальная ширина колонки — самое длинное слово (в заголовке жирное, оно шире) плюс поля ячейки"""
    char_cm = size_pt * 0.021   # средняя ширина символа Times New Roman с запасом
    padding_cm = 0.45           # поля ячейки Word слева и справа
    widths = []
    for c in range(len(rows[0])):
        longest = 1
        for r_idx, row in enumerate(rows):
            factor = 1.15 if r_idx == 0 else 1.0
            for word in str(row[c]).split():
                longest = max(longest, len(word) * factor)
        widths.append(min(longest, 20) * char_cm + padding_cm)
    return widths


def _table_font_size(rows: list, usable_width_cm: float) -> float:
    """Размер шрифта таблицы: уменьшается до 10 pt, если длинные слова не помещаются по ширине"""
    size = DocumentSettings.TABLE_FONT_SIZE_PT
    while size > MIN_TABLE_FONT_PT and sum(_min_widths_cm(rows, size)) > usable_width_cm:
        size -= 1
    return size


def _column_widths_cm(rows: list, usable_width_cm: float, size_pt: float) -> list:
    """Ширины колонок по содержимому: не уже самого длинного слова, остаток — пропорционально объёму текста"""
    min_w = _min_widths_cm(rows, size_pt)
    weight = [max(min(len(str(r[c])), 60) for r in rows) or 1 for c in range(len(rows[0]))]

    total_min = sum(min_w)
    if total_min >= usable_width_cm:
        return [w * usable_width_cm / total_min for w in min_w]
    # Сначала каждой колонке минимум, затем свободное место по весу
    free = usable_width_cm - total_min
    total_weight = sum(weight)
    return [m + free * w / total_weight for m, w in zip(min_w, weight)]


def _keep_with_next(cell):
    for p in cell.paragraphs:
        p.paragraph_format.keep_with_next = True


def render_table_block(renderer, block: TableBlock):
    """Рендеринг таблицы.

    Таблица выводится целиком: на страницы её разбивает Word, строка заголовка повторяется на каждой
    странице. Подписи «Продолжение таблицы N» по фактическим границам страниц расставляет
    tools/finalize_docx.ps1.
    """
    if not block.rows:
        return

    parsed_rows = [split_md_table_row(row) for row in block.rows]
    max_cols = max(len(r) for r in parsed_rows) if parsed_rows else 0
    if max_cols <= 0:
        return

    for row in parsed_rows:
        if len(row) < max_cols:
            row.extend([""] * (max_cols - len(row)))

    sec = renderer.doc.sections[0]
    usable_width_cm = sec.page_width.cm - sec.left_margin.cm - sec.right_margin.cm
    font_pt = _table_font_size(parsed_rows, usable_width_cm)
    widths = _column_widths_cm(parsed_rows, usable_width_cm, font_pt)

    # --- Подпись ---
    cap = renderer.doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_paragraph_formatting(
        cap,
        first_line_indent=None,
        left_indent=Cm(0),
        space_before=6,
        space_after=0,
        line_spacing=DocumentSettings.LINE_SPACING
    )
    cap.paragraph_format.keep_with_next = True
    caption_full = (
        f"Таблица {renderer.table_counter} — {block.caption}"
        if block.caption
        else f"Таблица {renderer.table_counter}"
    )
    for part_text, is_italic in apply_italic_formatting(caption_full):
        run = cap.add_run(part_text)
        set_run_font(run, size_pt=DocumentSettings.FONT_SIZE_PT, bold=False, italic=is_italic)

    # --- Таблица ---
    table = renderer.doc.add_table(rows=len(parsed_rows), cols=max_cols)
    try:
        table.style = 'Table Grid'
    except Exception:
        pass
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False

    for col_idx, width_cm in enumerate(widths):
        table.columns[col_idx].width = Cm(width_cm)
        for row in table.rows:
            row.cells[col_idx].width = Cm(width_cm)

    # Повтор заголовка средствами Word на каждой странице
    set_repeat_table_header(table.rows[0])

    # --- Заполнение ячеек ---
    for r_idx, row_data in enumerate(parsed_rows):
        row = table.rows[r_idx]

        # Запрет разрыва строки посередине
        trPr = row._tr.get_or_add_trPr()
        cantSplit = OxmlElement('w:cantSplit')
        cantSplit.set(qn('w:val'), '1')
        trPr.append(cantSplit)

        for c_idx, value in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.text = ""
            for p in cell.paragraphs:
                p.clear()

            p = cell.paragraphs[0]
            p.style = None
            p.paragraph_format.first_line_indent = Pt(0)
            p.paragraph_format.left_indent = Pt(0)
            p.paragraph_format.right_indent = Pt(0)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.0
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if r_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT

            cell_text = value if value is not None else ""
            parts = re.split(r'\$(.+?)\$', cell_text)

            for i, part in enumerate(parts):
                if i % 2 == 0:
                    # Обычный текст
                    for part_text, is_italic in apply_italic_formatting(part):
                        run = p.add_run(part_text)
                        set_run_font(
                            run,
                            size_pt=font_pt,
                            bold=r_idx == 0,
                            italic=is_italic
                        )
                else:
                    # Inline-формула $...$
                    add_inline_formula(p, part)

            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            # Заголовок не остаётся внизу страницы без первой строки данных
            if r_idx == 0:
                _keep_with_next(cell)

    set_table_borders(table)

    renderer.table_counter += 1
    renderer._mark_content()

