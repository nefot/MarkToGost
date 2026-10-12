"""Рендеринг блоков изображений"""

import os
import re
from copy import deepcopy

from docx.enum.section import WD_ORIENT, WD_SECTION_START
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.image.image import Image as DocxImage
from docx.oxml.ns import qn
from docx.shared import Cm

from MarkToGost.config import DocumentSettings
from MarkToGost.parser.blocks import ImageBlock
from MarkToGost.utils.drawio import is_drawio_path, split_drawio_path, export_drawio
from MarkToGost.utils.formatting import apply_italic_formatting, set_run_font, set_paragraph_formatting

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROJECT_ROOT = os.path.dirname(BASE_DIR)
INPUT_DIR = os.path.join(PROJECT_ROOT, "input")


CAPTION_RESERVE_CM = 1.5  # место под подпись рисунка на странице


def _start_section(renderer, landscape: bool):
    """Завершает текущий раздел DOCX и начинает новый с нужной ориентацией страницы"""
    body = renderer.doc.element.body
    body_sectPr = body.sectPr
    ended = deepcopy(body_sectPr)  # свойства завершаемого раздела
    last = body_sectPr.getprevious()
    if last is None or last.tag != qn('w:p') or last.find(qn('w:pPr') + '/' + qn('w:sectPr')) is not None:
        last = renderer.doc.add_paragraph()._p
    last.get_or_add_pPr().append(ended)

    sec = renderer.doc.sections[-1]
    sec.start_type = WD_SECTION_START.NEW_PAGE
    width, height = Cm(DocumentSettings.PAGE_WIDTH_CM), Cm(DocumentSettings.PAGE_HEIGHT_CM)
    sec.orientation = WD_ORIENT.LANDSCAPE if landscape else WD_ORIENT.PORTRAIT
    sec.page_width, sec.page_height = (height, width) if landscape else (width, height)
    renderer._landscape_open = landscape
    renderer._last_was_break = True


def close_landscape(renderer):
    """Возврат к книжной ориентации после альбомных рисунков"""
    if getattr(renderer, "_landscape_open", False):
        _start_section(renderer, landscape=False)


def _picture_width(img_path: str, landscape: bool) -> Cm:
    """Ширина рисунка: доля ширины страницы, но так, чтобы рисунок с подписью поместился по высоте"""
    s = DocumentSettings
    page_w, page_h = (s.PAGE_HEIGHT_CM, s.PAGE_WIDTH_CM) if landscape else (s.PAGE_WIDTH_CM, s.PAGE_HEIGHT_CM)
    avail_w = page_w - s.LEFT_MARGIN_CM - s.RIGHT_MARGIN_CM
    avail_h = page_h - s.TOP_MARGIN_CM - s.BOTTOM_MARGIN_CM - CAPTION_RESERVE_CM
    width = avail_w if landscape else avail_w * s.IMAGE_WIDTH_FRACTION
    try:
        img = DocxImage.from_file(img_path)
        width = min(width, avail_h * img.px_width / img.px_height)
    except Exception:
        pass
    return Cm(width)


def render_image_block(renderer, block: ImageBlock):
    """Рендеринг изображения"""
    # Схема draw.io: путь вида img/a.drawio#Страница
    is_drawio = is_drawio_path(block.path)
    src_path, drawio_page = split_drawio_path(block.path) if is_drawio else (block.path, "")
    img_path = src_path

    # Когда указан относительный путь — пробуем несколько мест в проекте (пакет, input/, корень проекта)
    if not os.path.isabs(img_path):
        candidates = [
            os.path.join(BASE_DIR, src_path),      # относительно пакета MarkToGost/
            os.path.join(INPUT_DIR, src_path),     # относительно input/
            os.path.join(PROJECT_ROOT, src_path),  # относительно корня проекта
            src_path                               # как есть (относительно cwd)
        ]
        found = None
        for p in candidates:
            if os.path.exists(p):
                found = p
                break
        if found:
            img_path = found
        else:
            # Оставляем первый вариант для сообщения об ошибке
            img_path = candidates[0]

    if is_drawio and os.path.exists(img_path):
        try:
            img_path = export_drawio(img_path, drawio_page)
        except Exception as e:
            p = renderer.doc.add_paragraph(f"[Ошибка экспорта draw.io: {os.path.basename(src_path)} — {e}]")
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            renderer._mark_content()
            return

    if os.path.exists(img_path):
        # Широкая схема — на отдельной альбомной странице; подряд идущие делят один альбомный раздел
        continue_landscape = block.landscape and getattr(renderer, "_landscape_open", False)
        if block.landscape and not continue_landscape:
            _start_section(renderer, landscape=True)

        # Изображение
        pic_para = renderer.doc.add_paragraph()
        pic_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pic_para.paragraph_format.first_line_indent = Cm(0)
        pic_para.paragraph_format.keep_with_next = True
        if continue_landscape:
            pic_para.paragraph_format.page_break_before = True

        img_width_cm = _picture_width(img_path, block.landscape)
        try:
            pic_para.add_run().add_picture(img_path, width=img_width_cm)
        except Exception:
            pic_para.add_run(f"[Ошибка вставки: {os.path.basename(img_path)}]")

        # Получаем номер изображения
        # Если img_id указан и был собран на первом проходе, используем сохраненный номер
        if block.img_id and block.img_id in renderer.image_refs:
            fig_num = renderer.image_refs[block.img_id]
        else:
            # Иначе увеличиваем счетчик
            renderer.figure_counter += 1
            fig_num = renderer.figure_counter
            # Сохраняем в image_refs, если есть ID
            if block.img_id:
                renderer.image_refs[block.img_id] = fig_num

        fname = os.path.basename(block.path)
        caption_text = block.caption or ""

        if caption_text:
            # Заглавной делаем только первую букву, регистр остального сохраняем (IDEF0, RuDelia)
            caption_text = re.sub(r'[^\W\d_]', lambda m: m.group().upper(), caption_text.strip(), count=1)
            caption_full = f"Рисунок {fig_num} — {caption_text}"  # тире по ГОСТ 7.32
        else:
            caption_full = f"Рисунок {fig_num}"

        cap_para = renderer.doc.add_paragraph(caption_full)
        cap_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_paragraph_formatting(cap_para, line_spacing=DocumentSettings.LINE_SPACING)

        # Применяем курсив
        cap_para.clear()
        for part_text, is_italic in apply_italic_formatting(caption_full):
            run = cap_para.add_run(part_text)
            set_run_font(run, size_pt=DocumentSettings.CAPTION_FONT_SIZE_PT,
                        italic=DocumentSettings.CAPTION_ITALIC or is_italic)

        renderer.image_map[fname] = str(fig_num)
        renderer._mark_content()
    else:
        p = renderer.doc.add_paragraph(f"[Изображение не найдено: {os.path.basename(img_path)}]")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        renderer._mark_content()

