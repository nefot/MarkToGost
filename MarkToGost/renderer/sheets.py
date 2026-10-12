"""Вставные листы: титульный лист, задание, отзыв и любые другие DOCX-шаблоны с полями {{имя}}.

Лист вставляется отдельным разделом DOCX со своей вёрсткой (размер и поля страницы, вертикальное
выравнивание) и начинается с новой страницы. Стили, на которые ссылается лист, и рисунки
переносятся из шаблона в документ.
"""

import io
import os
import re
from copy import deepcopy
from typing import Dict, List, Set, Tuple

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from MarkToGost.parser.blocks import SheetBlock
from MarkToGost.utils.xml_helpers import end_section_here

PLACEHOLDER = re.compile(r"\{\{\s*([^{}]+?)\s*\}\}")
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
R_EMBED = qn("r:embed")


def field_value(value) -> str:
    """Значение поля: список — строки через перенос, None — пусто"""
    if value is None:
        return ""
    if isinstance(value, (list, tuple)):
        return "\n".join(field_value(v) for v in value)
    return str(value).rstrip("\n")


# ----------------------------------------------------------------------
# Заполнение полей
# ----------------------------------------------------------------------

def _lines(p_el) -> List[list]:
    """Строки абзаца: списки прогонов, строка заканчивается прогоном с переносом"""
    lines, current = [], []
    for run in p_el.iterchildren(qn("w:r")):
        current.append(run)
        if run.find(qn("w:br")) is not None:
            lines.append(current)
            current = []
    if current:
        lines.append(current)
    return lines


def _set_multiline(t_el, text: str) -> None:
    """Текст с переносами строк: t, br, t, br, t в одном прогоне"""
    parts = text.split("\n")
    t_el.text = parts[0]
    t_el.set(XML_SPACE, "preserve")
    anchor = t_el
    for part in parts[1:]:
        br = OxmlElement("w:br")
        t = OxmlElement("w:t")
        t.text = part
        t.set(XML_SPACE, "preserve")
        anchor.addnext(br)
        br.addnext(t)
        anchor = t


def _fill_paragraph(p_el, fields: Dict[str, str], unknown: Set[str]) -> None:
    for runs in _lines(p_el):
        text = "".join(t.text or "" for r in runs for t in r.iter(qn("w:t")))
        names = PLACEHOLDER.findall(text)
        if not names:
            continue
        unknown.update(n for n in names if n not in fields)
        if any(not fields.get(n, "").strip() for n in names):
            # Строка с незаполненным полем удаляется целиком вместе с переносом
            for run in runs:
                p_el.remove(run)
            continue
        new_text = PLACEHOLDER.sub(lambda m: fields[m.group(1)], text)
        t_elements = [t for r in runs for t in r.iter(qn("w:t"))]
        for t in t_elements[1:]:
            t.text = ""
        _set_multiline(t_elements[0], new_text)

    # Удалённая последняя строка оставляет висячий перенос в конце абзаца
    runs = list(p_el.iterchildren(qn("w:r")))
    if runs:
        last = runs[-1]
        if len(last) and last[-1].tag == qn("w:br"):
            last.remove(last[-1])


def fill_template(path: str, fields: Dict[str, object]) -> Tuple[Document, List[str]]:
    """Открывает шаблон и подставляет поля. Возвращает документ и поля шаблона, которых нет в fields"""
    values = {k: field_value(v) for k, v in fields.items()}
    doc = Document(path)
    unknown: Set[str] = set()
    for p_el in list(doc.element.body.iter(qn("w:p"))):
        _fill_paragraph(p_el, values, unknown)
    return doc, sorted(unknown)


def template_fields(path: str) -> List[str]:
    """Имена всех полей шаблона — для заготовки документа и проверок"""
    doc = Document(path)
    # Поле может быть разбито на прогоны — собираем текст абзацев целиком
    text = "\n".join("".join(t.text or "" for t in p.iter(qn("w:t"))) for p in doc.element.body.iter(qn("w:p")))
    return list(dict.fromkeys(PLACEHOLDER.findall(text)))


# ----------------------------------------------------------------------
# Перенос стилей и рисунков
# ----------------------------------------------------------------------

def _copy_styles(src: Document, dst: Document, elements) -> None:
    """Переносит стили, на которые ссылается лист и которых нет в документе (с цепочкой basedOn)"""
    dst_styles = dst.styles.element
    src_styles = src.styles.element
    existing = {s.get(qn("w:styleId")) for s in dst_styles.iterchildren(qn("w:style"))}
    by_id = {s.get(qn("w:styleId")): s for s in src_styles.iterchildren(qn("w:style"))}

    wanted = []
    for el in elements:
        for tag in ("w:pStyle", "w:rStyle", "w:tblStyle"):
            for ref in el.iter(qn(tag)):
                wanted.append(ref.get(qn("w:val")))
    while wanted:
        style_id = wanted.pop()
        if style_id in existing or style_id not in by_id:
            continue
        style = deepcopy(by_id[style_id])
        dst_styles.append(style)
        existing.add(style_id)
        for link in ("w:basedOn", "w:next", "w:link"):
            ref = style.find(qn(link))
            if ref is not None:
                wanted.append(ref.get(qn("w:val")))


def _copy_images(src: Document, dst: Document, element) -> None:
    for blip in element.iter(qn("a:blip")):
        rid = blip.get(R_EMBED)
        part = src.part.related_parts.get(rid) if rid else None
        if part is None:
            continue
        new_rid, _ = dst.part.get_or_add_image(io.BytesIO(part.blob))
        blip.set(R_EMBED, new_rid)


def _strip_unportable(element) -> None:
    """Нумерация списков и ссылки на части шаблона в документе недействительны — убираем"""
    for num_pr in list(element.iter(qn("w:numPr"))):
        num_pr.getparent().remove(num_pr)
    for link in list(element.iter(qn("w:hyperlink"))):
        link.attrib.pop(qn("r:id"), None)


# ----------------------------------------------------------------------
# Вставка
# ----------------------------------------------------------------------

def _sheet_sectPr(src: Document, page_numbers: bool):
    sect = deepcopy(src.element.body.sectPr)
    # Колонтитулы шаблона в документ не переносятся: номера страниц — общие
    for tag in ("w:headerReference", "w:footerReference"):
        for el in sect.findall(qn(tag)):
            sect.remove(el)
    for el in sect.findall(qn("w:type")):
        sect.remove(el)
    sect_type = OxmlElement("w:type")
    sect_type.set(qn("w:val"), "nextPage")
    sect.insert(0, sect_type)
    title_pg = sect.find(qn("w:titlePg"))
    if not page_numbers and title_pg is None:
        sect.append(OxmlElement("w:titlePg"))
    elif page_numbers and title_pg is not None:
        sect.remove(title_pg)
    return sect


def insert_sheet(renderer, src: Document, page_numbers: bool = True) -> None:
    """Вставляет содержимое шаблона отдельным разделом с новой страницы"""
    doc = renderer.doc
    body = doc.element.body

    # Разрыв, отложенный предыдущим разделом, не нужен: лист и так начнётся с новой страницы
    renderer._pending_break = False
    prev = body.sectPr.getprevious()
    section_just_ended = (prev is not None and prev.tag == qn("w:p")
                          and prev.find(qn("w:pPr") + "/" + qn("w:sectPr")) is not None)
    # Перед листом есть текст основной части — завершаем её раздел; лист сразу после листа — нет
    if prev is not None and not section_just_ended:
        end_section_here(body, deepcopy(body.sectPr))

    elements = [deepcopy(el) for el in src.element.body if el.tag != qn("w:sectPr")]
    _copy_styles(src, doc, elements)
    for el in elements:
        _copy_images(src, doc, el)
        _strip_unportable(el)
        body.sectPr.addprevious(el)

    end_section_here(body, _sheet_sectPr(src, page_numbers))

    # Основная часть после листа — с новой страницы, без особенностей вёрстки листа
    main = body.sectPr
    for tag in ("w:vAlign", "w:titlePg", "w:type"):
        for el in main.findall(qn(tag)):
            main.remove(el)
    sect_type = OxmlElement("w:type")
    sect_type.set(qn("w:val"), "nextPage")
    main.insert(0, sect_type)

    renderer._last_was_break = True


def resolve_template(path: str, base_dir, titles_dir: str) -> str:
    """Шаблон: абсолютный путь, относительно документа или файл из profiles/titles/"""
    if os.path.isabs(path):
        return path
    for folder in (base_dir, titles_dir):
        if folder and os.path.isfile(os.path.join(folder, path)):
            return os.path.join(folder, path)
    return os.path.join(base_dir or titles_dir, path)


def render_sheet_block(renderer, block: SheetBlock) -> None:
    fields = getattr(renderer, "fields", {}) or {}
    titles_dir = getattr(renderer, "titles_dir", "")
    path = resolve_template(block.template, getattr(renderer, "base_dir", None), titles_dir)
    if not os.path.isfile(path):
        renderer.warnings.append(f"не найден шаблон листа: {block.template}")
        return
    src, unknown = fill_template(path, fields)
    for name in unknown:
        renderer.warnings.append(f"поле «{name}» в листе {os.path.basename(path)} не задано")
    insert_sheet(renderer, src, page_numbers=block.page_numbers)
