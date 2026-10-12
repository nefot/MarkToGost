# -*- coding: utf-8 -*-
"""
MarkToGost: конвертер Markdown → DOCX по ГОСТ 7.32

Использование:
    python -m MarkToGost.main                          # все .md из input/
    python -m MarkToGost.main отчёт.md                 # один файл (путь или имя в input/)
    python -m MarkToGost.main --new курсовая моя.md    # заготовка документа по типу
    python -m MarkToGost.main --types                  # список типов документов
    python -m MarkToGost.main --help
"""

import argparse
import dataclasses
import os
import re
import subprocess
import sys
from contextlib import contextmanager
from typing import List, Optional

import yaml
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt

from MarkToGost.config import DocumentSettings
from MarkToGost.parser.blocks import HeadingBlock, Section
from MarkToGost.parser.frontmatter import split_front_matter
from MarkToGost.parser.markdown_parser import MarkdownParser
from MarkToGost.parser.metadata import extract_metadata
from MarkToGost.profiles import (
    FORMAT_KEYS, TITLE_SHEET, TITLES_DIR, Profile, ProfileError, SheetSpec, available_profiles, load_profile,
    merge, parse_sheet, split_settings, unknown_format_keys
)
from MarkToGost.renderer.document_renderer import DocumentRenderer
from MarkToGost.renderer.sheets import template_fields
from MarkToGost.structure import check_fields, check_structure, group_sections, insert_toc, place_sheets
from MarkToGost.utils.formatting import set_run_font
from MarkToGost.utils.toc import get_toc_level, reset_heading_styles
from MarkToGost.utils.xml_helpers import add_page_number_centered, set_repeat_table_header, set_table_borders

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INPUT_DIR = os.path.join(PROJECT_ROOT, "input")
OUTPUT_DIR = os.path.join(PROJECT_ROOT, "output")
FINALIZE_SCRIPT = os.path.join(PROJECT_ROOT, "tools", "finalize_docx.ps1")
# Общие данные для всех документов (вуз, кафедра, студент…): ищется рядом с документом и в корне проекта
DEFAULTS_FILE = "marktogost.yaml"

# Параметр командной строки → атрибут DocumentSettings
CLI_TO_SETTINGS = {
    'font_name': 'FONT_NAME',
    'font_size': 'FONT_SIZE_PT',
    'line_spacing': 'LINE_SPACING',
    'first_line_indent': 'FIRST_LINE_INDENT_CM',
    'left_margin': 'LEFT_MARGIN_CM',
    'right_margin': 'RIGHT_MARGIN_CM',
    'top_margin': 'TOP_MARGIN_CM',
    'bottom_margin': 'BOTTOM_MARGIN_CM',
    'image_width': 'IMAGE_WIDTH_FRACTION',
    'caption_font_size': 'CAPTION_FONT_SIZE_PT',
    'caption_italic': 'CAPTION_ITALIC',
    'table_font_size': 'TABLE_FONT_SIZE_PT',
    'use_first_line_indent': 'USE_FIRST_LINE_INDENT',
    'code_font_size': 'CODE_FONT_SIZE_PT',
    'code_line_spacing': 'CODE_LINE_SPACING',
}
# Параметры командной строки сильнее профиля и шапки документа
_cli_overrides = {}


def apply_cli_overrides(args, settings=DocumentSettings):
    """Применяет параметры командной строки к DocumentSettings и запоминает их"""
    for arg_name, setting_name in CLI_TO_SETTINGS.items():
        value = getattr(args, arg_name, None)
        if value is not None:
            _cli_overrides[setting_name] = value
            setattr(settings, setting_name, value)


@contextmanager
def settings_scope():
    """Параметры, заданные профилем и шапкой, действуют только на время сборки одного документа"""
    snapshot = {k: getattr(DocumentSettings, k) for k in vars(DocumentSettings) if k.isupper()}
    try:
        yield
    finally:
        for key, value in snapshot.items():
            setattr(DocumentSettings, key, value)


def load_user_defaults(base_dir: Optional[str]) -> dict:
    """marktogost.yaml из корня проекта, поверх него — из папки документа"""
    result = {}
    for folder in dict.fromkeys([PROJECT_ROOT, base_dir or PROJECT_ROOT]):
        path = os.path.join(folder, DEFAULTS_FILE)
        if os.path.isfile(path):
            with open(path, encoding="utf-8") as f:
                data = yaml.safe_load(f) or {}
            if not isinstance(data, dict):
                raise ProfileError(f"{path}: ожидались строки «ключ: значение»")
            result = merge(result, data)
    return result


def _resolve_profile(front: dict, base_dir: Optional[str], warnings: List[str]):
    """Профиль с учётом marktogost.yaml и шапки: (профиль, поля листов, вставные листы)"""
    defaults = load_user_defaults(base_dir)
    profile = load_profile(front.get("тип") or defaults.get("тип"))
    d_fmt, d_fields = split_settings(defaults)
    f_fmt, f_fields = split_settings(front)
    for key in unknown_format_keys({**d_fmt, **f_fmt}):
        warnings.append(f"неизвестный параметр оформления «{key}»")
    formatting = merge(merge(profile.formatting, d_fmt), f_fmt)
    fields = {**profile.fields, **d_fields, **f_fields}
    profile = dataclasses.replace(profile, formatting=formatting)
    sheets = _sheet_specs(profile, {**defaults, **front}, warnings)
    return profile, fields, sheets


def _sheet_specs(profile: Profile, settings: dict, warnings: List[str]) -> List[SheetSpec]:
    """Титульный лист, листы профиля с «включён: да» и листы из ключа «листы» шапки"""
    specs: List[SheetSpec] = []

    title = settings.get("титульный_лист")
    if title is None:
        if profile.title_sheet():
            specs.append(profile.title_sheet())
    elif str(title).strip().lower() not in ("", "нет", "false", "none"):
        specs.append(SheetSpec(name=TITLE_SHEET, template=str(title), place="начало", page_numbers=False,
                               required_fields=profile.required_fields, enabled=True))

    requested = settings.get("листы") or []
    if not isinstance(requested, list):
        raise ProfileError("«листы» в шапке — список: имена листов профиля или словари с ключом «шаблон»")
    named = {str(item) for item in requested if isinstance(item, str)}
    specs += [spec for name, spec in profile.sheets.items() if spec.enabled and name not in named]
    for item in requested:
        if isinstance(item, str):
            spec = profile.sheets.get(item)
            if spec is None:
                known = ", ".join(profile.sheets) or "нет"
                warnings.append(f"в профиле «{profile.name}» нет листа «{item}» (есть: {known})")
            else:
                specs.append(spec)
        elif isinstance(item, dict):
            name = str(item.get("имя") or os.path.splitext(os.path.basename(str(item.get("шаблон", ""))))[0])
            specs.append(parse_sheet(name, item))
    return specs


def _add_signature_table(doc, metadata):
    """Таблица «Выполнил / Проверил» в конце документа (флаг is_table)"""
    doc.add_page_break()
    table = doc.add_table(rows=2, cols=2)
    try:
        table.style = 'Table Grid'
    except Exception:
        pass
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    col_width = Cm((DocumentSettings.PAGE_WIDTH_CM - DocumentSettings.LEFT_MARGIN_CM
                    - DocumentSettings.RIGHT_MARGIN_CM) / 2.0)
    for col in table.columns:
        for cell in col.cells:
            cell.width = col_width

    table.cell(0, 0).text = 'Выполнил:'
    table.cell(0, 1).text = metadata['fio']
    table.cell(1, 0).text = 'Проверил:'
    table.cell(1, 1).text = metadata['teacher']

    for row_idx, row in enumerate(table.rows):
        if row_idx == 0:
            set_repeat_table_header(row)
        for cell in row.cells:
            for p in cell.paragraphs:
                p.style = doc.styles['Normal']
                p.paragraph_format.first_line_indent = None
                p.paragraph_format.left_indent = Cm(0)
                p.paragraph_format.line_spacing = 1.5
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                for run in p.runs:
                    set_run_font(run, size_pt=12)
    set_table_borders(table)


def _build(md_text: str, base_dir: Optional[str], profile: Optional[Profile], fields: dict,
           sheets: List[SheetSpec], warnings: List[str]) -> Document:
    metadata = extract_metadata(md_text)
    if profile is not None:
        metadata['numerate'] = metadata['numerate'] and profile.page_numbers

    # Служебные строки не попадают в текст
    md_text = re.sub(r'__?ФИО__?[:\-]?.*\n?', '', md_text)
    md_text = re.sub(r'__?Группа__?[:\-]?.*\n?', '', md_text)
    md_text = re.sub(r'\[//\]:\s*#\s*\w+\s*=.*\n?', '', md_text, flags=re.I)

    doc = Document()

    # Страница и стили
    for s in doc.sections:
        s.page_height = Cm(DocumentSettings.PAGE_HEIGHT_CM)
        s.page_width = Cm(DocumentSettings.PAGE_WIDTH_CM)
        s.left_margin = Cm(DocumentSettings.LEFT_MARGIN_CM)
        s.right_margin = Cm(DocumentSettings.RIGHT_MARGIN_CM)
        s.top_margin = Cm(DocumentSettings.TOP_MARGIN_CM)
        s.bottom_margin = Cm(DocumentSettings.BOTTOM_MARGIN_CM)
    normal = doc.styles['Normal']
    normal.font.name = DocumentSettings.FONT_NAME
    normal.font.size = Pt(DocumentSettings.FONT_SIZE_PT)
    normal.paragraph_format.first_line_indent = Cm(DocumentSettings.FIRST_LINE_INDENT_CM)
    normal.paragraph_format.line_spacing = DocumentSettings.LINE_SPACING
    reset_heading_styles(doc)

    has_title = any(spec.name == TITLE_SHEET for spec in sheets)

    # Блок «Выполнил студент» в начале документа (ФИО без таблицы подписей)
    if not has_title and not metadata['is_table'] and (metadata['fio'] or metadata['group']):
        group_full = f"090302-{metadata['group']}" if metadata['group'] else ""
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        set_run_font(p.add_run(f"Выполнил студент: {metadata['fio']}\n"),
                     size_pt=DocumentSettings.FONT_SIZE_PT, bold=True)
        set_run_font(p.add_run(f"Группа: {group_full}"), size_pt=DocumentSettings.FONT_SIZE_PT, bold=True)

    # Разбор и структура
    blocks = MarkdownParser(md_text).parse()
    if profile is not None:
        blocks = group_sections(blocks, profile.sections_new_page)
        if profile.toc:
            blocks = insert_toc(blocks, profile)
        warnings.extend(check_structure(blocks, profile))
        for spec in sheets:
            warnings.extend(check_fields(fields, spec.required_fields, spec.name))
        blocks, sheet_warnings = place_sheets(blocks, sheets)
        warnings.extend(sheet_warnings)

    toc_entries = []

    def collect_headings(block):
        if isinstance(block, Section):
            for b in block.blocks:
                collect_headings(b)
        elif isinstance(block, HeadingBlock):
            toc_entries.append((get_toc_level(block.text), block.text))

    for block in blocks:
        collect_headings(block)

    # Рендеринг
    renderer = DocumentRenderer(doc, toc_entries, use_headings=metadata['use_headings'])
    renderer.base_dir = base_dir
    renderer.fields = fields
    renderer.titles_dir = TITLES_DIR
    renderer.warnings = warnings
    for block in blocks:
        renderer.render_block(block)

    if metadata['is_table']:
        _add_signature_table(doc, metadata)
    if metadata['numerate']:
        add_page_number_centered(doc)
    return doc


def create_document(md_text: str, base_dir: Optional[str] = None,
                    warnings: Optional[List[str]] = None) -> Document:
    """
    Создаёт DOCX из Markdown.

    Если в начале текста есть шапка «---», документ собирается по профилю его типа (ключ «тип»,
    по умолчанию — «гост»); без шапки — с параметрами по умолчанию, как раньше.

    Args:
        md_text: текст Markdown
        base_dir: папка документа — для marktogost.yaml, рисунков и шаблонов листов
        warnings: список, в который добавляются замечания по документу
    """
    warnings = warnings if warnings is not None else []
    front, body = split_front_matter(md_text)
    profile, fields, sheets = (None, {}, [])
    if front:
        profile, fields, sheets = _resolve_profile(front, base_dir, warnings)

    with settings_scope():
        if profile is not None:
            for key, value in profile.formatting.items():
                if key in FORMAT_KEYS and value is not None:
                    setattr(DocumentSettings, FORMAT_KEYS[key], value)
            for attr, value in _cli_overrides.items():
                setattr(DocumentSettings, attr, value)
        return _build(body, base_dir, profile, fields, sheets, warnings)


def finalize_docx(docx_path: str, pdf_path: Optional[str] = None) -> bool:
    """Доводка в Microsoft Word: оглавление, «Продолжение таблицы N», при необходимости PDF"""
    if os.name != "nt":
        print("⚠️ Доводка выполняется через Microsoft Word и доступна только в Windows")
        return False
    cmd = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", FINALIZE_SCRIPT, docx_path]
    if pdf_path:
        cmd += ["-Pdf", pdf_path]
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    output = (result.stdout + result.stderr).strip()
    if result.returncode != 0:
        print(f"❌ Доводка в Word не удалась: {output}")
        return False
    print(f"📝 Word: {output}")
    return True


def _display(path: str) -> str:
    """Путь для сообщений: относительно проекта, если файл внутри него"""
    rel = os.path.relpath(path, PROJECT_ROOT)
    return path if rel.startswith("..") else rel


def resolve_input(path: str) -> str:
    """Путь к .md: абсолютный, относительно текущей папки или имя файла в input/"""
    if os.path.isabs(path) or os.path.exists(path):
        return os.path.abspath(path)
    return os.path.join(INPUT_DIR, path)


def process_md_file(input_path: str, output_path: Optional[str] = None,
                    finalize: bool = False, pdf: bool = False) -> bool:
    """Собирает один .md в output/ (или в output_path). Возвращает True при успехе"""
    input_full_path = resolve_input(input_path)
    if not os.path.exists(input_full_path):
        print(f"❌ Файл не найден: {input_full_path}")
        return False

    if output_path is None:
        name = os.path.splitext(os.path.basename(input_full_path))[0]
        output_full_path = os.path.join(OUTPUT_DIR, f"{name}.docx")
    else:
        output_full_path = output_path if os.path.isabs(output_path) else os.path.join(OUTPUT_DIR, output_path)
    os.makedirs(os.path.dirname(output_full_path), exist_ok=True)

    try:
        with open(input_full_path, "r", encoding="utf-8") as f:
            md_text = f.read()

        warnings: List[str] = []
        doc = create_document(md_text, base_dir=os.path.dirname(input_full_path), warnings=warnings)
        doc.save(output_full_path)
        print(f"✅ {_display(input_full_path)} → {_display(output_full_path)}")
        for w in warnings:
            print(f"   ⚠️ {w}")

        if finalize or pdf:
            pdf_path = os.path.splitext(output_full_path)[0] + ".pdf" if pdf else None
            return finalize_docx(output_full_path, pdf_path)
        return True

    except Exception as e:
        print(f"❌ Ошибка при обработке {os.path.basename(input_full_path)}: {e}")
        import traceback
        traceback.print_exc()
        return False


def print_types():
    print("Типы документов (ключ «тип:» в шапке):\n")
    for p in available_profiles():
        names = ", ".join(p.names) or p.key
        print(f"  {p.name}")
        print(f"      тип: {names}")
        if p.description:
            print(f"      {p.description}")
        if p.structure:
            print(f"      структура: {' → '.join(i.title for i in p.structure)}")
        print()


def scaffold(type_name: str, path: str) -> bool:
    """Создаёт заготовку документа: шапка с полями титульного листа и разделы по профилю"""
    profile = load_profile(type_name)
    target = path if os.path.isabs(path) or os.path.dirname(path) else os.path.join(INPUT_DIR, path)
    if os.path.exists(target):
        print(f"❌ Файл уже существует: {target}")
        return False

    defaults = load_user_defaults(os.path.dirname(os.path.abspath(target)))
    _, known = split_settings(defaults)
    lines = ["---", f"тип: {profile.names[0] if profile.names else profile.key}"]
    for name in profile.required_fields:
        if name not in known:
            lines.append(f"{name}: ")
    optional = [k for k in profile.fields if k not in profile.required_fields and k not in known
                and not profile.fields[k]]  # поля со значением по умолчанию не подсказываем
    if optional:
        lines.append("# необязательные поля:")
        lines += [f"# {k}: " for k in optional]
    if profile.sheets:
        listed = set(profile.required_fields) | set(optional) | set(known) | set(profile.fields)
        lines.append(f"# дополнительные листы ({', '.join(profile.sheets)}) — включаются так:")
        lines.append(f"# листы: [{', '.join(profile.sheets)}]")
        for name, spec in profile.sheets.items():
            path = os.path.join(TITLES_DIR, spec.template)
            extra = [f for f in template_fields(path) if f not in listed] if os.path.isfile(path) else []
            if extra:
                lines.append(f"# поля листа «{name}»:")
                lines += [f"# {f}: " for f in extra]
                listed.update(extra)
    lines += ["---", ""]

    for item in profile.structure:
        if item.main_part:
            lines += ["# 1 НАЗВАНИЕ ПЕРВОГО РАЗДЕЛА", "", "## 1.1 Название подраздела", "", "Текст.", ""]
        elif item.prefix:
            lines += [f"# {item.title} А", "", "Текст приложения.", ""]
        else:
            lines += [f"# {item.title}", "", "Текст.", ""]
    if not profile.structure:
        lines += ["# Название документа", "", "Текст.", ""]

    os.makedirs(os.path.dirname(os.path.abspath(target)), exist_ok=True)
    with open(target, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"✅ Создан {target}")
    return True


def main():
    # Консоль Windows (cp1251) не умеет выводить эмодзи — без этого скрипт падает до обработки файлов
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(
        description="MarkToGost: конвертер Markdown → DOCX по ГОСТ 7.32",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Примеры:
  python -m MarkToGost.main                         все .md из папки input/
  python -m MarkToGost.main отчёт.md --pdf          один файл, доводка в Word и PDF
  python -m MarkToGost.main --new курсовая моя.md   заготовка курсовой в input/моя.md
  python -m MarkToGost.main --types                 список типов документов
""")
    parser.add_argument("file", nargs="?", default=None,
                        help="Markdown-файл: путь или имя файла в папке input/")
    parser.add_argument("--output", "-o", default=None, help="имя или путь выходного DOCX (по умолчанию output/<имя>.docx)")
    parser.add_argument("--finalize", action="store_true",
                        help="доводка в Word: оглавление и подписи «Продолжение таблицы N» (Windows)")
    parser.add_argument("--pdf", action="store_true", help="доводка в Word и сохранение PDF рядом с DOCX")
    parser.add_argument("--new", metavar="ТИП", default=None, help="создать заготовку документа указанного типа")
    parser.add_argument("--types", action="store_true", help="показать типы документов")

    doc_group = parser.add_argument_group('Оформление (сильнее профиля и шапки документа)')
    doc_group.add_argument('--font-name', default=None, help='шрифт')
    doc_group.add_argument('--font-size', default=None, type=float, help='размер шрифта, pt')
    doc_group.add_argument('--line-spacing', default=None, type=float, help='межстрочный интервал')
    doc_group.add_argument('--first-line-indent', default=None, type=float, help='абзацный отступ, см')
    doc_group.add_argument('--use-first-line-indent', default=None, type=lambda x: x.lower() == 'true',
                           help='абзацный отступ: true/false')
    doc_group.add_argument('--left-margin', default=None, type=float, help='левое поле, см')
    doc_group.add_argument('--right-margin', default=None, type=float, help='правое поле, см')
    doc_group.add_argument('--top-margin', default=None, type=float, help='верхнее поле, см')
    doc_group.add_argument('--bottom-margin', default=None, type=float, help='нижнее поле, см')
    doc_group.add_argument('--image-width', default=None, type=float, help='ширина рисунков, доля ширины текста')
    doc_group.add_argument('--caption-font-size', default=None, type=float, help='размер шрифта подписей рисунков, pt')
    doc_group.add_argument('--caption-italic', default=None, type=lambda x: x.lower() == 'true',
                           help='курсив подписей рисунков: true/false')
    doc_group.add_argument('--table-font-size', default=None, type=float, help='размер шрифта таблиц, pt')
    doc_group.add_argument('--code-font-size', default=None, type=float, help='размер шрифта блоков кода, pt')
    doc_group.add_argument('--code-line-spacing', default=None, type=float, help='интервал блоков кода')

    args = parser.parse_args()

    try:
        if args.types:
            print_types()
            return 0
        if args.new:
            if not args.file:
                parser.error("укажите имя файла: --new ТИП ФАЙЛ.md")
            return 0 if scaffold(args.new, args.file) else 1
    except ProfileError as e:
        print(f"❌ {e}")
        return 1

    apply_cli_overrides(args)

    if args.file:
        print(f"📄 Обработка файла: {args.file}")
        ok = process_md_file(args.file, args.output, finalize=args.finalize, pdf=args.pdf)
        return 0 if ok else 1

    os.makedirs(INPUT_DIR, exist_ok=True)
    all_md = sorted(f for f in os.listdir(INPUT_DIR) if f.lower().endswith(".md") and not f.startswith("."))
    if not all_md:
        print(f"⚠️ Не найдено ни одного Markdown-файла в {INPUT_DIR}")
        return 1
    print(f"🔍 Найдено файлов в {INPUT_DIR}: {len(all_md)}\n")
    ok_count = sum(process_md_file(os.path.join(INPUT_DIR, f), finalize=args.finalize, pdf=args.pdf)
                   for f in all_md)
    print(f"\n🎉 Обработано: {ok_count}/{len(all_md)} файлов")
    return 0 if ok_count == len(all_md) else 1


if __name__ == "__main__":
    sys.exit(main())
