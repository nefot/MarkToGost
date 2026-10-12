# structure.py
"""Структура документа по профилю: разделы из «# Заголовок», содержание, проверка состава"""

from typing import List, Tuple

from MarkToGost.parser.blocks import BaseBlock, HeadingBlock, Section, SheetBlock
from MarkToGost.profiles import Profile, SheetSpec, _norm


def group_sections(blocks: List[BaseBlock], new_page: bool) -> List[BaseBlock]:
    """Заголовок «# Название» начинает раздел, в который входит всё до следующего такого заголовка"""
    result: List[BaseBlock] = []
    current = None
    for block in blocks:
        if isinstance(block, HeadingBlock) and block.marker == 1:
            current = Section(section_id=block.text, blocks=[], heading_level=1, add_page_breaks=new_page)
            result.append(current)
        elif isinstance(block, Section):
            current = None
            result.append(block)
        elif current is not None:
            current.blocks.append(block)
        else:
            result.append(block)
    return result


def _top_sections(blocks: List[BaseBlock]) -> List[Section]:
    return [b for b in blocks if isinstance(b, Section) and b.section_id and not b.toc]


def insert_toc(blocks: List[BaseBlock], profile: Profile) -> List[BaseBlock]:
    """Вставляет содержание, если в документе его нет: после элементов «перед_содержанием»"""
    if any(isinstance(b, Section) and b.toc for b in blocks):
        return blocks
    position = 0
    for i, block in enumerate(blocks):
        if isinstance(block, Section) and any(
                item.before_toc and item.matches(block.section_id) for item in profile.structure):
            position = i + 1
    toc = Section(section_id=profile.toc_title, blocks=[], heading_level=1, toc=True)
    return blocks[:position] + [toc] + blocks[position:]


def check_structure(blocks: List[BaseBlock], profile: Profile) -> List[str]:
    """Сверяет разделы документа со структурой профиля, возвращает замечания"""
    titles = [s.section_id for s in _top_sections(blocks)]
    warnings = []
    last_found, last_title = -1, None
    for item in profile.structure:
        index = next((i for i, t in enumerate(titles) if item.matches(t)), None)
        if index is None:
            if item.required:
                warnings.append(f"нет обязательного раздела «{item.title}»")
            continue
        if index < last_found:
            warnings.append(f"раздел «{titles[index]}» должен идти после «{last_title}»")
        else:
            last_found, last_title = index, titles[index]
    return warnings


def check_fields(fields: dict, required: List[str], sheet: str = "титульный лист") -> List[str]:
    return [f"не заполнено поле «{name}» ({sheet})"
            for name in required if not str(fields.get(name, "")).strip()]


def _is_anchor(block: BaseBlock, anchor: str) -> bool:
    """Подходит ли блок как место вставки: лист по имени, содержание, раздел по заголовку или его началу"""
    target = _norm(anchor)
    if isinstance(block, SheetBlock):
        return bool(block.name) and _norm(block.name) == target
    if isinstance(block, Section):
        if block.toc and target in ("СОДЕРЖАНИЕ", "ОГЛАВЛЕНИЕ"):
            return True
        title = _norm(block.section_id)
        return title == target or title.startswith(target)
    return False


def place_sheets(blocks: List[BaseBlock], specs: List[SheetSpec]) -> Tuple[List[BaseBlock], List[str]]:
    """Расставляет вставные листы по местам из спецификаций (порядок спецификаций сохраняется)"""
    result = list(blocks)
    warnings = []
    for spec in specs:
        sheet = SheetBlock(template=spec.template, page_numbers=spec.page_numbers, name=spec.name)
        if spec.place == "начало":
            index = 0
            while index < len(result) and isinstance(result[index], SheetBlock):
                index += 1
        elif spec.place == "конец":
            index = len(result)
        else:
            found = next((i for i, b in enumerate(result) if _is_anchor(b, spec.anchor)), None)
            if found is None:
                warnings.append(f"лист «{spec.name}» не вставлен: нет раздела «{spec.anchor}»")
                continue
            if spec.place == "перед":
                index = found
            else:
                index = found + 1
                while index < len(result) and isinstance(result[index], SheetBlock):
                    index += 1
        result.insert(index, sheet)
    return result, warnings
