"""Профили типов документов: оформление, структура и титульный лист.

Профиль — YAML-файл в этой папке. Профили наследуются (ключ «наследует»): словари сливаются
рекурсивно, списки и значения заменяются целиком.
"""

import datetime
import os
from dataclasses import dataclass, field
from typing import Dict, List, Optional

import yaml

PROFILES_DIR = os.path.dirname(os.path.abspath(__file__))
TITLES_DIR = os.path.join(PROFILES_DIR, "titles")
DEFAULT_PROFILE = "гост"

# Ключ раздела «оформление» → атрибут DocumentSettings
FORMAT_KEYS = {
    "шрифт": "FONT_NAME",
    "размер_шрифта": "FONT_SIZE_PT",
    "интервал": "LINE_SPACING",
    "отступ_абзаца": "FIRST_LINE_INDENT_CM",
    "красная_строка": "USE_FIRST_LINE_INDENT",
    "поле_слева": "LEFT_MARGIN_CM",
    "поле_справа": "RIGHT_MARGIN_CM",
    "поле_сверху": "TOP_MARGIN_CM",
    "поле_снизу": "BOTTOM_MARGIN_CM",
    "ширина_рисунков": "IMAGE_WIDTH_FRACTION",
    "размер_подписей": "CAPTION_FONT_SIZE_PT",
    "курсив_подписей": "CAPTION_ITALIC",
    "размер_шрифта_таблиц": "TABLE_FONT_SIZE_PT",
    "размер_шрифта_кода": "CODE_FONT_SIZE_PT",
    "интервал_кода": "CODE_LINE_SPACING",
}
# Прочие параметры оформления, которыми управляет сборка документа, а не DocumentSettings
LAYOUT_KEYS = {"нумерация_страниц", "содержание", "заголовок_содержания", "разделы_с_новой_страницы"}

# Служебные ключи шапки документа — не поля листов
SERVICE_KEYS = {"тип", "оформление", "листы", "титульный_лист"}
TITLE_SHEET = "титульный лист"


class ProfileError(Exception):
    pass


@dataclass
class StructureItem:
    """Структурный элемент документа: ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ, основная часть…"""
    title: str
    synonyms: List[str] = field(default_factory=list)
    required: bool = True
    before_toc: bool = False   # элемент располагается перед содержанием (реферат ВКР)
    prefix: bool = False       # совпадение по началу заголовка (ПРИЛОЖЕНИЕ А, ПРИЛОЖЕНИЕ Б…)
    main_part: bool = False    # «основная часть» — разделы с номером

    def matches(self, heading: str) -> bool:
        text = _norm(heading)
        if self.main_part:
            return text[:1].isdigit()
        names = [_norm(n) for n in [self.title, *self.synonyms]]
        if self.prefix:
            return any(text.startswith(n) for n in names)
        return text in names


@dataclass
class SheetSpec:
    """Вставной лист: шаблон, место в документе, нумерация, обязательные поля"""
    name: str
    template: str
    place: str = "конец"            # начало | конец | после | перед
    anchor: Optional[str] = None    # заголовок раздела, «титульный лист» или «содержание»
    page_numbers: bool = True
    required_fields: List[str] = field(default_factory=list)
    enabled: bool = False           # включён без упоминания в шапке


def parse_sheet(name: str, spec: dict) -> SheetSpec:
    """Лист из словаря профиля или шапки: шаблон, после/перед/место, номер_страницы…"""
    if not isinstance(spec, dict) or not spec.get("шаблон"):
        raise ProfileError(f"лист «{name}»: нужен ключ «шаблон»")
    if spec.get("после"):
        place, anchor = "после", str(spec["после"])
    elif spec.get("перед"):
        place, anchor = "перед", str(spec["перед"])
    else:
        place, anchor = str(spec.get("место", "конец")).lower(), None
        if place not in ("начало", "конец"):
            raise ProfileError(f"лист «{name}»: место «{place}» — ожидалось «начало», «конец», «после» или «перед»")
    return SheetSpec(
        name=name, template=str(spec["шаблон"]), place=place, anchor=anchor,
        page_numbers=_flag(spec.get("номер_страницы", True)),
        required_fields=list(spec.get("обязательные_поля", [])),
        enabled=_flag(spec.get("включён", False)),
    )


def _flag(value) -> bool:
    if isinstance(value, str):
        return value.strip().lower() not in ("нет", "false", "no", "0", "")
    return bool(value)


@dataclass
class Profile:
    key: str
    name: str
    description: str
    names: List[str]
    formatting: Dict[str, object]
    structure: List[StructureItem]
    title_template: Optional[str]
    required_fields: List[str]
    fields: Dict[str, str]
    sheets: Dict[str, SheetSpec] = field(default_factory=dict)

    def title_sheet(self) -> Optional[SheetSpec]:
        if not self.title_template:
            return None
        return SheetSpec(name=TITLE_SHEET, template=self.title_template, place="начало",
                         page_numbers=False, required_fields=self.required_fields, enabled=True)

    @property
    def toc(self) -> bool:
        return bool(self.formatting.get("содержание", False))

    @property
    def toc_title(self) -> str:
        return str(self.formatting.get("заголовок_содержания", "СОДЕРЖАНИЕ"))

    @property
    def sections_new_page(self) -> bool:
        return bool(self.formatting.get("разделы_с_новой_страницы", True))

    @property
    def page_numbers(self) -> bool:
        return bool(self.formatting.get("нумерация_страниц", True))


def _norm(text: str) -> str:
    return " ".join(str(text).replace("Ё", "Е").replace("ё", "е").upper().split())


def merge(base: dict, override: dict) -> dict:
    """Рекурсивное слияние словарей: override поверх base"""
    result = dict(base)
    for key, value in (override or {}).items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = merge(result[key], value)
        else:
            result[key] = value
    return result


def _read_yaml(path: str) -> dict:
    with open(path, encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    if not isinstance(data, dict):
        raise ProfileError(f"{os.path.basename(path)}: ожидался словарь параметров")
    return data


def _profile_files() -> Dict[str, str]:
    """Все имена профилей (имя файла и синонимы из «имена») → путь к файлу"""
    result = {}
    for fname in sorted(os.listdir(PROFILES_DIR)):
        if not fname.endswith(".yaml"):
            continue
        path = os.path.join(PROFILES_DIR, fname)
        key = os.path.splitext(fname)[0]
        result[_norm(key)] = path
        for alias in _read_yaml(path).get("имена", []):
            result[_norm(alias)] = path
    return result


def _load_raw(path: str, seen=()) -> dict:
    data = _read_yaml(path)
    parent = data.get("наследует")
    if not parent:
        return data
    files = _profile_files()
    parent_path = files.get(_norm(parent))
    if not parent_path:
        raise ProfileError(f"{os.path.basename(path)}: неизвестный профиль-основа «{parent}»")
    if parent_path in seen:
        raise ProfileError(f"циклическое наследование профилей: {os.path.basename(path)}")
    base = _load_raw(parent_path, seen + (path,))
    # Имена и описание не наследуются
    for key in ("имена", "название", "описание"):
        base.pop(key, None)
    return merge(base, data)


def _structure(items) -> List[StructureItem]:
    result = []
    for item in items or []:
        if item == "основная_часть":
            result.append(StructureItem(title="основная часть (разделы с номером)", main_part=True))
        elif isinstance(item, str):
            result.append(StructureItem(title=item))
        else:
            result.append(StructureItem(
                title=item["элемент"],
                synonyms=list(item.get("синонимы", [])),
                required=bool(item.get("обязательный", True)),
                before_toc=bool(item.get("перед_содержанием", False)),
                prefix=bool(item.get("по_началу", False)),
            ))
    return result


def _field_text(value) -> str:
    """Значение поля как текст: список — строки через перенос"""
    if value is None:
        return ""
    if isinstance(value, (list, tuple)):
        return "\n".join(_field_text(v) for v in value)
    return str(value).rstrip("\n")


def _field_defaults(values: dict) -> Dict[str, str]:
    today = datetime.date.today()
    result = {}
    for key, value in (values or {}).items():
        if value == "текущий год":
            value = today.year
        result[key] = "" if value is None else str(value)
    return result


def available_profiles() -> List[Profile]:
    """Все профили (по одному на файл), отсортированные по имени файла"""
    seen, result = set(), []
    for path in _profile_files().values():
        if path not in seen:
            seen.add(path)
            result.append(load_profile(os.path.splitext(os.path.basename(path))[0]))
    return sorted(result, key=lambda p: p.key)


def load_profile(name: Optional[str]) -> Profile:
    files = _profile_files()
    path = files.get(_norm(name or DEFAULT_PROFILE))
    if not path:
        known = ", ".join(sorted({os.path.splitext(os.path.basename(p))[0] for p in files.values()}))
        raise ProfileError(f"неизвестный тип документа «{name}». Доступны: {known}")
    raw = _load_raw(path)
    title = raw.get("титульный_лист") or {}
    return Profile(
        key=os.path.splitext(os.path.basename(path))[0],
        name=str(raw.get("название", "")),
        description=str(raw.get("описание", "")),
        names=[str(n) for n in raw.get("имена", [])],
        formatting=dict(raw.get("оформление") or {}),
        structure=_structure(raw.get("структура")),
        title_template=title.get("шаблон"),
        required_fields=list(title.get("обязательные_поля", [])),
        fields=_field_defaults(raw.get("поля")),
        sheets={name: parse_sheet(name, spec) for name, spec in (raw.get("листы") or {}).items()},
    )


def split_settings(values: dict):
    """Разделяет словарь шапки на параметры оформления и поля титульного листа.

    Параметры оформления можно писать как в разделе «оформление», так и на верхнем уровне.
    """
    formatting = dict(values.get("оформление") or {})
    fields = {}
    for key, value in values.items():
        if key in SERVICE_KEYS:
            continue
        if key in FORMAT_KEYS or key in LAYOUT_KEYS:
            formatting[key] = value
        else:
            fields[key] = _field_text(value)
    return formatting, fields


def unknown_format_keys(formatting: dict) -> List[str]:
    return [k for k in formatting if k not in FORMAT_KEYS and k not in LAYOUT_KEYS]
