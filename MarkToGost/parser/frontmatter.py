# parser/frontmatter.py
"""Шапка документа в формате YAML между строками «---» в самом начале файла"""

import re
from typing import Tuple

import yaml

FRONT_MATTER = re.compile(r"\A﻿?---[ \t]*\r?\n(.*?)\r?\n---[ \t]*(?:\r?\n|\Z)", re.S)


class FrontMatterError(ValueError):
    pass


def split_front_matter(text: str) -> Tuple[dict, str]:
    """Возвращает (параметры шапки, текст документа без шапки)"""
    match = FRONT_MATTER.match(text)
    if not match:
        return {}, text
    try:
        data = yaml.safe_load(match.group(1)) or {}
    except yaml.YAMLError as e:
        raise FrontMatterError(f"ошибка в шапке документа: {e}") from e
    if not isinstance(data, dict):
        raise FrontMatterError("шапка документа должна состоять из строк «ключ: значение»")
    return data, text[match.end():]
