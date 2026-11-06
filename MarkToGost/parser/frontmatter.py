# parser/frontmatter.py
"""Шапка документа в формате YAML между строками «---» в самом начале файла"""

import re
from typing import Tuple

import yaml

FRONT_MATTER = re.compile(r"\A﻿?---[ \t]*\r?\n(.*?)\r?\n---[ \t]*(?:\r?\n|\Z)", re.S)


class FrontMatterError(ValueError):
    pass


class SourceFloat(float):
    """Число из шапки, которое помнит, как оно записано: «удк: 004.8» остаётся «004.8» в тексте"""
    source = ""

    def __str__(self):
        return self.source or float.__str__(self)


class SourceInt(int):
    source = ""

    def __str__(self):
        return self.source or int.__str__(self)


class _Loader(yaml.SafeLoader):
    pass


def _number(cls, construct):
    def constructor(loader, node):
        value = cls(construct(loader, node))
        value.source = node.value
        return value
    return constructor


_Loader.add_constructor("tag:yaml.org,2002:float", _number(SourceFloat, yaml.SafeLoader.construct_yaml_float))
_Loader.add_constructor("tag:yaml.org,2002:int", _number(SourceInt, yaml.SafeLoader.construct_yaml_int))


def load_yaml(text: str):
    """YAML шапки и marktogost.yaml: числа остаются числами для параметров оформления,
    а в полях выводятся как записаны («004.8», «0302»)"""
    return yaml.load(text, Loader=_Loader)


def split_front_matter(text: str) -> Tuple[dict, str]:
    """Возвращает (параметры шапки, текст документа без шапки)"""
    match = FRONT_MATTER.match(text)
    if not match:
        return {}, text
    try:
        data = load_yaml(match.group(1)) or {}
    except yaml.YAMLError as e:
        raise FrontMatterError(f"ошибка в шапке документа: {e}") from e
    if not isinstance(data, dict):
        raise FrontMatterError("шапка документа должна состоять из строк «ключ: значение»")
    return data, text[match.end():]
