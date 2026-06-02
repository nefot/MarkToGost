from dataclasses import dataclass
from typing import Optional, List
from enum import Enum


@dataclass
class BaseBlock:
    """Базовый класс блока"""
    pass


@dataclass
class TextBlock(BaseBlock):
    """Блок обычного текста"""
    text: str


class CellAlign(Enum):
    LEFT = "left"
    CENTER = "center"
    RIGHT = "right"


@dataclass
class HtmlTableCell:
    text: str
    colspan: int = 1
    rowspan: int = 1
    bold: bool = False
    italic: bool = False
    underline: bool = False
    align: CellAlign = CellAlign.LEFT
    formula: Optional[str] = None  # LaTeX формула

@dataclass
class HtmlTableRow:
    cells: List['HtmlTableCell']


@dataclass
class HtmlTableBlock(BaseBlock):
    rows: List['HtmlTableRow']
    caption: Optional[str] = None
    transparent: bool = False  # Скрыть границы таблицы


@dataclass
class HeadingBlock(BaseBlock):
    """Блок заголовка"""
    text: str
    level: int


@dataclass
class ImageBlock(BaseBlock):
    """Блок изображения"""
    path: str
    caption: str
    img_id: Optional[str] = None


@dataclass
class ListBlock(BaseBlock):
    """Блок списка"""
    items: List[str]
    ordered: bool = False


@dataclass
class TableBlock(BaseBlock):
    """Блок таблицы"""
    rows: List[str]
    caption: str = None


@dataclass
class CodeBlock(BaseBlock):
    """Блок кода"""
    code: str
    language: str = ""


@dataclass
class FormulaBlock(BaseBlock):
    """Блок формулы LaTeX (Pandoc + ГОСТ)"""
    latex: str
    formula_id: Optional[str] = None
    explanation: Optional[str] = None
    number: Optional[str] = None


@dataclass
class Section(BaseBlock):
    """Раздел документа с уникальным идентификатором и блоками"""
    section_id: str  # Уникальный идентификатор раздела (e.g. "Литература")
    blocks: List[BaseBlock]  # Блоки внутри раздела
    heading_level: int = 4  # Уровень заголовка раздела (2-6), по умолчанию 4 (####)
    add_page_breaks: bool = True  # Добавлять page breaks перед/после раздела (default=True)
