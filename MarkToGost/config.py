# config.py

from docx.shared import Pt, Cm


class DocumentSettings:
    """Настройки документа"""
    FONT_NAME = "Times New Roman"
    FONT_SIZE_PT = 14
    LINE_SPACING = 1.5
    FIRST_LINE_INDENT_CM = 1.25

    # Размеры страницы A4
    PAGE_WIDTH_CM = 21.0
    PAGE_HEIGHT_CM = 29.7

    # Поля по ГОСТ 7.32-2001
    LEFT_MARGIN_CM = 3.0   # не менее 30 мм
    RIGHT_MARGIN_CM = 1.0  # не менее 10 мм
    TOP_MARGIN_CM = 2.0    # не менее 20 мм
    BOTTOM_MARGIN_CM = 2.0 # не менее 20 мм

    # Настройки изображений
    IMAGE_WIDTH_FRACTION = 0.70

    # Настройки подписей
    CAPTION_FONT_SIZE_PT = 12
    CAPTION_ITALIC = True
    USE_FIRST_LINE_INDENT = True  # False — без отступа

    # Настройки таблиц
    TABLE_FONT_SIZE_PT = 12

class CaptionSettings:
    """Настройки подписей к элементам"""
    FONT_SIZE = Pt(12)
    ITALIC = True


class TextSettings:
    """Настройки текста"""
    FONT_NAME = "Times New Roman"
    FONT_SIZE = Pt(14)
    LINE_SPACING = 1.5
    FIRST_LINE_INDENT = Cm(1.25)

