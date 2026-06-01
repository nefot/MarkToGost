# -*- coding: utf-8 -*-
"""
Тесты для вспомогательных функций документа (utils/document_helpers.py)
"""

import pytest
from MarkToGost.utils.document_helpers import (
    split_md_table_row, is_md_table_separator, is_md_table_row,
    normalize_table_caption, replace_image_refs
)


class TestSplitMdTableRow:
    """Тесты для split_md_table_row"""
    
    def test_simple_row(self):
        result = split_md_table_row("| A | B | C |")
        assert result == ["A", "B", "C"]
    
    def test_row_with_extra_pipes(self):
        result = split_md_table_row("|A|B|C|")
        assert result == ["A", "B", "C"]
    
    def test_row_with_spaces(self):
        result = split_md_table_row("| A | B C | D |")
        assert result == ["A", "B C", "D"]
    
    def test_row_missing_pipes(self):
        result = split_md_table_row("A | B | C")
        assert result == ["A", "B", "C"]
    
    def test_empty_cells(self):
        result = split_md_table_row("| | B | |")
        assert result == ["", "B", ""]


class TestIsMdTableSeparator:
    """Тесты для is_md_table_separator"""
    
    def test_valid_separator(self):
        assert is_md_table_separator("| --- | --- | --- |") == True
    
    def test_separator_with_colons(self):
        assert is_md_table_separator("| :--- | ---: | :---: |") == True
    
    def test_invalid_separator_too_few_dashes(self):
        assert is_md_table_separator("| -- | --- | --- |") == False
    
    def test_invalid_separator_no_pipes(self):
        assert is_md_table_separator("--- --- ---") == False
    
    def test_invalid_separator_with_text(self):
        assert is_md_table_separator("| --- | abc | --- |") == False


class TestIsMdTableRow:
    """Тесты для is_md_table_row"""
    
    def test_valid_table_row(self):
        assert is_md_table_row("| A | B |") == True
    
    def test_valid_table_row_no_leading_pipe(self):
        assert is_md_table_row("A | B |") == True
    
    def test_empty_string(self):
        assert is_md_table_row("") == False
    
    def test_no_pipe(self):
        assert is_md_table_row("just text") == False
    
    def test_only_whitespace(self):
        assert is_md_table_row("   ") == False


class TestNormalizeTableCaption:
    """Тесты для normalize_table_caption"""
    
    def test_html_format(self):
        result = normalize_table_caption("<caption>Таблица данных</caption>")
        assert result == "Таблица данных"
    
    def test_markdown_format(self):
        result = normalize_table_caption("Таблица 1 — Мои данные")
        assert result == "Мои данные"
    
    def test_plain_format_with_dash(self):
        result = normalize_table_caption("Название таблицы: Результаты")
        assert result == "Результаты"
    
    def test_no_caption(self):
        result = normalize_table_caption(None)
        assert result is None
    
    def test_empty_string(self):
        result = normalize_table_caption("")
        assert result is None
    
    def test_extra_spaces_normalized(self):
        result = normalize_table_caption("Таблица   1  —   Данные   с   пробелами")
        assert result == "Данные с пробелами"


class TestReplaceImageRefs:
    """Тесты для replace_image_refs"""
    
    def test_single_replacement(self):
        text = "Смотрите @img1 для примера"
        refs = {"img1": 1}
        result = replace_image_refs(text, refs)
        assert result == "Смотрите рис. 1 для примера"
    
    def test_multiple_replacements(self):
        text = "На @img1 видно, а на @img2 видно тоже"
        refs = {"img1": 1, "img2": 2}
        result = replace_image_refs(text, refs)
        assert result == "На рис. 1 видно, а на рис. 2 видно тоже"
    
    def test_no_replacement_needed(self):
        text = "Просто текст без ссылок"
        refs = {"img1": 1}
        result = replace_image_refs(text, refs)
        assert result == "Просто текст без ссылок"
    
    def test_empty_refs(self):
        text = "Смотрите @img1"
        result = replace_image_refs(text, {})
        assert result == "Смотрите @img1"
    
    def test_empty_text(self):
        refs = {"img1": 1}
        result = replace_image_refs(None, refs)
        assert result is None


if __name__ == "__main__":
    pytest.main([__file__, "-v"])

