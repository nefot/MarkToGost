# tests/test_metadata.py

from MarkToGost.parser.metadata import extract_metadata


class TestExtractMetadata:

    def test_defaults(self):
        meta = extract_metadata("Просто текст")
        assert meta["use_headings"] is True
        assert meta["numerate"] is True
        assert meta["is_table"] is False
        assert meta["fio"] == ""

    def test_flags(self):
        md = "[//]: # numerate = false\n[//]: # use_headings = false\n[//]: # is_table = true"
        meta = extract_metadata(md)
        assert meta["numerate"] is False
        assert meta["use_headings"] is False
        assert meta["is_table"] is True

    def test_teacher(self):
        meta = extract_metadata('[//]: # teacher = "Черноморова Т. С."')
        assert meta["teacher"] == "Черноморова Т. С."

    def test_fio_and_group(self):
        meta = extract_metadata("__ФИО__: Иванов И. И.\n__Группа__: 3а")
        assert meta["fio"] == "Иванов И. И."
        assert meta["group"] == "3а"
