#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Собирает системный промпт для нейросети: общие правила + промпт типа документа.

Использование:
    python tools/prompt.py курсовая               # вывести промпт
    python tools/prompt.py вкр -o промпт_вкр.md   # сохранить в файл
    python tools/prompt.py --list                 # доступные типы
"""

import argparse
import sys
from pathlib import Path

PROMPTS = Path(__file__).resolve().parent.parent / "prompts"
COMMON = PROMPTS / "00_общие_правила.md"
ALIASES = {
    "курсовая": "курсовая", "курсовой проект": "курсовая",
    "вкр": "вкр", "диплом": "вкр",
    "практика": "практика", "отчёт по практике": "практика", "отчет по практике": "практика",
    "реферат": "реферат", "доклад": "реферат",
    "отчёт": "отчёт", "отчет": "отчёт", "гост": "отчёт",
    "договор": "договор", "соглашение": "договор",
}


def main():
    for stream in (sys.stdout, sys.stderr):  # консоль Windows (cp1251)
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(description="Промпт для подготовки документа в MarkToGost")
    parser.add_argument("type", nargs="?", help="тип документа")
    parser.add_argument("-o", "--output", help="сохранить в файл")
    parser.add_argument("--list", action="store_true", help="показать типы")
    args = parser.parse_args()

    types = sorted(p.stem for p in PROMPTS.glob("*.md") if p != COMMON and p.stem != "README")
    if args.list or not args.type:
        print("Типы:", ", ".join(types))
        return 0

    name = ALIASES.get(args.type.strip().lower(), args.type.strip().lower())
    type_file = PROMPTS / f"{name}.md"
    if not type_file.exists():
        print(f"Нет промпта для типа «{args.type}». Типы: {', '.join(types)}", file=sys.stderr)
        return 1

    text = COMMON.read_text(encoding="utf-8").rstrip() + "\n\n---\n\n" + type_file.read_text(encoding="utf-8")
    if args.output:
        Path(args.output).write_text(text, encoding="utf-8")
        print(f"Сохранено: {args.output}")
    else:
        print(text)
    return 0


if __name__ == "__main__":
    sys.exit(main())
