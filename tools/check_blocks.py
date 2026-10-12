#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Утилита для проверки отдельных блоков

Использование:
    python tools/check_blocks.py 01_TextBlock
    python tools/check_blocks.py all
    python tools/check_blocks.py --list
"""

import os
import sys
import argparse
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from MarkToGost.main import create_document  # noqa: E402

PROJECT_ROOT = Path(__file__).resolve().parent.parent
BLOCKS_DIR = PROJECT_ROOT / "examples" / "blocks"
OUTPUT_DIR = PROJECT_ROOT / "output" / "blocks"


def list_blocks():
    """Вывести список доступных блоков"""
    files = sorted(BLOCKS_DIR.glob("*.md"))
    files = [f for f in files if f.name != "README.md"]
    
    print("📚 Доступные блоки для проверки:\n")
    for i, f in enumerate(files, 1):
        name = f.stem
        print(f"{i}. {name}")
    
    return [f.stem for f in files]


def test_block(block_name: str, verbose: bool = True):
    """Проверить один блок"""
    block_file = BLOCKS_DIR / f"{block_name}.md"
    
    if not block_file.exists():
        print(f"❌ Файл не найден: {block_file}")
        return False
    
    # Создаём папку для вывода
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_file = OUTPUT_DIR / f"{block_name}.docx"
    
    try:
        # Читаем Markdown
        with open(block_file, "r", encoding="utf-8") as f:
            md_text = f.read()
        
        # Создаём документ
        doc = create_document(md_text)
        
        # Сохраняем
        doc.save(output_file)
        
        if verbose:
            size_kb = output_file.stat().st_size / 1024
            print(f"✅ {block_name}")
            print(f"   📄 {output_file}")
            print(f"   📊 Размер: {size_kb:.1f} KB")
        
        return True
    
    except Exception as e:
        print(f"❌ Ошибка при обработке {block_name}:")
        print(f"   {type(e).__name__}: {e}")
        return False


def main():
    parser = argparse.ArgumentParser(
        description="Проверить отдельные блоки MarkToGost"
    )
    parser.add_argument(
        "block",
        nargs="?",
        help="Имя блока (без расширения) или 'all' для всех"
    )
    parser.add_argument(
        "--list", "-l",
        action="store_true",
        help="Вывести список доступных блоков"
    )
    parser.add_argument(
        "--quiet", "-q",
        action="store_true",
        help="Тихий режим (без подробного вывода)"
    )
    
    args = parser.parse_args()
    
    # Список блоков
    if args.list or not args.block:
        blocks = list_blocks()
        if args.list:
            return 0
    else:
        blocks = None
    
    # Проверить один блок
    if args.block and args.block != "all":
        print(f"\n🔍 Проверка блока: {args.block}\n")
        success = test_block(args.block, verbose=not args.quiet)
        return 0 if success else 1
    
    # Проверить все блоки
    if args.block == "all":
        blocks = list_blocks()
        print("\n🔍 Проверка всех блоков:\n")
        
        results = []
        for block in blocks:
            success = test_block(block, verbose=not args.quiet)
            results.append((block, success))
            if not args.quiet:
                print()
        
        # Статистика
        passed = sum(1 for _, s in results if s)
        total = len(results)
        print(f"\n📊 Результаты: {passed}/{total} блоков ✅")
        
        if passed == total:
            print(f"🎉 Все блоки обработаны успешно!")
        else:
            print(f"⚠️  {total - passed} блок(ов) с ошибками")
            for block, success in results:
                if not success:
                    print(f"   ❌ {block}")
        
        return 0 if passed == total else 1
    
    return 0


if __name__ == "__main__":
    # Консоль Windows (cp1251) не умеет выводить эмодзи
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")
    sys.exit(main())


