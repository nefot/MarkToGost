# utils/drawio.py

"""
Экспорт схем draw.io в PNG для вставки в документ.

В Markdown схема указывается как обычное изображение:
    ![Подпись](img/schemes.drawio)            — первая страница
    ![Подпись](img/schemes.drawio#2)          — страница по номеру (с 1)
    ![Подпись](img/schemes.drawio#Контекст)   — страница по названию вкладки

Экспорт выполняется через CLI draw.io Desktop и кэшируется: PNG пересоздаётся,
только если .drawio-файл изменился.
"""

import os
import shutil
import subprocess
import xml.etree.ElementTree as ET
from typing import Optional, Tuple

from MarkToGost.config import DrawioSettings

DRAWIO_EXTENSIONS = (".drawio", ".dio")


def is_drawio_path(path: str) -> bool:
    return split_drawio_path(path)[0].lower().endswith(DRAWIO_EXTENSIONS)


def split_drawio_path(path: str) -> Tuple[str, str]:
    """'img/a.drawio#Контекст' -> ('img/a.drawio', 'Контекст')"""
    file_part, _, page = path.partition("#")
    return file_part, page.strip()


def find_drawio_exe() -> Optional[str]:
    """Путь к draw.io: переменная окружения DRAWIO_PATH, настройки, PATH"""
    candidates = [os.environ.get("DRAWIO_PATH"), *DrawioSettings.EXE_CANDIDATES]
    for c in candidates:
        if c and os.path.isfile(c):
            return c
    return shutil.which("draw.io") or shutil.which("drawio")


def resolve_page_index(drawio_file: str, page: str) -> int:
    """Номер страницы (с 1) по номеру или названию вкладки"""
    if not page:
        return 1
    if page.isdigit():
        return int(page)

    root = ET.parse(drawio_file).getroot()
    names = [d.get("name", "") for d in root.iter("diagram")]
    for i, name in enumerate(names, start=1):
        if name.strip().lower() == page.lower():
            return i
    raise ValueError(f"страница «{page}» не найдена, есть: {', '.join(names)}")


def export_drawio(drawio_file: str, page: str = "") -> str:
    """
    Экспортирует страницу .drawio в PNG и возвращает путь к PNG.
    Бросает RuntimeError/ValueError с понятным текстом при ошибке.
    """
    exe = find_drawio_exe()
    if not exe:
        raise RuntimeError("draw.io Desktop не найден (укажите путь в DRAWIO_PATH)")

    page_index = resolve_page_index(drawio_file, page)

    os.makedirs(DrawioSettings.CACHE_DIR, exist_ok=True)
    stem = os.path.splitext(os.path.basename(drawio_file))[0]
    mtime = int(os.path.getmtime(drawio_file))
    out_png = os.path.join(DrawioSettings.CACHE_DIR, f"{stem}_p{page_index}_{mtime}.png")

    if os.path.exists(out_png):
        return out_png

    cmd = [
        exe, "--export", "--format", "png",
        "--scale", str(DrawioSettings.SCALE),
        "--border", str(DrawioSettings.BORDER),
        "--page-index", str(page_index),
        "--output", out_png,
        drawio_file,
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=DrawioSettings.TIMEOUT_SEC)
    if result.returncode != 0 or not os.path.exists(out_png):
        details = (result.stderr or result.stdout or "").strip()
        raise RuntimeError(f"экспорт не удался (код {result.returncode}) {details}")

    return out_png
