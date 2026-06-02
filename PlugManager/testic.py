# -*- coding: utf-8 -*-
"""
db_to_md.py: Генерация pricelist2.md из таблицы Plug в PostgreSQL

Использование:
    python db_to_md.py
    python db_to_md.py --host localhost --port 5432 --db mydb --user postgres --password secret
    python db_to_md.py --output my_pricelist.md
"""

import argparse
import math
import os
from typing import List, Tuple

import psycopg2

# --- Настройки подключения по умолчанию ---
DEFAULT_HOST = "localhost"
DEFAULT_PORT = 5432
DEFAULT_DB   = "postgres"
DEFAULT_USER = "postgres"
DEFAULT_PASS = ""

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "input")
os.makedirs(OUTPUT_DIR, exist_ok=True)


# ---------------------------------------------------------------------------
# Категории и их маппинг из БД
# ---------------------------------------------------------------------------

# Точные значения поля "категория" в БД
CAT_SQUARE        = "Квадратная"
CAT_SQUARE_GREY   = "Квадратная (серые)"
CAT_RECT          = "Прямоугольная"
CAT_ROUND_INNER   = "Круглая внутренняя"
CAT_ROUND_OUTER   = "Круглая наружная"


def format_price(price: float) -> str:
    """Форматирует цену: 5.0 → '5,00 ₽'"""
    return f"{price:,.2f}".replace(",", " ").replace(".", ",") + " ₽"


def format_cell(size: str, price: float) -> str:
    """Форматирует ячейку таблицы: 'Ø18=2,20 ₽'"""
    return f"{size}={format_price(price)}"


def get_sort_key_square(size: str) -> float:
    """
    Сортировка квадратных/прямоугольных по площади.
    Формат: '10x10', '15x30', '120×120' и т.д.
    """
    size = size.replace("×", "x").replace("X", "x")
    parts = size.split("x")
    try:
        a, b = float(parts[0]), float(parts[1])
        return a * b
    except Exception:
        return 0.0


def get_sort_key_round(size: str) -> float:
    """
    Сортировка круглых по диаметру.
    Формат: 'Ø18', 'Ø102' и т.д.
    """
    try:
        return float(size.replace("Ø", "").replace("ø", "").strip())
    except Exception:
        return 0.0


def split_into_columns(items: List[Tuple[str, float]], n_cols: int) -> List[List[str]]:
    """
    Распределяет items равномерно по n_cols столбцам (по строкам сверху вниз).
    Возвращает список столбцов, каждый — список строк ячеек.
    """
    total = len(items)
    rows_per_col = math.ceil(total / n_cols)
    cols = []
    for i in range(n_cols):
        chunk = items[i * rows_per_col: (i + 1) * rows_per_col]
        cols.append([format_cell(s, p) for s, p in chunk])
    return cols


def fetch_data(conn) -> dict:
    """Загружает данные из таблицы Plug и возвращает словарь по категориям."""
    cur = conn.cursor()
    cur.execute('SELECT "категория", "размер", "цена_руб" FROM "Plug"')
    rows = cur.fetchall()
    cur.close()

    data = {
        CAT_SQUARE:      [],
        CAT_SQUARE_GREY: [],
        CAT_RECT:        [],
        CAT_ROUND_INNER: [],
        CAT_ROUND_OUTER: [],
    }

    for cat, size, price in rows:
        cat = cat.strip()
        if cat in data:
            data[cat].append((size.strip(), float(price)))

    # Сортировка
    for cat in [CAT_SQUARE, CAT_SQUARE_GREY, CAT_RECT]:
        data[cat].sort(key=lambda x: get_sort_key_square(x[0]))
    for cat in [CAT_ROUND_INNER, CAT_ROUND_OUTER]:
        data[cat].sort(key=lambda x: get_sort_key_round(x[0]))

    return data


def build_md(data: dict) -> str:
    """Собирает итоговый MD из данных."""
    lines = []

    lines.append("_Прайс-лист заглушки от 100 шт._\n")

    # -----------------------------------------------------------------------
    # Таблица 1: Квадратная | Серая | Прямоугольная
    # -----------------------------------------------------------------------
    sq   = [format_cell(s, p) for s, p in data[CAT_SQUARE]]
    grey = [format_cell(s, p) for s, p in data[CAT_SQUARE_GREY]]
    rect = [format_cell(s, p) for s, p in data[CAT_RECT]]

    max_rows_1 = max(len(sq), len(grey), len(rect))

    lines.append('<table class="transparent">')
    lines.append("    <tr>")
    lines.append('        <td colspan="1" bold>Квадратная</td>')
    lines.append("        <td bold>Серая</td>")
    lines.append("        <td bold>Прямоугольная</td>")
    lines.append("    </tr>")

    for i in range(max_rows_1):
        c1 = sq[i]   if i < len(sq)   else ""
        c2 = grey[i] if i < len(grey) else ""
        c3 = rect[i] if i < len(rect) else ""
        lines.append("    <tr>")
        lines.append(f'        <td colspan="1" >{c1}</td>')
        lines.append(f"        <td>{c2}</td>")
        lines.append(f"        <td>{c3}</td>")
        lines.append("    </tr>")

    lines.append("</table>\n")

    # -----------------------------------------------------------------------
    # Таблица 2: Круглые (внутренние 3 столбца + наружные 1 столбец)
    # -----------------------------------------------------------------------
    inner_cols = split_into_columns(data[CAT_ROUND_INNER], 3)
    outer_col  = [format_cell(s, p) for s, p in data[CAT_ROUND_OUTER]]

    # Выравниваем все 3 столбца внутренних до одной длины
    inner_max = max((len(c) for c in inner_cols), default=0)
    for col in inner_cols:
        while len(col) < inner_max:
            col.append("")

    max_rows_2 = max(inner_max, len(outer_col))

    lines.append('<table class="transparent">')
    lines.append("    <tr>")
    lines.append('        <td colspan="4" align="center" bold>Круглые</td>')
    lines.append("    </tr>")
    lines.append("    <tr>")
    lines.append('        <td colspan="3" bold>Внутренние</td>')
    lines.append("        <td bold>Наружные</td>")
    lines.append("    </tr>")

    for i in range(max_rows_2):
        c1 = inner_cols[0][i] if i < len(inner_cols[0]) else ""
        c2 = inner_cols[1][i] if i < len(inner_cols[1]) else ""
        c3 = inner_cols[2][i] if i < len(inner_cols[2]) else ""
        c4 = outer_col[i]     if i < len(outer_col)     else ""
        lines.append("    <tr>")
        lines.append(f"        <td>{c1}</td>")
        lines.append(f"        <td>{c2}</td>")
        lines.append(f"        <td>{c3}</td>")
        lines.append(f"        <td>{c4}</td>")
        lines.append("    </tr>")

    lines.append("</table>\n")

    lines.append("\n")
    lines.append("_ПРИ ЗАКАЗЕ НА СУММУ 5000₽ СКИДКА 5%_\n")
    lines.append("_ПРИ ЗАКАЗЕ ОТ 10000 И ВЫШЕ СКИДКА 10%_\n")

    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(
        description="db_to_md: Генерация прайс-листа MD из PostgreSQL (таблица Plug)"
    )
    parser.add_argument("--host",     default=DEFAULT_HOST, help=f"Хост БД (по умол.: {DEFAULT_HOST})")
    parser.add_argument("--port",     default=DEFAULT_PORT, type=int, help=f"Порт БД (по умол.: {DEFAULT_PORT})")
    parser.add_argument("--db",       default=DEFAULT_DB,   help=f"Имя БД (по умол.: {DEFAULT_DB})")
    parser.add_argument("--user",     default=DEFAULT_USER, help=f"Пользователь (по умол.: {DEFAULT_USER})")
    parser.add_argument("--password", default=DEFAULT_PASS, help="Пароль")
    parser.add_argument("--output",   default="pricelist2.md", help="Имя выходного MD файла (сохраняется в input/)")
    args = parser.parse_args()

    print(f"🔌 Подключение к {args.user}@{args.host}:{args.port}/{args.db} ...")
    try:
        conn = psycopg2.connect(
            host=args.host,
            port=args.port,
            dbname=args.db,
            user=args.user,
            password=args.password,
        )
    except Exception as e:
        print(f"❌ Ошибка подключения: {e}")
        return

    print("📦 Загрузка данных...")
    try:
        data = fetch_data(conn)
    except Exception as e:
        print(f"❌ Ошибка при чтении данных: {e}")
        conn.close()
        return
    conn.close()

    total = sum(len(v) for v in data.values())
    print(f"✅ Загружено записей: {total}")
    for cat, items in data.items():
        print(f"   {cat}: {len(items)} шт.")

    md = build_md(data)

    out_path = os.path.join(OUTPUT_DIR, args.output)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(md)

    print(f"\n✅ Файл сохранён: {out_path}")


if __name__ == "__main__":
    main()