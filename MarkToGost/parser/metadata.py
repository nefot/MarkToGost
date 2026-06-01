import os
import re
from typing import Dict, Any


def extract_metadata(md_text: str) -> Dict[str, Any]:
    """Извлечение метаданных из текста"""
    metadata = {
        'is_table': False,
        'teacher': '',
        'fio': '',
        'group': '',
        'image_map': {}
    }

    # Извлечение флагов

    prefix = re.escape('[//]:')
    m_headings = re.search(prefix + r"\s*#\s*use_headings\s*=\s*(true|false)", md_text, re.I)
    if m_headings:
        metadata['use_headings'] = m_headings.group(1).lower() != 'false'
    else:
        metadata['use_headings'] = True
    m_numerate = re.search(prefix + r"\s*#\s*numerate\s*=\s*(true|false)", md_text, re.I)
    if m_numerate:
        metadata['numerate'] = m_numerate.group(1).lower() != 'false'
    else:
        metadata['numerate'] = True  # по умолчанию нумерация включена
    m_flag = re.search(prefix + r"\s*#\s*is_table\s*=\s*(true|false)", md_text, re.I)
    if m_flag:
        metadata['is_table'] = m_flag.group(1).lower() == 'true'

    m_teacher = re.search(prefix + r"\s*#\s*teacher\s*=\s*\"?([^\"\n]+)\"?", md_text, re.I)
    if m_teacher:
        metadata['teacher'] = m_teacher.group(1).strip()

    # Извлечение ФИО и группы
    fio_match = re.search(r'__?ФИО__?[:\-]?\s*(.*)', md_text)
    group_match = re.search(r'__?Группа__?[:\-]?\s*(.*)', md_text)

    metadata['fio'] = fio_match.group(1).strip() if fio_match else ""
    metadata['group'] = group_match.group(1).strip() if group_match else ""

    # Карта изображений
    img_pattern = re.compile(r'!\[([^]]*)\]\(([^)\s]+)(?:\s"([^\"]*)")?\)')
    for m in img_pattern.finditer(md_text):
        alt = m.group(1).strip()
        path = m.group(2).strip()
        title = (m.group(3) or "").strip()
        key = os.path.basename(path)
        caption = alt or title or ""
        metadata['image_map'][key] = caption

    return metadata
def convert_inline_math(text: str) -> str:
    """
    Конвертирует inline math $...$ в читаемый текст.
    Убирает знаки доллара, оставляя содержимое.

    Примеры:
        $F$ → F
        $m_1$ → m₁  (нижний индекс через юникод)
        $6.674 \\times 10^{-11}$ → 6.674×10⁻¹¹
        $E_k$ → Eₖ
    """
    if not text:
        return text

    # Словари для конвертации индексов в юникод
    SUBSCRIPT_MAP = str.maketrans("0123456789aeinoruvxhklmnpst", "₀₁₂₃₄₅₆₇₈₉ₐₑᵢₙₒᵣᵤᵥₓₕₖₗₘₙₚₛₜ")
    SUPERSCRIPT_MAP = str.maketrans("0123456789+-=()ni", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿⁱ")

    def process_math(m: str) -> str:
        """Обрабатывает содержимое между знаками $"""
        # Замены LaTeX команд
        m = re.sub(r'\s*\\times\s*', '×', m)
        m = m.replace(r'\cdot', '·')
        m = m.replace(r'\pm', '±')
        m = m.replace(r'\leq', '≤')
        m = m.replace(r'\geq', '≥')
        m = m.replace(r'\neq', '≠')
        m = m.replace(r'\alpha', 'α')
        m = m.replace(r'\beta', 'β')
        m = m.replace(r'\gamma', 'γ')
        m = m.replace(r'\delta', 'δ')
        m = m.replace(r'\Delta', 'Δ')
        m = m.replace(r'\sigma', 'σ')
        m = m.replace(r'\rho', 'ρ')
        m = m.replace(r'\mu', 'μ')
        m = m.replace(r'\lambda', 'λ')
        m = m.replace(r'\omega', 'ω')
        m = m.replace(r'\Omega', 'Ω')
        m = m.replace(r'\pi', 'π')
        m = m.replace(r'\infty', '∞')
        m = m.replace(r'\frac', '')

        # Нижние индексы _x или _{xxx}
        def replace_sub(match):
            content = match.group(1) or match.group(2)
            return content.translate(SUBSCRIPT_MAP)

        m = re.sub(r'_\{([^}]+)\}|_([^{}\s])', replace_sub, m)

        # Верхние индексы ^x или ^{xxx}
        def replace_sup(match):
            content = match.group(1) or match.group(2)
            return content.translate(SUPERSCRIPT_MAP)

        m = re.sub(r'\^\{([^}]+)\}|\^([^{}\s])', replace_sup, m)

        # Убираем оставшиеся фигурные скобки
        m = m.replace('{', '').replace('}', '')

        return m.strip()

    # Заменяем $...$
    result = re.sub(r'\$([^$]+)\$', lambda match: process_math(match.group(1)), text)
    return result