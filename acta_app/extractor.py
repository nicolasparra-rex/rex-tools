"""
extractor.py — Lee Notas de Gemini (.docx) y extrae puntos y actividades.
"""

import re
from docx import Document


def parse_filename_date(filename: str) -> str:
    match = re.search(r'(\d{4})[/_: ](\d{2})[/_: ](\d{2})', filename)
    if match:
        y, m, d = match.groups()
        return f"{d}-{m}-{y}"
    return ""


def clean_text(text: str) -> str:
    text = re.sub(r'\(\d{2}:\d{2}:\d{2}\)', '', text)
    text = re.sub(r'\*\*([^*]+)\*\*', r'\1', text)
    text = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def extract_development_points(doc) -> list[str]:
    points = []
    in_detalles = False

    skip_patterns = [
        'califica', 'revisa las notas', 'cómo es la calidad',
        'responde una', 'obtén sugerencias', 'útil', 'poco útil',
    ]

    for para in doc.paragraphs:
        style = para.style.name.lower()
        text = para.text.strip()

        if not text:
            continue

        if 'heading' in style and text.lower() == 'detalles':
            in_detalles = True
            continue

        if in_detalles:
            if 'heading' in style or 'title' in style:
                break
            if any(p in text.lower() for p in skip_patterns):
                continue
            cleaned = clean_text(text)
            if len(cleaned) > 30 and ':' in cleaned[:80]:
                points.append(cleaned)

    return points


def smart_cut(text: str, max_len: int = 200) -> str:
    if len(text) <= max_len:
        return text
    segment = text[:max_len]
    for sep in ['. ', '; ', ', ']:
        pos = segment.rfind(sep)
        if pos > 40:
            return segment[:pos + 1]
    return segment.rsplit(' ', 1)[0]


def infer_activities(points: list[str]) -> list[tuple]:
    """
    Infiere estado R/P/NA de cada punto.
    Retorna lista de (nombre, R, P, NA, observacion).
    Default: R (Realizado), salvo que haya keywords de pendiente explícitas.
    """
    pending_kw = [
        'pendiente', 'falta', 'queda', 'requiere', 'necesita', 'se debe',
        'hay que', 'por definir', 'sin respuesta', 'todavía', 'no se ha',
        'por resolver', 'se definirá', 'por confirmar',
        'se requiere', 'debe completar', 'debe quedar',
    ]

    activities = []
    for point in points:
        lower = point.lower()

        colon = point.find(':')
        if 0 < colon < 80:
            title = point[:colon].strip()
            obs_text = point[colon+1:].strip()
            first_sentence = re.split(r'(?<=[.!?])\s+', obs_text)[0]
            obs = smart_cut(first_sentence)
        else:
            words = point.split()
            title = ' '.join(words[:7])
            obs = smart_cut(' '.join(words[7:]))

        is_pending = any(kw in lower for kw in pending_kw)

        if is_pending:
            r, p, na = False, True, False
        else:
            r, p, na = True, False, False  # default: Realizado

        activities.append((title, r, p, na, obs))

    return activities


def extract_all(file, filename: str = "") -> dict:
    doc = Document(file)
    dev_points = extract_development_points(doc)
    activities = infer_activities(dev_points)
    fecha = parse_filename_date(filename)

    return {
        "fecha": fecha,
        "dev_points": dev_points,
        "activities": activities,
    }
