"""
extractor.py — Lee Notas de Gemini (.docx) y extrae puntos y actividades.

Estructura del archivo:
  Heading 3 "Detalles"
  → párrafos 'normal' con texto "Título: descripción (timestamps)"
  Heading 3 / Title → fin de la sección
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
    text = re.sub(r'\(\d{2}:\d{2}:\d{2}\)', '', text)   # timestamps
    text = re.sub(r'\*\*([^*]+)\*\*', r'\1', text)       # **bold**
    text = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', text) # [link](url)
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def extract_development_points(doc) -> list[str]:
    """
    Recorre los párrafos del docx buscando el Heading 3 'Detalles',
    luego recopila todos los párrafos normales hasta el próximo heading/título.
    """
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

        # Detectar entrada a la sección Detalles
        if 'heading' in style and text.lower() == 'detalles':
            in_detalles = True
            continue

        # Salir de la sección al encontrar otro heading o sección nueva
        if in_detalles:
            if 'heading' in style or 'title' in style:
                break
            # Ignorar líneas de calificación/encuesta
            if any(p in text.lower() for p in skip_patterns):
                continue
            # Solo incluir si tiene contenido real (título + descripción)
            cleaned = clean_text(text)
            if len(cleaned) > 30 and ':' in cleaned[:80]:
                points.append(cleaned)

    return points


def smart_cut(text: str, max_len: int = 200) -> str:
    """Corta el texto en un punto natural (., ;, ,) sin exceder max_len."""
    if len(text) <= max_len:
        return text
    segment = text[:max_len]
    # Intentar cortar en punto, punto y coma o coma
    for sep in ['. ', '; ', ', ']:
        pos = segment.rfind(sep)
        if pos > 40:
            return segment[:pos + 1]
    # Fallback: cortar en espacio
    return segment.rsplit(' ', 1)[0]


def infer_activities(points: list[str]) -> list[tuple]:
    """
    Infiere estado R/P/NA de cada punto.
    Retorna lista de (nombre, R, P, NA, observacion).
    """
    pending_kw = [
        'pendiente', 'falta', 'queda', 'requiere', 'necesita', 'se debe',
        'hay que', 'por definir', 'sin respuesta', 'todavía', 'no se ha',
        'por resolver', 'se acordó que', 'se definirá', 'por confirmar',
        'se requiere', 'debe completar', 'debe quedar'
    ]
    done_kw = [
        'se acordó', 'se definió', 'se configuró', 'se creó', 'se resolvió',
        'se aclaró', 'se confirmó', 'quedó listo', 'se realizó', 'se completó',
        'se determinó', 'se estableció', 'se decidió', 'se mantuvo', 'se cargó'
    ]

    activities = []
    for point in points:
        lower = point.lower()

        # Título = texto antes del primer ':'
        colon = point.find(':')
        if 0 < colon < 80:
            title = point[:colon].strip()
            obs_text = point[colon+1:].strip()
            # Solo la primera oración, máximo 80 caracteres
            first_sentence = re.split(r'(?<=[.!?])\s+', obs_text)[0]
            obs = smart_cut(first_sentence)
        else:
            words = point.split()
            title = ' '.join(words[:7])
            obs = smart_cut(' '.join(words[7:]))

        is_done = any(kw in lower for kw in done_kw)
        is_pending = any(kw in lower for kw in pending_kw)

        if is_done and not is_pending:
            r, p, na = True, False, False
        else:
            r, p, na = False, True, False

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
