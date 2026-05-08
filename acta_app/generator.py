"""
generator.py
Genera el acta de implementación .docx a partir de:
- Un template.docx con tokens %%TOKEN%%
- Datos del formulario (header, asistentes)
- Puntos de desarrollo extraídos de Gemini
- Actividades con estado R/P/NA
"""

import zipfile
import io
import re
import random
from pathlib import Path

TEMPLATE_PATH = Path(__file__).parent / "template.docx"

# ── Helpers XML ──────────────────────────────────────────────────────────────

def esc(text: str) -> str:
    return (str(text)
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;'))


def uid() -> int:
    return random.randint(-999999999, 999999999)


def checkbox_cell(width: int, checked: bool = False) -> str:
    val = "1" if checked else "0"
    char = "&#x2612;" if checked else "&#x2610;"
    return (
        f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>'
        f'<w:tcBorders>'
        f'<w:top w:val="single" w:sz="4" w:space="0" w:color="B4C6E7" w:themeColor="accent1" w:themeTint="66"/>'
        f'<w:left w:val="single" w:sz="4" w:space="0" w:color="B4C6E7" w:themeColor="accent1" w:themeTint="66"/>'
        f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="B4C6E7" w:themeColor="accent1" w:themeTint="66"/>'
        f'<w:right w:val="single" w:sz="4" w:space="0" w:color="B4C6E7" w:themeColor="accent1" w:themeTint="66"/>'
        f'</w:tcBorders><w:noWrap/></w:tcPr>'
        f'<w:p><w:pPr><w:rPr>'
        f'<w:rFonts w:cstheme="minorHAnsi"/><w:bCs/><w:i/>'
        f'<w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/>'
        f'<w:sz w:val="20"/><w:szCs w:val="20"/>'
        f'</w:rPr></w:pPr>'
        f'<w:sdt><w:sdtPr>'
        f'<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
        f'<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
        f'<w:id w:val="{uid()}"/>'
        f'<w14:checkbox><w14:checked w14:val="{val}"/>'
        f'<w14:checkedState w14:val="2612" w14:font="MS Gothic"/>'
        f'<w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/>'
        f'</w14:checkbox></w:sdtPr>'
        f'<w:sdtContent><w:r>'
        f'<w:rPr><w:rFonts w:ascii="Segoe UI Symbol" w:hAnsi="Segoe UI Symbol" w:cs="Segoe UI Symbol"/>'
        f'<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
        f'<w:t>{char}</w:t>'
        f'</w:r></w:sdtContent></w:sdt></w:p></w:tc>'
    )


def obs_cell(text: str = "") -> str:
    inner = (
        f'<w:r><w:rPr>'
        f'<w:rFonts w:cstheme="minorHAnsi"/><w:bCs/>'
        f'<w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/>'
        f'<w:sz w:val="20"/><w:szCs w:val="20"/>'
        f'</w:rPr><w:t xml:space="preserve">{esc(text)}</w:t></w:r>'
    ) if text else ""
    return (
        f'<w:tc><w:tcPr><w:tcW w:w="4844" w:type="dxa"/>'
        f'<w:tcBorders>'
        f'<w:top w:val="single" w:sz="4" w:space="0" w:color="B4C6E7" w:themeColor="accent1" w:themeTint="66"/>'
        f'<w:left w:val="single" w:sz="4" w:space="0" w:color="B4C6E7" w:themeColor="accent1" w:themeTint="66"/>'
        f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="B4C6E7" w:themeColor="accent1" w:themeTint="66"/>'
        f'<w:right w:val="single" w:sz="4" w:space="0" w:color="B4C6E7" w:themeColor="accent1" w:themeTint="66"/>'
        f'</w:tcBorders><w:noWrap/></w:tcPr>'
        f'<w:p><w:pPr><w:rPr>'
        f'<w:rFonts w:cstheme="minorHAnsi"/><w:bCs/>'
        f'<w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/>'
        f'<w:sz w:val="20"/><w:szCs w:val="20"/>'
        f'</w:rPr></w:pPr>{inner}</w:p></w:tc>'
    )


def activity_row(name: str, r=False, p=False, na=False, obs="") -> str:
    name_cell = (
        f'<w:tc><w:tcPr><w:tcW w:w="3929" w:type="dxa"/>'
        f'<w:tcBorders>'
        f'<w:top w:val="single" w:sz="4" w:space="0" w:color="B4C6E7" w:themeColor="accent1" w:themeTint="66"/>'
        f'<w:left w:val="single" w:sz="4" w:space="0" w:color="B4C6E7" w:themeColor="accent1" w:themeTint="66"/>'
        f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="B4C6E7" w:themeColor="accent1" w:themeTint="66"/>'
        f'<w:right w:val="single" w:sz="4" w:space="0" w:color="B4C6E7" w:themeColor="accent1" w:themeTint="66"/>'
        f'</w:tcBorders><w:noWrap/></w:tcPr>'
        f'<w:p><w:pPr><w:jc w:val="left"/><w:rPr>'
        f'<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/>'
        f'<w:i w:val="0"/><w:iCs w:val="0"/>'
        f'<w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/>'
        f'<w:sz w:val="18"/><w:szCs w:val="18"/>'
        f'</w:rPr></w:pPr>'
        f'<w:r><w:rPr>'
        f'<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/>'
        f'<w:i w:val="0"/><w:iCs w:val="0"/>'
        f'<w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/>'
        f'<w:sz w:val="18"/><w:szCs w:val="18"/>'
        f'</w:rPr><w:t xml:space="preserve">{esc(name)}</w:t></w:r></w:p></w:tc>'
    )
    return (
        f'<w:tr><w:trPr><w:trHeight w:val="290"/></w:trPr>'
        f'{name_cell}'
        f'{checkbox_cell(389, r)}'
        f'{checkbox_cell(444, p)}'
        f'{checkbox_cell(469, na)}'
        f'{obs_cell(obs)}'
        f'</w:tr>'
    )


def bullet_para(text: str, para_id: str) -> str:
    return (
        f'<w:p w14:paraId="{para_id}" w14:textId="77777777" '
        f'w:rsidR="00A4228A" w:rsidRDefault="00A4228A" w:rsidP="00A4228A">'
        f'<w:pPr>'
        f'<w:pStyle w:val="Prrafodelista"/>'
        f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>'
        f'<w:jc w:val="both"/>'
        f'<w:spacing w:before="160" w:after="160"/>'
        f'<w:rPr><w:rFonts w:cstheme="minorHAnsi"/><w:bCs/>'
        f'<w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/>'
        f'<w:sz w:val="20"/><w:szCs w:val="20"/>'
        f'<w:lang w:eastAsia="es-CL"/></w:rPr>'
        f'</w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:cstheme="minorHAnsi"/><w:bCs/>'
        f'<w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/>'
        f'<w:sz w:val="20"/><w:szCs w:val="20"/>'
        f'<w:lang w:eastAsia="es-CL"/></w:rPr>'
        f'<w:t xml:space="preserve">{esc(text)}</w:t></w:r>'
        f'</w:p>'
    )


def asistente_row(nombre: str, cargo: str = "", gerencia: str = "", num: int = 1) -> str:
    """Genera una fila de la tabla de asistentes."""
    rpr = (
        '<w:rFonts w:ascii="Calibri" w:eastAsia="Times New Roman" w:hAnsi="Calibri" w:cs="Calibri"/>'
        '<w:color w:val="003287"/><w:sz w:val="18"/><w:szCs w:val="18"/>'
        '<w:lang w:eastAsia="es-CL"/>'
    )
    def tc(width, text, bold=False):
        b = '<w:b/><w:bCs/>' if bold else ''
        return (
            f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/></w:tcPr>'
            f'<w:p><w:pPr><w:rPr>{rpr}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{b}{rpr}</w:rPr>'
            f'<w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p></w:tc>'
        )
    return (
        f'<w:tr><w:trPr><w:trHeight w:val="309"/></w:trPr>'
        f'{tc(1938, str(num), bold=True)}'
        f'{tc(3099, nombre)}'
        f'{tc(2334, cargo)}'
        f'{tc(2802, gerencia)}'
        f'</w:tr>'
    )


# ── Función principal ────────────────────────────────────────────────────────

def generate_acta(
    header: dict,
    asistentes: list[dict],
    dev_points: list[str],
    activities: list[tuple],
) -> bytes:
    """
    Genera el acta .docx y retorna los bytes.

    header: {empresa, plan, acta_num, fecha, jefe_cliente, usuario_impl,
             email_impl, jefe_rex, email_jefe_rex, consultor}
    asistentes: [{"nombre": ..., "cargo": ..., "gerencia": ...}]
    dev_points: lista de strings (bullets del desarrollo)
    activities: lista de (nombre, R, P, NA, observacion)
    """

    with open(TEMPLATE_PATH, 'rb') as f:
        template_bytes = f.read()

    with zipfile.ZipFile(io.BytesIO(template_bytes), 'r') as zin:
        xml = zin.read('word/document.xml').decode('utf-8')
        other_files = {
            name: zin.read(name)
            for name in zin.namelist()
            if name != 'word/document.xml'
        }

    # ── 1. Token replacements simples ───────────────────────────────────────
    tokens = {
        '%%EMPRESA%%':         header.get('empresa', ''),
        '%%PLAN%%':            header.get('plan', 'BASE'),
        '%%ACTA_NUM%%':        header.get('acta_num', header.get('fecha', '')),
        '%%FECHA%%':           header.get('fecha', ''),
        '%%JEFE_CLIENTE%%':    header.get('jefe_cliente', ''),
        '%%USUARIO_IMPL%%':    header.get('usuario_impl', ''),
        '%%EMAIL_IMPL%%':      header.get('email_impl', ''),
        '%%JEFE_REX%%':        header.get('jefe_rex', ''),
        '%%EMAIL_JEFE_REX%%':  header.get('email_jefe_rex', ''),
        '%%CONSULTOR%%':       header.get('consultor', ''),
    }
    for token, value in tokens.items():
        xml = xml.replace(token, esc(value))

    # ── 2. Asistentes — reemplazar todas las filas dinámicamente ─────────────
    # Encontrar el bloque completo de filas de asistentes (fila 1 hasta fila 4)
    # y reemplazarlo por tantas filas como haya en la lista
    import re as _re

    # Buscar la primera fila de datos (contiene %%ASISTENTE_1_NOMBRE%%)
    pos_start = xml.find('%%ASISTENTE_1_NOMBRE%%')
    # Buscar la última fila fija (contiene <w:t>4</w:t>)
    pos_end_marker = xml.find('<w:t>4</w:t>')

    if pos_start != -1 and pos_end_marker != -1:
        # Retroceder hasta el <w:tr > de la primera fila
        start_tr = xml.rfind('<w:tr', 0, pos_start)
        while start_tr != -1 and xml[start_tr+5:start_tr+6] not in (' ', '\n', '\r'):
            start_tr = xml.rfind('<w:tr', 0, start_tr)
        # Avanzar hasta el </w:tr> de la última fila (fila 4)
        end_tr = xml.find('</w:tr>', pos_end_marker) + len('</w:tr>')

        # Generar solo las filas necesarias
        new_rows = '\n'.join([
            asistente_row(
                a.get('nombre',''), a.get('cargo',''), a.get('gerencia',''), i+1
            )
            for i, a in enumerate(asistentes)
        ]) if asistentes else asistente_row('', '', '', 1)

        xml = xml[:start_tr] + new_rows + xml[end_tr:]
    else:
        # Fallback: reemplazar tokens simples
        xml = xml.replace('%%ASISTENTE_1_NOMBRE%%',
            esc(asistentes[0]['nombre']) if asistentes else '')
        xml = xml.replace('%%ASISTENTE_2_NOMBRE%%',
            esc(asistentes[1]['nombre']) if len(asistentes) > 1 else '')

    # ── 3. Desarrollo de la sesión ───────────────────────────────────────────
    if dev_points:
        para_ids = [f'2E{i:06X}' for i in range(len(dev_points))]
        first_id = '0190D4DB'

        # Primer párrafo (reemplaza el placeholder)
        first_bullet = (
            f'<w:p w14:paraId="{first_id}" w14:textId="6640E4BA" '
            f'w:rsidR="009356C6" w:rsidRPr="00A4228A" '
            f'w:rsidRDefault="00A4228A" w:rsidP="00A4228A">'
            f'<w:pPr>'
            f'<w:pStyle w:val="Prrafodelista"/>'
            f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>'
            f'<w:jc w:val="both"/>'
            f'<w:rPr><w:rFonts w:cstheme="minorHAnsi"/><w:bCs/>'
            f'<w:i w:val="0"/><w:iCs w:val="0"/>'
            f'<w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/>'
            f'<w:sz w:val="20"/><w:szCs w:val="20"/>'
            f'<w:lang w:eastAsia="es-CL"/></w:rPr>'
            f'</w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:cstheme="minorHAnsi"/><w:bCs/>'
            f'<w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/>'
            f'<w:sz w:val="20"/><w:szCs w:val="20"/>'
            f'<w:lang w:eastAsia="es-CL"/></w:rPr>'
            f'<w:t xml:space="preserve">{esc(dev_points[0])} </w:t></w:r>'
            f'</w:p>'
        )

        rest_bullets = '\n'.join([
            bullet_para(p, para_ids[i])
            for i, p in enumerate(dev_points[1:])
        ])

        old_placeholder_para = re.search(
            r'<w:p [^>]*paraId="0190D4DB"[^>]*>.*?</w:p>',
            xml, re.DOTALL
        )
        if old_placeholder_para:
            replacement = first_bullet
            if rest_bullets:
                replacement += '\n' + rest_bullets
            xml = xml[:old_placeholder_para.start()] + replacement + xml[old_placeholder_para.end():]

    # ── 4. Actividades ───────────────────────────────────────────────────────
    if activities:
        new_rows = '\n'.join([activity_row(*a) for a in activities])

        start_marker = 'Creaci\u00f3n 1 formato de contrato'
        end_marker = 'Pruebas funcionales de emisi\u00f3n de documentos'

        start_pos = xml.find(start_marker)
        end_pos = xml.find(end_marker)

        if start_pos != -1 and end_pos != -1:
            # Buscar <w:tr > (no <w:trPr>) antes del marcador de inicio
            search_from = start_pos
            start_tr = None
            while search_from > 0:
                pos = xml.rfind('<w:tr', 0, search_from)
                if pos == -1:
                    break
                if xml[pos+5:pos+6] in (' ', '\n', '\r'):
                    start_tr = pos
                    break
                search_from = pos

            end_tr = xml.find('</w:tr>', end_pos) + len('</w:tr>')

            if start_tr is not None:
                xml = xml[:start_tr] + new_rows + xml[end_tr:]

    # ── 5. Armar el .docx final ──────────────────────────────────────────────
    output = io.BytesIO()
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zout:
        zout.writestr('word/document.xml', xml.encode('utf-8'))
        for name, data in other_files.items():
            zout.writestr(name, data)

    output.seek(0)
    return output.read()
