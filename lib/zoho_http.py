"""
Helpers HTTP para la API de Zoho Projects.

La API responde a veces con HTTP 204 (o 200 con cuerpo vacío) cuando no hay
resultados, lo que reventaba en r.json() con JSONDecodeError. Estas funciones
normalizan ese caso a "sin datos" y devuelven el error como valor, para que la
página decida si lo muestra. No llaman a st.error() por dentro a propósito:
casi todas las llamadas viven dentro de funciones @st.cache_data, y un
st.error() ahí queda congelado en la caché y no se repinta en los reruns.
"""

# Cuántos caracteres del cuerpo se incluyen en el mensaje de error.
_EXTRACTO = 300


def zoho_json(r, contexto=""):
    """Parsea la respuesta de Zoho.

    Devuelve (datos, error):
      - (dict, None)  respuesta válida
      - (None, None)  204 o cuerpo vacío -> sin resultados, no es error
      - (None, str)   status de error o JSON inválido
    """
    etiqueta = f" ({contexto})" if contexto else ""

    # El status manda: un error de servidor con cuerpo vacío es un error,
    # no "sin resultados". Por eso se evalúa antes del 204.
    if not r.ok:
        cuerpo = r.text[:_EXTRACTO] or "(cuerpo vacío)"
        return None, f"Zoho respondió HTTP {r.status_code}{etiqueta}: {cuerpo}"

    if r.status_code == 204 or not (r.content or b"").strip():
        return None, None

    try:
        return r.json(), None
    except ValueError:
        return None, (f"Respuesta no-JSON de Zoho{etiqueta} "
                      f"(HTTP {r.status_code}): {r.text[:_EXTRACTO]}")


def zoho_lista(r, clave, contexto=""):
    """Igual que zoho_json pero para endpoints de listado.

    Devuelve (lista, error); la lista siempre es una lista, nunca None.
    """
    datos, error = zoho_json(r, contexto)
    return (datos or {}).get(clave, []) or [], error
