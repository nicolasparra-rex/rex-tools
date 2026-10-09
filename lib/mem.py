"""
mem.py — Log del pico de memoria del proceso, sin dependencias nuevas.

Streamlit ejecuta cada script de pages/ de forma independiente: app.py no se
vuelve a correr al navegar a una pagina. Por eso el log vive aca y cada pagina
llama a log_mem() con su nombre, en vez de ponerlo solo en app.py.

Notas de portabilidad (todo stdlib):
  - `resource` solo existe en Unix. En Windows no se loguea nada: no hay
    equivalente en la stdlib y la idea es no agregar dependencias (psutil).
  - `ru_maxrss` viene en KB en Linux y en bytes en macOS. Se normaliza para
    que la cifra en MB sea correcta en ambos.
  - Es el pico (high-water mark) del proceso, no el uso actual: nunca baja.
    Sirve para ver crecimiento entre reruns, no para medir liberacion.
"""

import logging
import sys

try:
    import resource
except ImportError:  # Windows
    resource = None

logging.basicConfig(level=logging.INFO)

# macOS reporta ru_maxrss en bytes; Linux en kilobytes.
_DIVISOR = 1024 * 1024 if sys.platform == "darwin" else 1024


def pico_rss_mb():
    """Pico de RSS del proceso en MB, o None si la plataforma no lo expone."""
    if resource is None:
        return None
    return resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / _DIVISOR


def log_mem(pagina):
    """Loguea el pico de RSS a stderr via logging. Nunca lanza."""
    try:
        rss = pico_rss_mb()
        if rss is not None:
            logging.info(f"[mem] {pagina} · pico RSS: {rss:.0f} MB")
    except Exception:  # el log jamas debe tumbar la pagina
        pass
