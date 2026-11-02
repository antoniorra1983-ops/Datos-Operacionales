"""Entrada Streamlit para Datos-Operacionales con mejoras de estabilidad.

Esta versión compacta reconstruye el código original de la revisión fijada y
aplica las correcciones en memoria. Conserva el dashboard original completo.
La fuente original se descarga una sola vez y se almacena localmente.

Para una versión final independiente de GitHub, usa crear_app_completa.py sobre
el app.py original; el resultado app_actualizada.py es completamente autónomo.
"""
from __future__ import annotations

import hashlib
from pathlib import Path
from urllib.request import Request, urlopen

from aplicar_mejoras import transformar

_BASE = Path(__file__).resolve().parent
_ORIGINAL_SHA = '7e6f94b31a67fc913fa9061bbafd0d7457e5c1a8'
_COMMIT = 'a198bdaf58a7ae8c7fadda756c8046cc69e2b19e'
_URL = (
    'https://raw.githubusercontent.com/'
    f'antoniorra1983-ops/Datos-Operacionales/{_COMMIT}/app.py'
)
_CACHE = _BASE / '.datos_operacionales_cache' / f'app_{_ORIGINAL_SHA}.py'


def _sha_git(contenido: bytes) -> str:
    cabecera = b'blob ' + str(len(contenido)).encode('ascii') + b'\0'
    return hashlib.sha1(cabecera + contenido).hexdigest()


def _fuente_original() -> str:
    """Obtiene la versión original exacta, verificando su SHA de Git."""
    for ruta in (_BASE / 'app_original.py', _CACHE):
        if ruta.is_file():
            contenido = ruta.read_bytes()
            if _sha_git(contenido) == _ORIGINAL_SHA:
                return contenido.decode('utf-8-sig')

    peticion = Request(_URL, headers={'User-Agent': 'Datos-Operacionales-Streamlit/1.0'})
    try:
        with urlopen(peticion, timeout=30) as respuesta:
            contenido = respuesta.read(2_000_000)
    except Exception as exc:
        raise RuntimeError(
            'No se pudo obtener la versión base verificada. '
            'Coloca una copia del app.py original como app_original.py '
            'o comprueba la conexión con GitHub.'
        ) from exc
    if _sha_git(contenido) != _ORIGINAL_SHA:
        raise RuntimeError('El código base descargado no coincide con el original verificado.')
    try:
        _CACHE.parent.mkdir(parents=True, exist_ok=True)
        _CACHE.write_bytes(contenido)
    except OSError:
        pass  # La app también puede ejecutarse si el disco es de solo lectura.
    return contenido.decode('utf-8-sig')


# Se ejecuta el programa Streamlit completo con los fixes ya incorporados.
_codigo = transformar(_fuente_original().replace('\r\n', '\n'))
_ruta_codigo = _BASE / '.datos_operacionales_cache' / 'app_completa_actualizada.py'
try:
    _ruta_codigo.parent.mkdir(parents=True, exist_ok=True)
    if not _ruta_codigo.is_file() or _ruta_codigo.read_text(encoding='utf-8') != _codigo:
        _ruta_codigo.write_text(_codigo, encoding='utf-8')
except OSError:
    pass  # Si no se puede escribir, se ejecuta igualmente en memoria.
exec(compile(_codigo, str(_ruta_codigo), 'exec'), globals(), globals())
