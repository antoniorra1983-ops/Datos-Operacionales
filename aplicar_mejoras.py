"""Actualiza app.py del repositorio Datos-Operacionales sin rehacer el dashboard.

Uso desde la raíz del repositorio (junto a app.py y gestion_archivos.py):
    python aplicar_mejoras.py
    python -m unittest discover -s tests -v

Es idempotente: si ya se aplicó, no vuelve a escribir el archivo.
Se valida la sintaxis completa de app.py antes de guardar.
"""
from __future__ import annotations

from pathlib import Path
import re
import shutil


UPLOADS = (
    'f_v1', 'f_v2', 'f_umr', 'f_seat_files', 'f_bill_files',
    'f_carga_v1', 'f_carga_v2', 'f_viajes', 'f_ide', 'f_incid', 'f_oit',
)
FUENTES = (
    ('f_v1_all', 'v1'), ('f_v2_all', 'v2'), ('f_umr_all', 'umr'),
    ('f_seat_all', 'seat'), ('f_bill_all', 'bill'),
    ('f_carga_v1_all', 'carga_v1'), ('f_carga_v2_all', 'carga_v2'),
    ('f_viajes_all', 'viajes'), ('f_ide_all', 'ide_lb'),
    ('f_incid_all', 'incid'), ('f_oit_all', 'oit'),
)


def reemplazar_unico(texto: str, anterior: str, nuevo: str, descripcion: str) -> str:
    cuenta = texto.count(anterior)
    if cuenta != 1:
        raise RuntimeError(f'{descripcion}: esperaba una coincidencia, encontré {cuenta}. No se modificó app.py.')
    return texto.replace(anterior, nuevo, 1)


def transformar(texto: str) -> str:
    # 1. Sustituir funciones antiguas por el módulo de IO comprobable por separado.
    inicio = 'def guardar_archivo(uf, carpeta):\n'
    fin = 'def combinar_fuentes(ul, carpeta):\n'
    if texto.count(inicio) != 1 or texto.count(fin) != 1:
        raise RuntimeError('No se identificó un único bloque de almacenamiento original.')
    p, q = texto.index(inicio), texto.index(fin)
    if p > q:
        raise RuntimeError('Orden de funciones de almacenamiento inesperado.')
    texto = (texto[:p] +
             'from gestion_archivos import (\n'
             '    guardar_archivo, listar_archivos,\n'
             '    ArchivoEnDisco as _ArchivoEnDisco, firma_fuentes,\n'
             ')\n\n' + texto[q:])

    # 2. Al borrar un archivo, resetear el uploader activo en el siguiente rerun.
    marcador = '    def _badge(c): n=len(listar_archivos(c)); return f" ({n} guardados)" if n else ""'
    texto = reemplazar_unico(
        texto, marcador,
        '    _upload_revision = st.session_state.get("_upload_revision", 0)\n' + marcador,
        'Reinicio de los cargadores',
    )
    for variable in UPLOADS:
        patron = re.compile(
            r'(?m)^(\s*' + re.escape(variable) +
            r'\s*=\s*st\.file_uploader\([^\n]*?accept_multiple_files=True)\)\s*$'
        )
        texto, n = patron.subn(
            lambda m: m.group(1) + ', type=["xls", "xlsx", "xlsm", "csv"], '
                      + 'key=f"upload_' + variable + '_{_upload_revision}")',
            texto,
        )
        if n != 1:
            raise RuntimeError(f'Cargador {variable}: esperaba una coincidencia, encontré {n}.')

    # 3. Reemplazar contenidos con el mismo nombre; evita datos viejos y rutas inseguras.
    texto = reemplazar_unico(
        texto,
        '            dest=os.path.join(_ca,uf.name)\n'
        '            if not os.path.exists(dest): guardar_archivo(uf,_ca)',
        '            try:\n'
        '                guardar_archivo(uf, _ca)\n'
        '            except (ValueError, OSError) as _error_archivo:\n'
        '                st.error(f"No se pudo guardar {uf.name}: {_error_archivo}")',
        'Actualización segura de archivos subidos',
    )
    texto = reemplazar_unico(
        texto,
        '                    if cb2.button("🗑️",key=f"del_{_a}"): os.remove(_a); st.rerun()',
        '                    if cb2.button("🗑️", key=f"del_{_a}"):\n'
        '                        try:\n'
        '                            os.remove(_a)\n'
        '                        except FileNotFoundError:\n'
        '                            pass\n'
        '                        st.session_state["_upload_revision"] = _upload_revision + 1\n'
        '                        st.rerun()',
        'Eliminación que no restaura archivos desde el uploader',
    )

    # 4. Detectar reemplazos de archivos aun con nombre y tamaño idénticos.
    patron_cache = re.compile(
        r'_CACHE_VERSION = "v21_oit_sfe"\n'
        r'_cache_key = \(_CACHE_VERSION, str\(start_date\), str\(end_date\),\n'
        r'.*?\)\n[ \t]*\n(?=_hay_archivos =)', re.S,
    )
    contenido_cache = ('_CACHE_VERSION = "v22_integridad_archivos"\n'
                       '_cache_key = (_CACHE_VERSION, str(start_date), str(end_date),\n')
    for variable, clave in FUENTES:
        contenido_cache += f'              firma_fuentes({variable}, DATA_DIRS["{clave}"]),\n'
    contenido_cache = contenido_cache.rstrip(',\n') + ')\n\n'
    texto, n = patron_cache.subn(lambda m: contenido_cache, texto)
    if n != 1:
        raise RuntimeError(f'Clave de caché: esperaba una coincidencia, encontré {n}.')

    # 5. Priorizar aviso de desactualización sobre la ausencia de datos UMR.
    status_ant = (
        'elif df_ops.empty and not st.session_state.get(\'_do_load\'):\n'
        '    st.info("Configura el rango y los archivos en la barra lateral y aprieta **🔄 Cargar / actualizar datos** para procesar.")\n'
        'elif (\'df_ops\' in st.session_state) and (st.session_state.get(\'_cache_key\') != _cache_key):\n'
        '    st.warning("Cambiaron archivos o fechas desde la última carga. Aprieta **🔄 Cargar / actualizar datos** para refrescar.")'
    )
    status_nuevo = (
        'elif "_cache_key" not in st.session_state:\n'
        '    st.info("Configura el rango y presiona **🔄 Cargar / actualizar datos** para procesar los archivos guardados.")\n'
        'elif _recalcular:\n'
        '    st.warning("Hay archivos modificados o cambió el rango desde la última carga. Presiona **🔄 Cargar / actualizar datos** para actualizar los indicadores.")\n'
        'elif df_ops.empty:\n'
        '    st.info("No hay registros UMR para el rango cargado; puede haber datos de otras fuentes.")'
    )
    texto = reemplazar_unico(texto, status_ant, status_nuevo, 'Avisos de datos desactualizados')

    # No ejecutar la app durante el parche; solo comprobación sintáctica.
    compile(texto, 'app.py', 'exec')
    return texto


def aplicar(raiz: Path) -> bool:
    app = raiz / 'app.py'
    nuevo_modulo = raiz / 'gestion_archivos.py'
    if not app.is_file() or not nuevo_modulo.is_file():
        raise FileNotFoundError('Faltan app.py o gestion_archivos.py en la raíz del repositorio.')
    datos = app.read_bytes()
    newline = '\r\n' if b'\r\n' in datos else '\n'
    original = datos.decode('utf-8').replace('\r\n', '\n')
    if 'v22_integridad_archivos' in original:
        print('app.py ya tiene estas mejoras. No se modificó.')
        return False
    actualizado = transformar(original)
    copia = raiz / 'app.py.bak'
    if not copia.exists():
        shutil.copy2(app, copia)
    app.write_bytes(actualizado.replace('\n', newline).encode('utf-8'))
    requirements = raiz / 'requirements.txt'
    if requirements.exists():
        req = requirements.read_text(encoding='utf-8')
        if not any(l.strip().lower().startswith('requests') for l in req.splitlines()):
            with requirements.open('a', encoding='utf-8') as f:
                if not req.endswith(('\n', '\r')):
                    f.write('\n')
                f.write('requests>=2.31\n')
    print('OK: app.py actualizado, copia en app.py.bak y dependencia requests verificada.')
    return True


if __name__ == '__main__':
    aplicar(Path(__file__).resolve().parent)
