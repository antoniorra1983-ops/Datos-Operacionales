#!/usr/bin/env python3
"""Diseño lateral para Datos-Operacionales: navegación vertical y cargas desplegables.

No modifica los cálculos, nombres de variables, archivos de data/ ni filtros.
Uso:
  python aplicar_diseno_lateral.py --verificar
  python aplicar_diseno_lateral.py
  python -m py_compile app.py

Compatible con el bloque de navegación de app.py v23 y el gestor v24.
"""
from __future__ import annotations

import argparse
import ast
import hashlib
import os
from pathlib import Path
import re
import shutil
import tempfile

MARCADOR = '# ===== DISEÑO LATERAL v25 ====='
INICIO_SIDEBAR = '# --- 6. INTERFAZ Y SIDEBAR ---\n'
FIN_SIDEBAR = '# Combinar uploads\n'

NAVEGACION_ANTERIOR = '''with st.container(key="_topbar"):
    _cl, _cn = st.columns([1, 9], vertical_alignment="center")
    with _cl:
        _logo_header()
    with _cn:
        _seccion = st.radio("Sección", _SECCIONES, horizontal=True, key="_nav_seccion", label_visibility="collapsed")
'''

NAVEGACION_NUEVA = '''with st.container(key="_contenido_cabecera"):
    _hl, _hr = st.columns([5, 1], vertical_alignment="center")
    with _hl:
        st.caption("EFE VALPARAÍSO  ·  DATOS OPERACIONALES")
        st.markdown(f"### {_seccion}")
    with _hr:
        _logo_header()
'''

ESTILOS = '''# ===== DISEÑO LATERAL v25 =====
st.markdown("""<style>
[data-testid="stSidebar"] {border-right: 1px solid rgba(90,100,115,.18);}
.st-key-_nav_seccion [role="radiogroup"] {gap: .12rem;}
.st-key-_nav_seccion [role="radiogroup"] label {
    border: 1px solid transparent;
    border-radius: .58rem;
    padding: .40rem .50rem;
    transition: background-color .12s ease-in-out;
}
.st-key-_nav_seccion [role="radiogroup"] label:hover {
    background: rgba(0,81,149,.075);
}
.st-key-_nav_seccion [role="radiogroup"] label:has(input:checked) {
    background: rgba(0,81,149,.12);
    border-color: rgba(0,81,149,.25);
}
.st-key-_nav_seccion [role="radiogroup"] label p {
    font-size: .92rem;
    line-height: 1.28;
    white-space: normal;
}
</style>""", unsafe_allow_html=True)
'''


def _una(texto: str, bloque: str, nombre: str) -> None:
    n = texto.count(bloque)
    if n != 1:
        raise RuntimeError(f'{nombre}: se esperaba 1 coincidencia exacta; se encontraron {n}. No se modificó app.py.')


def _declaracion_secciones(codigo: str) -> str:
    """Extrae la lista real de secciones con AST; no inventa ni renombra pantallas."""
    arbol = ast.parse(codigo, filename='app.py')
    asignaciones = [
        n for n in arbol.body
        if isinstance(n, ast.Assign)
        and any(isinstance(t, ast.Name) and t.id == '_SECCIONES' for t in n.targets)
    ]
    if len(asignaciones) != 1:
        raise RuntimeError('No se encontró una única lista _SECCIONES de nivel superior.')
    n = asignaciones[0]
    lineas = codigo.splitlines(keepends=True)
    declaracion = ''.join(lineas[n.lineno - 1:n.end_lineno])
    _una(codigo, declaracion, 'Lista de secciones')
    if not isinstance(n.value, (ast.List, ast.Tuple)) or len(n.value.elts) < 8:
        raise RuntimeError('La lista de secciones cambió de formato. No se modifica la app.')
    return declaracion


def _separar_sidebar(bloque: str) -> tuple[str, str, str]:
    """Conserva filtros, cargadores y gestor existentes; solo cambia su posición."""
    if not bloque.startswith('with st.sidebar:\n'):
        raise RuntimeError('La barra lateral no coincide con la estructura conocida.')
    cuerpo = bloque[len('with st.sidebar:\n'):]
    ini_filtros = '    st.header("📅 Filtro Global")\n'
    ini_cargas = '    _upload_revision = st.session_state.get("_upload_revision", 0)\n'
    _una(cuerpo, ini_filtros, 'Filtro global')
    _una(cuerpo, ini_cargas, 'Cargadores')

    i_f = cuerpo.index(ini_filtros)
    i_c = cuerpo.index(ini_cargas)
    if i_f != 0 or i_c <= i_f:
        raise RuntimeError('Se alteró el orden conocido de filtros y cargas.')

    patron_guardados = re.compile(
        r'(?m)^    st\.divider\(\)\n(?=    (?:with st\.expander\("🗂️ Archivos guardados"\)|_gestor_archivos_v24\())'
    )
    coincidencias = list(patron_guardados.finditer(cuerpo))
    if len(coincidencias) != 1:
        raise RuntimeError('No se encontró de forma inequívoca el gestor de archivos guardados.')
    i_g = coincidencias[0].start()
    if i_g <= i_c:
        raise RuntimeError('Estructura de cargadores inesperada.')

    filtros = cuerpo[i_f:i_c].rstrip()
    cargas = cuerpo[i_c:i_g].rstrip()
    guardados = cuerpo[i_g:].strip('\n')

    # Mantiene el último separador en su nueva ubicación; no los duplica.
    if not filtros.endswith('    st.divider()'):
        raise RuntimeError('El bloque de filtros ya no termina donde se esperaba.')
    filtros = filtros[: -len('    st.divider()')].rstrip()
    if not guardados.startswith('    st.divider()\n'):
        raise RuntimeError('No se encontró el separador antes del gestor.')
    guardados = guardados[len('    st.divider()\n'):]

    if cargas.count('st.file_uploader(') != 11:
        raise RuntimeError('Se esperaban 11 cargadores. No se aplicará un parche parcial.')
    if 'guardar_archivo(uf, _ca)' not in cargas:
        raise RuntimeError('No se encontró la lógica de guardado de archivos.')
    return filtros, cargas, guardados


def transformar(codigo: str) -> str:
    if MARCADOR in codigo:
        # Idempotente: una segunda ejecución no vuelve a mover los componentes.
        ast.parse(codigo, filename='app.py')
        return codigo

    _una(codigo, INICIO_SIDEBAR, 'Inicio de la barra lateral')
    _una(codigo, FIN_SIDEBAR, 'Fin de la barra lateral')
    _una(codigo, NAVEGACION_ANTERIOR, 'Navegación superior')
    if '_CACHE_VERSION = "v23_estabilidad_navegacion"' not in codigo and '_CACHE_VERSION = "v24_gestor_fragmentado"' not in codigo:
        raise RuntimeError('Solo se admite app.py v23 o v24 verificado. No se hicieron cambios.')

    secciones = _declaracion_secciones(codigo)
    i = codigo.index(INICIO_SIDEBAR)
    j = codigo.index(FIN_SIDEBAR, i)
    filtros, cargas, guardados = _separar_sidebar(codigo[i+len(INICIO_SIDEBAR):j])

    cargas_indentadas = ''.join('    ' + linea if linea.strip() else '\n' for linea in cargas.splitlines(keepends=True))
    if not cargas_indentadas.endswith('\n'):
        cargas_indentadas += '\n'

    nueva_sidebar = (
        ESTILOS + '\n' + secciones + '\n' + INICIO_SIDEBAR
        + 'with st.sidebar:\n'
        + '    st.markdown("### 🚆 Datos operacionales")\n'
        + '    st.caption("EFE Valparaíso · Navegación")\n'
        + '    _seccion = st.radio("Secciones", _SECCIONES, key="_nav_seccion", horizontal=False, label_visibility="collapsed")\n'
        + '    st.divider()\n'
        + '    with st.expander("📥 Cargar archivos", expanded=False):\n'
        + '        st.caption("Formatos XLS, XLSX, XLSM y CSV · Los archivos guardados permanecen disponibles.")\n'
        + cargas_indentadas + '\n'
        + '    st.divider()\n'
        + filtros + '\n'
        + '    st.divider()\n'
        + guardados.rstrip() + '\n\n'
    )
    nuevo = codigo[:i] + nueva_sidebar + codigo[j:]
    # Quitar únicamente la lista antigua, que está debajo de # --- 8. TABS...
    # Antes de borrar, comprobar que ahora existe en dos ubicaciones.
    if nuevo.count(secciones) != 2:
        raise RuntimeError('La lista de secciones no aparece exactamente dos veces durante la migración.')
    ultimo = nuevo.rfind(secciones)
    if ultimo <= nuevo.index(INICIO_SIDEBAR):
        raise RuntimeError('No se identificó la declaración antigua de secciones.')
    nuevo = nuevo[:ultimo] + nuevo[ultimo+len(secciones):]
    nuevo = nuevo.replace(NAVEGACION_ANTERIOR, NAVEGACION_NUEVA, 1)

    arbol = ast.parse(nuevo, filename='app.py')
    compile(arbol, 'app.py', 'exec')
    if nuevo.count('st.file_uploader(') != codigo.count('st.file_uploader('):
        raise RuntimeError('No se conservó el número de cargadores.')
    if nuevo.count('_SECCIONES =') != 1:
        raise RuntimeError('La lista de pantallas se duplicó.')
    if len(nuevo.encode('utf-8')) < 500_000:
        raise RuntimeError('El resultado no parece contener el dashboard completo.')
    return nuevo


def aplicar(ruta: Path, verificar: bool = False) -> None:
    app = ruta / 'app.py'
    if not app.is_file():
        raise FileNotFoundError('Falta app.py en la misma carpeta que el actualizador.')
    datos = app.read_bytes()
    texto = datos.decode('utf-8-sig').replace('\r\n', '\n')
    if MARCADOR in texto:
        print('OK: diseño lateral ya instalado. No se modificó ningún archivo.')
        return
    modificado = transformar(texto)
    if verificar:
        print(f'OK SIMULACIÓN: navegación izquierda, carga desplegable, 11 cargadores; {len(modificado.encode("utf-8")):,} bytes; no se modificó app.py.')
        return

    backup_dir = ruta / '.respaldos_app'
    backup_dir.mkdir(exist_ok=True)
    sha = hashlib.sha256(datos).hexdigest()[:12]
    respaldo = backup_dir / f'app_antes_diseno_lateral_{sha}.py'
    if not respaldo.exists():
        shutil.copy2(app, respaldo)
    fd, temporal = tempfile.mkstemp(prefix='.diseno_lateral_', suffix='.tmp', dir=ruta)
    try:
        with os.fdopen(fd, 'wb') as f:
            f.write(modificado.encode('utf-8'))
            f.flush()
            os.fsync(f.fileno())
        os.replace(temporal, app)
    finally:
        if os.path.exists(temporal):
            os.unlink(temporal)
    print(f'OK: menú izquierdo y panel de carga ocultable instalados. app.py: {app.stat().st_size:,} bytes')
    print(f'Respaldo: {respaldo}')
    print('No se modificó ninguna carpeta data/.')


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('--verificar', action='store_true', help='Comprobar sin editar app.py')
    argumentos = p.parse_args()
    try:
        aplicar(Path(__file__).resolve().parent, verificar=argumentos.verificar)
    except (RuntimeError, OSError, ValueError, SyntaxError, UnicodeError) as e:
        print(f'ERROR: {e}')
        return 1
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
