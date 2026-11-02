#!/usr/bin/env python3
"""Mejoras v23 para el app.py COMPLETO de Datos-Operacionales.

Compatible con el app.py original de 622.271 bytes de GitHub y con v22.
No descarga el programa ni usa un lanzador. No modifica data/.

Uso en la raíz del repositorio:
    python mejorar_app.py --verificar   # simulación, no escribe
    python mejorar_app.py              # aplica la mejora al app.py original
    python -m py_compile app.py gestion_archivos.py
"""
from __future__ import annotations

import argparse
import ast
import hashlib
import os
from pathlib import Path
import shutil
import tempfile

from aplicar_mejoras import transformar as integrar_v22

_ORIGINAL_GIT_BLOB = "7e6f94b31a67fc913fa9061bbafd0d7457e5c1a8"
_V22 = '_CACHE_VERSION = "v22_integridad_archivos"'
_V23 = '_CACHE_VERSION = "v23_estabilidad_navegacion"'


def _reemplazar(texto: str, original: str, modificado: str, nombre: str) -> str:
    veces = texto.count(original)
    if veces != 1:
        raise RuntimeError(f'{nombre}: esperaba 1 coincidencia exacta, encontré {veces}. No se modifica app.py.')
    return texto.replace(original, modificado, 1)


def integrar_v23(texto: str) -> str:
    """Aplica la fase 2 solo sobre una app v22 previamente mejorada."""
    if _V23 in texto:
        return texto
    texto = _reemplazar(texto, _V22, _V23, 'Versión de cache')

    # Evita mezclar indicadores del período antiguo con la selección nueva.
    texto = _reemplazar(
        texto,
        "if _hay_archivos and 'df_ops' in st.session_state and not st.session_state.get('_do_load'):",
        "if _hay_archivos and 'df_ops' in st.session_state and not st.session_state.get('_do_load') and not _recalcular:",
        'Resultados desactualizados',
    )

    # Se borra la ficha vieja al iniciar un procesamiento nuevo sin facturación.
    texto = _reemplazar(
        texto,
        "elif _hay_archivos and st.session_state.get('_do_load'):\n    if f_umr_all:",
        "elif _hay_archivos and st.session_state.get('_do_load'):\n"
        "    st.session_state['_fichas_energia'] = []\n"
        "    if f_umr_all:",
        'Reinicio de fichas de facturación',
    )

    # En una carga sin incidencias ya no permanecen visibles los errores viejos.
    texto = _reemplazar(
        texto,
        "    if _errores_proc: st.session_state['_errores_proc']=_errores_proc",
        "    st.session_state['_errores_proc'] = _errores_proc",
        'Limpiar errores antiguos',
    )
    texto = _reemplazar(
        texto,
        "    if diag_thdr: st.session_state['diag_thdr']=diag_thdr",
        "    st.session_state['diag_thdr'] = diag_thdr",
        'Limpiar diagnóstico THDR antiguo',
    )

    # Si un archivo de THDR/Carga/Viajes está dañado se informa su error,
    # pero los demás archivos siguen procesándose en lugar de caer toda la app.
    helper = '''def _cargar_archivos_sin_interrumpir(_archivos, _lector, _tipo, _inicio, _fin, _errores):
    """Continúa procesando los demás archivos aunque uno falle."""
    _procesados = []
    for _f in _archivos:
        try:
            _dato = _lector(_f, _inicio, _fin)
            if _dato is not None:
                _procesados.append(_dato)
        except Exception as _exc:
            _errores[getattr(_f, 'name', 'archivo sin nombre')] = f"{_tipo}: {_exc}"
    return _procesados


'''
    texto = _reemplazar(
        texto,
        '# --- 6. INTERFAZ Y SIDEBAR ---\n',
        helper + '# --- 6. INTERFAZ Y SIDEBAR ---\n',
        'Procesamiento seguro de archivos',
    )
    cambios = {
        '        r1=[procesar_thdr_cacheado(f,start_date,end_date) for f in f_v1_all]':
        '        r1=_cargar_archivos_sin_interrumpir(f_v1_all, procesar_thdr_cacheado, "THDR V1", start_date, end_date, _errores_proc)',
        '        r2=[procesar_thdr_cacheado(f,start_date,end_date) for f in f_v2_all]':
        '        r2=_cargar_archivos_sin_interrumpir(f_v2_all, procesar_thdr_cacheado, "THDR V2", start_date, end_date, _errores_proc)',
        '        p_cv1 = [_carga_cacheado(f, start_date, end_date) for f in f_carga_v1_all]':
        '        p_cv1 = _cargar_archivos_sin_interrumpir(f_carga_v1_all, _carga_cacheado, "Carga V1", start_date, end_date, _errores_proc)',
        '        p_cv2 = [_carga_cacheado(f, start_date, end_date) for f in f_carga_v2_all]':
        '        p_cv2 = _cargar_archivos_sin_interrumpir(f_carga_v2_all, _carga_cacheado, "Carga V2", start_date, end_date, _errores_proc)',
        '        p_vj = [_viajes_cacheado(f, start_date, end_date) for f in f_viajes_all]':
        '        p_vj = _cargar_archivos_sin_interrumpir(f_viajes_all, _viajes_cacheado, "Viajes", start_date, end_date, _errores_proc)',
    }
    for original, modificado in cambios.items():
        texto = _reemplazar(texto, original, modificado, 'Lectura segura: ' + original[:27])

    # Asegura que los errores capturados por los lectores posteriores queden guardados.
    texto = _reemplazar(
        texto,
        "    st.session_state.update({'df_ops':df_ops",
        "    st.session_state['_errores_proc'] = _errores_proc\n"
        "    st.session_state.update({'df_ops':df_ops",
        'Guardar errores completos',
    )
    ast.parse(texto, filename='app.py')
    compile(texto, 'app.py', 'exec')
    return texto


def preparar(texto: str) -> str:
    if _V23 in texto:
        ast.parse(texto, filename='app.py')
        return texto
    if _V22 not in texto:
        texto = integrar_v22(texto)
    return integrar_v23(texto)


def _git_blob(contenido: bytes) -> str:
    return hashlib.sha1(b'blob ' + str(len(contenido)).encode('ascii') + b'\0' + contenido).hexdigest()


def _escribir_atomicamente(destino: Path, contenido: bytes) -> None:
    with tempfile.NamedTemporaryFile(dir=destino.parent, prefix='.mejora_', suffix='.tmp', delete=False) as tmp:
        ruta_tmp = Path(tmp.name)
        try:
            tmp.write(contenido)
            tmp.flush()
            os.fsync(tmp.fileno())
        except BaseException:
            ruta_tmp.unlink(missing_ok=True)
            raise
    try:
        os.replace(ruta_tmp, destino)
    finally:
        ruta_tmp.unlink(missing_ok=True)


def aplicar(ruta: Path, verificar: bool = False) -> tuple[bool, int]:
    ruta = ruta.resolve()
    archivos_necesarios = [ruta / 'app.py', ruta / 'aplicar_mejoras.py', ruta / 'gestion_archivos.py']
    faltantes = [p.name for p in archivos_necesarios if not p.is_file()]
    if faltantes:
        raise FileNotFoundError('Faltan archivos existentes del repositorio: ' + ', '.join(faltantes))
    app = ruta / 'app.py'
    original = app.read_bytes()
    linea = '\r\n' if b'\r\n' in original else '\n'
    origen = original.decode('utf-8-sig').replace('\r\n', '\n')
    if _V23 in origen:
        if len(origen) < 500_000 or 'from gestion_archivos import' not in origen:
            raise RuntimeError('El archivo dice v23 pero no contiene el dashboard completo. Se conserva sin cambios.')
        ast.parse(origen, filename=str(app))
        print('OK: app.py ya tiene las mejoras v23; no se modifica.')
        return False, len(original)
    if _V22 not in origen and _git_blob(original) != _ORIGINAL_GIT_BLOB:
        raise RuntimeError('app.py no coincide con la versión original validada ni con v22; se preservan tus modificaciones.')
    actualizado = preparar(origen)
    if len(actualizado) < 500_000 or '_CACHE_VERSION = "v23_estabilidad_navegacion"' not in actualizado:
        raise RuntimeError('El resultado está incompleto; NO se modificó app.py.')
    salida = actualizado.replace('\n', linea).encode('utf-8')
    if verificar:
        print(f'OK SIMULACIÓN: app completo {len(salida):,} bytes, {actualizado.count(chr(10)) + 1:,} líneas; sin escribir.')
        return True, len(salida)
    respaldos = ruta / '.respaldos_app'
    respaldos.mkdir(exist_ok=True)
    copia = respaldos / f'app_antes_mejoras_{_git_blob(original)[:12]}.py'
    if not copia.exists():
        shutil.copy2(app, copia)
    _escribir_atomicamente(app, salida)
    print(f'OK: app.py actualizado a v23 ({len(salida):,} bytes); respaldo en {copia.relative_to(ruta)}.')
    return True, len(salida)


def main() -> int:
    args = argparse.ArgumentParser(description=__doc__)
    args.add_argument('--verificar', action='store_true', help='Comprobar sin modificar archivos')
    opts = args.parse_args()
    try:
        aplicar(Path(__file__).resolve().parent, verificar=opts.verificar)
    except Exception as e:
        print('ERROR:', str(e))
        return 1
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
