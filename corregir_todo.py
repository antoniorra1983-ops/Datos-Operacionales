#!/usr/bin/env python3
"""Recupera y construye el app.py completo de Datos-Operacionales.

Colocar este archivo en la raíz del repo y ejecutar:
    python corregir_todo.py
    python corregir_todo.py --publicar     # además hace commit y push, si están autorizados

No toca data/, no modifica archivos de entrada y hace respaldo de app.py.
Rescata por SHA la versión completa del propio historial del repo, en lugar
 de depender del app.py de 71 líneas actualmente publicado.
"""
from __future__ import annotations

import argparse
import hashlib
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
from urllib.request import urlopen, Request

from aplicar_mejoras import transformar

ORIGINAL_COMMIT = 'a198bdaf58a7ae8c7fadda756c8046cc69e2b19e'
ORIGINAL_BLOB = '7e6f94b31a67fc913fa9061bbafd0d7457e5c1a8'
LANZADOR_BLOB = 'faee74cbca8832b7c5364eeebf810d401833897a'
REPO = 'antoniorra1983-ops/Datos-Operacionales'
VERSION = 'v22_integridad_archivos'
REQUIREMENTS = ('requests>=2.31',)


def sha_blob_git(content: bytes) -> str:
    return hashlib.sha1(b'blob ' + str(len(content)).encode() + b'\0' + content).hexdigest()


def _cmd(*args, cwd: Path, check=True):
    return subprocess.run(args, cwd=cwd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=check)


def recuperar_original(raiz: Path) -> bytes:
    """Usa un objeto local de Git (sin Internet) o la copia exacta del commit."""
    candidatos = [raiz / 'app_original.py', raiz / 'app.py.bak']
    for ruta in candidatos:
        if ruta.is_file():
            contenido = ruta.read_bytes()
            if sha_blob_git(contenido) == ORIGINAL_BLOB:
                return contenido
    try:
        raw = _cmd('git', 'show', f'{ORIGINAL_COMMIT}:app.py', cwd=raiz).stdout
        if sha_blob_git(raw) == ORIGINAL_BLOB:
            return raw
    except (OSError, subprocess.CalledProcessError):
        pass
    # Respaldo cuando el clon es superficial y el commit anterior no existe localmente.
    url = f'https://raw.githubusercontent.com/{REPO}/{ORIGINAL_COMMIT}/app.py'
    try:
        with urlopen(Request(url, headers={'User-Agent': 'Datos-Operacionales-Reparacion/1.0'}), timeout=30) as r:
            raw = r.read(2_000_000)
    except Exception as exc:
        raise RuntimeError(
            'No se encontró el código original completo en el historial local ni se pudo descargar. '
            'En Codespaces ejecuta: git fetch --unshallow, y vuelve a intentarlo.'
        ) from exc
    if sha_blob_git(raw) != ORIGINAL_BLOB:
        raise RuntimeError('La copia original no coincide con el SHA verificado. No se ha modificado app.py.')
    return raw


def escribir_atomico(destino: Path, contenido: bytes):
    destino.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(dir=destino.parent, prefix='.app_reparada_', suffix='.tmp', delete=False) as f:
        temporal = Path(f.name)
        try:
            f.write(contenido)
            f.flush()
        except Exception:
            temporal.unlink(missing_ok=True)
            raise
    try:
        temporal.replace(destino)
    finally:
        temporal.unlink(missing_ok=True)


def completar_requirements(raiz: Path):
    ruta = raiz / 'requirements.txt'
    if not ruta.is_file():
        raise FileNotFoundError('Falta requirements.txt; no se modifica app.py.')
    contenido = ruta.read_text(encoding='utf-8-sig')
    existentes = {s.strip().lower().split('=')[0].split('>')[0].split('<')[0].strip()
                  for s in contenido.splitlines() if s.strip() and not s.strip().startswith('#')}
    faltantes = [r for r in REQUIREMENTS if r.split('>=')[0].lower() not in existentes]
    if faltantes:
        nuevo = contenido.rstrip('\r\n') + '\n' + '\n'.join(faltantes) + '\n'
        escribir_atomico(ruta, nuevo.encode('utf-8'))
    return faltantes


def instalar(raiz: Path):
    raiz = raiz.resolve()
    app = raiz / 'app.py'
    required = ('app.py', 'aplicar_mejoras.py', 'gestion_archivos.py', 'requirements.txt')
    missing = [p for p in required if not (raiz / p).is_file()]
    if missing:
        raise FileNotFoundError('Faltan archivos en la raíz del repositorio: ' + ', '.join(missing))
    actual = app.read_bytes()
    actual_blob = sha_blob_git(actual)
    if VERSION.encode() in actual and actual_blob not in (ORIGINAL_BLOB, LANZADOR_BLOB):
        # No intervenir un código ya actualizado/modificado a mano.
        if len(actual) < 500_000 or b'from gestion_archivos import' not in actual:
            raise RuntimeError('Se detectó una etiqueta de mejoras, pero app.py no es el dashboard completo. No se sobrescribió.')
        compile(actual.decode('utf-8-sig'), str(app), 'exec')
        faltantes = completar_requirements(raiz)
        print('OK: app.py ya estaba actualizado y pasó la validación sintáctica.')
        if faltantes:
            print('Se agregaron dependencias:', ', '.join(faltantes))
        return False
    if actual_blob not in (ORIGINAL_BLOB, LANZADOR_BLOB):
        raise RuntimeError(
            'app.py cambió desde la última revisión. Se detuvo para no sobrescribir cambios tuyos. '
            f'SHA actual: {actual_blob}.')
    # Completar y verificar todos los datos ANTES de tocar el archivo activo.
    original = recuperar_original(raiz)
    texto = original.decode('utf-8-sig').replace('\r\n', '\n')
    actualizado = transformar(texto)
    if VERSION not in actualizado or len(actualizado) < 500_000:
        raise RuntimeError('El app.py resultante es demasiado corto o no incluye las mejoras. No se modifica.')
    compile(actualizado, str(app), 'exec')
    # Revisión de requisitos antes de sustituir app.py.
    completar_requirements(raiz)
    respaldo = raiz / 'app.py.bak'
    if not respaldo.exists():
        shutil.copy2(app, respaldo)
    salida = actualizado.encode('utf-8')
    escribir_atomico(app, salida)
    if app.stat().st_size != len(salida):
        raise RuntimeError('La escritura de app.py no coincide con los bytes esperados.')
    print('OK: app.py completo instalado:', f'{len(salida):,} bytes', '·', f'{actualizado.count(chr(10)) + 1:,} líneas')
    print('OK: respaldo previo preservado como app.py.bak (archivo local, no subir a GitHub).')
    print('OK: dependencias revisadas (requests incluido).')
    return True


def publicar(raiz: Path):
    # No absorber en el commit archivos ajenos que la persona tenga staged.
    candidatos = ['app.py', 'requirements.txt', 'aplicar_mejoras.py', 'gestion_archivos.py',
                  'README_MEJORAS.md', '.gitignore', '.github/workflows/verificacion.yml',
                  'tests/test_aplicar_mejoras.py', 'tests/test_gestion_archivos.py',
                  'tests/test_corregir_todo.py', 'corregir_todo.py']
    prev = _cmd('git', 'diff', '--cached', '--name-only', cwd=raiz).stdout.decode('utf-8', errors='replace').splitlines()
    if prev:
        raise RuntimeError('Hay archivos ya preparados (staged) en Git: ' + ', '.join(prev) + '. Publicación detenida para no mezclar cambios.')
    _cmd('git', 'add', '--', *(p for p in candidatos if (raiz / p).is_file()), cwd=raiz)
    status = _cmd('git', 'diff', '--cached', '--quiet', cwd=raiz, check=False)
    if status.returncode == 0:
        print('No hay diferencias nuevas para publicar.')
        return
    if status.returncode != 1:
        raise RuntimeError('No se pudo comprobar el estado de los archivos para commit.')
    _cmd('git', 'commit', '-m', 'Corregir y verificar dashboard Datos-Operacionales', cwd=raiz)
    _cmd('git', 'push', cwd=raiz)
    print('OK: cambios publicados en GitHub.')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--publicar', action='store_true', help='Crear commit y subir los cambios a GitHub tras instalar')
    args = parser.parse_args()
    raiz = Path(__file__).resolve().parent
    try:
        instalar(raiz)
        for nombre in ('app.py', 'gestion_archivos.py', 'aplicar_mejoras.py'):
            compile((raiz / nombre).read_bytes(), nombre, 'exec')
        print('OK: sintaxis de los tres módulos comprobada.')
        # Solo publicar después de comprobar sintaxis. Las pruebas pueden ejecutarse antes con unittest.
        if args.publicar:
            print('Comprobando pruebas unitarias...')
            _cmd(sys.executable, '-m', 'unittest', 'discover', '-s', 'tests', '-v', cwd=raiz)
            print('OK: pruebas aprobadas.')
            publicar(raiz)
        else:
            print('Para comprobar: python -m unittest discover -s tests -v')
            print('Para subir desde Codespaces: python corregir_todo.py --publicar')
    except (OSError, ValueError, RuntimeError, subprocess.CalledProcessError) as exc:
        if isinstance(exc, subprocess.CalledProcessError) and exc.stderr:
            detalle = exc.stderr.decode('utf-8', errors='replace').strip()
            print('ERROR: ' + str(exc) + '\n' + detalle, file=sys.stderr)
        else:
            print('ERROR: ' + str(exc), file=sys.stderr)
        return 1
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
