"""Funciones de almacenamiento de Datos-Operacionales.

No depende de Streamlit: se puede probar sin abrir la aplicación ni cargar el
historial ferroviario completo.
"""
from __future__ import annotations

import hashlib
from io import BytesIO
import os
from pathlib import Path
import tempfile

EXTENSIONES = frozenset({'.xls', '.xlsx', '.xlsm', '.csv'})


def _nombre_valido(nombre: str) -> str:
    if not isinstance(nombre, str) or not nombre or nombre in ('.', '..'):
        raise ValueError('Nombre de archivo vacío o inválido')
    if '/' in nombre or '\\' in nombre or Path(nombre).name != nombre:
        raise ValueError('No se admiten rutas dentro del nombre del archivo')
    if Path(nombre).suffix.lower() not in EXTENSIONES:
        raise ValueError('Solo se admiten archivos XLS, XLSX, XLSM y CSV')
    return nombre


def guardar_archivo(uf, carpeta: str) -> bool:
    """Guarda/actualiza un archivo de forma atómica.

    Devuelve True si el archivo fue creado o reemplazado; False si ya había una
    copia de contenido idéntico. Así una carga repetida no cambia la fecha de
    modificación y una versión nueva con el mismo nombre sí reemplaza la vieja.
    """
    nombre = _nombre_valido(uf.name)
    directorio = Path(carpeta)
    directorio.mkdir(parents=True, exist_ok=True)
    destino = directorio / nombre
    contenido = uf.getbuffer()

    if destino.is_file() and destino.stat().st_size == len(contenido):
        nuevo_hash = hashlib.sha256(contenido).digest()
        h = hashlib.sha256()
        with destino.open('rb') as original:
            for bloque in iter(lambda: original.read(1024 * 1024), b''):
                h.update(bloque)
        if h.digest() == nuevo_hash:
            return False

    descriptor, tmp = tempfile.mkstemp(dir=directorio, prefix='.carga_', suffix='.tmp')
    try:
        with os.fdopen(descriptor, 'wb') as salida:
            salida.write(contenido)
            salida.flush()
            os.fsync(salida.fileno())
        os.replace(tmp, destino)
    finally:
        if os.path.exists(tmp):
            os.unlink(tmp)
    return True


def listar_archivos(carpeta: str) -> list[str]:
    """Lista archivos reconocidos en disco, sin incluir temporales ni directorios."""
    try:
        return sorted(
            str(archivo)
            for archivo in Path(carpeta).iterdir()
            if archivo.is_file() and archivo.suffix.lower() in EXTENSIONES
        )
    except OSError:
        return []


class ArchivoEnDisco:
    """Envoltorio compatible con los lectores existentes, con lectura diferida.

    En cada clic de Streamlit solo se crean objetos livianos; los bytes del
    archivo se leen al comenzar el procesamiento, no al cambiar de pestaña.
    """
    def __init__(self, path: str):
        self.path = str(path)
        self.name = os.path.basename(path)
        self._bio: BytesIO | None = None

    def _buffer(self) -> BytesIO:
        if self._bio is None:
            with open(self.path, 'rb') as archivo:
                self._bio = BytesIO(archivo.read())
        return self._bio

    def read(self, *args, **kwargs):
        return self._buffer().read(*args, **kwargs)

    def seek(self, *args, **kwargs):
        return self._buffer().seek(*args, **kwargs)

    def tell(self, *args, **kwargs):
        return self._buffer().tell(*args, **kwargs)

    def seekable(self):
        return True

    def readable(self):
        return True

    def getbuffer(self):
        # Compatibilidad: la implementación anterior devolvía bytes.
        return self._buffer().getvalue()


def firma_fuentes(archivos, carpeta: str) -> tuple:
    """Huella barata por nombre, tamaño y metadatos del archivo persistido.

    La clave se actualiza cuando se reemplaza un archivo manteniendo su nombre
    (y tamaño), sin leer todo el historial en cada rerun de Streamlit.
    """
    firma = []
    for archivo in archivos:
        try:
            nombre = _nombre_valido(archivo.name)
        except ValueError:
            # Una carga rechazada en la interfaz no debe tumbar el dashboard
            # cuando Streamlit reconstruye las firmas de sus archivos.
            firma.append(('ARCHIVO_RECHAZADO', str(getattr(archivo, 'name', '')), getattr(archivo, 'size', None)))
            continue
        ruta = Path(carpeta) / nombre
        try:
            info = ruta.stat()
            firma.append((nombre, info.st_size, info.st_mtime_ns, info.st_ino))
        except OSError:
            # Si la carga al disco falló, no ocultar esa condición en la clave.
            firma.append((nombre, 'SIN_ARCHIVO_EN_DISCO', getattr(archivo, 'size', None)))
    return tuple(sorted(firma))
