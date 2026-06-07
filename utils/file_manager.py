import os
from io import BytesIO

# --- 1. CARPETAS DE DATOS LOCALES ---
DATA_DIRS = {
    "v1":"data/thdr_v1",
    "v2":"data/thdr_v2",
    "umr":"data/umr",
    "seat":"data/seat",
    "bill":"data/facturacion",
    "carga_v1":"data/carga_v1", 
    "carga_v2":"data/carga_v2"
}

# Crear carpetas si no existen físicamente
for _d in DATA_DIRS.values(): 
    os.makedirs(_d, exist_ok=True)

# --- 2. FUNCIONES DE LECTURA Y ESCRITURA ---
def guardar_archivo(uf, carpeta):
    """Guarda un archivo subido por Streamlit en el disco duro local."""
    with open(os.path.join(carpeta, uf.name), "wb") as out: 
        out.write(uf.getbuffer())

def listar_archivos(carpeta):
    """Devuelve la lista de Excels y CSVs guardados en una carpeta."""
    exts = ('.xls','.xlsx','.xlsm', '.csv')
    try: 
        return sorted([os.path.join(carpeta,f) for f in os.listdir(carpeta) if f.lower().endswith(exts)])
    except Exception: 
        return []

class _ArchivoEnDisco:
    """Simula un archivo subido para que Pandas lo lea directamente de la RAM (Acelera la carga)."""
    def __init__(self, path):
        self.name = os.path.basename(path)
        with open(path,'rb') as f: 
            self._bio = BytesIO(f.read())
    def read(self,*a,**kw):  return self._bio.read(*a,**kw)
    def seek(self,*a,**kw):  return self._bio.seek(*a,**kw)
    def tell(self,*a,**kw):  return self._bio.tell(*a,**kw)
    def seekable(self): return True
    def readable(self): return True
    def getbuffer(self): return self._bio.getvalue()

def combinar_fuentes(ul, carpeta):
    """Une los archivos recién subidos con los que ya estaban guardados en el historial."""
    nombres = {uf.name for uf in (ul or [])}
    return list(ul or []) + [_ArchivoEnDisco(p) for p in listar_archivos(carpeta)
                             if os.path.basename(p) not in nombres]
