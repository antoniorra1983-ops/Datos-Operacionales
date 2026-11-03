
python - <<'PY'
from pathlib import Path

archivo = Path("optimizar_app_v24.py")

codigo = r'''from __future__ import annotations

import argparse
import ast
import hashlib
import os
from pathlib import Path
import re
import shutil
import tempfile

VERSION_ANTERIOR = '_CACHE_VERSION = "v23_estabilidad_navegacion"'
VERSION_NUEVA = '_CACHE_VERSION = "v24_gestor_fragmentado"'

GESTOR = """
_fragmento_v24 = getattr(st, "fragment", None)

if _fragmento_v24 is None:
    def _fragmento_v24(func):
        return func

@_fragmento_v24
def _gestor_archivos_v24(carpetas, revision):
    with st.expander("🗂️ Archivos guardados", expanded=False):
        etiquetas = {
            "v1": "THDR Vía 1",
            "v2": "THDR Vía 2",
            "umr": "UMR",
            "seat": "SEAT",
            "bill": "Facturación",
            "carga_v1": "Pasajeros V1",
            "carga_v2": "Pasajeros V2",
            "viajes": "Viajes",
            "ide_lb": "Línea base",
            "incid": "Incidentes",
            "oit": "OIT",
        }

        categoria = st.selectbox(
            "Categoría",
            list(carpetas),
            format_func=lambda x: etiquetas.get(x, x),
            key="v24_categoria"
        )

        carpeta = carpetas[categoria]
        archivos = listar_archivos(carpeta)

        busqueda = st.text_input(
            "Buscar archivo",
            key="v24_busqueda"
        ).strip().casefold()

        encontrados = [
            p for p in archivos
            if busqueda in os.path.basename(p).casefold()
        ]

        total_paginas = max(1, (len(encontrados) + 24) // 25)

        pagina = st.number_input(
            "Página",
            min_value=1,
            max_value=total_paginas,
            value=1,
            key="v24_pagina"
        )

        inicio = (pagina - 1) * 25
        visibles = encontrados[inicio:inicio + 25]

        st.caption(
            f"{len(archivos)} archivos guardados. "
            f"{len(encontrados)} encontrados. "
            f"Página {pagina} de {total_paginas}."
        )

        if not visibles:
            st.info("No se encontraron archivos.")
            return

        st.dataframe(
            pd.DataFrame({
                "Archivo": [
                    os.path.basename(p) for p in visibles
                ]
            }),
            hide_index=True,
            use_container_width=True
        )

        elegido = st.selectbox(
            "Seleccionar archivo para eliminar",
            visibles,
            format_func=os.path.basename,
            key="v24_archivo"
        )

        confirmar = st.checkbox(
            "Confirmar eliminación",
            key="v24_confirmar"
        )

        if st.button(
            "🗑️ Eliminar archivo",
            disabled=not confirmar,
            key="v24_eliminar"
        ):
            if (
                os.path.dirname(os.path.realpath(elegido))
                != os.path.realpath(carpeta)
            ):
                st.error("Ruta de archivo inválida.")
                return

            try:
                os.remove(elegido)
            except OSError as exc:
                st.error(f"No se pudo eliminar: {exc}")
                return

            st.session_state["_upload_revision"] = revision + 1
            st.rerun()
"""

def reemplazar_unico(texto, anterior, nuevo):
    cantidad = texto.count(anterior)
    if cantidad != 1:
        raise RuntimeError(
            f"Se esperaban 1 coincidencia y se encontraron {cantidad}: "
            + anterior[:70]
        )
    return texto.replace(anterior, nuevo, 1)

def aplicar(verificar=False):
    ruta = Path(__file__).resolve().parent / "app.py"

    if not ruta.exists():
        raise FileNotFoundError("No se encontró app.py")

    original = ruta.read_text(encoding="utf-8-sig")

    if VERSION_NUEVA in original:
        print("La aplicación ya tiene la mejora v24.")
        return

    if VERSION_ANTERIOR not in original:
        raise RuntimeError(
            "La aplicación no está en v23. No se hicieron cambios."
        )

    texto = original.replace("\r\n", "\n")

    texto = reemplazar_unico(
        texto,
        "# --- 6. INTERFAZ Y SIDEBAR ---",
        GESTOR + "\n# --- 6. INTERFAZ Y SIDEBAR ---"
    )

    patron = re.compile(
        r'(?m)^    with st\\.expander\\("🗂️ Archivos guardados"\\):\\n'
        r'.*?(?=^# Combinar uploads)',
        re.DOTALL | re.MULTILINE
    )

    texto, cambios = patron.subn(
        '    _gestor_archivos_v24(DATA_DIRS, _upload_revision)\n\n',
        texto
    )

    if cambios != 1:
        raise RuntimeError(
            "No se encontró el bloque de archivos esperado."
        )

    texto = reemplazar_unico(
        texto,
        VERSION_ANTERIOR,
        VERSION_NUEVA
    )

    if len(texto.encode("utf-8")) < 500000:
        raise RuntimeError("El código resultante está incompleto.")

    ast.parse(texto)
    compile(texto, "app.py", "exec")

    if verificar:
        print(
            "OK SIMULACIÓN v24: "
            "código completo verificado; ningún archivo modificado."
        )
        return

    respaldos = ruta.parent / ".respaldos_app"
    respaldos.mkdir(exist_ok=True)

    sha = hashlib.sha256(
        ruta.read_bytes()
    ).hexdigest()[:12]

    respaldo = respaldos / f"app_antes_v24_{sha}.py"

    if not respaldo.exists():
        shutil.copy2(ruta, respaldo)

    with tempfile.NamedTemporaryFile(
        mode="w",
        encoding="utf-8",
        dir=ruta.parent,
        delete=False,
        suffix=".tmp"
    ) as temporal:
        temporal.write(texto)
        temporal.flush()
        os.fsync(temporal.fileno())
        temporal_path = Path(temporal.name)

    os.replace(temporal_path, ruta)

    print("OK: mejora v24 aplicada.")
    print("Respaldo:", respaldo)
    print("No se modificaron los archivos de data/.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--verificar", action="store_true")
    argumentos = parser.parse_args()
    aplicar(argumentos.verificar)
'''

archivo.write_text(codigo, encoding="utf-8")
print("Archivo creado:", archivo)
PY
