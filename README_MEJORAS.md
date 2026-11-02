# Revisión y mejoras: Datos-Operacionales

Repositorio de referencia: `antoniorra1983-ops/Datos-Operacionales` (rama `main`, inspeccionada 02-11-2026).

## Diagnóstico

1. La clase original `_ArchivoEnDisco` lee el archivo completo en su constructor. `f_*_all` se reconstruye en cada rerun de Streamlit, incluidos los cambios de sección; esto provoca lecturas del disco y consumo de RAM aunque no se pulse "Cargar / actualizar datos".
2. La clave `_cache_key` se compone de nombres: una nueva versión con igual nombre no se detecta por la clave.
3. `if not os.path.exists(dest): guardar_archivo(...)` ignora silenciosamente el contenido nuevo cuando el nombre ya existe.
4. Eliminar un archivo con el botón de papelera mientras sigue seleccionado en `file_uploader` puede restaurarlo en el rerun siguiente, dado que se vuelve a ejecutar el guardado antes de la tabla de archivos.
5. `guardar_archivo` une rutas con el nombre del usuario sin validarlo. No conviene permitir rutas arbitrarias en nombre del archivo.
6. La función de clima usa `requests`, pero la dependencia no figura expresamente en `requirements.txt`.
7. El `app.py` concentra más de 8.000 líneas, de modo que es recomendable modularizar progresivamente. Esta entrega evita rediseñar los cálculos energéticos.

## Cambios propuestos y preparados

- `gestion_archivos.py`: lectura de archivos **diferida**, guardado **atómico**, validación de extensión/nombre, detección de contenido diferente bajo el mismo nombre, firma por metadatos.
- `aplicar_mejoras.py`: cambia los bloques necesarios de `app.py`, agrega `requests>=2.31` a `requirements.txt`, verifica sintaxis antes de escribir y realiza respaldo de `app.py`.
- Cargadores de Streamlit con identificadores reiniciables: al eliminar un archivo, se limpian los archivos aún seleccionados para que no se reintroduzcan inmediatamente.
- Aviso visible cuando cambian las fechas o los archivos sin recargar los cálculos.
- Pruebas unitarias para persistencia, actualización, integridad, seguridad y lectura diferida.

## Aplicación de los cambios

1. Haz una copia de seguridad de tu repositorio o crea una rama en GitHub.
2. Coloca `gestion_archivos.py`, `aplicar_mejoras.py` y `tests/test_gestion_archivos.py` en la raíz y carpeta `tests` de tu proyecto respectivamente. **No borres `app.py`, `data/` ni `requirements.txt`.**
3. En el terminal de Codespaces (o tu computador) dentro de la raíz del repositorio, ejecuta:

   ```bash
   python aplicar_mejoras.py
   python -m unittest discover -s tests -v
   python -m py_compile app.py gestion_archivos.py
   ```

4. Prueba con un rango corto que tenga archivos disponibles (ej., enero de 2026). Verifica las tarjetas y descargas principales antes de publicar a usuarios.
5. Si todo funciona, añade `app.py`, `gestion_archivos.py`, `requirements.txt` y la carpeta `tests/` a GitHub. **No publiques `app.py.bak`**, es un respaldo local generado al aplicar el parche. Si usas Git CLI, añade los archivos explícitamente en vez de `git add .`.

## Recomendaciones pendientes

- Este parche es de **almacenamiento y estabilidad**, no certifica la exactitud de los indicadores de energía, THDR y pasajeros sin ejecutar un conjunto de datos de prueba conocido.
- Los datos incluidos en `data/` están en un repositorio público: revisar si sus contenidos son aptos para publicación.
- Streamlit Community Cloud puede reiniciar el proceso y borrar archivos añadidos solo al disco temporal. Para garantizar persistencia, los archivos deben guardarse en un repositorio o almacenamiento externo confiable.
- Fijar versiones de dependencias para despliegues reproducibles y añadir pruebas de integración de cada cargador de datos.

## Estado de publicación

La conexión de GitHub permitió inspeccionar el repositorio, pero **rechazó la creación de una rama con HTTP 403**, de modo que no se han aplicado cambios al repositorio remoto. El ZIP contiene cambios preparados y pruebas para aplicarlos con control de versiones.
