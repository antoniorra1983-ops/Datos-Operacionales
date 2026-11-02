# Datos-Operacionales — corrección integral

Repositorio: `antoniorra1983-ops/Datos-Operacionales`.

## Estado de esta entrega

Los archivos del paquete se validaron localmente (12 pruebas unitarias exitosas). La conexión actual de ChatGPT a GitHub puede leer el repositorio, pero **no puede escribir en él** (HTTP 403). Por tanto, **los cambios no están publicados en GitHub** hasta ejecutar el instalador en Codespaces y hacer `git push`.

## Correcciones incorporadas

1. **Recuperación del dashboard completo:** el `app.py` de 71 líneas que quedó publicado es un *lanzador*. El instalador recupera del commit `a198bdaf58a7ae8c7fadda756c8046cc69e2b19e` el `app.py` original de 8.244 líneas, verifica su huella Git y agrega las correcciones. El nuevo `app.py` generado mide aproximadamente 610 KB; no depende de descargar el código durante el arranque de Streamlit.
2. **Lectura diferida** para evitar que la navegación entre secciones lea todos los archivos en disco innecesariamente.
3. **Reemplazo de archivos subidos con el mismo nombre** cuando tienen nuevo contenido, mediante guardado atómico.
4. **Detección de modificaciones** por nombre y metadatos, aunque no cambie el nombre del archivo.
5. **Protección contra nombres de archivo inválidos** y control para que no se restablezcan automáticamente archivos eliminados desde los cargadores.
6. **Avisos de datos desactualizados** cuando se cambia el período o el conjunto de archivos sin recalcular.
7. **Dependencias actualizadas:** `requests>=2.31`, requerida por el módulo meteorológico.
8. **Pruebas y CI:** 12 pruebas locales y configuración de GitHub Actions para detectar errores de sintaxis o un `app.py` incompleto.

## Cómo instalar

1. Sube el ZIP completo a la raíz de tu Codespaces. No borres `data/`.
2. En el terminal, desde la raíz del repositorio, ejecuta:

   ```bash
   unzip -o Datos_Operacionales_CORRECCION_TOTAL_LISTO.zip -d .
   python corregir_todo.py --publicar
   ```

3. Después, revisa GitHub → **Actions** y abre tu aplicación Streamlit con un mes que tenga archivos.

**No ejecutas directamente `aplicar_mejoras.py` sobre el lanzador actual**, ya que ese parche espera el código original completo. `corregir_todo.py` realiza primero la recuperación y luego la transformación correcta.

Si prefieres revisar antes de publicar, ejecuta `python corregir_todo.py`, luego `python -m unittest discover -s tests -v`, y finalmente `python corregir_todo.py --publicar`.

## Seguridad y limitaciones

- Crea `app.py.bak` como respaldo local y se niega a sobrescribir una versión de `app.py` que no reconozca.
- **No edita ni borra los datos de `data/`**.
- Las pruebas automáticas comprueban código, integridad y partes del almacenamiento. No sustituyen una prueba de Streamlit con datos reales ni certifican los cálculos energéticos.
- Streamlit Cloud puede reiniciar sesiones por inactividad; eso no se elimina con estos cambios. Si existen datos publicados en `data/`, se pueden recargar desde la aplicación.
- La carpeta del repositorio es pública; conviene revisar que los datos operacionales estén autorizados para publicación.

Consulta `INSTRUCCIONES.txt` para una guía resumida.
