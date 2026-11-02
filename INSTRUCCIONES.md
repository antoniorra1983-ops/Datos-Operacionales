# Primera etapa de mejoras de Datos-Operacionales

**Objetivo:** actualizar el `app.py` completo existente (622 KB, rama main), sin reemplazarlo por un lanzador ni descargar códigos durante su arranque.

## Cambios

1. Integra la fase v22 de `aplicar_mejoras.py` ya disponible en el repositorio: lectura diferida de los archivos, reemplazo seguro y detección de modificaciones.
2. No muestra indicadores de un período anterior cuando el usuario cambia de rango y todavía no presiona «Cargar / actualizar datos».
3. Limpia diagnósticos y avisos de errores de cargas anteriores cuando ya no corresponden.
4. Continúa con los archivos restantes si falla uno de THDR V1/V2, Carga V1/V2 o Viajes, y registra el error por archivo.
5. Verifica que el resultado siga siendo el programa completo y que compile antes de escribir; guarda una copia de seguridad en `.respaldos_app/` y realiza la sustitución atómicamente.

## Instalación en Codespaces

El repositorio ya tiene `app.py` completo, `aplicar_mejoras.py` y `gestion_archivos.py`. **Solo necesitas agregar `mejorar_app.py` a la raíz**, junto a esos archivos. No borres `data/` ni reemplaces el `app.py` por este script.

Prueba sin modificar:

```bash
python mejorar_app.py --verificar
```

Instala y valida:

```bash
python mejorar_app.py
python -m py_compile app.py gestion_archivos.py aplicar_mejoras.py
```

Publícalo:

```bash
git add app.py mejorar_app.py
git commit -m "Mejorar estabilidad y carga del dashboard operacional"
git push
```

Si necesitas revertir, utiliza el respaldo de `.respaldos_app/` (no incluyas esa carpeta en tus commits). Ejecutar una segunda vez no duplica las modificaciones. **Este script no modifica `data/`.**

## Validaciones y límites

Las pruebas unitarias se encuentran en `tests/` del paquete para ejecutarlas localmente: `python -m unittest discover -s tests -v`. Es necesario probar la app con los datos reales después del despliegue para verificar el comportamiento funcional y el rendimiento; las pruebas de sintaxis por sí solas no demuestran que los indicadores sean correctos.

La integración de GitHub de esta conversación permite lectura pero sigue rechazando escrituras con error 403. Por ello la publicación en GitHub no se hizo desde esta conversación.
