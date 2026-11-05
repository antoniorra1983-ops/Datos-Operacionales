#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Notificar sobreconsumo de tracción vs LB mediante Issues de GitHub.

No usa SMTP ni requiere usar Watch: cada Issue de alerta se asigna al
propietario del repositorio (por defecto), de modo que GitHub pueda
notificarlo como responsable en su correo verificado si tiene email habilitado.

Ejecutar desde la raíz del repositorio:
  python alertas_energia_github.py --vista-previa
  python alertas_energia_github.py --publicar

El motor de línea base se obtiene del app.py confiable del repositorio.
"""
from __future__ import annotations

import argparse
import ast
import hashlib
import io
import json
import os
from pathlib import Path
import re
import tempfile
from datetime import datetime, timedelta
from urllib.request import Request, urlopen
from urllib.error import HTTPError
from urllib.parse import quote
from zoneinfo import ZoneInfo
from zipfile import ZipFile

import numpy as np
import pandas as pd

ZONA = ZoneInfo("America/Santiago")
RUTA_MANIFEST = Path("data/_historial_v28/manifest.json")
RUTA_REGISTRO = Path("data/_alertas_lb/enviadas.json")
FUNCIONES_LB = {
    "_sfe_unidad", "_sfe_mascara_comercial", "_bi_ide_km_sfe",
    "_bi_lb_energia_componentes", "_bi_lb_diaria_desde_umr",
}


def modelo_desde_app(ruta_app: Path):
    """Extrae exactamente el motor de LB sin ejecutar el resto del Dashboard.

    Esto evita mantener dos fórmulas que podrían quedar desincronizadas.
    El código fuente debe ser el app.py confiable del propio repositorio.
    """
    tree = ast.parse(ruta_app.read_text(encoding="utf-8"), filename=str(ruta_app))
    nodos = []
    halladas = set()
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) and node.name in FUNCIONES_LB:
            nodos.append(node)
            halladas.add(node.name)
        elif isinstance(node, ast.Assign) and any(
            isinstance(t, ast.Name) and t.id == "_SFE_COMERCIAL_DESDE" for t in node.targets
        ):
            nodos.append(node)
            halladas.add("_SFE_COMERCIAL_DESDE")
    faltantes = FUNCIONES_LB.union({"_SFE_COMERCIAL_DESDE"}) - halladas
    if faltantes:
        raise RuntimeError(f"Faltan funciones de línea base en app.py: {sorted(faltantes)}")
    mod = ast.fix_missing_locations(ast.Module(body=nodos, type_ignores=[]))
    ns = {"pd": pd, "np": np, "re": re}
    exec(compile(mod, str(ruta_app), "exec"), ns)  # solo fuente local de confianza
    return ns["_bi_lb_diaria_desde_umr"]


def snapshot_leer(ruta_manifest: Path, nombres=("df_ops", "df_ide_lb", "all_tr")):
    """Rearma ZIP por fragmentos validados y lee solo tres fuentes necesarias."""
    man = json.loads(ruta_manifest.read_text(encoding="utf-8"))
    if man.get("schema") != 2:
        raise ValueError("Esquema de historial desconocido")
    tipos = man.get("types", {})
    for nombre in nombres:
        if nombre not in tipos:
            raise ValueError(f"El snapshot no tiene {nombre}; requiere reconstruir historial")
    with tempfile.TemporaryDirectory(prefix="efe-alertas-lb-") as tmp:
        archivo = Path(tmp) / "snapshot.zip"
        with archivo.open("wb") as salida:
            for frag in man.get("parts", []):
                nombre = frag["name"]
                if not re.fullmatch(r"historial_\d{3}\.part", nombre):
                    raise ValueError("Nombre de fragmento de historial inválido")
                frag_ruta = ruta_manifest.parent / nombre
                digest = hashlib.sha256()
                total = 0
                with frag_ruta.open("rb") as origen:
                    for bloque in iter(lambda: origen.read(2 * 1024 * 1024), b""):
                        total += len(bloque)
                        digest.update(bloque)
                        salida.write(bloque)
                if total != int(frag["bytes"]) or digest.hexdigest() != frag["sha256"]:
                    raise ValueError(f"Fragmento dañado: {nombre}")
        with ZipFile(archivo) as z:
            resultado = {}
            for nombre in nombres:
                tipo = tipos[nombre]
                if tipo in ("frame", "records"):
                    tabla = pd.read_parquet(io.BytesIO(z.read(f"{nombre}.parquet")))
                    resultado[nombre] = tabla if tipo == "frame" else tabla.to_dict("records")
                elif tipo == "frame_json":
                    raw = json.loads(z.read(f"{nombre}.json"))
                    resultado[nombre] = pd.read_json(io.StringIO(raw), orient="table")
                elif tipo == "json":
                    resultado[nombre] = json.loads(z.read(f"{nombre}.json"))
                else:
                    raise ValueError(f"Formato no soportado de {nombre}: {tipo}")
            return resultado


def evaluar(datos: dict, funcion_lb, umbral_pct: float, umbral_kwh: float,
            dias_retroceso: int, hoy: datetime | None = None) -> pd.DataFrame:
    diario = datos["df_ops"]
    lb = datos["df_ide_lb"]
    umr = datos["all_tr"]
    if not isinstance(diario, pd.DataFrame) or diario.empty:
        raise RuntimeError("No existen datos operacionales diarios para evaluar")
    if not isinstance(lb, pd.DataFrame) or lb.empty:
        raise RuntimeError("Falta la fuente 'IDE Línea Base'; no enviar alertas sin LB")
    columnas = ["Fecha", "Odómetro [km]", "E_Tr", "IDE (kWh/km)"]
    faltan = set(columnas) - set(diario.columns)
    if faltan:
        raise RuntimeError(f"Faltan columnas operacionales: {sorted(faltan)}")
    copia = diario[columnas].copy()
    copia["Fecha"] = pd.to_datetime(copia["Fecha"], errors="coerce").dt.normalize()
    copia = copia.dropna(subset=["Fecha"])
    if copia["Fecha"].duplicated().any():
        raise RuntimeError("Existen varias filas por fecha: no se puede evaluar la LB sin duplicar")
    calculos = funcion_lb(copia, lb, umr)
    if calculos is None or calculos.empty:
        return pd.DataFrame()
    calculos = calculos[calculos["Línea base calculable"].fillna(False)].copy()
    if calculos.empty:
        return pd.DataFrame()
    fecha_local = (hoy or datetime.now(ZONA)).date()
    desde = fecha_local - timedelta(days=dias_retroceso)
    calculos = calculos[(calculos["Fecha"].dt.date >= desde) &
                        (calculos["Fecha"].dt.date <= fecha_local)]
    calculos["Sobreconsumo LB %"] = pd.to_numeric(calculos["Sobreconsumo LB %"], errors="coerce")
    calculos["Exceso LB (kWh)"] = pd.to_numeric(calculos["Exceso LB (kWh)"], errors="coerce")
    # Regla base: avisar ante CUALQUIER aumento real sobre la LB, aunque sea mínimo.
    # Los límites opcionales solo se aplican si el administrador configura un valor > 0.
    diferencia = calculos["E_real"] - calculos["E_LB"]
    supera = diferencia.gt(0)
    if umbral_pct > 0:
        supera &= calculos["Sobreconsumo LB %"].gt(umbral_pct)
    if umbral_kwh > 0:
        supera &= diferencia.ge(umbral_kwh)
    alertas = calculos[supera].copy()
    if alertas.empty:
        return alertas
    # Desgloses opcionales: están en df_ops y se cruzan por fecha, sin duplicar datos.
    extras = [c for c in ("Tipo Día", "PAX", "Servicios", "Tren-Km [km]") if c in diario.columns]
    if extras:
        extra = diario[["Fecha"] + extras].copy()
        extra["Fecha"] = pd.to_datetime(extra["Fecha"], errors="coerce").dt.normalize()
        alertas = alertas.merge(extra, on="Fecha", how="left", validate="one_to_one")
    return alertas.sort_values("Fecha", ascending=False).reset_index(drop=True)


def fmt(v, digitos=0):
    try:
        if pd.isna(v):
            return "—"
        return f"{float(v):,.{digitos}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (TypeError, ValueError):
        return "—"



def _jornada(v):
    return {'L': 'Laboral', 'S': 'Sábado', 'D/F': 'Domingo y festivo'}.get(str(v), 'Sin dato')


def _nombre_issue(fecha):
    return f"[SGE][LB] Sobreconsumo de tracción · {fecha}"


def _usuario_a_notificar(repositorio):
    """Asignatario de cada alerta. Por defecto es el dueño de este repositorio personal.

    Puede modificarse con GITHUB_ALERT_ASSIGNEE sin tocar el programa.
    La cuenta debe tener permiso para ser asignada a Issues del repositorio.
    """
    usuario = os.environ.get('GITHUB_ALERT_ASSIGNEE', '').strip() or repositorio.split('/', 1)[0]
    if not re.fullmatch(r'[A-Za-z0-9](?:[A-Za-z0-9-]{0,37}[A-Za-z0-9])?', usuario):
        raise ValueError('Nombre de usuario GitHub no válido para recibir alertas')
    return usuario


def _detalle_issue(fila, repositorio):
    dia = pd.Timestamp(fila['Fecha']).strftime('%Y-%m-%d')
    exceso = float(fila['Exceso LB (kWh)'])
    usuario = _usuario_a_notificar(repositorio)
    url = f"https://github.com/{repositorio}/actions"
    return (f"<!-- efe-alerta-lb:{dia} -->\n"
            f"**Responsable de revisar esta anomalía:** @{usuario}\n\n"
            f"## Sobreconsumo energético frente a línea base — {dia}\n\n"
            "| Indicador | Resultado |\n|---|---:|\n"
            f"| Tipo de jornada | {_jornada(fila.get('Tipo Día'))} |\n"
            f"| Tracción real | {fmt(fila['E_real'], 2)} kWh |\n"
            f"| Tracción esperada según LB | {fmt(fila['E_LB'], 2)} kWh |\n"
            f"| **Sobreconsumo** | **+{fmt(exceso, 2)} kWh** |\n"
            f"| **Desviación** | **+{fmt(fila['Sobreconsumo LB %'], 2)} %** |\n"
            f"| Odómetro UMR | {fmt(fila.get('Odómetro [km]'), 2)} km |\n"
            f"| Tren-km | {fmt(fila.get('Tren-Km [km]'), 2)} km |\n"
            f"| Pasajeros | {fmt(fila.get('PAX'))} |\n"
            f"| Servicios | {fmt(fila.get('Servicios'))} |\n\n"
            "**Criterio:** se publica cuando la energía real de tracción es estrictamente "
            "mayor que la LB ajustada, sin umbral mínimo de kWh ni porcentaje.\n\n"
            "**Modelo:** intercepción + coeficiente 100/M × km 100/M + coeficiente "
            "SFE × km comerciales + estimación SFE en pruebas, únicamente antes de "
            "la entrada comercial. Solo se consideran días con LB calculable.\n\n"
            "**Precaución:** la desviación indica un hallazgo que requiere investigación, "
            "no demuestra por sí sola una causa raíz.\n\n"
            f"[Ver ejecuciones y datos actualizados]({url}).\n")


def _api_github(ruta, *, metodo='GET', cuerpo=None, token=None, api='https://api.github.com'):
    token = token or os.environ.get('GITHUB_TOKEN', '')
    if not token:
        raise RuntimeError('Falta GITHUB_TOKEN: verifica el permiso Issues (write) en Actions')
    data = None if cuerpo is None else json.dumps(cuerpo, ensure_ascii=False).encode('utf-8')
    req = Request(f"{api.rstrip('/')}/{ruta.lstrip('/')}", data=data, method=metodo,
                  headers={'Authorization': f'Bearer {token}',
                           'Accept': 'application/vnd.github+json',
                           'X-GitHub-Api-Version': '2022-11-28',
                           'User-Agent': 'efe-datos-operacionales-alertas',
                           **({'Content-Type': 'application/json'} if data else {})})
    try:
        with urlopen(req, timeout=30) as response:
            return json.load(response)
    except HTTPError as e:
        detalle = e.read().decode('utf-8', errors='replace')[:700]
        raise RuntimeError(f'GitHub devolvió HTTP {e.code} en {metodo} {ruta}: {detalle}') from e


def _issues_registrados(repositorio, *, api_func=_api_github):
    """Detecta fechas notificadas aunque el issue original ya esté cerrado.

    No depende de un registro JSON ni de commits adicionales al repositorio.
    """
    rutas = f"repos/{repositorio}/issues"
    fechas = set()
    # Se esperan pocas incidencias, pero recorrer páginas evita omitir cerradas.
    for pagina in range(1, 31):
        issues = api_func(f'{rutas}?state=all&per_page=100&page={pagina}')
        for item in issues:
            if 'pull_request' in item:
                continue
            titulo = str(item.get('title', ''))
            encontrado = re.fullmatch(r'\[SGE\]\[LB\] Sobreconsumo de tracción · (\d{4}-\d{2}-\d{2})', titulo)
            if encontrado:
                fechas.add(encontrado.group(1))
        if len(issues) < 100:
            break
    else:
        raise RuntimeError('Demasiadas incidencias para comprobar duplicados. Revisa la paginación.')
    return fechas


def publicar_issues(alertas, repositorio, *, api_func=_api_github, existentes=None, simular=False):
    """Un issue por fecha; no duplicar incluso si el existente está cerrado."""
    anteriores = set(existentes) if existentes is not None else _issues_registrados(repositorio, api_func=api_func)
    publicados = []
    for _, fila in alertas.sort_values('Fecha').iterrows():
        dia = pd.Timestamp(fila['Fecha']).strftime('%Y-%m-%d')
        if dia in anteriores:
            print(f'{dia}: aviso ya registrado, no duplicar')
            continue
        if simular:
            print(f'{dia}: aviso pendiente (vista previa)')
            publicados.append({'fecha': dia, 'simulado': True})
            continue
        # Asignación directa: la notificación 'assigned' no requiere Watch.
        # El responsable es el propietario del repositorio salvo configuración explícita.
        usuario = _usuario_a_notificar(repositorio)
        creado = api_func(f'repos/{repositorio}/issues', metodo='POST', cuerpo={
            'title': _nombre_issue(dia),
            'body': _detalle_issue(fila, repositorio),
            'assignees': [usuario],
        })
        if not creado.get('number'):
            raise RuntimeError(f'GitHub no confirmó la creación de la alerta del {dia}')
        # GitHub puede aceptar una incidencia omitiendo un asignatario inválido.
        # No afirmar que llegará un email si no se confirmó la asignación.
        asignados = {str(a.get('login', '')).casefold() for a in creado.get('assignees', [])}
        if usuario.casefold() not in asignados:
            print(f'AVISO: la incidencia #{creado["number"]} se creó, pero GitHub no confirmó '
                  f'la asignación a @{usuario}. Revisa si el usuario tiene acceso al repositorio.')
        anteriores.add(dia)
        publicados.append({'fecha': dia, 'url': creado.get('html_url'), 'numero': creado['number'],
                           'asignado': usuario.casefold() in asignados})
        print(f"{dia}: creada incidencia #{creado['number']} — {creado.get('html_url', '')}")
    return publicados


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    modo = parser.add_mutually_exclusive_group(required=True)
    modo.add_argument('--vista-previa', action='store_true', help='Evalúa sin crear incidencias')
    modo.add_argument('--publicar', action='store_true', help='Crea incidencias y activa notificación de GitHub')
    parser.add_argument('--manifest', type=Path, default=RUTA_MANIFEST)
    parser.add_argument('--app', type=Path, default=Path('app.py'))
    parser.add_argument('--dias', type=int, default=30)
    args = parser.parse_args()
    if not 1 <= args.dias <= 365:
        raise ValueError('Los días de búsqueda deben estar entre 1 y 365')
    repositorio = os.environ.get('GITHUB_REPOSITORY', '').strip()
    if args.publicar and not re.fullmatch(r'[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+', repositorio):
        raise RuntimeError('Falta GITHUB_REPOSITORY (nombre propietario/repositorio)')
    modelo = modelo_desde_app(args.app)
    datos = snapshot_leer(args.manifest)
    hallazgos = evaluar(datos, modelo, umbral_pct=0, umbral_kwh=0, dias_retroceso=args.dias)
    print(f'Días con tracción real > LB calculable: {len(hallazgos)}')
    for _, fila in hallazgos.iterrows():
        print(f"{fila['Fecha']:%Y-%m-%d}: tracción {fmt(fila['E_real'], 2)} kWh; "
              f"LB {fmt(fila['E_LB'], 2)} kWh; exceso +{fmt(fila['Exceso LB (kWh)'], 2)} kWh")
    if not hallazgos.empty and args.publicar:
        anteriores = _issues_registrados(repositorio)
        nuevos = publicar_issues(hallazgos, repositorio, existentes=anteriores)
        print(f'Alertas nuevas publicadas: {len(nuevos)}; ya notificadas: {len(hallazgos) - len(nuevos)}')
    elif args.vista_previa:
        print('Solo vista previa: no se han publicado incidencias ni enviado correos')


if __name__ == '__main__':
    main()
