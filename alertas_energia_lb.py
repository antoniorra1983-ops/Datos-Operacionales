#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Alertas automáticas por sobreconsumo de TRACCIÓN respecto a la LB de EFE.

Utiliza las mismas funciones de LB de app.py (extraídas por AST, sin abrir Streamlit).
Consume exclusivamente el snapshot seguro de data/_historial_v28 y envía mensajes
solo si --enviar está presente y el servidor SMTP está configurado.

Prueba local: python alertas_energia_lb.py --vista-previa
Envío: python alertas_energia_lb.py --enviar
"""
from __future__ import annotations

import argparse
import ast
import hashlib
import html
import io
import json
import os
from pathlib import Path
import re
import smtplib
import ssl
import tempfile
from datetime import datetime, timedelta
from email.message import EmailMessage
from zoneinfo import ZoneInfo
from zipfile import ZipFile

import numpy as np
import pandas as pd

# Destinatario temporal. Puede cambiarse desde una variable de GitHub Actions.
# Admite una o varias direcciones separadas por coma o punto y coma.
DESTINATARIOS = tuple(
    direccion.strip() for direccion in
    re.split(r"[,;]", os.environ.get("LB_ALERT_DESTINATARIOS", "a.reyes.eie@hotmail.com"))
    if direccion.strip()
)
if not DESTINATARIOS or any("\n" in d or "\r" in d or "@" not in d for d in DESTINATARIOS):
    raise ValueError("LB_ALERT_DESTINATARIOS debe contener direcciones válidas")
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


def registro_leer(ruta: Path) -> dict:
    if not ruta.exists():
        return {"schema": 1, "enviadas": {}}
    estado = json.loads(ruta.read_text(encoding="utf-8"))
    if estado.get("schema") != 1 or not isinstance(estado.get("enviadas"), dict):
        raise ValueError("Registro de alertas incompatible")
    return estado


def fmt(v, digitos=0):
    try:
        if pd.isna(v):
            return "—"
        return f"{float(v):,.{digitos}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (TypeError, ValueError):
        return "—"


def mensaje(alertas: pd.DataFrame, umbral_pct: float, umbral_kwh: float) -> EmailMessage:
    fecha_hora = datetime.now(ZONA).strftime("%d/%m/%Y %H:%M")
    n = len(alertas)
    asunto = f"[SGE] Alerta energética LB: {n} día(s) con sobreconsumo"
    msg = EmailMessage()
    msg["Subject"] = asunto
    msg["To"] = ", ".join(DESTINATARIOS)
    msg["From"] = os.environ.get("SMTP_FROM", os.environ.get("SMTP_USER", ""))
    filas_txt = []
    filas_html = []
    for _, r in alertas.iterrows():
        d = pd.Timestamp(r["Fecha"]).strftime("%d/%m/%Y")
        jornada = {"L": "Laboral", "S": "Sábado", "D/F": "Domingo/festivo"}.get(
            str(r.get("Tipo Día", "")), "—")
        vals = [d, jornada, fmt(r.get("E_real")), fmt(r.get("E_LB")),
                fmt(r.get("Exceso LB (kWh)")), fmt(r.get("Sobreconsumo LB %"), 1),
                fmt(r.get("Odómetro [km]")), fmt(r.get("PAX")), fmt(r.get("Servicios"))]
        filas_txt.append(" | ".join(vals))
        filas_html.append("<tr>" + "".join(
            f"<td style='border:1px solid #bbb;padding:6px'>{html.escape(v)}</td>" for v in vals
        ) + "</tr>")
    criterio = "energía real de TRACCIÓN mayor que la LB ajustada"
    if umbral_pct > 0:
        criterio += f" (desviación superior a {fmt(umbral_pct, 1)} %)"
    if umbral_kwh > 0:
        criterio += f" (exceso de al menos {fmt(umbral_kwh)} kWh)"
    intro = (f"Revisión SGE, {fecha_hora} (Chile). Se detectaron {n} día(s) con "
             f"{criterio}.\n\n"
             "Fecha | Jornada | Real kWh | LB kWh | Exceso kWh | Exceso % | Odómetro km | PAX | Servicios\n")
    pie = ("\n\nLB ajustada = intercepción + coeficiente 100/M × km 100/M + "
           "coeficiente SFE × km comerciales + estimación SFE en pruebas anterior al alta. "
           "Las fechas sin LB calculable fueron excluidas; una anomalía no demuestra causa raíz.\n"
           "Dashboard: https://github.com/antoniorra1983-ops/Datos-Operacionales")
    msg.set_content(intro + "\n".join(filas_txt) + pie)
    msg.add_alternative(
        "<html><body style='font:14px Arial,sans-serif;color:#1d2735'>"
        "<h2 style='color:#07558b'>Alerta de energía vs línea base — EFE Valparaíso</h2>"
        f"<p>Revisión SGE: {html.escape(fecha_hora)} (Chile). "
        f"Se detectaron {n} día(s) con {html.escape(criterio)}.</p>"
        "<table style='border-collapse:collapse;font-size:12px'><thead><tr>" + "".join(
            f"<th style='border:1px solid #bbb;padding:7px;background:#e7f2fa'>{h}</th>" for h in
            ("Fecha", "Jornada", "Real kWh", "LB kWh", "Exceso kWh", "Exceso %", "Odómetro km", "PAX", "Servicios")
        ) + "</tr></thead><tbody>" + "".join(filas_html) + "</tbody></table>"
        "<p><b>Modelo:</b> LB ajustada 100/M + SFE comerciales + estimación de SFE en pruebas. "
        "Sin datos suficientes no se envía alerta; la desviación requiere investigación antes "
        "de atribuir causas.</p></body></html>", subtype="html"
    )
    return msg


def enviar_por_smtp(msg: EmailMessage):
    host = os.environ.get("SMTP_HOST", "").strip()
    user = os.environ.get("SMTP_USER", "").strip()
    pwd = os.environ.get("SMTP_PASSWORD", "")
    sender = os.environ.get("SMTP_FROM", user).strip()
    if not (host and sender):
        raise RuntimeError("Configurar secretos SMTP_HOST y SMTP_FROM (o SMTP_USER)")
    port = int(os.environ.get("SMTP_PORT") or "587")
    seguro = os.environ.get("SMTP_SSL", "false").strip().lower() in ("1", "true", "yes")
    if bool(user) != bool(pwd):
        raise RuntimeError("SMTP_USER y SMTP_PASSWORD deben configurarse juntos")
    contexto = ssl.create_default_context()
    if seguro:
        with smtplib.SMTP_SSL(host, port, timeout=45, context=contexto) as smtp:
            if user:
                smtp.login(user, pwd)
            smtp.send_message(msg)
    else:
        with smtplib.SMTP(host, port, timeout=45) as smtp:
            smtp.ehlo()
            smtp.starttls(context=contexto)
            smtp.ehlo()
            if user:
                smtp.login(user, pwd)
            smtp.send_message(msg)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    modo = parser.add_mutually_exclusive_group(required=True)
    modo.add_argument("--vista-previa", action="store_true", help="Solo lista anomalías; NO envía")
    modo.add_argument("--enviar", action="store_true", help="Envía correo y registra fechas notificadas")
    parser.add_argument("--manifest", type=Path, default=RUTA_MANIFEST)
    parser.add_argument("--registro", type=Path, default=RUTA_REGISTRO)
    parser.add_argument("--app", type=Path, default=Path("app.py"))
    a = parser.parse_args()
    umbral_pct = float(os.environ.get("LB_ALERT_UMBRAL_PCT", "0"))
    umbral_kwh = float(os.environ.get("LB_ALERT_UMBRAL_KWH", "0"))
    ventana = int(os.environ.get("LB_ALERT_DIAS_REVISAR", "30"))
    if not (0 <= umbral_pct <= 300 and 0 <= umbral_kwh and 1 <= ventana <= 365):
        raise ValueError("Umbrales o días a revisar fuera de rango")
    motor = modelo_desde_app(a.app)
    fuentes = snapshot_leer(a.manifest)
    hallazgos = evaluar(fuentes, motor, umbral_pct, umbral_kwh, ventana)
    estado = registro_leer(a.registro)
    nuevos = hallazgos[~hallazgos["Fecha"].dt.strftime("%Y-%m-%d").isin(estado["enviadas"])].copy() if not hallazgos.empty else hallazgos
    print(f"Alertas en ventana de {ventana} días: {len(hallazgos)}; sin notificar: {len(nuevos)}")
    for _, row in nuevos.iterrows():
        print(f"  {row['Fecha']:%Y-%m-%d}: real={fmt(row['E_real'])} kWh, "
              f"LB={fmt(row['E_LB'])} kWh, desvío=+{fmt(row['Sobreconsumo LB %'],1)}%")
    if a.vista_previa or nuevos.empty:
        return
    email = mensaje(nuevos, umbral_pct, umbral_kwh)
    enviar_por_smtp(email)  # Si falla, NO se marca como enviado.
    momento = datetime.now(ZONA).isoformat()
    for _, row in nuevos.iterrows():
        fecha = row["Fecha"].strftime("%Y-%m-%d")
        estado["enviadas"][fecha] = {
            "enviado_en": momento,
            "exceso_kwh": round(float(row["Exceso LB (kWh)"]), 2),
            "exceso_pct": round(float(row["Sobreconsumo LB %"]), 2),
        }
    a.registro.parent.mkdir(parents=True, exist_ok=True)
    temp = a.registro.with_suffix(".tmp")
    temp.write_text(json.dumps(estado, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    temp.replace(a.registro)
    print("Correo enviado y registro actualizado.")


if __name__ == "__main__":
    main()
