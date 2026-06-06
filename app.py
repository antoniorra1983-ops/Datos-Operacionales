import streamlit as st
import pandas as pd
import numpy as np
import re
import holidays
from io import BytesIO
from datetime import datetime, date, time
import plotly.graph_objects as go
import plotly.express as px
import os

# --- 1. CONFIGURACIÓN INICIAL ---
st.set_page_config(page_title="Gestión de Energía - Dashboard SGE", layout="wide", page_icon="🚆")
chile_holidays = holidays.Chile()
st.markdown("""<style>
.stMetric{background-color:#ffffff;padding:20px;border-radius:10px;
border-left:5px solid #005195;box-shadow:0 2px 4px rgba(0,0,0,0.05);}
div[data-testid="stMetricLabel"] > label {
    white-space: normal !important; 
    word-wrap: break-word !important; 
    min-height: 2.5rem;
    font-size: 0.95rem;
}
div[data-testid="stMetricValue"] {
    font-size: 1.6rem !important;
    word-wrap: break-word !important;
    white-space: normal !important;
}
.stDataFrame { overflow-x: auto; }
</style>""", unsafe_allow_html=True)

# --- 2. CONSTANTES DE RED Y CONFIGURACIONES ---
ESTACIONES = [
    'Puerto','Bellavista','Francia','Baron','Portales','Recreo','Miramar',
    'Viña del Mar','Hospital','Chorrillos','El Salto','Valencia','Quilpue',
    'El Sol','El Belloto','Las Americas','La Concepcion','Villa Alemana',
    'Sargento Aldea','Peñablanca','Limache'
]

# Diccionario inteligente de abreviaturas para visualización limpia (Data-to-Ink Ratio)
SHORT_NAMES_DICT = {
    'Puerto':'PUE', 'Bellavista':'BEL', 'Francia':'FRA', 'Baron':'BAR', 'Portales':'POR',
    'Recreo':'REC', 'Miramar':'MIR', 'Viña del Mar':'V.MAR', 'Hospital':'HOS',
    'Chorrillos':'CHO', 'El Salto':'SAL', 'Valencia':'VAL', 'Quilpue':'QUI',
    'El Sol':'SOL', 'El Belloto':'E.BEL', 'Las Americas':'AME', 'La Concepcion':'CON',
    'Villa Alemana':'V.ALE', 'Sargento Aldea':'S.ALD', 'Peñablanca':'PEÑ', 'Limache':'LIM'
}

KM_TRAMO = [0.7,0.7,0.8,1.7,2.1,1.4,0.9,0.9,1.0,1.5,7.4,2.3,1.9,2.0,1.1,1.2,0.9,0.6,1.3,12.73]
KM_ACUM  = [0.0]
for _k in KM_TRAMO: KM_ACUM.append(round(KM_ACUM[-1]+_k, 2))
KM_TOTAL = KM_ACUM[-1]

SPEED_PROFILE = [
    (90.6,    122.3,   31.7,   0,   0), (122.3,   215.3,   93.0,  52,  43), (215.3,   372.6,  157.3,  52,  43),
    (372.6,   577.2,  204.6,  52,  43), (577.2,   781.6,  204.4,  52,  43), (781.6,  1043.0,  261.4,  52,  43),
    (1043.0, 1377.0,  334.0,  52,  43), (1377.0, 1767.0,  390.0,  52,  43), (1767.0, 2202.0,  435.0,  42,  34),
    (2202.0, 2592.0,  390.0,  42,  34), (2592.0, 2960.5,  368.5,  74,  60), (2960.5, 3337.0,  376.5,  74,  60),
    (3337.0, 3448.4,  111.4,  74,  60), (3448.4, 3938.4,  490.0,  74,  60), (3938.4, 4328.4,  390.0,  66,  54),
    (4328.4, 4758.4,  430.0,  74,  60), (4758.4, 5188.4,  430.0,  52,  43), (5188.4, 5618.4,  430.0,  52,  43),
    (5618.4, 6034.4,  416.0,  52,  43), (6034.4, 6416.4,  382.0,  52,  43), (6416.4, 6913.0,  496.6,  74,  60),
    (6913.0, 7405.0,  492.0,  66,  54), (7405.0, 7816.4,  411.4,  66,  54), (7816.4, 8308.4,  492.0,  66,  54),
    (8308.4, 8695.0,  386.6,  66,  54), (8695.0, 9209.8,  514.8,  66,  54), (9209.8, 9622.2,  412.4,  66,  54),
    (9622.2,10171.1,  548.9,  66,  54), (10171.1,10530.5, 359.4,  52,  43), (10530.5,11020.5, 490.0,  74,  60),
    (11020.5,11513.5, 493.0,  74,  60), (11513.5,11920.0, 406.5,  74,  60), (11920.0,12088.4, 168.4,  74,  60),
    (12088.4,12176.0,  87.6,  74,  60), (12176.0,12578.0, 402.0,  74,  60), (12578.0,12724.8, 146.8,  74,  60),
    (12724.8,12861.7, 136.9,  74,  60), (12861.7,13359.7, 498.0, 120,  99), (13359.7,13847.7, 488.0, 120,  99),
    (13847.7,14337.7, 490.0,  74,  60), (14337.7,14828.7, 491.0,  52,  43), (14828.7,15325.7, 497.0,  52,  43),
    (15325.7,15823.7, 498.0,  52,  43), (15823.7,16321.7, 498.0,  52,  43), (16321.7,16812.7, 491.0,  52,  43),
    (16812.7,17317.7, 505.0,  52,  43), (17317.7,17809.7, 492.0,  52,  43), (17809.7,18301.7, 492.0,  74,  60),
    (18301.7,18788.7, 487.0,  74,  60), (18788.7,19281.7, 493.0,  74,  60), (19281.7,19772.7, 491.0,  74,  60),
    (19772.7,20265.7, 493.0,  74,  60), (20265.7,20754.7, 489.0,  74,  60), (20754.7,21250.7, 496.0,  66,  54),
    (21250.7,21337.7,  87.0,  52,  43), (21337.7,21632.1, 294.4,  52,  43), (21632.1,21739.7, 107.6,  74,  60),
    (21739.7,22061.7, 322.0,  74,  60), (22061.7,22251.2, 189.5, 102,  84), (22251.2,22357.7, 106.5, 102,  84),
    (22357.7,22812.7, 455.0,  74,  60), (22812.7,23265.7, 453.0,  74,  60), (23265.7,23660.7, 395.0,  74,  60),
    (23660.7,24155.7, 495.0, 102,  84), (24155.7,24650.7, 495.0, 102,  84), (24650.7,25145.7, 495.0,  74,  60),
    (25145.7,25343.7, 198.0,  74,  60), (25343.7,25483.0, 139.3,  74,  60), (25483.0,25725.0, 242.0,  74,  60),
    (25725.0,26219.0, 494.0,  74,  60), (26219.0,26614.0, 395.0,  74,  60), (26614.0,27025.5, 411.5,  74,  60),
    (27025.5,27457.0, 431.5,  74,  60), (27457.0,27837.0, 380.0,  74,  60), (27837.0,28317.0, 480.0,  74,  60),
    (28317.0,28712.0, 395.0,  74,  60), (28712.0,29180.0, 468.0,  74,  60), (29180.0,29565.0, 385.0,  74,  60),
    (29565.0,29817.0, 252.0,  74,  60), (29817.0,30122.0, 305.0,  74,  60), (30122.0,30464.0, 342.0,  66,  54),
    (30464.0,30849.0, 385.0,  74,  60), (30849.0,31332.6, 483.6, 102,  84), (31332.6,31817.6, 485.0, 120,  99),
    (31817.6,32307.6, 490.0, 120,  99), (32307.6,32802.6, 495.0, 120,  99), (32802.6,33297.6, 495.0, 120,  99),
    (33297.6,33792.6, 495.0, 120,  99), (33792.6,34282.6, 490.0, 120,  99), (34282.6,34767.6, 485.0, 120,  99),
    (34767.6,35246.6, 479.0, 120,  99), (35246.6,35725.3, 478.7, 120,  99), (35725.3,36223.3, 498.0, 102,  84),
    (36223.3,36704.5, 481.2,  74,  60), (36704.5,37194.0, 489.5,  74,  60), (37194.0,37683.5, 489.5,  74,  60),
    (37683.5,38172.0, 488.5, 102,  84), (38172.0,38665.3, 493.3, 120,  99), (38665.3,39153.0, 487.7, 120,  99),
    (39153.0,39642.4, 489.4, 120,  99), (39642.4,40134.0, 491.6, 120,  99), (40134.0,40621.8, 487.8, 120,  99),
    (40621.8,41100.8, 479.0, 120,  99), (41100.8,41601.5, 500.7, 120,  99), (41601.5,42089.1, 487.6, 102,  84),
    (42089.1,42588.5, 499.4,  66,  54), (42588.5,42785.5, 197.0,  66,  54), (42785.5,43057.2, 271.7,  42,  34),
    (43057.2,43273.1, 215.9,  42,  34), (43273.1,43305.0,  31.9,   0,   0)
]

# --- 3. FUNCIONES DE APOYO Y SEGURIDAD ---
def make_columns_unique(df):
    if not isinstance(df, pd.DataFrame) or df.empty: return df
    cols = pd.Series(df.columns)
    for dup in cols[cols.duplicated()].unique():
        cols[cols == dup] = [f"{dup}_{i}" if i != 0 else dup for i in range(sum(cols == dup))]
    df.columns = cols
    return df

def parse_latam_number(val):
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    s = str(val).strip().replace(' ','').replace('$','')
    s = re.sub(r'[^\d.,-]','',s)
    if not s: return 0.0
    if ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'): s = s.replace('.','').replace(',','.')
        else: s = s.replace(',','')
    elif ',' in s: s = s.replace(',','.')
    try: return float(s)
    except ValueError: return 0.0

def get_tipo_dia(fch):
    if fch is None: return "N/A"
    if fch in chile_holidays or fch.weekday() == 6: return "D/F"
    if fch.weekday() == 5: return "S"
    return "L"

def obtener_nombre_feriado(fch):
    if fch is None: return "No aplica"
    return chile_holidays.get(fch, "No es feriado")

def obtener_fecha_es(fecha):
    if pd.isna(fecha): return ""
    meses = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
    dias = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]
    return f"{dias[fecha.weekday()]} {fecha.day} {meses[fecha.month - 1]} {fecha.year}"

def minutos_a_hhmmss(minutos_float):
    if pd.isna(minutos_float): return "00:00:00"
    sign = "-" if minutos_float < 0 else ""
    m_abs = abs(minutos_float)
    h = int(m_abs // 60)
    m = int(m_abs % 60)
    s = int(round((m_abs - int(m_abs)) * 60))
    if s == 60:
        s = 0
        m += 1
    if m == 60:
        m = 0
        h += 1
    return f"{sign}{h:02d}:{m:02d}:{s:02d}"

# --- MOTOR DE EMPAREJAMIENTO DE ESTACIONES ---
def _norm(s):
    return str(s).upper().translate(str.maketrans("ÁÉÍÓÚÜÑ", "AEIOUUN"))

def get_col_thdr(df, estacion, tipo):
    if df is None or df.empty: return None
    est_n = _norm(estacion)
    for c in df.columns:
        c_n = _norm(c)
        if not c_n.endswith("_MIN"): continue
        if 'PROGRAMADA' in c_n: continue
        
        if tipo == 'SALIDA' and 'SALIDA' not in c_n: continue
        if tipo == 'LLEGADA' and 'LLEGADA' not in c_n: continue
        
        if est_n in c_n: return c
        
        alias_map = {
            "VINA DEL MAR": ["VINA", "V. MAR", "V MAR", "VIÑA"],
            "EL BELLOTO": ["BELLOTO"],
            "LAS AMERICAS": ["AMERICAS"],
            "LA CONCEPCION": ["CONCEPCION"],
            "VILLA ALEMANA": ["VILLA", "ALEMANA", "V. ALEMANA"],
            "SARGENTO ALDEA": ["SARGENTO", "ALDEA", "S. ALDEA"],
            "PENABLANCA": ["PENA BLANCA", "PENABLANCA", "PEÑA BLANCA"],
            "EL SALTO": ["SALTO"]
        }
        if est_n in alias_map:
            for alias in alias_map[est_n]:
                if alias in c_n: return c
    return None

def extract_series(df, col_name):
    s = df[col_name]
    if isinstance(s, pd.DataFrame):
        return pd.to_numeric(s.iloc[:, 0], errors='coerce')
    return pd.to_numeric(s, errors='coerce')

# --- 3b. MOTOR DE DIAGNÓSTICO DE ANOMALÍAS ---
_MAD_K = 1.4826
_METRICAS_ANOM = {
    "E_Total": "energía total", "E_Tr": "tracción", "E_12": "auxiliares 12 kV",
    "IDE (kWh/km)": "eficiencia (kWh/km)", "kWh_por_PAX": "kWh/pasajero",
    "Noche_kWh": "consumo nocturno", "Pico_kW": "pico de potencia", "Servicios": "servicios",
}

def _robust_z(serie):
    s = serie.astype(float)
    med = s.median()
    mad = (s - med).abs().median()
    esc = _MAD_K * mad if mad > 0 else s.std(ddof=0)
    if not esc or np.isnan(esc):
        return pd.Series(0.0, index=s.index)
    return (s - med) / esc

def _perfil_horario_diario(all_prmte_full, all_fact_full, df_ops=None):
    datos, freq = (all_prmte_full, 15) if all_prmte_full else (all_fact_full, 60)
    if not datos:
        return pd.DataFrame(columns=["Fecha", "Noche_kWh", "Pico_kW"])
    h = pd.DataFrame(datos)
    h["Fecha"] = pd.to_datetime(h["Fecha"]).dt.normalize()
    h["Hora_n"] = h["Hora"].astype(str).str.slice(0, 2).apply(lambda x: int(x) if str(x).isdigit() else -1)
    
    if df_ops is not None and not df_ops.empty:
        mapa_tipo = df_ops.set_index('Fecha')['Tipo Día'].to_dict()
        h['Tipo Día'] = h['Fecha'].map(mapa_tipo)
    else:
        h['Tipo Día'] = h['Fecha'].apply(lambda x: get_tipo_dia(x.date()))

    def _is_noche_dinamico(row):
        limite = 6 if row['Tipo Día'] == 'L' else (7 if row['Tipo Día'] == 'S' else 8)
        return 0 <= row['Hora_n'] < limite
        
    h['Es_Noche'] = h.apply(_is_noche_dinamico, axis=1)
    noche = h[h['Es_Noche']].groupby("Fecha")["Consumo"].sum().rename("Noche_kWh")
    pico = (h.groupby("Fecha")["Consumo"].max() * (60.0 / freq)).rename("Pico_kW")
    return pd.concat([noche, pico], axis=1).reset_index()

def _contexto_dia(fecha, cc, tt):
    fecha = pd.to_datetime(fecha).normalize()
    out = {"Doble_pct": np.nan, "Est_critica": None, "Ocup_max": np.nan, "Fuente_op": "—"}
    fuentes = []
    if cc is not None and not cc.empty:
        g = cc[cc["Fecha"] == fecha]
        if not g.empty:
            fuentes.append("Pasajeros")
            cols = {_norm(c): c for c in g.columns}
            c_m2 = next((cols[k] for k in cols if "MOTRIZ 2" in k), None)
            c_est = next((cols[k] for k in cols if "ESTACI" in k and "MAX" in k), None)
            c_cm = next((cols[k] for k in cols if "CARGA" in k and "MAX" in k), None)
            if c_m2 is not None:
                out["Doble_pct"] = 100.0 * (pd.to_numeric(g[c_m2], errors="coerce").fillna(0) > 0).mean()
            if c_cm is not None:
                out["Ocup_max"] = pd.to_numeric(g[c_cm], errors="coerce").max()
            if c_est is not None:
                m = g[c_est].astype(str).replace("nan", np.nan).dropna()
                if not m.empty and not m.mode().empty:
                    out["Est_critica"] = m.mode().iloc[0]
    if tt is not None and not tt.empty and "Fecha" in tt.columns:
        g = tt[tt["Fecha"] == fecha]
        if not g.empty:
            fuentes.append("THDR")
            if pd.isna(out["Doble_pct"]) and "Unidad" in g.columns:
                out["Doble_pct"] = 100.0 * g["Unidad"].astype(str).str.upper().eq("M").mean()
    if fuentes:
        out["Fuente_op"] = " + ".join(fuentes)
    return out

def _dur_via(t, est_sal, est_lleg):
    if t is None or t.empty or "Fecha_Op" not in t.columns:
        return pd.DataFrame(columns=["Fecha", "dur"])
    
    c_sal = get_col_thdr(t, est_sal, "SALIDA")
    c_lleg = get_col_thdr(t, est_lleg, "LLEGADA")
    
    if not c_sal or not c_lleg:
        return pd.DataFrame(columns=["Fecha", "dur"])
        
    s_sal = extract_series(t, c_sal)
    s_lleg = extract_series(t, c_lleg)
    
    dur = s_lleg - s_sal
    dur = dur.apply(lambda x: x + 1440 if pd.notna(x) and x < -200 else x)
    out = pd.DataFrame({"Fecha": pd.to_datetime(t["Fecha_Op"]).dt.normalize(), "dur": dur})
    return out[(out["dur"] > 30) & (out["dur"] < 120)]

def _thdr_tiempos(df_thdr_v1, df_thdr_v2):
    alld = pd.concat([_dur_via(df_thdr_v1, "PUERTO", "LIMACHE"),
                      _dur_via(df_thdr_v2, "LIMACHE", "PUERTO")], ignore_index=True)
    if alld.empty:
        return pd.DataFrame(columns=["Fecha", "Viaje_prom", "Brecha_min"])
    gb = alld.groupby("Fecha")["dur"]
    return pd.DataFrame({"Viaje_prom": gb.mean(), "Brecha_min": gb.max() - gb.min()}).reset_index()

def diagnosticar_anomalias(df_ops, all_prmte_full=None, all_fact_full=None,
                           df_carga_v1=None, df_carga_v2=None, df_thdr_v1=None, df_thdr_v2=None,
                           z_alerta=2.5, z_fuerte=3.5):
    if df_ops is None or df_ops.empty:
        return pd.DataFrame()
    d = df_ops[df_ops["E_Total"] > 0].copy().reset_index(drop=True)
    if d.empty:
        return d
    d["Fecha"] = pd.to_datetime(d["Fecha"]).dt.normalize()
    if "kWh_por_PAX" not in d.columns:
        d["kWh_por_PAX"] = d["E_Tr"] / d["PAX"].replace(0, np.nan)
        
    perfil = _perfil_horario_diario(all_prmte_full, all_fact_full, d)
    if not perfil.empty:
        d = d.merge(perfil, on="Fecha", how="left")
    for c in ["Noche_kWh", "Pico_kW"]:
        if c not in d.columns:
            d[c] = np.nan

    zcols = {}
    for col in _METRICAS_ANOM:
        if col not in d.columns:
            continue
        zc = "z_" + col
        zcols[col] = zc
        d[zc] = np.nan
        for tipo, idx in d.groupby("Tipo Día").groups.items():
            sub = d.loc[idx, col]
            valid = sub[(sub.notna()) & (sub != 0)]
            if len(valid) < 4:
                continue
            d.loc[valid.index, zc] = _robust_z(valid)

    _cl = [x for x in [df_carga_v1, df_carga_v2] if x is not None and not x.empty]
    cc = pd.concat(_cl, ignore_index=True) if _cl else pd.DataFrame()
    if not cc.empty:
        cc["Fecha"] = pd.to_datetime(cc["Fecha"]).dt.normalize()
    _tl = [x for x in [df_thdr_v1, df_thdr_v2] if x is not None and not x.empty]
    tt = pd.concat(_tl, ignore_index=True) if _tl else pd.DataFrame()
    if not tt.empty and "Fecha_Op" in tt.columns:
        tt["Fecha"] = pd.to_datetime(tt["Fecha_Op"]).dt.normalize()
    ctx = pd.DataFrame([_contexto_dia(f, cc, tt) for f in d["Fecha"]], index=d.index)
    for c in ["Doble_pct", "Est_critica", "Ocup_max", "Fuente_op"]:
        d[c] = ctx[c].values
    d = d.merge(_thdr_tiempos(df_thdr_v1, df_thdr_v2), on="Fecha", how="left")

    niveles, sevs, diags = [], [], []
    for _, r in d.iterrows():
        fired = {c: r[zcols[c]] for c in zcols if pd.notna(r.get(zcols[c])) and abs(r[zcols[c]]) >= z_alerta}
        if not fired:
            niveles.append("OK"); sevs.append(0.0); diags.append(""); continue
            
        sev = max(abs(v) for v in fired.values())
        niveles.append("ANOMALÍA" if sev >= z_fuerte else "ATENCIÓN")
        sevs.append(sev)
        
        z_en = fired.get("E_Total", fired.get("E_Tr", None))
        sintoma_principal = "Desviación Operativa Detectada"
        explicacion = []

        if z_en is not None and z_en > 0:
            sintoma_principal = "📈 Sobreconsumo Crítico de Energía"
            if fired.get("IDE (kWh/km)", 0) >= z_alerta:
                explicacion.append("Pérdida severa de eficiencia traccional (IDE disparado). Posible conducción en 'Stop-and-Go' o exceso de masa inercial.")
            elif fired.get("Servicios", 0) >= z_alerta:
                explicacion.append("Exceso de Oferta: Aumento de energía justificado por un despacho masivo de trenes superior al promedio de este tipo de día.")
            else:
                explicacion.append("Alza en consumo bruto sin justificación aparente en los kilómetros recorridos.")
        elif z_en is not None and z_en < 0:
            sintoma_principal = "📉 Caída Atípica de Consumo"
            if fired.get("Servicios", 0) <= -z_alerta:
                explicacion.append("Reducción de Oferta: Operaron significativamente menos trenes.")
            elif fired.get("IDE (kWh/km)", 0) <= -z_alerta:
                explicacion.append("Alta Eficiencia Traccional detectada (menos kWh/km de lo normal).")
            else:
                explicacion.append("Posible pérdida de datos de telemetría/facturación en este día.")

        if fired.get("E_12", 0) >= z_alerta and fired.get("Noche_kWh", 0) >= z_alerta:
            sintoma_principal = "🌙 Alerta de Consumo Parásito"
            explicacion = ["Consumo nocturno disparado. Posibles trenes energizados operando en vacío durante la madrugada."]
        elif fired.get("Noche_kWh", 0) >= z_alerta:
            explicacion.append("Pico anómalo de demanda eléctrica durante la ventana nocturna.")

        otras = [f"{_METRICAS_ANOM.get(m, m)} {'alto' if z>0 else 'bajo'}" for m, z in fired.items() if m not in ["E_Total", "E_Tr", "IDE (kWh/km)", "Servicios", "E_12", "Noche_kWh"]]
        if otras:
            explicacion.append("Anomalías adicionales en: " + ", ".join(otras) + ".")

        texto_diag = f"**{sintoma_principal}**\n\n" + "\n".join([f"- {e}" for e in explicacion])
        diags.append(texto_diag)

    d["Nivel"] = niveles
    d["Severidad"] = sevs
    d["Diagnóstico"] = diags
    return d

# --- 4. PERSISTENCIA EN DISCO ---
DATA_DIRS = {
    "v1":"data/thdr_v1","v2":"data/thdr_v2","umr":"data/umr",
    "seat":"data/seat","bill":"data/facturacion",
    "carga_v1":"data/carga_v1", "carga_v2":"data/carga_v2"
}
for _d in DATA_DIRS.values(): os.makedirs(_d, exist_ok=True)

def guardar_archivo(uf, carpeta):
    with open(os.path.join(carpeta, uf.name), "wb") as out: out.write(uf.getbuffer())

def listar_archivos(carpeta):
    exts = ('.xls','.xlsx','.xlsm', '.csv')
    try: return sorted([os.path.join(carpeta,f) for f in os.listdir(carpeta) if f.lower().endswith(exts)])
    except Exception: return []

class _ArchivoEnDisco:
    def __init__(self, path):
        self.name = os.path.basename(path)
        with open(path,'rb') as f: self._bio = BytesIO(f.read())
    def read(self,*a,**kw):  return self._bio.read(*a,**kw)
    def seek(self,*a,**kw):  return self._bio.seek(*a,**kw)
    def tell(self,*a,**kw):  return self._bio.tell(*a,**kw)
    def seekable(self): return True
    def readable(self): return True
    def getbuffer(self): return self._bio.getvalue()

def combinar_fuentes(ul, carpeta):
    nombres = {uf.name for uf in (ul or [])}
    return list(ul or []) + [_ArchivoEnDisco(p) for p in listar_archivos(carpeta)
                             if os.path.basename(p) not in nombres]

# --- 5. FUNCIONES DE PROCESAMIENTO CORE ---
def convertir_a_minutos(val):
    if pd.isna(val) or str(val).strip() == "": return None
    try:
        if isinstance(val, (datetime, time)): return val.hour*60 + val.minute + val.second/60.0
        sv = str(val).strip()
        m = re.search(r'(\d{1,2}):(\d{2}):(\d{2})', sv)
        if m: return int(m.group(1))*60 + int(m.group(2)) + int(m.group(3))/60.0
        m = re.search(r'(\d{1,2}):(\d{2})', sv)
        if m: return int(m.group(1))*60 + int(m.group(2))
        return None
    except Exception: return None

def parsear_fecha_nombre(nombre):
    nombre = str(nombre)
    patrones = [
        (r'(\d{2})[-_](\d{2})[-_](\d{4})', lambda m: date(int(m.group(3)), int(m.group(2)), int(m.group(1)))),
        (r'(\d{4})[-_](\d{2})[-_](\d{2})', lambda m: date(int(m.group(1)), int(m.group(2)), int(m.group(3)))),
        (r'(\d{8})', lambda m: date(int(m.group(1)[4:]), int(m.group(1)[2:4]), int(m.group(1)[:2]))),
        (r'(\d{6})', lambda m: date(2000+int(m.group(1)[4:]), int(m.group(1)[2:4]), int(m.group(1)[:2])))
    ]
    for patron, parser in patrones:
        m = re.search(patron, nombre)
        if m:
            try: return parser(m), f"Match ({m.group()})"
            except ValueError: pass
    return None, f"sin fecha en: '{nombre}'"

def procesar_thdr_eficiente(file, start_date, end_date):
    nombre = getattr(file, 'name', str(file))
    diag = {"archivo": nombre, "fecha_parseada": None, "en_rango": None, "filas": 0, "error": None}
    try:
        fch_date, desc = parsear_fecha_nombre(nombre)
        diag["fecha_parseada"] = desc
        if fch_date is None:
            diag["error"] = "No se encontró fecha en el nombre"; return pd.DataFrame(), diag
        
        diag["en_rango"] = f"{start_date}≤{fch_date}≤{end_date}"
        if not (start_date <= fch_date <= end_date):
            diag["error"] = "Fecha fuera del rango"; return pd.DataFrame(), diag
            
        fch_dt = pd.to_datetime(fch_date).normalize()
        engine = "xlrd" if nombre.lower().endswith(".xls") else "openpyxl"
        df_raw = pd.read_excel(file, header=None, engine=engine)
        
        r0 = df_raw.iloc[0].copy(); r0[0] = np.nan
        h1 = r0.ffill().astype(str); h2 = df_raw.iloc[1].fillna('').astype(str)
        
        cols = []
        for stn, tip in zip(h1, h2):
            stn, tip = str(stn).strip(), str(tip).strip()
            if stn == 'nan' or stn == '': cols.append(tip if tip else '_vacio')
            else: cols.append(f"{stn}_{tip}" if tip else stn)
            
        df = df_raw.iloc[5:].copy().reset_index(drop=True)
        n = len(df.columns)
        cols_adj = cols[:n] if len(cols) >= n else cols + [f"_C{j}" for j in range(n-len(cols))]
        df.columns = cols_adj
        df = make_columns_unique(df).dropna(how='all', axis=0).reset_index(drop=True)
        
        for col in df.columns:
            if any(k in str(col) for k in ['Hora Llegada','Hora Salida','Hora Salida Programada']):
                df[f"{col}_min"] = df[col].apply(convertir_a_minutos)
                
        if 'Unidad' in df.columns:
            df['Unidad'] = df['Unidad'].fillna('S').replace('','S')
        else:
            c_m2 = next((c for c in df.columns if 'Motriz 2' in str(c)), None)
            df['Unidad'] = df[c_m2].apply(lambda x: 'M' if parse_latam_number(x)>0 else 'S') if c_m2 else 'S'
            
        df['Tren-Km'] = 43.13 * df['Unidad'].apply(lambda x: 2 if str(x).strip()=='M' else 1)
        df['Fecha_Op'] = fch_dt
        
        col_ref = next((c for c in df.columns if ('PUERTO' in str(c).upper() or 'LIMACHE' in str(c).upper()) and 'Salida' in str(c) and '_min' in str(c)), None)
        if col_ref: df['Hora_Ref_Min'] = df[col_ref]
        
        diag["filas"] = len(df)
        return df, diag
    except Exception as e:
        diag["error"] = str(e); return pd.DataFrame(), diag

def procesar_carga_pasajeros(f, start_date, end_date):
    try:
        is_csv = f.name.lower().endswith('.csv')
        if is_csv:
            try: df = pd.read_csv(f, header=None, encoding='utf-8')
            except UnicodeDecodeError: 
                f.seek(0); df = pd.read_csv(f, header=None, encoding='latin-1')
        else:
            eu = "xlrd" if f.name.lower().endswith(".xls") else "openpyxl"
            df = pd.read_excel(f, engine=eu, header=None)
            
        h_idx = next((i for i in range(min(30, len(df))) if 'THDR' in str(df.iloc[i].values).upper() or 'VIAJE' in str(df.iloc[i].values).upper()), None)
        
        if h_idx is not None:
            f.seek(0)
            if is_csv:
                try: df = pd.read_csv(f, header=h_idx, encoding='utf-8')
                except UnicodeDecodeError:
                    f.seek(0); df = pd.read_csv(f, header=h_idx, encoding='latin-1')
            else:
                df = pd.read_excel(f, engine=eu, header=h_idx)
                
            df.columns = [str(c).strip() for c in df.columns]
            
            c_fecha = next((c for c in df.columns if 'FECHA' in str(c).upper() or 'DIA' in str(c).upper() or 'DATE' in str(c).upper()), None)
            if c_fecha:
                df['Fecha'] = pd.to_datetime(df[c_fecha], dayfirst=True, errors='coerce').dt.normalize()
                df = df.dropna(subset=['Fecha'])
                df = df[(df['Fecha'].dt.date >= start_date) & (df['Fecha'].dt.date <= end_date)]
            elif 'Fecha' in df.columns:
                df['Fecha'] = pd.to_datetime(df['Fecha'], dayfirst=True, errors='coerce').dt.normalize()
                df = df.dropna(subset=['Fecha'])
                df = df[(df['Fecha'].dt.date >= start_date) & (df['Fecha'].dt.date <= end_date)]
            
            if df.empty:
                return pd.DataFrame()
                
            c_tot = next((c for c in df.columns if 'TOTAL' in str(c).upper() and 'BORDO' in str(c).upper()), None)
            if c_tot:
                df['Total a Bordo'] = pd.to_numeric(df[c_tot], errors='coerce').fillna(0)
            elif 'Total a Bordo' in df.columns:
                df['Total a Bordo'] = pd.to_numeric(df['Total a Bordo'], errors='coerce').fillna(0)
                
            c_est_max = next((c for c in df.columns if 'ESTACI' in str(c).upper() and 'MAX' in _norm(c)), None)
            if c_est_max:
                df['Estación Máxima'] = df[c_est_max]
                
            return df
        return pd.DataFrame()
    except Exception:
        return pd.DataFrame()

def fig_perfil_velocidades():
    kms = [(s[0]+s[1])/2/1000 for s in SPEED_PROFILE]
    vels_n = [s[3] if s[3] > 0 else 0 for s in SPEED_PROFILE]
    vels_r = [s[4] if s[4] > 0 else 0 for s in SPEED_PROFILE]
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=kms, y=vels_n, mode='lines', name='Vel. Normal',
                              line=dict(color='#005195', width=2), fill='tozeroy', fillcolor='rgba(0,81,149,0.12)'))
    fig.add_trace(go.Scatter(x=kms, y=vels_r, mode='lines', name='Vel. RM', line=dict(color='#E85500', width=1.5, dash='dot')))
    for est, km_est in zip(ESTACIONES, KM_ACUM):
        fig.add_vline(x=km_est, line_width=1, line_dash='dot', line_color='gray')
        fig.add_annotation(x=km_est, y=125, text=est[:3], showarrow=False, font=dict(size=8, color='#555'), textangle=-90)
    fig.update_layout(title='Perfil de velocidades — Vía 1', xaxis_title='km desde Puerto', yaxis_title='km/h', height=300, margin=dict(t=40, b=20))
    return fig

# --- 6. INTERFAZ Y SIDEBAR ---
with st.sidebar:
    st.header("📅 Filtro Global")
    dr = st.date_input("Rango", value=(date(2026,1,1), date(2026,1,31)))
    start_date, end_date = (dr[0],dr[1]) if isinstance(dr,tuple) and len(dr)==2 else (dr,dr)
    st.divider()
    
    def _badge(c): n=len(listar_archivos(c)); return f" ({n} guardados)" if n else ""
    
    f_v1         = st.file_uploader(f"1. THDR Vía 1{_badge(DATA_DIRS['v1'])}", accept_multiple_files=True)
    f_v2         = st.file_uploader(f"2. THDR Vía 2{_badge(DATA_DIRS['v2'])}", accept_multiple_files=True)
    f_umr        = st.file_uploader(f"3. UMR / Odómetros{_badge(DATA_DIRS['umr'])}", accept_multiple_files=True)
    f_seat_files = st.file_uploader(f"4. Energía SEAT{_badge(DATA_DIRS['seat'])}", accept_multiple_files=True)
    f_bill_files = st.file_uploader(f"5. Facturación y PRMTE{_badge(DATA_DIRS['bill'])}", accept_multiple_files=True)
    f_carga_v1   = st.file_uploader(f"6. Carga Pasajeros V1{_badge(DATA_DIRS['carga_v1'])}", accept_multiple_files=True)
    f_carga_v2   = st.file_uploader(f"7. Carga Pasajeros V2{_badge(DATA_DIRS['carga_v2'])}", accept_multiple_files=True)
    
    for _ul,_ca in [(f_v1,DATA_DIRS["v1"]),(f_v2,DATA_DIRS["v2"]),(f_umr,DATA_DIRS["umr"]),
                    (f_seat_files,DATA_DIRS["seat"]),(f_bill_files,DATA_DIRS["bill"]),
                    (f_carga_v1,DATA_DIRS["carga_v1"]),(f_carga_v2,DATA_DIRS["carga_v2"])]:
        for uf in (_ul or []):
            dest=os.path.join(_ca,uf.name)
            if not os.path.exists(dest): guardar_archivo(uf,_ca)
            
    st.divider()
    with st.expander("🗂️ Archivos guardados"):
        _labels={"v1":"Vía 1","v2":"Vía 2","umr":"UMR","seat":"SEAT","bill":"Facturación", "carga_v1":"Pasajeros V1", "carga_v2":"Pasajeros V2"}
        for _key,_carpeta in DATA_DIRS.items():
            _arch=listar_archivos(_carpeta)
            if _arch:
                st.markdown(f"**{_labels[_key]}** — {len(_arch)} archivo(s)")
                for _a in _arch:
                    ca2,cb2=st.columns([5,1]); ca2.caption(os.path.basename(_a))
                    if cb2.button("🗑️",key=f"del_{_a}"): os.remove(_a); st.rerun()
            else: st.caption(f"{_labels[_key]}: sin archivos")

# Combinar uploads
f_v1_all   = combinar_fuentes(f_v1,         DATA_DIRS["v1"])
f_v2_all   = combinar_fuentes(f_v2,         DATA_DIRS["v2"])
f_umr_all  = combinar_fuentes(f_umr,        DATA_DIRS["umr"])
f_seat_all = combinar_fuentes(f_seat_files, DATA_DIRS["seat"])
f_bill_all = combinar_fuentes(f_bill_files, DATA_DIRS["bill"])
f_carga_v1_all = combinar_fuentes(f_carga_v1, DATA_DIRS["carga_v1"])
f_carga_v2_all = combinar_fuentes(f_carga_v2, DATA_DIRS["carga_v2"])

# --- 7. LÓGICA DE CACHÉ Y PROCESAMIENTO ---
_CACHE_VERSION = "v18_pax_local_time"
_cache_key = (_CACHE_VERSION, str(start_date), str(end_date),
              tuple(sorted(f.name for f in f_v1_all)), tuple(sorted(f.name for f in f_v2_all)),
              tuple(sorted(f.name for f in f_umr_all)), tuple(sorted(f.name for f in f_seat_all)),
              tuple(sorted(f.name for f in f_bill_all)),
              tuple(sorted(f.name for f in f_carga_v1_all)), tuple(sorted(f.name for f in f_carga_v2_all)))
              
_hay_archivos = any([f_v1_all,f_v2_all,f_umr_all,f_seat_all,f_bill_all,f_carga_v1_all,f_carga_v2_all])
_recalcular   = st.session_state.get('_cache_key') != _cache_key

df_ops=pd.DataFrame(); df_thdr_v1=pd.DataFrame(); df_thdr_v2=pd.DataFrame()
df_carga_v1=pd.DataFrame(); df_carga_v2=pd.DataFrame()
all_ops,all_tr,all_seat,all_fact_full,all_prmte_full=[],[],[],[],[]
_errores_proc={}

if _hay_archivos and not _recalcular and 'df_ops' in st.session_state:
    df_ops=st.session_state['df_ops']
    df_thdr_v1=st.session_state['df_thdr_v1']
    df_thdr_v2=st.session_state['df_thdr_v2']
    all_tr=st.session_state['all_tr']
    all_seat=st.session_state['all_seat']
    all_fact_full=st.session_state['all_fact_full']
    all_prmte_full=st.session_state['all_prmte_full']
    df_carga_v1=st.session_state.get('df_carga_v1', pd.DataFrame())
    df_carga_v2=st.session_state.get('df_carga_v2', pd.DataFrame())

elif _hay_archivos and _recalcular:
    if f_umr_all:
        for f in f_umr_all:
            try:
                eu="xlrd" if f.name.lower().endswith(".xls") else "openpyxl"
                xl=pd.ExcelFile(f,engine=eu)
                for sn in xl.sheet_names:
                    f.seek(0); df_raw=pd.read_excel(f,sheet_name=sn,header=None,engine=eu)
                    h_r=next((i for i in range(min(100,len(df_raw))) if any(k in str(df_raw.iloc[i].tolist()).upper() for k in ['FECHA','ODO','KILOM'])),None)
                    if h_r is not None:
                        f.seek(0); df_p=pd.read_excel(f,sheet_name=sn,header=h_r,engine=eu)
                        df_p.columns=[str(c).upper().replace('Ó','O').strip() for c in df_p.columns]
                        c_f=next((c for c in df_p.columns if 'FECHA' in c),None)
                        c_o=next((c for c in df_p.columns if 'ODO' in c),None)
                        c_t=next((c for c in df_p.columns if 'KM' in c),None)
                        if c_f and c_o:
                            df_p['_dt']=pd.to_datetime(df_p[c_f],errors='coerce').dt.normalize()
                            mask=(df_p['_dt'].dt.date>=start_date)&(df_p['_dt'].dt.date<=end_date)
                            df_valid = df_p[mask].dropna(subset=['_dt']).copy()
                            for _, r in df_valid.iterrows():
                                all_ops.append({"Fecha":r['_dt'],"Tipo Día":get_tipo_dia(r['_dt'].date()),
                                                "Odómetro [km]":parse_latam_number(r[c_o]),
                                                "Tren-Km [km]":parse_latam_number(r[c_t]) if c_t else 0.0})
                    if any(k in sn.upper() for k in ['KIL','ODO']):
                        for i in range(len(df_raw)-2):
                            for j in range(1,len(df_raw.columns)):
                                v_f=pd.to_datetime(df_raw.iloc[i,j],errors='coerce')
                                if pd.notna(v_f) and start_date<=v_f.date()<=end_date:
                                    for k in range(i+3,min(i+50,len(df_raw))):
                                        t=str(df_raw.iloc[k,0]).strip().upper()
                                        if re.match(r'^(M|XM)',t):
                                            all_tr.append({"Tren":t,"Fecha":v_f.normalize(),"Valor":parse_latam_number(df_raw.iloc[k,j])})
            except Exception as e: _errores_proc[f.name]=f"UMR: {e}"
            
    if f_seat_all:
        for f in f_seat_all:
            try:
                es="xlrd" if f.name.lower().endswith(".xls") else "openpyxl"
                df_s=pd.read_excel(f,header=None,engine=es)
                for i in range(len(df_s)):
                    fs=pd.to_datetime(df_s.iloc[i,1],errors='coerce')
                    if pd.notna(fs) and start_date<=fs.normalize().date()<=end_date:
                        all_seat.append({"Fecha":fs.normalize(),"E_Total":parse_latam_number(df_s.iloc[i,3]),
                                         "E_Tr":parse_latam_number(df_s.iloc[i,5]),"E_12":parse_latam_number(df_s.iloc[i,7])})
            except Exception as e: _errores_proc[f.name]=f"SEAT: {e}"
            
    if f_bill_all:
        for f in f_bill_all:
            try:
                eb="xlrd" if f.name.lower().endswith(".xls") else "openpyxl"
                f.seek(0); xl=pd.ExcelFile(f,engine=eb)
                for sn in xl.sheet_names:
                    if sn=="Consumo Factura":
                        f.seek(0); df_ff=pd.read_excel(f,sheet_name=sn,engine=eb)
                        c_f=next((c for c in df_ff.columns if 'FECHA' in str(c).upper()),df_ff.columns[0])
                        c_v=next((c for c in df_ff.columns if 'CONSUMO' in str(c).upper() or 'VALOR' in str(c).upper()),df_ff.columns[1])
                        df_ff['dt']=pd.to_datetime(df_ff[c_f],errors='coerce')
                        df_valid = df_ff.dropna(subset=['dt'])
                        for _,r in df_valid.iterrows():
                            if "TOTAL" in str(r[c_f]).upper(): continue
                            v=abs(parse_latam_number(r[c_v]))
                            all_fact_full.append({"Fecha":r['dt'].normalize(),"Hora":f"{r['dt'].hour:02d}:00",
                                                  "15min":f"{r['dt'].hour:02d}:{(r['dt'].minute//15)*15:02d}","Consumo":v})
                                                  
                    if 'PRMTE' in sn.upper():
                        f.seek(0); df_pd_raw=pd.read_excel(f,sheet_name=sn,header=None,engine=eb)
                        h=next((i for i in range(min(20,len(df_pd_raw))) if any(k in str(df_pd_raw.iloc[i]).upper() for k in ['AÑO','ANO','YEAR'])),0)
                        f.seek(0); df_pd=pd.read_excel(f,sheet_name=sn,header=h,engine=eb).dropna(how='all')
                        c_anio=next((c for c in df_pd.columns if str(c).upper().replace('Ñ','N').startswith('AN')),None)
                        c_mes=next((c for c in df_pd.columns if str(c).upper().startswith('MES')),None)
                        c_dia=next((c for c in df_pd.columns if str(c).upper().startswith('DIA')),None)
                        c_hora=next((c for c in df_pd.columns if str(c).upper()=='HORA'),None)
                        c_ini=next((c for c in df_pd.columns if 'INICIO' in str(c).upper()),None)
                        
                        if not (c_anio and c_mes and c_dia and c_hora):
                            raise ValueError(f"Columnas de fecha no encontradas: {list(df_pd.columns)}")
                            
                        def _build_ts(r):
                            try:
                                min_=int(r[c_ini]) if c_ini and not pd.isna(r[c_ini]) else 0
                                return pd.Timestamp(year=int(r[c_anio]),month=int(r[c_mes]),day=int(r[c_dia]),hour=int(r[c_hora]),minute=min_)
                            except Exception: return pd.NaT
                            
                        df_pd['ts']=df_pd.apply(_build_ts,axis=1)
                        cols_retiro=[c for c in df_pd.columns if 'Retiro_Energia_Activa' in str(c)]
                        if not cols_retiro:
                            cols_retiro=[c for c in df_pd.columns if 'RETIRO' in str(c).upper() or ('ENERGIA' in str(c).upper() and 'ACTIVA' in str(c).upper())]
                            
                        for _,r in df_pd.dropna(subset=['ts']).iterrows():
                            ts=r['ts']
                            if pd.isna(ts) or not (start_date<=ts.date()<=end_date): continue
                            consumo=sum(parse_latam_number(r.get(c,0)) for c in cols_retiro)
                            all_prmte_full.append({"Fecha":ts.normalize(),"Hora":f"{ts.hour:02d}:00","15min":f"{ts.hour:02d}:{ts.minute:02d}","Consumo":consumo})
            except Exception as e: _errores_proc[f.name]=f"Factura/PRMTE: {e}"
            
    if _errores_proc: st.session_state['_errores_proc']=_errores_proc

    if all_ops:
        df_ops=pd.DataFrame(all_ops).groupby("Fecha").agg({"Odómetro [km]":"sum","Tren-Km [km]":"sum"}).reset_index()
        df_f_d=(pd.DataFrame(all_fact_full).groupby("Fecha")["Consumo"].sum().reset_index().rename(columns={"Consumo":"E_Fact"}) if all_fact_full else pd.DataFrame(columns=["Fecha","E_Fact"]))
        df_p_d=(pd.DataFrame(all_prmte_full).groupby("Fecha")["Consumo"].sum().reset_index().rename(columns={"Consumo":"E_Prmte"}) if all_prmte_full else pd.DataFrame(columns=["Fecha","E_Prmte"]))
        df_s_d=(pd.DataFrame(all_seat).groupby("Fecha").agg({"E_Total":"sum","E_Tr":"sum","E_12":"sum"}).reset_index().rename(columns={"E_Total":"E_Seat_T","E_Tr":"E_Seat_Tr","E_12":"E_Seat_12"}) if all_seat else pd.DataFrame(columns=["Fecha","E_Seat_T","E_Seat_Tr","E_Seat_12"]))
        
        for dff in [df_ops,df_f_d,df_p_d,df_s_d]: dff['Fecha']=pd.to_datetime(dff['Fecha']).dt.normalize()
        df_ops=(df_ops.merge(df_f_d,on="Fecha",how="left").merge(df_p_d,on="Fecha",how="left").merge(df_s_d,on="Fecha",how="left").fillna(0))
        
        df_ops['Tipo Día'] = df_ops['Fecha'].apply(lambda x: get_tipo_dia(x.date() if pd.notna(x) else None))
        df_ops['Nombre Feriado'] = df_ops['Fecha'].apply(lambda x: obtener_nombre_feriado(x.date() if pd.notna(x) else None))
        df_ops['Fecha (ES)'] = df_ops['Fecha'].apply(obtener_fecha_es)

        def jerarquia(row):
            if row['E_Fact']>0:    tot,src=row['E_Fact'],"Factura"
            elif row['E_Prmte']>0:  tot,src=row['E_Prmte'],"PRMTE"
            elif row['E_Seat_T']>0: tot,src=row['E_Seat_T'],"SEAT"
            else: return 0,0,0,0,0,"N/A"
            r_tr=row['E_Seat_Tr']/row['E_Seat_T'] if row['E_Seat_T']>0 else 0.85
            r_12=row['E_Seat_12']/row['E_Seat_T'] if row['E_Seat_T']>0 else 0.15
            return tot,tot*r_tr,tot*r_12,r_tr*100,r_12*100,src
            
        df_ops[['E_Total','E_Tr','E_12','% Tracción','% 12 kV','Fuente']]=df_ops.apply(jerarquia,axis=1,result_type='expand')
        df_ops['IDE (kWh/km)']=df_ops.apply(lambda r: r['E_Tr']/r['Odómetro [km]'] if r['Odómetro [km]']>0 else 0,axis=1)
        df_ops['UMR (%)'] = df_ops.apply(lambda r: (r['Tren-Km [km]'] / r['Odómetro [km]']) * 100 if r['Odómetro [km]'] > 0 else 0, axis=1)

    diag_thdr=[]
    if f_v1_all:
        r1=[procesar_thdr_eficiente(f,start_date,end_date) for f in f_v1_all]
        diag_thdr+=[r[1] for r in r1]
        p1=[r[0] for r in r1 if not r[0].empty]
        df_thdr_v1=pd.concat(p1,ignore_index=True) if p1 else pd.DataFrame()
    if f_v2_all:
        r2=[procesar_thdr_eficiente(f,start_date,end_date) for f in f_v2_all]
        diag_thdr+=[r[1] for r in r2]
        p2=[r[0] for r in r2 if not r[0].empty]
        df_thdr_v2=pd.concat(p2,ignore_index=True) if p2 else pd.DataFrame()
    if diag_thdr: st.session_state['diag_thdr']=diag_thdr
    
    if f_carga_v1_all:
        p_cv1 = [procesar_carga_pasajeros(f, start_date, end_date) for f in f_carga_v1_all]
        p_cv1 = [d for d in p_cv1 if not d.empty]
        df_carga_v1 = pd.concat(p_cv1, ignore_index=True) if p_cv1 else pd.DataFrame()

    if f_carga_v2_all:
        p_cv2 = [procesar_carga_pasajeros(f, start_date, end_date) for f in f_carga_v2_all]
        p_cv2 = [d for d in p_cv2 if not d.empty]
        df_carga_v2 = pd.concat(p_cv2, ignore_index=True) if p_cv2 else pd.DataFrame()

    if not df_ops.empty:
        if not df_thdr_v1.empty or not df_thdr_v2.empty:
            s1 = df_thdr_v1.groupby('Fecha_Op').size().reset_index(name='V1_S') if not df_thdr_v1.empty else pd.DataFrame(columns=['Fecha_Op', 'V1_S'])
            s2 = df_thdr_v2.groupby('Fecha_Op').size().reset_index(name='V2_S') if not df_thdr_v2.empty else pd.DataFrame(columns=['Fecha_Op', 'V2_S'])
            df_servicios = pd.merge(s1, s2, on='Fecha_Op', how='outer').fillna(0)
            df_servicios['Servicios'] = df_servicios['V1_S'] + df_servicios['V2_S']
            df_servicios = df_servicios.rename(columns={'Fecha_Op': 'Fecha'})
            df_servicios['Fecha'] = pd.to_datetime(df_servicios['Fecha']).dt.normalize()
            df_ops = df_ops.merge(df_servicios[['Fecha', 'Servicios']], on='Fecha', how='left').fillna({'Servicios': 0})
        else:
            df_ops['Servicios'] = 0

        if not df_carga_v1.empty or not df_carga_v2.empty:
            p1 = df_carga_v1.groupby('Fecha')['Total a Bordo'].sum().reset_index(name='PAX_V1') if (not df_carga_v1.empty and 'Fecha' in df_carga_v1.columns) else pd.DataFrame(columns=['Fecha', 'PAX_V1'])
            p2 = df_carga_v2.groupby('Fecha')['Total a Bordo'].sum().reset_index(name='PAX_V2') if (not df_carga_v2.empty and 'Fecha' in df_carga_v2.columns) else pd.DataFrame(columns=['Fecha', 'PAX_V2'])
            df_pax = pd.merge(p1, p2, on='Fecha', how='outer').fillna(0)
            df_pax['PAX'] = df_pax['PAX_V1'] + df_pax['PAX_V2']
            df_pax['Fecha'] = pd.to_datetime(df_pax['Fecha']).dt.normalize()
            df_ops = df_ops.merge(df_pax[['Fecha', 'PAX']], on='Fecha', how='left').fillna({'PAX': 0})
        else:
            df_ops['PAX'] = 0

    st.session_state.update({'df_ops':df_ops,'df_thdr_v1':df_thdr_v1,'df_thdr_v2':df_thdr_v2,
                              'all_tr':all_tr,'all_seat':all_seat,'all_fact_full':all_fact_full,
                              'all_prmte_full':all_prmte_full,'_cache_key':_cache_key,
                              'df_carga_v1':df_carga_v1, 'df_carga_v2':df_carga_v2})

# --- 8. TABS DE VISUALIZACIÓN ---
tabs=st.tabs(["📊 Resumen","📑 Operaciones","📑 Trenes","⚡ Energía","⚖️ Perfil Horario & Anomalías",
              "📈 Regresión","🚨 Atípicos","📋 THDR","🔬 Análisis Multivariante", "👥 Pasajeros", "📝 Informe Ejecutivo", "🩺 Diagnóstico de Causas"])

with tabs[0]:
    _ep=st.session_state.get('_errores_proc',{})
    if _ep:
        with st.expander(f"⚠️ {len(_ep)} archivo(s) con error",expanded=True):
            for _n,_m in _ep.items(): st.error(f"**{_n}**: {_m}")
    if not df_ops.empty:
        st.markdown("### 🎛️ Filtros de Resumen")
        filtro_dia = st.multiselect(
            "Tipo de Jornada:",
            options=["L", "S", "D/F"],
            default=["L", "S", "D/F"],
            format_func=lambda x: {"L": "Laboral (L)", "S": "Sábado (S)", "D/F": "Domingo y Festivo (D/F)"}.get(x, x)
        )
        
        df_resumen = df_ops[df_ops['Tipo Día'].isin(filtro_dia)]
        
        if 'drilldown_date' not in st.session_state:
            st.session_state.drilldown_date = None
            
        if st.session_state.drilldown_date is not None:
            df_resumen = df_resumen[df_resumen['Fecha'] == st.session_state.drilldown_date]
        
        if df_resumen.empty:
            st.warning("No hay datos operacionales para los filtros seleccionados.")
        else:
            st.markdown("### 🚄 DATOS OPERACIONALES")
            
            if st.session_state.drilldown_date is not None:
                c_info, c_btn = st.columns([4, 1])
                c_info.info(f"🔍 **Modo Detalle Activo:** Estás viendo los datos exclusivos del día **{st.session_state.drilldown_date.strftime('%d-%m-%Y')}**.")
                if c_btn.button("❌ Quitar filtro de día", use_container_width=True):
                    st.session_state.drilldown_date = None
                    st.rerun()
            
            hover_config = {}
            if 'Fecha (ES)' in df_resumen.columns:
                hover_config['Fecha'] = False       
                hover_config['Fecha (ES)'] = True   
            else:
                hover_config['Fecha'] = True        
                
            for col in ['Tipo Día', 'Nombre Feriado']:
                if col in df_resumen.columns: 
                    hover_config[col] = True
            
            c_chart_s, c_card_s, c_chart_p, c_card_p = st.columns([2.5, 1, 2.5, 1]) 
            
            with c_chart_s:
                fig_serv = px.bar(df_resumen, x='Fecha', y='Servicios', 
                                  color_discrete_sequence=["#005195"],
                                  hover_data=hover_config, title="Servicios Programados")
                fig_serv.update_traces(texttemplate='%{y:,.0f}', textposition='inside', insidetextanchor='middle', textfont=dict(color='white', size=13))
                fig_serv.update_layout(margin=dict(t=50, b=0, l=0, r=0), title=dict(font=dict(size=15), automargin=True),
                                       bargap=0.15, uniformtext=dict(minsize=9, mode='hide'))
                ev_serv = st.plotly_chart(fig_serv, use_container_width=True, config={'locale': 'es'}, on_select="rerun", key="chart_serv")
                
            with c_card_s:
                st.markdown("<br><br>", unsafe_allow_html=True)
                st.metric("Total Servicios", f"{int(df_resumen['Servicios'].sum()):,}")

            with c_chart_p:
                fig_pax = px.bar(df_resumen, x='Fecha', y='PAX', 
                                  color_discrete_sequence=["#E85500"], 
                                  hover_data=hover_config, title="Pasajeros Transportados (PAX)")
                fig_pax.update_traces(texttemplate='%{y:,.0f}', textposition='inside', insidetextanchor='middle', textfont=dict(color='white', size=13))
                fig_pax.update_layout(margin=dict(t=50, b=0, l=0, r=0), title=dict(font=dict(size=15), automargin=True),
                                      bargap=0.15, uniformtext=dict(minsize=9, mode='hide'))
                ev_pax = st.plotly_chart(fig_pax, use_container_width=True, config={'locale': 'es'}, on_select="rerun", key="chart_pax")
                
            with c_card_p:
                st.markdown("<br><br>", unsafe_allow_html=True)
                st.metric("Total PAX", f"{int(df_resumen['PAX'].sum()):,}")

            st.divider()
            
            c_chart_k, c_card_k, c_chart_u, c_card_u = st.columns([2.5, 1, 2.5, 1]) 
            
            with c_chart_k:
                fig_km = px.bar(df_resumen, x='Fecha', y=['Odómetro [km]', 'Tren-Km [km]'], 
                                barmode='group',
                                color_discrete_map={'Odómetro [km]': '#005195', 'Tren-Km [km]': '#66A5D9'}, 
                                hover_data=hover_config, title="Kilometraje (Odómetro vs Tren-Km)")
                fig_km.update_traces(texttemplate='%{y:,.2f}', textposition='inside', insidetextanchor='middle', textangle=-90, textfont=dict(color='white', size=11))
                fig_km.update_layout(margin=dict(t=50, b=0, l=0, r=0), title=dict(font=dict(size=15), automargin=True),
                                     legend=dict(title="", orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                                     bargap=0.15, uniformtext=dict(minsize=8, mode='hide'))
                ev_km = st.plotly_chart(fig_km, use_container_width=True, config={'locale': 'es'}, on_select="rerun", key="chart_km")
                
            with c_card_k:
                st.markdown("<br>", unsafe_allow_html=True)
                st.metric("Odómetro Total", f"{df_resumen['Odómetro [km]'].sum():,.2f} km")
                st.metric("Tren-Km Total", f"{df_resumen['Tren-Km [km]'].sum():,.2f} km")

            with c_chart_u:
                fig_umr = px.bar(df_resumen, x='Fecha', y='UMR (%)', 
                                  color_discrete_sequence=["#E85500"], 
                                  hover_data=hover_config, title="Tasa Acoplamiento (UMR %)")
                fig_umr.update_traces(texttemplate='%{y:,.2f}%', textposition='inside', insidetextanchor='middle', textfont=dict(color='white', size=13))
                fig_umr.update_layout(margin=dict(t=50, b=0, l=0, r=0), title=dict(font=dict(size=15), automargin=True),
                                      bargap=0.15, uniformtext=dict(minsize=9, mode='hide'))
                ev_umr = st.plotly_chart(fig_umr, use_container_width=True, config={'locale': 'es'}, on_select="rerun", key="chart_umr")
                
            with c_card_u:
                st.markdown("<br><br>", unsafe_allow_html=True)
                tot_tren_km = df_resumen['Tren-Km [km]'].sum()
                tot_odometro = df_resumen['Odómetro [km]'].sum()
                umr_global = (tot_tren_km / tot_odometro * 100) if tot_odometro > 0 else 0
                st.metric("Tasa UMR Global", f"{umr_global:,.2f} %")
                
            st.divider() 
            
            c_chart_e, c_card_e, c_chart_i, c_card_i = st.columns([2.5, 1, 2.5, 1]) 
            
            df_plot_ener = df_resumen.rename(columns={'E_Tr': 'Tracción', 'E_12': 'Baja Tensión'})
            
            with c_chart_e:
                fig_ener = px.bar(df_plot_ener, x='Fecha', y=['Tracción', 'Baja Tensión'], 
                                  barmode='stack',
                                  color_discrete_map={'Tracción': '#E85500', 'Baja Tensión': '#005195'},
                                  hover_data=hover_config, title="Consumo Energético (kWh)")
                fig_ener.update_traces(texttemplate='%{y:,.2f}', textposition='inside', insidetextanchor='middle', textangle=-90, textfont=dict(color='white', size=11)) 
                fig_ener.update_layout(margin=dict(t=50, b=0, l=0, r=0), title=dict(font=dict(size=15), automargin=True),
                                     legend=dict(title="", orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                                     bargap=0.15, uniformtext=dict(minsize=8, mode='hide'))
                ev_ener = st.plotly_chart(fig_ener, use_container_width=True, config={'locale': 'es'}, on_select="rerun", key="chart_ener")
                
            with c_card_e:
                st.markdown("<br>", unsafe_allow_html=True)
                st.metric("Total Tracción", f"{df_plot_ener['Tracción'].sum():,.2f} kWh")
                st.metric("Total Baja Tensión", f"{df_plot_ener['Baja Tensión'].sum():,.2f} kWh")

            with c_chart_i:
                fig_ide_bar = px.bar(df_resumen, x='Fecha', y='IDE (kWh/km)', 
                                  color_discrete_sequence=["#E85500"], 
                                  hover_data=hover_config, title="Desempeño Energético (IDE)")
                fig_ide_bar.update_traces(texttemplate='%{y:,.2f}', textposition='inside', insidetextanchor='middle', textfont=dict(color='white', size=13))
                fig_ide_bar.update_layout(margin=dict(t=50, b=0, l=0, r=0), title=dict(font=dict(size=15), automargin=True),
                                          bargap=0.15, uniformtext=dict(minsize=9, mode='hide'))
                ev_ide_bar = st.plotly_chart(fig_ide_bar, use_container_width=True, config={'locale': 'es'}, on_select="rerun", key="chart_ide")
                
            with c_card_i:
                st.markdown("<br><br>", unsafe_allow_html=True)
                tot_traccion_real = df_resumen['E_Tr'].sum()
                ide_global = (tot_traccion_real / tot_odometro) if tot_odometro > 0 else 0
                st.metric("IDE Global", f"{ide_global:,.2f} kWh/km")

            chart_events = [ev_serv, ev_pax, ev_km, ev_umr, ev_ener, ev_ide_bar]
            
            for ev in chart_events:
                if ev and isinstance(ev, dict) and ev.get('selection') and ev['selection'].get('points'):
                    clicked_x = ev['selection']['points'][0].get('x')
                    if clicked_x:
                        try:
                            clicked_dt = pd.to_datetime(clicked_x).normalize()
                            if st.session_state.drilldown_date != clicked_dt:
                                st.session_state.drilldown_date = clicked_dt
                                st.rerun()
                        except Exception:
                            pass
            
    else: st.info("📂 Sube archivos desde el panel lateral para ver el resumen.")

with tabs[1]:
    if not df_ops.empty:
        dv=df_ops.copy()
        dv['Fecha'] = dv['Fecha'].dt.strftime('%Y-%m-%d')
        st.write("### Datos Consolidados de Operaciones y Energía")
        st.dataframe(make_columns_unique(dv), use_container_width=True)
    else: st.info("No hay datos de operaciones en el rango seleccionado.")

with tabs[2]:
    if all_tr:
        st.write("### Detalle por Unidad (Tren)")
        df_tr = pd.DataFrame(all_tr)
        df_tr['Fecha'] = pd.to_datetime(df_tr['Fecha']).dt.strftime('%Y-%m-%d')
        st.dataframe(make_columns_unique(df_tr), use_container_width=True)
    else: st.info("No hay datos detallados de trenes cargados.")

with tabs[3]:
    if not df_ops.empty and 'E_Total' in df_ops.columns and df_ops['E_Total'].sum() > 0:
        st.write("### Consumo de Energía por Tipo")
        fig_e = go.Figure()
        fig_e.add_trace(go.Bar(x=df_ops['Fecha'], y=df_ops['E_Tr'], name='Tracción', marker_color='#E85500'))
        fig_e.add_trace(go.Bar(x=df_ops['Fecha'], y=df_ops['E_12'], name='12 kV', marker_color='#005195'))
        fig_e.update_layout(barmode='stack', title="Consumo Energético: Tracción vs Servicios Auxiliares (12kV)", xaxis_title="Fecha", yaxis_title="Consumo (kWh)")
        st.plotly_chart(fig_e, use_container_width=True)
        
        st.write("### Desglose Detallado")
        dv_ener = df_ops[['Fecha', 'E_Total', 'E_Tr', 'E_12', '% Tracción', '% 12 kV', 'Fuente']].copy()
        dv_ener['Fecha'] = dv_ener['Fecha'].dt.strftime('%Y-%m-%d')
        st.dataframe(make_columns_unique(dv_ener), use_container_width=True)
    else: st.info("No hay datos de energía procesados (Facturación, PRMTE o SEAT).")

with tabs[4]:
    if all_prmte_full:
        st.markdown("### 🔍 Análisis Granular de Consumo (15 min y Horario)")
        st.markdown("Este panel permite auditar el comportamiento eléctrico de la flota detectando consumos parásitos (nocturnos) y picos de demanda críticos.")
        
        df_prmte = pd.DataFrame(all_prmte_full)
        df_prmte['Fecha'] = pd.to_datetime(df_prmte['Fecha']).dt.date
        
        # --- 1. AUDITORÍA DE CARGA BASE (CONSUMO NOCTURNO DINÁMICO) ---
        st.markdown("#### 🌙 Auditoría de Consumo Nocturno Dinámico (Carga Base)")
        st.caption("Horario evaluado: 00:00 a 06:00 (Laborales), a 07:00 (Sábados) y a 08:00 (Dom/Fest)")
        
        mapa_tipo_prmte = df_ops.set_index(df_ops['Fecha'].dt.date)['Tipo Día'].to_dict()
        df_prmte['Tipo Día'] = df_prmte['Fecha'].map(mapa_tipo_prmte)
        df_prmte['Tipo Día'] = df_prmte['Tipo Día'].fillna(df_prmte['Fecha'].apply(lambda x: get_tipo_dia(x)))
        df_prmte['Hora_n'] = df_prmte['Hora'].str.slice(0, 2).astype(int)
        
        def is_noche_tab4(row):
            limite = 6 if row['Tipo Día'] == 'L' else (7 if row['Tipo Día'] == 'S' else 8)
            return 0 <= row['Hora_n'] < limite
            
        df_prmte['Es_Noche'] = df_prmte.apply(is_noche_tab4, axis=1)
        df_noche = df_prmte[df_prmte['Es_Noche']]
        
        if not df_noche.empty:
            consumo_noche_diario = df_noche.groupby('Fecha')['Consumo'].sum().reset_index()
            promedio_noche = consumo_noche_diario['Consumo'].mean()
            
            c_noct1, c_noct2 = st.columns([3, 1])
            with c_noct1:
                fig_noche = px.line(consumo_noche_diario, x='Fecha', y='Consumo', markers=True,
                                    title="Consumo Total Nocturno por Día (kWh)",
                                    color_discrete_sequence=["#1f77b4"])
                fig_noche.add_hline(y=promedio_noche, line_dash="dash", line_color="red", annotation_text="Promedio Base")
                anomalos_noche = consumo_noche_diario[consumo_noche_diario['Consumo'] > promedio_noche * 1.2]
                if not anomalos_noche.empty:
                    fig_noche.add_trace(go.Scatter(x=anomalos_noche['Fecha'], y=anomalos_noche['Consumo'],
                                                   mode='markers', marker=dict(color='red', size=12, symbol='x'),
                                                   name='Posible Tren Encendido'))
                fig_noche.update_layout(margin=dict(t=40, b=0, l=0, r=0))
                st.plotly_chart(fig_noche, use_container_width=True, config={'locale': 'es'})
            
            with c_noct2:
                st.markdown("<br>", unsafe_allow_html=True)
                st.metric("Consumo Nocturno Promedio", f"{promedio_noche:,.0f} kWh/noche")
                st.info("💡 **Insight:** Picos en este gráfico indican que los trenes no fueron apagados correctamente en cocheras, manteniendo equipos auxiliares/climatización operando de forma 'vampira'.")
        
        st.divider()

        # --- 2. MAPA DE CALOR DE 15 MINUTOS ---
        st.markdown("#### 🔥 Mapa de Calor: Consumo cada 15 Minutos")
        st.markdown("Identifica fácilmente a qué hora y qué día se producen los picos de demanda (colores cálidos) o valles (colores fríos).")
        
        matriz_15m = df_prmte.groupby(['Fecha', '15min'])['Consumo'].sum().reset_index()
        
        fig_heat = go.Figure(data=go.Heatmap(
            z=matriz_15m['Consumo'],
            x=matriz_15m['15min'],
            y=matriz_15m['Fecha'],
            colorscale='Turbo',
            hoverongaps=False,
            hovertemplate='Día: %{y}<br>Hora: %{x}<br>Consumo: %{z:,.1f} kWh<extra></extra>'
        ))
        
        fig_heat.update_layout(
            xaxis_title="Franja de 15 Minutos",
            yaxis_title="Fecha",
            height=500,
            margin=dict(t=20, b=50, l=0, r=0)
        )
        st.plotly_chart(fig_heat, use_container_width=True)
        
        st.divider()

        # --- 3. CURVA DE DEMANDA HORARIA (PERFIL ESTADÍSTICO) ---
        st.markdown("#### 📈 Curva de Demanda Promedio y Tolerancia Estadística")
        
        df_hr_stats = df_prmte.groupby('Hora')['Consumo'].agg(['mean', 'std']).reset_index()
        df_hr_stats['upper_band'] = df_hr_stats['mean'] + (1.5 * df_hr_stats['std'])
        df_hr_stats['lower_band'] = df_hr_stats['mean'] - (1.5 * df_hr_stats['std'])
        df_hr_stats['lower_band'] = df_hr_stats['lower_band'].clip(lower=0)
        
        fig_curva = go.Figure()
        
        fig_curva.add_trace(go.Scatter(
            x=pd.concat([df_hr_stats['Hora'], df_hr_stats['Hora'][::-1]]),
            y=pd.concat([df_hr_stats['upper_band'], df_hr_stats['lower_band'][::-1]]),
            fill='toself',
            fillcolor='rgba(0, 81, 149, 0.2)',
            line=dict(color='rgba(255,255,255,0)'),
            name='Rango Normal Operativo (±1.5σ)'
        ))
        
        fig_curva.add_trace(go.Scatter(
            x=df_hr_stats['Hora'], y=df_hr_stats['mean'],
            line=dict(color='#005195', width=3),
            mode='lines+markers',
            name='Consumo Promedio'
        ))
        
        fig_curva.update_layout(xaxis_title="Hora del Día", yaxis_title="Consumo (kWh)", hovermode="x unified",
                                margin=dict(t=30, b=0, l=0, r=0))
        st.plotly_chart(fig_curva, use_container_width=True)
        
        st.info("💡 **Peak Shaving:** El área sombreada celeste representa el 'consumo normal' histórico de la flota a esa hora. Si un día el consumo rompe la barrera superior, podría generar altos **cargos por potencia máxima**. Asegúrate de que los despachos de trenes (THDR) no coincidan exactamente en el mismo segundo para aplanar esta curva.")
        
    else: 
        st.info("Se necesita cargar el archivo de **PRMTE (Energía cada 15 min)** para habilitar el Centro de Control de Anomalías.")

with tabs[5]:
    if not df_ops.empty and df_ops['E_Tr'].sum() > 0:
        st.write("### Relación entre Kilometraje y Consumo de Tracción")
        df_reg = df_ops.dropna(subset=["Odómetro [km]", "E_Tr"])
        df_reg = df_reg[(df_reg["Odómetro [km]"] > 0) & (df_reg["E_Tr"] > 0)].sort_values("Odómetro [km]")
        
        fig_reg = px.scatter(df_reg, x="Odómetro [km]", y="E_Tr", hover_data=["Fecha"], title="Regresión: Odómetro vs Energía de Tracción", color_discrete_sequence=["#005195"])
        
        if len(df_reg) > 1:
            x_vals = df_reg["Odómetro [km]"].values
            y_vals = df_reg["E_Tr"].values
            slope, intercept = np.polyfit(x_vals, y_vals, 1)
            y_pred = slope * x_vals + intercept
            
            corr_matrix = np.corrcoef(x_vals, y_vals)
            r_squared = corr_matrix[0, 1] ** 2
            
            fig_reg.add_trace(go.Scatter(x=x_vals, y=y_pred, mode='lines', name=f'Ajuste Lineal (R²={r_squared:.4f})', line=dict(color='#E85500', width=2.5)))
            st.metric("Fórmula de Ajuste Matemático", f"y = {slope:.4f} * x + {intercept:.2f}", help="Ecuación lineal de regresión por mínimos cuadrados. y = Tracción (kWh), x = Distancia (km)")
        
        fig_reg.update_layout(xaxis_title="Odómetro Total (km)", yaxis_title="Energía de Tracción (kWh)")
        st.plotly_chart(fig_reg, use_container_width=True)
    else: st.info("No hay datos cruzados suficientes de kilometraje y consumo energético para calcular la regresión.")

with tabs[6]:
    if not df_ops.empty and df_ops['IDE (kWh/km)'].sum() > 0:
        st.write("### Detección de Valores Atípicos (IDE)")
        mean_ide = df_ops['IDE (kWh/km)'].mean()
        std_ide = df_ops['IDE (kWh/km)'].std()
        
        df_ops['Z-Score'] = (df_ops['IDE (kWh/km)'] - mean_ide) / std_ide
        df_ops['Es_Atípico'] = df_ops['Z-Score'].abs() > 2
        
        fig_out = go.Figure()
        fig_out.add_trace(go.Scatter(x=df_ops['Fecha'], y=df_ops['IDE (kWh/km)'], mode='markers',
                                     marker=dict(color=df_ops['Es_Atípico'].map({True: 'red', False: '#005195'}), size=8), name='IDE'))
        
        fig_out.add_hline(y=mean_ide, line_dash="dash", line_color="green", annotation_text="Media")
        fig_out.add_hline(y=mean_ide + 2*std_ide, line_dash="dot", line_color="orange", annotation_text="+2 Desv. Est.")
        fig_out.add_hline(y=mean_ide - 2*std_ide, line_dash="dot", line_color="orange", annotation_text="-2 Desv. Est.")
        
        fig_out.update_layout(title="IDE Diario (Identificando días más allá de ±2 desviaciones estándar)", xaxis_title="Fecha", yaxis_title="IDE (kWh/km)")
        st.plotly_chart(fig_out, use_container_width=True)
        
        atipicos = df_ops[df_ops['Es_Atípico']][['Fecha', 'IDE (kWh/km)', 'Odómetro [km]', 'E_Tr']]
        if not atipicos.empty:
            st.warning("⚠️ Se han detectado los siguientes días con comportamiento anómalo en el consumo:")
            atipicos['Fecha'] = atipicos['Fecha'].dt.strftime('%Y-%m-%d')
            st.dataframe(atipicos, use_container_width=True)
        else: st.success("✅ No se detectaron valores atípicos significativos (Z-score > 2) en el periodo analizado.")
    else: st.info("No hay datos de IDE calculados para analizar.")

with tabs[7]:
    st.write("### Perfil de Velocidades Vía 1 y 2")
    st.plotly_chart(fig_perfil_velocidades(), use_container_width=True)

    c_v1, c_v2 = st.columns(2)
    with c_v1:
        st.write("#### Datos THDR Vía 1")
        if not df_thdr_v1.empty:
            st.dataframe(df_thdr_v1.head(50), use_container_width=True)
            st.caption("Mostrando hasta 50 registros.")
        else: st.info("No se ha cargado/procesado THDR Vía 1.")
            
    with c_v2:
        st.write("#### Datos THDR Vía 2")
        if not df_thdr_v2.empty:
            st.dataframe(df_thdr_v2.head(50), use_container_width=True)
            st.caption("Mostrando hasta 50 registros.")
        else: st.info("No se ha cargado/procesado THDR Vía 2.")

with tabs[8]:
    st.markdown("### 🔬 Análisis Multivariante: PAX vs Tiempos vs Energía")
    st.markdown("Este módulo utiliza estadística robusta para cruzar la demanda real con la fricción operativa (Tiempos de Viaje/Detenciones) y explicar el gasto de Tracción.")
    
    if not df_thdr_v1.empty and not df_thdr_v2.empty and not df_ops.empty and not df_carga_v1.empty:
        
        # --- NUEVO FILTRO DE ESCENARIO POR TIPO DE JORNADA ---
        st.markdown("#### 🎛️ Filtro de Escenario Analítico")
        filtro_dia_multi = st.multiselect(
            "Selecciona el Tipo de Jornada a analizar en el Ecosistema y Mapas de Calor:",
            options=["L", "S", "D/F"],
            default=["L", "S", "D/F"],
            key="filtro_multi",
            format_func=lambda x: {"L": "Laboral (L)", "S": "Sábado (S)", "D/F": "Domingo y Festivo (D/F)"}.get(x, x)
        )
        
        # Filtrar matemáticamente todas las bases de datos subyacentes
        fechas_validas = df_ops[df_ops['Tipo Día'].isin(filtro_dia_multi)]['Fecha']
        df_ops_filt = df_ops[df_ops['Fecha'].isin(fechas_validas)]
        df_thdr_v1_filt = df_thdr_v1[df_thdr_v1['Fecha_Op'].isin(fechas_validas)]
        df_thdr_v2_filt = df_thdr_v2[df_thdr_v2['Fecha_Op'].isin(fechas_validas)]
        df_carga_v1_filt = df_carga_v1[df_carga_v1['Fecha'].isin(fechas_validas)] if not df_carga_v1.empty else pd.DataFrame()
        df_carga_v2_filt = df_carga_v2[df_carga_v2['Fecha'].isin(fechas_validas)] if not df_carga_v2.empty else pd.DataFrame()

        if df_ops_filt.empty:
            st.warning("No hay datos para los filtros seleccionados en este rango de fechas.")
        else:
            # --- 1. PREPARACIÓN BUBBLE CHART (Macro) ---
            def extr_tiempos_bubble(df_t, sal_str, lleg_str):
                if df_t.empty: return pd.DataFrame()
                c_sal = get_col_thdr(df_t, sal_str, 'SALIDA')
                c_lleg = get_col_thdr(df_t, lleg_str, 'LLEGADA')
                if not c_sal or not c_lleg: return pd.DataFrame()
                
                s_sal = extract_series(df_t, c_sal)
                s_lleg = extract_series(df_t, c_lleg)
                
                t_v = pd.DataFrame({'Fecha_Op': df_t['Fecha_Op'], 'Salida': s_sal, 'Llegada': s_lleg}).dropna()
                t_v['Dur'] = t_v['Llegada'] - t_v['Salida']
                t_v['Dur'] = t_v['Dur'].apply(lambda x: x + 1440 if pd.notna(x) and x < -500 else x)
                t_v = t_v[(t_v['Dur'] > 30) & (t_v['Dur'] < 120)]
                return t_v.groupby('Fecha_Op')['Dur'].median().reset_index()

            tv1 = extr_tiempos_bubble(df_thdr_v1_filt, 'PUERTO', 'LIMACHE')
            tv2 = extr_tiempos_bubble(df_thdr_v2_filt, 'LIMACHE', 'PUERTO')
            
            df_tiempos = pd.DataFrame(columns=['Fecha'])
            if not tv1.empty:
                tv1.columns = ['Fecha', 'Tiempo_V1']
                df_tiempos = pd.merge(df_tiempos, tv1, on='Fecha', how='outer') if not df_tiempos.empty else tv1
            if not tv2.empty:
                tv2.columns = ['Fecha', 'Tiempo_V2']
                df_tiempos = pd.merge(df_tiempos, tv2, on='Fecha', how='outer') if not df_tiempos.empty else tv2
                
            if 'Tiempo_V1' not in df_tiempos.columns: df_tiempos['Tiempo_V1'] = np.nan
            if 'Tiempo_V2' not in df_tiempos.columns: df_tiempos['Tiempo_V2'] = np.nan
                
            df_tiempos['Tiempo_Mediana_Red'] = df_tiempos[['Tiempo_V1', 'Tiempo_V2']].mean(axis=1)
            df_tiempos['Fecha'] = pd.to_datetime(df_tiempos['Fecha']).dt.normalize()

            df_mixto = pd.merge(df_ops_filt, df_tiempos[['Fecha', 'Tiempo_Mediana_Red']], on='Fecha', how='inner')
            df_plot = df_mixto.dropna(subset=['Tiempo_Mediana_Red', 'E_Tr', 'PAX']).copy()
            df_plot = df_plot[df_plot['PAX'] > 0] 
            
            if not df_plot.empty and df_plot['E_Tr'].sum() > 0:
                
                # --- 2. BUBBLE CHART 4D ---
                st.markdown("#### 🫧 Ecosistema Operativo Diario (Macro)")
                st.caption("Eje X: Mediana de Tiempos de Viaje | Eje Y: Consumo Tracción | Tamaño: Volumen de Pasajeros")
                
                df_plot['Tiempo Promedio HH:MM:SS'] = df_plot['Tiempo_Mediana_Red'].apply(minutos_a_hhmmss)
                
                # 🛡️ CORRECCIÓN: Tooltip robusto con listas (evita error de coerción en basevalidators de Plotly)
                fig_mix = px.scatter(df_plot, 
                                     x='Tiempo_Mediana_Red', 
                                     y='E_Tr', 
                                     size='PAX',
                                     color='Tipo Día', 
                                     hover_name='Fecha (ES)',
                                     hover_data=['Tiempo Promedio HH:MM:SS', 'IDE (kWh/km)'],
                                     labels={'Tiempo_Mediana_Red': 'Tiempo Mediano de Viaje', 'E_Tr': 'Energía de Tracción (kWh)'},
                                     color_discrete_map={'L': '#005195', 'S': '#E85500', 'D/F': '#2CA02C'})
                
                fig_mix.update_layout(margin=dict(t=20, b=0, l=0, r=0), height=400)
                st.plotly_chart(fig_mix, use_container_width=True)
                
                st.divider()
                
                # --- FUNCIÓN DE EXTRACCIÓN DE PAX (MOTOR ESPACIO-TEMPORAL CORRECTO) ---
                def extraer_pax_heatmap(df_carga_filt, df_thdr_filt, valid_stations):
                    if df_carga_filt.empty or df_thdr_filt.empty: return []
                    df_c = df_carga_filt.copy()
                    
                    # Identificar columnas de servicio para el cruce
                    c_serv_c = next((c for c in df_c.columns if 'THDR' in str(c).upper() or 'VIAJE' in str(c).upper() or 'SERV' in str(c).upper()), None)
                    c_serv_t = next((c for c in df_thdr_filt.columns if 'SERV' in str(c).upper() or 'VIAJE' in str(c).upper() or 'THDR' in str(c).upper()), None)
                    
                    if not c_serv_c or not c_serv_t:
                        return []

                    # Limpiar Servicio para que sea puro número
                    df_c['_srv_clean'] = pd.to_numeric(df_c[c_serv_c].astype(str).apply(lambda x: re.sub(r'\D', '', x)), errors='coerce')
                    t_sub = df_thdr_filt.copy()
                    t_sub['_srv_clean'] = pd.to_numeric(t_sub[c_serv_t].astype(str).apply(lambda x: re.sub(r'\D', '', x)), errors='coerce')

                    alias_map_pax = {
                        "PUERTO": ["PTO", "PUERTO"],
                        "BELLAVISTA": ["BELLA", "BELLAVISTA"],
                        "FRANCIA": ["FRANCIA"],
                        "BARON": ["BARON"],
                        "PORTALES": ["PORTALES"],
                        "RECREO": ["RECREO"],
                        "MIRAMAR": ["MIRAMAR"],
                        "VIÑA DEL MAR": ["VINA", "V. MAR", "V MAR", "VIÑA", "V.MAR"],
                        "HOSPITAL": ["HOSPITAL", "HOSP"],
                        "CHORRILLOS": ["CHORRILLOS", "CHORR", "CHORRILLO"],
                        "EL SALTO": ["SALTO", "E. SALTO"],
                        "VALENCIA": ["VALENCIA"],
                        "QUILPUE": ["QUILPUE", "QUILPUÉ"],
                        "EL SOL": ["EL SOL"],
                        "EL BELLOTO": ["BELLOTO", "E. BELLOTO"],
                        "LAS AMERICAS": ["AMERICAS", "L. AMERICAS"],
                        "LA CONCEPCION": ["CONCEPCION", "L. CONCEPCION"],
                        "VILLA ALEMANA": ["VILLA", "ALEMANA", "V. ALEMANA", "V.ALEMANA", "V. ALE"],
                        "SARGENTO ALDEA": ["SARGENTO", "ALDEA", "S. ALDEA", "S.ALDEA"],
                        "PEÑABLANCA": ["PENA BLANCA", "PENABLANCA", "PEÑA BLANCA"],
                        "LIMACHE": ["LIMACHE", "LIM"]
                    }
                    
                    pax_data = []
                    
                    for est in valid_stations:
                        est_n = _norm(est)
                        
                        # 1. Buscar la columna de pasajeros para esta estación
                        c_est = None
                        for c in df_c.columns:
                            cn = _norm(c)
                            if 'MAX' in cn or 'MIN' in cn or 'TOTAL' in cn or 'PROMEDIO' in cn: continue
                            if est_n in cn: 
                                c_est = c; break
                            if est_n in alias_map_pax:
                                for alias in alias_map_pax[est_n]:
                                    if alias in cn: c_est = c; break
                                if c_est: break
                                
                        # 2. Buscar la hora de paso por esta estación en la THDR
                        c_time = get_col_thdr(t_sub, est, 'SALIDA')
                        if not c_time:
                            c_time = get_col_thdr(t_sub, est, 'LLEGADA')
                            
                        if c_est and c_time:
                            s_time = extract_series(t_sub, c_time)
                            
                            # Crear un dataframe temporal de THDR solo con el servicio y la hora en ESTA estación
                            t_est = pd.DataFrame({
                                'Fecha_Op': t_sub['Fecha_Op'], 
                                '_srv_clean': t_sub['_srv_clean'], 
                                'Hora_Estacion': (s_time // 60).astype(float)
                            })
                            
                            # Cruzar con los pasajeros de esa estación
                            merged = pd.merge(df_c[['Fecha', '_srv_clean', c_est]], t_est,
                                              left_on=['Fecha', '_srv_clean'], 
                                              right_on=['Fecha_Op', '_srv_clean'], 
                                              how='inner')
                            
                            merged = merged.rename(columns={c_est: 'Pax'})
                            merged['Pax'] = pd.to_numeric(merged['Pax'], errors='coerce')
                            merged = merged.dropna(subset=['Hora_Estacion', 'Pax'])
                            merged['Hora_Estacion'] = merged['Hora_Estacion'].astype(int)
                            merged = merged[(merged['Hora_Estacion'] >= 5) & (merged['Hora_Estacion'] <= 23)]
                            
                            if not merged.empty:
                                merged['Estacion'] = SHORT_NAMES_DICT[est]
                                pax_data.append(merged[['Hora_Estacion', 'Estacion', 'Pax']])
                                
                    # 🛡️ Fallback si no encontró estaciones individuales
                    if not pax_data:
                        c_tot = next((c for c in df_c.columns if 'TOTAL' in str(c).upper() and 'BORDO' in str(c).upper()), None)
                        c_orig = get_col_thdr(t_sub, valid_stations[0], 'SALIDA') if valid_stations else None
                        if c_tot and c_orig:
                            s_orig = extract_series(t_sub, c_orig)
                            t_est = pd.DataFrame({'Fecha_Op': t_sub['Fecha_Op'], '_srv_clean': t_sub['_srv_clean'], 'Hora_Estacion': (s_orig // 60).astype(float)})
                            merged = pd.merge(df_c[['Fecha', '_srv_clean', c_tot]], t_est, left_on=['Fecha', '_srv_clean'], right_on=['Fecha_Op', '_srv_clean'], how='inner')
                            merged = merged.rename(columns={c_tot: 'Pax'})
                            merged['Pax'] = pd.to_numeric(merged['Pax'], errors='coerce')
                            merged = merged.dropna(subset=['Hora_Estacion', 'Pax'])
                            merged['Hora_Estacion'] = merged['Hora_Estacion'].astype(int)
                            merged = merged[(merged['Hora_Estacion'] >= 5) & (merged['Hora_Estacion'] <= 23)]
                            if not merged.empty:
                                merged['Estacion'] = 'Total en el Tren (Malla)'
                                pax_data.append(merged[['Hora_Estacion', 'Estacion', 'Pax']])
                                
                    return pax_data

                # --- ANÁLISIS TOPOLÓGICO POR HORA DE SALIDA (HEATMAPS) VÍA 1 ---
                st.markdown("#### 🔍 Análisis Topológico Horario (Vía 1)")
                st.markdown("Dinámica cruzada de Pasajeros, Circulación y Detenciones para los trenes en dirección Puerto ➔ Limache.")

                if not df_thdr_v1_filt.empty:
                    valid_stations_v1 = []
                    for est in ESTACIONES:
                        if get_col_thdr(df_thdr_v1_filt, est, 'LLEGADA') or get_col_thdr(df_thdr_v1_filt, est, 'SALIDA'):
                            valid_stations_v1.append(est)

                    # 3. EXTRAER DWELL TIMES
                    dwell_data = []
                    for est in valid_stations_v1:
                        c_lleg = get_col_thdr(df_thdr_v1_filt, est, 'LLEGADA')
                        c_sal = get_col_thdr(df_thdr_v1_filt, est, 'SALIDA')
                        
                        if c_lleg and c_sal:
                            s_lleg = extract_series(df_thdr_v1_filt, c_lleg)
                            s_sal = extract_series(df_thdr_v1_filt, c_sal)
                            
                            temp = pd.DataFrame({'Fecha_Op': df_thdr_v1_filt['Fecha_Op'], 'Llegada': s_lleg, 'Salida': s_sal})
                            temp['Dwell'] = temp['Salida'] - temp['Llegada']
                            temp['Dwell'] = temp['Dwell'].apply(lambda x: x + 1440 if pd.notna(x) and x < -1000 else x)
                            temp = temp[(temp['Dwell'] >= 0) & (temp['Dwell'] < 15)].dropna()
                            
                            if not temp.empty:
                                temp['Hora_Estacion'] = (temp['Salida'] // 60).astype(int)
                                temp = temp[(temp['Hora_Estacion'] >= 5) & (temp['Hora_Estacion'] <= 23)]
                                temp['Estacion'] = SHORT_NAMES_DICT[est] 
                                dwell_data.append(temp[['Hora_Estacion', 'Estacion', 'Dwell']])

                    # 4. EXTRAER RUNNING TIMES
                    running_data = []
                    for i in range(len(valid_stations_v1)-1):
                        e_A = valid_stations_v1[i]
                        e_B = valid_stations_v1[i+1]
                        c_s = get_col_thdr(df_thdr_v1_filt, e_A, 'SALIDA')
                        c_l = get_col_thdr(df_thdr_v1_filt, e_B, 'LLEGADA')
                        
                        if c_s and c_l:
                            s_sal = extract_series(df_thdr_v1_filt, c_s)
                            s_lleg = extract_series(df_thdr_v1_filt, c_l)
                            
                            temp = pd.DataFrame({'Fecha_Op': df_thdr_v1_filt['Fecha_Op'], 'Salida': s_sal, 'Llegada': s_lleg})
                            temp['RunTime'] = temp['Llegada'] - temp['Salida']
                            temp['RunTime'] = temp['RunTime'].apply(lambda x: x + 1440 if pd.notna(x) and x < -1000 else x)
                            temp = temp[(temp['RunTime'] > 0) & (temp['RunTime'] < 30)].dropna()
                            
                            if not temp.empty:
                                temp['Hora_Estacion'] = (temp['Salida'] // 60).astype(int)
                                temp = temp[(temp['Hora_Estacion'] >= 5) & (temp['Hora_Estacion'] <= 23)]
                                temp['Tramo'] = f"{SHORT_NAMES_DICT[e_A]}-{SHORT_NAMES_DICT[e_B]}"
                                running_data.append(temp[['Hora_Estacion', 'Tramo', 'RunTime']])

                    # 5. EXTRAER PAX
                    pax_data_v1 = extraer_pax_heatmap(df_carga_v1_filt, df_thdr_v1_filt, valid_stations_v1)

                    c_h1, c_h2, c_h3 = st.columns(3)

                    with c_h1:
                        st.markdown("##### 🛑 Dwell Time (Detención)")
                        if dwell_data:
                            df_dwell_full = pd.concat(dwell_data)
                            df_dwell_heat = df_dwell_full.groupby(['Estacion', 'Hora_Estacion'])['Dwell'].median().reset_index()
                            
                            pivot_dwell = df_dwell_heat.pivot(index='Estacion', columns='Hora_Estacion', values='Dwell')
                            pivot_dwell = pivot_dwell.reindex([SHORT_NAMES_DICT[e] for e in ESTACIONES[::-1] if e in valid_stations_v1])
                            pivot_fmt_dw = pivot_dwell.apply(lambda col: col.apply(minutos_a_hhmmss))
                            
                            fig_heat_dw = px.imshow(pivot_dwell, 
                                                    labels=dict(x="Hora Local", y="Estación", color="Tiempo"),
                                                    x=pivot_dwell.columns,
                                                    y=pivot_dwell.index,
                                                    color_continuous_scale="Reds",
                                                    aspect="auto")
                            fig_heat_dw.update_traces(customdata=pivot_fmt_dw.values, hovertemplate="Hora Local: %{x}:00<br>Estación: %{y}<br>Detención: %{customdata}<extra></extra>")
                            fig_heat_dw.update_layout(margin=dict(t=20, b=20, l=0, r=0), height=500)
                            st.plotly_chart(fig_heat_dw, use_container_width=True)
                            
                            est_max = df_dwell_heat.loc[df_dwell_heat['Dwell'].idxmax()]
                            st.caption(f"**Insight:** Máx. detención típica en **{est_max['Estacion']}** a las **{est_max['Hora_Estacion']:02d}:00 hrs** ({minutos_a_hhmmss(est_max['Dwell'])}).")
                        else:
                            st.info("Sin datos.")

                    with c_h2:
                        st.markdown("##### 🛤️ Circulación (Tramo)")
                        if running_data:
                            df_run_full = pd.concat(running_data)
                            df_run_heat = df_run_full.groupby(['Tramo', 'Hora_Estacion'])['RunTime'].median().reset_index()
                            
                            tramos_orden = [f"{SHORT_NAMES_DICT[valid_stations_v1[i]]}-{SHORT_NAMES_DICT[valid_stations_v1[i+1]]}" for i in range(len(valid_stations_v1)-1)]
                            pivot_run = df_run_heat.pivot(index='Tramo', columns='Hora_Estacion', values='RunTime')
                            pivot_run = pivot_run.reindex(tramos_orden[::-1]) 
                            pivot_fmt_run = pivot_run.apply(lambda col: col.apply(minutos_a_hhmmss))
                            
                            fig_heat_run = px.imshow(pivot_run, 
                                                    labels=dict(x="Hora Local", y="Tramo", color="Tiempo"),
                                                    x=pivot_run.columns,
                                                    y=pivot_run.index,
                                                    color_continuous_scale="Blues",
                                                    aspect="auto")
                            fig_heat_run.update_traces(customdata=pivot_fmt_run.values, hovertemplate="Hora Local: %{x}:00<br>Tramo: %{y}<br>Tiempo: %{customdata}<extra></extra>")
                            fig_heat_run.update_layout(margin=dict(t=20, b=20, l=0, r=0), height=500)
                            st.plotly_chart(fig_heat_run, use_container_width=True)
                            
                            tramo_max = df_run_heat.loc[df_run_heat['RunTime'].idxmax()]
                            st.caption(f"**Insight:** Tramo más lento **{tramo_max['Tramo']}** a las **{tramo_max['Hora_Estacion']:02d}:00 hrs** ({minutos_a_hhmmss(tramo_max['RunTime'])}).")
                        else:
                            st.info("Sin datos.")

                    with c_h3:
                        st.markdown("##### 👥 Carga de Pasajeros")
                        if pax_data_v1:
                            df_pax_full = pd.concat(pax_data_v1)
                            # UTILIZANDO PROMEDIO (MEAN) PARA VOLUMETRÍA
                            df_pax_heat = df_pax_full.groupby(['Estacion', 'Hora_Estacion'])['Pax'].mean().reset_index()
                            
                            pivot_pax = df_pax_heat.pivot(index='Estacion', columns='Hora_Estacion', values='Pax')
                            
                            if 'Total en el Tren (Malla)' in df_pax_heat['Estacion'].values:
                                pivot_pax = pivot_pax.reindex(['Total en el Tren (Malla)'])
                            else:
                                pivot_pax = pivot_pax.reindex([SHORT_NAMES_DICT[e] for e in ESTACIONES[::-1] if e in valid_stations_v1])
                            
                            fig_heat_pax = px.imshow(pivot_pax, 
                                                    labels=dict(x="Hora Local", y="Métrica", color="PAX Prom"),
                                                    x=pivot_pax.columns,
                                                    y=pivot_pax.index,
                                                    color_continuous_scale="Greens",
                                                    aspect="auto")
                            fig_heat_pax.update_traces(hovertemplate="Hora Local: %{x}:00<br>Sector: %{y}<br>Carga PAX: %{z:,.0f}<extra></extra>")
                            fig_heat_pax.update_layout(margin=dict(t=20, b=20, l=0, r=0), height=500)
                            st.plotly_chart(fig_heat_pax, use_container_width=True)
                            
                            pax_max = df_pax_heat.loc[df_pax_heat['Pax'].idxmax()]
                            st.caption(f"**Insight:** Mayor carga prom. en **{pax_max['Estacion']}** a las **{pax_max['Hora_Estacion']:02d}:00 hrs** ({pax_max['Pax']:,.0f} PAX).")
                        else:
                            st.info("Formato de pasajeros no detectado (ni por estación ni consolidado).")

                else:
                    st.info("Se requiere procesar archivo THDR Vía 1 para este filtro.")

                st.divider()

                # --- ANÁLISIS TOPOLÓGICO POR HORA DE SALIDA (HEATMAPS) VÍA 2 ---
                st.markdown("#### 🔍 Análisis Topológico Horario (Vía 2)")
                st.markdown("Dinámica cruzada para los trenes en dirección Limache ➔ Puerto.")

                if not df_thdr_v2_filt.empty:
                    valid_stations_v2 = []
                    for est in ESTACIONES:
                        if get_col_thdr(df_thdr_v2_filt, est, 'LLEGADA') or get_col_thdr(df_thdr_v2_filt, est, 'SALIDA'):
                            valid_stations_v2.append(est)

                    dwell_data_v2 = []
                    for est in valid_stations_v2:
                        c_lleg = get_col_thdr(df_thdr_v2_filt, est, 'LLEGADA')
                        c_sal = get_col_thdr(df_thdr_v2_filt, est, 'SALIDA')
                        if c_lleg and c_sal:
                            s_lleg = extract_series(df_thdr_v2_filt, c_lleg)
                            s_sal = extract_series(df_thdr_v2_filt, c_sal)
                            temp = pd.DataFrame({'Fecha_Op': df_thdr_v2_filt['Fecha_Op'], 'Llegada': s_lleg, 'Salida': s_sal})
                            temp['Dwell'] = temp['Salida'] - temp['Llegada']
                            temp['Dwell'] = temp['Dwell'].apply(lambda x: x + 1440 if pd.notna(x) and x < -1000 else x)
                            temp = temp[(temp['Dwell'] >= 0) & (temp['Dwell'] < 15)].dropna()
                            if not temp.empty:
                                temp['Hora_Estacion'] = (temp['Salida'] // 60).astype(int)
                                temp = temp[(temp['Hora_Estacion'] >= 5) & (temp['Hora_Estacion'] <= 23)]
                                temp['Estacion'] = SHORT_NAMES_DICT[est]
                                dwell_data_v2.append(temp[['Hora_Estacion', 'Estacion', 'Dwell']])

                    running_data_v2 = []
                    estaciones_orden_v2 = [e for e in ESTACIONES[::-1] if e in valid_stations_v2]
                    for i in range(len(estaciones_orden_v2)-1):
                        e_A = estaciones_orden_v2[i]
                        e_B = estaciones_orden_v2[i+1]
                        c_s = get_col_thdr(df_thdr_v2_filt, e_A, 'SALIDA')
                        c_l = get_col_thdr(df_thdr_v2_filt, e_B, 'LLEGADA')
                        if c_s and c_l:
                            s_sal = extract_series(df_thdr_v2_filt, c_s)
                            s_lleg = extract_series(df_thdr_v2_filt, c_l)
                            temp = pd.DataFrame({'Fecha_Op': df_thdr_v2_filt['Fecha_Op'], 'Salida': s_sal, 'Llegada': s_lleg})
                            temp['RunTime'] = temp['Llegada'] - temp['Salida']
                            temp['RunTime'] = temp['RunTime'].apply(lambda x: x + 1440 if pd.notna(x) and x < -1000 else x)
                            temp = temp[(temp['RunTime'] > 0) & (temp['RunTime'] < 30)].dropna()
                            if not temp.empty:
                                temp['Hora_Estacion'] = (temp['Salida'] // 60).astype(int)
                                temp = temp[(temp['Hora_Estacion'] >= 5) & (temp['Hora_Estacion'] <= 23)]
                                temp['Tramo'] = f"{SHORT_NAMES_DICT[e_A]}-{SHORT_NAMES_DICT[e_B]}"
                                running_data_v2.append(temp[['Hora_Estacion', 'Tramo', 'RunTime']])

                    pax_data_v2 = extraer_pax_heatmap(df_carga_v2_filt, df_thdr_v2_filt, valid_stations_v2)

                    c_h4, c_h5, c_h6 = st.columns(3)

                    with c_h4:
                        st.markdown("##### 🛑 Dwell Time (Vía 2)")
                        if dwell_data_v2:
                            df_dwell_full_v2 = pd.concat(dwell_data_v2)
                            df_dwell_heat_v2 = df_dwell_full_v2.groupby(['Estacion', 'Hora_Estacion'])['Dwell'].median().reset_index()
                            pivot_dwell_v2 = df_dwell_heat_v2.pivot(index='Estacion', columns='Hora_Estacion', values='Dwell')
                            pivot_dwell_v2 = pivot_dwell_v2.reindex([SHORT_NAMES_DICT[e] for e in ESTACIONES[::-1] if e in valid_stations_v2])
                            pivot_fmt_dw_v2 = pivot_dwell_v2.apply(lambda col: col.apply(minutos_a_hhmmss))
                            fig_heat_dw_v2 = px.imshow(pivot_dwell_v2, 
                                                    labels=dict(x="Hora Local", y="Estación", color="Tiempo"),
                                                    x=pivot_dwell_v2.columns, y=pivot_dwell_v2.index,
                                                    color_continuous_scale="Oranges", aspect="auto")
                            fig_heat_dw_v2.update_traces(customdata=pivot_fmt_dw_v2.values, hovertemplate="Hora Local: %{x}:00<br>Estación: %{y}<br>Detención: %{customdata}<extra></extra>")
                            fig_heat_dw_v2.update_layout(margin=dict(t=20, b=20, l=0, r=0), height=500)
                            st.plotly_chart(fig_heat_dw_v2, use_container_width=True)
                            est_max_v2 = df_dwell_heat_v2.loc[df_dwell_heat_v2['Dwell'].idxmax()]
                            st.caption(f"**Insight:** Máx. detención típica en **{est_max_v2['Estacion']}** a las **{est_max_v2['Hora_Estacion']:02d}:00 hrs** ({minutos_a_hhmmss(est_max_v2['Dwell'])}).")
                        else:
                            st.info("Sin datos.")

                    with c_h5:
                        st.markdown("##### 🛤️ Circulación (Vía 2)")
                        if running_data_v2:
                            df_run_full_v2 = pd.concat(running_data_v2)
                            df_run_heat_v2 = df_run_full_v2.groupby(['Tramo', 'Hora_Estacion'])['RunTime'].median().reset_index()
                            tramos_orden_v2 = [f"{SHORT_NAMES_DICT[estaciones_orden_v2[i]]}-{SHORT_NAMES_DICT[estaciones_orden_v2[i+1]]}" for i in range(len(estaciones_orden_v2)-1)]
                            pivot_run_v2 = df_run_heat_v2.pivot(index='Tramo', columns='Hora_Estacion', values='RunTime')
                            pivot_run_v2 = pivot_run_v2.reindex(tramos_orden_v2[::-1]) 
                            pivot_fmt_run_v2 = pivot_run_v2.apply(lambda col: col.apply(minutos_a_hhmmss))
                            fig_heat_run_v2 = px.imshow(pivot_run_v2, 
                                                    labels=dict(x="Hora Local", y="Tramo", color="Tiempo"),
                                                    x=pivot_run_v2.columns, y=pivot_run_v2.index,
                                                    color_continuous_scale="Purples", aspect="auto")
                            fig_heat_run_v2.update_traces(customdata=pivot_fmt_run_v2.values, hovertemplate="Hora Local: %{x}:00<br>Tramo: %{y}<br>Tiempo: %{customdata}<extra></extra>")
                            fig_heat_run_v2.update_layout(margin=dict(t=20, b=20, l=0, r=0), height=500)
                            st.plotly_chart(fig_heat_run_v2, use_container_width=True)
                            tramo_max_v2 = df_run_heat_v2.loc[df_run_heat_v2['RunTime'].idxmax()]
                            st.caption(f"**Insight:** Tramo más lento **{tramo_max_v2['Tramo']}** a las **{tramo_max_v2['Hora_Estacion']:02d}:00 hrs** ({minutos_a_hhmmss(tramo_max_v2['RunTime'])}).")
                        else:
                            st.info("Sin datos.")

                    with c_h6:
                        st.markdown("##### 👥 Carga de Pasajeros (Vía 2)")
                        if pax_data_v2:
                            df_pax_full_v2 = pd.concat(pax_data_v2)
                            # UTILIZANDO PROMEDIO (MEAN) PARA VOLUMETRÍA
                            df_pax_heat_v2 = df_pax_full_v2.groupby(['Estacion', 'Hora_Estacion'])['Pax'].mean().reset_index()
                            pivot_pax_v2 = df_pax_heat_v2.pivot(index='Estacion', columns='Hora_Estacion', values='Pax')
                            
                            if 'Total en el Tren (Malla)' in df_pax_heat_v2['Estacion'].values:
                                pivot_pax_v2 = pivot_pax_v2.reindex(['Total en el Tren (Malla)'])
                            else:
                                pivot_pax_v2 = pivot_pax_v2.reindex([SHORT_NAMES_DICT[e] for e in ESTACIONES[::-1] if e in valid_stations_v2])
                                
                            fig_heat_pax_v2 = px.imshow(pivot_pax_v2, 
                                                    labels=dict(x="Hora Local", y="Métrica", color="PAX Prom"),
                                                    x=pivot_pax_v2.columns, y=pivot_pax_v2.index,
                                                    color_continuous_scale="Greens", aspect="auto")
                            fig_heat_pax_v2.update_traces(hovertemplate="Hora Local: %{x}:00<br>Sector: %{y}<br>Carga PAX: %{z:,.0f}<extra></extra>")
                            fig_heat_pax_v2.update_layout(margin=dict(t=20, b=20, l=0, r=0), height=500)
                            st.plotly_chart(fig_heat_pax_v2, use_container_width=True)
                            pax_max_v2 = df_pax_heat_v2.loc[df_pax_heat_v2['Pax'].idxmax()]
                            st.caption(f"**Insight:** Mayor carga prom. en **{pax_max_v2['Estacion']}** a las **{pax_max_v2['Hora_Estacion']:02d}:00 hrs** ({pax_max_v2['Pax']:,.0f} PAX).")
                        else:
                            st.info("Formato de pasajeros no detectado (ni por estación ni consolidado).")
                else:
                    st.info("Se requiere procesar archivo THDR Vía 2 para este filtro.")
            else:
                st.warning("No hay suficientes datos superpuestos para realizar la regresión en base a los filtros actuales.")
    else: 
        st.info("⚠️ Carga archivos de **THDR (Vía 1 y 2), Facturación/PRMTE/SEAT y Carga de Pasajeros** para habilitar el Microscopio Operacional.")

with tabs[9]:
    st.write("### Flujo y Carga de Pasajeros")
    if not df_carga_v1.empty or not df_carga_v2.empty:
        c_p1, c_p2 = st.columns(2)
        with c_p1:
            st.write("#### Total de Pasajeros por Día")
            df_c1_agg = df_carga_v1.groupby('Fecha')['Total a Bordo'].sum().reset_index() if (not df_carga_v1.empty and 'Fecha' in df_carga_v1.columns) else pd.DataFrame(columns=['Fecha', 'Total a Bordo'])
            df_c2_agg = df_carga_v2.groupby('Fecha')['Total a Bordo'].sum().reset_index() if (not df_carga_v2.empty and 'Fecha' in df_carga_v2.columns) else pd.DataFrame(columns=['Fecha', 'Total a Bordo'])
            
            fig_pas = go.Figure()
            if not df_c1_agg.empty:
                fig_pas.add_trace(go.Bar(x=df_c1_agg['Fecha'], y=df_c1_agg['Total a Bordo'], name='Vía 1 (Puerto->Limache)', marker_color='#005195'))
            if not df_c2_agg.empty:
                fig_pas.add_trace(go.Bar(x=df_c2_agg['Fecha'], y=df_c2_agg['Total a Bordo'], name='Vía 2 (Limache->Puerto)', marker_color='#E85500'))
            
            fig_pas.update_layout(barmode='group', xaxis_title="Fecha", yaxis_title="Total Pasajeros", margin=dict(t=30))
            st.plotly_chart(fig_pas, use_container_width=True)

        with c_p2:
            st.write("#### Estaciones con Carga Máxima (Frecuencia)")
            est_v1 = df_carga_v1['Estación Máxima'].value_counts().reset_index() if not df_carga_v1.empty else pd.DataFrame(columns=['Estación Máxima', 'count'])
            est_v2 = df_carga_v2['Estación Máxima'].value_counts().reset_index() if not df_carga_v2.empty else pd.DataFrame(columns=['Estación Máxima', 'count'])
            
            if not est_v1.empty: est_v1.columns = ['Estación', 'Frecuencia']; est_v1['Vía'] = 'Vía 1'
            if not est_v2.empty: est_v2.columns = ['Estación', 'Frecuencia']; est_v2['Vía'] = 'Vía 2'
            
            df_est = pd.concat([est_v1, est_v2]).sort_values('Frecuencia', ascending=True).tail(15)
            
            if not df_est.empty:
                fig_est = px.bar(df_est, x='Frecuencia', y='Estación', color='Vía', orientation='h', color_discrete_map={'Vía 1': '#005195', 'Vía 2': '#E85500'})
                fig_est.update_layout(xaxis_title="Frecuencia (N° de Viajes)", yaxis_title="Estación", margin=dict(t=30))
                st.plotly_chart(fig_est, use_container_width=True)

        st.write("#### Detalle de Viajes (Muestra)")
        if not df_carga_v1.empty:
            st.caption("Vía 1 (Puerto -> Limache)")
            dv_c1 = df_carga_v1.copy()
            dv_c1['Fecha'] = dv_c1['Fecha'].dt.strftime('%Y-%m-%d')
            st.dataframe(make_columns_unique(dv_c1.head(100)), use_container_width=True)
            
        if not df_carga_v2.empty:
            st.caption("Vía 2 (Limache -> Puerto)")
            dv_c2 = df_carga_v2.copy()
            dv_c2['Fecha'] = dv_c2['Fecha'].dt.strftime('%Y-%m-%d')
            st.dataframe(make_columns_unique(dv_c2.head(100)), use_container_width=True)
            
    else: st.info("No se han procesado datos de carga de pasajeros. Verifica los archivos subidos o tu Rango de Fechas.")

with tabs[10]:
    st.markdown("### 📝 Análisis Ejecutivo Automático")
    if not df_ops.empty:
        if 'filtro_dia' in locals():
            df_reporte = df_ops[df_ops['Tipo Día'].isin(filtro_dia)]
        else:
            df_reporte = df_ops
            
        if 'drilldown_date' in st.session_state and st.session_state.drilldown_date is not None:
            df_reporte = df_reporte[df_reporte['Fecha'] == st.session_state.drilldown_date]
        
        if df_reporte.empty:
            st.warning("No hay datos para generar el reporte con los filtros o días seleccionados.")
        else:
            st.markdown("Este reporte es generado automáticamente aplicando **algoritmos de estadística descriptiva**, segmentando la operación por sus distintos perfiles de demanda (Tipos de Jornada).")
            
            tot_traccion_global = df_reporte['E_Tr'].sum()
            tot_pax_global = df_reporte['PAX'].sum()
            kwh_per_pax_global = (tot_traccion_global / tot_pax_global) if tot_pax_global > 0 else 0
            
            st.info(f"🎯 **KPI Global de Sostenibilidad (Estándar UIC):** En toda la selección analizada, la empresa consumió en promedio **{kwh_per_pax_global:.2f} kWh de tracción por cada pasajero transportado**.")
            
            tipos_ordenados = ["L", "S", "D/F"]
            nombres_tipos = {"L": "Días Laborales (L)", "S": "Sábados (S)", "D/F": "Domingos y Festivos (D/F)"}
            
            for tipo in tipos_ordenados:
                df_tipo = df_reporte[df_reporte['Tipo Día'] == tipo]
                if df_tipo.empty: continue
                
                with st.expander(f"📌 Análisis de Operación: {nombres_tipos[tipo]}", expanded=True):
                    
                    dia_max_ide = df_tipo.loc[df_tipo['IDE (kWh/km)'].idxmax()]
                    dias_validos_ide = df_tipo[df_tipo['IDE (kWh/km)'] > 0]
                    dia_min_ide = dias_validos_ide.loc[dias_validos_ide['IDE (kWh/km)'].idxmin()] if not dias_validos_ide.empty else dia_max_ide
                    
                    tot_tren_km = df_tipo['Tren-Km [km]'].sum()
                    tot_odo = df_tipo['Odómetro [km]'].sum()
                    umr_global = (tot_tren_km / tot_odo * 100) if tot_odo > 0 else 0
                    dia_max_pax = df_tipo.loc[df_tipo['PAX'].idxmax()] if df_tipo['PAX'].sum() > 0 else None

                    fechas_tipo = df_tipo['Fecha'].tolist()
                    
                    peak_hr_msg = "No hay datos horarios."
                    noche_msg = ""
                    datos_hr = all_prmte_full if all_prmte_full else all_fact_full
                    if datos_hr:
                        df_hr = pd.DataFrame(datos_hr)
                        df_hr['Fecha'] = pd.to_datetime(df_hr['Fecha'])
                        df_hr_filt = df_hr[df_hr['Fecha'].isin(fechas_tipo)]
                        if not df_hr_filt.empty:
                            df_hr_filt = df_hr_filt.copy() 
                            hr_agrupado = df_hr_filt.groupby('Hora')['Consumo'].mean()
                            hora_peak = hr_agrupado.idxmax()
                            consumo_peak = hr_agrupado.max()
                            peak_hr_msg = f"La 'Hora Punta Eléctrica' ocurre a las **{hora_peak}** ({consumo_peak:,.0f} kWh prom.)."
                            
                            df_hr_filt['Hora_n'] = df_hr_filt['Hora'].astype(str).str.slice(0, 2).astype(int)
                            limite_hora = 6 if tipo == "L" else (7 if tipo == "S" else 8)
                            
                            df_noche = df_hr_filt[(df_hr_filt['Hora_n'] >= 0) & (df_hr_filt['Hora_n'] < limite_hora)]
                            if not df_noche.empty:
                                noche_diario = df_noche.groupby('Fecha')['Consumo'].sum().reset_index()
                                promedio_noche = noche_diario['Consumo'].mean()
                                max_noche = noche_diario.loc[noche_diario['Consumo'].idxmax()]
                                if max_noche['Consumo'] > (promedio_noche * 1.2) and promedio_noche > 0:
                                    noche_msg = f"🌙 **Alerta Parásita:** Pico de **{max_noche['Consumo']:,.0f} kWh** la madrugada del {max_noche['Fecha'].strftime('%d/%m')} (Ventana: 00:00 a 0{limite_hora}:00 hrs)."
                                else:
                                    noche_msg = f"🌙 **Auditoría Nocturna:** Estable ({promedio_noche:,.0f} kWh de 00:00 a 0{limite_hora}:00 hrs)."

                    estacion_msg = "Sin datos de estaciones."
                    df_c_filt = pd.DataFrame()
                    if not df_carga_v1.empty or not df_carga_v2.empty:
                        c1 = df_carga_v1[df_carga_v1['Fecha'].isin(fechas_tipo)] if not df_carga_v1.empty else pd.DataFrame()
                        c2 = df_carga_v2[df_carga_v2['Fecha'].isin(fechas_tipo)] if not df_carga_v2.empty else pd.DataFrame()
                        df_c_filt = pd.concat([c1, c2])
                        if not df_c_filt.empty and 'Estación Máxima' in df_c_filt.columns:
                            estacion_critica = df_c_filt['Estación Máxima'].value_counts().idxmax()
                            frecuencia_critica = df_c_filt['Estación Máxima'].value_counts().max()
                            estacion_msg = f"**{estacion_critica}** es el cuello de botella físico ({frecuencia_critica} viajes a máxima capacidad)."

                    thdr_msg = "Sin datos de THDR."
                    tiempo_msg = ""
                    if not df_thdr_v1.empty or not df_thdr_v2.empty:
                        t1 = df_thdr_v1[df_thdr_v1['Fecha_Op'].isin(fechas_tipo)] if not df_thdr_v1.empty else pd.DataFrame()
                        t2 = df_thdr_v2[df_thdr_v2['Fecha_Op'].isin(fechas_tipo)] if not df_thdr_v2.empty else pd.DataFrame()
                        df_t_filt = pd.concat([t1, t2])
                        if not df_t_filt.empty:
                            total_viajes = len(df_t_filt)
                            if 'Unidad' in df_t_filt.columns:
                                v_doble = len(df_t_filt[df_t_filt['Unidad'].astype(str).str.contains('M', case=False, na=False)])
                                thdr_msg = f"**{total_viajes:,} servicios** en total. Uso de Tracción Doble: **{(v_doble/total_viajes*100) if total_viajes>0 else 0:.1f}%**."

                        def formato_hora(h):
                            if pd.isna(h): return "N/A"
                            if isinstance(h, (datetime, time)): return h.strftime('%H:%M')
                            return str(h)[:5]

                        msg_v1, msg_v2, msg_insight = "", "", ""
                        brecha_max = 0
                        
                        if not t1.empty:
                            c_serv_v1 = t1.columns[0]
                            c_p_sal = get_col_thdr(t1, 'PUERTO', 'SALIDA')
                            c_l_lleg = get_col_thdr(t1, 'LIMACHE', 'LLEGADA')
                            if c_p_sal and c_l_lleg:
                                s_sal = extract_series(t1, c_p_sal)
                                s_lleg = extract_series(t1, c_l_lleg)
                                raw_sal = t1[c_p_sal].iloc[:, 0] if isinstance(t1[c_p_sal], pd.DataFrame) else t1[c_p_sal]
                                t1_v = pd.DataFrame({c_serv_v1: t1[c_serv_v1], 'Fecha_Op': t1['Fecha_Op'], 'Salida_raw': raw_sal, 'Salida': s_sal, 'Llegada': s_lleg}).dropna()
                                t1_v['Dur'] = t1_v['Llegada'] - t1_v['Salida']
                                t1_v['Dur'] = t1_v['Dur'].apply(lambda x: x + 1440 if x < -500 else x)
                                t1_v = t1_v[(t1_v['Dur'] > 30) & (t1_v['Dur'] < 120)]
                                if not t1_v.empty:
                                    r_min, r_max = t1_v.loc[t1_v['Dur'].idxmin()], t1_v.loc[t1_v['Dur'].idxmax()]
                                    msg_v1 = f"**V1 (PU→LI) Promedio: {minutos_a_hhmmss(t1_v['Dur'].mean())}**\n\n- 🟢 *Rápido:* {minutos_a_hhmmss(r_min['Dur'])} ({r_min['Fecha_Op'].strftime('%d/%m')}, Serv. {r_min[c_serv_v1]}, {formato_hora(r_min['Salida_raw'])})\n- 🔴 *Lento:* {minutos_a_hhmmss(r_max['Dur'])} ({r_max['Fecha_Op'].strftime('%d/%m')}, Serv. {r_max[c_serv_v1]}, {formato_hora(r_max['Salida_raw'])})"
                                    brecha_max = max(brecha_max, r_max['Dur'] - r_min['Dur'])

                        if not t2.empty:
                            c_serv_v2 = t2.columns[0]
                            c_l_sal = get_col_thdr(t2, 'LIMACHE', 'SALIDA')
                            c_p_lleg = get_col_thdr(t2, 'PUERTO', 'LLEGADA')
                            if c_l_sal and c_p_lleg:
                                s_sal = extract_series(t2, c_l_sal)
                                s_lleg = extract_series(t2, c_p_lleg)
                                raw_sal = t2[c_l_sal].iloc[:, 0] if isinstance(t2[c_l_sal], pd.DataFrame) else t2[c_l_sal]
                                t2_v = pd.DataFrame({c_serv_v2: t2[c_serv_v2], 'Fecha_Op': t2['Fecha_Op'], 'Salida_raw': raw_sal, 'Salida': s_sal, 'Llegada': s_lleg}).dropna()
                                t2_v['Dur'] = t2_v['Llegada'] - t2_v['Salida']
                                t2_v['Dur'] = t2_v['Dur'].apply(lambda x: x + 1440 if x < -500 else x)
                                t2_v = t2_v[(t2_v['Dur'] > 30) & (t2_v['Dur'] < 120)]
                                if not t2_v.empty:
                                    r_min, r_max = t2_v.loc[t2_v['Dur'].idxmin()], t2_v.loc[t2_v['Dur'].idxmax()]
                                    msg_v2 = f"**V2 (LI→PU) Promedio: {minutos_a_hhmmss(t2_v['Dur'].mean())}**\n\n- 🟢 *Rápido:* {minutos_a_hhmmss(r_min['Dur'])} ({r_min['Fecha_Op'].strftime('%d/%m')}, Serv. {r_min[c_serv_v2]}, {formato_hora(r_min['Salida_raw'])})\n- 🔴 *Lento:* {minutos_a_hhmmss(r_max['Dur'])} ({r_max['Fecha_Op'].strftime('%d/%m')}, Serv. {r_max[c_serv_v2]}, {formato_hora(r_max['Salida_raw'])})"
                                    brecha_max = max(brecha_max, r_max['Dur'] - r_min['Dur'])

                        if msg_v1 or msg_v2:
                            if brecha_max > 10: msg_insight = f"⚠️ *Alta inestabilidad ({minutos_a_hhmmss(brecha_max)} de brecha).* Fuerte impacto negativo en consumo de tracción."
                            else: msg_insight = f"✅ *Buena regularidad ({minutos_a_hhmmss(brecha_max)} de brecha).* Operación estable que favorece la conducción eficiente."

                    c_rep1, c_rep2 = st.columns(2)
                    with c_rep1:
                        st.markdown("##### 📊 Desempeño y Demanda")
                        st.success(f"🏆 **Mayor Eficiencia:** {dia_min_ide['Fecha (ES)']} (IDE: **{dia_min_ide['IDE (kWh/km)']:.2f} kWh/km**)")
                        st.warning(f"🚨 **Día Crítico (Ineficiente):** {dia_max_ide['Fecha (ES)']} (IDE: **{dia_max_ide['IDE (kWh/km)']:.2f} kWh/km**)")
                        if dia_max_pax is not None and dia_max_pax['PAX'] > 0:
                            st.info(f"👥 **Peak de Demanda:** {dia_max_pax['Fecha (ES)']} con **{int(dia_max_pax['PAX']):,}** personas.")
                        st.info(f"🚆 **Tasa UMR Promedio:** {umr_global:.1f}%.")
                        
                    with c_rep2:
                        st.markdown("##### 🔬 Diagnóstico Operativo")
                        st.info(f"⚡ **Red Eléctrica:** {peak_hr_msg}")
                        if noche_msg:
                            if "Alerta" in noche_msg: st.error(noche_msg)
                            else: st.success(noche_msg)
                        st.error(f"🛑 **Cuello de Botella:** {estacion_msg}")
                        st.info(f"📋 **Despachos:** {thdr_msg}")
                        
                    if msg_v1 or msg_v2:
                        st.markdown("##### ⏱️ Análisis de Tiempos de Viaje")
                        c_t1, c_t2 = st.columns(2)
                        with c_t1:
                            if msg_v1: st.info(msg_v1)
                        with c_t2:
                            if msg_v2: st.info(msg_v2)
                        if msg_insight:
                            st.markdown(msg_insight)
    else:
        st.info("No hay datos consolidados para generar el análisis.")

with tabs[11]:
    st.markdown("### 🩺 Diagnóstico Automático de Anomalías de Consumo")
    st.markdown("Compara **cada día con los de su mismo tipo** (Laboral / Sábado / Domingo-Festivo) "
                "con estadística robusta, detecta los que se salen de lo normal y **cruza la THDR y la "
                "carga de pasajeros** para explicar la causa raíz en lenguaje operativo y de negocio.")

    with st.expander("❓ ¿Qué significa el análisis y cómo leer esto?"):
        st.markdown(
            "El sistema utiliza estadística avanzada (**Z-score robusto basado en MAD**) para buscar qué días "
            "tuvieron consumos que rompieron los patrones normales de la empresa.\n\n"
            "- Cuanto más alta la barra, más grave: se activan alertas naranjas 🟠 (**Atención**) y rojas 🔴 (**Anomalía**).\n"
            "- A diferencia de un simple promedio mensual, esto compara lunes contra lunes o domingo contra domingos, "
            "evitando generar alarmas falsas en días de fin de semana.\n\n"
            "Lo más importante: El motor diagnóstico **traduce la matemática a ingeniería**. Al cruzar con "
            "frecuencias de trenes, ocupación y tiempos, no solo te dirá 'hubo sobreconsumo', sino 'por qué' ocurrió."
        )

    if not df_ops.empty:
        df_diag = diagnosticar_anomalias(df_ops, all_prmte_full, all_fact_full,
                                         df_carga_v1, df_carga_v2, df_thdr_v1, df_thdr_v2)
        if df_diag.empty:
            st.info("No hay días con energía medida en el rango para diagnosticar.")
        else:
            usa_odo = ("Odómetro [km]" in df_diag.columns) and (df_diag["Odómetro [km]"] > 0).any()
            st.caption("IDE calculado con el odómetro real (UMR)." if usa_odo
                       else "⚠ Sin UMR en el rango: el IDE puede no ser exacto.")

            c1, c2, c3 = st.columns(3)
            c1.metric("Días analizados", len(df_diag))
            c2.metric("🔴 Anomalías (Críticas)", int((df_diag["Nivel"] == "ANOMALÍA").sum()))
            c3.metric("🟠 Atención (Leves)", int((df_diag["Nivel"] == "ATENCIÓN").sum()))

            st.markdown("#### Línea base por tipo de día (mediana)")
            filas = []
            for tipo in ["L", "S", "D/F"]:
                sub = df_diag[df_diag["Tipo Día"] == tipo]
                if sub.empty:
                    continue
                filas.append({
                    "Tipo": {"L": "Laboral", "S": "Sábado", "D/F": "Dgo/Festivo"}[tipo],
                    "Días": len(sub),
                    "E. Total (kWh)": round(sub["E_Total"].median()),
                    "Tracción (kWh)": round(sub["E_Tr"].median()),
                    "12 kV (kWh)": round(sub["E_12"].median()),
                    "IDE (kWh/km)": round(sub["IDE (kWh/km)"].median(), 2),
                    "Servicios": int(sub["Servicios"].median()),
                    "Tracción doble %": (round(sub["Doble_pct"].median(), 0)
                                          if "Doble_pct" in sub and sub["Doble_pct"].notna().any() else None),
                })
            st.dataframe(pd.DataFrame(filas), use_container_width=True)

            anom = df_diag[df_diag["Nivel"] != "OK"].sort_values("Severidad", ascending=False)
            if anom.empty:
                st.success("✓ Operación Saludable: Sin desviaciones relevantes dentro de cada tipo de día.")
            else:
                st.markdown("#### Días con desviación (Narrativa de Causas Raíz)")
                for _, r in anom.iterrows():
                    icon = "🔴" if r["Nivel"] == "ANOMALÍA" else "🟠"
                    titulo = f"{icon} {pd.to_datetime(r['Fecha']).strftime('%d-%m-%Y')} ({r['Tipo Día']}) · Severidad de Falla: z={r['Severidad']:.1f}"
                    
                    with st.expander(titulo, expanded=(r["Nivel"] == "ANOMALÍA")):
                        
                        if r["Nivel"] == "ANOMALÍA": st.error(r['Diagnóstico'])
                        else: st.warning(r['Diagnóstico'])
                            
                        st.markdown("##### 🔍 Evidencia Numérica del Día")
                        a1, a2, a3, a4 = st.columns(4)
                        a1.metric("E. Total", f"{r['E_Total']:,.0f} kWh")
                        a2.metric("Tracción", f"{r['E_Tr']:,.0f} kWh")
                        a3.metric("12 kV", f"{r['E_12']:,.0f} kWh")
                        a4.metric("IDE", f"{r['IDE (kWh/km)']:,.2f} kWh/km")
                        
                        b1, b2, b3, b4 = st.columns(4)
                        noche_ok = ("Noche_kWh" in df_diag.columns) and pd.notna(r["Noche_kWh"])
                        b1.metric("Nocturno", f"{r['Noche_kWh']:,.0f} kWh" if noche_ok else "—")
                        b2.metric("Servicios", f"{int(r['Servicios']):,}")
                        b3.metric("PAX", f"{int(r['PAX']):,}" if pd.notna(r["PAX"]) else "—")
                        odo_ok = ("Odómetro [km]" in df_diag.columns) and pd.notna(r["Odómetro [km]"])
                        b4.metric("Odómetro", f"{r['Odómetro [km]']:,.0f} km" if odo_ok else "—")

                        st.markdown("##### 🚆 Análisis Logístico y Contexto")
                        cols_ctx = st.columns(2)
                        bits = []
                        if pd.notna(r.get("Doble_pct")): bits.append(f"**Uso Tracción Doble:** {r['Doble_pct']:.0f}%")
                        if r.get("Est_critica"): bits.append(f"**Punto de Saturación:** Estación {r['Est_critica']}")
                        if pd.notna(r.get("Ocup_max")): bits.append(f"**Peak Carga de Tren:** {int(r['Ocup_max']):,} PAX")
                        if pd.notna(r.get("Viaje_prom")): bits.append(f"**Viaje Promedio (Malla):** {minutos_a_hhmmss(r['Viaje_prom'])}")
                        if pd.notna(r.get("Brecha_min")): bits.append(f"**Brecha Irregularidad:** {minutos_a_hhmmss(r['Brecha_min'])}")
                            
                        if bits:
                            with cols_ctx[0]:
                                for bit in bits[:3]: st.markdown(f"- {bit}")
                            with cols_ctx[1]:
                                for bit in bits[3:]: st.markdown(f"- {bit}")

                        ins = []
                        if pd.notna(r.get("Brecha_min")) and r["Brecha_min"] > 12:
                            ins.append(f"Alta inestabilidad (Brecha de {minutos_a_hhmmss(r['Brecha_min'])}). Obliga a patrones 'Stop-and-Go', justificando el alza del consumo de tracción.")
                        if ("Eficiencia" in r["Diagnóstico"] or "Sobreconsumo" in r["Diagnóstico"]) and pd.notna(r.get("Doble_pct")) and r["Doble_pct"] > 25:
                            ins.append(f"Despacho elevado de Tracción Doble ({r['Doble_pct']:.0f}%). Más toneladas inerciales movilizadas impactan el indicador de kWh/km.")
                        if ("Volumen" in r["Diagnóstico"] or "Oferta" in r["Diagnóstico"]) and r.get("Est_critica"):
                            ins.append(f"La fricción de red (cuello de botella) se concentró fuertemente en {r['Est_critica']}.")
                        
                        if ins:
                            st.info("💡 **Apreciación de Ingeniería:** " + " ".join(ins))

            st.markdown("#### Tabla completa del diagnóstico")
            cols_show = ["Fecha", "Tipo Día", "E_Total", "E_Tr", "E_12", "Noche_kWh",
                         "IDE (kWh/km)", "Servicios", "Doble_pct", "Ocup_max", "Est_critica",
                         "Viaje_prom", "Brecha_min", "PAX", "Nivel", "Diagnóstico"]
            cols_show = [c for c in cols_show if c in df_diag.columns]
            tabla = df_diag[cols_show].copy()
            tabla["Fecha"] = pd.to_datetime(tabla["Fecha"]).dt.strftime("%Y-%m-%d")
            for c in ["Doble_pct", "Viaje_prom", "Brecha_min"]:
                if c in tabla.columns:
                    tabla[c] = pd.to_numeric(tabla[c], errors="coerce").round(1)
            st.dataframe(make_columns_unique(tabla), use_container_width=True)
    else:
        st.info("📂 Sube archivos desde el panel lateral para generar el diagnóstico.")
