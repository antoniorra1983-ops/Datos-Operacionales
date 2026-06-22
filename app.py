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

# Códigos de 3 letras usados en los export de Carga de Pasajeros (encabezados de estación)
PAX_COL_CODE = {
    'Puerto':'PUE', 'Bellavista':'BEL', 'Francia':'FRA', 'Baron':'BAR', 'Portales':'POR',
    'Recreo':'REC', 'Miramar':'MIR', 'Viña del Mar':'VIN', 'Hospital':'HOS', 'Chorrillos':'CHO',
    'El Salto':'SLT', 'Valencia':'VAL', 'Quilpue':'QUI', 'El Sol':'SOL', 'El Belloto':'BTO',
    'Las Americas':'AME', 'La Concepcion':'CON', 'Villa Alemana':'VAM', 'Sargento Aldea':'SGA',
    'Peñablanca':'PEN', 'Limache':'LIM'
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

def _srv_clean_series(s):
    # Limpia el número de servicio de forma robusta: usa el valor numérico tal cual
    # (evita que floats como 1.0 se vuelvan "10" al quitar el punto) y solo extrae
    # dígitos como respaldo para textos tipo 'V123'.
    num = pd.to_numeric(s, errors='coerce')
    if num.isna().any():
        resto = pd.to_numeric(pd.Series(s).astype(str).str.replace(r'\D', '', regex=True), errors='coerce')
        num = num.fillna(resto)
    return num

def clasificar_od_thdr(df_thdr):
    # Clasifica cada servicio por su patrón Origen->Destino usando las horas por estación.
    # Origen = estación con salida y sin llegada; Destino = estación con llegada y sin salida.
    if df_thdr is None or df_thdr.empty:
        return pd.Series(dtype=object)
    sal = pd.DataFrame(index=df_thdr.index)
    lle = pd.DataFrame(index=df_thdr.index)
    for est in ESTACIONES:
        cs = get_col_thdr(df_thdr, est, 'SALIDA')
        cl = get_col_thdr(df_thdr, est, 'LLEGADA')
        sal[est] = pd.to_numeric(df_thdr[cs], errors='coerce') if cs else np.nan
        lle[est] = pd.to_numeric(df_thdr[cl], errors='coerce') if cl else np.nan
    def _od(i):
        s = sal.loc[i]; l = lle.loc[i]
        orig = [e for e in ESTACIONES if pd.notna(s[e]) and pd.isna(l[e])]
        dest = [e for e in ESTACIONES if pd.notna(l[e]) and pd.isna(s[e])]
        o = orig[0] if orig else (s.dropna().sort_values().index[0] if s.notna().any() else None)
        d = dest[-1] if dest else (l.dropna().sort_values().index[-1] if l.notna().any() else None)
        if o is None or d is None:
            return pd.NA
        return f"{SHORT_NAMES_DICT.get(o, o)}→{SHORT_NAMES_DICT.get(d, d)}"
    return df_thdr.index.to_series().apply(_od)

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
# --- Servicios por tipo de tren (material rodante) + export Excel ---
def _no_huecos(_fig, _fmt='%d-%m-%y'):
    """Eje X de fechas categorico: oculta dias/meses/anios sin datos (sin huecos),
    etiquetas dd-mm-aa en orden cronologico. Sirve para px y go (alinea trazas extra)."""
    try:
        _pares = []
        for _tr in _fig.data:
            _xv = getattr(_tr, 'x', None)
            if _xv is None:
                continue
            _dt = pd.to_datetime(pd.Series(list(_xv)), errors='coerce')
            _lab = _dt.dt.strftime(_fmt)
            _tr.x = tuple('' if pd.isna(_l) else _l for _l in _lab)
            for _d, _l in zip(_dt, _lab):
                if pd.notna(_d):
                    _pares.append((_d, _l))
        if _pares:
            _pares.sort(key=lambda _p: _p[0])
            _seen = set(); _orden = []
            for _d, _l in _pares:
                if _l not in _seen:
                    _seen.add(_l); _orden.append(_l)
            _fig.update_xaxes(type='category', categoryorder='array', categoryarray=_orden)
    except Exception:
        pass
    return _fig
try:
    import plotly.io as _pio
    for _tn in ("plotly", "plotly_white", "simple_white", "none", "ggplot2", "seaborn", "plotly_dark", "presentation"):
        if _tn in _pio.templates: _pio.templates[_tn].layout.separators = ",."
except Exception:
    pass

def _ncl(x, dec=0):
    try: v = float(x)
    except (ValueError, TypeError): return str(x)
    s = f"{v:.{dec}f}"
    ent, _, frac = s.partition(".")
    neg = ent.startswith("-"); ent = ent.lstrip("-")
    grp = ""
    while len(ent) > 3:
        grp = "." + ent[-3:] + grp; ent = ent[:-3]
    ent = ent + grp
    out = ent + ("," + frac if frac else "")
    return ("-" if neg else "") + out

def _tipo_tren(m1):
    n = pd.to_numeric(m1, errors='coerce')
    if pd.isna(n): return "Sin asignar"
    n = int(n)
    if 1 <= n <= 27:  return "XT-100"     # unidades M01-M27
    if 28 <= n <= 35: return "XT-M"       # unidades XM28-XM35
    return "SFE"                          # otras unidades (SFE siempre simple)

def _col_tt(d, name):
    return next((c for c in d.columns if str(c).strip().upper() == name), None)

def _servicios_norm(df_thdr_v1, df_thdr_v2, fechas=None):
    parts = []
    for t in (df_thdr_v1, df_thdr_v2):
        if t is None or getattr(t, 'empty', True): continue
        x = t.copy()
        if fechas is not None and 'Fecha_Op' in x.columns:
            x = x[pd.to_datetime(x['Fecha_Op']).dt.normalize().isin(pd.to_datetime(list(fechas)))]
        if x.empty: continue
        # Tipo de servicio = patrón Origen->Destino real de la malla (todos los patrones)
        if 'Tipo_Servicio' in x.columns:
            tserv = x['Tipo_Servicio']
        else:
            try: tserv = clasificar_od_thdr(x)
            except Exception: tserv = pd.Series(pd.NA, index=x.index)
        tserv = pd.Series(tserv, index=x.index).astype(object)
        tserv = tserv.where(tserv.notna(), 'Sin clasificar')
        c1, c2 = _col_tt(x, 'MOTRIZ 1'), _col_tt(x, 'MOTRIZ 2')
        ct, cv = _col_tt(x, 'TREN'), _col_tt(x, 'VIAJE')
        df = pd.DataFrame({'Fecha': pd.to_datetime(x['Fecha_Op']).dt.date.values})
        df['Tipo de servicio'] = tserv.values
        df['N. Servicio'] = x[ct].values if ct else None
        df['N. Viaje']    = x[cv].values if cv else None
        df['Motriz 1']    = pd.to_numeric(x[c1], errors='coerce').values if c1 else np.nan
        df['Motriz 2']    = pd.to_numeric(x[c2], errors='coerce').values if c2 else np.nan
        df['Tipo de tren'] = df['Motriz 1'].apply(_tipo_tren)
        doble = df['Motriz 2'].notna().values
        if 'Unidad' in x.columns:
            doble = doble | x['Unidad'].astype(str).str.upper().str.strip().eq('M').values
        df['Composicion'] = np.where(doble, 'Doble', 'Simple')
        df.loc[df['Tipo de tren'] == 'SFE', 'Composicion'] = 'Simple'
        parts.append(df)
    if not parts: return pd.DataFrame()
    cols = ['Fecha', 'Tipo de servicio', 'N. Servicio', 'N. Viaje', 'Tipo de tren', 'Composicion', 'Motriz 1', 'Motriz 2']
    return pd.concat(parts, ignore_index=True)[cols].sort_values(['Fecha', 'Tipo de servicio', 'N. Servicio']).reset_index(drop=True)

def detalle_servicios(df_thdr_v1, df_thdr_v2, fechas=None):
    return _servicios_norm(df_thdr_v1, df_thdr_v2, fechas)

def excel_servicios(detalle):
    por_dia = (detalle.groupby(['Fecha', 'Tipo de servicio', 'Tipo de tren', 'Composicion'])
               .size().reset_index(name='Servicios'))
    resumen = (detalle.groupby(['Tipo de servicio', 'Tipo de tren', 'Composicion'])
               .size().reset_index(name='Servicios'))
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as w:
        por_dia.to_excel(w, sheet_name='Por dia', index=False)
        resumen.to_excel(w, sheet_name='Resumen', index=False)
        detalle.to_excel(w, sheet_name='Detalle', index=False)
    return buf.getvalue()

def _fmt_mmss(m):
    if m is None or (isinstance(m, float) and np.isnan(m)) or pd.isna(m): return "—"
    m = float(m); _h = int(m // 60); _mn = int(m % 60); _s = int(round((m - int(m)) * 60))
    if _s == 60: _s = 0; _mn += 1
    if _mn == 60: _mn = 0; _h += 1
    return f"{_h:02d}:{_mn:02d}:{_s:02d}"

def _od_y_dur(df_thdr):
    if df_thdr is None or getattr(df_thdr, 'empty', True):
        return pd.DataFrame(columns=['OD', 'Dur'])
    _ests = ESTACIONES
    _S = np.full((len(df_thdr), len(_ests)), np.nan)
    _L = np.full((len(df_thdr), len(_ests)), np.nan)
    for _j, _e in enumerate(_ests):
        _cs = get_col_thdr(df_thdr, _e, 'SALIDA'); _cl = get_col_thdr(df_thdr, _e, 'LLEGADA')
        if _cs: _S[:, _j] = pd.to_numeric(df_thdr[_cs], errors='coerce').values
        if _cl: _L[:, _j] = pd.to_numeric(df_thdr[_cl], errors='coerce').values
    _ods = []; _durs = []
    for _r in range(len(df_thdr)):
        _s = _S[_r]; _l = _L[_r]
        _orig = [k for k in range(len(_ests)) if not np.isnan(_s[k]) and np.isnan(_l[k])]
        _dest = [k for k in range(len(_ests)) if not np.isnan(_l[k]) and np.isnan(_s[k])]
        _oi = _orig[0] if _orig else (int(np.nanargmin(_s)) if not np.isnan(_s).all() else None)
        _di = _dest[-1] if _dest else (int(np.nanargmax(_l)) if not np.isnan(_l).all() else None)
        if _oi is None or _di is None:
            _ods.append(pd.NA); _durs.append(np.nan); continue
        _ods.append(f"{SHORT_NAMES_DICT.get(_ests[_oi], _ests[_oi])}→{SHORT_NAMES_DICT.get(_ests[_di], _ests[_di])}")
        _d = _l[_di] - _s[_oi]
        if not np.isnan(_d) and _d < -500: _d += 1440
        _durs.append(_d if (not np.isnan(_d) and 1 <= _d <= 180) else np.nan)
    return pd.DataFrame({'OD': _ods, 'Dur': _durs}, index=df_thdr.index)

def tiempos_servicios(df_thdr_v1, df_thdr_v2, fechas=None):
    parts = []
    for t in (df_thdr_v1, df_thdr_v2):
        if t is None or getattr(t, 'empty', True): continue
        x = t.copy()
        if fechas is not None and 'Fecha_Op' in x.columns:
            x = x[pd.to_datetime(x['Fecha_Op']).dt.normalize().isin(pd.to_datetime(list(fechas)))]
        if x.empty: continue
        _r = _od_y_dur(x)
        if 'Tipo_Servicio' in x.columns:
            _ts = pd.Series(x['Tipo_Servicio'].values, index=x.index).astype(object)
            _ts = _ts.where(_ts.notna(), _r['OD'])
        else:
            _ts = _r['OD']
        c1 = _col_tt(x, 'MOTRIZ 1'); c2 = _col_tt(x, 'MOTRIZ 2')
        df = pd.DataFrame({'Fecha': pd.to_datetime(x['Fecha_Op']).dt.date.values})
        df['Tipo de servicio'] = pd.Series(_ts).where(pd.Series(_ts).notna(), 'Sin clasificar').values
        df['Dur'] = _r['Dur'].values
        _m1 = pd.to_numeric(x[c1], errors='coerce') if c1 else pd.Series(np.nan, index=x.index)
        df['Tipo de tren'] = _m1.apply(_tipo_tren).values
        _doble = pd.to_numeric(x[c2], errors='coerce').notna() if c2 else pd.Series(False, index=x.index)
        if 'Unidad' in x.columns:
            _doble = _doble | x['Unidad'].astype(str).str.upper().str.strip().eq('M')
        df['Composicion'] = np.where(_doble.values, 'Doble', 'Simple')
        df.loc[df['Tipo de tren'] == 'SFE', 'Composicion'] = 'Simple'
        parts.append(df)
    if not parts: return pd.DataFrame()
    out = pd.concat(parts, ignore_index=True)
    return out[out['Dur'].notna()].reset_index(drop=True)

def _stats_dur(df, col):
    cols = col if isinstance(col, list) else [col]
    g = df.dropna(subset=['Dur']).groupby(cols, sort=False)['Dur']
    return g.agg(['mean', 'median', 'max', 'min', 'count']).reset_index()

def _render_tv_cards(stats, col, badge=None, badge_col=None):
    filas = stats.to_dict('records')
    for j in range(0, len(filas), 3):
        chunk = filas[j:j + 3]
        cols = st.columns(3)
        for k, r in enumerate(chunk):
            if isinstance(badge_col, list):
                _bk = [_bc for _bc in badge_col if r.get(_bc) not in (None, '')]
                bdg = ''.join(f'<span class="tv-badge{"" if _i == 0 else " tv-badge-alt"}">{r[_bc]}</span>' for _i, _bc in enumerate(_bk))
            else:
                _b = r[badge_col] if badge_col else badge
                bdg = f'<span class="tv-badge">{_b}</span>' if _b else ''
            html = (f'<div class="tv-card"><div class="tv-head">{r[col]} {bdg}</div>'
                    f'<div class="tv-grid">'
                    f'<div class="tv-stat"><div class="tv-lbl">Promedio</div><div class="tv-val">{_fmt_mmss(r["mean"])}</div></div>'
                    f'<div class="tv-stat"><div class="tv-lbl">Mediana</div><div class="tv-val">{_fmt_mmss(r["median"])}</div></div>'
                    f'<div class="tv-stat"><div class="tv-lbl">Máxima</div><div class="tv-val">{_fmt_mmss(r["max"])}</div></div>'
                    f'<div class="tv-stat"><div class="tv-lbl">Mínima</div><div class="tv-val">{_fmt_mmss(r["min"])}</div></div>'
                    f'</div><div class="tv-foot">N = {_ncl(int(r["count"]), 0)} servicios</div></div>')
            cols[k].markdown(html, unsafe_allow_html=True)

def _mat_sal_lle(df_thdr):
    _n = len(df_thdr); _ne = len(ESTACIONES)
    _S = np.full((_n, _ne), np.nan); _L = np.full((_n, _ne), np.nan)
    for _j, _e in enumerate(ESTACIONES):
        _cs = get_col_thdr(df_thdr, _e, 'SALIDA'); _cl = get_col_thdr(df_thdr, _e, 'LLEGADA')
        if _cs: _S[:, _j] = pd.to_numeric(df_thdr[_cs], errors='coerce').values
        if _cl: _L[:, _j] = pd.to_numeric(df_thdr[_cl], errors='coerce').values
    return _S, _L

def _dwell_estaciones(v1, v2):
    parts = []
    for t in (v1, v2):
        if t is None or getattr(t, 'empty', True): continue
        _S, _L = _mat_sal_lle(t)
        for _j, _e in enumerate(ESTACIONES):
            d = _S[:, _j] - _L[:, _j]
            d = d[~np.isnan(d)]; d = d[(d >= 0) & (d <= 30)]
            if len(d): parts.append(pd.DataFrame({'Estacion': SHORT_NAMES_DICT.get(_e, _e), 'Dur': d}))
    if not parts: return pd.DataFrame(columns=['Estacion', 'Dur'])
    return pd.concat(parts, ignore_index=True)

def _segmentos(v1, v2):
    parts = []; E = ESTACIONES
    if v1 is not None and not getattr(v1, 'empty', True):
        _S, _L = _mat_sal_lle(v1)
        for _j in range(len(E) - 1):
            d = _L[:, _j + 1] - _S[:, _j]
            d = d[~np.isnan(d)]; d = d[(d > 0) & (d <= 60)]
            if len(d): parts.append(pd.DataFrame({'Segmento': f"{SHORT_NAMES_DICT.get(E[_j], E[_j])}→{SHORT_NAMES_DICT.get(E[_j + 1], E[_j + 1])}", 'Dur': d}))
    if v2 is not None and not getattr(v2, 'empty', True):
        _S, _L = _mat_sal_lle(v2)
        for _j in range(len(E) - 1):
            d = _L[:, _j] - _S[:, _j + 1]
            d = d[~np.isnan(d)]; d = d[(d > 0) & (d <= 60)]
            if len(d): parts.append(pd.DataFrame({'Segmento': f"{SHORT_NAMES_DICT.get(E[_j + 1], E[_j + 1])}→{SHORT_NAMES_DICT.get(E[_j], E[_j])}", 'Dur': d}))
    if not parts: return pd.DataFrame(columns=['Segmento', 'Dur'])
    return pd.concat(parts, ignore_index=True)

def _diagrama_marey(v1, v2):
    """Diagrama tiempo-distancia (Marey / string-line). Cada servicio es una linea:
    via1 sube (Puerto->Limache), via2 baja (Limache->Puerto). Donde una linea de via1
    cruza una de via2 hay un cruzamiento; el marcador muestra km, servicios y recorrido."""
    import plotly.graph_objects as _go
    _N = len(ESTACIONES)
    _kmacum = np.array([float(_k) for _k in KM_ACUM[:_N]], dtype=float)
    _ix = np.arange(_N, dtype=float)
    _fig = _go.Figure()
    _COL = {'V1': '#005195', 'V2': '#E85500'}
    _NOM = {'V1': 'Vía 1 (Puerto→Limache)', 'V2': 'Vía 2 (Limache→Puerto)'}
    _serv = {'V1': [], 'V2': []}

    def _meta(_row, _name):
        if _name in _row.index and pd.notna(_row[_name]):
            try:
                return str(int(float(_row[_name])))
            except Exception:
                return str(_row[_name])
        return "?"

    for _via, _t in (('V1', v1), ('V2', v2)):
        if _t is None or getattr(_t, 'empty', True):
            continue
        _S, _L = _mat_sal_lle(_t)
        _xs = []; _ys = []
        for _r in range(len(_t)):
            _pts = []
            for _j in range(_N):
                _l = _L[_r, _j]; _s = _S[_r, _j]
                if not np.isnan(_l): _pts.append((_l, _j))
                if not np.isnan(_s): _pts.append((_s, _j))
            if len(_pts) < 2:
                continue
            _pts.sort(key=lambda _p: _p[0])
            for _p in _pts:
                _xs.append(_p[0]); _ys.append(_p[1])
            _xs.append(None); _ys.append(None)
            _row = _t.iloc[_r]
            _o = int(_pts[0][1]); _d = int(_pts[-1][1])
            _od = f"{SHORT_NAMES_DICT.get(ESTACIONES[_o], ESTACIONES[_o])}→{SHORT_NAMES_DICT.get(ESTACIONES[_d], ESTACIONES[_d])}"
            _serv[_via].append({
                't': np.array([_q[0] for _q in _pts], dtype=float),
                'y': np.array([_q[1] for _q in _pts], dtype=float),
                'tmin': _pts[0][0], 'tmax': _pts[-1][0],
                'num': _meta(_row, 'Viaje'), 'tren': _meta(_row, 'Tren'), 'od': _od})
        if _xs:
            _fig.add_trace(_go.Scatter(x=_xs, y=_ys, mode='lines',
                line=dict(color=_COL[_via], width=1.1), opacity=0.55,
                name=_NOM[_via], legendgroup=_via, hoverinfo='skip', connectgaps=False))

    if not _serv['V1'] and not _serv['V2']:
        return None, []

    _cx = []; _cy = []; _ctxt = []; _cruces = []
    for _a in _serv['V1']:
        for _b in _serv['V2']:
            _tlo = max(_a['tmin'], _b['tmin']); _thi = min(_a['tmax'], _b['tmax'])
            if _tlo >= _thi:
                continue
            _tg = np.linspace(_tlo, _thi, 60)
            _y1 = np.interp(_tg, _a['t'], _a['y'])
            _y2 = np.interp(_tg, _b['t'], _b['y'])
            _dif = _y1 - _y2
            _chg = np.where(np.diff(np.sign(_dif)) != 0)[0]
            if len(_chg) == 0:
                continue
            _k = int(_chg[0])
            _den = _dif[_k + 1] - _dif[_k]
            _tcr = _tg[_k] if _den == 0 else _tg[_k] - _dif[_k] * (_tg[_k + 1] - _tg[_k]) / _den
            _ycr = float(np.interp(_tcr, _a['t'], _a['y']))
            _kmcr = float(np.interp(_ycr, _ix, _kmacum))
            _lo = int(np.clip(np.floor(_ycr), 0, _N - 1)); _hi = int(np.clip(np.ceil(_ycr), 0, _N - 1))
            if _lo == _hi:
                _tramo = SHORT_NAMES_DICT.get(ESTACIONES[_lo], ESTACIONES[_lo])
            else:
                _tramo = f"{SHORT_NAMES_DICT.get(ESTACIONES[_lo], ESTACIONES[_lo])}–{SHORT_NAMES_DICT.get(ESTACIONES[_hi], ESTACIONES[_hi])}"
            _hh = int(_tcr) // 60; _mm = int(_tcr) % 60
            _cx.append(_tcr); _cy.append(_ycr)
            _ctxt.append(
                f"<b>Cruce {_hh:02d}:{_mm:02d}</b><br>"
                f"Km ≈ {_kmcr:.1f} · tramo {_tramo}<br>"
                f"<b>Vía 1</b> → Viaje {_a['num']} · Tren {_a['tren']} · {_a['od']}<br>"
                f"<b>Vía 2</b> → Viaje {_b['num']} · Tren {_b['tren']} · {_b['od']}")
            _cruces.append({'hora': f"{_hh:02d}:{_mm:02d}", 'hora_h': int(_hh), 'km': round(_kmcr, 1), 'tramo': _tramo})

    if _cx:
        _fig.add_trace(_go.Scatter(x=_cx, y=_cy, mode='markers',
            marker=dict(size=7, color='#15803d', symbol='x', line=dict(color='#ffffff', width=0.5)),
            name=f'Cruzamientos ({len(_cx)})', legendgroup='CX',
            text=_ctxt, hovertemplate='%{text}<extra></extra>'))

    _fig.update_yaxes(tickmode='array', tickvals=list(range(_N)),
                      ticktext=[SHORT_NAMES_DICT.get(_e, _e) for _e in ESTACIONES],
                      title='Estación', gridcolor='#eef2f7')
    _tk = list(range(0, 1441, 60))
    _fig.update_xaxes(tickmode='array', tickvals=_tk,
                      ticktext=[f"{_m // 60:02d}:{_m % 60:02d}" for _m in _tk],
                      title='Hora del día', gridcolor='#eef2f7')
    _fig.update_layout(height=640, margin=dict(t=30, b=0, l=0, r=0), hovermode='closest',
                       legend=dict(orientation='h', yanchor='bottom', y=1.02, x=0))
    return _fig, _cruces

def _fechas_thdr(v1, v2):
    _f = set()
    for _t in (v1, v2):
        if _t is not None and not getattr(_t, 'empty', True) and 'Fecha_Op' in _t.columns:
            _f |= set(pd.to_datetime(_t['Fecha_Op'], errors='coerce').dropna().dt.date.tolist())
    return sorted(_f)

def _filtra_fecha_op(df, fecha):
    if df is None or getattr(df, 'empty', True) or 'Fecha_Op' not in df.columns:
        return df
    return df[pd.to_datetime(df['Fecha_Op'], errors='coerce').dt.date == fecha]

def _km_por_tren(all_tr):
    """Desde all_tr (Tren, Fecha, Valor) separa odómetro acumulado (valores grandes) y
    kilometraje diario (valores chicos) de la hoja UMR, y resume km recorridos por tren.
    Devuelve (resumen_por_tren, km_diario_flota, pivote_km_diario)."""
    _vacio = (pd.DataFrame(), pd.DataFrame(), pd.DataFrame())
    if not all_tr:
        return _vacio
    _d = pd.DataFrame(all_tr)
    if not {'Tren', 'Fecha', 'Valor'}.issubset(_d.columns):
        return _vacio
    _d['Fecha'] = pd.to_datetime(_d['Fecha'], errors='coerce').dt.normalize()
    _d['Valor'] = pd.to_numeric(_d['Valor'], errors='coerce')
    _d = _d.dropna(subset=['Fecha', 'Valor'])
    if _d.empty:
        return _vacio
    _km = _d[_d['Valor'] < 50000]
    _odo = _d[_d['Valor'] >= 50000]
    if _km.empty:
        return _vacio
    _pivk = _km.pivot_table(index='Tren', columns='Fecha', values='Valor', aggfunc='mean').sort_index(axis=1)
    _res = pd.DataFrame(index=_pivk.index)
    _res['Km recorridos'] = _pivk.sum(axis=1, min_count=1).round(0)
    _res['Días activos'] = (_pivk > 0).sum(axis=1)
    _res['Km/día (prom)'] = (_res['Km recorridos'] / _res['Días activos'].replace(0, np.nan)).round(0)
    _res['Km máx (día)'] = _pivk.max(axis=1).round(0)
    if not _odo.empty:
        _res['Odómetro actual'] = _odo.sort_values('Fecha').groupby('Tren')['Valor'].last()
    _res = _res.reset_index().sort_values('Km recorridos', ascending=False).reset_index(drop=True)
    def _tipo_id(_s):
        _s = str(_s).upper().strip()
        if _s.startswith('SFE'): return 'SFE'
        if _s.startswith('XM'): return 'XT-M'
        if _s.startswith('M'): return 'XT-100'
        return 'Otro'
    _res.insert(1, 'Tipo', _res['Tren'].map(_tipo_id))
    _flota = _pivk.sum(axis=0, min_count=1).reset_index()
    _flota.columns = ['Fecha', 'Km flota']
    return _res, _flota, _pivk

def _orden_serv_key(s):
    s = str(s)
    if '→' not in s: return (9, 99, s)
    _o, _d = s.split('→', 1)
    _PR = {'LIM': 0, 'S.ALD': 1, 'E.BEL': 2}
    if _o == 'PUE': return (0, _PR.get(_d, 50), _d)
    if _d == 'PUE': return (1, _PR.get(_o, 50), _o)
    return (2, 0, s)

def _ordenar_serv(stats, servcol):
    if stats is None or stats.empty: return stats
    k = list(stats[servcol].map(_orden_serv_key))
    s = stats.copy()
    s['_g'] = [x[0] for x in k]; s['_p'] = [x[1] for x in k]; s['_l'] = [x[2] for x in k]
    return s.sort_values(['_g', '_p', '_l']).drop(columns=['_g', '_p', '_l'])

def _ordenar_serv_tren(stats, servcol, trencol):
    if stats is None or stats.empty: return stats
    _TR = {'XT-100': 0, 'XT-M': 1, 'SFE': 2, 'Sin asignar': 3}
    k = list(stats[servcol].map(_orden_serv_key))
    s = stats.copy()
    s['_g'] = [x[0] for x in k]; s['_p'] = [x[1] for x in k]; s['_l'] = [x[2] for x in k]
    s['_t'] = list(stats[trencol].map(lambda x: _TR.get(x, 9)))
    return s.sort_values(['_g', '_p', '_l', '_t']).drop(columns=['_g', '_p', '_l', '_t'])

def _ordenar_serv_comp(stats, servcol, compcol):
    if stats is None or stats.empty: return stats
    _CP = {'Simple': 0, 'Doble': 1}
    k = list(stats[servcol].map(_orden_serv_key))
    s = stats.copy()
    s['_g'] = [x[0] for x in k]; s['_p'] = [x[1] for x in k]; s['_l'] = [x[2] for x in k]
    s['_c'] = list(stats[compcol].map(lambda x: _CP.get(x, 9)))
    return s.sort_values(['_g', '_p', '_l', '_c']).drop(columns=['_g', '_p', '_l', '_c'])

def _ordenar_serv_tren_comp(stats, servcol, trencol, compcol):
    if stats is None or stats.empty: return stats
    _TR = {'XT-100': 0, 'XT-M': 1, 'SFE': 2, 'Sin asignar': 3}
    _CP = {'Simple': 0, 'Doble': 1}
    k = list(stats[servcol].map(_orden_serv_key))
    s = stats.copy()
    s['_g'] = [x[0] for x in k]; s['_p'] = [x[1] for x in k]; s['_l'] = [x[2] for x in k]
    s['_t'] = list(stats[trencol].map(lambda x: _TR.get(x, 9)))
    s['_c'] = list(stats[compcol].map(lambda x: _CP.get(x, 9)))
    return s.sort_values(['_g', '_p', '_l', '_t', '_c']).drop(columns=['_g', '_p', '_l', '_t', '_c'])

def _thdr_filtros():
    v1 = st.session_state.get('df_thdr_v1', pd.DataFrame())
    v2 = st.session_state.get('df_thdr_v2', pd.DataFrame())
    _f = []
    for t in (v1, v2):
        if t is not None and not getattr(t, 'empty', True) and 'Fecha_Op' in t.columns:
            _f.append(pd.to_datetime(t['Fecha_Op'], errors='coerce'))
    if not _f: return v1, v2
    _all = pd.concat(_f).dropna()
    if _all.empty: return v1, v2
    _fmin = _all.dt.date.min(); _fmax = _all.dt.date.max()
    _anios = sorted({d.year for d in _all.dt.date}, reverse=True)
    _MES_t = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    _op_mes_t = ["Todos"] + [_MES_t[mn - 1] for mn in sorted(set(_all.dt.month))]
    _mesnum_t = {_MES_t[i]: i + 1 for i in range(12)}
    _dts = sorted(set(_all.dt.date))
    _iso = pd.to_datetime(pd.Series(_dts)).dt.isocalendar()
    _wkdf = pd.DataFrame({'d': _dts, 'iy': _iso['year'].values, 'iw': _iso['week'].values})
    _smap = {}
    for (_iy, _iw), _g in _wkdf.groupby(['iy', 'iw']):
        _smap[f"Sem {int(_iw):02d} ({_g['d'].min().strftime('%d/%m')}–{_g['d'].max().strftime('%d/%m')})"] = (int(_iy), int(_iw))
    _JC = {"Laboral": "L", "Sábado": "S", "Domingo/Festivo": "D/F"}
    _tsig = (str(_fmin), str(_fmax), ",".join(map(str, _anios)))
    if st.session_state.get('_t_sig') != _tsig:
        st.session_state['_t_sig'] = _tsig
        for _k in ('_t_anio', '_t_mes', '_t_sem', '_t_jor', '_t_fec'): st.session_state.pop(_k, None)
    with st.container(border=True):
        st.markdown("**🎛️ Filtros THDR** — Año · Semana · Jornada · Fecha")
        _c1, _cmm, _c2, _c3, _c4 = st.columns([0.9, 1.2, 1.7, 2, 1.7])
        with _c1: _a = st.selectbox("Año", ["Todos"] + [str(x) for x in _anios], key="_t_anio")
        with _cmm: _me = st.selectbox("Mes", _op_mes_t, key="_t_mes")
        with _c2: _sw = st.selectbox("Semana", ["Todas"] + list(_smap.keys()), key="_t_sem")
        with _c3:
            _oj = list(_JC.keys())
            _j = st.pills("Tipo de jornada", _oj, selection_mode="multi", default=_oj, key="_t_jor") if hasattr(st, "pills") else st.multiselect("Tipo de jornada", _oj, default=_oj, key="_t_jor")
            _j = _j or _oj
        with _c4:
            _rg = st.date_input("Fecha", value=(_fmin, _fmax), min_value=_fmin, max_value=_fmax, key="_t_fec")
            _fi, _fe = (_rg[0], _rg[1]) if isinstance(_rg, tuple) and len(_rg) == 2 else (_rg, _rg)
    _cods = {_JC[x] for x in _j}
    def _mk(serie):
        f = pd.to_datetime(serie, errors='coerce')
        mm = f.notna() & (f.dt.date >= _fi) & (f.dt.date <= _fe)
        if _a != "Todos": mm = mm & (f.dt.year == int(_a))
        if _me != "Todos": mm = mm & (f.dt.month == _mesnum_t[_me])
        if _sw != "Todas":
            _iy2, _iw2 = _smap[_sw]; _ic = f.dt.isocalendar()
            mm = mm & (_ic['year'] == _iy2) & (_ic['week'] == _iw2)
        if len(_cods) < 3:
            _td = f.dt.date.map(lambda d: get_tipo_dia(d) if pd.notna(d) else None); mm = mm & _td.isin(_cods)
        return mm.values
    def _ap(t):
        if t is None or getattr(t, 'empty', True) or 'Fecha_Op' not in t.columns: return t
        return t[_mk(t['Fecha_Op'])].reset_index(drop=True)
    return _ap(v1), _ap(v2)

def _prep_noche(registros, df_ops):
    if not registros: return pd.DataFrame()
    d = pd.DataFrame(registros)
    d['Fecha'] = pd.to_datetime(d['Fecha']).dt.date
    _m = df_ops.set_index(df_ops['Fecha'].dt.date)['Tipo Día'].to_dict() if (df_ops is not None and not df_ops.empty) else {}
    d['TD'] = d['Fecha'].map(_m)
    d['TD'] = d['TD'].fillna(d['Fecha'].apply(get_tipo_dia))
    d['Hora_n'] = d['Hora'].str.slice(0, 2).astype(int)
    d['_lim'] = d['TD'].map(lambda t: 6 if t == 'L' else (7 if t == 'S' else 8))
    d = d[d['Hora_n'] < d['_lim']]
    _nom = {'L': 'Laboral', 'S': 'Sábado', 'D/F': 'Domingo/Festivo'}
    d['Tipo'] = d['TD'].map(lambda t: _nom.get(t, t))
    return d


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
    st.divider()
    if st.button("🔄 Cargar / actualizar datos", type="primary", use_container_width=True):
        st.session_state["_do_load"] = True
    st.caption("Los datos se procesan al apretar este botón; no se cargan solos al abrir.")
    
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
_CACHE_VERSION = "v20_tipo_servicio"
_cache_key = (_CACHE_VERSION, str(start_date), str(end_date),
              tuple(sorted(f.name for f in f_v1_all)), tuple(sorted(f.name for f in f_v2_all)),
              tuple(sorted(f.name for f in f_umr_all)), tuple(sorted(f.name for f in f_seat_all)),
              tuple(sorted(f.name for f in f_bill_all)),
              tuple(sorted(f.name for f in f_carga_v1_all)), tuple(sorted(f.name for f in f_carga_v2_all)))
              
_hay_archivos = any([f_v1_all,f_v2_all,f_umr_all,f_seat_all,f_bill_all,f_carga_v1_all,f_carga_v2_all])
_recalcular   = st.session_state.get('_cache_key') != _cache_key

df_ops=pd.DataFrame(); df_thdr_v1=pd.DataFrame(); df_thdr_v2=pd.DataFrame()
df_carga_v1=pd.DataFrame(); df_carga_v2=pd.DataFrame()
df_serv_tipo=pd.DataFrame(columns=['Fecha','Tipo_Servicio','Servicios','TrenKm']); df_pax_tipo=pd.DataFrame(columns=['Fecha','Tipo_Servicio','PAX'])
all_ops,all_tr,all_seat,all_fact_full,all_prmte_full=[],[],[],[],[]
all_prmte_2025=[]
all_kmserv=[]
_errores_proc={}

if _hay_archivos and 'df_ops' in st.session_state and not st.session_state.get('_do_load'):
    df_ops=st.session_state['df_ops']
    df_thdr_v1=st.session_state['df_thdr_v1']
    df_thdr_v2=st.session_state['df_thdr_v2']
    all_tr=st.session_state['all_tr']
    all_seat=st.session_state['all_seat']
    all_fact_full=st.session_state['all_fact_full']
    all_prmte_full=st.session_state['all_prmte_full']
    all_prmte_2025=st.session_state.get('all_prmte_2025',[])
    all_kmserv=st.session_state.get('all_kmserv',[])
    df_carga_v1=st.session_state.get('df_carga_v1', pd.DataFrame())
    df_carga_v2=st.session_state.get('df_carga_v2', pd.DataFrame())
    df_serv_tipo=st.session_state.get('df_serv_tipo', pd.DataFrame(columns=['Fecha','Tipo_Servicio','Servicios','TrenKm']))
    df_pax_tipo=st.session_state.get('df_pax_tipo', pd.DataFrame(columns=['Fecha','Tipo_Servicio','PAX']))

elif _hay_archivos and st.session_state.get('_do_load'):
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
                                    for k in range(i+3,min(i+100,len(df_raw))):
                                        t=str(df_raw.iloc[k,0]).strip().upper()
                                        if re.match(r'^(M|XM|SFE)',t):
                                            all_tr.append({"Tren":t,"Fecha":v_f.normalize(),"Valor":parse_latam_number(df_raw.iloc[k,j])})
                    if 'SERV' in sn.upper() and 'KM' in sn.upper():
                        _fe_kms = pd.to_datetime(df_raw.iloc[4:, 0], errors='coerce').ffill()
                        for _ri in range(4, len(df_raw)):
                            _od = df_raw.iloc[_ri, 5]
                            _fe = _fe_kms.get(_ri)
                            if pd.notna(_od) and pd.notna(_fe) and re.match(r'^[A-Z]{2,3}-[A-Z]{2,3}$', str(_od).strip().upper()) and start_date <= _fe.date() <= end_date:
                                all_kmserv.append({"Fecha": _fe.normalize(),
                                                   "KmsxTrenes": parse_latam_number(df_raw.iloc[_ri, 9]),
                                                   "KmTrenR": parse_latam_number(df_raw.iloc[_ri, 22]),
                                                   "KmTrenP": parse_latam_number(df_raw.iloc[_ri, 23])})
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
                            if pd.isna(ts): continue
                            _en_rango=(start_date<=ts.date()<=end_date); _es_2025=(ts.year==2025)
                            if not (_en_rango or _es_2025): continue
                            consumo=sum(parse_latam_number(r.get(c,0)) for c in cols_retiro)
                            _reg_p={"Fecha":ts.normalize(),"Hora":f"{ts.hour:02d}:00","15min":f"{ts.hour:02d}:{ts.minute:02d}","Consumo":consumo}
                            if _en_rango: all_prmte_full.append(_reg_p)
                            if _es_2025: all_prmte_2025.append(_reg_p)
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

    df_serv_tipo = pd.DataFrame(columns=['Fecha', 'Tipo_Servicio', 'Servicios', 'TrenKm'])
    df_pax_tipo = pd.DataFrame(columns=['Fecha', 'Tipo_Servicio', 'PAX'])

    if not df_ops.empty:
        if not df_thdr_v1.empty or not df_thdr_v2.empty:
            if not df_thdr_v1.empty:
                df_thdr_v1['Tipo_Servicio'] = clasificar_od_thdr(df_thdr_v1)
            if not df_thdr_v2.empty:
                df_thdr_v2['Tipo_Servicio'] = clasificar_od_thdr(df_thdr_v2)
            _thdr_all = pd.concat([d for d in [df_thdr_v1, df_thdr_v2] if not d.empty], ignore_index=True)
            _thdr_all = _thdr_all[_thdr_all['Tipo_Servicio'].notna()]
            if not _thdr_all.empty:
                _g = _thdr_all.groupby(['Fecha_Op', 'Tipo_Servicio'])
                df_serv_tipo = _g.size().reset_index(name='Servicios')
                if 'Tren-Km' in _thdr_all.columns:
                    _tk = _g['Tren-Km'].sum().reset_index(name='TrenKm')
                    df_serv_tipo = df_serv_tipo.merge(_tk, on=['Fecha_Op', 'Tipo_Servicio'], how='left')
                else:
                    df_serv_tipo['TrenKm'] = 0.0
                df_serv_tipo = df_serv_tipo.rename(columns={'Fecha_Op': 'Fecha'})
                df_serv_tipo['Fecha'] = pd.to_datetime(df_serv_tipo['Fecha']).dt.normalize()
            _serv_tot = df_serv_tipo.groupby('Fecha')['Servicios'].sum().reset_index() if not df_serv_tipo.empty else pd.DataFrame(columns=['Fecha', 'Servicios'])
            df_ops = df_ops.merge(_serv_tot, on='Fecha', how='left').fillna({'Servicios': 0})
        else:
            df_ops['Servicios'] = 0

        if not df_carga_v1.empty or not df_carga_v2.empty:
            _carga_all = pd.concat([d for d in [df_carga_v1, df_carga_v2] if not d.empty], ignore_index=True)
            _ct = next((c for c in _carga_all.columns if 'TOTAL' in str(c).upper() and 'BORDO' in str(c).upper()), None)
            _carga_all['_pax_val'] = pd.to_numeric(_carga_all[_ct], errors='coerce').fillna(0) if _ct else 0
            _pax_tot = _carga_all.groupby('Fecha')['_pax_val'].sum().reset_index(name='PAX')
            _pax_tot['Fecha'] = pd.to_datetime(_pax_tot['Fecha']).dt.normalize()
            df_ops = df_ops.merge(_pax_tot, on='Fecha', how='left').fillna({'PAX': 0})
            _thdr_map = pd.concat([d for d in [df_thdr_v1, df_thdr_v2] if not d.empty], ignore_index=True) if (not df_thdr_v1.empty or not df_thdr_v2.empty) else pd.DataFrame()
            if not _thdr_map.empty and 'Tipo_Servicio' in _thdr_map.columns:
                _cs_t = next((c for c in _thdr_map.columns if 'VIAJE' in str(c).upper() or 'SERV' in str(c).upper() or 'THDR' in str(c).upper()), None)
                _cs_c = next((c for c in _carga_all.columns if 'THDR' in str(c).upper() or 'VIAJE' in str(c).upper() or 'SERV' in str(c).upper()), None)
                if _cs_t and _cs_c:
                    _tm = _thdr_map[['Fecha_Op', _cs_t, 'Tipo_Servicio']].copy()
                    _tm['_srv'] = _srv_clean_series(_tm[_cs_t])
                    _tm = _tm.dropna(subset=['Tipo_Servicio']).drop_duplicates(['Fecha_Op', '_srv'])
                    _cc = _carga_all[['Fecha', _cs_c, '_pax_val']].copy()
                    _cc['_srv'] = _srv_clean_series(_cc[_cs_c])
                    _mp = _cc.merge(_tm[['Fecha_Op', '_srv', 'Tipo_Servicio']], left_on=['Fecha', '_srv'], right_on=['Fecha_Op', '_srv'], how='inner')
                    if not _mp.empty:
                        df_pax_tipo = _mp.groupby(['Fecha', 'Tipo_Servicio'])['_pax_val'].sum().reset_index(name='PAX')
                        df_pax_tipo['Fecha'] = pd.to_datetime(df_pax_tipo['Fecha']).dt.normalize()
        else:
            df_ops['PAX'] = 0

    st.session_state.update({'df_ops':df_ops,'df_thdr_v1':df_thdr_v1,'df_thdr_v2':df_thdr_v2,
                              'all_tr':all_tr,'all_seat':all_seat,'all_fact_full':all_fact_full,
                              'all_prmte_full':all_prmte_full,'all_prmte_2025':all_prmte_2025,'_cache_key':_cache_key,'all_kmserv':all_kmserv,
                              'df_carga_v1':df_carga_v1, 'df_carga_v2':df_carga_v2, 'df_serv_tipo':df_serv_tipo, 'df_pax_tipo':df_pax_tipo})
    st.session_state['_do_load'] = False

# --- 8. TABS DE VISUALIZACIÓN ---
# Avisos de estado de carga
if not _hay_archivos:
    st.info("Sube archivos en la barra lateral para comenzar.")
elif df_ops.empty and not st.session_state.get('_do_load'):
    st.info("Configura el rango y los archivos en la barra lateral y aprieta **🔄 Cargar / actualizar datos** para procesar.")
elif ('df_ops' in st.session_state) and (st.session_state.get('_cache_key') != _cache_key):
    st.warning("Cambiaron archivos o fechas desde la última carga. Aprieta **🔄 Cargar / actualizar datos** para refrescar.")

_SECCIONES = ["📊 Resumen", "📑 Operaciones", "📑 Trenes", "⚡ Energía", "⚖️ Perfil Horario & Anomalías",
              "🌙 Consumo Nocturno", "🚨 Atípicos", "📋 THDR", "🔬 Análisis Multivariante", "👥 Pasajeros", "📝 Informe Ejecutivo", "🩺 Diagnóstico de Causas", "📈 Servicios", "💡 Ahorro de energía"]
_seccion = st.radio("Sección", _SECCIONES, horizontal=True, key="_nav_seccion", label_visibility="collapsed")

# ===== BARRA DE FILTROS GLOBAL (post-carga, visible en todas las pestañas) =====
st.markdown("<style>.barra-filtros-tit{font-size:1.02rem;font-weight:700;color:#005195;margin:.1rem 0 .35rem 0}</style>", unsafe_allow_html=True)
_J_COD = {"Laboral": "L", "Sábado": "S", "Domingo/Festivo": "D/F"}
if not df_ops.empty and _seccion != _SECCIONES[7]:
    _ff = pd.to_datetime(df_ops['Fecha'], errors='coerce')
    _fmin, _fmax = _ff.min().date(), _ff.max().date()
    _anios = sorted({d.year for d in _ff.dropna().dt.date}, reverse=True)
    _MES_g = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    _op_mes_g = ["Todos"] + [_MES_g[mn - 1] for mn in sorted(set(_ff.dropna().dt.month))]
    _mesnum_g = {_MES_g[i]: i + 1 for i in range(12)}
    _dts = sorted(set(_ff.dropna().dt.date))
    _iso = pd.to_datetime(pd.Series(_dts)).dt.isocalendar()
    _wkdf = pd.DataFrame({'d': _dts, 'iy': _iso['year'].values, 'iw': _iso['week'].values})
    _sem_map = {}
    for (_iy, _iw), _g in _wkdf.groupby(['iy', 'iw']):
        _lbl = f"Sem {int(_iw):02d} ({_g['d'].min().strftime('%d/%m')}–{_g['d'].max().strftime('%d/%m')})"
        _sem_map[_lbl] = (int(_iy), int(_iw))
    _op_sem = ["Todas"] + list(_sem_map.keys())
    _sig = (str(_fmin), str(_fmax), ",".join(map(str, _anios)))
    if st.session_state.get('_f_sig') != _sig:
        st.session_state['_f_sig'] = _sig
        for _k in ('_f_anio', '_f_mes', '_f_semana', '_f_jornada', '_f_fecha'): st.session_state.pop(_k, None)
    with st.container(border=True):
        st.markdown('<div class="barra-filtros-tit">🎛️ Filtros</div>', unsafe_allow_html=True)
        _cf1, _cfm, _cf2, _cf3, _cf4 = st.columns([0.9, 1.2, 1.7, 2, 1.7])
        with _cf1:
            _sel_a = st.selectbox("Año", ["Todos"] + [str(a) for a in _anios], key="_f_anio")
        with _cfm:
            _sel_m = st.selectbox("Mes", _op_mes_g, key="_f_mes")
        with _cf2:
            _sel_s = st.selectbox("Semana", _op_sem, key="_f_semana")
        with _cf3:
            _opj = list(_J_COD.keys())
            if hasattr(st, "pills"):
                _selj = st.pills("Tipo de jornada", _opj, selection_mode="multi", default=_opj, key="_f_jornada")
            else:
                _selj = st.multiselect("Tipo de jornada", _opj, default=_opj, key="_f_jornada")
            _selj = _selj or _opj
        with _cf4:
            _rg = st.date_input("Fecha", value=(_fmin, _fmax), min_value=_fmin, max_value=_fmax, key="_f_fecha")
            _fi, _fe = (_rg[0], _rg[1]) if isinstance(_rg, tuple) and len(_rg) == 2 else (_rg, _rg)
    _cods = {_J_COD[j] for j in _selj}
    def _msk_f(serie):
        f = pd.to_datetime(serie, errors='coerce')
        m = f.notna() & (f.dt.date >= _fi) & (f.dt.date <= _fe)
        if _sel_a != "Todos": m = m & (f.dt.year == int(_sel_a))
        if _sel_m != "Todos": m = m & (f.dt.month == _mesnum_g[_sel_m])
        if _sel_s != "Todas":
            _iy2, _iw2 = _sem_map[_sel_s]
            _ic = f.dt.isocalendar()
            m = m & (_ic['year'] == _iy2) & (_ic['week'] == _iw2)
        if len(_cods) < 3:
            td = f.dt.date.map(lambda d: get_tipo_dia(d) if pd.notna(d) else None)
            m = m & td.isin(_cods)
        return m.values
    def _filt_df(df, col):
        if df is None or getattr(df, 'empty', True) or col not in df.columns: return df
        return df[_msk_f(df[col])].reset_index(drop=True)
    def _filt_regs(regs):
        if not regs: return regs
        _d = pd.DataFrame(regs)
        if 'Fecha' not in _d.columns: return regs
        return _d[_msk_f(_d['Fecha'])].to_dict('records')
    df_ops = _filt_df(df_ops, 'Fecha')
    df_thdr_v1 = _filt_df(df_thdr_v1, 'Fecha_Op')
    df_thdr_v2 = _filt_df(df_thdr_v2, 'Fecha_Op')
    df_carga_v1 = _filt_df(df_carga_v1, 'Fecha')
    df_carga_v2 = _filt_df(df_carga_v2, 'Fecha')
    df_serv_tipo = _filt_df(df_serv_tipo, 'Fecha')
    df_pax_tipo = _filt_df(df_pax_tipo, 'Fecha')
    all_prmte_full = _filt_regs(all_prmte_full)
    all_fact_full = _filt_regs(all_fact_full)
    all_tr = _filt_regs(all_tr)
    all_seat = _filt_regs(all_seat)
    all_kmserv = _filt_regs(all_kmserv)
    _rf = []
    if _sel_a != "Todos": _rf.append(f"Año {_sel_a}")
    if _sel_m != "Todos": _rf.append(_sel_m)
    if _sel_s != "Todas": _rf.append(_sel_s)
    _rf.append("Jornada: " + ("todas" if len(_cods) == 3 else ", ".join(_selj)))
    _rf.append(f"{_fi.strftime('%d-%m-%Y')} → {_fe.strftime('%d-%m-%Y')}")
    st.caption("Filtros activos · " + " · ".join(_rf) + f" · {_ncl(len(df_ops), 0)} día(s)")

if _seccion == _SECCIONES[0]:
    _ep=st.session_state.get('_errores_proc',{})
    if _ep:
        with st.expander(f"⚠️ {len(_ep)} archivo(s) con error",expanded=True):
            for _n,_m in _ep.items(): st.error(f"**{_n}**: {_m}")
    if not df_ops.empty:
        # Jornada/Año/Fecha se filtran en la barra global (arriba).
        
        df_resumen = df_ops
        
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
            
            # ====== Gráficos por tipo de servicio (uno por fila, con su tarjeta) ======
            _fechas_ok = set(df_resumen['Fecha'])
            _st = df_serv_tipo[df_serv_tipo['Fecha'].isin(_fechas_ok)].copy() if not df_serv_tipo.empty else df_serv_tipo
            _pt = df_pax_tipo[df_pax_tipo['Fecha'].isin(_fechas_ok)].copy() if not df_pax_tipo.empty else df_pax_tipo
            _SHORT_INV = {v: k for k, v in SHORT_NAMES_DICT.items()}
            _idx_est = {e: i for i, e in enumerate(ESTACIONES)}
            def _via_de(_t):
                # Sentido de circulación: origen antes que destino en la línea => Vía 1 (hacia Limache); al revés => Vía 2 (hacia Puerto).
                _p = str(_t).split("→")
                if len(_p) != 2:
                    return 'Otros'
                _o = _SHORT_INV.get(_p[0].strip(), _p[0].strip())
                _d = _SHORT_INV.get(_p[1].strip(), _p[1].strip())
                _io = _idx_est.get(_o)
                _idd = _idx_est.get(_d)
                if _io is None or _idd is None:
                    return 'Otros'
                if _io < _idd:
                    return 'Vía 1'
                if _io > _idd:
                    return 'Vía 2'
                return 'Otros'
            _orden_via = {'Vía 1': 0, 'Vía 2': 1, 'Otros': 2}
            _tot_tipo = df_serv_tipo.groupby('Tipo_Servicio')['Servicios'].sum() if not df_serv_tipo.empty else pd.Series(dtype=float)
            _tipos = sorted(df_serv_tipo['Tipo_Servicio'].dropna().unique().tolist(), key=lambda t: (_orden_via.get(_via_de(t), 9), -float(_tot_tipo.get(t, 0)), t)) if not df_serv_tipo.empty else []
            _pal_v1 = ['#005195', '#3B8BD0', '#9ECAE9', '#C6DBEF']
            _pal_v2 = ['#E85500', '#F4A06B', '#FBD0B0', '#FDE3D0']
            _pal_otros = ['#2CA02C', '#9467BD', '#8C564B']
            _cmap = {}
            _iv1 = _iv2 = _io = 0
            for _t in _tipos:
                _v = _via_de(_t)
                if _v == 'Vía 1':
                    _cmap[_t] = _pal_v1[_iv1 % len(_pal_v1)]; _iv1 += 1
                elif _v == 'Vía 2':
                    _cmap[_t] = _pal_v2[_iv2 % len(_pal_v2)]; _iv2 += 1
                else:
                    _cmap[_t] = _pal_otros[_io % len(_pal_otros)]; _io += 1
            def _card_metric(_col, _label, _value, _unit="", _tip=""):
                _u = ("<span style='font-size:.68rem;font-weight:600;color:#94a3b8;'>&nbsp;" + str(_unit) + "</span>") if _unit else ""
                _html = ("<div style='padding:2px 8px 4px 2px;' title='" + str(_tip) + "'>"
                         "<div style='font-size:.72rem;color:#64748b;font-weight:600;line-height:1.18;white-space:normal;word-break:break-word;min-height:2.4em;display:flex;align-items:flex-end;'>" + str(_label) + "</div>"
                         "<div style='font-size:1.15rem;font-weight:700;color:#0f172a;line-height:1.2;white-space:nowrap;'>" + str(_value) + _u + "</div></div>")
                _col.markdown(_html, unsafe_allow_html=True)
            def _tarjeta_por_via(_df, _col, _fmt, _label_total, _unit=""):
                _serie = _df.groupby('Tipo_Servicio')[_col].sum() if (not _df.empty and _col in _df.columns) else None
                _tipos_ord = []
                if _serie is not None:
                    for _vlbl in ['Vía 1', 'Vía 2', 'Otros']:
                        _tipos_ord += sorted([t for t in _serie.index if _via_de(t) == _vlbl], key=lambda t: _serie[t], reverse=True)
                with st.container(border=True):
                    _cols = st.columns([1.3] + [1] * len(_tipos_ord))
                    _card_metric(_cols[0], _label_total, _fmt(_df[_col].sum() if _serie is not None else 0), _unit)
                    for _i, _t2 in enumerate(_tipos_ord):
                        _card_metric(_cols[_i + 1], _t2, _fmt(_serie[_t2]), _unit, _via_de(_t2))
            _fmt_int = lambda v: f"{_ncl(int(round(v)), 0)}"
            _fmt_km = lambda v: f"{_ncl(v, 1)}"
            ev_serv = ev_pax = ev_tk = None

            def _layout_tipo(_fig):
                _fig.update_layout(margin=dict(t=10, b=80, l=0, r=0),
                                   legend=dict(title="", orientation="h", yanchor="top", y=-0.28, xanchor="center", x=0.5,
                                               font=dict(size=11)),
                                   bargap=0.28, xaxis_title=None, yaxis_title=None)
                return _fig

            # --- 1. Servicios por tipo de servicio ---
            st.markdown("**Servicios por tipo de servicio**")
            if not _st.empty:
                fig_serv = px.bar(_st, x='Fecha', y='Servicios', color='Tipo_Servicio', barmode='stack',
                                  color_discrete_map=_cmap, category_orders={'Tipo_Servicio': _tipos})
                _layout_tipo(fig_serv)
                ev_serv = st.plotly_chart(_no_huecos(fig_serv), use_container_width=True, config={'locale': 'es'}, on_select="rerun", key="chart_serv")
            else:
                st.info("Sin datos de servicios (THDR) para el filtro actual.")
            _tarjeta_por_via(_st, 'Servicios', _fmt_int, "Total Servicios")

            st.divider()

            # --- 1b. Servicios por tipo de tren (material rodante) ---
            st.markdown("**Servicios por tipo de tren (XT-100 / XT-M / SFE)**")
            _det_tren = detalle_servicios(df_thdr_v1, df_thdr_v2, df_resumen['Fecha'].unique())
            if _det_tren.empty:
                st.info("Sin datos de THDR (con Motriz) para el desglose por tipo de tren.")
            else:
                _piv = _det_tren.groupby(['Tipo de tren', 'Composicion']).size().unstack(fill_value=0)
                for _c in ['Simple', 'Doble']:
                    if _c not in _piv.columns: _piv[_c] = 0
                _piv = _piv[['Simple', 'Doble']]; _piv['Total'] = _piv['Simple'] + _piv['Doble']
                _piv = _piv.reindex([t for t in ['XT-100', 'XT-M', 'SFE', 'Sin asignar'] if t in _piv.index])
                _cta, _ctb = st.columns([1, 1.5])
                with _cta:
                    st.dataframe(_piv.rename_axis("Tipo").reset_index(), use_container_width=True, hide_index=True)
                    st.caption(f"Total: {_ncl(int(_piv['Total'].sum()), 0)} servicios")
                with _ctb:
                    _dd = _det_tren.copy()
                    _dd['Tren · Comp'] = _dd['Tipo de tren'] + " · " + _dd['Composicion']
                    _ppd = _dd.groupby(['Fecha', 'Tren · Comp']).size().reset_index(name='Servicios')
                    _fig_tren = px.bar(_ppd, x='Fecha', y='Servicios', color='Tren · Comp', barmode='stack')
                    _fig_tren.update_layout(margin=dict(t=10, b=0, l=0, r=0), height=320,
                                            legend=dict(orientation='h', y=1.18, x=0, font=dict(size=10)), xaxis_title=None)
                    st.plotly_chart(_no_huecos(_fig_tren), use_container_width=True, config={'locale': 'es'})
            st.caption("Tipo de servicio = patrón Origen→Destino de la malla. XT-100 = M01-M27 · XT-M = XM28-XM35 · SFE = otras unidades (siempre simple).")

            st.divider()

            # --- 2. Pasajeros por tipo de servicio ---
            st.markdown("**Pasajeros por tipo de servicio (PAX)**")
            if not _pt.empty:
                fig_pax = px.bar(_pt, x='Fecha', y='PAX', color='Tipo_Servicio', barmode='stack',
                                 color_discrete_map=_cmap, category_orders={'Tipo_Servicio': _tipos})
                _layout_tipo(fig_pax)
                ev_pax = st.plotly_chart(_no_huecos(fig_pax), use_container_width=True, config={'locale': 'es'}, on_select="rerun", key="chart_pax")
            else:
                st.info("Sin datos de pasajeros cruzables con la malla THDR para el filtro actual.")
            _tarjeta_por_via(_pt, 'PAX', _fmt_int, "Total PAX")

            st.divider()

            # --- 3. Tren-Km por tipo de servicio (THDR) ---
            st.markdown("**Tren-Km por tipo de servicio (THDR)**")
            if not _st.empty and 'TrenKm' in _st.columns:
                fig_tk = px.bar(_st, x='Fecha', y='TrenKm', color='Tipo_Servicio', barmode='stack',
                                color_discrete_map=_cmap, category_orders={'Tipo_Servicio': _tipos})
                _layout_tipo(fig_tk)
                ev_tk = st.plotly_chart(_no_huecos(fig_tk), use_container_width=True, config={'locale': 'es'}, on_select="rerun", key="chart_tk")
            else:
                st.info("Sin datos de Tren-Km (THDR) para el filtro actual.")
            _tarjeta_por_via(_st, 'TrenKm', _fmt_km, "Tren-Km Total", "km")
            st.caption("Tipo de servicio = patrón Origen→Destino detectado en la malla THDR. El Tren-Km usa 43,13 km por servicio (×2 en tracción doble); los servicios cortos comparten esa base, así que su Tren-Km es una cota superior, no la distancia exacta.")

            st.divider()

            # --- 4. Odómetro real (UMR) ---
            c_chart, c_card = st.columns([3, 1])
            with c_chart:
                fig_odo = px.bar(df_resumen, x='Fecha', y='Odómetro [km]',
                                 color_discrete_sequence=["#005195"],
                                 hover_data=hover_config, title="Odómetro Real (UMR)")
                fig_odo.update_traces(texttemplate='%{_ncl(y, 2)}', textposition='inside', insidetextanchor='middle', textangle=-90, textfont=dict(color='white', size=10))
                fig_odo.update_layout(margin=dict(t=50, b=0, l=0, r=0), title=dict(font=dict(size=15), automargin=True),
                                      bargap=0.2, uniformtext=dict(minsize=8, mode='hide'))
                ev_odo = st.plotly_chart(_no_huecos(fig_odo), use_container_width=True, config={'locale': 'es'}, on_select="rerun", key="chart_odo")
            with c_card:
                st.markdown("<br><br>", unsafe_allow_html=True)
                st.metric("Odómetro Total", f"{_ncl(df_resumen['Odómetro [km]'].sum(), 2)} km")

            st.divider()

            # --- 5. Tasa de acoplamiento (UMR %) ---
            c_chart, c_card = st.columns([3, 1])
            with c_chart:
                fig_umr = px.bar(df_resumen, x='Fecha', y='UMR (%)',
                                 color_discrete_sequence=["#E85500"],
                                 hover_data=hover_config, title="Tasa de Acoplamiento (UMR %)")
                fig_umr.update_traces(texttemplate='%{_ncl(y, 2)}%', textposition='inside', insidetextanchor='middle', textfont=dict(color='white', size=11))
                fig_umr.update_layout(margin=dict(t=50, b=0, l=0, r=0), title=dict(font=dict(size=15), automargin=True),
                                      bargap=0.2, uniformtext=dict(minsize=8, mode='hide'))
                ev_umr = st.plotly_chart(_no_huecos(fig_umr), use_container_width=True, config={'locale': 'es'}, on_select="rerun", key="chart_umr")
            with c_card:
                st.markdown("<br><br>", unsafe_allow_html=True)
                _tot_tk_umr = df_resumen['Tren-Km [km]'].sum()
                _tot_odo_umr = df_resumen['Odómetro [km]'].sum()
                umr_global = (_tot_tk_umr / _tot_odo_umr * 100) if _tot_odo_umr > 0 else 0
                st.metric("Tasa UMR Global", f"{_ncl(umr_global, 2)} %")

            st.divider()

            # --- 6. Consumo energético (Tracción + Baja Tensión) ---
            df_plot_ener = df_resumen.rename(columns={'E_Tr': 'Tracción', 'E_12': 'Baja Tensión'})
            c_chart, c_card = st.columns([3, 1])
            with c_chart:
                fig_ener = px.bar(df_plot_ener, x='Fecha', y=['Tracción', 'Baja Tensión'],
                                  barmode='stack',
                                  color_discrete_map={'Tracción': '#E85500', 'Baja Tensión': '#005195'},
                                  hover_data=hover_config, title="Consumo Energético (kWh)")
                fig_ener.update_traces(texttemplate='%{_ncl(y, 2)}', textposition='inside', insidetextanchor='middle', textangle=-90, textfont=dict(color='white', size=10))
                fig_ener.update_layout(margin=dict(t=50, b=0, l=0, r=0), title=dict(font=dict(size=15), automargin=True),
                                       legend=dict(title="", orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                                       bargap=0.2, uniformtext=dict(minsize=8, mode='hide'))
                ev_ener = st.plotly_chart(_no_huecos(fig_ener), use_container_width=True, config={'locale': 'es'}, on_select="rerun", key="chart_ener")
            with c_card:
                st.markdown("<br>", unsafe_allow_html=True)
                st.metric("Total Tracción", f"{_ncl(df_plot_ener['Tracción'].sum(), 2)} kWh")
                st.metric("Total Baja Tensión", f"{_ncl(df_plot_ener['Baja Tensión'].sum(), 2)} kWh")

            st.divider()

            # --- 7. Desempeño energético (IDE) ---
            c_chart, c_card = st.columns([3, 1])
            with c_chart:
                fig_ide_bar = px.bar(df_resumen, x='Fecha', y='IDE (kWh/km)',
                                     color_discrete_sequence=["#E85500"],
                                     hover_data=hover_config, title="Desempeño Energético (IDE)")
                fig_ide_bar.update_traces(texttemplate='%{_ncl(y, 2)}', textposition='inside', insidetextanchor='middle', textfont=dict(color='white', size=11))
                fig_ide_bar.update_layout(margin=dict(t=50, b=0, l=0, r=0), title=dict(font=dict(size=15), automargin=True),
                                          bargap=0.2, uniformtext=dict(minsize=8, mode='hide'))
                ev_ide_bar = st.plotly_chart(_no_huecos(fig_ide_bar), use_container_width=True, config={'locale': 'es'}, on_select="rerun", key="chart_ide")
            with c_card:
                st.markdown("<br><br>", unsafe_allow_html=True)
                _tot_tr_ide = df_resumen['E_Tr'].sum()
                _tot_odo_ide = df_resumen['Odómetro [km]'].sum()
                ide_global = (_tot_tr_ide / _tot_odo_ide) if _tot_odo_ide > 0 else 0
                st.metric("IDE Global", f"{_ncl(ide_global, 2)} kWh/km")

            chart_events = [ev_serv, ev_pax, ev_tk, ev_odo, ev_umr, ev_ener, ev_ide_bar]
            
            for ev in chart_events:
                if ev and isinstance(ev, dict) and ev.get('selection') and ev['selection'].get('points'):
                    clicked_x = ev['selection']['points'][0].get('x')
                    if clicked_x:
                        try:
                            clicked_dt = pd.to_datetime(clicked_x, dayfirst=True).normalize()
                            if st.session_state.drilldown_date != clicked_dt:
                                st.session_state.drilldown_date = clicked_dt
                                st.rerun()
                        except Exception:
                            pass
            
    else: st.info("📂 Sube archivos desde el panel lateral para ver el resumen.")

if _seccion == _SECCIONES[1]:
    if not df_ops.empty:
        dv=df_ops.copy()
        dv['Fecha'] = dv['Fecha'].dt.strftime('%Y-%m-%d')
        st.write("### Datos Consolidados de Operaciones y Energía")
        st.dataframe(make_columns_unique(dv), use_container_width=True)
    else: st.info("No hay datos de operaciones en el rango seleccionado.")

if _seccion == _SECCIONES[2]:
    if all_tr:
        _res_tr, _flota_tr, _pivk_tr = _km_por_tren(all_tr)
        if _res_tr.empty:
            st.write("### Detalle por Unidad (Tren)")
            df_tr = pd.DataFrame(all_tr)
            df_tr['Fecha'] = pd.to_datetime(df_tr['Fecha']).dt.strftime('%Y-%m-%d')
            st.dataframe(make_columns_unique(df_tr), use_container_width=True)
        else:
            st.markdown("### 🚆 Kilometraje por tren (odómetro UMR)")
            _km_tot_tr = float(_res_tr['Km recorridos'].sum())
            _act_tr = _res_tr[_res_tr['Km recorridos'] > 0]
            _n_act = int(len(_act_tr))
            _c1, _c2, _c3, _c4 = st.columns(4)
            _c1.metric("Trenes con actividad", _ncl(_n_act, 0))
            _c2.metric("Km total flota", f"{_ncl(_km_tot_tr, 0)} km")
            _c3.metric("Km promedio/tren", f"{_ncl(_km_tot_tr / _n_act if _n_act else 0, 0)} km")
            _c4.metric("Tren más usado", str(_res_tr.iloc[0]['Tren']), f"{_ncl(_res_tr.iloc[0]['Km recorridos'], 0)} km")
            _por_tipo = _res_tr.groupby('Tipo').agg(**{'Trenes': ('Tren', 'count'), 'Km total': ('Km recorridos', 'sum')}).reset_index()
            _tot_km_tp = float(_por_tipo['Km total'].sum())
            _por_tipo['% uso'] = (_por_tipo['Km total'] / _tot_km_tp * 100).round(1) if _tot_km_tp > 0 else 0.0
            _por_tipo = _por_tipo.sort_values('Km total', ascending=False).reset_index(drop=True)
            st.markdown("#### Uso por tipo de tren")
            _ct1, _ct2 = st.columns([5, 6])
            with _ct1:
                _fig_tp = px.pie(_por_tipo, names='Tipo', values='Km total', hole=0.55,
                                 color='Tipo', color_discrete_map={'XT-100': '#005195', 'XT-M': '#0a7c6e', 'SFE': '#E85500', 'Otro': '#888888'})
                _fig_tp.update_traces(textinfo='percent+label', sort=False)
                _fig_tp.update_layout(height=280, margin=dict(t=10, b=0, l=0, r=0), showlegend=False)
                st.plotly_chart(_fig_tp, use_container_width=True, config={'locale': 'es'})
            with _ct2:
                _tp_show = _por_tipo.copy()
                _tp_show['Km total'] = _tp_show['Km total'].map(lambda _v: f"{_ncl(_v, 0)} km")
                _tp_show['% uso'] = _tp_show['% uso'].map(lambda _v: f"{_ncl(_v, 1)} %")
                st.dataframe(_tp_show, use_container_width=True, hide_index=True)
                st.caption("% uso = participación de cada tipo en el km total de la flota.")
            st.markdown("#### Km recorridos por tren")
            _fig_kt = px.bar(_res_tr, x='Km recorridos', y='Tren', orientation='h', text='Km recorridos',
                             color='Tipo', color_discrete_map={'XT-100': '#005195', 'XT-M': '#0a7c6e', 'SFE': '#E85500', 'Otro': '#888888'})
            _fig_kt.update_layout(height=max(340, 20 * len(_res_tr)), margin=dict(t=10, b=0, l=0, r=0),
                                  yaxis=dict(autorange='reversed', title=''), title='')
            _fig_kt.update_traces(textposition='outside', cliponaxis=False)
            st.plotly_chart(_fig_kt, use_container_width=True, config={'locale': 'es'})
            st.caption("Km recorridos por cada tren en el período (kilometraje diario del odómetro UMR). Color por tipo: XT-100 (azul), XT-M (verde), SFE (naranja). Los trenes en 0 estuvieron detenidos o en mantención.")
            if not _flota_tr.empty:
                st.markdown("#### Km diario de la flota")
                _fig_fl = px.bar(_flota_tr, x='Fecha', y='Km flota', color_discrete_sequence=['#E85500'])
                _fig_fl.update_layout(height=300, margin=dict(t=10, b=0, l=0, r=0), yaxis=dict(title='km'))
                st.plotly_chart(_no_huecos(_fig_fl), use_container_width=True, config={'locale': 'es'})
                st.caption("Suma de km de todos los trenes por día.")
            st.markdown("#### Detalle por tren")
            st.dataframe(_res_tr, use_container_width=True)
            with st.expander("Ver kilometraje diario por tren (matriz tren × día) — mostrar / ocultar", expanded=False):
                _pk_show = _pivk_tr.copy()
                _pk_show.columns = [pd.to_datetime(_c).strftime('%d-%m') for _c in _pk_show.columns]
                st.dataframe(_pk_show, use_container_width=True)
    else: st.info("No hay datos de odómetro (UMR) cargados para el análisis por tren.")
    if all_kmserv:
        _dks = pd.DataFrame(all_kmserv)
        _dks['Fecha'] = pd.to_datetime(_dks['Fecha'], errors='coerce').dt.normalize()
        _dks['KmsxTrenes'] = pd.to_numeric(_dks['KmsxTrenes'], errors='coerce')
        _dks['KmTrenR'] = pd.to_numeric(_dks['KmTrenR'], errors='coerce')
        _dks = _dks.dropna(subset=['Fecha'])
        _agg_ks = _dks.groupby('Fecha', as_index=False).agg(**{'Kms.xTrenes': ('KmsxTrenes', 'sum'), 'KmTren R': ('KmTrenR', 'sum')})
        _agg_ks['Diferencia'] = (_agg_ks['Kms.xTrenes'] - _agg_ks['KmTren R']).round(2)
        _agg_ks = _agg_ks.sort_values('Fecha').reset_index(drop=True)
        st.divider()
        st.markdown("### 📏 Km por servicio: teórico vs real (KM-Servicio UMR)")
        _kt1 = float(_agg_ks['Kms.xTrenes'].sum()); _kt2 = float(_agg_ks['KmTren R'].sum()); _ktd = _kt1 - _kt2
        _mk1, _mk2, _mk3 = st.columns(3)
        _mk1.metric("Kms.xTrenes (total)", f"{_ncl(_kt1, 0)} km")
        _mk2.metric("KmTren R (total)", f"{_ncl(_kt2, 0)} km")
        _mk3.metric("Diferencia (no realizado)", f"{_ncl(_ktd, 0)} km", f"{_ncl(_ktd / _kt1 * 100 if _kt1 else 0, 2)} %")
        _aml = _agg_ks.melt(id_vars='Fecha', value_vars=['Kms.xTrenes', 'KmTren R'], var_name='Métrica', value_name='Km')
        _fig_ks = px.bar(_aml, x='Fecha', y='Km', color='Métrica', barmode='group',
                         color_discrete_map={'Kms.xTrenes': '#005195', 'KmTren R': '#E85500'})
        _fig_ks.update_layout(height=320, margin=dict(t=10, b=0, l=0, r=0))
        st.plotly_chart(_no_huecos(_fig_ks), use_container_width=True, config={'locale': 'es'})
        _fig_kd = px.bar(_agg_ks, x='Fecha', y='Diferencia', color_discrete_sequence=['#b91c1c'])
        _fig_kd.update_layout(height=260, margin=dict(t=10, b=0, l=0, r=0), yaxis=dict(title='km no realizados'))
        st.plotly_chart(_no_huecos(_fig_kd), use_container_width=True, config={'locale': 'es'})
        st.caption("Kms.xTrenes = km teórico (trenes × recorrido). KmTren R = km real recorrido. Diferencia = Kms.xTrenes − KmTren R (positivo: se recorrió menos de lo previsto).")
        with st.expander("Ver tabla diaria (Kms.xTrenes, KmTren R, diferencia) — mostrar / ocultar", expanded=False):
            _at_ks = _agg_ks.copy()
            _at_ks['Fecha'] = _at_ks['Fecha'].dt.strftime('%d-%m-%Y')
            st.dataframe(_at_ks, use_container_width=True, hide_index=True)

if _seccion == _SECCIONES[3]:
    if not df_ops.empty and 'E_Total' in df_ops.columns and df_ops['E_Total'].sum() > 0:
        st.header("🔋 Comparación de energías: Total · Tracción · 12 kV")
        _de = df_ops[df_ops['E_Total'] > 0].copy()
        _et = float(_de['E_Total'].sum()); _etr = float(_de['E_Tr'].sum()); _e12 = float(_de['E_12'].sum())
        _ptr = (_etr / _et * 100) if _et > 0 else 0.0
        _p12 = (_e12 / _et * 100) if _et > 0 else 0.0
        _k1, _k2, _k3 = st.columns(3)
        _k1.metric("Energía Total", f"{_ncl(_et, 0)} kWh")
        _k2.metric("Tracción", f"{_ncl(_etr, 0)} kWh", f"{_ncl(_ptr, 1)} % del total")
        _k3.metric("12 kV (auxiliares)", f"{_ncl(_e12, 0)} kWh", f"{_ncl(_p12, 1)} % del total")
        _cc1, _cc2 = st.columns([1, 2])
        with _cc1:
            _fcomp = px.pie(names=['Tracción', '12 kV'], values=[_etr, _e12], hole=0.55,
                            color=['Tracción', '12 kV'], color_discrete_map={'Tracción': '#E85500', '12 kV': '#005195'})
            _fcomp.update_traces(textinfo='percent+label', sort=False,
                                 hovertemplate='%{label}: %{value:,.0f} kWh (%{percent})<extra></extra>')
            _fcomp.update_layout(height=300, margin=dict(t=34, b=0, l=0, r=0), showlegend=False, title="Composición Tracción vs 12 kV")
            st.plotly_chart(_fcomp, use_container_width=True, config={'locale': 'es'})
        with _cc2:
            _mv = _de.melt(id_vars='Fecha', value_vars=['E_Total', 'E_Tr', 'E_12'], var_name='Energía', value_name='kWh')
            _mv['Energía'] = _mv['Energía'].map({'E_Total': 'Total', 'E_Tr': 'Tracción', 'E_12': '12 kV'})
            _fevo = px.line(_mv, x='Fecha', y='kWh', color='Energía', markers=True,
                            color_discrete_map={'Total': '#15803d', 'Tracción': '#E85500', '12 kV': '#005195'})
            _fevo.update_layout(height=300, margin=dict(t=34, b=0, l=0, r=0), yaxis_title='kWh', xaxis_title='', legend_title='', title="Evolución diaria de las 3 energías")
            st.plotly_chart(_no_huecos(_fevo), use_container_width=True, config={'locale': 'es'})
        _b1, _b2 = st.columns(2)
        with _b1:
            fig_e = go.Figure()
            fig_e.add_trace(go.Bar(x=_de['Fecha'], y=_de['E_Tr'], name='Tracción', marker_color='#E85500'))
            fig_e.add_trace(go.Bar(x=_de['Fecha'], y=_de['E_12'], name='12 kV', marker_color='#005195'))
            fig_e.update_layout(barmode='stack', height=330, margin=dict(t=40, b=0, l=0, r=0),
                                title="Tracción + 12 kV apiladas (kWh)", yaxis_title="kWh", xaxis_title='', legend_title='')
            st.plotly_chart(_no_huecos(fig_e), use_container_width=True)
        with _b2:
            _d2 = _de.copy()
            _d2['% Tracción'] = _d2['E_Tr'] / _d2['E_Total'] * 100
            _d2['% 12 kV'] = _d2['E_12'] / _d2['E_Total'] * 100
            _fpct = px.bar(_d2, x='Fecha', y=['% Tracción', '% 12 kV'], barmode='stack',
                           color_discrete_map={'% Tracción': '#E85500', '% 12 kV': '#005195'})
            _fpct.update_layout(height=330, margin=dict(t=40, b=0, l=0, r=0), title="Proporción diaria (%)",
                                yaxis_title="%", xaxis_title='', legend_title='')
            st.plotly_chart(_no_huecos(_fpct), use_container_width=True, config={'locale': 'es'})
        st.caption("Total = Tracción + 12 kV. Tracción = energía para mover los trenes; 12 kV = servicios auxiliares (iluminación, climatización, señalización, etc.).")
        with st.expander("Ver desglose diario — mostrar / ocultar", expanded=False):
            dv_ener = df_ops[['Fecha', 'E_Total', 'E_Tr', 'E_12', '% Tracción', '% 12 kV', 'Fuente']].copy()
            dv_ener['Fecha'] = dv_ener['Fecha'].dt.strftime('%d-%m-%Y')
            st.dataframe(make_columns_unique(dv_ener), use_container_width=True, hide_index=True)
    else: st.info("No hay datos de energía procesados (Facturación, PRMTE o SEAT).")

if _seccion == _SECCIONES[4]:
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
                st.plotly_chart(_no_huecos(fig_noche), use_container_width=True, config={'locale': 'es'})
            
            with c_noct2:
                st.markdown("<br>", unsafe_allow_html=True)
                st.metric("Consumo Nocturno Promedio", f"{_ncl(promedio_noche, 0)} kWh/noche")
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
            hovertemplate='Día: %{y}<br>Hora: %{x}<br>Consumo: %{_ncl(z, 1)} kWh<extra></extra>'
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

if _seccion == _SECCIONES[5]:
    st.markdown("### 🌙 Consumo Base Nocturno (Mediana)")
    if all_prmte_full:
        st.caption("Ventana nocturna por tipo de día — Laboral: 00:00–06:00 · Sábado: 00:00–07:00 · Domingo/Festivo: 00:00–08:00. Se usa la mediana porque es robusta frente a días atípicos.")
        _dfp = _prep_noche(all_prmte_full, df_ops)
        _dfp25 = _prep_noche(all_prmte_2025, df_ops)
        _cmap_td = {'Laboral': '#005195', 'Sábado': '#E85500', 'Domingo/Festivo': '#2CA02C'}
        _orden_td = [x for x in ['Laboral', 'Sábado', 'Domingo/Festivo'] if x in set(_dfp['Tipo'])]
        if _dfp.empty:
            st.info("No hay registros de PRMTE dentro de la ventana nocturna para el periodo cargado.")
        else:
            _tot_dia = _dfp.groupby(['Tipo', 'Fecha'])['Consumo'].sum().reset_index()
            _med_tot = _tot_dia.groupby('Tipo')['Consumo'].median()
            st.markdown("#### Mediana del consumo nocturno total por día")
            _cm = st.columns(max(1, len(_orden_td)))
            for _i, _t in enumerate(_orden_td):
                _cm[_i].metric(_t, f"{_ncl(_med_tot.get(_t, 0), 0)} kWh")
            st.caption("Mediana del total acumulado en la ventana nocturna de cada día (kWh/noche).")

            st.divider()
            st.markdown("#### 📐 Línea base nocturna (kWh/hora)")
            _hd = _dfp.groupby(['Tipo', 'Fecha', 'Hora'])['Consumo'].sum().reset_index()
            # --- Base anclada a la mediana de 2025 (año base ISO 50001) ---
            _hay_2025 = (_dfp25 is not None) and (not _dfp25.empty)
            _hd_base = (_dfp25 if _hay_2025 else _dfp).groupby(['Tipo', 'Fecha', 'Hora'])['Consumo'].sum().reset_index()
            _base_hora = _hd_base.groupby('Tipo')['Consumo'].median()
            _base_global = float(_hd_base['Consumo'].median())
            _suf = "2025" if _hay_2025 else "periodo cargado"
            _cb = st.columns(max(1, len(_orden_td)))
            for _i, _t in enumerate(_orden_td):
                _cb[_i].metric(f"Base {_t} ({_suf})", f"{_ncl(_base_hora.get(_t, 0), 0)} kWh/h")
            if _hay_2025:
                st.success(f"**Línea base nocturna general (2025): {_ncl(_base_global, 0)} kWh/hora** — mediana del consumo nocturno por hora durante 2025 (año base de referencia).")
            else:
                st.warning(f"No hay PRMTE de 2025 cargado: la base usa el periodo cargado ({_ncl(_base_global, 0)} kWh/hora). Carga el PRMTE de 2025 para anclar la base al año base.")
            st.caption("Valor de referencia (ISO 50001): la base es la mediana nocturna de 2025; si el consumo horario nocturno la supera de forma sostenida, suele indicar equipos operando sin necesidad (climatización, trenes sin apagar, auxiliares).")

            st.divider()
            st.markdown("#### Mediana cada 15 minutos")
            _p15 = _dfp.groupby(['Tipo', '15min'])['Consumo'].median().reset_index()
            fig15 = px.line(_p15, x='15min', y='Consumo', color='Tipo', markers=True,
                            color_discrete_map=_cmap_td, category_orders={'Tipo': _orden_td},
                            labels={'15min': 'Franja de 15 min', 'Consumo': 'Mediana (kWh/15min)', 'Tipo': ''})
            fig15.update_layout(margin=dict(t=10, b=0, l=0, r=0),
                                legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='left', x=0))
            st.plotly_chart(fig15, use_container_width=True, config={'locale': 'es'})
            st.caption("Para cada franja de 15 min, mediana del consumo de esa franja a lo largo de los días del mismo tipo.")

            st.divider()
            st.markdown("#### Mediana por hora")
            _ph = _hd.groupby(['Tipo', 'Hora'])['Consumo'].median().reset_index()
            figh = px.bar(_ph, x='Hora', y='Consumo', color='Tipo', barmode='group',
                          color_discrete_map=_cmap_td, category_orders={'Tipo': _orden_td},
                          labels={'Hora': 'Hora', 'Consumo': 'Mediana (kWh/hora)', 'Tipo': ''})
            figh.update_layout(margin=dict(t=10, b=0, l=0, r=0),
                               legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='left', x=0))
            figh.add_hline(y=_base_global, line_dash="dash", line_color="#475569", annotation_text=f"Base {_suf} · {_ncl(_base_global, 0)} kWh/h", annotation_position="top left")
            st.plotly_chart(figh, use_container_width=True, config={'locale': 'es'})
            st.caption("Por hora se suman los cuatro tramos de 15 min de cada día (kWh/hora) y luego se toma la mediana entre días.")

            with st.expander("Ver tablas de medianas"):
                st.markdown("**Cada 15 min (kWh/15min)**")
                st.dataframe(_p15.pivot(index='15min', columns='Tipo', values='Consumo').round(1), use_container_width=True)
                st.markdown("**Por hora (kWh/hora)**")
                st.dataframe(_ph.pivot(index='Hora', columns='Tipo', values='Consumo').round(1), use_container_width=True)
    else:
        st.info("Se necesita cargar el archivo de **PRMTE (Energía cada 15 min)** para este análisis de consumo nocturno.")

if _seccion == _SECCIONES[6]:
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
        st.plotly_chart(_no_huecos(fig_out), use_container_width=True)
        
        atipicos = df_ops[df_ops['Es_Atípico']][['Fecha', 'IDE (kWh/km)', 'Odómetro [km]', 'E_Tr']]
        if not atipicos.empty:
            st.warning("⚠️ Se han detectado los siguientes días con comportamiento anómalo en el consumo:")
            atipicos['Fecha'] = atipicos['Fecha'].dt.strftime('%Y-%m-%d')
            st.dataframe(atipicos, use_container_width=True)
        else: st.success("✅ No se detectaron valores atípicos significativos (Z-score > 2) en el periodo analizado.")
    else: st.info("No hay datos de IDE calculados para analizar.")

if _seccion == _SECCIONES[7]:
    st.markdown("### 📋 THDR — Servicios y Tiempos de Viaje")
    st.markdown("<style>.tv-card{border:1px solid #e2e8f0;border-radius:14px;padding:14px 16px;background:linear-gradient(135deg,#ffffff,#f7fafc);box-shadow:0 1px 4px rgba(15,23,42,.07)}.tv-head{font-weight:800;color:#005195;font-size:1.02rem;margin-bottom:.55rem;display:flex;align-items:center;gap:.45rem;flex-wrap:wrap}.tv-badge{background:#005195;color:#fff;font-size:.68rem;padding:2px 9px;border-radius:999px;font-weight:700}.tv-badge-alt{background:#0a7c6e}.tv-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px}.tv-stat{background:#fff;border:1px solid #eef2f7;border-radius:10px;padding:7px 10px;text-align:center}.tv-lbl{font-size:.68rem;color:#64748b;text-transform:uppercase;letter-spacing:.5px}.tv-val{font-size:1.22rem;font-weight:800;color:#0f172a;font-variant-numeric:tabular-nums}.tv-foot{margin-top:.5rem;font-size:.74rem;color:#64748b;text-align:right}</style>", unsafe_allow_html=True)
    df_thdr_v1, df_thdr_v2 = _thdr_filtros()
    if df_thdr_v1.empty and df_thdr_v2.empty:
        st.info("No se ha cargado/procesado THDR (o los filtros dejaron 0 registros).")
    else:
        _det = detalle_servicios(df_thdr_v1, df_thdr_v2)
        _tv = tiempos_servicios(df_thdr_v1, df_thdr_v2)
        _k1, _k2, _k3 = st.columns(3)
        _k1.metric("Servicios totales", f"{_ncl(len(_det), 0)}")
        _k2.metric("Dobles", f"{_ncl(int((_det['Composicion'] == 'Doble').sum()), 0)}" if not _det.empty else "0")
        _k3.metric("Simples", f"{_ncl(int((_det['Composicion'] == 'Simple').sum()), 0)}" if not _det.empty else "0")
        if _tv.empty:
            st.info("No se pudieron calcular tiempos de viaje (faltan horas de salida/llegada en la THDR).")
        else:
            st.markdown("#### ⏱️ Tiempos de viaje")
            st.caption("Tiempo de viaje = llegada al destino − salida del origen (malla THDR). Orden: Vía 1 (Puerto→...) y luego Vía 2 (...→Puerto). Formato HH:MM:SS. Cada bloque se puede mostrar u ocultar.")
            with st.expander("⏱️ Por tipo de servicio (mostrar / ocultar)", expanded=True):
                _ss = _ordenar_serv(_stats_dur(_tv, 'Tipo de servicio'), 'Tipo de servicio')
                _render_tv_cards(_ss, 'Tipo de servicio')
            with st.expander("🚆 Por composición — tren simple / doble (mostrar / ocultar)", expanded=False):
                _sc = _ordenar_serv_comp(_stats_dur(_tv, ['Tipo de servicio', 'Composicion']), 'Tipo de servicio', 'Composicion')
                if _sc.empty: st.info("Sin datos por composición.")
                else: _render_tv_cards(_sc, 'Tipo de servicio', badge_col='Composicion')
                st.caption("Cada servicio separado en tren Simple y Doble (cuando hay registros de ambos).")
            with st.expander("🚇 Por tipo de tren y tipo de servicio (mostrar / ocultar)", expanded=False):
                _stt = _ordenar_serv_tren(_stats_dur(_tv, ['Tipo de tren', 'Tipo de servicio']), 'Tipo de servicio', 'Tipo de tren')
                _render_tv_cards(_stt, 'Tipo de servicio', badge_col='Tipo de tren')
                st.caption("XT-100 / XT-M / SFE por cada servicio (vista general).")
            with st.expander("🚈 Por tipo de tren + composición — simple / doble (mostrar / ocultar)", expanded=False):
                _stc = _ordenar_serv_tren_comp(_stats_dur(_tv, ['Tipo de servicio', 'Tipo de tren', 'Composicion']), 'Tipo de servicio', 'Tipo de tren', 'Composicion')
                if _stc.empty: st.info("Sin datos por tipo de tren y composición.")
                else: _render_tv_cards(_stc, 'Tipo de servicio', badge_col=['Tipo de tren', 'Composicion'])
                st.caption("Cada servicio por material rodante (XT-100/XT-M/SFE) y además simple/doble. La vista general (sin separar simple/doble) está en el bloque anterior.")
            with st.expander("🚉 Detención en estaciones y tiempo entre estaciones (mostrar / ocultar)", expanded=False):
                _dw = _dwell_estaciones(df_thdr_v1, df_thdr_v2)
                st.markdown("**🛑 Tiempo detenido en cada estación**")
                if _dw.empty: st.info("Sin datos de detención por estación.")
                else: _render_tv_cards(_stats_dur(_dw, 'Estacion'), 'Estacion')
                _sg = _segmentos(df_thdr_v1, df_thdr_v2)
                st.markdown("**↔️ Tiempo de viaje entre estaciones consecutivas**")
                if _sg.empty: st.info("Sin datos de tiempo entre estaciones.")
                else: _render_tv_cards(_stats_dur(_sg, 'Segmento'), 'Segmento')
        st.markdown("#### 🚦 Cruzamientos entre vía 1 y vía 2")
        with st.expander("🚦 Diagrama de cruzamientos (tiempo × estación) — mostrar / ocultar", expanded=False):
            _v1raw = st.session_state.get('df_thdr_v1', pd.DataFrame())
            _v2raw = st.session_state.get('df_thdr_v2', pd.DataFrame())
            _fm = _fechas_thdr(_v1raw, _v2raw)
            if not _fm:
                st.info("Sin fechas disponibles para el diagrama de cruzamientos.")
            else:
                st.caption("Este diagrama usa su propio selector de día (independiente de los filtros de la pestaña). El cruzamiento solo es real dentro de un mismo día.")
                _fsel = st.selectbox("📅 Día a graficar", _fm, format_func=lambda _d: _d.strftime('%d-%m-%Y'), key="marey_fecha")
                _figm, _cruces = _diagrama_marey(_filtra_fecha_op(_v1raw, _fsel), _filtra_fecha_op(_v2raw, _fsel))
                if _figm is None:
                    st.info("Sin datos suficientes para el diagrama de cruzamientos en ese día.")
                else:
                    st.plotly_chart(_figm, use_container_width=True, config={'locale': 'es'})
                    st.caption("Cada línea es un servicio del día elegido: las azules (vía 1) suben de Puerto a Limache y las naranjas (vía 2) bajan de Limache a Puerto. Las ✕ verdes son los cruzamientos: pasa el cursor sobre una para ver el km, el número de servicio (Viaje), el tren y el recorrido de ambos trenes.")
                    if _cruces:
                        _dfc = pd.DataFrame(_cruces)
                        _por_tramo = _dfc.groupby('tramo').size().sort_values(ascending=False)
                        _por_hora = _dfc.groupby('hora_h').size().sort_values(ascending=False)
                        st.markdown("##### 📊 Resumen de cruzamientos del día")
                        _rm1, _rm2, _rm3 = st.columns(3)
                        _rm1.metric("Total de cruzamientos", _ncl(len(_dfc), 0))
                        _rm2.metric("Tramo con más cruces", _por_tramo.index[0], f"{_ncl(int(_por_tramo.iloc[0]), 0)} cruces")
                        _rm3.metric("Hora con más cruces", f"{int(_por_hora.index[0]):02d}:00 h", f"{_ncl(int(_por_hora.iloc[0]), 0)} cruces")
                        _bt = _por_tramo.reset_index(); _bt.columns = ['Tramo', 'Cruces']
                        _figb = px.bar(_bt, x='Cruces', y='Tramo', orientation='h', text='Cruces', color_discrete_sequence=['#15803d'])
                        _figb.update_layout(height=max(280, 24 * len(_bt)), margin=dict(t=10, b=0, l=0, r=0), yaxis=dict(autorange='reversed', title=''))
                        _figb.update_traces(textposition='outside', cliponaxis=False)
                        st.plotly_chart(_figb, use_container_width=True, config={'locale': 'es'})
                        st.caption("Tramos donde más se cruzan los trenes: mientras más larga la barra, mayor concentración de cruzamientos en ese tramo de la línea.")
        if not _det.empty:
            st.divider()
            st.markdown("#### 📥 Servicios por tipo de tren (descargable)")
            st.download_button("⬇️ Descargar Excel (Por día + Resumen + Detalle)", data=excel_servicios(_det), file_name="servicios_por_tipo_tren.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            _ca, _cb = st.columns(2)
            with _ca:
                st.markdown("**Por día**")
                st.dataframe(_det.groupby(['Fecha', 'Tipo de servicio', 'Tipo de tren', 'Composicion']).size().reset_index(name='Servicios'), use_container_width=True, hide_index=True)
            with _cb:
                st.markdown("**Resumen (totales)**")
                st.dataframe(_det.groupby(['Tipo de servicio', 'Tipo de tren', 'Composicion']).size().reset_index(name='Servicios'), use_container_width=True, hide_index=True)
    st.divider()
    st.markdown("#### Datos THDR crudos")
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

if _seccion == _SECCIONES[8]:
    st.markdown("### 🔬 Análisis Multivariante: PAX vs Tiempos vs Energía")
    st.markdown("Este módulo utiliza estadística robusta para cruzar la demanda real con la fricción operativa (Tiempos de Viaje/Detenciones) y explicar el gasto de Tracción.")
    
    if not df_thdr_v1.empty and not df_thdr_v2.empty and not df_ops.empty and not df_carga_v1.empty:
        
        # El filtro de jornada/año/fecha viene de la barra global (arriba).
        
        # Filtrar matemáticamente todas las bases de datos subyacentes
        fechas_validas = df_ops['Fecha']
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
            df_plot = df_mixto.copy()
            # Blindaje: numéricos reales y sin NaN/inf (evita el ValueError de Plotly en 'size')
            for _c in ['Tiempo_Mediana_Red', 'E_Tr', 'PAX', 'IDE (kWh/km)']:
                if _c in df_plot.columns:
                    df_plot[_c] = pd.to_numeric(df_plot[_c], errors='coerce')
            df_plot = df_plot.replace([np.inf, -np.inf], np.nan)
            df_plot = df_plot.dropna(subset=['Tiempo_Mediana_Red', 'E_Tr', 'PAX'])
            df_plot = df_plot[df_plot['PAX'] > 0]
            
            if not df_plot.empty and df_plot['E_Tr'].sum() > 0:
                
                # --- 2. BUBBLE CHART 4D ---
                st.markdown("#### 🫧 Ecosistema Operativo Diario (Macro)")
                st.caption("Eje X: Mediana de Tiempos de Viaje | Eje Y: Consumo Tracción | Tamaño: Volumen de Pasajeros")
                
                df_plot['Tiempo Promedio HH:MM:SS'] = df_plot['Tiempo_Mediana_Red'].apply(minutos_a_hhmmss)
                
                # 🛡️ CORRECCIÓN DE ESTABILIDAD: Eliminación de diccionario en hover_data por lista simple.
                try:
                    fig_mix = px.scatter(
                        df_plot,
                        x='Tiempo_Mediana_Red',
                        y='E_Tr',
                        size='PAX', size_max=40,
                        color='Tipo Día',
                        hover_name='Fecha (ES)',
                        hover_data=['Tiempo Promedio HH:MM:SS', 'IDE (kWh/km)'],
                        labels={'Tiempo_Mediana_Red': 'Tiempo Mediano de Viaje', 'E_Tr': 'Energía de Tracción (kWh)'},
                        color_discrete_map={'L': '#005195', 'S': '#E85500', 'D/F': '#2CA02C'})
                except Exception:
                    # Fallback sin 'size' si una versión estricta de Plotly rechaza tamaños no finitos
                    fig_mix = px.scatter(
                        df_plot,
                        x='Tiempo_Mediana_Red',
                        y='E_Tr',
                        color='Tipo Día',
                        hover_name='Fecha (ES)',
                        hover_data=['Tiempo Promedio HH:MM:SS', 'IDE (kWh/km)', 'PAX'],
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
                    df_c['_srv_clean'] = _srv_clean_series(df_c[c_serv_c])
                    t_sub = df_thdr_filt.copy()
                    t_sub['_srv_clean'] = _srv_clean_series(t_sub[c_serv_t])

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
                        est_code = _norm(PAX_COL_CODE.get(est, ''))
                        # 1a. Match exacto por código de 3 letras del export (PUE, BEL, VIN, SLT, VAM, SGA...)
                        if est_code:
                            for c in df_c.columns:
                                if _norm(c).strip() == est_code:
                                    c_est = c; break
                        # 1b. Respaldo: nombre o alias contenido en el encabezado
                        if not c_est:
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
                        c_orig = None
                        for _e in valid_stations:
                            c_orig = get_col_thdr(t_sub, _e, 'SALIDA')
                            if c_orig: break
                        if not c_orig and valid_stations:
                            c_orig = get_col_thdr(t_sub, valid_stations[0], 'LLEGADA')
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
                            fig_heat_pax.update_traces(hovertemplate="Hora Local: %{x}:00<br>Sector: %{y}<br>Carga PAX: %{_ncl(z, 0)}<extra></extra>")
                            fig_heat_pax.update_layout(margin=dict(t=20, b=20, l=0, r=0), height=500)
                            st.plotly_chart(fig_heat_pax, use_container_width=True)
                            
                            pax_max = df_pax_heat.loc[df_pax_heat['Pax'].idxmax()]
                            st.caption(f"**Insight:** Mayor carga prom. en **{pax_max['Estacion']}** a las **{pax_max['Hora_Estacion']:02d}:00 hrs** ({_ncl(pax_max['Pax'], 0)} PAX).")
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
                            fig_heat_pax_v2.update_traces(hovertemplate="Hora Local: %{x}:00<br>Sector: %{y}<br>Carga PAX: %{_ncl(z, 0)}<extra></extra>")
                            fig_heat_pax_v2.update_layout(margin=dict(t=20, b=20, l=0, r=0), height=500)
                            st.plotly_chart(fig_heat_pax_v2, use_container_width=True)
                            pax_max_v2 = df_pax_heat_v2.loc[df_pax_heat_v2['Pax'].idxmax()]
                            st.caption(f"**Insight:** Mayor carga prom. en **{pax_max_v2['Estacion']}** a las **{pax_max_v2['Hora_Estacion']:02d}:00 hrs** ({_ncl(pax_max_v2['Pax'], 0)} PAX).")
                        else:
                            st.info("Formato de pasajeros no detectado (ni por estación ni consolidado).")
                else:
                    st.info("Se requiere procesar archivo THDR Vía 2 para este filtro.")
            else:
                st.warning("No hay suficientes datos superpuestos para realizar la regresión en base a los filtros actuales.")
    else: 
        st.info("⚠️ Carga archivos de **THDR (Vía 1 y 2), Facturación/PRMTE/SEAT y Carga de Pasajeros** para habilitar el Microscopio Operacional.")

if _seccion == _SECCIONES[9]:
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
            st.plotly_chart(_no_huecos(fig_pas), use_container_width=True)

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

if _seccion == _SECCIONES[10]:
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
                            peak_hr_msg = f"La 'Hora Punta Eléctrica' ocurre a las **{hora_peak}** ({_ncl(consumo_peak, 0)} kWh prom.)."
                            
                            df_hr_filt['Hora_n'] = df_hr_filt['Hora'].astype(str).str.slice(0, 2).astype(int)
                            limite_hora = 6 if tipo == "L" else (7 if tipo == "S" else 8)
                            
                            df_noche = df_hr_filt[(df_hr_filt['Hora_n'] >= 0) & (df_hr_filt['Hora_n'] < limite_hora)]
                            if not df_noche.empty:
                                noche_diario = df_noche.groupby('Fecha')['Consumo'].sum().reset_index()
                                promedio_noche = noche_diario['Consumo'].mean()
                                max_noche = noche_diario.loc[noche_diario['Consumo'].idxmax()]
                                if max_noche['Consumo'] > (promedio_noche * 1.2) and promedio_noche > 0:
                                    noche_msg = f"🌙 **Alerta Parásita:** Pico de **{_ncl(max_noche['Consumo'], 0)} kWh** la madrugada del {max_noche['Fecha'].strftime('%d/%m')} (Ventana: 00:00 a 0{limite_hora}:00 hrs)."
                                else:
                                    noche_msg = f"🌙 **Auditoría Nocturna:** Estable ({_ncl(promedio_noche, 0)} kWh de 00:00 a 0{limite_hora}:00 hrs)."

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
                                thdr_msg = f"**{_ncl(total_viajes, 0)} servicios** en total. Uso de Tracción Doble: **{(v_doble/total_viajes*100) if total_viajes>0 else 0:.1f}%**."

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
                            st.info(f"👥 **Peak de Demanda:** {dia_max_pax['Fecha (ES)']} con **{_ncl(int(dia_max_pax['PAX']), 0)}** personas.")
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

if _seccion == _SECCIONES[11]:
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
                        a1.metric("E. Total", f"{_ncl(r['E_Total'], 0)} kWh")
                        a2.metric("Tracción", f"{_ncl(r['E_Tr'], 0)} kWh")
                        a3.metric("12 kV", f"{_ncl(r['E_12'], 0)} kWh")
                        a4.metric("IDE", f"{_ncl(r['IDE (kWh/km)'], 2)} kWh/km")
                        
                        b1, b2, b3, b4 = st.columns(4)
                        noche_ok = ("Noche_kWh" in df_diag.columns) and pd.notna(r["Noche_kWh"])
                        b1.metric("Nocturno", f"{_ncl(r['Noche_kWh'], 0)} kWh" if noche_ok else "—")
                        b2.metric("Servicios", f"{_ncl(int(r['Servicios']), 0)}")
                        b3.metric("PAX", f"{_ncl(int(r['PAX']), 0)}" if pd.notna(r["PAX"]) else "—")
                        odo_ok = ("Odómetro [km]" in df_diag.columns) and pd.notna(r["Odómetro [km]"])
                        b4.metric("Odómetro", f"{_ncl(r['Odómetro [km]'], 0)} km" if odo_ok else "—")

                        st.markdown("##### 🚆 Análisis Logístico y Contexto")
                        cols_ctx = st.columns(2)
                        bits = []
                        if pd.notna(r.get("Doble_pct")): bits.append(f"**Uso Tracción Doble:** {r['Doble_pct']:.0f}%")
                        if r.get("Est_critica"): bits.append(f"**Punto de Saturación:** Estación {r['Est_critica']}")
                        if pd.notna(r.get("Ocup_max")): bits.append(f"**Peak Carga de Tren:** {_ncl(int(r['Ocup_max']), 0)} PAX")
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


if _seccion == _SECCIONES[12]:
    _det_sv = detalle_servicios(df_thdr_v1, df_thdr_v2, None)
    if _det_sv is None or _det_sv.empty:
        st.info("No hay datos de servicios (THDR) en el rango seleccionado. Sube archivos THDR desde el panel lateral.")
    else:
        st.markdown("### 📈 Distribución de servicios")
        _tot_sv = len(_det_sv)
        _n_simp = int((_det_sv['Composicion'] == 'Simple').sum())
        _n_dob = int((_det_sv['Composicion'] == 'Doble').sum())
        _ma, _mb, _mc = st.columns(3)
        _ma.metric("Total de servicios", _ncl(_tot_sv, 0))
        _mb.metric("Simples", _ncl(_n_simp, 0), f"{_ncl(_n_simp / _tot_sv * 100 if _tot_sv else 0, 1)} %")
        _mc.metric("Dobles", _ncl(_n_dob, 0), f"{_ncl(_n_dob / _tot_sv * 100 if _tot_sv else 0, 1)} %")
        _viz_opts = ["🍩 Torta", "📊 Barra", "🪪 Tarjetas"]
        try:
            if hasattr(st, "segmented_control"):
                _viz = st.segmented_control("Ver como", _viz_opts, default=_viz_opts[0], key="_sv_viz")
            elif hasattr(st, "pills"):
                _viz = st.pills("Ver como", _viz_opts, default=_viz_opts[0], key="_sv_viz")
            else:
                _viz = st.radio("Ver como", _viz_opts, horizontal=True, key="_sv_viz")
        except Exception:
            _viz = st.radio("Ver como", _viz_opts, horizontal=True, key="_sv_viz2")
        if not _viz:
            _viz = _viz_opts[0]
        def _render_dist(_col, _titulo, _cmap=None):
            _vc = _det_sv[_col].value_counts().reset_index()
            _vc.columns = [_col, 'Cantidad']
            _vc['pct'] = (_vc['Cantidad'] / _tot_sv * 100).round(1)
            st.markdown(f"#### {_titulo}")
            if "Torta" in _viz:
                if _cmap:
                    _f = px.pie(_vc, names=_col, values='Cantidad', hole=0.5, color=_col, color_discrete_map=_cmap)
                else:
                    _f = px.pie(_vc, names=_col, values='Cantidad', hole=0.5)
                _f.update_traces(textinfo='percent+label', sort=False,
                                 hovertemplate='%{label}: %{value} servicios (%{percent})<extra></extra>')
                _f.update_layout(height=330, margin=dict(t=10, b=0, l=0, r=0), showlegend=False)
                st.plotly_chart(_f, use_container_width=True, config={'locale': 'es'})
            elif "Barra" in _viz:
                if _cmap:
                    _f = px.bar(_vc, x='Cantidad', y=_col, orientation='h', text='Cantidad', color=_col, color_discrete_map=_cmap)
                else:
                    _f = px.bar(_vc, x='Cantidad', y=_col, orientation='h', text='Cantidad', color_discrete_sequence=['#005195'])
                _f.update_traces(textposition='outside', cliponaxis=False, customdata=_vc[['pct']],
                                 hovertemplate='%{y}: %{x} servicios (%{customdata[0]}%)<extra></extra>')
                _f.update_layout(height=max(220, 48 * len(_vc)), margin=dict(t=10, b=0, l=0, r=0),
                                 yaxis=dict(autorange='reversed', title=''), xaxis_title='Servicios', showlegend=False)
                st.plotly_chart(_f, use_container_width=True, config={'locale': 'es'})
            else:
                _pal = {'XT-100': '#005195', 'XT-M': '#0a7c6e', 'SFE': '#E85500', 'Simple': '#005195', 'Doble': '#E85500'}
                _cl = st.columns(min(len(_vc), 4))
                for _i, (_, _r) in enumerate(_vc.iterrows()):
                    _clr = (_cmap or {}).get(str(_r[_col])) or _pal.get(str(_r[_col]), '#005195')
                    _cl[_i % len(_cl)].markdown(
                        f'<div style="border:1px solid #e5e7eb;border-radius:12px;padding:12px 14px;background:#fff;margin-bottom:8px">'
                        f'<div style="font-size:.78rem;color:#6b7280;font-weight:600">{_r[_col]}</div>'
                        f'<div style="font-size:1.6rem;font-weight:800;color:{_clr}">{_ncl(_r["Cantidad"], 0)}</div>'
                        f'<div style="font-size:.85rem;color:#374151">{_ncl(_r["pct"], 1)} % del total</div></div>',
                        unsafe_allow_html=True)
        _CM_TREN = {'XT-100': '#005195', 'XT-M': '#0a7c6e', 'SFE': '#E85500', 'Sin asignar': '#9aa0a6'}
        _render_dist('Composicion', 'Cantidad de trenes (simple / doble)', {'Simple': '#005195', 'Doble': '#E85500'})
        _render_dist('Tipo de tren', 'Tipos de tren (XT-100 / XT-M / SFE)', _CM_TREN)
        _render_dist('Tipo de servicio', 'Tipo de servicio (recorrido)')
        st.caption("Distribución de los servicios THDR del período seleccionado. Los porcentajes son sobre el total de servicios.")
        st.markdown("#### Composición según tipo de servicio y tipo de tren")
        def _render_cruce(_dim, _titulo, _cmap=None):
            _g = _det_sv.groupby(['Composicion', _dim]).size().reset_index(name='Cantidad')
            _tc = _g.groupby('Composicion')['Cantidad'].transform('sum')
            _g['pct'] = (_g['Cantidad'] / _tc * 100).round(1)
            _g['pct_str'] = _g['pct'].map(lambda _v: _ncl(_v, 1))
            st.markdown(f"**{_titulo}**")
            if "Barra" in _viz:
                _f = px.bar(_g, x='Composicion', y='Cantidad', color=_dim, barmode='stack',
                            custom_data=['pct_str'], color_discrete_map=_cmap, text='Cantidad')
                _f.update_traces(textposition='inside',
                                 hovertemplate='<b>%{fullData.name}</b><br>%{x}: %{y} servicios (%{customdata[0]}%)<extra></extra>')
                _f.update_layout(height=390, margin=dict(t=10, b=0, l=0, r=0), yaxis_title='Servicios', xaxis_title='', legend_title='')
                st.plotly_chart(_f, use_container_width=True, config={'locale': 'es'})
            elif "Torta" in _viz:
                _f = px.sunburst(_g, path=['Composicion', _dim], values='Cantidad')
                _f.update_traces(hovertemplate='<b>%{label}</b><br>%{value} servicios<br>%{percentParent} de %{parent}<extra></extra>')
                _f.update_layout(height=390, margin=dict(t=10, b=0, l=0, r=0))
                st.plotly_chart(_f, use_container_width=True, config={'locale': 'es'})
            else:
                _pal = {'XT-100': '#005195', 'XT-M': '#0a7c6e', 'SFE': '#E85500', 'Simple': '#005195', 'Doble': '#E85500'}
                for _comp in ['Simple', 'Doble']:
                    _sub = _g[_g['Composicion'] == _comp].sort_values('Cantidad', ascending=False)
                    if _sub.empty:
                        continue
                    st.markdown(f"*{_comp}* ({_ncl(int(_sub['Cantidad'].sum()), 0)} servicios)")
                    _cl = st.columns(min(len(_sub), 4))
                    for _i, (_, _r) in enumerate(_sub.iterrows()):
                        _clr = (_cmap or {}).get(str(_r[_dim])) or _pal.get(str(_r[_dim]), '#005195')
                        _cl[_i % len(_cl)].markdown(
                            f'<div style="border:1px solid #e5e7eb;border-radius:12px;padding:10px 12px;background:#fff;margin-bottom:8px">'
                            f'<div style="font-size:.74rem;color:#6b7280;font-weight:600">{_r[_dim]}</div>'
                            f'<div style="font-size:1.4rem;font-weight:800;color:{_clr}">{_ncl(_r["Cantidad"], 0)}</div>'
                            f'<div style="font-size:.8rem;color:#374151">{_r["pct_str"]} % de los {_comp.lower()}s</div></div>',
                            unsafe_allow_html=True)
        _cM = {'XT-100': '#005195', 'XT-M': '#0a7c6e', 'SFE': '#E85500', 'Sin asignar': '#9aa0a6'}
        if "Tarjetas" in _viz:
            _render_cruce('Tipo de servicio', 'Simple / Doble por tipo de servicio')
            _render_cruce('Tipo de tren', 'Simple / Doble por tipo de tren', _cM)
        else:
            _cc1, _cc2 = st.columns(2)
            with _cc1:
                _render_cruce('Tipo de servicio', 'Simple / Doble por tipo de servicio')
            with _cc2:
                _render_cruce('Tipo de tren', 'Simple / Doble por tipo de tren', _cM)
        st.caption("Cada composición (Simple / Doble) se desglosa por tipo de servicio y por tipo de tren, con la cantidad y el % dentro de esa composición.")
        def _excel_dist(_det):
            _buf = BytesIO(); _t = len(_det)
            with pd.ExcelWriter(_buf, engine='openpyxl') as _w:
                for _c, _nm in [('Composicion', 'Composicion'), ('Tipo de tren', 'Tipo de tren'), ('Tipo de servicio', 'Tipo de servicio')]:
                    _d = _det[_c].value_counts().reset_index(); _d.columns = [_nm, 'Cantidad']
                    _d['%'] = (_d['Cantidad'] / _t * 100).round(1)
                    _d.to_excel(_w, sheet_name=_nm[:31], index=False)
                pd.crosstab(_det['Tipo de servicio'], _det['Composicion'], margins=True, margins_name='Total').to_excel(_w, sheet_name='Comp x Servicio')
                pd.crosstab(_det['Tipo de tren'], _det['Composicion'], margins=True, margins_name='Total').to_excel(_w, sheet_name='Comp x Tren')
                _det.to_excel(_w, sheet_name='Detalle', index=False)
            return _buf.getvalue()
        st.download_button("⬇️ Descargar tablas en Excel", _excel_dist(_det_sv),
                           file_name="servicios_distribucion.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key="_dl_serv_xlsx")
        with st.expander("Ver tablas con cantidades y % — mostrar / ocultar", expanded=False):
            for _cc, _tt in [('Composicion', 'Composición'), ('Tipo de tren', 'Tipo de tren'), ('Tipo de servicio', 'Tipo de servicio')]:
                _vc = _det_sv[_cc].value_counts().reset_index()
                _vc.columns = [_tt, 'Cantidad']
                _vc['%'] = (_vc['Cantidad'] / _tot_sv * 100).round(1).map(lambda _v: f"{_ncl(_v, 1)} %")
                st.markdown(f"**{_tt}**")
                st.dataframe(_vc, use_container_width=True, hide_index=True)
            st.markdown("**Composición × Tipo de servicio**")
            st.dataframe(pd.crosstab(_det_sv['Tipo de servicio'], _det_sv['Composicion'], margins=True, margins_name='Total'), use_container_width=True)
            st.markdown("**Composición × Tipo de tren**")
            st.dataframe(pd.crosstab(_det_sv['Tipo de tren'], _det_sv['Composicion'], margins=True, margins_name='Total'), use_container_width=True)

# --- Pestaña: Ahorro de energía (UMR vs meta) ---
if _seccion == _SECCIONES[13]:
    st.header("💡 Ahorro de energía (UMR vs meta)")
    if df_ops is None or df_ops.empty or float(df_ops['Odómetro [km]'].sum()) <= 0:
        st.info("No hay datos de odómetro/energía cargados en el período para calcular el ahorro.")
    else:
        _meta = st.number_input("Meta UMR", min_value=0.500, max_value=1.000, value=0.964, step=0.001, format="%.3f",
                                help="UMR = Tren-Km ÷ Odómetro. La meta por defecto es 0,964 (96,4%).")
        _b = df_ops.copy()
        _b['_f'] = pd.to_datetime(_b['Fecha']).dt.date
        _b['Meta km'] = _b['Odómetro [km]'] * _meta
        # 1) UMR — usa el Tren-Km del dashboard (df_ops, archivo de operación)
        _b['Δ UMR'] = _b['Tren-Km [km]'] - _b['Meta km']
        _b['Ahorro UMR'] = _b['Δ UMR'] * _b['IDE (kWh/km)']
        # Tren-Km R (real) y Kms.xTrenes (teórico) — de la hoja KM-Servicio del Excel
        if all_kmserv:
            _dkp = pd.DataFrame(all_kmserv)
            for _c in ['KmTrenR', 'KmsxTrenes']:
                _dkp[_c] = pd.to_numeric(_dkp.get(_c), errors='coerce')
            _aggp = _dkp.groupby(pd.to_datetime(_dkp['Fecha']).dt.date, as_index=False).agg(_r=('KmTrenR', 'sum'), _t=('KmsxTrenes', 'sum'))
            _aggp.columns = ['_f', 'Tren-Km R', 'Kms.xTrenes']
            _b = _b.merge(_aggp, on='_f', how='left')
        else:
            _b['Tren-Km R'] = np.nan; _b['Kms.xTrenes'] = np.nan
        # 2) Tren-Km R (real)
        _b['Δ R'] = _b['Tren-Km R'] - _b['Meta km']
        _b['Ahorro Tren-Km R'] = _b['Δ R'] * _b['IDE (kWh/km)']
        # 3) Kms.xTrenes (teórico)
        _b['Δ T'] = _b['Kms.xTrenes'] - _b['Meta km']
        _b['Ahorro Kms.xTrenes'] = _b['Δ T'] * _b['IDE (kWh/km)']
        _has = bool(_b['Tren-Km R'].notna().any())
        _m1, _m2, _m3 = st.columns(3)
        _uR = float(_b['Ahorro UMR'].sum()); _ukm = float(_b['Δ UMR'].sum())
        _m1.metric("Ahorro UMR (dashboard)", f"{_ncl(_uR, 0)} kWh", f"{_ncl(_ukm, 0)} km vs meta")
        if _has:
            _rR = float(_b['Ahorro Tren-Km R'].sum()); _rkm = float(_b['Δ R'].sum())
            _m2.metric("Ahorro Tren-Km R (real)", f"{_ncl(_rR, 0)} kWh", f"{_ncl(_rkm, 0)} km vs meta")
            _tR = float(_b['Ahorro Kms.xTrenes'].sum()); _tkm = float(_b['Δ T'].sum())
            _m3.metric("Ahorro Kms.xTrenes (teórico)", f"{_ncl(_tR, 0)} kWh", f"{_ncl(_tkm, 0)} km vs meta")
        else:
            _m2.metric("Ahorro Tren-Km R (real)", "—", help="Carga el Excel UMR (hoja KM-Servicio).")
            _m3.metric("Ahorro Kms.xTrenes (teórico)", "—", help="Carga el Excel UMR (hoja KM-Servicio).")
        _vv = ['Ahorro UMR'] + (['Ahorro Tren-Km R', 'Ahorro Kms.xTrenes'] if _has else [])
        _mv = _b.melt(id_vars='Fecha', value_vars=_vv, var_name='Ahorro', value_name='kWh').dropna(subset=['kWh'])
        _cmap_ah = {'Ahorro UMR': '#005195', 'Ahorro Tren-Km R': '#E85500', 'Ahorro Kms.xTrenes': '#0a7c6e'}
        _fig_ah = px.bar(_mv, x='Fecha', y='kWh', color='Ahorro', barmode='group',
                         color_discrete_map=_cmap_ah, title="Ahorro de energía por día (kWh) — UMR · Tren-Km R · Kms.xTrenes")
        _fig_ah.update_layout(margin=dict(t=46, b=0, l=0, r=0), yaxis_title='kWh', xaxis_title='', legend_title='')
        st.plotly_chart(_no_huecos(_fig_ah), use_container_width=True, config={'locale': 'es'})
        st.caption(f"Cada ahorro = (Tren-Km − Odómetro × {_ncl(_meta, 3)}) × IDE del día. Positivo = ahorro (sobre la meta); negativo = sobreconsumo. IDE = E_Tr ÷ Odómetro de cada día. El UMR usa el Tren-Km del dashboard; Tren-Km R (real) y Kms.xTrenes (teórico) vienen de la hoja KM-Servicio del Excel.")
        with st.expander("Ver tabla diaria — mostrar / ocultar", expanded=False):
            _cols_t = ['_f', 'Odómetro [km]', 'Meta km', 'IDE (kWh/km)', 'Tren-Km [km]', 'Ahorro UMR', 'Tren-Km R', 'Ahorro Tren-Km R', 'Kms.xTrenes', 'Ahorro Kms.xTrenes']
            _tab = _b[_cols_t].copy()
            _tab.columns = ['Fecha', 'Odómetro', 'Meta km', 'IDE', 'Tren-Km (UMR)', 'Ahorro UMR', 'Tren-Km R', 'Ahorro Tren-Km R', 'Kms.xTrenes', 'Ahorro Kms.xTrenes']
            for _c in ['Odómetro', 'Meta km', 'Tren-Km (UMR)', 'Ahorro UMR', 'Tren-Km R', 'Ahorro Tren-Km R', 'Kms.xTrenes', 'Ahorro Kms.xTrenes']:
                _tab[_c] = _tab[_c].map(lambda _v: _ncl(_v, 2) if pd.notna(_v) else '—')
            _tab['IDE'] = _tab['IDE'].map(lambda _v: _ncl(_v, 3) if pd.notna(_v) else '—')
            st.dataframe(_tab, use_container_width=True, hide_index=True)
