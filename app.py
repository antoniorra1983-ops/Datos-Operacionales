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
/* Evitar que los títulos de las tarjetas (Métricas) se corten: Responsive wrapping */
div[data-testid="stMetricLabel"] > label {
    white-space: normal !important; 
    word-wrap: break-word !important; 
    min-height: 2.5rem;
    font-size: 0.95rem;
}
/* Garantizar que los valores numéricos grandes siempre se vean completos */
div[data-testid="stMetricValue"] {
    font-size: 1.6rem !important;
    word-wrap: break-word !important;
    white-space: normal !important;
}
/* Forzar scroll horizontal SOLO en tablas gigantes, no en el layout principal */
.stDataFrame { overflow-x: auto; }
</style>""", unsafe_allow_html=True)

# --- 2. CONSTANTES DE RED Y CONFIGURACIONES ---
ESTACIONES = [
    'Puerto','Bellavista','Francia','Baron','Portales','Recreo','Miramar',
    'Viña del Mar','Hospital','Chorrillos','El Salto','Valencia','Quilpue',
    'El Sol','El Belloto','Las Americas','La Concepcion','Villa Alemana',
    'Sargento Aldea','Peñablanca','Limache'
]
ESTACIONES_CORTO = [e[:2].upper() for e in ESTACIONES]
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

# --- 3b. MOTOR DE DIAGNÓSTICO DE ANOMALÍAS (segmentado por tipo de día) ---
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

def _perfil_horario_diario(all_prmte_full, all_fact_full):
    datos, freq = (all_prmte_full, 15) if all_prmte_full else (all_fact_full, 60)
    if not datos:
        return pd.DataFrame(columns=["Fecha", "Noche_kWh", "Pico_kW"])
    h = pd.DataFrame(datos)
    h["Fecha"] = pd.to_datetime(h["Fecha"]).dt.normalize()
    h["Hora_n"] = h["Hora"].astype(str).str.slice(0, 2).apply(lambda x: int(x) if str(x).isdigit() else -1)
    noche = h[h["Hora_n"].isin([1, 2, 3, 4])].groupby("Fecha")["Consumo"].sum().rename("Noche_kWh")
    pico = (h.groupby("Fecha")["Consumo"].max() * (60.0 / freq)).rename("Pico_kW")
    return pd.concat([noche, pico], axis=1).reset_index()

def diagnosticar_anomalias(df_ops, all_prmte_full=None, all_fact_full=None, z_alerta=2.5, z_fuerte=3.5):
    if df_ops is None or df_ops.empty:
        return pd.DataFrame()
    d = df_ops[df_ops["E_Total"] > 0].copy().reset_index(drop=True)
    if d.empty:
        return d
    if "kWh_por_PAX" not in d.columns:
        d["kWh_por_PAX"] = d["E_Tr"] / d["PAX"].replace(0, np.nan)
    perfil = _perfil_horario_diario(all_prmte_full, all_fact_full)
    if not perfil.empty:
        d["Fecha"] = pd.to_datetime(d["Fecha"]).dt.normalize()
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

    niveles, sevs, diags = [], [], []
    for _, r in d.iterrows():
        fired = {c: r[zcols[c]] for c in zcols if pd.notna(r.get(zcols[c])) and abs(r[zcols[c]]) >= z_alerta}
        if not fired:
            niveles.append("OK"); sevs.append(0.0); diags.append(""); continue
        sev = max(abs(v) for v in fired.values())
        niveles.append("ANOMALÍA" if sev >= z_fuerte else "ATENCIÓN")
        sevs.append(sev)
        z_en = fired.get("E_Total", fired.get("E_Tr", None))
        if z_en is not None:
            cabeza = "SOBRECONSUMO" if z_en > 0 else "BAJA DE CONSUMO"
        else:
            cabeza = "DESVIACIÓN"
        partes = []
        if z_en is not None and z_en > 0:
            if fired.get("IDE (kWh/km)", 0) >= z_alerta:
                partes.append("por EFICIENCIA: más kWh por km (acoplamiento/conducción/clima)")
            elif fired.get("Servicios", 0) >= z_alerta:
                partes.append("por VOLUMEN: más servicios/km de lo normal para este tipo de día")
            else:
                partes.append("tracción alta sin más servicios -> revisar eficiencia / odómetro")
        elif z_en is not None and z_en < 0:
            if fired.get("Servicios", 0) <= -z_alerta:
                partes.append("por VOLUMEN: menos servicios/km (reducción de oferta?)")
            elif fired.get("IDE (kWh/km)", 0) <= -z_alerta:
                partes.append("por mejor EFICIENCIA: menos kWh por km")
            else:
                partes.append("energía baja -> confirmar que no falten datos del día")
        if fired.get("E_12", 0) >= z_alerta and fired.get("Noche_kWh", 0) >= z_alerta:
            partes.append("CONSUMO PARÁSITO: 12 kV y nocturno altos -> unidades encendidas en cocheras")
        ya = {"E_Total", "E_Tr"}
        if "PARÁSITO" in " ".join(partes):
            ya |= {"E_12", "Noche_kWh"}
        for m, z in sorted(fired.items(), key=lambda kv: -abs(kv[1])):
            if m in ya:
                continue
            direc = "alto" if z > 0 else "bajo"
            partes.append(f"{_METRICAS_ANOM.get(m, m)} {direc} (z={z:+.1f})")
        diags.append(cabeza + " · " + "; ".join(partes) if partes else cabeza)
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
            
        h_idx = next((i for i in range(min(30, len(df))) if 'N° THDR' in str(df.iloc[i].values).upper() or 'N° VIAJE' in str(df.iloc[i].values).upper()), None)
        if h_idx is not None:
            f.seek(0)
            if is_csv:
                try: df = pd.read_csv(f, header=h_idx, encoding='utf-8')
                except UnicodeDecodeError:
                    f.seek(0); df = pd.read_csv(f, header=h_idx, encoding='latin-1')
            else:
                df = pd.read_excel(f, engine=eu, header=h_idx)
                
            df.columns = [str(c).strip() for c in df.columns]
            if 'Fecha' in df.columns:
                df['Fecha'] = pd.to_datetime(df['Fecha'], format='%d-%m-%Y', errors='coerce').dt.normalize()
                df = df[(df['Fecha'].dt.date >= start_date) & (df['Fecha'].dt.date <= end_date)]
            if 'Total a Bordo' in df.columns:
                df['Total a Bordo'] = pd.to_numeric(df['Total a Bordo'], errors='coerce').fillna(0)
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
_CACHE_VERSION = "v13_release_estable"
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
            p1 = df_carga_v1.groupby('Fecha')['Total a Bordo'].sum().reset_index(name='PAX_V1') if not df_carga_v1.empty else pd.DataFrame(columns=['Fecha', 'PAX_V1'])
            p2 = df_carga_v2.groupby('Fecha')['Total a Bordo'].sum().reset_index(name='PAX_V2') if not df_carga_v2.empty else pd.DataFrame(columns=['Fecha', 'PAX_V2'])
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
        
        # --- 1. AUDITORÍA DE CARGA BASE (CONSUMO NOCTURNO) ---
        st.markdown("#### 🌙 Auditoría de Consumo Nocturno (01:00 - 05:00 hrs)")
        
        df_noche = df_prmte[df_prmte['Hora'].isin(['01:00', '02:00', '03:00', '04:00'])]
        
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
    st.markdown("Este módulo cruza la demanda real (Pasajeros) con la fricción operativa (Tiempos de Viaje/Detenciones) para explicar el gasto de Tracción.")
    
    if not df_thdr_v1.empty and not df_thdr_v2.empty and not df_ops.empty and not df_carga_v1.empty:
        # --- 1. PREPARACIÓN DEL CRUCE DE DATOS ---
        def extr_tiempos(df_t, sal_str, lleg_str):
            if df_t.empty: return pd.DataFrame()
            c_sal = next((c for c in df_t.columns if sal_str in str(c).upper() and '_min' in str(c).lower()), None)
            c_lleg = next((c for c in df_t.columns if lleg_str in str(c).upper() and '_min' in str(c).lower()), None)
            if not c_sal or not c_lleg: return pd.DataFrame()
            
            t_v = df_t[['Fecha_Op', c_sal, c_lleg]].dropna().copy()
            t_v['Dur'] = t_v[c_lleg] - t_v[c_sal]
            t_v['Dur'] = t_v['Dur'].apply(lambda x: x + 1440 if x < -500 else x)
            t_v = t_v[(t_v['Dur'] > 30) & (t_v['Dur'] < 120)]
            
            return t_v.groupby('Fecha_Op')['Dur'].mean().reset_index()

        tv1 = extr_tiempos(df_thdr_v1, 'PUERTO', 'LIMACHE')
        tv2 = extr_tiempos(df_thdr_v2, 'LIMACHE', 'PUERTO')
        
        df_tiempos = pd.DataFrame(columns=['Fecha'])
        if not tv1.empty:
            tv1.columns = ['Fecha', 'Tiempo_V1']
            df_tiempos = pd.merge(df_tiempos, tv1, on='Fecha', how='outer') if not df_tiempos.empty else tv1
        if not tv2.empty:
            tv2.columns = ['Fecha', 'Tiempo_V2']
            df_tiempos = pd.merge(df_tiempos, tv2, on='Fecha', how='outer') if not df_tiempos.empty else tv2
            
        df_tiempos['Tiempo_Promedio_Red'] = df_tiempos[['Tiempo_V1', 'Tiempo_V2']].mean(axis=1)
        df_tiempos['Fecha'] = pd.to_datetime(df_tiempos['Fecha']).dt.normalize()

        df_mixto = pd.merge(df_ops, df_tiempos[['Fecha', 'Tiempo_Promedio_Red']], on='Fecha', how='inner')
        df_plot = df_mixto.dropna(subset=['Tiempo_Promedio_Red', 'E_Tr', 'PAX']).copy()
        
        if not df_plot.empty and df_plot['E_Tr'].sum() > 0:
            
            # --- 2. BUBBLE CHART 4D ---
            st.markdown("#### 🫧 Ecosistema Operativo Diario")
            st.caption("Eje X: Lento/Rápido | Eje Y: Consumo Tracción | Tamaño: Volumen de Pasajeros")
            
            fig_mix = px.scatter(df_plot, 
                                 x='Tiempo_Promedio_Red', 
                                 y='E_Tr', 
                                 size='PAX',
                                 color='Tipo Día', 
                                 hover_data=['Fecha', 'IDE (kWh/km)'],
                                 labels={'Tiempo_Promedio_Red': 'Tiempo de Viaje Promedio (min)', 'E_Tr': 'Energía de Tracción (kWh)'},
                                 color_discrete_map={'L': '#005195', 'S': '#E85500', 'D/F': '#2CA02C'})
            
            fig_mix.update_layout(margin=dict(t=20, b=0, l=0, r=0), height=450)
            st.plotly_chart(fig_mix, use_container_width=True)
            
            # --- 3. ANÁLISIS DE CORRELACIÓN ---
            corr_tiempo = df_plot['Tiempo_Promedio_Red'].corr(df_plot['E_Tr'])
            corr_pax = df_plot['PAX'].corr(df_plot['E_Tr'])
            
            c_ins1, c_ins2 = st.columns(2)
            with c_ins1:
                st.info(f"📈 **Impacto del Tiempo (Fricción):** Correlación de **{corr_tiempo:.2f}**. (Cercano a 1 indica que al subir los tiempos de viaje por detenciones largas, la energía se dispara violentamente).")
            with c_ins2:
                st.info(f"👥 **Impacto de la Demanda (Peso):** Correlación de **{corr_pax:.2f}**. (Cercano a 1 indica que el puro peso inercial de la gente obliga a gastar más energía).")

            st.divider()

            # --- 4. EXTRACCIÓN DE DETENCIONES (DWELL TIME) ---
            st.markdown("#### 🛑 El Costo Energético de las Detenciones (Dwell Time)")
            
            estacion_critica = df_carga_v1['Estación Máxima'].mode()[0] if not df_carga_v1.empty else None
            
            if estacion_critica and not df_thdr_v1.empty:
                c_lleg = next((c for c in df_thdr_v1.columns if estacion_critica.upper() in str(c).upper() and 'LLEGADA' in str(c).upper() and '_min' in str(c).lower()), None)
                c_sal = next((c for c in df_thdr_v1.columns if estacion_critica.upper() in str(c).upper() and 'SALIDA' in str(c).upper() and '_min' in str(c).lower()), None)
                
                if c_lleg and c_sal:
                    d_thdr = df_thdr_v1.copy()
                    d_thdr['Dwell'] = d_thdr[c_sal] - d_thdr[c_lleg]
                    d_thdr['Dwell'] = d_thdr['Dwell'].apply(lambda x: x + 1440 if x < -1000 else x)
                    d_thdr = d_thdr[(d_thdr['Dwell'] >= 0) & (d_thdr['Dwell'] < 15)] 
                    
                    dwell_diario = d_thdr.groupby('Fecha_Op')['Dwell'].mean().reset_index()
                    dwell_diario.columns = ['Fecha', 'Dwell_Promedio_Min']
                    dwell_diario['Fecha'] = pd.to_datetime(dwell_diario['Fecha']).dt.normalize()
                    
                    df_dwell = pd.merge(df_plot, dwell_diario, on='Fecha', how='inner')
                    
                    if not df_dwell.empty:
                        df_dwell['Dwell_Secs'] = df_dwell['Dwell_Promedio_Min'] * 60
                        
                        c_dw1, c_dw2 = st.columns([2, 1])
                        with c_dw1:
                            fig_dwell = px.scatter(df_dwell, x='Dwell_Secs', y='IDE (kWh/km)', color='Tipo Día',
                                                   title=f"Efecto de detención en {estacion_critica} vs Desempeño Energético",
                                                   labels={'Dwell_Secs': 'Tiempo de Detención Promedio (Segundos)'},
                                                   trendline="ols",
                                                   color_discrete_sequence=["#E85500", "#005195", "#2CA02C"])
                            st.plotly_chart(fig_dwell, use_container_width=True)
                            
                        with c_dw2:
                            promedio_segundos = df_dwell['Dwell_Secs'].mean()
                            max_segundos = df_dwell['Dwell_Secs'].max()
                            st.markdown("<br><br>", unsafe_allow_html=True)
                            st.metric(f"Dwell Time Prom. en {estacion_critica}", f"{promedio_segundos:.0f} s")
                            st.metric(f"Dwell Time Máximo Registrado", f"{max_segundos:.0f} s", delta="Cuello de Botella Activo", delta_color="inverse")
                            st.caption("*Nota:* Valores sostenidos por sobre los 45-50 segundos indican fallas de flujo de pasajeros en andén, obligando al tren a recuperar itinerario con aceleraciones severas (penalizando el IDE).")
                else:
                    st.info(f"No se detectaron columnas de tiempos para la estación {estacion_critica} en la Vía 1.")
            else:
                st.info("No hay datos de carga suficientes para determinar la estación crítica.")

            st.divider()

            # --- 5. TIEMPO ENTRE ESTACIONES (RUNNING TIME) ---
            st.markdown("#### 🛤️ Fricción en Vía: Tiempos entre Estaciones (Running Time)")
            st.markdown("Identificación de los tramos que más retrasan la operación general de la Vía 1, consumiendo valioso tiempo de itinerario.")
            
            running_avgs = []
            for i in range(len(ESTACIONES)-1):
                e_A = ESTACIONES[i]
                e_B = ESTACIONES[i+1]
                c_s = next((c for c in df_thdr_v1.columns if e_A.upper() in str(c).upper() and 'SALIDA' in str(c).upper() and '_min' in str(c).lower()), None)
                c_l = next((c for c in df_thdr_v1.columns if e_B.upper() in str(c).upper() and 'LLEGADA' in str(c).upper() and '_min' in str(c).lower()), None)
                
                if c_s and c_l:
                    d = df_thdr_v1[c_l] - df_thdr_v1[c_s]
                    d = d.apply(lambda x: x + 1440 if x < -1000 else x).dropna()
                    d = d[(d > 0) & (d < 30)] 
                    if not d.empty:
                        running_avgs.append({'Tramo': f"{e_A[:3]}-{e_B[:3]}", 'Minutos': d.mean()})
                        
            if running_avgs:
                df_run = pd.DataFrame(running_avgs)
                fig_run = px.line(
                    df_run,
                    x='Tramo',
                    y='Minutos',
                    markers=True,
                    title="Tiempo Promedio de Viaje entre Estaciones Consecutivas (V1)",
                    line_shape="linear"
                )
                fig_run.update_traces(line_color="#005195", marker=dict(size=8))
                fig_run.update_layout(xaxis_title="Tramo", yaxis_title="Minutos de Viaje", margin=dict(t=40, b=0, l=0, r=0))
                st.plotly_chart(fig_run, use_container_width=True)
                
                tramo_critico = df_run.loc[df_run['Minutos'].idxmax()]
                st.warning(f"⚠️ **Tramo de mayor fricción:** El segmento **{tramo_critico['Tramo']}** consume en promedio **{tramo_critico['Minutos']:.1f} minutos** del itinerario. Este es el punto principal donde se debe velar por no tener interrupciones de velocidad (precauciones), ya que inyectar tracción aquí es lo más costoso para el sistema global.")

        else:
            st.warning("No hay suficientes datos superpuestos para realizar el análisis multivariante. Revisa las fechas.")
    else: 
        st.info("⚠️ Carga archivos de **THDR (Vía 1 y 2), Facturación/PRMTE/SEAT y Carga de Pasajeros** para habilitar el Microscopio Operacional.")

with tabs[9]:
    st.write("### Flujo y Carga de Pasajeros")
    if not df_carga_v1.empty or not df_carga_v2.empty:
        c_p1, c_p2 = st.columns(2)
        with c_p1:
            st.write("#### Total de Pasajeros por Día")
            df_c1_agg = df_carga_v1.groupby('Fecha')['Total a Bordo'].sum().reset_index() if not df_carga_v1.empty else pd.DataFrame(columns=['Fecha', 'Total a Bordo'])
            df_c2_agg = df_carga_v2.groupby('Fecha')['Total a Bordo'].sum().reset_index() if not df_carga_v2.empty else pd.DataFrame(columns=['Fecha', 'Total a Bordo'])
            
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
            
            # --- KPI GLOBAL DE LA SELECCIÓN ---
            tot_traccion_global = df_reporte['E_Tr'].sum()
            tot_pax_global = df_reporte['PAX'].sum()
            kwh_per_pax_global = (tot_traccion_global / tot_pax_global) if tot_pax_global > 0 else 0
            
            st.info(f"🎯 **KPI Global de Sostenibilidad (Estándar UIC):** En toda la selección analizada, la empresa consumió en promedio **{kwh_per_pax_global:.2f} kWh de tracción por cada pasajero transportado**.")
            
            # --- ITERACIÓN POR TIPO DE JORNADA ---
            tipos_ordenados = ["L", "S", "D/F"]
            nombres_tipos = {"L": "Días Laborales (L)", "S": "Sábados (S)", "D/F": "Domingos y Festivos (D/F)"}
            
            for tipo in tipos_ordenados:
                df_tipo = df_reporte[df_reporte['Tipo Día'] == tipo]
                if df_tipo.empty: continue
                
                with st.expander(f"📌 Análisis de Operación: {nombres_tipos[tipo]}", expanded=True):
                    
                    # 1. MOTOR DE CÁLCULOS MACRO PARA EL TIPO DE DÍA
                    dia_max_ide = df_tipo.loc[df_tipo['IDE (kWh/km)'].idxmax()]
                    dias_validos_ide = df_tipo[df_tipo['IDE (kWh/km)'] > 0]
                    dia_min_ide = dias_validos_ide.loc[dias_validos_ide['IDE (kWh/km)'].idxmin()] if not dias_validos_ide.empty else dia_max_ide
                    
                    tot_tren_km = df_tipo['Tren-Km [km]'].sum()
                    tot_odo = df_tipo['Odómetro [km]'].sum()
                    umr_global = (tot_tren_km / tot_odo * 100) if tot_odo > 0 else 0
                    dia_max_pax = df_tipo.loc[df_tipo['PAX'].idxmax()] if df_tipo['PAX'].sum() > 0 else None

                    # 2. MOTOR DE CÁLCULOS DE ALTA FRECUENCIA PARA EL TIPO DE DÍA
                    fechas_tipo = df_tipo['Fecha'].tolist()
                    
                    # Perfil Horario y Nocturno
                    peak_hr_msg = "No hay datos horarios."
                    noche_msg = ""
                    datos_hr = all_prmte_full if all_prmte_full else all_fact_full
                    if datos_hr:
                        df_hr = pd.DataFrame(datos_hr)
                        df_hr['Fecha'] = pd.to_datetime(df_hr['Fecha'])
                        df_hr_filt = df_hr[df_hr['Fecha'].isin(fechas_tipo)]
                        if not df_hr_filt.empty:
                            hr_agrupado = df_hr_filt.groupby('Hora')['Consumo'].mean()
                            hora_peak = hr_agrupado.idxmax()
                            consumo_peak = hr_agrupado.max()
                            peak_hr_msg = f"La 'Hora Punta Eléctrica' ocurre a las **{hora_peak}** ({consumo_peak:,.0f} kWh prom.)."
                            
                            df_noche = df_hr_filt[df_hr_filt['Hora'].isin(['01:00', '02:00', '03:00', '04:00'])]
                            if not df_noche.empty:
                                noche_diario = df_noche.groupby('Fecha')['Consumo'].sum().reset_index()
                                promedio_noche = noche_diario['Consumo'].mean()
                                max_noche = noche_diario.loc[noche_diario['Consumo'].idxmax()]
                                if max_noche['Consumo'] > (promedio_noche * 1.2) and promedio_noche > 0:
                                    noche_msg = f"🌙 **Alerta Parásita:** Pico de **{max_noche['Consumo']:,.0f} kWh** la madrugada del {max_noche['Fecha'].strftime('%d/%m')}."
                                else:
                                    noche_msg = f"🌙 **Auditoría Nocturna:** Estable ({promedio_noche:,.0f} kWh/noche)."

                    # Cuellos de Botella (Estaciones)
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

                    # Despacho THDR y Velocidades
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

                        # Cálculo de Tiempos
                        def formato_hora(h):
                            if pd.isna(h): return "N/A"
                            if isinstance(h, (datetime, time)): return h.strftime('%H:%M')
                            return str(h)[:5]

                        msg_v1, msg_v2, msg_insight = "", "", ""
                        brecha_max = 0
                        
                        if not t1.empty:
                            c_serv_v1 = t1.columns[0]
                            c_p_sal, c_l_lleg = next((c for c in t1.columns if 'PUERTO' in str(c).upper() and 'SALIDA' in str(c).upper() and '_min' not in str(c).lower()), None), next((c for c in t1.columns if 'LIMACHE' in str(c).upper() and 'LLEGADA' in str(c).upper() and '_min' not in str(c).lower()), None)
                            if c_p_sal and c_l_lleg:
                                t1_v = t1[[c_serv_v1, 'Fecha_Op', c_p_sal, f"{c_p_sal}_min", f"{c_l_lleg}_min"]].dropna().copy()
                                t1_v['Dur'] = t1_v[f"{c_l_lleg}_min"] - t1_v[f"{c_p_sal}_min"]
                                t1_v['Dur'] = t1_v['Dur'].apply(lambda x: x + 1440 if x < -500 else x)
                                t1_v = t1_v[(t1_v['Dur'] > 30) & (t1_v['Dur'] < 120)]
                                if not t1_v.empty:
                                    r_min, r_max = t1_v.loc[t1_v['Dur'].idxmin()], t1_v.loc[t1_v['Dur'].idxmax()]
                                    msg_v1 = f"**V1 (PU→LI) Promedio: {t1_v['Dur'].mean():.1f} min**\n\n- 🟢 *Rápido:* {r_min['Dur']:.0f} min ({r_min['Fecha_Op'].strftime('%d/%m')}, Serv. {r_min[c_serv_v1]}, {formato_hora(r_min[c_p_sal])})\n- 🔴 *Lento:* {r_max['Dur']:.0f} min ({r_max['Fecha_Op'].strftime('%d/%m')}, Serv. {r_max[c_serv_v1]}, {formato_hora(r_max[c_p_sal])})"
                                    brecha_max = max(brecha_max, r_max['Dur'] - r_min['Dur'])

                        if not t2.empty:
                            c_serv_v2 = t2.columns[0]
                            c_l_sal, c_p_lleg = next((c for c in t2.columns if 'LIMACHE' in str(c).upper() and 'SALIDA' in str(c).upper() and '_min' not in str(c).lower()), None), next((c for c in t2.columns if 'PUERTO' in str(c).upper() and 'LLEGADA' in str(c).upper() and '_min' not in str(c).lower()), None)
                            if c_l_sal and c_p_lleg:
                                t2_v = t2[[c_serv_v2, 'Fecha_Op', c_l_sal, f"{c_l_sal}_min", f"{c_p_lleg}_min"]].dropna().copy()
                                t2_v['Dur'] = t2_v[f"{c_p_lleg}_min"] - t2_v[f"{c_l_sal}_min"]
                                t2_v['Dur'] = t2_v['Dur'].apply(lambda x: x + 1440 if x < -500 else x)
                                t2_v = t2_v[(t2_v['Dur'] > 30) & (t2_v['Dur'] < 120)]
                                if not t2_v.empty:
                                    r_min, r_max = t2_v.loc[t2_v['Dur'].idxmin()], t2_v.loc[t2_v['Dur'].idxmax()]
                                    msg_v2 = f"**V2 (LI→PU) Promedio: {t2_v['Dur'].mean():.1f} min**\n\n- 🟢 *Rápido:* {r_min['Dur']:.0f} min ({r_min['Fecha_Op'].strftime('%d/%m')}, Serv. {r_min[c_serv_v2]}, {formato_hora(r_min[c_l_sal])})\n- 🔴 *Lento:* {r_max['Dur']:.0f} min ({r_max['Fecha_Op'].strftime('%d/%m')}, Serv. {r_max[c_serv_v2]}, {formato_hora(r_max[c_l_sal])})"
                                    brecha_max = max(brecha_max, r_max['Dur'] - r_min['Dur'])

                        if msg_v1 or msg_v2:
                            if brecha_max > 10: msg_insight = f"⚠️ *Alta inestabilidad ({brecha_max:.0f} min de brecha máxima).* Fuerte impacto negativo en consumo de tracción."
                            else: msg_insight = f"✅ *Buena regularidad ({brecha_max:.0f} min de brecha máxima).* Operación estable que favorece la conducción eficiente."

                    # --- MAQUETACIÓN DEL SUB-REPORTE POR JORNADA ---
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
                "con estadística robusta y detecta los que se salen de lo normal, nombrando la causa "
                "probable: volumen, eficiencia, auxiliares 12 kV, consumo nocturno/parásito o pico de potencia.")
    if not df_ops.empty:
        df_diag = diagnosticar_anomalias(df_ops, all_prmte_full, all_fact_full)
        if df_diag.empty:
            st.info("No hay días con energía medida en el rango para diagnosticar.")
        else:
            usa_odo = ("Odómetro [km]" in df_diag.columns) and (df_diag["Odómetro [km]"] > 0).any()
            st.caption("IDE calculado con el odómetro real (UMR)." if usa_odo
                       else "⚠ Sin UMR en el rango: el IDE puede no ser exacto.")

            c1, c2, c3 = st.columns(3)
            c1.metric("Días analizados", len(df_diag))
            c2.metric("🔴 Anomalías", int((df_diag["Nivel"] == "ANOMALÍA").sum()))
            c3.metric("🟠 Atención", int((df_diag["Nivel"] == "ATENCIÓN").sum()))

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
                })
            st.dataframe(pd.DataFrame(filas), use_container_width=True)

            anom = df_diag[df_diag["Nivel"] != "OK"].sort_values("Severidad", ascending=False)
            if anom.empty:
                st.success("✓ Sin desviaciones relevantes dentro de cada tipo de día.")
            else:
                st.markdown("#### Días con desviación (ordenados por severidad)")
                for _, r in anom.iterrows():
                    icon = "🔴" if r["Nivel"] == "ANOMALÍA" else "🟠"
                    titulo = f"{icon} {pd.to_datetime(r['Fecha']).strftime('%d-%m-%Y')} ({r['Tipo Día']}) · z={r['Severidad']:.1f} · {r['Nivel']}"
                    with st.expander(titulo, expanded=(r["Nivel"] == "ANOMALÍA")):
                        st.markdown(f"**{r['Diagnóstico']}**")
                        a1, a2, a3, a4 = st.columns(4)
                        a1.metric("E. Total", f"{r['E_Total']:,.0f} kWh")
                        a2.metric("Tracción", f"{r['E_Tr']:,.0f} kWh")
                        a3.metric("12 kV", f"{r['E_12']:,.0f} kWh")
                        a4.metric("IDE", f"{r['IDE (kWh/km)']:,.2f}")
                        b1, b2, b3, b4 = st.columns(4)
                        noche_ok = ("Noche_kWh" in df_diag.columns) and pd.notna(r["Noche_kWh"])
                        b1.metric("Nocturno", f"{r['Noche_kWh']:,.0f} kWh" if noche_ok else "—")
                        b2.metric("Servicios", f"{int(r['Servicios']):,}")
                        b3.metric("PAX", f"{int(r['PAX']):,}" if pd.notna(r["PAX"]) else "—")
                        odo_ok = ("Odómetro [km]" in df_diag.columns) and pd.notna(r["Odómetro [km]"])
                        b4.metric("Odómetro", f"{r['Odómetro [km]']:,.0f} km" if odo_ok else "—")

            st.markdown("#### Tabla completa del diagnóstico")
            cols_show = ["Fecha", "Tipo Día", "E_Total", "E_Tr", "E_12", "Noche_kWh",
                         "IDE (kWh/km)", "Servicios", "PAX", "Nivel", "Diagnóstico"]
            cols_show = [c for c in cols_show if c in df_diag.columns]
            tabla = df_diag[cols_show].copy()
            tabla["Fecha"] = pd.to_datetime(tabla["Fecha"]).dt.strftime("%Y-%m-%d")
            st.dataframe(make_columns_unique(tabla), use_container_width=True)
    else:
        st.info("📂 Sube archivos desde el panel lateral para generar el diagnóstico.")
