import streamlit as st
import pandas as pd
import numpy as np
import re
import holidays
from io import BytesIO
from datetime import datetime, date, time
import plotly.graph_objects as go

# --- 0. SEGURIDAD DE COLUMNAS ---
def make_columns_unique(df):
    if not isinstance(df, pd.DataFrame) or df.empty: return df
    cols = pd.Series(df.columns)
    for dup in cols[cols.duplicated()].unique():
        cols[cols == dup] = [f"{dup}_{i}" if i != 0 else dup for i in range(sum(cols == dup))]
    df.columns = cols
    return df

# --- 1. CONFIGURACIÓN ---
st.set_page_config(page_title="Gestión de Energía - Dashboard SGE", layout="wide", page_icon="🚆")
chile_holidays = holidays.Chile()
st.markdown("""<style>
.stMetric{background-color:#ffffff;padding:20px;border-radius:10px;
border-left:5px solid #005195;box-shadow:0 2px 4px rgba(0,0,0,0.05);}
</style>""", unsafe_allow_html=True)

# --- 2. CONSTANTES DE RED ---
ESTACIONES = [
    'Puerto','Bellavista','Francia','Baron','Portales','Recreo','Miramar',
    'Viña del Mar','Hospital','Chorrillos','El Salto','Valencia','Quilpue',
    'El Sol','El Belloto','Las Americas','La Concepcion','Villa Alemana',
    'Sargento Aldea','Peñablanca','Limache'
]
ESTACIONES_CORTO = ['PU','BE','FR','BA','PO','RE','MI','VM','HO','CH',
                    'ES','VAL','QU','SO','EB','AM','CO','VL','SA','PE','LI']
KM_TRAMO = [0.7,0.7,0.8,1.7,2.1,1.4,0.9,0.9,1.0,1.5,7.4,2.3,1.9,2.0,1.1,1.2,0.9,0.6,1.3,12.73]
KM_ACUM  = [0.0]
for _k in KM_TRAMO: KM_ACUM.append(round(KM_ACUM[-1]+_k, 2))
KM_TOTAL = KM_ACUM[-1]

# Coordenadas reales interpoladas para las 21 estaciones
_ANCHORS_KM  = [0.0, 8.3, 21.4, 28.5, 43.13]
_ANCHORS_LAT = [-33.0385, -33.0264, -33.0453, -33.0426, -32.9843]
_ANCHORS_LON = [-71.6271, -71.5518, -71.4445, -71.3735, -71.2777]
EST_LATS = [float(np.interp(k, _ANCHORS_KM, _ANCHORS_LAT)) for k in KM_ACUM]
EST_LONS = [float(np.interp(k, _ANCHORS_KM, _ANCHORS_LON)) for k in KM_ACUM]

def interpolar_posicion(km_pos):
    km_pos = max(0.0, min(float(km_pos), KM_TOTAL))
    return (float(np.interp(km_pos, KM_ACUM, EST_LATS)),
            float(np.interp(km_pos, KM_ACUM, EST_LONS)))

# ══════════════════════════════════════════════════════════════════════════════
# PERFIL DE VELOCIDAD (metros, km/h)
# Columnas: (km_ini_m, km_fin_m, dist_m, vel_normal, vel_rm)
# ══════════════════════════════════════════════════════════════════════════════
SPEED_PROFILE = [
    (90.6,    122.3,   31.7,   0,   0),
    (122.3,   215.3,   93.0,  52,  43),
    (215.3,   372.6,  157.3,  52,  43),
    (372.6,   577.2,  204.6,  52,  43),
    (577.2,   781.6,  204.4,  52,  43),
    (781.6,  1043.0,  261.4,  52,  43),
    (1043.0, 1377.0,  334.0,  52,  43),
    (1377.0, 1767.0,  390.0,  52,  43),
    (1767.0, 2202.0,  435.0,  42,  34),
    (2202.0, 2592.0,  390.0,  42,  34),
    (2592.0, 2960.5,  368.5,  74,  60),
    (2960.5, 3337.0,  376.5,  74,  60),
    (3337.0, 3448.4,  111.4,  74,  60),
    (3448.4, 3938.4,  490.0,  74,  60),
    (3938.4, 4328.4,  390.0,  66,  54),
    (4328.4, 4758.4,  430.0,  74,  60),
    (4758.4, 5188.4,  430.0,  52,  43),
    (5188.4, 5618.4,  430.0,  52,  43),
    (5618.4, 6034.4,  416.0,  52,  43),
    (6034.4, 6416.4,  382.0,  52,  43),
    (6416.4, 6913.0,  496.6,  74,  60),
    (6913.0, 7405.0,  492.0,  66,  54),
    (7405.0, 7816.4,  411.4,  66,  54),
    (7816.4, 8308.4,  492.0,  66,  54),
    (8308.4, 8695.0,  386.6,  66,  54),
    (8695.0, 9209.8,  514.8,  66,  54),
    (9209.8, 9622.2,  412.4,  66,  54),
    (9622.2,10171.1,  548.9,  66,  54),
    (10171.1,10530.5, 359.4,  52,  43),
    (10530.5,11020.5, 490.0,  74,  60),
    (11020.5,11513.5, 493.0,  74,  60),
    (11513.5,11920.0, 406.5,  74,  60),
    (11920.0,12088.4, 168.4,  74,  60),
    (12088.4,12176.0,  87.6,  74,  60),
    (12176.0,12578.0, 402.0,  74,  60),
    (12578.0,12724.8, 146.8,  74,  60),
    (12724.8,12861.7, 136.9,  74,  60),
    (12861.7,13359.7, 498.0, 120,  99),
    (13359.7,13847.7, 488.0, 120,  99),
    (13847.7,14337.7, 490.0,  74,  60),
    (14337.7,14828.7, 491.0,  52,  43),
    (14828.7,15325.7, 497.0,  52,  43),
    (15325.7,15823.7, 498.0,  52,  43),
    (15823.7,16321.7, 498.0,  52,  43),
    (16321.7,16812.7, 491.0,  52,  43),
    (16812.7,17317.7, 505.0,  52,  43),
    (17317.7,17809.7, 492.0,  52,  43),
    (17809.7,18301.7, 492.0,  74,  60),
    (18301.7,18788.7, 487.0,  74,  60),
    (18788.7,19281.7, 493.0,  74,  60),
    (19281.7,19772.7, 491.0,  74,  60),
    (19772.7,20265.7, 493.0,  74,  60),
    (20265.7,20754.7, 489.0,  74,  60),
    (20754.7,21250.7, 496.0,  66,  54),
    (21250.7,21337.7,  87.0,  52,  43),
    (21337.7,21632.1, 294.4,  52,  43),
    (21632.1,21739.7, 107.6,  74,  60),
    (21739.7,22061.7, 322.0,  74,  60),
    (22061.7,22251.2, 189.5, 102,  84),
    (22251.2,22357.7, 106.5, 102,  84),
    (22357.7,22812.7, 455.0,  74,  60),
    (22812.7,23265.7, 453.0,  74,  60),
    (23265.7,23660.7, 395.0,  74,  60),
    (23660.7,24155.7, 495.0, 102,  84),
    (24155.7,24650.7, 495.0, 102,  84),
    (24650.7,25145.7, 495.0,  74,  60),
    (25145.7,25343.7, 198.0,  74,  60),
    (25343.7,25483.0, 139.3,  74,  60),
    (25483.0,25725.0, 242.0,  74,  60),
    (25725.0,26219.0, 494.0,  74,  60),
    (26219.0,26614.0, 395.0,  74,  60),
    (26614.0,27025.5, 411.5,  74,  60),
    (27025.5,27457.0, 431.5,  74,  60),
    (27457.0,27837.0, 380.0,  74,  60),
    (27837.0,28317.0, 480.0,  74,  60),
    (28317.0,28712.0, 395.0,  74,  60),
    (28712.0,29180.0, 468.0,  74,  60),
    (29180.0,29565.0, 385.0,  74,  60),
    (29565.0,29817.0, 252.0,  74,  60),
    (29817.0,30122.0, 305.0,  74,  60),
    (30122.0,30464.0, 342.0,  66,  54),
    (30464.0,30849.0, 385.0,  74,  60),
    (30849.0,31332.6, 483.6, 102,  84),
    (31332.6,31817.6, 485.0, 120,  99),
    (31817.6,32307.6, 490.0, 120,  99),
    (32307.6,32802.6, 495.0, 120,  99),
    (32802.6,33297.6, 495.0, 120,  99),
    (33297.6,33792.6, 495.0, 120,  99),
    (33792.6,34282.6, 490.0, 120,  99),
    (34282.6,34767.6, 485.0, 120,  99),
    (34767.6,35246.6, 479.0, 120,  99),
    (35246.6,35725.3, 478.7, 120,  99),
    (35725.3,36223.3, 498.0, 102,  84),
    (36223.3,36704.5, 481.2,  74,  60),
    (36704.5,37194.0, 489.5,  74,  60),
    (37194.0,37683.5, 489.5,  74,  60),
    (37683.5,38172.0, 488.5, 102,  84),
    (38172.0,38665.3, 493.3, 120,  99),
    (38665.3,39153.0, 487.7, 120,  99),
    (39153.0,39642.4, 489.4, 120,  99),
    (39642.4,40134.0, 491.6, 120,  99),
    (40134.0,40621.8, 487.8, 120,  99),
    (40621.8,41100.8, 479.0, 120,  99),
    (41100.8,41601.5, 500.7, 120,  99),
    (41601.5,42089.1, 487.6, 102,  84),
    (42089.1,42588.5, 499.4,  66,  54),
    (42588.5,42785.5, 197.0,  66,  54),
    (42785.5,43057.2, 271.7,  42,  34),
    (43057.2,43273.1, 215.9,  42,  34),
    (43273.1,43305.0,  31.9,   0,   0),
]

# ── Pre-calcular perfil tiempo→posición para V1 y V2 ──────────────────────
_MIN_VEL = 5.0   # km/h mínimo para tramos con v=0 (zona de depósito)

def _build_profile(use_rm: bool, via: int):
    """
    Devuelve (km_array_m, cum_time_array_s) — arrays del mismo largo.
    via=1: Puerto→Limache (segmentos en orden creciente de km)
    via=2: Limache→Puerto (segmentos en orden decreciente de km)
    """
    segs = SPEED_PROFILE if via == 1 else list(reversed(SPEED_PROFILE))
    km_pts  = []
    t_pts   = []
    cum_t   = 0.0

    for km_ini, km_fin, dist_m, v_n, v_rm in segs:
        v = (v_rm if use_rm else v_n)
        if v <= 0: v = _MIN_VEL
        seg_s = (dist_m / 1000.0) / v * 3600.0   # segundos

        if via == 1:
            km_pts.append(km_ini)
        else:
            km_pts.append(km_fin)   # recorriendo de mayor a menor

        t_pts.append(cum_t)
        cum_t += seg_s

    # punto final
    last = SPEED_PROFILE[-1] if via == 1 else SPEED_PROFILE[0]
    km_pts.append(last[1] if via == 1 else last[0])
    t_pts.append(cum_t)

    return np.array(km_pts, dtype=float), np.array(t_pts, dtype=float)

# Cachear los 4 perfiles (v1/v2 × normal/rm)
_PROFILES = {
    (1, False): _build_profile(False, 1),
    (1, True ): _build_profile(True,  1),
    (2, False): _build_profile(False, 2),
    (2, True ): _build_profile(True,  2),
}

def km_en_tiempo_real(t_ini_min, t_fin_min, t_actual_min, via, use_rm=False):
    """
    Posición del tren (en km desde Puerto) usando perfil de velocidades real.
    Los tiempos son en minutos (float).
    """
    dur_min = t_fin_min - t_ini_min
    if dur_min <= 0: return 0.0
    elapsed_min = t_actual_min - t_ini_min
    frac = max(0.0, min(1.0, elapsed_min / dur_min))   # 0 → 1

    km_arr, t_arr = _PROFILES[(via, use_rm)]
    total_s = t_arr[-1]
    t_interp = frac * total_s                           # tiempo teórico escalado

    km_m = float(np.interp(t_interp, t_arr, km_arr))   # posición en metros
    km   = km_m / 1000.0                                # → km

    # Para Vía 2 el perfil tiene km decreciente: necesitamos reflejar
    if via == 2:
        km = KM_TOTAL - (km_arr[0]/1000.0 - km)
        # Más simple: interpolar directamente en km decreciente
        # km_arr para v2 va de ~43.3km a ~0.09km
        km = float(np.interp(t_interp, t_arr, km_arr)) / 1000.0

    return max(0.0, min(km, KM_TOTAL))

# ── Función para construir el gráfico de velocidades del perfil ───────────
def fig_perfil_velocidades():
    kms = [(s[0]+s[1])/2/1000 for s in SPEED_PROFILE]
    vels_n = [s[3] if s[3] > 0 else 0 for s in SPEED_PROFILE]
    vels_r = [s[4] if s[4] > 0 else 0 for s in SPEED_PROFILE]
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=kms, y=vels_n, mode='lines', name='Vel. Normal',
                              line=dict(color='#005195', width=2), fill='tozeroy',
                              fillcolor='rgba(0,81,149,0.12)'))
    fig.add_trace(go.Scatter(x=kms, y=vels_r, mode='lines', name='Vel. RM',
                              line=dict(color='#E85500', width=1.5, dash='dot')))
    # Marcar estaciones
    for est, km_est in zip(ESTACIONES, KM_ACUM):
        fig.add_vline(x=km_est, line_width=1, line_dash='dot', line_color='gray')
        fig.add_annotation(x=km_est, y=125, text=est[:3], showarrow=False,
                           font=dict(size=8, color='#555'), textangle=-90)
    fig.update_layout(title='Perfil de velocidades — Vía 1',
                      xaxis_title='km desde Puerto', yaxis_title='km/h',
                      height=300, margin=dict(t=40, b=20))
    return fig

# --- 2b. FUNCIONES DE APOYO ---
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
    except: return 0.0

def get_tipo_dia(fch):
    if fch in chile_holidays or fch.weekday() == 6: return "D/F"
    if fch.weekday() == 5: return "S"
    return "L"

def format_hm_short(minutos_float):
    if pd.isna(minutos_float): return "00:00:00"
    total_seg = int(round(minutos_float * 60))
    h = total_seg // 3600; m = (total_seg % 3600) // 60; s = total_seg % 60
    return f"{h:02d}:{m:02d}:{s:02d}"

# --- 3. MOTOR THDR ---
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
    except: return None

def parsear_fecha_nombre(nombre):
    nombre = str(nombre)
    m = re.search(r'(\d{2})[-_](\d{2})[-_](\d{4})', nombre)
    if m:
        try: return date(int(m.group(3)), int(m.group(2)), int(m.group(1))), f"DD-MM-YYYY ({m.group()})"
        except: pass
    m = re.search(r'(\d{4})[-_](\d{2})[-_](\d{2})', nombre)
    if m:
        try: return date(int(m.group(1)), int(m.group(2)), int(m.group(3))), f"YYYY-MM-DD ({m.group()})"
        except: pass
    m = re.search(r'(\d{8})', nombre)
    if m:
        s = m.group(1)
        try: return date(int(s[4:]), int(s[2:4]), int(s[:2])), f"DDMMYYYY ({s})"
        except: pass
    m = re.search(r'(\d{6})', nombre)
    if m:
        s = m.group(1)
        try: return date(2000+int(s[4:]), int(s[2:4]), int(s[:2])), f"DDMMYY ({s})"
        except: pass
    return None, f"sin fecha en: '{nombre}'"

def procesar_thdr_eficiente(file, start_date, end_date):
    nombre = getattr(file, 'name', str(file))
    diag = {"archivo": nombre, "fecha_parseada": None, "en_rango": None, "filas": 0, "error": None}
    try:
        fch_date, desc = parsear_fecha_nombre(nombre)
        diag["fecha_parseada"] = desc
        if fch_date is None:
            diag["error"] = "No se encontró fecha en el nombre"; return pd.DataFrame(), diag
        diag["en_rango"] = f"{start_date}≤{fch_date}≤{end_date} → {start_date<=fch_date<=end_date}"
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
        col_ref = next((c for c in df.columns
                        if ('PUERTO' in str(c).upper() or 'LIMACHE' in str(c).upper())
                        and 'Salida' in str(c) and '_min' in str(c)), None)
        if col_ref: df['Hora_Ref_Min'] = df[col_ref]
        diag["filas"] = len(df)
        return df, diag
    except Exception as e:
        diag["error"] = str(e); return pd.DataFrame(), diag

# --- 4. PERSISTENCIA EN DISCO ---
import os
DATA_DIRS = {"v1":"data/thdr_v1","v2":"data/thdr_v2","umr":"data/umr",
             "seat":"data/seat","bill":"data/facturacion"}
for _d in DATA_DIRS.values(): os.makedirs(_d, exist_ok=True)

def guardar_archivo(uf, carpeta):
    with open(os.path.join(carpeta, uf.name), "wb") as out: out.write(uf.getbuffer())

def listar_archivos(carpeta):
    exts = ('.xls','.xlsx','.xlsm')
    try: return sorted([os.path.join(carpeta,f) for f in os.listdir(carpeta) if f.lower().endswith(exts)])
    except: return []

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

# --- 5. INICIALIZACIÓN ---
df_ops=pd.DataFrame(); df_thdr_v1=pd.DataFrame(); df_thdr_v2=pd.DataFrame()
all_ops,all_tr,all_seat,all_fact_full,all_prmte_full=[],[],[],[],[]

# --- 6. SIDEBAR ---
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
    for _ul,_ca in [(f_v1,DATA_DIRS["v1"]),(f_v2,DATA_DIRS["v2"]),(f_umr,DATA_DIRS["umr"]),
                    (f_seat_files,DATA_DIRS["seat"]),(f_bill_files,DATA_DIRS["bill"])]:
        for uf in (_ul or []):
            dest=os.path.join(_ca,uf.name)
            if not os.path.exists(dest): guardar_archivo(uf,_ca)
    st.divider()
    with st.expander("🗂️ Archivos guardados"):
        _labels={"v1":"Vía 1","v2":"Vía 2","umr":"UMR","seat":"SEAT","bill":"Facturación"}
        for _key,_carpeta in DATA_DIRS.items():
            _arch=listar_archivos(_carpeta)
            if _arch:
                st.markdown(f"**{_labels[_key]}** — {len(_arch)} archivo(s)")
                for _a in _arch:
                    ca2,cb2=st.columns([5,1]); ca2.caption(os.path.basename(_a))
                    if cb2.button("🗑️",key=f"del_{_a}"): os.remove(_a); st.rerun()
            else: st.caption(f"{_labels[_key]}: sin archivos")

f_v1_all   = combinar_fuentes(f_v1,         DATA_DIRS["v1"])
f_v2_all   = combinar_fuentes(f_v2,         DATA_DIRS["v2"])
f_umr_all  = combinar_fuentes(f_umr,        DATA_DIRS["umr"])
f_seat_all = combinar_fuentes(f_seat_files, DATA_DIRS["seat"])
f_bill_all = combinar_fuentes(f_bill_files, DATA_DIRS["bill"])

_CACHE_VERSION = "v5_perfil_vel"
_cache_key = (_CACHE_VERSION, str(start_date), str(end_date),
              tuple(sorted(f.name for f in f_v1_all)), tuple(sorted(f.name for f in f_v2_all)),
              tuple(sorted(f.name for f in f_umr_all)), tuple(sorted(f.name for f in f_seat_all)),
              tuple(sorted(f.name for f in f_bill_all)))
_hay_archivos = any([f_v1_all,f_v2_all,f_umr_all,f_seat_all,f_bill_all])
_recalcular   = st.session_state.get('_cache_key') != _cache_key

if _hay_archivos and not _recalcular and 'df_ops' in st.session_state:
    df_ops=st.session_state['df_ops']; df_thdr_v1=st.session_state['df_thdr_v1']
    df_thdr_v2=st.session_state['df_thdr_v2']; all_tr=st.session_state['all_tr']
    all_seat=st.session_state['all_seat']; all_fact_full=st.session_state['all_fact_full']
    all_prmte_full=st.session_state['all_prmte_full']

_errores_proc={}

if _hay_archivos and _recalcular:
    if f_umr_all:
        for f in f_umr_all:
            try:
                eu="xlrd" if f.name.lower().endswith(".xls") else "openpyxl"
                xl=pd.ExcelFile(f,engine=eu)
                for sn in xl.sheet_names:
                    f.seek(0); df_raw=pd.read_excel(f,sheet_name=sn,header=None,engine=eu)
                    h_r=next((i for i in range(min(100,len(df_raw)))
                               if any(k in str(df_raw.iloc[i].tolist()).upper() for k in ['FECHA','ODO','KILOM'])),None)
                    if h_r is not None:
                        f.seek(0); df_p=pd.read_excel(f,sheet_name=sn,header=h_r,engine=eu)
                        df_p.columns=[str(c).upper().replace('Ó','O').strip() for c in df_p.columns]
                        c_f=next((c for c in df_p.columns if 'FECHA' in c),None)
                        c_o=next((c for c in df_p.columns if 'ODO' in c),None)
                        c_t=next((c for c in df_p.columns if 'KM' in c),None)
                        if c_f and c_o:
                            df_p['_dt']=pd.to_datetime(df_p[c_f],errors='coerce').dt.normalize()
                            mask=(df_p['_dt'].dt.date>=start_date)&(df_p['_dt'].dt.date<=end_date)
                            for _,r in df_p[mask].dropna(subset=['_dt']).iterrows():
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
                                            all_tr.append({"Tren":t,"Fecha":v_f.normalize(),
                                                           "Valor":parse_latam_number(df_raw.iloc[k,j])})
            except Exception as e: _errores_proc[f.name]=f"UMR: {e}"
    if f_seat_all:
        for f in f_seat_all:
            try:
                es="xlrd" if f.name.lower().endswith(".xls") else "openpyxl"
                df_s=pd.read_excel(f,header=None,engine=es)
                for i in range(len(df_s)):
                    fs=pd.to_datetime(df_s.iloc[i,1],errors='coerce')
                    if pd.notna(fs):
                        fs=fs.normalize()
                        if start_date<=fs.date()<=end_date:
                            all_seat.append({"Fecha":fs,"E_Total":parse_latam_number(df_s.iloc[i,3]),
                                             "E_Tr":parse_latam_number(df_s.iloc[i,5]),
                                             "E_12":parse_latam_number(df_s.iloc[i,7])})
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
                        for _,r in df_ff.dropna(subset=['dt']).iterrows():
                            if "TOTAL" in str(r[c_f]).upper(): continue
                            v=abs(parse_latam_number(r[c_v]))
                            all_fact_full.append({"Fecha":r['dt'].normalize(),"Hora":f"{r['dt'].hour:02d}:00",
                                                   "15min":f"{r['dt'].hour:02d}:{(r['dt'].minute//15)*15:02d}","Consumo":v})
                    if 'PRMTE' in sn.upper():
                        f.seek(0); df_pd_raw=pd.read_excel(f,sheet_name=sn,header=None,engine=eb)
                        h=next((i for i in range(min(20,len(df_pd_raw)))
                                if any(k in str(df_pd_raw.iloc[i]).upper() for k in ['AÑO','ANO','YEAR'])),0)
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
                                return pd.Timestamp(year=int(r[c_anio]),month=int(r[c_mes]),
                                                    day=int(r[c_dia]),hour=int(r[c_hora]),minute=min_)
                            except: return pd.NaT
                        df_pd['ts']=df_pd.apply(_build_ts,axis=1)
                        cols_retiro=[c for c in df_pd.columns if 'Retiro_Energia_Activa' in str(c)]
                        if not cols_retiro:
                            cols_retiro=[c for c in df_pd.columns if 'RETIRO' in str(c).upper()
                                         or ('ENERGIA' in str(c).upper() and 'ACTIVA' in str(c).upper())]
                        for _,r in df_pd.dropna(subset=['ts']).iterrows():
                            ts=r['ts']
                            if pd.isna(ts) or not (start_date<=ts.date()<=end_date): continue
                            consumo=sum(parse_latam_number(r.get(c,0)) for c in cols_retiro)
                            all_prmte_full.append({"Fecha":ts.normalize(),"Hora":f"{ts.hour:02d}:00",
                                                    "15min":f"{ts.hour:02d}:{ts.minute:02d}","Consumo":consumo})
            except Exception as e: _errores_proc[f.name]=f"Factura/PRMTE: {e}"
    if _errores_proc: st.session_state['_errores_proc']=_errores_proc

    if all_ops:
        df_ops=pd.DataFrame(all_ops).groupby("Fecha").agg(
            {"Odómetro [km]":"sum","Tren-Km [km]":"sum","Tipo Día":"first"}).reset_index()
        df_f_d=(pd.DataFrame(all_fact_full).groupby("Fecha")["Consumo"].sum().reset_index()
                .rename(columns={"Consumo":"E_Fact"}) if all_fact_full else pd.DataFrame(columns=["Fecha","E_Fact"]))
        df_p_d=(pd.DataFrame(all_prmte_full).groupby("Fecha")["Consumo"].sum().reset_index()
                .rename(columns={"Consumo":"E_Prmte"}) if all_prmte_full else pd.DataFrame(columns=["Fecha","E_Prmte"]))
        df_s_d=(pd.DataFrame(all_seat).groupby("Fecha").agg({"E_Total":"sum","E_Tr":"sum","E_12":"sum"}).reset_index()
                .rename(columns={"E_Total":"E_Seat_T","E_Tr":"E_Seat_Tr","E_12":"E_Seat_12"}) if all_seat
                else pd.DataFrame(columns=["Fecha","E_Seat_T","E_Seat_Tr","E_Seat_12"]))
        for dff in [df_ops,df_f_d,df_p_d,df_s_d]: dff['Fecha']=pd.to_datetime(dff['Fecha']).dt.normalize()
        df_ops=(df_ops.merge(df_f_d,on="Fecha",how="left").merge(df_p_d,on="Fecha",how="left")
                      .merge(df_s_d,on="Fecha",how="left").fillna(0))
        def jerarquia(row):
            if row['E_Fact']>0:     tot,src=row['E_Fact'],"Factura"
            elif row['E_Prmte']>0:  tot,src=row['E_Prmte'],"PRMTE"
            elif row['E_Seat_T']>0: tot,src=row['E_Seat_T'],"SEAT"
            else: return 0,0,0,0,0,"N/A"
            r_tr=row['E_Seat_Tr']/row['E_Seat_T'] if row['E_Seat_T']>0 else 0.85
            r_12=row['E_Seat_12']/row['E_Seat_T'] if row['E_Seat_T']>0 else 0.15
            return tot,tot*r_tr,tot*r_12,r_tr*100,r_12*100,src
        df_ops[['E_Total','E_Tr','E_12','% Tracción','% 12 kV','Fuente']]=df_ops.apply(jerarquia,axis=1,result_type='expand')
        df_ops['IDE (kWh/km)']=df_ops.apply(lambda r: r['E_Tr']/r['Odómetro [km]'] if r['Odómetro [km]']>0 else 0,axis=1)

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
    st.session_state.update({'df_ops':df_ops,'df_thdr_v1':df_thdr_v1,'df_thdr_v2':df_thdr_v2,
                              'all_tr':all_tr,'all_seat':all_seat,'all_fact_full':all_fact_full,
                              'all_prmte_full':all_prmte_full,'_cache_key':_cache_key})

# --- 7. TABS ---
tabs=st.tabs(["📊 Resumen","📑 Operaciones","📑 Trenes","⚡ Energía","⚖️ Comparación hr",
              "📈 Regresión","🚨 Atípicos","📋 THDR","🔬 Servicios vs Energía","🗺️ Mapa de Trenes"])

# TAB 0
with tabs[0]:
    _ep=st.session_state.get('_errores_proc',{})
    if _ep:
        with st.expander(f"⚠️ {len(_ep)} archivo(s) con error",expanded=True):
            for _n,_m in _ep.items(): st.error(f"**{_n}**: {_m}")
    if not df_ops.empty:
        c1,c2,c3=st.columns(3)
        c1.metric("Odómetro Total",f"{df_ops['Odómetro [km]'].sum():,.1f} km")
        c2.metric("Tren-Km Total", f"{df_ops['Tren-Km [km]'].sum():,.1f} km")
        c3.metric("IDE Promedio",  f"{df_ops['IDE (kWh/km)'].mean():.4f} kWh/km")
        st.plotly_chart(go.Figure(go.Bar(x=df_ops['Fecha'],y=df_ops['Odómetro [km]'],marker_color="#005195"))
                        .update_layout(title="Odómetro Diario",xaxis_title="Fecha",yaxis_title="km"),use_container_width=True)
        st.plotly_chart(go.Figure(go.Scatter(x=df_ops['Fecha'],y=df_ops['IDE (kWh/km)'],mode='lines+markers',line=dict(color="#E85500")))
                        .update_layout(title="IDE Diario (kWh/km)",xaxis_title="Fecha",yaxis_title="kWh/km"),use_container_width=True)
    else: st.info("📂 Sube archivos desde el panel lateral para ver el resumen.")

# TAB 1
with tabs[1]:
    if not df_ops.empty:
        dv=df_ops.copy(); dv['Fecha']=dv['Fecha'].dt.strftime('%Y-%m-%d')
        st.dataframe(make_columns_unique(dv).style.format({'Odómetro [km]':"{:,.1f}",'Tren-Km [km]':"{:,.1f}",
            'E_Total':"{:,.0f}",'E_Tr':"{:,.0f}",'E_12':"{:,.0f}",'% Tracción':"{:.1f}%",
            '% 12 kV':"{:.1f}%",'IDE (kWh/km)':"{:.4f}"}),use_container_width=True)
    else: st.info("📂 Sin datos de operación.")

# TAB 2
with tabs[2]:
    if all_tr:
        df_tr=pd.DataFrame(all_tr); df_tr['Fecha']=df_tr['Fecha'].dt.strftime('%Y-%m-%d')
        pivot=df_tr.pivot_table(index='Tren',columns='Fecha',values='Valor',aggfunc='sum').fillna(0)
        st.dataframe(pivot.style.format("{:,.1f}"),use_container_width=True)
    else: st.info("📂 Sin datos de trenes.")

# TAB 3
with tabs[3]:
    e_tabs=st.tabs(["🔹 SEAT","🔹 PRMTE","🔹 Facturación"])
    with e_tabs[0]:
        if all_seat:
            dsv=pd.DataFrame(all_seat); dsv['Fecha']=dsv['Fecha'].dt.strftime('%Y-%m-%d')
            st.dataframe(dsv.style.format({'E_Total':"{:,.0f}",'E_Tr':"{:,.0f}",'E_12':"{:,.0f}"}),use_container_width=True)
        else: st.info("📂 Sin datos SEAT.")
    with e_tabs[1]:
        if all_prmte_full:
            dp=pd.DataFrame(all_prmte_full); dp['Fecha_Str']=dp['Fecha'].dt.strftime('%Y-%m-%d')
            pr=dp.loc[dp['Consumo'].idxmax()]
            m1,m2,m3=st.columns(3)
            m1.metric("Total kWh",f"{dp['Consumo'].sum():,.0f}")
            m2.metric("Días cargados",f"{dp['Fecha_Str'].nunique()}")
            m3.metric("Pico 15 min",f"{pr['Consumo']:,.0f} kWh",f"{pr['Fecha_Str']} {pr['15min']}")
            fp=st.selectbox("Fecha",sorted(dp['Fecha_Str'].unique()),key="prmte_fecha")
            vp=st.radio("Vista",["15 min","Horario","Diario"],horizontal=True,key="prmte_vista")
            dps=dp[dp['Fecha_Str']==fp]
            if vp=="15 min":   dsh=dps.groupby("15min")["Consumo"].sum().reset_index().rename(columns={"15min":"Franja","Consumo":"kWh"}).sort_values("Franja")
            elif vp=="Horario":dsh=dps.groupby("Hora")["Consumo"].sum().reset_index().rename(columns={"Hora":"Franja","Consumo":"kWh"}).sort_values("Franja")
            else:              dsh=dp.groupby("Fecha_Str")["Consumo"].sum().reset_index().rename(columns={"Fecha_Str":"Franja","Consumo":"kWh"})
            fig_p=go.Figure(go.Bar(x=dsh['Franja'],y=dsh['kWh'],marker_color='#005195',
                                    hovertemplate='<b>%{x}</b><br>%{y:,.0f} kWh<extra></extra>'))
            fig_p.update_layout(title=f"PRMTE — {fp} ({vp})" if vp!="Diario" else "PRMTE — Consumo diario",
                                xaxis_title="Franja",yaxis_title="kWh",xaxis=dict(tickangle=-45),height=380)
            st.plotly_chart(fig_p,use_container_width=True)
        else: st.info("📂 Sin datos PRMTE.")
    with e_tabs[2]:
        if all_fact_full:
            df_f=pd.DataFrame(all_fact_full); df_f['Fecha_Str']=df_f['Fecha'].dt.strftime('%Y-%m-%d')
            st.dataframe(df_f.groupby("Fecha_Str")["Consumo"].sum().reset_index()
                         .style.format({'Consumo':"{:,.0f}"}),use_container_width=True)
        else: st.info("📂 Sin datos de facturación.")

# TAB 4
with tabs[4]:
    st.header("⚖️ Comparación Horaria")
    if all_prmte_full or all_fact_full:
        fuentes={}
        if all_prmte_full:
            dph=pd.DataFrame(all_prmte_full); dph['Hora_int']=dph['Hora'].str[:2].astype(int)
            fuentes['PRMTE']=dph.groupby('Hora_int')['Consumo'].sum().reset_index()
        if all_fact_full:
            dfh=pd.DataFrame(all_fact_full); dfh['Hora_int']=dfh['Hora'].str[:2].astype(int)
            fuentes['Factura']=dfh.groupby('Hora_int')['Consumo'].sum().reset_index()
        fig=go.Figure()
        for nb,dfh2 in fuentes.items():
            fig.add_trace(go.Bar(x=dfh2['Hora_int'],y=dfh2['Consumo'],name=nb,
                                  marker_color={'PRMTE':'#005195','Factura':'#E85500'}.get(nb,'#888')))
        fig.update_layout(title="Consumo Acumulado por Hora",barmode='group',xaxis_title="Hora",yaxis_title="kWh")
        st.plotly_chart(fig,use_container_width=True)
    else: st.info("📂 Sube datos de PRMTE o Facturación para comparar.")

# TAB 5
with tabs[5]:
    st.header("📈 Regresión IDE vs Odómetro")
    if not df_ops.empty and df_ops['IDE (kWh/km)'].sum()>0:
        dr2=df_ops[df_ops['IDE (kWh/km)']>0].copy()
        cmap={"L":"#005195","S":"#FFA500","D/F":"#E85500"}
        fig=go.Figure()
        for tp,grp in dr2.groupby('Tipo Día'):
            fig.add_trace(go.Scatter(x=grp['Odómetro [km]'],y=grp['IDE (kWh/km)'],mode='markers',
                                     name=tp,marker=dict(color=cmap.get(tp,'#888'),size=8)))
        xa,ya=dr2['Odómetro [km]'].values,dr2['IDE (kWh/km)'].values
        if len(xa)>=2:
            coef=np.polyfit(xa,ya,1); xl=np.linspace(xa.min(),xa.max(),100)
            r2=np.corrcoef(xa,ya)[0,1]**2
            fig.add_trace(go.Scatter(x=xl,y=np.polyval(coef,xl),mode='lines',
                                     name=f'Tendencia (R²={r2:.3f})',line=dict(color='gray',dash='dash',width=2)))
        fig.update_layout(title='IDE vs Odómetro por Tipo de Día',xaxis_title='Odómetro [km]',yaxis_title='kWh/km')
        st.plotly_chart(fig,use_container_width=True)
    else: st.info("📂 Sin datos suficientes para regresión.")

# TAB 6
with tabs[6]:
    st.header("🚨 Detección de Atípicos")
    if not df_ops.empty and df_ops['IDE (kWh/km)'].sum()>0:
        dat=df_ops[df_ops['IDE (kWh/km)']>0].copy()
        media,std=dat['IDE (kWh/km)'].mean(),dat['IDE (kWh/km)'].std()
        umbral=st.slider("Umbral σ",1.0,3.0,2.0,0.1)
        dat['Atípico']=(dat['IDE (kWh/km)']-media).abs()>umbral*std
        c1,c2=st.columns(2); c1.metric("Media IDE",f"{media:.4f}"); c2.metric("Atípicos",int(dat['Atípico'].sum()))
        fig=go.Figure()
        fig.add_trace(go.Scatter(x=dat[~dat['Atípico']]['Fecha'],y=dat[~dat['Atípico']]['IDE (kWh/km)'],
                                  mode='markers',name='Normal',marker=dict(color='#005195')))
        fig.add_trace(go.Scatter(x=dat[dat['Atípico']]['Fecha'],y=dat[dat['Atípico']]['IDE (kWh/km)'],
                                  mode='markers',name='Atípico',marker=dict(color='red',size=10,symbol='x')))
        fig.add_hline(y=media+umbral*std,line_dash="dash",line_color="orange",annotation_text=f"+{umbral}σ")
        fig.add_hline(y=media-umbral*std,line_dash="dash",line_color="orange",annotation_text=f"-{umbral}σ")
        fig.update_layout(title="IDE Diario con Atípicos",xaxis_title="Fecha",yaxis_title="kWh/km")
        st.plotly_chart(fig,use_container_width=True)
    else: st.info("📂 Sin datos suficientes para detección de atípicos.")

# TAB 7
def render_via_thdr(df_via, label):
    if df_via.empty: st.info(f"📂 No hay datos para {label}."); return
    df=df_via.copy(); df['Fecha']=df['Fecha_Op'].dt.strftime('%Y-%m-%d')
    c1,c2,c3,c4=st.columns(4)
    c1.metric("Total viajes",len(df)); c2.metric("Tren-Km",f"{df['Tren-Km'].sum():,.1f}")
    c3.metric("Días cargados",df['Fecha'].nunique())
    c4.metric("Viajes doble (M)",(df['Unidad'].astype(str).str.strip()=='M').sum())
    resumen=df.groupby('Fecha').agg(Viajes=('Unidad','count'),
                                     Viajes_M=('Unidad',lambda x:(x.astype(str).str.strip()=='M').sum()),
                                     TrenKm=('Tren-Km','sum')).reset_index()
    st.dataframe(resumen.style.format({'TrenKm':"{:,.1f}"}),use_container_width=True)
    fecha_sel=st.selectbox(f"Seleccionar fecha ({label})",sorted(df['Fecha'].unique()),key=f"sel_{label}")
    df_sel=df[df['Fecha']==fecha_sel].copy()
    for cm in [c for c in df_sel.columns if '_min' in c]:
        df_sel[cm.replace('_min','_hms')]=df_sel[cm].apply(format_hm_short)
    cols_base=['Viaje','Tren','Hora Salida Programada','Motriz 1','Motriz 2','Unidad','Maquinista','Tren-Km']
    cols_hms=[c for c in df_sel.columns if '_hms' in c]
    cols_show=[c for c in cols_base+cols_hms if c in df_sel.columns]
    st.dataframe(make_columns_unique(df_sel[cols_show]).reset_index(drop=True),use_container_width=True)
    st.caption(f"{len(df_sel)} viajes el {fecha_sel}")

with tabs[7]:
    st.header("📋 Análisis THDR")
    tv1,tv2=st.tabs(["🔵 Vía 1 (Puerto → Limache)","🟠 Vía 2 (Limache → Puerto)"])
    with tv1: render_via_thdr(df_thdr_v1,"Vía 1")
    with tv2: render_via_thdr(df_thdr_v2,"Vía 2")

# TAB 8
with tabs[8]:
    st.header("🔬 Servicios vs Consumo de Energía (15 min)")
    _tp=len(all_prmte_full)>0; _tt=not df_thdr_v1.empty or not df_thdr_v2.empty
    if not _tp and not _tt: st.info("📂 Sube PRMTE y THDR para este análisis."); st.stop()
    col_av,col_at=st.columns(2)
    col_av.metric("PRMTE","✅" if _tp else "❌"); col_at.metric("THDR","✅" if _tt else "❌")
    def s2m(s):
        try: h,m=map(int,s.split(':')); return h*60+m
        except: return 0
    df_svc=pd.DataFrame()
    if _tt:
        partes=[df for df in [df_thdr_v1,df_thdr_v2] if not df.empty]
        dta=pd.concat(partes,ignore_index=True); dta['Fecha_str']=dta['Fecha_Op'].dt.strftime('%Y-%m-%d')
        def _ps(row):
            vals=[row[c] for c in row.index if 'Salida' in c and '_min' in c and pd.notna(row[c])]; return min(vals) if vals else np.nan
        def _ul(row):
            vals=[row[c] for c in row.index if 'Llegada' in c and '_min' in c and pd.notna(row[c])]; return max(vals) if vals else np.nan
        dta['t_ini']=dta.apply(_ps,axis=1); dta['t_fin']=dta.apply(_ul,axis=1)
        dta=dta.dropna(subset=['t_ini','t_fin'])
        def _kf(t_i,t_f,t_fr,un):
            dur=t_f-t_i
            if dur<=0: return 0.0
            dist=KM_TOTAL*(2 if str(un).strip()=='M' else 1)
            return round((dist/dur)*max(0.0,min(t_f,t_fr+15)-max(t_i,t_fr)),3)
        tf_all=[f"{h:02d}:{m:02d}" for h in range(24) for m in range(0,60,15)]
        filas=[]
        for fg,grp in dta.groupby('Fecha_str'):
            for fr in tf_all:
                t_f=s2m(fr); mask=(grp['t_ini']<=t_f)&(grp['t_fin']>t_f)
                if mask.sum()==0: continue
                ga=grp[mask]
                filas.append({'Fecha':fg,'Franja':fr,'Servicios':int(mask.sum()),
                               'Servicios_M':int((ga['Unidad'].astype(str).str.strip()=='M').sum()),
                               'Tren_Km':sum(_kf(r['t_ini'],r['t_fin'],t_f,r['Unidad']) for _,r in ga.iterrows())})
        if filas: df_svc=pd.DataFrame(filas)
    df_en=pd.DataFrame()
    if _tp:
        dpr=pd.DataFrame(all_prmte_full); dpr['Fecha']=dpr['Fecha'].dt.strftime('%Y-%m-%d')
        df_en=(dpr.groupby(['Fecha','15min'])['Consumo'].sum().reset_index()
               .rename(columns={'15min':'Franja','Consumo':'kWh'}))
    if df_svc.empty and df_en.empty: st.warning("Sin datos suficientes."); st.stop()
    if not df_svc.empty and not df_en.empty:
        dm=pd.merge(df_en,df_svc,on=['Fecha','Franja'],how='outer').fillna(0)
    elif not df_en.empty:
        dm=df_en.copy(); dm['Servicios']=0; dm['Servicios_M']=0; dm['Tren_Km']=0.0
    else:
        dm=df_svc.copy(); dm['kWh']=0
    if 'Tren_Km' not in dm.columns: dm['Tren_Km']=0.0
    dm['_o']=dm['Franja'].apply(s2m); dm=dm.sort_values(['Fecha','_o']).drop(columns='_o')
    fd=sorted(dm['Fecha'].unique())
    if not fd: st.warning("Sin fechas."); st.stop()
    st.divider()
    cf1,cf2,cf3=st.columns([2,2,1])
    with cf1: modo=st.radio("Vista",["Por día","Promedio del período"],horizontal=True)
    with cf2:
        if modo=="Por día":
            fsel=st.selectbox("Fecha",fd); dp2=dm[dm['Fecha']==fsel].copy()
        else:
            dp2=dm.groupby('Franja').agg(kWh=('kWh','mean'),Servicios=('Servicios','mean'),
                                          Servicios_M=('Servicios_M','mean'),Tren_Km=('Tren_Km','mean')).reset_index()
            dp2['_o']=dp2['Franja'].apply(s2m); dp2=dp2.sort_values('_o').drop(columns='_o')
    with cf3: mm=st.checkbox("Solo doble (M)",value=False)
    cs='Servicios_M' if mm else 'Servicios'; ls='Tracción doble' if mm else 'Servicios totales'
    if not dp2.empty:
        m1,m2,m3,m4,m5=st.columns(5)
        m1.metric("Total kWh",f"{dp2['kWh'].sum():,.0f}")
        m2.metric("Pico kWh",f"{dp2['kWh'].max():,.0f}",dp2.loc[dp2['kWh'].idxmax(),'Franja'])
        m3.metric("Total servicios",f"{dp2[cs].sum():.0f}")
        m4.metric("Pico servicios",f"{dp2[cs].max():.0f}",dp2.loc[dp2[cs].idxmax(),'Franja'] if dp2[cs].max()>0 else "—")
        m5.metric("Tren-Km",f"{dp2['Tren_Km'].sum():,.1f} km")
    st.divider()
    if not dp2.empty:
        fd2=go.Figure()
        fd2.add_trace(go.Bar(x=dp2['Franja'],y=dp2['kWh'],name='Energía PRMTE (kWh)',
                              marker_color='rgba(0,81,149,0.7)',yaxis='y1'))
        fd2.add_trace(go.Scatter(x=dp2['Franja'],y=dp2[cs],name=ls,mode='lines+markers',
                                  line=dict(color='#E85500',width=2),marker=dict(size=5),yaxis='y2'))
        if dp2['Tren_Km'].sum()>0:
            fd2.add_trace(go.Scatter(x=dp2['Franja'],y=dp2['Tren_Km'],name='Tren-Km',mode='lines',
                                      line=dict(color='#00AA44',width=2,dash='dot'),yaxis='y3'))
        fd2.update_layout(
            title=(f"Energía vs Servicios — {fsel}" if modo=="Por día" else f"Promedio {fd[0]} a {fd[-1]}"),
            xaxis=dict(title="Franja 15 min",tickangle=-45,tickmode='array',tickvals=dp2['Franja'][::4].tolist()),
            yaxis=dict(title="kWh",side='left',showgrid=True),
            yaxis2=dict(title="Servicios",side='right',overlaying='y',showgrid=False),
            yaxis3=dict(title="Tren-Km",side='right',overlaying='y',showgrid=False,anchor='free',position=1.0,showticklabels=False),
            legend=dict(orientation='h',y=1.08),hovermode='x unified',height=450)
        st.plotly_chart(fd2,use_container_width=True)
    dco=dp2.dropna(subset=['kWh',cs]); dco=dco[(dco['kWh']>0)&(dco[cs]>0)]
    if len(dco)>=5:
        st.divider(); corr=np.corrcoef(dco['kWh'].values,dco[cs].values)[0,1]
        st.subheader(f"📐 Correlación: **{corr:.3f}**")
        coef=np.polyfit(dco[cs].values,dco['kWh'].values,1)
        xl=np.linspace(dco[cs].min(),dco[cs].max(),100)
        fsc=go.Figure()
        fsc.add_trace(go.Scatter(x=dco[cs],y=dco['kWh'],mode='markers',text=dco['Franja'],
                                  hovertemplate='<b>%{text}</b><br>Servicios:%{x}<br>kWh:%{y:,.0f}<extra></extra>',
                                  marker=dict(color='#005195',size=7,opacity=0.7)))
        fsc.add_trace(go.Scatter(x=xl,y=np.polyval(coef,xl),mode='lines',
                                  line=dict(color='#E85500',dash='dash'),name=f'R²={corr**2:.3f}'))
        fsc.update_layout(title='Dispersión Servicios vs kWh',xaxis_title=ls,yaxis_title='kWh',height=380)
        st.plotly_chart(fsc,use_container_width=True)
    with st.expander("📋 Ver tabla"):
        cs2=[c for c in ['Franja','kWh','Servicios','Servicios_M','Tren_Km'] if c in dp2.columns]
        st.dataframe(dp2[cs2].style.format({'kWh':'{:,.1f}','Servicios':'{:.1f}','Servicios_M':'{:.1f}','Tren_Km':'{:.2f}'}),
                     use_container_width=True,height=300)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 9 — MAPA DE TRENES (con perfil de velocidades real)
# ══════════════════════════════════════════════════════════════════════════════
with tabs[9]:
    st.header("🗺️ Mapa de Trenes — Posición calculada por perfil de velocidades")

    _tiene_thdr_m = not df_thdr_v1.empty or not df_thdr_v2.empty
    if not _tiene_thdr_m:
        st.info("📂 Sube archivos THDR (Vía 1 y/o Vía 2) para ver el mapa."); st.stop()

    # ── Preparar dataframe unificado ─────────────────────────────────────────
    partes_m=[]
    if not df_thdr_v1.empty:
        tmp=df_thdr_v1.copy(); tmp['Via']=1; partes_m.append(tmp)
    if not df_thdr_v2.empty:
        tmp=df_thdr_v2.copy(); tmp['Via']=2; partes_m.append(tmp)
    df_m=pd.concat(partes_m,ignore_index=True)
    df_m['Fecha_str']=df_m['Fecha_Op'].dt.strftime('%Y-%m-%d')

    def _ps2(row):
        vals=[row[c] for c in row.index if 'Salida' in c and '_min' in c and pd.notna(row[c])]; return min(vals) if vals else np.nan
    def _ul2(row):
        vals=[row[c] for c in row.index if 'Llegada' in c and '_min' in c and pd.notna(row[c])]; return max(vals) if vals else np.nan
    df_m['t_ini']=df_m.apply(_ps2,axis=1); df_m['t_fin']=df_m.apply(_ul2,axis=1)
    df_m=df_m.dropna(subset=['t_ini','t_fin']); df_m=df_m[df_m['t_fin']>df_m['t_ini']]

    cv=next((c for c in df_m.columns if 'Viaje' in str(c) and '_min' not in c),None)
    ct=next((c for c in df_m.columns if str(c).strip()=='Tren'),None)
    df_m['_id']=((df_m[cv].astype(str) if cv else '')+" "+(df_m[ct].astype(str) if ct else '')).str.strip()

    fechas_m=sorted(df_m['Fecha_str'].unique())
    if not fechas_m: st.warning("Sin fechas disponibles."); st.stop()

    # ── Controles ─────────────────────────────────────────────────────────────
    cc1,cc2,cc3=st.columns([2,3,1])
    with cc1:
        fecha_m=st.selectbox("📅 Fecha",fechas_m,key="mapa_fecha")
    with cc2:
        hora_m=st.slider("🕐 Hora (minutos desde 00:00)",min_value=0,max_value=1439,
                          value=360,step=1,key="mapa_minuto")
        st.caption(f"Hora seleccionada: **{hora_m//60:02d}:{hora_m%60:02d}**")
    with cc3:
        use_rm=st.checkbox("Vel. RM",value=False,help="Usar velocidades de Restricción de Marcha")

    hora_s=f"{hora_m//60:02d}:{hora_m%60:02d}"

    cp1,cp2,cp3,cp4,_=st.columns([1,1,1,1,2])
    if cp1.button("−1 min", key="btn_m1"):  st.session_state['mapa_minuto']=max(0,hora_m-1);   st.rerun()
    if cp2.button("+1 min", key="btn_p1"):  st.session_state['mapa_minuto']=min(1439,hora_m+1); st.rerun()
    if cp3.button("−15 min",key="btn_m15"): st.session_state['mapa_minuto']=max(0,hora_m-15);  st.rerun()
    if cp4.button("+15 min",key="btn_p15"): st.session_state['mapa_minuto']=min(1439,hora_m+15);st.rerun()
    st.caption(f"Trenes a las **{hora_s}** · **{fecha_m}** · Velocidades: {'RM' if use_rm else 'Normales'}")

    # ── Calcular posiciones usando perfil real ────────────────────────────────
    df_dia=df_m[df_m['Fecha_str']==fecha_m].copy()
    df_act=df_dia[(df_dia['t_ini']<=hora_m)&(df_dia['t_fin']>hora_m)].copy()

    df_act['km_pos']=df_act.apply(
        lambda r: km_en_tiempo_real(r['t_ini'],r['t_fin'],hora_m,r['Via'],use_rm), axis=1)
    df_act['lat']=df_act['km_pos'].apply(lambda k: interpolar_posicion(k)[0])
    df_act['lon']=df_act['km_pos'].apply(lambda k: interpolar_posicion(k)[1])
    df_act['dir']=df_act['Via'].map({1:'→ Limache',2:'← Puerto'})

    # Velocidad instantánea teórica (para mostrar en tooltip)
    def vel_instantanea(km_pos_km, via, use_rm_):
        km_m=km_pos_km*1000
        segs=SPEED_PROFILE if via==1 else list(reversed(SPEED_PROFILE))
        for ki,kf,_,vn,vr in segs:
            if ki<=km_m<=kf: return vr if use_rm_ else vn
        return 0
    df_act['vel_inst']=df_act.apply(lambda r: vel_instantanea(r['km_pos'],r['Via'],use_rm),axis=1)

    df_act['tooltip']=(df_act['_id']+'<br>'+df_act['dir']+
                        '<br>km '+df_act['km_pos'].round(2).astype(str)+
                        '<br>'+df_act['vel_inst'].astype(int).astype(str)+' km/h'+
                        '<br>'+df_act['t_ini'].apply(format_hm_short)+
                        ' – '+df_act['t_fin'].apply(format_hm_short))

    # ── Figura mapa ───────────────────────────────────────────────────────────
    fig_m=go.Figure()
    fig_m.add_trace(go.Scattermapbox(lat=EST_LATS,lon=EST_LONS,mode='lines',
                                      line=dict(width=3,color='#888'),name='Línea',hoverinfo='skip'))
    fig_m.add_trace(go.Scattermapbox(lat=EST_LATS,lon=EST_LONS,mode='markers+text',
                                      marker=dict(size=7,color='#444'),text=ESTACIONES_CORTO,
                                      textposition='top right',textfont=dict(size=9,color='#333'),
                                      name='Estaciones',hovertext=ESTACIONES,
                                      hovertemplate='<b>%{hovertext}</b><br>km %{customdata:.1f}<extra></extra>',
                                      customdata=KM_ACUM))
    dv1a=df_act[df_act['Via']==1]; dv2a=df_act[df_act['Via']==2]
    if not dv1a.empty:
        fig_m.add_trace(go.Scattermapbox(lat=dv1a['lat'],lon=dv1a['lon'],mode='markers',
                                          marker=dict(size=18,color='#005195'),
                                          name='Vía 1 → Limache',hovertext=dv1a['tooltip'],
                                          hovertemplate='%{hovertext}<extra></extra>'))
    if not dv2a.empty:
        fig_m.add_trace(go.Scattermapbox(lat=dv2a['lat'],lon=dv2a['lon'],mode='markers',
                                          marker=dict(size=18,color='#E85500'),
                                          name='Vía 2 ← Puerto',hovertext=dv2a['tooltip'],
                                          hovertemplate='%{hovertext}<extra></extra>'))
    fig_m.update_layout(
        mapbox=dict(style='open-street-map',
                    center=dict(lat=float(np.mean(EST_LATS)),lon=float(np.mean(EST_LONS))),zoom=10),
        margin=dict(l=0,r=0,t=40,b=0),height=540,
        title=f"Trenes — {fecha_m} {hora_s} ({'RM' if use_rm else 'vel. normal'})",
        legend=dict(orientation='h',y=1.02,x=0))
    st.plotly_chart(fig_m,use_container_width=True)

    # ── Métricas ──────────────────────────────────────────────────────────────
    c1,c2,c3=st.columns(3)
    c1.metric("Trenes en circulación",len(df_act))
    c2.metric("Vía 1 (→ Limache)",len(dv1a))
    c3.metric("Vía 2 (← Puerto)",len(dv2a))

    if not df_act.empty:
        with st.expander("📋 Detalle de trenes activos"):
            dt2=df_act[['_id','Via','dir','km_pos','vel_inst','Unidad','t_ini','t_fin']].copy()
            dt2.columns=['Viaje/Tren','Vía','Dirección','Posición km','Vel. inst. km/h','Unidad','Salida (min)','Llegada (min)']
            dt2['Salida']=dt2['Salida (min)'].apply(format_hm_short)
            dt2['Llegada']=dt2['Llegada (min)'].apply(format_hm_short)
            dt2=dt2.drop(columns=['Salida (min)','Llegada (min)'])
            st.dataframe(dt2.style.format({'Posición km':'{:.2f}','Vel. inst. km/h':'{:.0f}'}),use_container_width=True)
    else:
        st.info("No hay trenes en circulación en esta franja horaria.")

    # ── Perfil de velocidades de la línea ─────────────────────────────────────
    with st.expander("📉 Perfil de velocidades de la línea"):
        st.plotly_chart(fig_perfil_velocidades(),use_container_width=True)

    # ── Diagrama Marey ────────────────────────────────────────────────────────
    with st.expander("📈 Diagrama espacio-tiempo (Marey)"):
        st.caption("Trayectoria real usando perfil de velocidades. Azul→Limache · Naranja→Puerto.")
        fig_mr=go.Figure()
        cvia={1:'#005195',2:'#E85500'}
        N_PUNTOS=60  # resolución de la curva

        for _,row in df_dia.iterrows():
            if pd.isna(row['t_ini']) or pd.isna(row['t_fin']): continue
            ts=np.linspace(row['t_ini'],row['t_fin'],N_PUNTOS)
            kms=[km_en_tiempo_real(row['t_ini'],row['t_fin'],t,row['Via'],use_rm) for t in ts]
            fig_mr.add_trace(go.Scatter(
                x=list(ts),y=kms,mode='lines',
                line=dict(color=cvia[row['Via']],width=1.5),showlegend=False,
                hovertemplate=(f"<b>{row['_id']}</b><br>"
                               f"Salida:{format_hm_short(row['t_ini'])}<br>"
                               f"Llegada:{format_hm_short(row['t_fin'])}<extra></extra>")))

        fig_mr.add_vline(x=hora_m,line_dash="dash",line_color="green",
                          annotation_text=hora_s,annotation_position="top right")
        # Líneas de estaciones
        for est,km_est in zip(ESTACIONES_CORTO,KM_ACUM):
            fig_mr.add_hline(y=km_est,line_width=0.5,line_color='#ccc')
        fig_mr.update_layout(
            xaxis=dict(title="Hora",tickmode='array',tickvals=list(range(0,1440,60)),
                       ticktext=[f"{h:02d}:00" for h in range(24)]),
            yaxis=dict(title="km desde Puerto",tickmode='array',tickvals=KM_ACUM,ticktext=ESTACIONES_CORTO),
            height=520,title=f"Diagrama Marey — {fecha_m} ({'RM' if use_rm else 'vel. normal'})",
            plot_bgcolor='#f8f8f8')
        st.plotly_chart(fig_mr,use_container_width=True)
