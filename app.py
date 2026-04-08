import streamlit as st
import time
import pandas as pd
import numpy as np
import re
import holidays
from io import BytesIO
from datetime import datetime, date, timedelta, time
import plotly.graph_objects as go
import plotly.express as px

# --- 0. SEGURIDAD DE COLUMNAS (Evita error de PyArrow) ---
def make_columns_unique(df):
    if not isinstance(df, pd.DataFrame) or df.empty:
        return df
    cols = pd.Series(df.columns)
    for dup in cols[cols.duplicated()].unique():
        cols[cols == dup] = [f"{dup}_{i}" if i != 0 else dup for i in range(sum(cols == dup))]
    df.columns = cols
    return df

# --- 1. CONFIGURACIÓN Y ESTILOS ---
st.set_page_config(page_title="Gestión de Energía - Dashboard SGE", layout="wide", page_icon="🚆")
chile_holidays = holidays.Chile()

st.markdown("""
    <style>
    .stMetric { background-color: #ffffff; padding: 20px; border-radius: 10px; border-left: 5px solid #005195; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    </style>
    """, unsafe_allow_html=True)

# --- 2. CONSTANTES DE RED ---
ESTACIONES = [
    'Puerto','Bellavista','Francia','Baron','Portales','Recreo','Miramar',
    'Viña del Mar','Hospital','Chorrillos','El Salto','Valencia','Quilpue',
    'El Sol','El Belloto','Las Americas','La Concepcion','Villa Alemana',
    'Sargento Aldea','Peñablanca','Limache'
]
ESTACIONES_CORTO = [
    'PU','BE','FR','BA','PO','RE','MI','VM','HO','CH',
    'ES','VAL','QU','SO','EB','AM','CO','VL','SA','PE','LI'
]
KM_TRAMO = [0.7,0.7,0.8,1.7,2.1,1.4,0.9,0.9,1.0,1.5,7.4,2.3,1.9,2.0,1.1,1.2,0.9,0.6,1.3,12.73]
KM_ACUM  = [0.0]
for _k in KM_TRAMO: KM_ACUM.append(round(KM_ACUM[-1]+_k, 2))
KM_TOTAL = KM_ACUM[-1]
N_EST    = len(ESTACIONES)

# Coordenadas reales (ancla) interpoladas para las 21 estaciones
_ANCHORS_KM  = [0.0,    8.3,     21.4,    28.5,     43.13]
_ANCHORS_LAT = [-33.0385, -33.0264, -33.0453, -33.0426, -32.9843]
_ANCHORS_LON = [-71.6271, -71.5518, -71.4445, -71.3735, -71.2777]

EST_LATS = [float(np.interp(k, _ANCHORS_KM, _ANCHORS_LAT)) for k in KM_ACUM]
EST_LONS = [float(np.interp(k, _ANCHORS_KM, _ANCHORS_LON)) for k in KM_ACUM]

def interpolar_posicion(km_pos):
    km_pos = max(0.0, min(float(km_pos), KM_TOTAL))
    return (float(np.interp(km_pos, KM_ACUM, EST_LATS)),
            float(np.interp(km_pos, KM_ACUM, EST_LONS)))

# --- 2b. FUNCIONES DE APOYO ---
def parse_latam_number(val):
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    s = str(val).strip().replace(' ', '').replace('$', '')
    s = re.sub(r'[^\d.,-]', '', s)
    if not s: return 0.0
    if ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'): s = s.replace('.', '').replace(',', '.')
        else: s = s.replace(',', '')
    elif ',' in s: s = s.replace(',', '.')
    try: return float(s)
    except: return 0.0

def get_tipo_dia(fch):
    if fch in chile_holidays or fch.weekday() == 6: return "D/F"
    if fch.weekday() == 5: return "S"
    return "L"

def format_hm_short(minutos_float):
    if pd.isna(minutos_float): return "00:00:00"
    total_seg = int(round(minutos_float * 60))
    h = total_seg // 3600
    m = (total_seg % 3600) // 60
    s = total_seg % 60
    return f"{h:02d}:{m:02d}:{s:02d}"

# --- 3. MOTOR THDR ---
def convertir_a_minutos(val):
    if pd.isna(val) or str(val).strip() == "": return None
    try:
        if isinstance(val, (datetime, time)): return val.hour * 60 + val.minute + (val.second / 60.0)
        s_val = str(val).strip()
        m_ss = re.search(r'(\d{1,2}):(\d{2}):(\d{2})', s_val)
        if m_ss: return int(m_ss.group(1)) * 60 + int(m_ss.group(2)) + (int(m_ss.group(3)) / 60.0)
        m_mm = re.search(r'(\d{1,2}):(\d{2})', s_val)
        if m_mm: return int(m_mm.group(1)) * 60 + int(m_mm.group(2))
        return None
    except: return None

def parsear_fecha_nombre(nombre_archivo):
    nombre = str(nombre_archivo)
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
        try: return date(2000 + int(s[4:]), int(s[2:4]), int(s[:2])), f"DDMMYY ({s})"
        except: pass
    return None, f"sin fecha en: '{nombre}'"

def procesar_thdr_eficiente(file, start_date, end_date):
    nombre = getattr(file, 'name', str(file))
    diag = {"archivo": nombre, "fecha_parseada": None, "en_rango": None, "filas": 0, "error": None}
    try:
        fch_date, desc = parsear_fecha_nombre(nombre)
        diag["fecha_parseada"] = desc
        if fch_date is None:
            diag["error"] = "No se encontró fecha en el nombre del archivo"; return pd.DataFrame(), diag
        diag["en_rango"] = f"{start_date} ≤ {fch_date} ≤ {end_date} → {start_date <= fch_date <= end_date}"
        if not (start_date <= fch_date <= end_date):
            diag["error"] = "Fecha fuera del rango del Sidebar"; return pd.DataFrame(), diag
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
            if any(k in str(col) for k in ['Hora Llegada', 'Hora Salida', 'Hora Salida Programada']):
                df[f"{col}_min"] = df[col].apply(convertir_a_minutos)
        if 'Unidad' in df.columns:
            df['Unidad'] = df['Unidad'].fillna('S').replace('', 'S')
        else:
            c_m2 = next((c for c in df.columns if 'Motriz 2' in str(c)), None)
            df['Unidad'] = df[c_m2].apply(lambda x: 'M' if parse_latam_number(x) > 0 else 'S') if c_m2 else 'S'
        df['Tren-Km'] = 43.13 * df['Unidad'].apply(lambda x: 2 if str(x).strip() == 'M' else 1)
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

DATA_DIRS = {"v1":"data/thdr_v1","v2":"data/thdr_v2","umr":"data/umr","seat":"data/seat","bill":"data/facturacion"}
for _d in DATA_DIRS.values(): os.makedirs(_d, exist_ok=True)

def guardar_archivo(uploaded_file, carpeta):
    with open(os.path.join(carpeta, uploaded_file.name), "wb") as out:
        out.write(uploaded_file.getbuffer())

def listar_archivos(carpeta):
    exts = ('.xls', '.xlsx', '.xlsm')
    try: return sorted([os.path.join(carpeta, f) for f in os.listdir(carpeta) if f.lower().endswith(exts)])
    except: return []

class _ArchivoEnDisco:
    def __init__(self, path):
        self.name = os.path.basename(path); self._path = path
        with open(path, 'rb') as f: self._bio = BytesIO(f.read())
    def read(self, *a, **kw):  return self._bio.read(*a, **kw)
    def seek(self, *a, **kw):  return self._bio.seek(*a, **kw)
    def tell(self, *a, **kw):  return self._bio.tell(*a, **kw)
    def seekable(self): return True
    def readable(self): return True
    def getbuffer(self): return self._bio.getvalue()
    def __str__(self): return self._path

def combinar_fuentes(uploaded_list, carpeta):
    nombres = {uf.name for uf in (uploaded_list or [])}
    return list(uploaded_list or []) + [_ArchivoEnDisco(p) for p in listar_archivos(carpeta)
                                         if os.path.basename(p) not in nombres]

# --- 5. INICIALIZACIÓN ---
df_ops = pd.DataFrame(); df_thdr_v1 = pd.DataFrame(); df_thdr_v2 = pd.DataFrame()
all_ops, all_tr, all_seat, all_fact_full, all_prmte_full = [], [], [], [], []

# --- 6. SIDEBAR ---
with st.sidebar:
    st.header("📅 Filtro Global")
    dr = st.date_input("Rango", value=(date(2026, 1, 1), date(2026, 1, 31)))
    start_date, end_date = (dr[0], dr[1]) if isinstance(dr, tuple) and len(dr) == 2 else (dr, dr)
    st.divider()
    def _badge(c): n = len(listar_archivos(c)); return f" ({n} guardados)" if n else ""
    f_v1         = st.file_uploader(f"1. THDR Vía 1{_badge(DATA_DIRS['v1'])}", accept_multiple_files=True)
    f_v2         = st.file_uploader(f"2. THDR Vía 2{_badge(DATA_DIRS['v2'])}", accept_multiple_files=True)
    f_umr        = st.file_uploader(f"3. UMR / Odómetros{_badge(DATA_DIRS['umr'])}", accept_multiple_files=True)
    f_seat_files = st.file_uploader(f"4. Energía SEAT{_badge(DATA_DIRS['seat'])}", accept_multiple_files=True)
    f_bill_files = st.file_uploader(f"5. Facturación y PRMTE{_badge(DATA_DIRS['bill'])}", accept_multiple_files=True)
    for _ul, _ca in [(f_v1,DATA_DIRS["v1"]),(f_v2,DATA_DIRS["v2"]),(f_umr,DATA_DIRS["umr"]),
                     (f_seat_files,DATA_DIRS["seat"]),(f_bill_files,DATA_DIRS["bill"])]:
        for uf in (_ul or []):
            dest = os.path.join(_ca, uf.name)
            if not os.path.exists(dest): guardar_archivo(uf, _ca)
    st.divider()
    with st.expander("🗂️ Archivos guardados"):
        _labels = {"v1":"Vía 1","v2":"Vía 2","umr":"UMR","seat":"SEAT","bill":"Facturación"}
        for _key, _carpeta in DATA_DIRS.items():
            _arch = listar_archivos(_carpeta)
            if _arch:
                st.markdown(f"**{_labels[_key]}** — {len(_arch)} archivo(s)")
                for _a in _arch:
                    _ca2, _cb2 = st.columns([5, 1]); _ca2.caption(os.path.basename(_a))
                    if _cb2.button("🗑️", key=f"del_{_a}"): os.remove(_a); st.rerun()
            else: st.caption(f"{_labels[_key]}: sin archivos")

f_v1_all   = combinar_fuentes(f_v1,         DATA_DIRS["v1"])
f_v2_all   = combinar_fuentes(f_v2,         DATA_DIRS["v2"])
f_umr_all  = combinar_fuentes(f_umr,        DATA_DIRS["umr"])
f_seat_all = combinar_fuentes(f_seat_files, DATA_DIRS["seat"])
f_bill_all = combinar_fuentes(f_bill_files, DATA_DIRS["bill"])

_CACHE_VERSION = "v4_mapa"
_cache_key = (_CACHE_VERSION, str(start_date), str(end_date),
              tuple(sorted(f.name for f in f_v1_all)), tuple(sorted(f.name for f in f_v2_all)),
              tuple(sorted(f.name for f in f_umr_all)), tuple(sorted(f.name for f in f_seat_all)),
              tuple(sorted(f.name for f in f_bill_all)))
_hay_archivos = any([f_v1_all, f_v2_all, f_umr_all, f_seat_all, f_bill_all])
_recalcular   = st.session_state.get('_cache_key') != _cache_key

if _hay_archivos and not _recalcular and 'df_ops' in st.session_state:
    df_ops = st.session_state['df_ops']; df_thdr_v1 = st.session_state['df_thdr_v1']
    df_thdr_v2 = st.session_state['df_thdr_v2']; all_tr = st.session_state['all_tr']
    all_seat = st.session_state['all_seat']; all_fact_full = st.session_state['all_fact_full']
    all_prmte_full = st.session_state['all_prmte_full']

_errores_proc = {}

if _hay_archivos and _recalcular:
    if f_umr_all:
        for f in f_umr_all:
            try:
                engine_umr = "xlrd" if f.name.lower().endswith(".xls") else "openpyxl"
                xl = pd.ExcelFile(f, engine=engine_umr)
                for sn in xl.sheet_names:
                    f.seek(0)
                    df_raw = pd.read_excel(f, sheet_name=sn, header=None, engine=engine_umr)
                    h_r = next((i for i in range(min(100, len(df_raw)))
                                if any(k in str(df_raw.iloc[i].tolist()).upper() for k in ['FECHA','ODO','KILOM'])), None)
                    if h_r is not None:
                        f.seek(0)
                        df_p = pd.read_excel(f, sheet_name=sn, header=h_r, engine=engine_umr)
                        df_p.columns = [str(c).upper().replace('Ó','O').strip() for c in df_p.columns]
                        c_f = next((c for c in df_p.columns if 'FECHA' in c), None)
                        c_o = next((c for c in df_p.columns if 'ODO' in c), None)
                        c_t = next((c for c in df_p.columns if 'KM' in c), None)
                        if c_f and c_o:
                            df_p['_dt'] = pd.to_datetime(df_p[c_f], errors='coerce').dt.normalize()
                            mask = (df_p['_dt'].dt.date >= start_date) & (df_p['_dt'].dt.date <= end_date)
                            for _, r in df_p[mask].dropna(subset=['_dt']).iterrows():
                                all_ops.append({"Fecha": r['_dt'], "Tipo Día": get_tipo_dia(r['_dt'].date()),
                                                "Odómetro [km]": parse_latam_number(r[c_o]),
                                                "Tren-Km [km]": parse_latam_number(r[c_t]) if c_t else 0.0})
                    if any(k in sn.upper() for k in ['KIL','ODO']):
                        for i in range(len(df_raw)-2):
                            for j in range(1, len(df_raw.columns)):
                                v_f = pd.to_datetime(df_raw.iloc[i, j], errors='coerce')
                                if pd.notna(v_f) and start_date <= v_f.date() <= end_date:
                                    for k in range(i+3, min(i+50, len(df_raw))):
                                        t = str(df_raw.iloc[k, 0]).strip().upper()
                                        if re.match(r'^(M|XM)', t):
                                            all_tr.append({"Tren": t, "Fecha": v_f.normalize(),
                                                           "Valor": parse_latam_number(df_raw.iloc[k, j])})
            except Exception as e: _errores_proc[f.name] = f"UMR: {e}"

    if f_seat_all:
        for f in f_seat_all:
            try:
                engine_seat = "xlrd" if f.name.lower().endswith(".xls") else "openpyxl"
                df_s = pd.read_excel(f, header=None, engine=engine_seat)
                for i in range(len(df_s)):
                    fs = pd.to_datetime(df_s.iloc[i, 1], errors='coerce')
                    if pd.notna(fs):
                        fs = fs.normalize()
                        if start_date <= fs.date() <= end_date:
                            all_seat.append({"Fecha": fs, "E_Total": parse_latam_number(df_s.iloc[i, 3]),
                                             "E_Tr": parse_latam_number(df_s.iloc[i, 5]),
                                             "E_12": parse_latam_number(df_s.iloc[i, 7])})
            except Exception as e: _errores_proc[f.name] = f"SEAT: {e}"

    if f_bill_all:
        for f in f_bill_all:
            try:
                engine_bill = "xlrd" if f.name.lower().endswith(".xls") else "openpyxl"
                f.seek(0); xl = pd.ExcelFile(f, engine=engine_bill)
                for sn in xl.sheet_names:
                    if sn == "Consumo Factura":
                        f.seek(0)
                        df_ff = pd.read_excel(f, sheet_name=sn, engine=engine_bill)
                        c_f = next((c for c in df_ff.columns if 'FECHA' in str(c).upper()), df_ff.columns[0])
                        c_v = next((c for c in df_ff.columns if 'CONSUMO' in str(c).upper() or 'VALOR' in str(c).upper()), df_ff.columns[1])
                        df_ff['dt'] = pd.to_datetime(df_ff[c_f], errors='coerce')
                        for _, r in df_ff.dropna(subset=['dt']).iterrows():
                            if "TOTAL" in str(r[c_f]).upper(): continue
                            v = abs(parse_latam_number(r[c_v]))
                            all_fact_full.append({"Fecha": r['dt'].normalize(), "Hora": f"{r['dt'].hour:02d}:00",
                                                   "15min": f"{r['dt'].hour:02d}:{(r['dt'].minute//15)*15:02d}", "Consumo": v})
                    if 'PRMTE' in sn.upper():
                        f.seek(0)
                        df_pd_raw = pd.read_excel(f, sheet_name=sn, header=None, engine=engine_bill)
                        h = next((i for i in range(min(20, len(df_pd_raw)))
                                  if any(k in str(df_pd_raw.iloc[i]).upper() for k in ['AÑO','ANO','YEAR'])), 0)
                        f.seek(0)
                        df_pd = pd.read_excel(f, sheet_name=sn, header=h, engine=engine_bill).dropna(how='all')
                        c_anio = next((c for c in df_pd.columns if str(c).upper().replace('Ñ','N').startswith('AN')), None)
                        c_mes  = next((c for c in df_pd.columns if str(c).upper().startswith('MES')), None)
                        c_dia  = next((c for c in df_pd.columns if str(c).upper().startswith('DIA')), None)
                        c_hora = next((c for c in df_pd.columns if str(c).upper() == 'HORA'), None)
                        c_ini  = next((c for c in df_pd.columns if 'INICIO' in str(c).upper()), None)
                        if not (c_anio and c_mes and c_dia and c_hora):
                            raise ValueError(f"Columnas de fecha no encontradas. Disponibles: {list(df_pd.columns)}")
                        def _build_ts(r):
                            try:
                                minuto = int(r[c_ini]) if c_ini and not pd.isna(r[c_ini]) else 0
                                return pd.Timestamp(year=int(r[c_anio]), month=int(r[c_mes]),
                                                    day=int(r[c_dia]), hour=int(r[c_hora]), minute=minuto)
                            except: return pd.NaT
                        df_pd['ts'] = df_pd.apply(_build_ts, axis=1)
                        cols_retiro = [c for c in df_pd.columns if 'Retiro_Energia_Activa' in str(c)]
                        if not cols_retiro:
                            cols_retiro = [c for c in df_pd.columns if 'RETIRO' in str(c).upper()
                                           or ('ENERGIA' in str(c).upper() and 'ACTIVA' in str(c).upper())]
                        for _, r in df_pd.dropna(subset=['ts']).iterrows():
                            ts = r['ts']
                            if pd.isna(ts) or not (start_date <= ts.date() <= end_date): continue
                            consumo = sum(parse_latam_number(r.get(c, 0)) for c in cols_retiro)
                            all_prmte_full.append({"Fecha": ts.normalize(), "Hora": f"{ts.hour:02d}:00",
                                                    "15min": f"{ts.hour:02d}:{ts.minute:02d}", "Consumo": consumo})
            except Exception as e: _errores_proc[f.name] = f"Factura/PRMTE: {e}"

    if _errores_proc: st.session_state['_errores_proc'] = _errores_proc

    if all_ops:
        df_ops = pd.DataFrame(all_ops).groupby("Fecha").agg(
            {"Odómetro [km]": "sum", "Tren-Km [km]": "sum", "Tipo Día": "first"}).reset_index()
        df_f_d = (pd.DataFrame(all_fact_full).groupby("Fecha")["Consumo"].sum().reset_index()
                  .rename(columns={"Consumo": "E_Fact"}) if all_fact_full else pd.DataFrame(columns=["Fecha","E_Fact"]))
        df_p_d = (pd.DataFrame(all_prmte_full).groupby("Fecha")["Consumo"].sum().reset_index()
                  .rename(columns={"Consumo": "E_Prmte"}) if all_prmte_full else pd.DataFrame(columns=["Fecha","E_Prmte"]))
        df_s_d = (pd.DataFrame(all_seat).groupby("Fecha").agg({"E_Total":"sum","E_Tr":"sum","E_12":"sum"}).reset_index()
                  .rename(columns={"E_Total":"E_Seat_T","E_Tr":"E_Seat_Tr","E_12":"E_Seat_12"}) if all_seat
                  else pd.DataFrame(columns=["Fecha","E_Seat_T","E_Seat_Tr","E_Seat_12"]))
        for dff in [df_ops, df_f_d, df_p_d, df_s_d]: dff['Fecha'] = pd.to_datetime(dff['Fecha']).dt.normalize()
        df_ops = (df_ops.merge(df_f_d, on="Fecha", how="left").merge(df_p_d, on="Fecha", how="left")
                        .merge(df_s_d, on="Fecha", how="left").fillna(0))
        def jerarquia_energia(row):
            if row['E_Fact'] > 0:     tot, src = row['E_Fact'], "Factura"
            elif row['E_Prmte'] > 0:  tot, src = row['E_Prmte'], "PRMTE"
            elif row['E_Seat_T'] > 0: tot, src = row['E_Seat_T'], "SEAT"
            else: return 0, 0, 0, 0, 0, "N/A"
            r_tr = row['E_Seat_Tr'] / row['E_Seat_T'] if row['E_Seat_T'] > 0 else 0.85
            r_12 = row['E_Seat_12'] / row['E_Seat_T'] if row['E_Seat_T'] > 0 else 0.15
            return tot, tot*r_tr, tot*r_12, r_tr*100, r_12*100, src
        df_ops[['E_Total','E_Tr','E_12','% Tracción','% 12 kV','Fuente']] = df_ops.apply(
            jerarquia_energia, axis=1, result_type='expand')
        df_ops['IDE (kWh/km)'] = df_ops.apply(
            lambda r: r['E_Tr'] / r['Odómetro [km]'] if r['Odómetro [km]'] > 0 else 0, axis=1)

    diagnosticos_thdr = []
    if f_v1_all:
        res_v1 = [procesar_thdr_eficiente(f, start_date, end_date) for f in f_v1_all]
        diagnosticos_thdr += [r[1] for r in res_v1]
        partes_v1 = [r[0] for r in res_v1 if not r[0].empty]
        df_thdr_v1 = pd.concat(partes_v1, ignore_index=True) if partes_v1 else pd.DataFrame()
    if f_v2_all:
        res_v2 = [procesar_thdr_eficiente(f, start_date, end_date) for f in f_v2_all]
        diagnosticos_thdr += [r[1] for r in res_v2]
        partes_v2 = [r[0] for r in res_v2 if not r[0].empty]
        df_thdr_v2 = pd.concat(partes_v2, ignore_index=True) if partes_v2 else pd.DataFrame()
    if diagnosticos_thdr: st.session_state['diag_thdr'] = diagnosticos_thdr

    st.session_state.update({'df_ops': df_ops, 'df_thdr_v1': df_thdr_v1, 'df_thdr_v2': df_thdr_v2,
                              'all_tr': all_tr, 'all_seat': all_seat, 'all_fact_full': all_fact_full,
                              'all_prmte_full': all_prmte_full, '_cache_key': _cache_key})

# --- 7. TABS ---
tabs = st.tabs(["📊 Resumen","📑 Operaciones","📑 Trenes","⚡ Energía","⚖️ Comparación hr",
                "📈 Regresión","🚨 Atípicos","📋 THDR","🔬 Servicios vs Energía","🗺️ Mapa de Trenes"])

# TAB 0: RESUMEN
with tabs[0]:
    _ep = st.session_state.get('_errores_proc', {})
    if _ep:
        with st.expander(f"⚠️ {len(_ep)} archivo(s) con error", expanded=True):
            for _n, _m in _ep.items(): st.error(f"**{_n}**: {_m}")
    if not df_ops.empty:
        c1, c2, c3 = st.columns(3)
        c1.metric("Odómetro Total", f"{df_ops['Odómetro [km]'].sum():,.1f} km")
        c2.metric("Tren-Km Total",  f"{df_ops['Tren-Km [km]'].sum():,.1f} km")
        c3.metric("IDE Promedio",   f"{df_ops['IDE (kWh/km)'].mean():.4f} kWh/km")
        st.plotly_chart(go.Figure(go.Bar(x=df_ops['Fecha'], y=df_ops['Odómetro [km]'],
                                          marker_color="#005195")).update_layout(
            title="Odómetro Diario", xaxis_title="Fecha", yaxis_title="km"), use_container_width=True)
        st.plotly_chart(go.Figure(go.Scatter(x=df_ops['Fecha'], y=df_ops['IDE (kWh/km)'],
                                              mode='lines+markers', line=dict(color="#E85500"))).update_layout(
            title="IDE Diario (kWh/km)", xaxis_title="Fecha", yaxis_title="kWh/km"), use_container_width=True)
    else:
        st.info("📂 Sube archivos desde el panel lateral para ver el resumen.")

# TAB 1: OPERACIONES
with tabs[1]:
    if not df_ops.empty:
        df_view = df_ops.copy(); df_view['Fecha'] = df_view['Fecha'].dt.strftime('%Y-%m-%d')
        st.write("### 📑 Detalle Operacional e IDE")
        st.dataframe(make_columns_unique(df_view).style.format({
            'Odómetro [km]': "{:,.1f}", 'Tren-Km [km]': "{:,.1f}", 'E_Total': "{:,.0f}",
            'E_Tr': "{:,.0f}", 'E_12': "{:,.0f}", '% Tracción': "{:.1f}%",
            '% 12 kV': "{:.1f}%", 'IDE (kWh/km)': "{:.4f}"}), use_container_width=True)
    else: st.info("📂 Sin datos de operación disponibles.")

# TAB 2: TRENES
with tabs[2]:
    if all_tr:
        df_tr = pd.DataFrame(all_tr); df_tr['Fecha'] = df_tr['Fecha'].dt.strftime('%Y-%m-%d')
        st.write("### 📑 Kilómetros por Tren")
        pivot = df_tr.pivot_table(index='Tren', columns='Fecha', values='Valor', aggfunc='sum').fillna(0)
        st.dataframe(pivot.style.format("{:,.1f}"), use_container_width=True)
    else: st.info("📂 Sin datos de trenes disponibles.")

# TAB 3: ENERGÍA
with tabs[3]:
    e_tabs = st.tabs(["🔹 SEAT","🔹 PRMTE","🔹 Facturación"])
    with e_tabs[0]:
        if all_seat:
            df_sv = pd.DataFrame(all_seat); df_sv['Fecha'] = df_sv['Fecha'].dt.strftime('%Y-%m-%d')
            st.dataframe(df_sv.style.format({'E_Total':"{:,.0f}",'E_Tr':"{:,.0f}",'E_12':"{:,.0f}"}), use_container_width=True)
        else: st.info("📂 Sin datos SEAT.")
    with e_tabs[1]:
        if all_prmte_full:
            df_p = pd.DataFrame(all_prmte_full); df_p['Fecha_Str'] = df_p['Fecha'].dt.strftime('%Y-%m-%d')
            pico_row = df_p.loc[df_p['Consumo'].idxmax()]
            m1, m2, m3 = st.columns(3)
            m1.metric("Total kWh", f"{df_p['Consumo'].sum():,.0f}")
            m2.metric("Días cargados", f"{df_p['Fecha_Str'].nunique()}")
            m3.metric("Pico 15 min", f"{pico_row['Consumo']:,.0f} kWh", f"{pico_row['Fecha_Str']} {pico_row['15min']}")
            fechas_p = sorted(df_p['Fecha_Str'].unique())
            col_f, col_v = st.columns([2,1])
            fp = col_f.selectbox("Fecha", fechas_p, key="prmte_fecha")
            vp = col_v.radio("Vista", ["15 min","Horario","Diario"], horizontal=True, key="prmte_vista")
            df_ps = df_p[df_p['Fecha_Str'] == fp]
            if vp == "15 min":   df_sh = df_ps.groupby("15min")["Consumo"].sum().reset_index().rename(columns={"15min":"Franja","Consumo":"kWh"}).sort_values("Franja")
            elif vp == "Horario":df_sh = df_ps.groupby("Hora")["Consumo"].sum().reset_index().rename(columns={"Hora":"Franja","Consumo":"kWh"}).sort_values("Franja")
            else:                df_sh = df_p.groupby("Fecha_Str")["Consumo"].sum().reset_index().rename(columns={"Fecha_Str":"Franja","Consumo":"kWh"})
            fig_p = go.Figure(go.Bar(x=df_sh['Franja'], y=df_sh['kWh'], marker_color='#005195',
                                      hovertemplate='<b>%{x}</b><br>%{y:,.0f} kWh<extra></extra>'))
            fig_p.update_layout(title=f"PRMTE — {fp} ({vp})" if vp != "Diario" else "PRMTE — Consumo diario",
                                xaxis_title="Franja", yaxis_title="kWh", xaxis=dict(tickangle=-45), height=380)
            st.plotly_chart(fig_p, use_container_width=True)
            with st.expander("📋 Ver tabla"):
                st.dataframe(df_sh.style.format({'kWh':"{:,.1f}"}), use_container_width=True, height=300)
        else: st.info("📂 Sin datos PRMTE.")
    with e_tabs[2]:
        if all_fact_full:
            df_f = pd.DataFrame(all_fact_full); df_f['Fecha_Str'] = df_f['Fecha'].dt.strftime('%Y-%m-%d')
            st.dataframe(df_f.groupby("Fecha_Str")["Consumo"].sum().reset_index()
                         .style.format({'Consumo':"{:,.0f}"}), use_container_width=True)
            st.dataframe(df_f[['Fecha_Str','15min','Consumo']].style.format({'Consumo':"{:,.0f}"}), use_container_width=True)
        else: st.info("📂 Sin datos de facturación.")

# TAB 4: COMPARACIÓN HORARIA
with tabs[4]:
    st.header("⚖️ Comparación Horaria")
    if all_prmte_full or all_fact_full:
        fuentes = {}
        if all_prmte_full:
            df_ph = pd.DataFrame(all_prmte_full); df_ph['Hora_int'] = df_ph['Hora'].str[:2].astype(int)
            fuentes['PRMTE'] = df_ph.groupby('Hora_int')['Consumo'].sum().reset_index()
        if all_fact_full:
            df_fh = pd.DataFrame(all_fact_full); df_fh['Hora_int'] = df_fh['Hora'].str[:2].astype(int)
            fuentes['Factura'] = df_fh.groupby('Hora_int')['Consumo'].sum().reset_index()
        fig = go.Figure()
        colors = {'PRMTE':'#005195','Factura':'#E85500'}
        for nb, dfh in fuentes.items():
            fig.add_trace(go.Bar(x=dfh['Hora_int'], y=dfh['Consumo'], name=nb, marker_color=colors.get(nb,'#888')))
        fig.update_layout(title="Consumo Acumulado por Hora", barmode='group', xaxis_title="Hora", yaxis_title="kWh")
        st.plotly_chart(fig, use_container_width=True)
    else: st.info("📂 Sube datos de PRMTE o Facturación para comparar.")

# TAB 5: REGRESIÓN
with tabs[5]:
    st.header("📈 Regresión IDE vs Odómetro")
    if not df_ops.empty and df_ops['IDE (kWh/km)'].sum() > 0:
        df_reg = df_ops[df_ops['IDE (kWh/km)'] > 0].copy()
        color_map = {"L":"#005195","S":"#FFA500","D/F":"#E85500"}
        fig = go.Figure()
        for tipo, grp in df_reg.groupby('Tipo Día'):
            fig.add_trace(go.Scatter(x=grp['Odómetro [km]'], y=grp['IDE (kWh/km)'], mode='markers', name=tipo,
                                     marker=dict(color=color_map.get(tipo,'#888'), size=8)))
        xa, ya = df_reg['Odómetro [km]'].values, df_reg['IDE (kWh/km)'].values
        if len(xa) >= 2:
            coef = np.polyfit(xa, ya, 1); xl = np.linspace(xa.min(), xa.max(), 100)
            r2 = np.corrcoef(xa, ya)[0,1]**2
            fig.add_trace(go.Scatter(x=xl, y=np.polyval(coef, xl), mode='lines',
                                     name=f'Tendencia (R²={r2:.3f})', line=dict(color='gray', dash='dash', width=2)))
        fig.update_layout(title='IDE vs Odómetro por Tipo de Día', xaxis_title='Odómetro [km]', yaxis_title='kWh/km')
        st.plotly_chart(fig, use_container_width=True)
    else: st.info("📂 Sin datos suficientes para regresión.")

# TAB 6: ATÍPICOS
with tabs[6]:
    st.header("🚨 Detección de Atípicos")
    if not df_ops.empty and df_ops['IDE (kWh/km)'].sum() > 0:
        df_at = df_ops[df_ops['IDE (kWh/km)'] > 0].copy()
        media, std = df_at['IDE (kWh/km)'].mean(), df_at['IDE (kWh/km)'].std()
        umbral = st.slider("Umbral σ", 1.0, 3.0, 2.0, 0.1)
        df_at['Atípico'] = (df_at['IDE (kWh/km)'] - media).abs() > umbral * std
        c1, c2 = st.columns(2); c1.metric("Media IDE", f"{media:.4f}"); c2.metric("Atípicos", int(df_at['Atípico'].sum()))
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=df_at[~df_at['Atípico']]['Fecha'], y=df_at[~df_at['Atípico']]['IDE (kWh/km)'],
                                  mode='markers', name='Normal', marker=dict(color='#005195')))
        fig.add_trace(go.Scatter(x=df_at[df_at['Atípico']]['Fecha'], y=df_at[df_at['Atípico']]['IDE (kWh/km)'],
                                  mode='markers', name='Atípico', marker=dict(color='red', size=10, symbol='x')))
        fig.add_hline(y=media+umbral*std, line_dash="dash", line_color="orange", annotation_text=f"+{umbral}σ")
        fig.add_hline(y=media-umbral*std, line_dash="dash", line_color="orange", annotation_text=f"-{umbral}σ")
        fig.update_layout(title="IDE Diario con Atípicos", xaxis_title="Fecha", yaxis_title="kWh/km")
        st.plotly_chart(fig, use_container_width=True)
        if df_at['Atípico'].any():
            df_sh = df_at[df_at['Atípico']][['Fecha','Tipo Día','Odómetro [km]','IDE (kWh/km)','Fuente']].copy()
            df_sh['Fecha'] = df_sh['Fecha'].dt.strftime('%Y-%m-%d')
            st.dataframe(df_sh.style.format({'Odómetro [km]':"{:,.1f}",'IDE (kWh/km)':"{:.4f}"}))
    else: st.info("📂 Sin datos suficientes para detección de atípicos.")

# TAB 7: THDR
def render_via_thdr(df_via, label):
    if df_via.empty:
        st.info(f"📂 No hay datos cargados para {label}."); return
    df = df_via.copy(); df['Fecha'] = df['Fecha_Op'].dt.strftime('%Y-%m-%d')
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total viajes", len(df)); c2.metric("Tren-Km", f"{df['Tren-Km'].sum():,.1f}")
    c3.metric("Días cargados", df['Fecha'].nunique())
    c4.metric("Viajes doble (M)", (df['Unidad'].astype(str).str.strip() == 'M').sum())
    resumen = df.groupby('Fecha').agg(Viajes=('Unidad','count'),
                                       Viajes_M=('Unidad', lambda x: (x.astype(str).str.strip()=='M').sum()),
                                       TrenKm=('Tren-Km','sum')).reset_index()
    st.dataframe(resumen.style.format({'TrenKm':"{:,.1f}"}), use_container_width=True)
    fechas_disp = sorted(df['Fecha'].unique())
    fecha_sel = st.selectbox(f"Seleccionar fecha ({label})", fechas_disp, key=f"sel_{label}")
    df_sel = df[df['Fecha'] == fecha_sel].copy()
    for col_min in [c for c in df_sel.columns if '_min' in c]:
        df_sel[col_min.replace('_min','_hms')] = df_sel[col_min].apply(format_hm_short)
    cols_base = ['Viaje','Tren','Hora Salida Programada','Motriz 1','Motriz 2','Unidad','Maquinista','Tren-Km']
    cols_hms  = [c for c in df_sel.columns if '_hms' in c]
    cols_show = [c for c in cols_base + cols_hms if c in df_sel.columns]
    st.dataframe(make_columns_unique(df_sel[cols_show]).reset_index(drop=True), use_container_width=True)
    st.caption(f"{len(df_sel)} viajes el {fecha_sel}")

with tabs[7]:
    st.header("📋 Análisis THDR")
    t_v1, t_v2 = st.tabs(["🔵 Vía 1 (Puerto → Limache)","🟠 Vía 2 (Limache → Puerto)"])
    with t_v1: render_via_thdr(df_thdr_v1, "Vía 1")
    with t_v2: render_via_thdr(df_thdr_v2, "Vía 2")

# TAB 8: SERVICIOS VS ENERGÍA
with tabs[8]:
    st.header("🔬 Servicios vs Consumo de Energía (15 min)")
    _tiene_prmte = len(all_prmte_full) > 0
    _tiene_thdr  = not df_thdr_v1.empty or not df_thdr_v2.empty
    if not _tiene_prmte and not _tiene_thdr:
        st.info("📂 Sube archivos PRMTE y THDR para este análisis."); st.stop()
    col_av, col_at = st.columns(2)
    col_av.metric("PRMTE disponible", "✅" if _tiene_prmte else "❌ Sin datos")
    col_at.metric("THDR disponible",  "✅" if _tiene_thdr  else "❌ Sin datos")

    def str_franja_a_minutos(s):
        try: h, m = map(int, s.split(':')); return h*60+m
        except: return 0

    df_servicios = pd.DataFrame()
    if _tiene_thdr:
        partes = [df for df in [df_thdr_v1, df_thdr_v2] if not df.empty]
        df_all_thdr = pd.concat(partes, ignore_index=True)
        df_all_thdr['Fecha_str'] = df_all_thdr['Fecha_Op'].dt.strftime('%Y-%m-%d')
        def _pri_sal(row):
            vals = [row[c] for c in row.index if 'Salida' in c and '_min' in c and pd.notna(row[c])]
            return min(vals) if vals else np.nan
        def _ult_lle(row):
            vals = [row[c] for c in row.index if 'Llegada' in c and '_min' in c and pd.notna(row[c])]
            return max(vals) if vals else np.nan
        df_all_thdr['t_ini'] = df_all_thdr.apply(_pri_sal, axis=1)
        df_all_thdr['t_fin'] = df_all_thdr.apply(_ult_lle, axis=1)
        df_all_thdr = df_all_thdr.dropna(subset=['t_ini','t_fin'])
        def _km_franja(t_ini, t_fin, t_f, unidad):
            dur = t_fin - t_ini
            if dur <= 0: return 0.0
            dist = KM_TOTAL * (2 if str(unidad).strip() == 'M' else 1)
            mins_act = max(0.0, min(t_fin, t_f+15) - max(t_ini, t_f))
            return round((dist/dur) * mins_act, 3)
        todas_franjas = [f"{h:02d}:{m:02d}" for h in range(24) for m in range(0,60,15)]
        filas_srv = []
        for fg, grp in df_all_thdr.groupby('Fecha_str'):
            for fr in todas_franjas:
                t_f = int(fr[:2])*60 + int(fr[3:])
                mask = (grp['t_ini'] <= t_f) & (grp['t_fin'] > t_f)
                if mask.sum() == 0: continue
                ga = grp[mask]
                filas_srv.append({'Fecha': fg, 'Franja': fr, 'Servicios': int(mask.sum()),
                                   'Servicios_M': int((ga['Unidad'].astype(str).str.strip()=='M').sum()),
                                   'Tren_Km': sum(_km_franja(r['t_ini'], r['t_fin'], t_f, r['Unidad']) for _, r in ga.iterrows())})
        if filas_srv: df_servicios = pd.DataFrame(filas_srv)

    df_energia = pd.DataFrame()
    if _tiene_prmte:
        df_prmte = pd.DataFrame(all_prmte_full); df_prmte['Fecha'] = df_prmte['Fecha'].dt.strftime('%Y-%m-%d')
        df_energia = (df_prmte.groupby(['Fecha','15min'])['Consumo'].sum().reset_index()
                      .rename(columns={'15min':'Franja','Consumo':'kWh'}))

    if df_servicios.empty and df_energia.empty:
        st.warning("Sin datos suficientes."); st.stop()
    if not df_servicios.empty and not df_energia.empty:
        df_merge = pd.merge(df_energia, df_servicios, on=['Fecha','Franja'], how='outer').fillna(0)
    elif not df_energia.empty:
        df_merge = df_energia.copy(); df_merge['Servicios'] = 0; df_merge['Servicios_M'] = 0; df_merge['Tren_Km'] = 0.0
    else:
        df_merge = df_servicios.copy(); df_merge['kWh'] = 0
    if 'Tren_Km' not in df_merge.columns: df_merge['Tren_Km'] = 0.0
    df_merge['_ord'] = df_merge['Franja'].apply(str_franja_a_minutos)
    df_merge = df_merge.sort_values(['Fecha','_ord']).drop(columns='_ord')

    fechas_disp = sorted(df_merge['Fecha'].unique())
    if not fechas_disp: st.warning("Sin fechas en el rango."); st.stop()

    st.divider()
    col_f1, col_f2, col_f3 = st.columns([2,2,1])
    with col_f1: modo = st.radio("Vista", ["Por día","Promedio del período"], horizontal=True)
    with col_f2:
        if modo == "Por día":
            fecha_sel = st.selectbox("Fecha", fechas_disp)
            df_plot = df_merge[df_merge['Fecha'] == fecha_sel].copy()
        else:
            df_plot = df_merge.groupby('Franja').agg(kWh=('kWh','mean'), Servicios=('Servicios','mean'),
                                                      Servicios_M=('Servicios_M','mean'), Tren_Km=('Tren_Km','mean')).reset_index()
            df_plot['_ord'] = df_plot['Franja'].apply(str_franja_a_minutos)
            df_plot = df_plot.sort_values('_ord').drop(columns='_ord')
    with col_f3: mostrar_m = st.checkbox("Solo tracción doble (M)", value=False)

    col_srv = 'Servicios_M' if mostrar_m else 'Servicios'
    lbl_srv = 'Servicios tracción doble' if mostrar_m else 'Servicios totales'

    if not df_plot.empty:
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Total kWh", f"{df_plot['kWh'].sum():,.0f}")
        m2.metric("Pico kWh", f"{df_plot['kWh'].max():,.0f}", df_plot.loc[df_plot['kWh'].idxmax(),'Franja'])
        m3.metric("Total servicios", f"{df_plot[col_srv].sum():.0f}")
        m4.metric("Pico servicios", f"{df_plot[col_srv].max():.0f}",
                  df_plot.loc[df_plot[col_srv].idxmax(),'Franja'] if df_plot[col_srv].max() > 0 else "—")
        m5.metric("Tren-Km período", f"{df_plot['Tren_Km'].sum():,.1f} km")

    st.divider()
    if not df_plot.empty:
        fig_dual = go.Figure()
        fig_dual.add_trace(go.Bar(x=df_plot['Franja'], y=df_plot['kWh'], name='Energía PRMTE (kWh)',
                                   marker_color='rgba(0,81,149,0.7)', yaxis='y1'))
        fig_dual.add_trace(go.Scatter(x=df_plot['Franja'], y=df_plot[col_srv], name=lbl_srv,
                                       mode='lines+markers', line=dict(color='#E85500', width=2),
                                       marker=dict(size=5), yaxis='y2'))
        if df_plot['Tren_Km'].sum() > 0:
            fig_dual.add_trace(go.Scatter(x=df_plot['Franja'], y=df_plot['Tren_Km'], name='Tren-Km recorridos',
                                           mode='lines', line=dict(color='#00AA44', width=2, dash='dot'), yaxis='y3'))
        titulo = (f"Energía vs Servicios — {fecha_sel}" if modo == "Por día"
                  else f"Energía vs Servicios — Promedio {fechas_disp[0]} a {fechas_disp[-1]}")
        fig_dual.update_layout(title=titulo,
            xaxis=dict(title="Franja 15 min", tickangle=-45, tickmode='array', tickvals=df_plot['Franja'][::4].tolist()),
            yaxis=dict(title="kWh", side='left', showgrid=True),
            yaxis2=dict(title="Servicios", side='right', overlaying='y', showgrid=False),
            yaxis3=dict(title="Tren-Km", side='right', overlaying='y', showgrid=False,
                        anchor='free', position=1.0, showticklabels=False),
            legend=dict(orientation='h', y=1.08), hovermode='x unified', height=450)
        st.plotly_chart(fig_dual, use_container_width=True)

    df_corr = df_plot.dropna(subset=['kWh',col_srv])
    df_corr = df_corr[(df_corr['kWh']>0) & (df_corr[col_srv]>0)]
    if len(df_corr) >= 5:
        st.divider()
        corr = np.corrcoef(df_corr['kWh'].values, df_corr[col_srv].values)[0,1]
        st.subheader(f"📐 Correlación energía ↔ servicios: **{corr:.3f}**")
        interp = ("muy alta 🟢" if abs(corr)>0.8 else "alta 🟡" if abs(corr)>0.6
                  else "moderada 🟠" if abs(corr)>0.4 else "baja 🔴")
        st.caption(f"Correlación {interp}. {'Positiva: más servicios → más consumo.' if corr>0 else 'Negativa: relación inversa.'}")
        coef = np.polyfit(df_corr[col_srv].values, df_corr['kWh'].values, 1)
        xl = np.linspace(df_corr[col_srv].min(), df_corr[col_srv].max(), 100)
        fig_sc = go.Figure()
        fig_sc.add_trace(go.Scatter(x=df_corr[col_srv], y=df_corr['kWh'], mode='markers', text=df_corr['Franja'],
                                     hovertemplate='<b>%{text}</b><br>Servicios: %{x}<br>kWh: %{y:,.0f}<extra></extra>',
                                     marker=dict(color='#005195', size=7, opacity=0.7), name='Franjas'))
        fig_sc.add_trace(go.Scatter(x=xl, y=np.polyval(coef, xl), mode='lines',
                                     line=dict(color='#E85500', dash='dash'), name=f'Tendencia (R²={corr**2:.3f})'))
        fig_sc.update_layout(title='Dispersión: Servicios vs kWh por franja',
                              xaxis_title=lbl_srv, yaxis_title='kWh', height=380)
        st.plotly_chart(fig_sc, use_container_width=True)

    with st.expander("📋 Ver tabla"):
        cols_show = [c for c in ['Franja','kWh','Servicios','Servicios_M','Tren_Km'] if c in df_plot.columns]
        st.dataframe(df_plot[cols_show].style.format(
            {'kWh':'{:,.1f}','Servicios':'{:.1f}','Servicios_M':'{:.1f}','Tren_Km':'{:.2f}'}),
            use_container_width=True, height=300)

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 9: MAPA DE TRENES
# ═══════════════════════════════════════════════════════════════════════════════
with tabs[9]:
    st.header("🗺️ Mapa de Trenes — Posición cada 15 min")

    _tiene_thdr_mapa = not df_thdr_v1.empty or not df_thdr_v2.empty
    if not _tiene_thdr_mapa:
        st.info("📂 Sube archivos THDR (Vía 1 y/o Vía 2) para ver el mapa."); st.stop()

    # ── Preparar DataFrame unificado ─────────────────────────────────────────
    partes_mapa = []
    if not df_thdr_v1.empty:
        tmp = df_thdr_v1.copy(); tmp['Via'] = 1; partes_mapa.append(tmp)
    if not df_thdr_v2.empty:
        tmp = df_thdr_v2.copy(); tmp['Via'] = 2; partes_mapa.append(tmp)
    df_mapa = pd.concat(partes_mapa, ignore_index=True)
    df_mapa['Fecha_str'] = df_mapa['Fecha_Op'].dt.strftime('%Y-%m-%d')

    def _pri_s(row):
        vals = [row[c] for c in row.index if 'Salida' in c and '_min' in c and pd.notna(row[c])]
        return min(vals) if vals else np.nan
    def _ult_l(row):
        vals = [row[c] for c in row.index if 'Llegada' in c and '_min' in c and pd.notna(row[c])]
        return max(vals) if vals else np.nan

    df_mapa['t_ini'] = df_mapa.apply(_pri_s, axis=1)
    df_mapa['t_fin'] = df_mapa.apply(_ult_l, axis=1)
    df_mapa = df_mapa.dropna(subset=['t_ini','t_fin'])
    df_mapa = df_mapa[df_mapa['t_fin'] > df_mapa['t_ini']]

    col_viaje = next((c for c in df_mapa.columns if 'Viaje' in str(c) and '_min' not in c), None)
    col_tren  = next((c for c in df_mapa.columns if str(c).strip() == 'Tren'), None)
    df_mapa['_id'] = ((df_mapa[col_viaje].astype(str) if col_viaje else '') + ' ' +
                       (df_mapa[col_tren].astype(str)  if col_tren  else '')).str.strip()

    fechas_mapa = sorted(df_mapa['Fecha_str'].unique())
    if not fechas_mapa:
        st.warning("Sin fechas disponibles en los datos THDR."); st.stop()

    # ── Controles ─────────────────────────────────────────────────────────────
    col_c1, col_c2 = st.columns([2, 3])
    with col_c1:
        fecha_mapa = st.selectbox("📅 Fecha", fechas_mapa, key="mapa_fecha")
    with col_c2:
        franjas_dia = [f"{h:02d}:{m:02d}" for h in range(24) for m in range(0,60,15)]
        idx_fr = st.select_slider("🕐 Franja horaria (15 min)", options=list(range(len(franjas_dia))),
                                   value=24, format_func=lambda i: franjas_dia[i], key="mapa_franja")

    hora_str = franjas_dia[idx_fr]
    hora_min = int(hora_str[:2])*60 + int(hora_str[3:])

    # Botones prev/next
    cp1, cp2, _ = st.columns([1,1,4])
    if cp1.button("⏮ −15 min", key="btn_prev"):
        st.session_state['mapa_franja'] = max(0, idx_fr-1); st.rerun()
    if cp2.button("⏭ +15 min", key="btn_next"):
        st.session_state['mapa_franja'] = min(len(franjas_dia)-1, idx_fr+1); st.rerun()

    st.caption(f"Trenes en circulación a las **{hora_str}** · **{fecha_mapa}**")

    # ── Calcular posiciones ───────────────────────────────────────────────────
    df_dia     = df_mapa[df_mapa['Fecha_str'] == fecha_mapa].copy()
    df_activos = df_dia[(df_dia['t_ini'] <= hora_min) & (df_dia['t_fin'] > hora_min)].copy()

    def km_en_t(row, t):
        dur = row['t_fin'] - row['t_ini']
        if dur <= 0: return 0.0
        frac = max(0.0, min(1.0, (t - row['t_ini']) / dur))
        return frac * KM_TOTAL if row['Via'] == 1 else (1-frac) * KM_TOTAL

    df_activos['km_pos'] = df_activos.apply(lambda r: km_en_t(r, hora_min), axis=1)
    df_activos['lat']    = df_activos['km_pos'].apply(lambda k: interpolar_posicion(k)[0])
    df_activos['lon']    = df_activos['km_pos'].apply(lambda k: interpolar_posicion(k)[1])
    df_activos['dir']    = df_activos['Via'].map({1:'→ Limache', 2:'← Puerto'})
    df_activos['tooltip'] = (df_activos['_id'] + '<br>' + df_activos['dir'] +
                              '<br>km ' + df_activos['km_pos'].round(1).astype(str) +
                              '<br>' + df_activos['t_ini'].apply(format_hm_short) +
                              ' – ' + df_activos['t_fin'].apply(format_hm_short))

    # ── Figura mapa ───────────────────────────────────────────────────────────
    fig_mapa = go.Figure()

    # Línea de la ruta
    fig_mapa.add_trace(go.Scattermapbox(
        lat=EST_LATS, lon=EST_LONS, mode='lines',
        line=dict(width=3, color='#888888'), name='Línea', hoverinfo='skip'))

    # Estaciones
    fig_mapa.add_trace(go.Scattermapbox(
        lat=EST_LATS, lon=EST_LONS, mode='markers+text',
        marker=dict(size=7, color='#444444'),
        text=ESTACIONES_CORTO, textposition='top right',
        textfont=dict(size=9, color='#333333'),
        name='Estaciones', hovertext=ESTACIONES,
        hovertemplate='<b>%{hovertext}</b><br>km %{customdata:.1f}<extra></extra>',
        customdata=KM_ACUM))

    # Trenes Vía 1 (azul)
    df_v1a = df_activos[df_activos['Via'] == 1]
    if not df_v1a.empty:
        fig_mapa.add_trace(go.Scattermapbox(
            lat=df_v1a['lat'], lon=df_v1a['lon'], mode='markers',
            marker=dict(size=18, color='#005195'),
            name='Vía 1 → Limache', hovertext=df_v1a['tooltip'],
            hovertemplate='%{hovertext}<extra></extra>'))

    # Trenes Vía 2 (naranja)
    df_v2a = df_activos[df_activos['Via'] == 2]
    if not df_v2a.empty:
        fig_mapa.add_trace(go.Scattermapbox(
            lat=df_v2a['lat'], lon=df_v2a['lon'], mode='markers',
            marker=dict(size=18, color='#E85500'),
            name='Vía 2 ← Puerto', hovertext=df_v2a['tooltip'],
            hovertemplate='%{hovertext}<extra></extra>'))

    fig_mapa.update_layout(
        mapbox=dict(style='open-street-map',
                    center=dict(lat=float(np.mean(EST_LATS)), lon=float(np.mean(EST_LONS))), zoom=10),
        margin=dict(l=0, r=0, t=40, b=0), height=540,
        title=f"Trenes en circulación — {fecha_mapa} {hora_str}",
        legend=dict(orientation='h', y=1.02, x=0))
    st.plotly_chart(fig_mapa, use_container_width=True)

    # ── Métricas ──────────────────────────────────────────────────────────────
    c1, c2, c3 = st.columns(3)
    c1.metric("Trenes en circulación", len(df_activos))
    c2.metric("Vía 1 (→ Limache)",    len(df_v1a))
    c3.metric("Vía 2 (← Puerto)",     len(df_v2a))

    if not df_activos.empty:
        with st.expander("📋 Detalle de trenes activos"):
            df_tab = df_activos[['_id','Via','dir','km_pos','Unidad','t_ini','t_fin']].copy()
            df_tab.columns = ['Viaje/Tren','Vía','Dirección','Posición km','Unidad','Salida (min)','Llegada (min)']
            df_tab['Salida']  = df_tab['Salida (min)'].apply(format_hm_short)
            df_tab['Llegada'] = df_tab['Llegada (min)'].apply(format_hm_short)
            df_tab = df_tab.drop(columns=['Salida (min)','Llegada (min)'])
            st.dataframe(df_tab.style.format({'Posición km':'{:.1f}'}), use_container_width=True)
    else:
        st.info("No hay trenes en circulación en esta franja horaria.")

    # ── Diagrama Marey (espacio-tiempo) ───────────────────────────────────────
    with st.expander("📈 Diagrama espacio-tiempo (Marey)"):
        st.caption("Cada línea = un viaje. Azul → Limache · Naranja → Puerto. Línea verde = franja actual.")
        fig_marey = go.Figure()
        color_via = {1:'#005195', 2:'#E85500'}
        for _, row in df_dia.iterrows():
            if pd.isna(row['t_ini']) or pd.isna(row['t_fin']): continue
            km_ini = 0.0 if row['Via'] == 1 else KM_TOTAL
            km_fin = KM_TOTAL if row['Via'] == 1 else 0.0
            fig_marey.add_trace(go.Scatter(
                x=[row['t_ini'], row['t_fin']], y=[km_ini, km_fin],
                mode='lines', line=dict(color=color_via[row['Via']], width=1.5),
                showlegend=False,
                hovertemplate=(f"<b>{row['_id']}</b><br>"
                               f"Salida: {format_hm_short(row['t_ini'])}<br>"
                               f"Llegada: {format_hm_short(row['t_fin'])}<extra></extra>")))
        fig_marey.add_vline(x=hora_min, line_dash="dash", line_color="green",
                             annotation_text=hora_str, annotation_position="top right")
        fig_marey.update_layout(
            xaxis=dict(title="Hora", tickmode='array',
                       tickvals=list(range(0,1440,60)),
                       ticktext=[f"{h:02d}:00" for h in range(24)]),
            yaxis=dict(title="km desde Puerto", tickmode='array',
                       tickvals=KM_ACUM, ticktext=ESTACIONES_CORTO),
            height=500, title=f"Diagrama Marey — {fecha_mapa}", plot_bgcolor='#f8f8f8')
        st.plotly_chart(fig_marey, use_container_width=True)
