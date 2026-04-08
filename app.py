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

# --- 2. CONSTANTES DE RED (estaciones, distancias) ---
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
# Distancias en km entre estaciones consecutivas (Puerto→Limache)
KM_TRAMO = [0.7,0.7,0.8,1.7,2.1,1.4,0.9,0.9,1.0,1.5,7.4,2.3,1.9,2.0,1.1,1.2,0.9,0.6,1.3,12.73]
KM_ACUM  = [0.0]
for _k in KM_TRAMO: KM_ACUM.append(round(KM_ACUM[-1]+_k, 2))
KM_TOTAL = KM_ACUM[-1]  # 43.13 km total
N_EST    = len(ESTACIONES)

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
    if pd.isna(minutos_float): return "00:00"
    h, m = divmod(int(minutos_float), 60)
    return f"{h:02d}:{m:02d}"

# --- 3. MOTOR THDR (A1: FECHA | 2 CABECERAS | 3 FILAS VACÍAS | DATA) ---
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
    """
    Extrae la fecha desde el nombre del archivo.
    Soporta: DD-MM-YYYY, DD_MM_YYYY, YYYY-MM-DD, YYYY_MM_DD, DDMMYYYY, DDMMYY
    Ej: Escenario_6_21-12-2025_XT32.xlsx → 2025-12-21
    """
    nombre = str(nombre_archivo)

    # DD-MM-YYYY o DD_MM_YYYY
    m = re.search(r'(\d{2})[-_](\d{2})[-_](\d{4})', nombre)
    if m:
        try:
            return date(int(m.group(3)), int(m.group(2)), int(m.group(1))), f"DD-MM-YYYY desde nombre ({m.group()})"
        except: pass

    # YYYY-MM-DD o YYYY_MM_DD
    m = re.search(r'(\d{4})[-_](\d{2})[-_](\d{2})', nombre)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3))), f"YYYY-MM-DD desde nombre ({m.group()})"
        except: pass

    # DDMMYYYY (8 dígitos seguidos)
    m = re.search(r'(\d{8})', nombre)
    if m:
        s = m.group(1)
        try:
            return date(int(s[4:]), int(s[2:4]), int(s[:2])), f"DDMMYYYY desde nombre ({s})"
        except: pass

    # DDMMYY (6 dígitos seguidos)
    m = re.search(r'(\d{6})', nombre)
    if m:
        s = m.group(1)
        try:
            return date(2000 + int(s[4:]), int(s[2:4]), int(s[:2])), f"DDMMYY desde nombre ({s})"
        except: pass

    return None, f"sin fecha reconocible en: '{nombre}'"


def procesar_thdr_eficiente(file, start_date, end_date):
    nombre = getattr(file, 'name', str(file))
    diag = {"archivo": nombre, "fecha_parseada": None, "en_rango": None, "filas": 0, "error": None}
    try:
        fch_date, desc = parsear_fecha_nombre(nombre)
        diag["fecha_parseada"] = desc

        if fch_date is None:
            diag["error"] = "No se encontró fecha en el nombre del archivo"
            return pd.DataFrame(), diag

        diag["en_rango"] = f"{start_date} ≤ {fch_date} ≤ {end_date} → {start_date <= fch_date <= end_date}"
        if not (start_date <= fch_date <= end_date):
            diag["error"] = "Fecha fuera del rango del Sidebar"
            return pd.DataFrame(), diag

        fch_dt = pd.to_datetime(fch_date).normalize()

        # Leer con xlrd para .xls, openpyxl para .xlsx
        engine = "xlrd" if nombre.lower().endswith(".xls") else "openpyxl"
        df_raw = pd.read_excel(file, header=None, engine=engine)

        # --- Construcción de cabeceras ---
        # Fila 0: fecha en col 0, estaciones desde col 11 (con merged cells = ffill)
        # Fila 1: nombres de columnas base (Viaje, Tren, ..., Hora Llegada, Hora Salida, ...)
        r0 = df_raw.iloc[0].copy()
        r0[0] = np.nan          # quitar la fecha del ffill
        h1 = r0.ffill().astype(str)
        h2 = df_raw.iloc[1].fillna('').astype(str)

        cols = []
        for stn, tip in zip(h1, h2):
            stn, tip = str(stn).strip(), str(tip).strip()
            if stn == 'nan' or stn == '':
                cols.append(tip if tip else '_vacio')
            else:
                cols.append(f"{stn}_{tip}" if tip else stn)

        # --- Datos desde fila 5 (índice 5, saltando las 3 filas vacías) ---
        df = df_raw.iloc[5:].copy().reset_index(drop=True)
        n = len(df.columns)
        cols_adj = cols[:n] if len(cols) >= n else cols + [f"_C{j}" for j in range(n - len(cols))]
        df.columns = cols_adj
        df = make_columns_unique(df).dropna(how='all', axis=0).reset_index(drop=True)

        # Convertir columnas de hora a minutos
        for col in df.columns:
            if any(k in str(col) for k in ['Hora Llegada', 'Hora Salida', 'Hora Salida Programada']):
                df[f"{col}_min"] = df[col].apply(convertir_a_minutos)

        # Unidad: leer directo de la columna (ya tiene 'M' o vacío)
        if 'Unidad' in df.columns:
            df['Unidad'] = df['Unidad'].fillna('S').replace('', 'S')
        else:
            # Fallback: derivar de Motriz 2
            c_m2 = next((c for c in df.columns if 'Motriz 2' in str(c)), None)
            df['Unidad'] = df[c_m2].apply(lambda x: 'M' if parse_latam_number(x) > 0 else 'S') if c_m2 else 'S'

        df['Tren-Km'] = 43.13 * df['Unidad'].apply(lambda x: 2 if str(x).strip() == 'M' else 1)
        df['Fecha_Op'] = fch_dt

        # Hora de referencia (salida desde Puerto o Limache)
        col_ref = next((c for c in df.columns
                        if ('PUERTO' in str(c).upper() or 'LIMACHE' in str(c).upper())
                        and 'Salida' in str(c) and '_min' in str(c)), None)
        if col_ref:
            df['Hora_Ref_Min'] = df[col_ref]

        diag["filas"] = len(df)
        return df, diag

    except Exception as e:
        diag["error"] = str(e)
        return pd.DataFrame(), diag

# --- 4. PERSISTENCIA EN DISCO ---
import os

DATA_DIRS = {
    "v1":   "data/thdr_v1",
    "v2":   "data/thdr_v2",
    "umr":  "data/umr",
    "seat": "data/seat",
    "bill": "data/facturacion",
}
for _d in DATA_DIRS.values():
    os.makedirs(_d, exist_ok=True)

def guardar_archivo(uploaded_file, carpeta):
    dest = os.path.join(carpeta, uploaded_file.name)
    with open(dest, "wb") as out:
        out.write(uploaded_file.getbuffer())

def listar_archivos(carpeta):
    exts = ('.xls', '.xlsx', '.xlsm')
    try:
        return sorted([os.path.join(carpeta, f) for f in os.listdir(carpeta) if f.lower().endswith(exts)])
    except:
        return []

class _ArchivoEnDisco:
    """Wrapper de archivo en disco compatible con pd.read_excel y getattr(f, 'name')."""
    def __init__(self, path):
        from io import BytesIO
        self.name = os.path.basename(path)
        self._path = path
        with open(path, 'rb') as f:
            self._bio = BytesIO(f.read())
    def read(self, *a, **kw):     return self._bio.read(*a, **kw)
    def seek(self, *a, **kw):     return self._bio.seek(*a, **kw)
    def tell(self, *a, **kw):     return self._bio.tell(*a, **kw)
    def seekable(self):           return True
    def readable(self):           return True
    def getbuffer(self):          return self._bio.getvalue()
    def __str__(self):            return self._path

def combinar_fuentes(uploaded_list, carpeta):
    nombres_subidos = {uf.name for uf in (uploaded_list or [])}
    desde_disco = [_ArchivoEnDisco(p) for p in listar_archivos(carpeta)
                   if os.path.basename(p) not in nombres_subidos]
    return list(uploaded_list or []) + desde_disco

# --- 5. INICIALIZACIÓN ---
df_ops = pd.DataFrame()
df_thdr_v1 = pd.DataFrame()
df_thdr_v2 = pd.DataFrame()
all_ops, all_tr, all_seat, all_fact_full, all_prmte_full = [], [], [], [], []

# --- 6. SIDEBAR ---
with st.sidebar:
    st.header("📅 Filtro Global")
    dr = st.date_input("Rango", value=(date(2026, 1, 1), date(2026, 1, 31)))
    start_date, end_date = (dr[0], dr[1]) if isinstance(dr, tuple) and len(dr) == 2 else (dr, dr)
    st.divider()

    def _badge(carpeta):
        n = len(listar_archivos(carpeta))
        return f" ({n} guardados)" if n else ""

    f_v1         = st.file_uploader(f"1. THDR Vía 1{_badge(DATA_DIRS['v1'])}", accept_multiple_files=True)
    f_v2         = st.file_uploader(f"2. THDR Vía 2{_badge(DATA_DIRS['v2'])}", accept_multiple_files=True)
    f_umr        = st.file_uploader(f"3. UMR / Odómetros{_badge(DATA_DIRS['umr'])}", accept_multiple_files=True)
    f_seat_files = st.file_uploader(f"4. Energía SEAT{_badge(DATA_DIRS['seat'])}", accept_multiple_files=True)
    f_bill_files = st.file_uploader(f"5. Facturación y PRMTE{_badge(DATA_DIRS['bill'])}", accept_multiple_files=True)

    # Guardar al disco archivos recién subidos
    for _uploaded_list, _carpeta in [
        (f_v1, DATA_DIRS["v1"]), (f_v2, DATA_DIRS["v2"]),
        (f_umr, DATA_DIRS["umr"]), (f_seat_files, DATA_DIRS["seat"]),
        (f_bill_files, DATA_DIRS["bill"]),
    ]:
        for uf in (_uploaded_list or []):
            dest = os.path.join(_carpeta, uf.name)
            if not os.path.exists(dest):
                guardar_archivo(uf, _carpeta)

    st.divider()
    with st.expander("🗂️ Archivos guardados"):
        _labels = {"v1":"Vía 1","v2":"Vía 2","umr":"UMR","seat":"SEAT","bill":"Facturación"}
        for _key, _carpeta in DATA_DIRS.items():
            _archivos = listar_archivos(_carpeta)
            if _archivos:
                st.markdown(f"**{_labels[_key]}** — {len(_archivos)} archivo(s)")
                for _a in _archivos:
                    _ca, _cb = st.columns([5, 1])
                    _ca.caption(os.path.basename(_a))
                    if _cb.button("🗑️", key=f"del_{_a}"):
                        os.remove(_a)
                        st.rerun()
            else:
                st.caption(f"{_labels[_key]}: sin archivos")

# Combinar subidos ahora + guardados en disco
f_v1_all        = combinar_fuentes(f_v1,         DATA_DIRS["v1"])
f_v2_all        = combinar_fuentes(f_v2,         DATA_DIRS["v2"])
f_umr_all       = combinar_fuentes(f_umr,        DATA_DIRS["umr"])
f_seat_all      = combinar_fuentes(f_seat_files, DATA_DIRS["seat"])
f_bill_all      = combinar_fuentes(f_bill_files, DATA_DIRS["bill"])

# Clave de caché: tuple con rango + nombres de archivos
_CACHE_VERSION = "v3_15min"  # cambiar esto fuerza recálculo en todos los usuarios
_cache_key = (
    _CACHE_VERSION,
    str(start_date), str(end_date),
    tuple(sorted(f.name for f in f_v1_all)),
    tuple(sorted(f.name for f in f_v2_all)),
    tuple(sorted(f.name for f in f_umr_all)),
    tuple(sorted(f.name for f in f_seat_all)),
    tuple(sorted(f.name for f in f_bill_all)),
)
_hay_archivos = any([f_v1_all, f_v2_all, f_umr_all, f_seat_all, f_bill_all])
_recalcular   = st.session_state.get('_cache_key') != _cache_key

# Recuperar caché si existe y no cambió nada
if _hay_archivos and not _recalcular and 'df_ops' in st.session_state:
    df_ops         = st.session_state['df_ops']
    df_thdr_v1     = st.session_state['df_thdr_v1']
    df_thdr_v2     = st.session_state['df_thdr_v2']
    all_tr         = st.session_state['all_tr']
    all_seat       = st.session_state['all_seat']
    all_fact_full  = st.session_state['all_fact_full']
    all_prmte_full = st.session_state['all_prmte_full']

_errores_proc = {}  # {nombre_archivo: mensaje_error}

if _hay_archivos and _recalcular:
    # UMR / TRENES
    if f_umr_all:
        for f in f_umr_all:
            try:
                engine_umr = "xlrd" if f.name.lower().endswith(".xls") else "openpyxl"
                xl = pd.ExcelFile(f, engine=engine_umr)
                for sn in xl.sheet_names:
                    f.seek(0)
                    df_raw = pd.read_excel(f, sheet_name=sn, header=None, engine=engine_umr)
                    h_r = next((i for i in range(min(100, len(df_raw))) if any(k in str(df_raw.iloc[i].tolist()).upper() for k in ['FECHA', 'ODO', 'KILOM'])), None)
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
                            df_filt = df_p[mask].dropna(subset=['_dt'])
                            for _, r in df_filt.iterrows():
                                all_ops.append({
                                    "Fecha": r['_dt'],
                                    "Tipo Día": get_tipo_dia(r['_dt'].date()),
                                    "Odómetro [km]": parse_latam_number(r[c_o]),
                                    "Tren-Km [km]": parse_latam_number(r[c_t]) if c_t else 0.0
                                })
                    if any(k in sn.upper() for k in ['KIL', 'ODO']):
                        for i in range(len(df_raw)-2):
                            for j in range(1, len(df_raw.columns)):
                                v_f = pd.to_datetime(df_raw.iloc[i, j], errors='coerce')
                                if pd.notna(v_f) and start_date <= v_f.date() <= end_date:
                                    for k in range(i+3, min(i+50, len(df_raw))):
                                        t = str(df_raw.iloc[k, 0]).strip().upper()
                                        if re.match(r'^(M|XM)', t):
                                            all_tr.append({
                                                "Tren": t,
                                                "Fecha": v_f.normalize(),
                                                "Valor": parse_latam_number(df_raw.iloc[k, j])
                                            })
            except Exception as e:
                _errores_proc[f.name] = f"UMR: {e}"

    # SEAT
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
                            all_seat.append({
                                "Fecha": fs,
                                "E_Total": parse_latam_number(df_s.iloc[i, 3]),
                                "E_Tr":    parse_latam_number(df_s.iloc[i, 5]),
                                "E_12":    parse_latam_number(df_s.iloc[i, 7])
                            })
            except Exception as e:
                _errores_proc[f.name] = f"SEAT: {e}"

    # FACTURA / PRMTE
    if f_bill_all:
        for f in f_bill_all:
            try:
                engine_bill = "xlrd" if f.name.lower().endswith(".xls") else "openpyxl"
                f.seek(0)
                xl = pd.ExcelFile(f, engine=engine_bill)
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
                            all_fact_full.append({
                                "Fecha": r['dt'].normalize(),
                                "Hora":  f"{r['dt'].hour:02d}:00",
                                "15min": f"{r['dt'].hour:02d}:{(r['dt'].minute//15)*15:02d}",
                                "Consumo": v
                            })
                    if 'PRMTE' in sn.upper():
                        f.seek(0)
                        df_pd_raw = pd.read_excel(f, sheet_name=sn, header=None, engine=engine_bill)
                        # Buscar fila cabecera (contiene AÑO o ANO)
                        h = next((i for i in range(min(20, len(df_pd_raw)))
                                  if any(k in str(df_pd_raw.iloc[i]).upper()
                                         for k in ['AÑO','ANO','YEAR'])), 0)
                        f.seek(0)
                        df_pd = pd.read_excel(f, sheet_name=sn, header=h, engine=engine_bill)
                        df_pd = df_pd.dropna(how='all')

                        # Columnas de timestamp
                        c_anio  = next((c for c in df_pd.columns if str(c).upper().replace('Ñ','N').startswith('AN')), None)
                        c_mes   = next((c for c in df_pd.columns if str(c).upper().startswith('MES')), None)
                        c_dia   = next((c for c in df_pd.columns if str(c).upper().startswith('DIA')), None)
                        c_hora  = next((c for c in df_pd.columns if str(c).upper() == 'HORA'), None)
                        c_ini   = next((c for c in df_pd.columns if 'INICIO' in str(c).upper()), None)

                        if not (c_anio and c_mes and c_dia and c_hora):
                            raise ValueError(f"Columnas de fecha no encontradas. Disponibles: {list(df_pd.columns)}")

                        # Construir timestamp con minutos del intervalo
                        def _build_ts(r):
                            try:
                                minuto = int(r[c_ini]) if c_ini and not pd.isna(r[c_ini]) else 0
                                return pd.Timestamp(year=int(r[c_anio]), month=int(r[c_mes]),
                                                    day=int(r[c_dia]),   hour=int(r[c_hora]),
                                                    minute=minuto)
                            except:
                                return pd.NaT

                        df_pd['ts'] = df_pd.apply(_build_ts, axis=1)

                        # Sumar todas las columnas Retiro_Energia_Activa (H1 + H2 + ...)
                        cols_retiro = [c for c in df_pd.columns if 'Retiro_Energia_Activa' in str(c)]
                        if not cols_retiro:
                            cols_retiro = [c for c in df_pd.columns if 'RETIRO' in str(c).upper()
                                           or ('ENERGIA' in str(c).upper() and 'ACTIVA' in str(c).upper())]

                        for _, r in df_pd.dropna(subset=['ts']).iterrows():
                            ts = r['ts']
                            if pd.isna(ts) or not (start_date <= ts.date() <= end_date):
                                continue
                            consumo = sum(parse_latam_number(r.get(c, 0)) for c in cols_retiro)
                            all_prmte_full.append({
                                "Fecha": ts.normalize(),
                                "Hora":  f"{ts.hour:02d}:00",
                                "15min": f"{ts.hour:02d}:{ts.minute:02d}",
                                "Consumo": consumo
                            })
            except Exception as e:
                _errores_proc[f.name] = f"Factura/PRMTE: {e}"

    if _errores_proc:
        st.session_state['_errores_proc'] = _errores_proc

    # CONSOLIDACIÓN
    if all_ops:
        df_ops = pd.DataFrame(all_ops).groupby("Fecha").agg({
            "Odómetro [km]": "sum",
            "Tren-Km [km]": "sum",
            "Tipo Día": "first"
        }).reset_index()

        df_f_d = (pd.DataFrame(all_fact_full).groupby("Fecha")["Consumo"].sum().reset_index()
                  .rename(columns={"Consumo": "E_Fact"})
                  if all_fact_full else pd.DataFrame(columns=["Fecha", "E_Fact"]))
        df_p_d = (pd.DataFrame(all_prmte_full).groupby("Fecha")["Consumo"].sum().reset_index()
                  .rename(columns={"Consumo": "E_Prmte"})
                  if all_prmte_full else pd.DataFrame(columns=["Fecha", "E_Prmte"]))
        df_s_d = (pd.DataFrame(all_seat).groupby("Fecha").agg({"E_Total":"sum","E_Tr":"sum","E_12":"sum"}).reset_index()
                  .rename(columns={"E_Total":"E_Seat_T","E_Tr":"E_Seat_Tr","E_12":"E_Seat_12"})
                  if all_seat else pd.DataFrame(columns=["Fecha","E_Seat_T","E_Seat_Tr","E_Seat_12"]))

        for dff in [df_ops, df_f_d, df_p_d, df_s_d]:
            dff['Fecha'] = pd.to_datetime(dff['Fecha']).dt.normalize()

        df_ops = (df_ops
                  .merge(df_f_d, on="Fecha", how="left")
                  .merge(df_p_d, on="Fecha", how="left")
                  .merge(df_s_d, on="Fecha", how="left")
                  .fillna(0))

        def jerarquia_energia(row):
            if row['E_Fact'] > 0:
                tot, src = row['E_Fact'], "Factura"
            elif row['E_Prmte'] > 0:
                tot, src = row['E_Prmte'], "PRMTE"
            elif row['E_Seat_T'] > 0:
                tot, src = row['E_Seat_T'], "SEAT"
            else:
                return 0, 0, 0, 0, 0, "N/A"
            r_tr = row['E_Seat_Tr'] / row['E_Seat_T'] if row['E_Seat_T'] > 0 else 0.85
            r_12 = row['E_Seat_12'] / row['E_Seat_T'] if row['E_Seat_T'] > 0 else 0.15
            return tot, tot * r_tr, tot * r_12, r_tr * 100, r_12 * 100, src

        df_ops[['E_Total','E_Tr','E_12','% Tracción','% 12 kV','Fuente']] = df_ops.apply(
            jerarquia_energia, axis=1, result_type='expand'
        )
        df_ops['IDE (kWh/km)'] = df_ops.apply(
            lambda r: r['E_Tr'] / r['Odómetro [km]'] if r['Odómetro [km]'] > 0 else 0, axis=1
        )

    diagnosticos_thdr = []
    if f_v1_all:
        resultados_v1 = [procesar_thdr_eficiente(f, start_date, end_date) for f in f_v1_all]
        diagnosticos_thdr += [r[1] for r in resultados_v1]
        partes_v1 = [r[0] for r in resultados_v1 if not r[0].empty]
        df_thdr_v1 = pd.concat(partes_v1, ignore_index=True) if partes_v1 else pd.DataFrame()
    if f_v2_all:
        resultados_v2 = [procesar_thdr_eficiente(f, start_date, end_date) for f in f_v2_all]
        diagnosticos_thdr += [r[1] for r in resultados_v2]
        partes_v2 = [r[0] for r in resultados_v2 if not r[0].empty]
        df_thdr_v2 = pd.concat(partes_v2, ignore_index=True) if partes_v2 else pd.DataFrame()
    if diagnosticos_thdr:
        st.session_state['diag_thdr'] = diagnosticos_thdr

    # Guardar resultados en session_state y marcar caché
    st.session_state['df_ops']     = df_ops
    st.session_state['df_thdr_v1'] = df_thdr_v1
    st.session_state['df_thdr_v2'] = df_thdr_v2
    st.session_state['all_tr']        = all_tr
    st.session_state['all_seat']      = all_seat
    st.session_state['all_fact_full'] = all_fact_full
    st.session_state['all_prmte_full']= all_prmte_full
    # Guardar en session_state
    st.session_state['df_ops']        = df_ops
    st.session_state['df_thdr_v1']    = df_thdr_v1
    st.session_state['df_thdr_v2']    = df_thdr_v2
    st.session_state['all_tr']        = all_tr
    st.session_state['all_seat']      = all_seat
    st.session_state['all_fact_full'] = all_fact_full
    st.session_state['all_prmte_full']= all_prmte_full
    st.session_state['_cache_key']    = _cache_key

# --- 7. TABS ---
tabs = st.tabs([
    "📊 Resumen",
    "📑 Operaciones",
    "📑 Trenes",
    "⚡ Energía",
    "⚖️ Comparación hr",
    "📈 Regresión",
    "🚨 Atípicos",
    "📋 THDR",
    "🔬 Servicios vs Energía"
])

# TAB 0: RESUMEN
with tabs[0]:
    # Mostrar errores de procesamiento si los hay
    _ep = st.session_state.get('_errores_proc', {})
    if _ep:
        with st.expander(f"⚠️ {len(_ep)} archivo(s) con error al procesar", expanded=True):
            for _nombre, _msg in _ep.items():
                st.error(f"**{_nombre}**: {_msg}")

    if not df_ops.empty:
        df_rf = df_ops.copy()
        c1, c2, c3 = st.columns(3)
        c1.metric("Odómetro Total", f"{df_rf['Odómetro [km]'].sum():,.1f} km")
        c2.metric("Tren-Km Total", f"{df_rf['Tren-Km [km]'].sum():,.1f} km")
        c3.metric("IDE Promedio", f"{df_rf['IDE (kWh/km)'].mean():.4f} kWh/km")
        st.plotly_chart(
            go.Figure(data=[go.Bar(
                x=df_rf['Fecha'],
                y=df_rf['Odómetro [km]'],
                marker_color="#005195",
                name="Odómetro [km]"
            )]).update_layout(title="Odómetro Diario", xaxis_title="Fecha", yaxis_title="km"),
            use_container_width=True
        )
        st.plotly_chart(
            go.Figure(data=[go.Scatter(
                x=df_rf['Fecha'],
                y=df_rf['IDE (kWh/km)'],
                mode='lines+markers',
                line=dict(color="#E85500"),
                name="IDE"
            )]).update_layout(title="IDE Diario (kWh/km)", xaxis_title="Fecha", yaxis_title="kWh/km"),
            use_container_width=True
        )
    else:
        st.info("📂 Sube archivos desde el panel lateral para ver el resumen.")

# TAB 1: OPERACIONES
with tabs[1]:
    if not df_ops.empty:
        df_view = df_ops.copy()
        df_view['Fecha'] = df_view['Fecha'].dt.strftime('%Y-%m-%d')
        st.write("### 📑 Detalle Operacional e IDE")
        st.dataframe(
            make_columns_unique(df_view).style.format({
                'Odómetro [km]': "{:,.1f}",
                'Tren-Km [km]': "{:,.1f}",
                'E_Total': "{:,.0f}",
                'E_Tr': "{:,.0f}",
                'E_12': "{:,.0f}",
                '% Tracción': "{:.1f}%",
                '% 12 kV': "{:.1f}%",
                'IDE (kWh/km)': "{:.4f}"
            }),
            use_container_width=True
        )
    else:
        st.info("📂 Sin datos de operación disponibles.")

# TAB 2: TRENES
with tabs[2]:
    if all_tr:
        df_tr = pd.DataFrame(all_tr)
        df_tr['Fecha'] = df_tr['Fecha'].dt.strftime('%Y-%m-%d')
        st.write("### 📑 Kilómetros por Tren")
        pivot = df_tr.pivot_table(index='Tren', columns='Fecha', values='Valor', aggfunc='sum').fillna(0)
        st.dataframe(pivot.style.format("{:,.1f}"), use_container_width=True)
    else:
        st.info("📂 Sin datos de trenes disponibles.")

# TAB 3: ENERGÍA
with tabs[3]:
    e_tabs = st.tabs(["🔹 SEAT", "🔹 PRMTE", "🔹 Facturación"])
    with e_tabs[0]:
        if all_seat:
            df_s_view = pd.DataFrame(all_seat)
            df_s_view['Fecha'] = df_s_view['Fecha'].dt.strftime('%Y-%m-%d')
            st.write("#### 📅 Datos SEAT Diarios")
            st.dataframe(df_s_view.style.format({
                'E_Total': "{:,.0f}",
                'E_Tr': "{:,.0f}",
                'E_12': "{:,.0f}"
            }), use_container_width=True)
        else:
            st.info("📂 Sin datos SEAT.")
    with e_tabs[1]:
        if all_prmte_full:
            df_p = pd.DataFrame(all_prmte_full)
            df_p['Fecha_Str'] = df_p['Fecha'].dt.strftime('%Y-%m-%d')

            # Métricas rápidas
            total_kwh = df_p['Consumo'].sum()
            dias = df_p['Fecha_Str'].nunique()
            pico_row = df_p.loc[df_p['Consumo'].idxmax()]
            m1, m2, m3 = st.columns(3)
            m1.metric("Total kWh", f"{total_kwh:,.0f}")
            m2.metric("Días cargados", f"{dias}")
            m3.metric("Pico 15 min", f"{pico_row['Consumo']:,.0f} kWh",
                      f"{pico_row['Fecha_Str']} {pico_row['15min']}")

            # Selector de fecha
            fechas_p = sorted(df_p['Fecha_Str'].unique())
            col_f, col_v = st.columns([2, 1])
            fecha_p_sel = col_f.selectbox("Fecha", fechas_p, key="prmte_fecha")
            vista_p = col_v.radio("Vista", ["15 min", "Horario", "Diario"], horizontal=True, key="prmte_vista")

            df_p_sel = df_p[df_p['Fecha_Str'] == fecha_p_sel]

            if vista_p == "15 min":
                df_show = (df_p_sel.groupby("15min")["Consumo"].sum()
                           .reset_index().rename(columns={"15min": "Franja", "Consumo": "kWh"}))
                df_show = df_show.sort_values("Franja")
            elif vista_p == "Horario":
                df_show = (df_p_sel.groupby("Hora")["Consumo"].sum()
                           .reset_index().rename(columns={"Hora": "Franja", "Consumo": "kWh"}))
                df_show = df_show.sort_values("Franja")
            else:
                df_show = (df_p.groupby("Fecha_Str")["Consumo"].sum()
                           .reset_index().rename(columns={"Fecha_Str": "Franja", "Consumo": "kWh"}))

            # Gráfico de barras
            fig_p = go.Figure(go.Bar(
                x=df_show['Franja'], y=df_show['kWh'],
                marker_color='#005195',
                hovertemplate='<b>%{x}</b><br>%{y:,.0f} kWh<extra></extra>'
            ))
            fig_p.update_layout(
                title=f"PRMTE — {fecha_p_sel} ({vista_p})" if vista_p != "Diario" else "PRMTE — Consumo diario",
                xaxis_title="Franja", yaxis_title="kWh",
                xaxis=dict(tickangle=-45),
                height=380
            )
            st.plotly_chart(fig_p, use_container_width=True)

            # Tabla
            with st.expander("📋 Ver tabla"):
                st.dataframe(
                    df_show.style.format({'kWh': "{:,.1f}"}),
                    use_container_width=True, height=300
                )
        else:
            st.info("📂 Sin datos PRMTE.")
    with e_tabs[2]:
        if all_fact_full:
            df_f = pd.DataFrame(all_fact_full)
            df_f['Fecha_Str'] = df_f['Fecha'].dt.strftime('%Y-%m-%d')
            st.write("#### 📅 Factura Diario")
            st.dataframe(
                df_f.groupby("Fecha_Str")["Consumo"].sum().reset_index()
                .style.format({'Consumo': "{:,.0f}"}),
                use_container_width=True
            )
            st.write("#### ⏲️ Factura cada 15 min")
            st.dataframe(df_f[['Fecha_Str', '15min', 'Consumo']]
                         .style.format({'Consumo': "{:,.0f}"}),
                         use_container_width=True)
        else:
            st.info("📂 Sin datos de facturación.")

# TAB 4: COMPARACIÓN HORARIA
with tabs[4]:
    st.header("⚖️ Comparación Horaria")
    if all_prmte_full or all_fact_full:
        fuentes = {}
        if all_prmte_full:
            df_ph = pd.DataFrame(all_prmte_full)
            df_ph['Hora_int'] = df_ph['Hora'].str[:2].astype(int)
            fuentes['PRMTE'] = df_ph.groupby('Hora_int')['Consumo'].sum().reset_index()
        if all_fact_full:
            df_fh = pd.DataFrame(all_fact_full)
            df_fh['Hora_int'] = df_fh['Hora'].str[:2].astype(int)
            fuentes['Factura'] = df_fh.groupby('Hora_int')['Consumo'].sum().reset_index()
        fig = go.Figure()
        colors = {'PRMTE': '#005195', 'Factura': '#E85500'}
        for nombre, df_h in fuentes.items():
            fig.add_trace(go.Bar(
                x=df_h['Hora_int'], y=df_h['Consumo'],
                name=nombre, marker_color=colors.get(nombre, '#888')
            ))
        fig.update_layout(
            title="Consumo Acumulado por Hora", barmode='group',
            xaxis_title="Hora", yaxis_title="kWh"
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("📂 Sube datos de PRMTE o Facturación para comparar.")

# TAB 5: REGRESIÓN
with tabs[5]:
    st.header("📈 Regresión IDE vs Odómetro")
    if not df_ops.empty and df_ops['IDE (kWh/km)'].sum() > 0:
        df_reg = df_ops[df_ops['IDE (kWh/km)'] > 0].copy()
        color_map = {"L": "#005195", "S": "#FFA500", "D/F": "#E85500"}
        fig = go.Figure()
        for tipo, grp in df_reg.groupby('Tipo Día'):
            fig.add_trace(go.Scatter(
                x=grp['Odómetro [km]'], y=grp['IDE (kWh/km)'],
                mode='markers', name=tipo,
                marker=dict(color=color_map.get(tipo, '#888'), size=8)
            ))
        # Línea de regresión global con numpy
        x_all = df_reg['Odómetro [km]'].values
        y_all = df_reg['IDE (kWh/km)'].values
        if len(x_all) >= 2:
            coef = np.polyfit(x_all, y_all, 1)
            x_line = np.linspace(x_all.min(), x_all.max(), 100)
            y_line = np.polyval(coef, x_line)
            r2 = np.corrcoef(x_all, y_all)[0, 1] ** 2
            fig.add_trace(go.Scatter(
                x=x_line, y=y_line, mode='lines',
                name=f'Tendencia (R²={r2:.3f})',
                line=dict(color='gray', dash='dash', width=2)
            ))
        fig.update_layout(title='IDE vs Odómetro por Tipo de Día',
                          xaxis_title='Odómetro [km]', yaxis_title='kWh/km')
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("📂 Sin datos suficientes para regresión.")

# TAB 6: ATÍPICOS
with tabs[6]:
    st.header("🚨 Detección de Atípicos")
    if not df_ops.empty and df_ops['IDE (kWh/km)'].sum() > 0:
        df_at = df_ops[df_ops['IDE (kWh/km)'] > 0].copy()
        media = df_at['IDE (kWh/km)'].mean()
        std = df_at['IDE (kWh/km)'].std()
        umbral = st.slider("Umbral σ", 1.0, 3.0, 2.0, 0.1)
        df_at['Atípico'] = (df_at['IDE (kWh/km)'] - media).abs() > umbral * std
        col1, col2 = st.columns(2)
        col1.metric("Media IDE", f"{media:.4f}")
        col2.metric("Atípicos detectados", int(df_at['Atípico'].sum()))
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=df_at[~df_at['Atípico']]['Fecha'],
            y=df_at[~df_at['Atípico']]['IDE (kWh/km)'],
            mode='markers', name='Normal', marker=dict(color='#005195')
        ))
        fig.add_trace(go.Scatter(
            x=df_at[df_at['Atípico']]['Fecha'],
            y=df_at[df_at['Atípico']]['IDE (kWh/km)'],
            mode='markers', name='Atípico',
            marker=dict(color='red', size=10, symbol='x')
        ))
        fig.add_hline(y=media + umbral*std, line_dash="dash", line_color="orange", annotation_text=f"+{umbral}σ")
        fig.add_hline(y=media - umbral*std, line_dash="dash", line_color="orange", annotation_text=f"-{umbral}σ")
        fig.update_layout(title="IDE Diario con Atípicos", xaxis_title="Fecha", yaxis_title="kWh/km")
        st.plotly_chart(fig, use_container_width=True)
        if df_at['Atípico'].any():
            st.write("#### Registros atípicos")
            df_show = df_at[df_at['Atípico']][['Fecha','Tipo Día','Odómetro [km]','IDE (kWh/km)','Fuente']].copy()
            df_show['Fecha'] = df_show['Fecha'].dt.strftime('%Y-%m-%d')
            st.dataframe(df_show.style.format({'Odómetro [km]':"{:,.1f}",'IDE (kWh/km)':"{:.4f}"}))
    else:
        st.info("📂 Sin datos suficientes para detección de atípicos.")

# TAB 7: THDR
def render_via_thdr(df_via, label):
    """Renderiza el contenido de una vía THDR."""
    if df_via.empty:
        st.info(f"📂 No hay datos cargados para {label}. Sube archivos en el panel lateral.")
        return

    df = df_via.copy()
    df['Fecha'] = df['Fecha_Op'].dt.strftime('%Y-%m-%d')

    # Métricas rápidas
    total_viajes = len(df)
    total_trenkm = df['Tren-Km'].sum()
    fechas_unicas = df['Fecha'].nunique()
    trenes_M = (df['Unidad'].astype(str).str.strip() == 'M').sum()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total viajes", f"{total_viajes}")
    c2.metric("Tren-Km", f"{total_trenkm:,.1f}")
    c3.metric("Días cargados", f"{fechas_unicas}")
    c4.metric("Viajes doble (M)", f"{trenes_M}")

    # Resumen por fecha
    st.write("#### 📅 Resumen por Fecha")
    resumen = df.groupby('Fecha').agg(
        Viajes=('Unidad', 'count'),
        Viajes_M=('Unidad', lambda x: (x.astype(str).str.strip() == 'M').sum()),
        TrenKm=('Tren-Km', 'sum')
    ).reset_index()
    st.dataframe(
        resumen.style.format({'TrenKm': "{:,.1f}"}),
        use_container_width=True
    )

    # Detalle completo con selector de fecha
    st.write("#### 🔍 Detalle por Fecha")
    fechas_disp = sorted(df['Fecha'].unique())
    fecha_sel = st.selectbox(f"Seleccionar fecha ({label})", fechas_disp, key=f"sel_{label}")
    df_sel = df[df['Fecha'] == fecha_sel]

    # Columnas a mostrar: base + columnas _min de estaciones
    cols_base = ['Viaje', 'Tren', 'Hora Salida Programada', 'Motriz 1', 'Motriz 2', 'Unidad', 'Maquinista', 'Tren-Km']
    cols_min = [c for c in df_sel.columns if '_min' in c and 'Hora Salida Programada' not in c]
    cols_show = [c for c in cols_base + cols_min if c in df_sel.columns]
    st.dataframe(make_columns_unique(df_sel[cols_show]).reset_index(drop=True), use_container_width=True)
    st.caption(f"{len(df_sel)} viajes el {fecha_sel}")


with tabs[7]:
    st.header("📋 Análisis THDR")

    # --- Sub-pestañas V1 / V2 ---
    t_v1, t_v2 = st.tabs(["🔵 Vía 1 (Puerto → Limache)", "🟠 Vía 2 (Limache → Puerto)"])
    with t_v1:
        render_via_thdr(df_thdr_v1, "Vía 1")
    with t_v2:
        render_via_thdr(df_thdr_v2, "Vía 2")



# TAB 8: SERVICIOS VS ENERGÍA
with tabs[8]:
    st.header("🔬 Servicios vs Consumo de Energía (15 min)")

    # --- Verificar datos disponibles ---
    _tiene_prmte = len(all_prmte_full) > 0
    _tiene_thdr  = not df_thdr_v1.empty or not df_thdr_v2.empty

    if not _tiene_prmte and not _tiene_thdr:
        st.info("📂 Sube archivos PRMTE (en Facturación) y THDR (Vía 1 y/o Vía 2) para este análisis.")
        st.stop()

    col_av, col_at = st.columns(2)
    col_av.metric("PRMTE disponible", "✅" if _tiene_prmte else "❌ Sin datos")
    col_at.metric("THDR disponible",  "✅" if _tiene_thdr  else "❌ Sin datos")

    # ── Helpers ──────────────────────────────────────────────────────────────
    def minutos_a_franja15(minutos):
        """Convierte minutos flotantes a string HH:MM redondeado a 15 min."""
        if minutos is None or (isinstance(minutos, float) and np.isnan(minutos)):
            return None
        total = int(minutos)
        h = total // 60
        m = (total % 60 // 15) * 15
        if h >= 24:
            return None
        return f"{h:02d}:{m:02d}"

    def str_franja_a_minutos(s):
        """HH:MM → minutos para ordenar."""
        try:
            h, m = map(int, s.split(':'))
            return h * 60 + m
        except:
            return 0

    # ── Construir series de SERVICIOS por franja 15 min ──────────────────────
    # Lógica: un tren cuenta en una franja si está circulando en ese intervalo,
    # es decir t_salida_origen <= franja < t_llegada_destino.
    # Esto refleja que un viaje dura ~1hr y ocupa ~4 franjas de 15 min.
    df_servicios = pd.DataFrame()
    if _tiene_thdr:
        partes = [df for df in [df_thdr_v1, df_thdr_v2] if not df.empty]
        df_all_thdr = pd.concat(partes, ignore_index=True)
        df_all_thdr['Fecha_str'] = df_all_thdr['Fecha_Op'].dt.strftime('%Y-%m-%d')

        # t_inicio: primera salida del viaje (salida desde origen)
        # t_fin:    última llegada del viaje (llegada a destino)
        def _primera_salida(row):
            cols_sal = [c for c in row.index if 'Salida' in c and '_min' in c]
            vals = [row[c] for c in cols_sal if pd.notna(row[c])]
            return min(vals) if vals else np.nan

        def _ultima_llegada(row):
            cols_lle = [c for c in row.index if 'Llegada' in c and '_min' in c]
            vals = [row[c] for c in row.index if 'Llegada' in c and '_min' in c and pd.notna(row[c])]
            return max(vals) if vals else np.nan

        df_all_thdr['t_ini'] = df_all_thdr.apply(_primera_salida, axis=1)
        df_all_thdr['t_fin'] = df_all_thdr.apply(_ultima_llegada, axis=1)
        df_all_thdr = df_all_thdr.dropna(subset=['t_ini', 't_fin'])

        # Calcular km recorridos por cada tren en su viaje
        # Velocidad = KM_TOTAL / duración_viaje → km recorridos en los min activos de la franja
        def _km_en_franja(t_ini, t_fin, t_f, unidad):
            """Km recorridos por un tren en la franja [t_f, t_f+15)."""
            duracion = t_fin - t_ini
            if duracion <= 0: return 0.0
            dist_total = KM_TOTAL * (2 if str(unidad).strip() == 'M' else 1)
            vel = dist_total / duracion          # km/min
            t_activo_ini = max(t_ini, t_f)
            t_activo_fin = min(t_fin, t_f + 15)
            mins_activos = max(0.0, t_activo_fin - t_activo_ini)
            return round(vel * mins_activos, 3)

        # Para cada franja de 15 min del día, contar trenes activos y km
        todas_franjas = [f"{h:02d}:{m:02d}" for h in range(24) for m in range(0,60,15)]
        filas_srv = []
        for fecha_g, grp in df_all_thdr.groupby('Fecha_str'):
            for franja in todas_franjas:
                h_f, m_f = int(franja[:2]), int(franja[3:])
                t_f = h_f * 60 + m_f
                # Tren activo si t_ini <= t_franja < t_fin
                mask = (grp['t_ini'] <= t_f) & (grp['t_fin'] > t_f)
                n_total = mask.sum()
                if n_total == 0:
                    continue
                grp_act = grp[mask]
                n_M = (grp_act['Unidad'].astype(str).str.strip() == 'M').sum()
                km_franja = sum(
                    _km_en_franja(row['t_ini'], row['t_fin'], t_f, row['Unidad'])
                    for _, row in grp_act.iterrows()
                )
                filas_srv.append({
                    'Fecha':       fecha_g,
                    'Franja':      franja,
                    'Servicios':   int(n_total),
                    'Servicios_M': int(n_M),
                    'Tren_Km':     km_franja
                })

        if filas_srv:
            df_servicios = pd.DataFrame(filas_srv)

    # ── Construir series de ENERGÍA por franja 15 min ────────────────────────
    df_energia = pd.DataFrame()
    if _tiene_prmte:
        df_prmte = pd.DataFrame(all_prmte_full)
        df_prmte['Fecha'] = df_prmte['Fecha'].dt.strftime('%Y-%m-%d')
        df_energia = (df_prmte
            .groupby(['Fecha', '15min'])['Consumo'].sum()
            .reset_index()
            .rename(columns={'15min': 'Franja', 'Consumo': 'kWh'}))

    # ── Merge ─────────────────────────────────────────────────────────────────
    if df_servicios.empty and df_energia.empty:
        st.warning("Sin datos suficientes para el análisis.")
        st.stop()

    if not df_servicios.empty and not df_energia.empty:
        df_merge = pd.merge(df_energia, df_servicios, on=['Fecha', 'Franja'], how='outer').fillna(0)
    elif not df_energia.empty:
        df_merge = df_energia.copy()
        df_merge['Servicios'] = 0
        df_merge['Servicios_M'] = 0
        df_merge['Tren_Km'] = 0.0
    else:
        df_merge = df_servicios.copy()
        df_merge['kWh'] = 0
    if 'Tren_Km' not in df_merge.columns:
        df_merge['Tren_Km'] = 0.0

    df_merge['_ord'] = df_merge['Franja'].apply(str_franja_a_minutos)
    df_merge = df_merge.sort_values(['Fecha', '_ord']).drop(columns='_ord')

    # ── Filtros ───────────────────────────────────────────────────────────────
    fechas_disp = sorted(df_merge['Fecha'].unique())
    if not fechas_disp:
        st.warning("Sin fechas en el rango seleccionado.")
        st.stop()

    st.divider()
    col_f1, col_f2, col_f3 = st.columns([2, 2, 1])
    with col_f1:
        modo = st.radio("Vista", ["Por día", "Promedio del período"], horizontal=True)
    with col_f2:
        if modo == "Por día":
            fecha_sel = st.selectbox("Fecha", fechas_disp)
            df_plot = df_merge[df_merge['Fecha'] == fecha_sel].copy()
        else:
            df_plot = df_merge.groupby('Franja').agg(
                kWh=('kWh','mean'),
                Servicios=('Servicios','mean'),
                Servicios_M=('Servicios_M','mean'),
                Tren_Km=('Tren_Km','mean')
            ).reset_index()
            df_plot['_ord'] = df_plot['Franja'].apply(str_franja_a_minutos)
            df_plot = df_plot.sort_values('_ord').drop(columns='_ord')
    with col_f3:
        mostrar_m = st.checkbox("Solo tracción doble (M)", value=False)

    col_srv = 'Servicios_M' if mostrar_m else 'Servicios'
    lbl_srv = 'Servicios tracción doble' if mostrar_m else 'Servicios totales'

    # ── Métricas rápidas ─────────────────────────────────────────────────────
    if not df_plot.empty:
        m1, m2, m3, m4 = st.columns(4)
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Total kWh",         f"{df_plot['kWh'].sum():,.0f}")
        m2.metric("Pico kWh (franja)", f"{df_plot['kWh'].max():,.0f}",
                  df_plot.loc[df_plot['kWh'].idxmax(), 'Franja'])
        m3.metric("Total servicios",   f"{df_plot[col_srv].sum():.0f}")
        m4.metric("Pico servicios",    f"{df_plot[col_srv].max():.0f}",
                  df_plot.loc[df_plot[col_srv].idxmax(), 'Franja'] if df_plot[col_srv].max() > 0 else "—")
        m5.metric("Tren-Km período",   f"{df_plot['Tren_Km'].sum():,.1f} km")

    st.divider()

    # ── Gráfico principal: dual axis ──────────────────────────────────────────
    if not df_plot.empty:
        fig_dual = go.Figure()

        fig_dual.add_trace(go.Bar(
            x=df_plot['Franja'], y=df_plot['kWh'],
            name='Energía PRMTE (kWh)',
            marker_color='rgba(0,81,149,0.7)',
            yaxis='y1'
        ))
        fig_dual.add_trace(go.Scatter(
            x=df_plot['Franja'], y=df_plot[col_srv],
            name=lbl_srv,
            mode='lines+markers',
            line=dict(color='#E85500', width=2),
            marker=dict(size=5),
            yaxis='y2'
        ))
        if 'Tren_Km' in df_plot.columns and df_plot['Tren_Km'].sum() > 0:
            fig_dual.add_trace(go.Scatter(
                x=df_plot['Franja'], y=df_plot['Tren_Km'],
                name='Tren-Km recorridos',
                mode='lines',
                line=dict(color='#00AA44', width=2, dash='dot'),
                yaxis='y3'
            ))

        titulo = (f"Energía vs Servicios — {fecha_sel}"
                  if modo == "Por día"
                  else f"Energía vs Servicios — Promedio {fechas_disp[0]} a {fechas_disp[-1]}")

        fig_dual.update_layout(
            title=titulo,
            xaxis=dict(title="Franja 15 min", tickangle=-45,
                       tickmode='array',
                       tickvals=df_plot['Franja'][::4].tolist()),
            yaxis =dict(title="kWh",       side='left',  showgrid=True),
            yaxis2=dict(title="Servicios", side='right', overlaying='y', showgrid=False),
            yaxis3=dict(title="Tren-Km",   side='right', overlaying='y', showgrid=False,
                        anchor='free', position=1.0, showticklabels=False),
            legend=dict(orientation='h', y=1.08),
            hovermode='x unified',
            height=450,
        )
        st.plotly_chart(fig_dual, use_container_width=True)

    # ── Correlación ───────────────────────────────────────────────────────────
    df_corr = df_plot.dropna(subset=['kWh', col_srv])
    df_corr = df_corr[(df_corr['kWh'] > 0) & (df_corr[col_srv] > 0)]

    if len(df_corr) >= 5:
        st.divider()
        corr = np.corrcoef(df_corr['kWh'].values, df_corr[col_srv].values)[0, 1]
        st.subheader(f"📐 Correlación energía ↔ servicios: **{corr:.3f}**")

        interp = ("muy alta 🟢" if abs(corr) > 0.8
                  else "alta 🟡" if abs(corr) > 0.6
                  else "moderada 🟠" if abs(corr) > 0.4
                  else "baja 🔴")
        st.caption(f"Correlación {interp}. {'Positiva: más servicios → más consumo.' if corr > 0 else 'Negativa: relación inversa.'}")

        # Scatter correlación
        coef = np.polyfit(df_corr[col_srv].values, df_corr['kWh'].values, 1)
        x_line = np.linspace(df_corr[col_srv].min(), df_corr[col_srv].max(), 100)
        y_line = np.polyval(coef, x_line)

        fig_sc = go.Figure()
        fig_sc.add_trace(go.Scatter(
            x=df_corr[col_srv], y=df_corr['kWh'],
            mode='markers',
            text=df_corr['Franja'],
            hovertemplate='<b>%{text}</b><br>Servicios: %{x}<br>kWh: %{y:,.0f}<extra></extra>',
            marker=dict(color='#005195', size=7, opacity=0.7),
            name='Franjas'
        ))
        fig_sc.add_trace(go.Scatter(
            x=x_line, y=y_line, mode='lines',
            line=dict(color='#E85500', dash='dash'),
            name=f'Tendencia (R²={corr**2:.3f})'
        ))
        fig_sc.update_layout(
            title='Dispersión: Servicios vs kWh por franja',
            xaxis_title=lbl_srv,
            yaxis_title='kWh',
            height=380
        )
        st.plotly_chart(fig_sc, use_container_width=True)

    # ── Tabla detalle ─────────────────────────────────────────────────────────
    with st.expander("📋 Ver tabla de datos"):
        cols_show = ['Franja', 'kWh', 'Servicios', 'Servicios_M', 'Tren_Km']
        cols_show = [c for c in cols_show if c in df_plot.columns]
        st.dataframe(
            df_plot[cols_show].style.format({
                'kWh': '{:,.1f}',
                'Servicios': '{:.1f}',
                'Servicios_M': '{:.1f}',
                'Tren_Km': '{:.2f}'
            }),
            use_container_width=True,
            height=300
        )
