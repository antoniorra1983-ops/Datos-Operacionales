import streamlit as st
import pandas as pd
import numpy as np
import re
import holidays
from io import BytesIO
from datetime import datetime, date, timedelta, time
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# --- 1. CONFIGURACIÓN Y ESTILOS ---
st.set_page_config(page_title="Gestión de Energía - Dashboard SGE", layout="wide", page_icon="🚆")
chile_holidays = holidays.Chile()
ORDEN_TIPO_DIA = ["L", "S", "D/F"]

st.markdown("""
    <style>
    .stMetric { background-color: #ffffff; padding: 20px; border-radius: 10px; border-left: 5px solid #005195; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    </style>
    """, unsafe_allow_html=True)

# --- 2. FUNCIONES DE EXPORTACIÓN (DEFINIDAS AL INICIO PARA EVITAR NAMEERROR) ---

def to_excel_consolidado(df_ops, df_tr, df_tr_acum, df_seat, df_p_d, df_p_15, df_fact_h, df_fact_d):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dict_dfs = {
            'Operaciones': df_ops, 'Kms_Diarios_Tren': df_tr, 
            'Odometros_Acum_Tren': df_tr_acum, 'SEAT': df_seat, 
            'PRMTE_D': df_p_d, 'PRMTE_15': df_p_15, 
            'Fact_H': df_fact_h, 'Fact_D': df_fact_d
        }
        for name, df in dict_dfs.items():
            if df is not None and not df.empty:
                df.to_excel(writer, index=False, sheet_name=name)
    return output.getvalue()

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

def convertir_a_minutos(val):
    if pd.isna(val) or str(val).strip() == "": return None
    try:
        if isinstance(val, (datetime, time)): return val.hour * 60 + val.minute + (val.second / 60.0)
        s = str(val).strip()
        m_ss = re.search(r'(\d{1,2}):(\d{2}):(\d{2})', s)
        if m_ss: return int(m_ss.group(1)) * 60 + int(m_ss.group(2)) + (int(m_ss.group(3)) / 60.0)
        m_mm = re.search(r'(\d{1,2}):(\d{2})', s)
        if m_mm: return int(m_mm.group(1)) * 60 + int(m_mm.group(2))
        return None
    except: return None

# --- 3. PROCESAMIENTO THDR (ROBUSTO) ---

DISTANCIAS = {"PU-LI": 43.13, "LI-PU": 43.13, "PU-SA": 29.11, "SA-PU": 29.11, "EB-PU": 25.40, "PU-EB": 25.40, "VM-LI": 34.03, "LI-VM": 34.03, "VM-PU": 9.10, "PU-VM": 9.10}

def procesar_thdr_avanzado(file):
    try:
        df_raw = pd.read_excel(file, header=None)
        # Limpiar encabezados de celdas combinadas
        h0 = df_raw.iloc[0].ffill().astype(str)
        h1 = df_raw.iloc[1].fillna('').astype(str)
        
        cols_raw = []
        for i in range(len(h0)):
            base, sub = h0[i].strip(), h1[i].strip()
            if "Hora" in sub: cols_raw.append(f"{base} ({sub})")
            else: cols_raw.append(base)
            
        # Deduplicar columnas para Streamlit/PyArrow
        final_cols, counts = [], {}
        for name in cols_raw:
            if name in counts:
                counts[name] += 1
                final_cols.append(f"{name}_{counts[name]}")
            else:
                counts[name] = 0
                final_cols.append(name)
        
        df = df_raw.iloc[2:].copy()
        df.columns = final_cols
        
        # Identificar estaciones dinámicamente
        c_salida = [c for c in df.columns if 'Hora Salida' in c]
        c_llegada = [c for c in df.columns if 'Hora Llegada' in c]
        
        def detectar_viaje(row):
            ini_v, ini_n, fin_v, fin_n = None, None, None, None
            for c in c_salida:
                v = convertir_a_minutos(row[c])
                if v is not None: ini_v, ini_n = v, c.split('(')[0].strip(); break
            for c in reversed(c_llegada):
                v = convertir_a_minutos(row[c])
                if v is not None: fin_v, fin_n = v, c.split('(')[0].strip(); break
            return pd.Series([ini_v, ini_n, fin_v, fin_n])

        df[['T_Ini', 'Origen', 'T_Fin', 'Destino']] = df.apply(detectar_viaje, axis=1)
        
        def find_col(keys):
            for c in df.columns:
                if any(k.lower() in c.lower() for k in keys): return c
            return None

        c_serv = find_col(['Servicio', 'N°'])
        c_prog = find_col(['Prog'])
        c_m1 = find_col(['Motriz 1', 'M1'])
        c_m2 = find_col(['Motriz 2', 'M2'])
        
        df['Servicio'] = pd.to_numeric(df[c_serv], errors='coerce').fillna(0).astype(int) if c_serv else 0
        df['Motriz 1'] = pd.to_numeric(df[c_m1], errors='coerce').fillna(0).astype(int) if c_m1 else 0
        df['Motriz 2'] = pd.to_numeric(df[c_m2], errors='coerce').fillna(0).astype(int) if c_m2 else 0
        df['Unidad'] = df['Motriz 2'].apply(lambda x: 'M' if x > 0 else 'S')
        df['Min_Prog'] = df[c_prog].apply(convertir_a_minutos) if c_prog else 0
        df['Retraso'] = df['T_Ini'] - df['Min_Prog']
        
        def calc_km(r):
            o, d = str(r['Origen'])[:2].upper(), str(r['Destino'])[:2].upper()
            map_e = {"PU":"PU", "VA":"PU", "LI":"LI", "VI":"VM", "EL":"EB"}
            k = f"{map_e.get(o,o)}-{map_e.get(d,d)}"
            return DISTANCIAS.get(k, 43.13) * (2 if r['Unidad'] == 'M' else 1)
            
        df['Tren-Km'] = df.apply(calc_km, axis=1)
        try:
            f_str = str(df_raw.iloc[0, 0]).split('.')[0].strip().zfill(6)
            df['Fecha_Op'] = f"{f_str[0:2]}/{f_str[2:4]}/20{f_str[4:6]}"
        except: df['Fecha_Op'] = ""
        
        return df[df['Servicio'] > 0]
    except Exception as e:
        st.error(f"Error procesando archivo: {e}"); return pd.DataFrame()

# --- 4. INICIALIZACIÓN DE DATAFRAMES ---
df_ops = df_tr = df_tr_acum = df_seat = df_p_d = df_f_d = df_thdr_v1 = df_thdr_v2 = pd.DataFrame()
all_ops, all_tr, all_tr_acum, all_seat, all_comp_full = [], [], [], [], []

# --- 5. SIDEBAR ---
with st.sidebar:
    st.header("📅 Filtro Global")
    date_range = st.date_input("Rango de análisis", value=(date(2026, 1, 1), date(2026, 1, 31)))
    start_date, end_date = (date_range[0], date_range[1]) if len(date_range)==2 else (date_range, date_range)
    
    st.header("📂 Carga de Archivos")
    f_v1 = st.file_uploader("1. THDR Vía 1", accept_multiple_files=True)
    f_v2 = st.file_uploader("2. THDR Vía 2", accept_multiple_files=True)
    f_umr = st.file_uploader("3. UMR / Odómetros", accept_multiple_files=True)
    f_seat = st.file_uploader("4. Energía SEAT", accept_multiple_files=True)
    f_bill = st.file_uploader("5. Facturación y PRMTE", accept_multiple_files=True)

# --- 6. PROCESAMIENTO ---
if any([f_v1, f_v2, f_umr, f_seat, f_bill]):
    if f_umr:
        for f in f_umr:
            xl = pd.ExcelFile(f)
            for sn in xl.sheet_names:
                if any(k in sn.upper() for k in ['UMR', 'RESUMEN']):
                    df_u = pd.read_excel(f, sheet_name=sn, header=None)
                    h_idx = next((i for i in range(min(50, len(df_u))) if 'ODO' in str(df_u.iloc[i]).upper()), None)
                    if h_idx is not None:
                        df_p = pd.read_excel(f, sheet_name=sn, header=h_idx)
                        df_p.columns = [re.sub(r'[^A-Z]', '', str(c).upper()) for c in df_p.columns]
                        if 'FECHA' in df_p.columns:
                            df_p['DT'] = pd.to_datetime(df_p['FECHA'], errors='coerce')
                            for _, r in df_p.dropna(subset=['DT']).iterrows():
                                if start_date <= r['DT'].date() <= end_date:
                                    all_ops.append({"Fecha": r['DT'].normalize(), "Tipo Día": get_tipo_dia(r['DT']), "N° Semana": r['DT'].isocalendar()[1], "Odómetro [km]": parse_latam_number(r.get('ODO',0)), "Tren-Km [km]": parse_latam_number(r.get('TRENKM',0))})
                if 'KIL' in sn.upper() and 'ODO' in sn.upper():
                    df_tr_raw = pd.read_excel(f, sheet_name=sn, header=None)
                    for i in range(len(df_tr_raw)-2):
                        for j in range(1, len(df_tr_raw.columns)):
                            v_fch = pd.to_datetime(df_tr_raw.iloc[i,j], errors='coerce')
                            if pd.notna(v_fch) and start_date <= v_fch.date() <= end_date:
                                for k in range(i+3, min(i+40, len(df_tr_raw))):
                                    n_tren = str(df_tr_raw.iloc[k,0]).strip().upper()
                                    if n_tren.startswith(('M','XM')):
                                        all_tr.append({"Tren": n_tren, "Fecha": v_fch.normalize(), "Valor": parse_latam_number(df_tr_raw.iloc[k,j])})

    if f_seat:
        for f in f_seat:
            df_s = pd.read_excel(f, header=None)
            for i in range(len(df_s)):
                dt = pd.to_datetime(df_s.iloc[i,1], errors='coerce')
                if pd.notna(dt) and start_date <= dt.date() <= end_date:
                    all_seat.append({"Fecha": dt.normalize(), "E_Total": parse_latam_number(df_s.iloc[i,3]), "E_Tr": parse_latam_number(df_s.iloc[i,5]), "E_12": parse_latam_number(df_s.iloc[i,7])})

    if f_v1:
        l1 = [procesar_thdr_avanzado(f) for f in f_v1]; df_thdr_v1 = pd.concat(l1) if l1 else pd.DataFrame()
    if f_v2:
        l2 = [procesar_thdr_avanzado(f) for f in f_v2]; df_thdr_v2 = pd.concat(l2) if l2 else pd.DataFrame()

    if all_ops:
        df_ops = pd.DataFrame(all_ops).drop_duplicates(subset=['Fecha']).sort_values("Fecha")
        if all_seat:
            df_ops = pd.merge(df_ops, pd.DataFrame(all_seat), on="Fecha", how="left").fillna(0)
            df_ops['IDE'] = df_ops.apply(lambda r: r['E_Tr']/r['Odómetro [km]'] if r['Odómetro [km]']>0 else 0, axis=1)

# --- 7. DASHBOARD ---
tabs = st.tabs(["📊 Resumen", "📑 Operaciones", "📑 Trenes", "⚡ Energía", "⚖️ Comparativa", "📈 Regresión", "🚨 Atípicos", "📋 THDR"])

with tabs[0]: # RESUMEN CON FILTROS RESTAURADOS
    if not df_ops.empty:
        c_f1, c_f2, c_f3 = st.columns(3)
        anios = sorted(df_ops['Fecha'].dt.year.unique())
        meses = sorted(df_ops['Fecha'].dt.month.unique())
        semanas = sorted(df_ops['N° Semana'].unique())
        f_ano = c_f1.multiselect("Año", anios, default=anios)
        f_mes = c_f2.multiselect("Mes", meses, default=meses)
        f_sem = c_f3.multiselect("Semana", semanas, default=semanas)
        
        df_filt = df_ops[df_ops['Fecha'].dt.year.isin(f_ano) & df_ops['Fecha'].dt.month.isin(f_mes) & df_ops['N° Semana'].isin(f_sem)]
        if not df_filt.empty:
            m1, m2, m3 = st.columns(3)
            m1.metric("Odómetro Total", f"{df_filt['Odómetro [km]'].sum():,.1f} km")
            m2.metric("Tren-Km Total", f"{df_filt['Tren-Km [km]'].sum():,.1f} km")
            m3.metric("IDE Promedio", f"{df_filt['IDE'].mean():.4f}")
            # Porcentajes de energía
            e_tr, e_12 = df_filt['E_Tr'].sum(), df_filt['E_12'].sum()
            if (e_tr + e_12) > 0:
                st.info(f"⚡ Composición Energía: Tracción **{e_tr/(e_tr+e_12)*100:.1f}%** | Otros 12kV **{e_12/(e_tr+e_12)*100:.1f}%**")
    else: st.warning("No hay datos para las fechas seleccionadas en el sidebar.")

with tabs[2]: # TRENES (TABLAS DIARIAS Y ACUMULADAS)
    if all_tr:
        df_tr = pd.DataFrame(all_tr)
        st.write("#### Kilometraje Diario por Unidad [km]")
        st.dataframe(df_tr.pivot_table(index="Tren", columns=df_tr["Fecha"].dt.day, values="Valor", aggfunc='sum').fillna(0))

with tabs[7]: # THDR (VISUALIZACIÓN DINÁMICA)
    st.write("### 📋 Datos THDR Dinámicos")
    col_v1, col_v2 = st.columns(2)
    with col_v1:
        st.write("#### Vía 1")
        if not df_thdr_v1.empty:
            st.dataframe(df_thdr_v1[['Fecha_Op', 'Servicio', 'Origen', 'Destino', 'Unidad', 'Tren-Km']], use_container_width=True)
        else: st.info("Sube archivos de Vía 1.")
    with col_v2:
        st.write("#### Vía 2")
        if not df_thdr_v2.empty:
            st.dataframe(df_thdr_v2[['Fecha_Op', 'Servicio', 'Origen', 'Destino', 'Unidad', 'Tren-Km']], use_container_width=True)
        else: st.info("Sube archivos de Vía 2.")

# BOTÓN DE DESCARGA EN EL SIDEBAR
if not df_ops.empty:
    st.sidebar.download_button(
        "📥 Reporte Completo", 
        to_excel_consolidado(df_ops, df_tr, pd.DataFrame(), df_seat, pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()), 
        "Reporte_SGE_EFE.xlsx"
    )
