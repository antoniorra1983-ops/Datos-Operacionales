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
import plotly.express as px

# --- 0. SEGURIDAD DE COLUMNAS ---
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
ORDEN_TIPO_DIA = ["L", "S", "D/F"]

st.markdown("""
    <style>
    .stMetric { background-color: #ffffff; padding: 20px; border-radius: 10px; border-left: 5px solid #005195; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    </style>
    """, unsafe_allow_html=True)

# --- 2. FUNCIONES DE APOYO ---
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

# --- 3. PROCESAMIENTO THDR ESPECIALIZADO ---
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

def procesar_thdr_eficiente(file, start_date, end_date):
    try:
        df_raw = pd.read_excel(file, header=None)
        
        # 1. Extraer Fecha desde A1 (Formato 10126 o 010126)
        fecha_str = str(df_raw.iloc[0, 0]).strip().split('.')[0]
        if len(fecha_str) == 5: fecha_str = "0" + fecha_str
        dia, mes, anio = int(fecha_str[:2]), int(fecha_str[2:4]), 2000 + int(fecha_str[4:])
        fecha_dt = pd.to_datetime(date(anio, mes, dia)).normalize()
        
        if not (start_date <= fecha_dt.date() <= end_date): return pd.DataFrame()

        # 2. Construir Cabeceras (Fila 1: Estaciones, Fila 2: Llegada/Salida)
        h1 = df_raw.iloc[0].fillna(method='ffill').astype(str)
        h2 = df_raw.iloc[1].fillna('').astype(str)
        cols = [f"{st.strip()}_{tipo.strip()}" if tipo else st.strip() for st, tipo in zip(h1, h2)]
        
        # 3. Limpiar Datos (Saltar 3 filas vacías -> Data empieza en fila index 5)
        df = df_raw.iloc[5:].copy()
        df.columns = cols
        df = make_columns_unique(df)
        df = df.dropna(how='all', axis=0)
        
        # 4. Procesar Tiempos y Metadatos
        for col in df.columns:
            if 'Hora' in col:
                df[f"{col}_min"] = df[col].apply(convertir_a_minutos)
        
        # Identificar Motrices para Tren-Km
        c_m1 = next((c for c in df.columns if 'Motriz 1' in str(c)), None)
        c_m2 = next((c for c in df.columns if 'Motriz 2' in str(c)), None)
        df['Unidad'] = df[c_m2].apply(lambda x: 'M' if parse_latam_number(x) > 0 else 'S') if c_m2 else 'S'
        
        # Distancia (Puerto-Limache = 43.13)
        es_v1 = any('PUERTO' in str(c).upper() for c in df.columns)
        df['Tren-Km'] = 43.13 * df['Unidad'].apply(lambda x: 2 if x == 'M' else 1)
        df['Fecha_Op'] = fecha_dt
        
        # Columna de Referencia para Frecuencias (Salida de la primera estación)
        col_ref = next((c for c in df.columns if ('PUERTO' in c.upper() or 'LIMACHE' in c.upper()) and 'Salida' in c and '_min' in c), None)
        if col_ref: df['Hora_Ref_Min'] = df[col_ref]
        
        return df
    except: return pd.DataFrame()

# --- 4. INICIALIZACIÓN ---
df_ops, df_tr, df_seat = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
all_ops, all_tr, all_seat, all_comp_full, all_prmte_full, all_fact_full = [], [], [], [], [], []

# --- 5. SIDEBAR ---
with st.sidebar:
    st.header("📅 Filtro Global")
    dr = st.date_input("Período", value=(date(2026, 1, 1), date(2026, 1, 31)))
    start_date, end_date = (dr[0], dr[1]) if isinstance(dr, tuple) and len(dr) == 2 else (dr, dr)
    st.divider()
    f_v1 = st.file_uploader("1. THDR Vía 1", accept_multiple_files=True)
    f_v2 = st.file_uploader("2. THDR Vía 2", accept_multiple_files=True)
    f_umr = st.file_uploader("3. UMR / Odómetros", accept_multiple_files=True)
    f_seat_files = st.file_uploader("4. Energía SEAT", accept_multiple_files=True)
    f_bill_files = st.file_uploader("5. Facturación y PRMTE", accept_multiple_files=True)

# --- 6. PROCESAMIENTO TOTAL (MANTENIENDO OPERACIONES, TRENES Y ENERGÍA) ---
if any([f_v1, f_v2, f_umr, f_seat_files, f_bill_files]):
    if f_umr:
        for f in f_umr:
            try:
                xl = pd.ExcelFile(f)
                for sn in xl.sheet_names:
                    df_raw = pd.read_excel(f, sheet_name=sn, header=None)
                    h_r = next((i for i in range(min(100, len(df_raw))) if any(k in str(df_raw.iloc[i].tolist()).upper() for k in ['FECHA', 'ODO', 'KILOM'])), None)
                    if h_r is not None:
                        df_p = pd.read_excel(f, sheet_name=sn, header=h_r)
                        df_p.columns = [str(c).upper().replace('Ó','O').strip() for c in df_p.columns]
                        c_f, c_o, c_t = next((c for c in df_p.columns if 'FECHA' in c), None), next((c for c in df_p.columns if 'ODO' in c), None), next((c for c in df_p.columns if 'KM' in c), None)
                        if c_f and c_o:
                            df_p['_dt'] = pd.to_datetime(df_p[c_f], errors='coerce').dt.normalize()
                            mask = (df_p['_dt'].dt.date >= start_date) & (df_p['_dt'].dt.date <= end_date)
                            df_filt = df_p[mask].dropna(subset=['_dt'])
                            for _, r in df_filt.iterrows():
                                all_ops.append({"Fecha": r['_dt'], "Tipo Día": get_tipo_dia(r['_dt']), "Odómetro [km]": parse_latam_number(r[c_o]), "Tren-Km [km]": parse_latam_number(r[c_t]) if c_t else 0.0})
                    if any(k in sn.upper() for k in ['KIL', 'ODO']):
                        for i in range(len(df_raw)-2):
                            for j in range(1, len(df_raw.columns)):
                                v_f = pd.to_datetime(df_raw.iloc[i, j], errors='coerce')
                                if pd.notna(v_f) and start_date <= v_f.date() <= end_date:
                                    for k in range(i+3, min(i+60, len(df_raw))):
                                        tren = str(df_raw.iloc[k, 0]).strip().upper()
                                        if re.match(r'^(M|XM)', tren): all_tr.append({"Tren": tren, "Fecha": v_f.normalize(), "Valor": parse_latam_number(df_raw.iloc[k, j])})
            except: pass

    if f_seat_files:
        for f in f_seat_files:
            try:
                df_s = pd.read_excel(f, header=None)
                for i in range(len(df_s)):
                    fs = pd.to_datetime(df_s.iloc[i, 1], errors='coerce')
                    if pd.notna(fs) and start_date <= fs.date() <= end_date:
                        all_seat.append({"Fecha": fs.normalize(), "Total [kWh]": parse_latam_number(df_s.iloc[i, 3]), "Tracción [kWh]": parse_latam_number(df_s.iloc[i, 5])})
            except: pass

    if f_bill_files:
        for f in f_bill_files:
            try:
                xl = pd.ExcelFile(f)
                for sn in xl.sheet_names:
                    if 'FACT' in sn.upper():
                        df_f = pd.read_excel(f, sheet_name=sn); df_f.columns = ['F', 'V']; df_f['dt'] = pd.to_datetime(df_f['F'], errors='coerce')
                        for _, r in df_f.dropna(subset=['dt']).iterrows(): all_comp_full.append({"Fecha": r['dt'].normalize(), "Hora": r['dt'].hour, "Consumo": abs(parse_latam_number(r['V'])), "Fuente": "Factura"})
                    if 'PRMTE' in sn.upper():
                        df_pd_raw = pd.read_excel(f, sheet_name=sn, header=None); h = next((i for i in range(len(df_pd_raw)) if 'AÑO' in str(df_pd_raw.iloc[i]).upper()), 0)
                        df_pd = pd.read_excel(f, sheet_name=sn, header=h); df_pd['ts'] = pd.to_datetime(df_pd[['AÑO', 'MES', 'DIA', 'HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'}))
                        for _, r in df_pd.iterrows(): all_comp_full.append({"Fecha": r['ts'].normalize(), "Hora": r['ts'].hour, "Consumo": parse_latam_number(r.get('Retiro_Energia_Activa (kWhD)', 0)), "Fuente": "PRMTE"})
            except: pass

    # Consolidación IDE
    if all_ops:
        df_ops = pd.DataFrame(all_ops).groupby("Fecha").agg({"Odómetro [km]":"sum", "Tren-Km [km]":"sum", "Tipo Día":"first"}).reset_index()
        df_ops['E_Total'], df_ops['E_Tr'], df_ops['IDE (kWh/km)'] = 0.0, 0.0, 0.0
        df_em = pd.DataFrame()
        if all_seat: df_em = pd.DataFrame(all_seat).drop_duplicates(subset=['Fecha']).rename(columns={"Total [kWh]":"E_Total", "Tracción [kWh]":"E_Tr"})
        if not df_em.empty:
            df_ops = pd.merge(df_ops.drop(columns=['E_Total', 'E_Tr', 'IDE (kWh/km)']), df_em, on="Fecha", how="left").fillna(0)
            df_ops['IDE (kWh/km)'] = df_ops.apply(lambda r: r['E_Tr'] / r['Odómetro [km]'] if r['Odómetro [km]'] > 0 else 0, axis=1)

    if f_v1: df_thdr_v1 = pd.concat([procesar_thdr_eficiente(f, start_date, end_date) for f in f_v1], ignore_index=True)
    if f_v2: df_thdr_v2 = pd.concat([procesar_thdr_eficiente(f, start_date, end_date) for f in f_v2], ignore_index=True)

# --- 7. TABS ---
tabs = st.tabs(["📊 Resumen", "📑 Operaciones", "📑 Trenes", "⚡ Energía", "⚖️ Comparación hr", "📈 Regresión", "🚨 Atípicos", "📋 THDR"])

with tabs[0]: # RESUMEN
    if not df_ops.empty:
        df_rf = df_ops[df_ops['Fecha'].dt.year.isin(st.multiselect("Año", sorted(df_ops['Fecha'].dt.year.unique()), sorted(df_ops['Fecha'].dt.year.unique())))]
        if not df_rf.empty:
            c1, c2, c3 = st.columns(3); c1.metric("Odómetro", f"{df_rf['Odómetro [km]'].sum():,.1f} km"); c2.metric("Tren-Km", f"{df_rf['Tren-Km [km]'].sum():,.1f} km"); c3.metric("IDE Prom", f"{df_rf['IDE (kWh/km)'].mean():.4f}")
            st.plotly_chart(go.Figure(data=[go.Bar(x=df_rf['Fecha'], y=df_rf['Odómetro [km]'], marker_color="#005195")]), use_container_width=True)

with tabs[2]: # TRENES
    if all_tr:
        st.dataframe(pd.DataFrame(all_tr).pivot_table(index="Tren", columns="Fecha", values="Valor", aggfunc='sum').fillna(0).style.format("{:,.1f}"))

with tabs[3]: # ENERGÍA
    e_tabs = st.tabs(["🔹 SEAT", "🔹 PRMTE", "🔹 Facturación"])
    with e_tabs[1]:
        if all_comp_full: st.dataframe(pd.DataFrame(all_comp_full).groupby("Fecha")["Consumo"].sum().reset_index())

with tabs[7]: # 📋 THDR (CON TABLAS DE FRECUENCIA CORREGIDAS)
    st.header("📋 Análisis THDR")
    if not df_thdr_v1.empty or not df_thdr_v2.empty:
        # Tabla 1: Servicios por Hora
        st.subheader("⏱️ Servicios por Hora")
        freq_h = []
        if not df_thdr_v1.empty and 'Hora_Ref_Min' in df_thdr_v1.columns:
            v1_h = (df_thdr_v1['Hora_Ref_Min'] // 60).value_counts().reset_index(); v1_h.columns = ['Hora', 'Vía 1']; freq_h.append(v1_h)
        if not df_thdr_v2.empty and 'Hora_Ref_Min' in df_thdr_v2.columns:
            v2_h = (df_thdr_v2['Hora_Ref_Min'] // 60).value_counts().reset_index(); v2_h.columns = ['Hora', 'Vía 2']; freq_h.append(v2_h)
        if freq_h:
            df_fh = freq_h[0]
            if len(freq_h) > 1: df_fh = pd.merge(df_fh, freq_h[1], on='Hora', how='outer').fillna(0)
            df_fh['Hora'] = df_fh['Hora'].apply(lambda x: f"{int(x):02d}:00")
            st.table(df_fh.sort_values('Hora').set_index('Hora'))

        # Tabla 2: Servicios cada 15 Minutos
        st.subheader("⏲️ Frecuencia cada 15 Minutos")
        freq_15 = []
        if not df_thdr_v1.empty and 'Hora_Ref_Min' in df_thdr_v1.columns:
            v1_15 = ((df_thdr_v1['Hora_Ref_Min'] // 15) * 15).value_counts().reset_index(); v1_15.columns = ['Min', 'Vía 1']; freq_15.append(v1_15)
        if not df_thdr_v2.empty and 'Hora_Ref_Min' in df_thdr_v2.columns:
            v2_15 = ((df_thdr_v2['Hora_Ref_Min'] // 15) * 15).value_counts().reset_index(); v2_15.columns = ['Min', 'Vía 2']; freq_15.append(v2_15)
        if freq_15:
            df_f15 = freq_15[0]
            if len(freq_15) > 1: df_f15 = pd.merge(df_f15, freq_15[1], on='Min', how='outer').fillna(0)
            df_f15['Intervalo'] = df_f15['Min'].apply(format_hm_short)
            st.dataframe(df_f15.sort_values('Min')[['Intervalo', 'Vía 1', 'Vía 2']].set_index('Intervalo'))

        st.divider()
        st.subheader("📄 Registros THDR")
        c1, c2 = st.columns(2)
        with c1: st.write("Vía 1"); st.dataframe(make_columns_unique(df_thdr_v1))
        with c2: st.write("Vía 2"); st.dataframe(make_columns_unique(df_thdr_v2))
    else: st.info("Sube archivos THDR y verifica que el rango de fecha coincida con la celda A1.")
