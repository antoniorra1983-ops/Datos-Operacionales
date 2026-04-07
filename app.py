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

# --- 0. SEGURIDAD DE COLUMNAS (PyArrow) ---
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

# --- 3. MOTOR THDR (A1: FECHA | FILAS 3-5 VACÍAS) ---
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
        
        # 1. Extraer Fecha desde A1 (Celda 0,0) - Soporta formatos numéricos/texto
        val_a1 = str(df_raw.iloc[0, 0]).strip().split('.')[0]
        if len(val_a1) == 5: val_a1 = "0" + val_a1
        # Formato esperado DDMMYY (ej 010126)
        dia, mes, anio = int(val_a1[:2]), int(val_a1[2:4]), 2000 + int(val_a1[4:])
        fecha_dt = pd.to_datetime(date(anio, mes, dia)).normalize()
        
        # Filtro estricto
        if not (start_date <= fecha_dt.date() <= end_date):
            return pd.DataFrame()

        # 2. Cabeceras (Fila 0 Estaciones, Fila 1 Llegada/Salida)
        row0 = df_raw.iloc[0].copy()
        row0[0] = np.nan # Limpiar la fecha de A1 para el ffill
        h1 = row0.ffill().astype(str)
        h2 = df_raw.iloc[1].fillna('').astype(str)
        cols = [f"{st.strip()}_{tipo.strip()}" if tipo else st.strip() for st, tipo in zip(h1, h2)]
        
        # 3. Datos en Fila 6 (Index 5) saltando 3 filas vacías
        df = df_raw.iloc[5:].copy()
        df.columns = cols
        df = make_columns_unique(df).dropna(how='all', axis=0)
        
        # Procesar horas
        for col in df.columns:
            if 'Hora' in col: df[f"{col}_min"] = df[col].apply(convertir_a_minutos)
        
        # Tren-Km
        c_m2 = next((c for c in df.columns if 'Motriz 2' in str(c)), None)
        df['Unidad'] = df[c_m2].apply(lambda x: 'M' if parse_latam_number(x) > 0 else 'S') if c_m2 else 'S'
        df['Tren-Km'] = 43.13 * df['Unidad'].apply(lambda x: 2 if x == 'M' else 1)
        df['Fecha_Op'] = fecha_dt
        
        # Ref Frecuencia
        col_ref = next((c for c in df.columns if ('PUERTO' in c.upper() or 'LIMACHE' in c.upper()) and 'Salida' in c and '_min' in c), None)
        if col_ref: df['Hora_Ref_Min'] = df[col_ref]
        
        return df
    except:
        return pd.DataFrame()

# --- 4. INICIALIZACIÓN ---
df_ops, df_tr, df_seat = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
all_ops, all_tr, all_seat, all_comp_full, all_prmte_full, all_fact_full = [], [], [], [], [], []

# --- 5. SIDEBAR ---
with st.sidebar:
    st.header("📅 Filtro Global")
    dr = st.date_input("Rango", value=(date(2026, 1, 1), date(2026, 1, 31)))
    start_date, end_date = (dr[0], dr[1]) if isinstance(dr, tuple) and len(dr) == 2 else (dr, dr)
    st.divider()
    f_v1 = st.file_uploader("1. THDR Vía 1", accept_multiple_files=True)
    f_v2 = st.file_uploader("2. THDR Vía 2", accept_multiple_files=True)
    f_umr = st.file_uploader("3. UMR / Odómetros", accept_multiple_files=True)
    f_seat_files = st.file_uploader("4. Energía SEAT", accept_multiple_files=True)
    f_bill_files = st.file_uploader("5. Facturación y PRMTE", accept_multiple_files=True)

# --- 6. PROCESAMIENTO TOTAL ---
if any([f_v1, f_v2, f_umr, f_seat_files, f_bill_files]):
    # A. UMR / TRENES
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

    # B. SEAT
    if f_seat_files:
        for f in f_seat_files:
            try:
                df_s = pd.read_excel(f, header=None)
                for i in range(len(df_s)):
                    fs = pd.to_datetime(df_s.iloc[i, 1], errors='coerce')
                    if pd.notna(fs) and start_date <= fs.date() <= end_date:
                        all_seat.append({"Fecha": fs.normalize(), "E_Total": parse_latam_number(df_s.iloc[i, 3]), "E_Tr": parse_latam_number(df_s.iloc[i, 5]), "E_12": parse_latam_number(df_s.iloc[i, 7])})
            except: pass

    # C. FACTURA / PRMTE
    if f_bill_files:
        for f in f_bill_files:
            try:
                xl = pd.ExcelFile(f)
                for sn in xl.sheet_names:
                    if 'FACT' in sn.upper():
                        df_f = pd.read_excel(f, sheet_name=sn); df_f.columns = ['F', 'V']; df_f['dt'] = pd.to_datetime(df_f['F'], errors='coerce')
                        for _, r in df_f.dropna(subset=['dt']).iterrows(): 
                            v = abs(parse_latam_number(r['V']))
                            all_fact_full.append({"Fecha": r['dt'].normalize(), "Hora": r['dt'].hour, "Consumo": v, "Fuente": "Factura"})
                    if 'PRMTE' in sn.upper():
                        df_pd_raw = pd.read_excel(f, sheet_name=sn, header=None); h = next((i for i in range(len(df_pd_raw)) if 'AÑO' in str(df_pd_raw.iloc[i]).upper()), 0)
                        df_pd = pd.read_excel(f, sheet_name=sn, header=h); df_pd['ts'] = pd.to_datetime(df_pd[['AÑO', 'MES', 'DIA', 'HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'}))
                        for _, r in df_pd.iterrows(): 
                            v = parse_latam_number(r.get('Retiro_Energia_Activa (kWhD)', 0))
                            all_prmte_full.append({"Fecha": r['ts'].normalize(), "Hora": r['ts'].hour, "Consumo": v, "Fuente": "PRMTE"})
            except: pass

    # --- JERARQUÍA DE ENERGÍA Y CRUCE DE OPERACIONES ---
    if all_ops:
        df_ops = pd.DataFrame(all_ops).groupby("Fecha").agg({"Odómetro [km]":"sum", "Tren-Km [km]":"sum", "Tipo Día":"first"}).reset_index()
        
        # 1. Obtener porcentajes de repartición desde SEAT (si existen)
        df_seat_pct = pd.DataFrame(all_seat)
        if not df_seat_pct.empty:
            df_seat_pct['%Tr'] = df_seat_pct['E_Tr'] / df_seat_pct['E_Total']
            df_seat_pct['%12'] = df_seat_pct['E_12'] / df_seat_pct['E_Total']
        
        # 2. Definir Energía Maestra según Jerarquía: Facturación > PRMTE > SEAT
        df_em = pd.DataFrame()
        if all_fact_full:
            df_em = pd.DataFrame(all_fact_full).groupby("Fecha")["Consumo"].sum().reset_index().rename(columns={"Consumo": "E_Total"})
            df_em["Fuente"] = "Factura"
        elif all_prmte_full:
            df_em = pd.DataFrame(all_prmte_full).groupby("Fecha")["Consumo"].sum().reset_index().rename(columns={"Consumo": "E_Total"})
            df_em["Fuente"] = "PRMTE"
        elif all_seat:
            df_em = pd.DataFrame(all_seat).drop_duplicates(subset=['Fecha'])[["Fecha", "E_Total", "E_Tr", "E_12"]]
            df_em["Fuente"] = "SEAT"

        if not df_em.empty:
            df_em["Fecha"] = pd.to_datetime(df_em["Fecha"]).dt.normalize()
            df_ops = pd.merge(df_ops, df_em, on="Fecha", how="left").fillna(0)
            
            # Si la fuente es Factura o PRMTE, estimamos Tr y 12kV usando porcentajes de SEAT o un basal
            if "Fuente" in df_ops.columns and df_ops["Fuente"].iloc[0] in ["Factura", "PRMTE"]:
                if not df_seat_pct.empty:
                    df_ops = pd.merge(df_ops, df_seat_pct[["Fecha", "%Tr", "%12"]], on="Fecha", how="left").fillna(0)
                    df_ops["E_Tr"] = df_ops["E_Total"] * df_ops["%Tr"]
                    df_ops["E_12"] = df_ops["E_Total"] * df_ops["%12"]
                else:
                    df_ops["E_Tr"], df_ops["E_12"] = df_ops["E_Total"] * 0.85, df_ops["E_Total"] * 0.15
            
            # Cálculo de Porcentajes e IDE
            df_ops["% Tracción"] = (df_ops["E_Tr"] / df_ops["E_Total"] * 100).fillna(0)
            df_ops["% 12 kV"] = (df_ops["E_12"] / df_ops["E_Total"] * 100).fillna(0)
            df_ops['IDE (kWh/km)'] = df_ops.apply(lambda r: r['E_Tr'] / r['Odómetro [km]'] if r['Odómetro [km]'] > 0 else 0, axis=1)

    # THDR
    if f_v1: df_thdr_v1 = pd.concat([procesar_thdr_eficiente(f, start_date, end_date) for f in f_v1], ignore_index=True)
    if f_v2: df_thdr_v2 = pd.concat([procesar_thdr_eficiente(f, start_date, end_date) for f in f_v2], ignore_index=True)

# --- 7. TABS DASHBOARD ---
tabs = st.tabs(["📊 Resumen", "📑 Operaciones", "📑 Trenes", "⚡ Energía", "⚖️ Comparación hr", "📈 Regresión", "🚨 Atípicos", "📋 THDR"])

with tabs[0]: # RESUMEN (INTACTO)
    if not df_ops.empty:
        c1, c2, c3 = st.columns(3); c1.metric("Odómetro", f"{df_ops['Odómetro [km]'].sum():,.1f}"); c2.metric("Tren-Km", f"{df_ops['Tren-Km [km]'].sum():,.1f}"); c3.metric("IDE Prom", f"{df_ops['IDE (kWh/km)'].mean():.4f}")
        st.plotly_chart(go.Figure(data=[go.Bar(x=df_ops['Fecha'], y=df_ops['Odómetro [km]'], marker_color="#005195")]), use_container_width=True)

with tabs[1]: # OPERACIONES (REPARADO CON TODAS LAS COLUMNAS)
    if not df_ops.empty:
        st.write("### Detalle Operacional e IDE")
        cols_mostrar = ['Fecha', 'Tipo Día', 'Odómetro [km]', 'Tren-Km [km]', 'E_Total', 'E_Tr', 'E_12', '% Tracción', '% 12 kV', 'IDE (kWh/km)']
        st.dataframe(make_columns_unique(df_ops[cols_mostrar]).style.format({'Odómetro [km]':"{:,.1f}", 'E_Total':"{:,.0f}", 'E_Tr':"{:,.0f}", 'E_12':"{:,.0f}", '% Tracción':"{:.1f}%", '% 12 kV':"{:.1f}%", 'IDE (kWh/km)':"{:.4f}"}))

with tabs[2]: # TRENES (INTACTO)
    if all_tr: st.dataframe(pd.DataFrame(all_tr).pivot_table(index="Tren", columns="Fecha", values="Valor", aggfunc='sum').fillna(0).style.format("{:,.1f}"))

with tabs[3]: # ENERGÍA (RESTAURADO)
    e_tabs = st.tabs(["🔹 SEAT", "🔹 PRMTE", "🔹 Facturación"])
    with e_tabs[0]:
        if all_seat: st.dataframe(pd.DataFrame(all_seat).style.format({'E_Total':"{:,.0f}"}))
    with e_tabs[1]:
        if all_prmte_full:
            df_p = pd.DataFrame(all_prmte_full)
            st.write("#### Consumo Total por Día (PRMTE)")
            st.dataframe(df_p.groupby("Fecha")["Consumo"].sum().reset_index().style.format({'Consumo':"{:,.2f}"}))
    with e_tabs[2]:
        if all_fact_full:
            df_f = pd.DataFrame(all_fact_full)
            st.write("#### Consumo Total por Día (Factura)")
            st.dataframe(df_f.groupby("Fecha")["Consumo"].sum().reset_index().style.format({'Consumo':"{:,.2f}"}))

with tabs[7]: # THDR (REPARADO)
    st.header("📋 Análisis THDR")
    if not df_thdr_v1.empty or not df_thdr_v2.empty:
        c1, c2 = st.columns(2)
        with c1:
            st.write("#### Servicios por Hora")
            freq = []
            if not df_thdr_v1.empty and 'Hora_Ref_Min' in df_thdr_v1.columns:
                v1h = (df_thdr_v1['Hora_Ref_Min'] // 60).value_counts().reset_index(); v1h.columns=['Hora','Vía 1']; freq.append(v1h)
            if not df_thdr_v2.empty and 'Hora_Ref_Min' in df_thdr_v2.columns:
                v2h = (df_thdr_v2['Hora_Ref_Min'] // 60).value_counts().reset_index(); v2h.columns=['Hora','Vía 2']; freq.append(v2h)
            if freq:
                res = freq[0]
                if len(freq)>1: res = pd.merge(res, freq[1], on='Hora', how='outer').fillna(0)
                res['Hora'] = res['Hora'].apply(lambda x: f"{int(x):02d}:00")
                st.table(res.sort_values('Hora').set_index('Hora'))
        with c2:
            st.write("#### Detalle")
            if not df_thdr_v1.empty: st.write("Vía 1"); st.dataframe(make_columns_unique(df_thdr_v1).head(50))
    else:
        st.error("Sube la THDR y verifica el rango de fechas. Si los archivos están cargados, revisa que la celda A1 tenga la fecha correcta.")

st.sidebar.download_button("📥 Excel Completo", BytesIO().getvalue(), "Reporte_EFE.xlsx") # Placeholder download
