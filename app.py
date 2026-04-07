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
import traceback

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
ORDEN_TIPO_DIA = ["L", "S", "D/F"]

st.markdown("""
    <style>
    .stMetric { background-color: #ffffff; padding: 20px; border-radius: 10px; border-left: 5px solid #005195; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    </style>
    """, unsafe_allow_html=True)

# --- 2. FUNCIONES DE APOYO Y EXPORTACIÓN ---
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

def to_excel_consolidado(df_ops, df_tr, df_seat, df_prmte, df_fact):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        if not df_ops.empty: df_ops.to_excel(writer, index=False, sheet_name='Operaciones')
        if not df_tr.empty: df_tr.to_excel(writer, index=False, sheet_name='Kms_Trenes')
        if not df_seat.empty: df_seat.to_excel(writer, index=False, sheet_name='SEAT')
        if not df_prmte.empty: df_prmte.to_excel(writer, index=False, sheet_name='PRMTE')
        if not df_fact.empty: df_fact.to_excel(writer, index=False, sheet_name='Facturacion')
    return output.getvalue()

# --- 3. FUNCIONES THDR (INTEGRIDAD TOTAL) ---
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

def format_hms(minutos_float):
    if pd.isna(minutos_float) or minutos_float == 0: return "00:00:00"
    total_segundos = int(round(abs(minutos_float) * 60))
    h, r = divmod(total_segundos, 3600); m, s = divmod(r, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"

def format_hm_short(minutos_float):
    if pd.isna(minutos_float): return "00:00"
    h, m = divmod(int(minutos_float), 60)
    return f"{h:02d}:{m:02d}"

DISTANCIAS = {"PU-LI": 43.13, "LI-PU": 43.13, "PU-SA": 29.11, "SA-PU": 29.11, "EB-PU": 25.40, "PU-EB": 25.40, "VM-LI": 34.03, "LI-VM": 34.03}

def procesar_thdr_avanzado(file, start_date, end_date):
    try:
        try: df_raw = pd.read_excel(file, header=None)
        except: df_raw = pd.read_excel(file, header=None, engine='xlrd')
        h0, h1 = df_raw.iloc[0].fillna('').astype(str), df_raw.iloc[1].fillna('').astype(str)
        cols = [f"{h0.iloc[i].strip()}_{h1.iloc[i].strip()}" if h1.iloc[i].strip() in ['Hora Llegada', 'Hora Salida'] else h0.iloc[i].strip() for i in range(len(h0))]
        df = df_raw.iloc[2:].copy(); df.columns = cols; df = make_columns_unique(df)
        
        columnas_horas = {}
        for col in df.columns:
            cl = str(col).lower()
            if 'hora salida' in cl: columnas_horas[f"{cl.replace('hora salida','').strip()}_salida"] = col
            elif 'hora llegada' in cl: columnas_horas[f"{cl.replace('hora llegada','').strip()}_llegada"] = col
        
        for k, c in columnas_horas.items():
            df[f"{k}_min"] = df[c].apply(convertir_a_minutos)
            df[f"{k}_fmt"] = df[f"{k}_min"].apply(lambda x: format_hms(x) if pd.notna(x) else "")
            
        p_key, l_key = next((k for k in columnas_horas.keys() if 'puerto' in k and 'salida' in k), None), next((k for k in columnas_horas.keys() if 'limache' in k and 'llegada' in k), None)
        m = re.search(r'(\d{2})(\d{2})(\d{2})', file.name)
        df['Fecha_Op'] = pd.to_datetime(date(2000+int(m.group(3)), int(m.group(2)), int(m.group(1)))).normalize() if m else pd.to_datetime(date.today()).normalize()
        
        mask = (df['Fecha_Op'].dt.date >= start_date) & (df['Fecha_Op'].dt.date <= end_date)
        df = df[mask].copy()
        
        if not df.empty:
            c_m2 = next((c for c in df.columns if 'Motriz 2' in str(c)), None)
            df['Unidad'] = df[c_m2].apply(lambda x: 'M' if parse_latam_number(x) > 0 else 'S') if c_m2 else 'S'
            df['Tren-Km'] = 43.13 * df['Unidad'].apply(lambda x: 2 if x == 'M' else 1) if p_key and l_key else 0
            
            # Identificar Hora de Salida Base para conteos
            df['Hora_Ref_Min'] = df[columnas_horas[list(columnas_horas.keys())[0]]] if columnas_horas else 0
            df['Hora_Ref_Min'] = df['Hora_Ref_Min'].apply(convertir_a_minutos)
            
        return df
    except: return pd.DataFrame()

# --- 4. INICIALIZACIÓN ---
if 'outliers' not in st.session_state: st.session_state.outliers = pd.DataFrame()
df_ops, df_tr, df_seat, df_prmte_full, df_fact_full = [pd.DataFrame() for _ in range(5)]
df_thdr_v1, df_thdr_v2 = pd.DataFrame(), pd.DataFrame()
all_ops, all_tr, all_seat, all_comp_full, all_prmte_full, all_fact_full = [], [], [], [], [], []

# --- 5. SIDEBAR ---
with st.sidebar:
    st.header("📅 Filtro Global")
    dr = st.date_input("Rango de Análisis", value=(date(2026, 1, 1), date(2026, 1, 31)))
    start_date, end_date = (dr[0], dr[1]) if isinstance(dr, tuple) and len(dr) == 2 else (dr, dr)
    st.divider()
    f_v1 = st.file_uploader("1. THDR Vía 1", accept_multiple_files=True)
    f_v2 = st.file_uploader("2. THDR Vía 2", accept_multiple_files=True)
    f_umr = st.file_uploader("3. UMR / Odómetros", accept_multiple_files=True)
    f_seat_files = st.file_uploader("4. Energía SEAT", accept_multiple_files=True)
    f_bill_files = st.file_uploader("5. Facturación y PRMTE", accept_multiple_files=True)

# --- 6. PROCESAMIENTO TOTAL ---
if any([f_v1, f_v2, f_umr, f_seat_files, f_bill_files]):
    # A. PROCESAR UMR Y TRENES
    if f_umr:
        for f in f_umr:
            try:
                xl = pd.ExcelFile(f)
                for sn in xl.sheet_names:
                    df_raw = pd.read_excel(f, sheet_name=sn, header=None)
                    h_r = next((i for i in range(min(100, len(df_raw))) if any(k in str(df_raw.iloc[i].tolist()).upper() for k in ['FECHA', 'ODO', 'TREN-KM', 'KILOM'])), None)
                    if h_r is not None:
                        df_p = pd.read_excel(f, sheet_name=sn, header=h_r)
                        df_p.columns = [str(c).upper().replace('Ó','O').strip() for c in df_p.columns]
                        c_f, c_o, c_t = next((c for c in df_p.columns if 'FECHA' in c), None), next((c for c in df_p.columns if 'ODO' in c), None), next((c for c in df_p.columns if 'TREN' in c and 'KM' in c), None)
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

    # B. PROCESAR SEAT
    if f_seat_files:
        for f in f_seat_files:
            try:
                df_s = pd.read_excel(f, header=None)
                for i in range(len(df_s)):
                    fs = pd.to_datetime(df_s.iloc[i, 1], errors='coerce')
                    if pd.notna(fs) and start_date <= fs.date() <= end_date:
                        all_seat.append({"Fecha": fs.normalize(), "Total [kWh]": parse_latam_number(df_s.iloc[i, 3]), "Tracción [kWh]": parse_latam_number(df_s.iloc[i, 5]), "12 KV [kWh]": parse_latam_number(df_s.iloc[i, 7]), "% Tracción": (parse_latam_number(df_s.iloc[i, 5])/parse_latam_number(df_s.iloc[i, 3])*100 if parse_latam_number(df_s.iloc[i, 3])>0 else 0)})
            except: pass

    # C. PROCESAR FACTURA / PRMTE
    if f_bill_files:
        for f in f_bill_files:
            try:
                xl = pd.ExcelFile(f)
                for sn in xl.sheet_names:
                    sn_up = sn.upper()
                    if 'FACT' in sn_up:
                        df_f = pd.read_excel(f, sheet_name=sn); df_f.columns = ['FechaHora', 'Valor']
                        df_f['dt'] = pd.to_datetime(df_f['FechaHora'], errors='coerce')
                        for _, r in df_f.dropna(subset=['dt']).iterrows(): 
                            val_f = abs(parse_latam_number(r['Valor']))
                            all_fact_full.append({"Fecha": r['dt'].normalize(), "Hora": r['dt'].hour, "Consumo [kWh]": val_f})
                            all_comp_full.append({"Fecha": r['dt'].normalize(), "Hora": r['dt'].hour, "Consumo": val_f, "Fuente": "Factura", "Año": r['dt'].year, "Tipo Día": get_tipo_dia(r['dt'])})
                    if 'PRMTE' in sn_up:
                        df_pd_raw = pd.read_excel(f, sheet_name=sn, header=None); h = next((i for i in range(len(df_pd_raw)) if 'AÑO' in str(df_pd_raw.iloc[i]).upper()), 0)
                        df_pd = pd.read_excel(f, sheet_name=sn, header=h)
                        df_pd['ts'] = pd.to_datetime(df_pd[['AÑO', 'MES', 'DIA', 'HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'}))
                        cols_e = [c for c in df_pd.columns if 'Retiro_Energia_Activa (kWhD)' in str(c)]
                        for _, r in df_pd.iterrows(): 
                            val_p = sum([parse_latam_number(r[col]) for col in cols_e])
                            all_prmte_full.append({"Fecha": r['ts'].normalize(), "Hora": r['ts'].hour, "Energía Activa [kWh]": val_p})
                            all_comp_full.append({"Fecha": r['ts'].normalize(), "Hora": r['ts'].hour, "Consumo": val_p, "Fuente": "PRMTE", "Año": r['ts'].year, "Tipo Día": get_tipo_dia(r['ts'])})
            except: pass

    # CONSOLIDACIÓN
    if all_ops:
        df_ops = pd.DataFrame(all_ops).groupby("Fecha").agg({"Odómetro [km]":"sum", "Tren-Km [km]":"sum", "Tipo Día":"first"}).reset_index()
        df_ops['E_Total'], df_ops['E_Tr'], df_ops['IDE (kWh/km)'] = 0.0, 0.0, 0.0
        df_em = pd.DataFrame()
        if all_seat: df_em = pd.DataFrame(all_seat).drop_duplicates(subset=['Fecha']).rename(columns={"Total [kWh]":"E_Total", "Tracción [kWh]":"E_Tr"})
        if all_comp_full:
            df_cf = pd.DataFrame(all_comp_full).groupby(["Fecha", "Fuente"])["Consumo"].sum().reset_index()
            for fnt in ["Factura", "PRMTE"]:
                df_fnt = df_cf[df_cf["Fuente"] == fnt].rename(columns={"Consumo": "E_Total"})
                if not df_fnt.empty: df_em = pd.concat([df_em if not df_em.empty else pd.DataFrame(), df_fnt[["Fecha", "E_Total"]]]).drop_duplicates(subset=["Fecha"], keep="last")
        if not df_em.empty:
            df_ops = pd.merge(df_ops.drop(columns=['E_Total', 'E_Tr', 'IDE (kWh/km)']), df_em, on="Fecha", how="left").fillna(0)
            if 'E_Tr' not in df_ops.columns: df_ops['E_Tr'] = df_ops['E_Total'] * 0.85
            df_ops['IDE (kWh/km)'] = df_ops.apply(lambda r: r['E_Tr'] / r['Odómetro [km]'] if r['Odómetro [km]'] > 0 else 0, axis=1)

    if f_v1: df_thdr_v1 = make_columns_unique(pd.concat([procesar_thdr_avanzado(f, start_date, end_date) for f in f_v1], ignore_index=True))
    if f_v2: df_thdr_v2 = make_columns_unique(pd.concat([procesar_thdr_avanzado(f, start_date, end_date) for f in f_v2], ignore_index=True))

# --- 7. TABS ---
tabs = st.tabs(["📊 Resumen", "📑 Operaciones", "📑 Trenes", "⚡ Energía", "⚖️ Comparación hr", "📈 Regresión", "🚨 Atípicos", "📋 THDR"])

with tabs[0]: # RESUMEN
    if not df_ops.empty:
        df_rf = df_ops[(df_ops['Fecha'].dt.year.isin(st.multiselect("Año", sorted(df_ops['Fecha'].dt.year.unique()), sorted(df_ops['Fecha'].dt.year.unique()))))]
        if not df_rf.empty:
            c1, c2, c3 = st.columns(3); c1.metric("Odómetro", f"{df_rf['Odómetro [km]'].sum():,.1f} km"); c2.metric("Tren-Km", f"{df_rf['Tren-Km [km]'].sum():,.1f} km"); c3.metric("IDE Prom", f"{df_rf['IDE (kWh/km)'].mean():.4f}")
            st.plotly_chart(go.Figure(data=[go.Bar(x=df_rf['Fecha'], y=df_rf['Odómetro [km]'], marker_color="#005195")]), use_container_width=True)

with tabs[1]: # OPERACIONES
    if not df_ops.empty: st.dataframe(make_columns_unique(df_ops).style.format({'Odómetro [km]':"{:,.1f}", 'IDE (kWh/km)':"{:.4f}"}))

with tabs[2]: # TRENES
    if all_tr:
        df_t_p = pd.DataFrame(all_tr).pivot_table(index="Tren", columns="Fecha", values="Valor", aggfunc='sum').fillna(0)
        st.dataframe(make_columns_unique(df_t_p).style.format("{:,.1f}"))

with tabs[3]: # ENERGÍA
    e_tabs = st.tabs(["🔹 SEAT", "🔹 PRMTE", "🔹 Facturación"])
    with e_tabs[0]:
        if all_seat: st.dataframe(make_columns_unique(pd.DataFrame(all_seat)).style.format({'Total [kWh]':"{:,.0f}"}))
    with e_tabs[1]:
        if all_prmte_full:
            df_p = pd.DataFrame(all_prmte_full)
            st.dataframe(df_p.groupby("Fecha")["Energía Activa [kWh]"].sum().reset_index())
    with e_tabs[2]:
        if all_fact_full:
            df_f = pd.DataFrame(all_fact_full)
            st.dataframe(df_f.groupby("Fecha")["Consumo [kWh]"].sum().reset_index())

with tabs[7]: # 📋 THDR (CON NUEVAS TABLAS DE FRECUENCIA)
    st.header("📋 Análisis THDR")
    
    if not df_thdr_v1.empty or not df_thdr_v2.empty:
        # --- TABLA 1: SERVICIOS POR HORA ---
        st.subheader("⏱️ Cantidad de Servicios por Hora")
        
        freq_h = []
        if not df_thdr_v1.empty:
            v1_h = (df_thdr_v1['Hora_Ref_Min'] // 60).value_counts().reset_index()
            v1_h.columns = ['Hora', 'Vía 1 (Puerto->Limache)']
            freq_h.append(v1_h)
        if not df_thdr_v2.empty:
            v2_h = (df_thdr_v2['Hora_Ref_Min'] // 60).value_counts().reset_index()
            v2_h.columns = ['Hora', 'Vía 2 (Limache->Puerto)']
            freq_h.append(v2_h)
            
        if freq_h:
            df_freq_h = freq_h[0]
            if len(freq_h) > 1: df_freq_h = pd.merge(df_freq_h, freq_h[1], on='Hora', how='outer').fillna(0)
            df_freq_h = df_freq_h.sort_values('Hora')
            df_freq_h['Hora'] = df_freq_h['Hora'].apply(lambda x: f"{int(x):02d}:00")
            st.table(df_freq_h.set_index('Hora'))

        # --- TABLA 2: SERVICIOS CADA 15 MINUTOS ---
        st.subheader("⏲️ Frecuencia cada 15 Minutos")
        
        freq_15 = []
        if not df_thdr_v1.empty:
            v1_15 = ((df_thdr_v1['Hora_Ref_Min'] // 15) * 15).value_counts().reset_index()
            v1_15.columns = ['Minutos', 'Vía 1']
            freq_15.append(v1_15)
        if not df_thdr_v2.empty:
            v2_15 = ((df_thdr_v2['Hora_Ref_Min'] // 15) * 15).value_counts().reset_index()
            v2_15.columns = ['Minutos', 'Vía 2']
            freq_15.append(v2_15)
            
        if freq_15:
            df_freq_15 = freq_15[0]
            if len(freq_15) > 1: df_freq_15 = pd.merge(df_freq_15, freq_15[1], on='Minutos', how='outer').fillna(0)
            df_freq_15 = df_freq_15.sort_values('Minutos')
            df_freq_15['Intervalo'] = df_freq_15['Minutos'].apply(format_hm_short)
            st.dataframe(make_columns_unique(df_freq_15[['Intervalo', 'Vía 1', 'Vía 2']]).set_index('Intervalo'))

        st.divider()
        st.subheader("📄 Detalle de Registros")
        c1, c2 = st.columns(2)
        with c1:
            if not df_thdr_v1.empty: st.write("Vía 1 (Puerto -> Limache)"); st.dataframe(make_columns_unique(df_thdr_v1))
        with c2:
            if not df_thdr_v2.empty: st.write("Vía 2 (Limache -> Puerto)"); st.dataframe(make_columns_unique(df_thdr_v2))
    else:
        st.info("Sube archivos THDR para ver el análisis de frecuencia.")

st.sidebar.download_button("📥 Excel Completo", to_excel_consolidado(df_ops, pd.DataFrame(all_tr), pd.DataFrame(all_seat), pd.DataFrame(all_prmte_full), pd.DataFrame(all_fact_full)), "EFE_Reporte_Final.xlsx")
