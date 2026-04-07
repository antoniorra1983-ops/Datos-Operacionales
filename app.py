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
from plotly.subplots import make_subplots
import traceback

# --- 0. FUNCIÓN DE SEGURIDAD PARA COLUMNAS DUPLICADAS ---
def make_columns_unique(df):
    """Evita el error de PyArrow en st.dataframe añadiendo sufijos a columnas repetidas."""
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

def to_excel_consolidado(df_ops, df_tr, df_seat, df_prmte, df_factura):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        if not df_ops.empty: df_ops.to_excel(writer, index=False, sheet_name='Operaciones')
        if not df_tr.empty: df_tr.to_excel(writer, index=False, sheet_name='Trenes')
        if not df_seat.empty: df_seat.to_excel(writer, index=False, sheet_name='SEAT')
    return output.getvalue()

# --- 3. FUNCIONES THDR ---
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
            
        p_key = next((k for k in columnas_horas.keys() if 'puerto' in k and 'salida' in k), None)
        l_key = next((k for k in columnas_horas.keys() if 'limache' in k and 'llegada' in k), None)
        
        m = re.search(r'(\d{2})(\d{2})(\d{2})', file.name)
        df['Fecha_Op'] = pd.to_datetime(date(2000+int(m.group(3)), int(m.group(2)), int(m.group(1)))).normalize() if m else pd.to_datetime(date.today()).normalize()
        
        mask = (df['Fecha_Op'].dt.date >= start_date) & (df['Fecha_Op'].dt.date <= end_date)
        df = df[mask].copy()
        
        if not df.empty:
            c_m2 = next((c for c in df.columns if 'Motriz 2' in str(c)), None)
            df['Unidad'] = df[c_m2].apply(lambda x: 'M' if parse_latam_number(x) > 0 else 'S') if c_m2 else 'S'
            df['Tren-Km'] = 43.13 * df['Unidad'].apply(lambda x: 2 if x == 'M' else 1) if p_key and l_key else 0
            
        return df
    except: return pd.DataFrame()

# --- 4. INICIALIZACIÓN ---
if 'outliers' not in st.session_state: st.session_state.outliers = pd.DataFrame()
df_ops, df_tr, df_seat, df_energy_master = [pd.DataFrame() for _ in range(4)]
all_ops, all_tr, all_seat, all_comp_full = [], [], [], []

# --- 5. SIDEBAR ---
with st.sidebar:
    st.header("📅 Filtro Global")
    dr = st.date_input("Selecciona el Rango", value=(date(2026, 1, 1), date(2026, 1, 31)))
    if isinstance(dr, tuple) and len(dr) == 2: start_date, end_date = dr[0], dr[1]
    else: start_date = end_date = dr

    st.divider()
    f_v1 = st.file_uploader("1. THDR Vía 1", accept_multiple_files=True)
    f_v2 = st.file_uploader("2. THDR Vía 2", accept_multiple_files=True)
    f_umr = st.file_uploader("3. UMR / Odómetros", accept_multiple_files=True)
    f_seat_files = st.file_uploader("4. Energía SEAT", accept_multiple_files=True)
    f_bill_files = st.file_uploader("5. Facturación y PRMTE", accept_multiple_files=True)

# --- 6. PROCESAMIENTO DE DATOS ---
if any([f_v1, f_v2, f_umr, f_seat_files, f_bill_files]):
    # A. PROCESAR UMR (BASE DEL RESUMEN)
    if f_umr:
        for f in f_umr:
            try:
                xl = pd.ExcelFile(f)
                for sn in xl.sheet_names:
                    df_raw = pd.read_excel(f, sheet_name=sn, header=None)
                    h_r = None
                    for i in range(min(60, len(df_raw))):
                        fila_str = " ".join([str(x).upper() for x in df_raw.iloc[i].values])
                        if any(k in fila_str for k in ['FECHA', 'ODO', 'KILOM']):
                            h_r = i; break
                    
                    if h_r is not None:
                        df_p = pd.read_excel(f, sheet_name=sn, header=h_r)
                        df_p.columns = [str(c).upper().strip() for c in df_p.columns]
                        c_f = next((c for c in df_p.columns if 'FECHA' in c), None)
                        c_o = next((c for c in df_p.columns if 'ODO' in c), None)
                        c_t = next((c for c in df_p.columns if 'TREN' in c and 'KM' in c), None)
                        
                        if c_f and c_o:
                            df_p['_dt'] = pd.to_datetime(df_p[c_f], errors='coerce').dt.normalize()
                            # Filtro estricto por fecha
                            df_p = df_p[(df_p['_dt'].dt.date >= start_date) & (df_p['_dt'].dt.date <= end_date)].dropna(subset=['_dt'])
                            for _, r in df_p.iterrows():
                                all_ops.append({
                                    "Fecha": r['_dt'], 
                                    "Tipo Día": get_tipo_dia(r['_dt']),
                                    "Odómetro [km]": parse_latam_number(r[c_o]),
                                    "Tren-Km [km]": parse_latam_number(r[c_t]) if c_t else 0.0
                                })
            except: continue

    # B. PROCESAR TRENES
    if f_umr:
        for f in f_umr:
            xl = pd.ExcelFile(f)
            for sn in xl.sheet_names:
                if 'KIL' in sn.upper() and 'ODO' in sn.upper():
                    df_tr_raw = pd.read_excel(f, sheet_name=sn, header=None)
                    for i in range(len(df_tr_raw)-2):
                        for j in range(1, len(df_tr_raw.columns)):
                            v_f = pd.to_datetime(df_tr_raw.iloc[i, j], errors='coerce')
                            if pd.notna(v_f) and start_date <= v_f.date() <= end_date:
                                for k in range(i+3, min(i+45, len(df_tr_raw))):
                                    tren = str(df_tr_raw.iloc[k, 0]).strip().upper()
                                    if re.match(r'^(M|XM)', tren):
                                        all_tr.append({"Tren": tren, "Fecha": v_f.normalize(), "Valor": parse_latam_number(df_tr_raw.iloc[k, j])})

    # C. PROCESAR SEAT
    if f_seat_files:
        for f in f_seat_files:
            try:
                df_s = pd.read_excel(f, sheet_name=0, header=None)
                for i in range(len(df_s)):
                    fs = pd.to_datetime(df_s.iloc[i, 1], errors='coerce')
                    if pd.notna(fs) and start_date <= fs.date() <= end_date:
                        all_seat.append({
                            "Fecha": fs.normalize(), 
                            "E_Total": parse_latam_number(df_s.iloc[i, 3]),
                            "E_Tr": parse_latam_number(df_s.iloc[i, 5]),
                            "E_12": parse_latam_number(df_s.iloc[i, 7]),
                            "Fuente": "SEAT"
                        })
            except: continue

    # D. PROCESAR FACTURA / PRMTE
    if f_bill_files:
        for f in f_bill_files:
            xl = pd.ExcelFile(f)
            for sn in xl.sheet_names:
                sn_up = sn.upper()
                if 'FACT' in sn_up:
                    df_f = pd.read_excel(f, sheet_name=sn)
                    df_f.columns = ['Fch', 'Val']
                    df_f['dt'] = pd.to_datetime(df_f['Fch'], errors='coerce')
                    for _, r in df_f.dropna(subset=['dt']).iterrows():
                        all_comp_full.append({"Fecha": r['dt'].normalize(), "Hora": r['dt'].hour, "Consumo": abs(parse_latam_number(r['Val'])), "Fuente": "Factura"})
                if 'PRMTE' in sn_up:
                    df_prm = pd.read_excel(f, sheet_name=sn, header=None)
                    h_idx = next((i for i in range(len(df_prm)) if 'AÑO' in str(df_prm.iloc[i]).upper()), None)
                    if h_idx is not None:
                        df_pd = pd.read_excel(f, sheet_name=sn, header=h_idx)
                        df_pd['ts'] = pd.to_datetime(df_pd[['AÑO', 'MES', 'DIA', 'HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'}))
                        for _, r in df_pd.iterrows():
                            all_comp_full.append({"Fecha": r['ts'].normalize(), "Hora": r['ts'].hour, "Consumo": parse_latam_number(r.get('Retiro_Energia_Activa (kWhD)', 0)), "Fuente": "PRMTE"})

    # --- E. CONSOLIDACIÓN FINAL PARA EL RESUMEN ---
    if all_ops:
        df_ops = pd.DataFrame(all_ops).groupby("Fecha").agg({"Odómetro [km]":"sum", "Tren-Km [km]":"sum", "Tipo Día":"first"}).reset_index()
        
        # Consolidar Energía Maestra
        df_em = pd.DataFrame()
        if all_seat:
            df_em = pd.DataFrame(all_seat).drop_duplicates(subset=['Fecha'])
        
        if all_comp_full:
            df_comp = pd.DataFrame(all_comp_full).groupby(["Fecha", "Fuente"])["Consumo"].sum().reset_index()
            # Prioridad: Factura > PRMTE > SEAT
            for fnt in ["Factura", "PRMTE"]:
                df_fnt = df_comp[df_comp["Fuente"] == fnt].rename(columns={"Consumo": "E_Total"})
                if not df_fnt.empty:
                    if df_em.empty: df_em = df_fnt[["Fecha", "E_Total", "Fuente"]]
                    else:
                        df_em = pd.concat([df_em, df_fnt]).drop_duplicates(subset=["Fecha"], keep="last")

        if not df_em.empty:
            df_em["Fecha"] = pd.to_datetime(df_em["Fecha"]).dt.normalize()
            df_ops = pd.merge(df_ops, df_em, on="Fecha", how="left")
            # IDE Provisorio si no hay tracción desagregada
            if 'E_Tr' not in df_ops.columns and 'E_Total' in df_ops.columns:
                df_ops['E_Tr'] = df_ops['E_Total'] * 0.85 # Estimación basal si no hay SEAT
            df_ops['IDE (kWh/km)'] = df_ops.apply(lambda r: r['E_Tr'] / r['Odómetro [km]'] if r.get('Odómetro [km]', 0) > 0 else 0, axis=1)

    if f_v1: df_thdr_v1 = pd.concat([procesar_thdr_avanzado(f, start_date, end_date) for f in f_v1], ignore_index=True)
    if f_v2: df_thdr_v2 = pd.concat([procesar_thdr_avanzado(f, start_date, end_date) for f in f_v2], ignore_index=True)

# --- 7. TABS ---
tabs = st.tabs(["📊 Resumen", "📑 Operaciones", "📑 Trenes", "⚡ Energía", "⚖️ Comparación hr", "📈 Regresión", "🚨 Atípicos", "📋 THDR"])

with tabs[0]: # PESTAÑA RESUMEN REPARADA
    if not df_ops.empty:
        anios = sorted(df_ops['Fecha'].dt.year.unique())
        meses = sorted(df_ops['Fecha'].dt.month.unique())
        
        c1, c2 = st.columns(2)
        f_a = c1.multiselect("Año", anios, default=anios)
        f_m = c2.multiselect("Mes", meses, default=meses)
        
        df_res = df_ops[(df_ops['Fecha'].dt.year.isin(f_a)) & (df_ops['Fecha'].dt.month.isin(f_m))]
        
        if not df_res.empty:
            m1, m2, m3 = st.columns(3)
            m1.metric("Odómetro Total", f"{df_res['Odómetro [km]'].sum():,.1f} km")
            m2.metric("Tren-Km Total", f"{df_res['Tren-Km [km]'].sum():,.1f} km")
            m3.metric("IDE Promedio", f"{df_res['IDE (kWh/km)'].mean():.4f} kWh/km")
            
            fig = go.Figure()
            fig.add_trace(go.Bar(x=df_res['Fecha'], y=df_res['Odómetro [km]'], name="Kms Diarios", marker_color="#005195"))
            fig.update_layout(title="Evolución Kilometraje", xaxis_title="Fecha", yaxis_title="Km")
            st.plotly_chart(fig, use_container_width=True)
            
            st.write("### 📈 Eficiencia por Tipo de Día")
            df_jor = df_res.groupby("Tipo Día")["IDE (kWh/km)"].mean().reindex(ORDEN_TIPO_DIA).dropna()
            st.bar_chart(df_jor)
        else:
            st.warning("No hay datos para los filtros seleccionados.")
    else:
        st.info("Sube archivos (especialmente UMR) para ver el resumen.")

with tabs[1]: # OPERACIONES
    if not df_ops.empty: st.dataframe(make_columns_unique(df_ops).style.format({'Odómetro [km]': "{:,.1f}", 'IDE (kWh/km)': "{:.4f}"}))

with tabs[2]: # TRENES
    if all_tr:
        df_tr_piv = pd.DataFrame(all_tr).pivot_table(index="Tren", columns="Fecha", values="Valor", aggfunc='sum').fillna(0)
        st.dataframe(make_columns_unique(df_tr_piv).style.format("{:,.1f}"))

with tabs[3]: # ENERGÍA
    if all_seat:
        st.dataframe(pd.DataFrame(all_seat).style.format({'E_Total': "{:,.0f}"}))

with tabs[4]: # COMPARACIÓN
    if all_comp_full:
        df_c = pd.DataFrame(all_comp_full).groupby(['Fecha', 'Hora', 'Fuente'])['Consumo'].sum().reset_index()
        fig_c = px.line(df_c, x="Hora", y="Consumo", color="Fuente", title="Perfil de Carga Horario")
        st.plotly_chart(fig_c, use_container_width=True)

with tabs[5]: # REGRESIÓN
    if all_comp_full:
        df_reg = pd.DataFrame(all_comp_full)[pd.DataFrame(all_comp_full)['Hora'] <= 5]
        if len(df_reg) > 2:
            x = np.arange(len(df_reg))
            y = df_reg['Consumo'].values
            m, n = np.polyfit(x, y, 1)
            st.write(f"Ecuación Basal: $y = {m:.4f}x + {n:.2f}$")
            st.line_chart(df_reg['Consumo'])

with tabs[7]: # THDR
    st.header("📋 Datos THDR")
    c_t1, c_t2 = st.columns(2)
    with c_t1:
        if 'df_thdr_v1' in locals() and not df_thdr_v1.empty:
            st.write("Vía 1 (Puerto -> Limache)")
            st.dataframe(make_columns_unique(df_thdr_v1))
    with c_t2:
        if 'df_thdr_v2' in locals() and not df_thdr_v2.empty:
            st.write("Vía 2 (Limache -> Puerto)")
            st.dataframe(make_columns_unique(df_thdr_v2))

# REPORTE EXCEL
st.sidebar.download_button("📥 Reporte Consolidado", to_excel_consolidado(df_ops, pd.DataFrame(all_tr), pd.DataFrame(all_seat), pd.DataFrame(), pd.DataFrame()), "Reporte_EFE_SGE.xlsx")
