import streamlit as st
import pandas as pd
import numpy as np
import re
import holidays
from io import BytesIO
from datetime import datetime, date

# --- 1. CONFIGURACIÓN Y UI PREMIUM ---
st.set_page_config(page_title="EFE SGE - Dashboard Profesional", layout="wide", page_icon="🚆")
chile_holidays = holidays.Chile()

# Inyección de CSS para Diseño de Fondo y Tarjetas
st.markdown("""
    <style>
    /* Fondo General con degradado */
    .stApp {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        background-attachment: fixed;
    }
    
    /* Estilo del Sidebar (Azul EFE) */
    [data-testid="stSidebar"] {
        background-color: #005195 !important;
        border-right: 1px solid #e0e0e0;
    }
    [data-testid="stSidebar"] .stMarkdown p, [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2 {
        color: white !important;
    }
    
    /* Contenedores de contenido (Glassmorphism) */
    .stTable, .stDataFrame, div[data-testid="stMetric"] {
        background-color: rgba(255, 255, 255, 0.8) !important;
        backdrop-filter: blur(10px);
        border-radius: 15px !important;
        padding: 15px;
        box-shadow: 0 8px 32px 0 rgba(31, 38, 135, 0.1);
        border: 1px solid rgba(255, 255, 255, 0.18);
    }

    /* Ajuste de métricas */
    [data-testid="stMetricValue"] {
        color: #005195 !important;
        font-weight: bold;
    }
    
    /* Títulos con sombra sutil */
    h1, h2, h3 {
        color: #003366;
        text-shadow: 1px 1px 2px rgba(0,0,0,0.1);
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. FUNCIONES TÉCNICAS (SIN CAMBIOS) ---
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

def to_excel_consolidado(df_ops, df_tr, df_tr_acum, df_seat, df_prm_d, df_fact_d):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dfs = {'Operaciones': df_ops, 'Diario_Trenes': df_tr, 'Lectura_Odometros': df_tr_acum, 'SEAT': df_seat, 'PRMTE': df_prm_d, 'Factura': df_fact_d}
        for name, df in dfs.items():
            if not df.empty: df.to_excel(writer, index=False, sheet_name=name)
    return output.getvalue()

# --- 3. SIDEBAR CON CALENDARIO ---
with st.sidebar:
    st.title("EFE Valparaíso")
    st.write("---")
    st.header("📅 Período de Análisis")
    
    # Filtro calendario estético
    date_range = st.date_input(
        "Seleccionar Rango",
        value=(date(2026, 3, 1), date.today()),
        help="Navega por el calendario para elegir inicio y fin"
    )
    
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = date_range
    else:
        start_date = end_date = (date_range[0] if isinstance(date_range, tuple) else date_range)

    st.write("---")
    st.header("📂 Carga de Archivos")
    f_umr = st.file_uploader("Odómetros / UMR", type=["xlsx"], accept_multiple_files=True)
    f_seat = st.file_uploader("Energía SEAT", type=["xlsx"], accept_multiple_files=True)
    f_prmte_bill = st.file_uploader("Facturación / PRMTE", type=["xlsx"], accept_multiple_files=True)

# --- 4. MOTOR DE DATOS (REFORZADO) ---
all_ops, all_tr, all_tr_acum, all_seat, all_prmte_15, all_fact_h = [], [], [], [], [], []
todos = (f_umr or []) + (f_seat or []) + (f_prmte_bill or [])

for f in todos:
    try:
        xl = pd.ExcelFile(f)
        for sn in xl.sheet_names:
            sn_up = sn.upper()
            
            # --- OPS / UMR ---
            if any(k in sn_up for k in ['UMR', 'RESUMEN']):
                df_raw = pd.read_excel(f, sheet_name=sn, header=None)
                h_idx = next((i for i in range(min(100, len(df_raw))) if 'FECHA' in str(df_raw.iloc[i]).upper()), None)
                if h_idx is not None:
                    df_p = pd.read_excel(f, sheet_name=sn, header=h_idx)
                    df_p.columns = [re.sub(r'[^A-Z]', '', str(c).upper()) for c in df_p.columns]
                    col_f = next((c for c in df_p.columns if 'FECHA' in c), None)
                    col_o = next((c for c in df_p.columns if 'ODO' in c and 'ACUM' not in c), None)
                    col_t = next((c for c in df_p.columns if 'TREN' in c and 'KM' in c), None)
                    if col_f:
                        df_p['_dt'] = pd.to_datetime(df_p[col_f], errors='coerce')
                        mask = (df_p['_dt'].dt.date >= start_date) & (df_p['_dt'].dt.date <= end_date)
                        for _, r in df_p[mask].dropna(subset=['_dt']).iterrows():
                            all_ops.append({"Fecha": r['_dt'].normalize(), "Tipo Día": get_tipo_dia(r['_dt']), "Odómetro [km]": parse_latam_number(r[col_o]), "Tren-Km [km]": parse_latam_number(r[col_t]), "UMR [%]": (parse_latam_number(r[col_t])/parse_latam_number(r[col_o])*100 if parse_latam_number(r[col_o])>0 else 0)})

            # --- TRENES (DOBLE TABLA) ---
            if 'ODO' in sn_up and 'KIL' in sn_up:
                df_tr_raw = pd.read_excel(f, sheet_name=sn, header=None)
                blocks = []
                for i in range(len(df_tr_raw)-2):
                    for j in range(1, len(df_tr_raw.columns)):
                        dv = pd.to_datetime(df_tr_raw.iloc[i, j], errors='coerce')
                        if pd.notna(dv) and start_date <= dv.date() <= end_date:
                            if i not in [b[0] for b in blocks]: blocks.append((i, dv))
                
                for idx, (ri, sd) in enumerate(blocks):
                    ctx = str(df_tr_raw.iloc[ri:ri+3, 0:3]).upper()
                    is_acum = any(k in ctx for k in ['ACUM', 'ODO', 'LECTURA'])
                    c_map = {j: pd.to_datetime(df_tr_raw.iloc[ri, j]).normalize() for j in range(1, len(df_tr_raw.columns)) if pd.notna(pd.to_datetime(df_tr_raw.iloc[ri, j], errors='coerce'))}
                    for k in range(ri+3, min(ri+40, len(df_tr_raw))):
                        tr_name = str(df_tr_raw.iloc[k, 0]).strip().upper()
                        if re.match(r'^(M|XM)', tr_name):
                            for ci, cf in c_map.items():
                                val = parse_latam_number(df_tr_raw.iloc[k, ci])
                                dp = {"Tren": tr_name, "Fecha": cf, "Día": cf.day, "Valor": val}
                                if is_acum or idx > 0: all_tr_acum.append(dp)
                                else: all_tr.append(dp)

            # --- SEAT ---
            if 'SEAT' in sn_up and 'SER' in sn_up:
                df_s = pd.read_excel(f, sheet_name=sn, header=None)
                for i in range(len(df_s)):
                    fs = pd.to_datetime(df_s.iloc[i, 1], errors='coerce')
                    if pd.notna(fs) and start_date <= fs.date() <= end_date:
                        all_seat.append({"Fecha": fs.normalize(), "Total [kWh]": parse_latam_number(df_s.iloc[i, 3]), "Tracción [kWh]": parse_latam_number(df_s.iloc[i, 5]), "12 KV [kWh]": parse_latam_number(df_s.iloc[i, 7]), "% Tracción": (parse_latam_number(df_s.iloc[i, 5])/parse_latam_number(df_s.iloc[i, 3])*100 if parse_latam_number(df_s.iloc[i, 3])>0 else 0), "% 12 KV": (parse_latam_number(df_s.iloc[i, 7])/parse_latam_number(df_s.iloc[i, 3])*100 if parse_latam_number(df_s.iloc[i, 3])>0 else 0)})
    except: continue

# --- 5. JERARQUÍA Y RENDERIZADO ---
df_ops_f, df_tr_f, df_tr_a_f, df_seat_f, df_e_m = [pd.DataFrame()] * 5

if any([all_ops, all_tr, all_tr_acum, all_seat]):
    if all_ops: df_ops_f = pd.DataFrame(all_ops).drop_duplicates(subset=['Fecha']).sort_values("Fecha")
    if all_tr: df_tr_f = pd.DataFrame(all_tr).sort_values(["Fecha", "Tren"])
    if all_tr_acum: df_tr_a_f = pd.DataFrame(all_tr_acum).sort_values(["Fecha", "Tren"])
    if all_seat: df_seat_f = pd.DataFrame(all_seat).drop_duplicates(subset=['Fecha']).sort_values("Fecha")
    
    # Jerarquía: Factura > PRMTE > SEAT (Lógica maestra)
    if not df_seat_f.empty:
        df_e_m = df_seat_f[["Fecha", "Total [kWh]", "Tracción [kWh]", "12 KV [kWh]"]].copy().rename(columns={"Total [kWh]":"E_Total", "Tracción [kWh]":"E_Tr", "12 KV [kWh]":"E_12"})
        df_e_m["Fuente"] = "SEAT"
        if not df_ops_f.empty: df_ops_f = pd.merge(df_ops_f, df_e_m, on="Fecha", how="left")

    tabs = st.tabs(["📊 Resumen", "📑 Operaciones", "📑 Trenes", "⚡ SEAT", "📈 PRMTE", "💰 Facturación"])
    
    with tabs[0]: # Resumen Estético
        st.write(f"### Análisis Período: {start_date} al {end_date}")
        c1, c2, c3 = st.columns(3)
        to = df_ops_f["Odómetro [km]"].sum() if not df_ops_f.empty else 0
        tk = df_ops_f["Tren-Km [km]"].sum() if not df_ops_f.empty else 0
        c1.metric("Odómetro Total", f"{to:,.1f} km")
        c2.metric("Tren-Km Total", f"{tk:,.1f} km")
        c3.metric("UMR Promedio", f"{(tk/to*100 if to>0 else 0):.2f} %")
        
        if not df_ops_f.empty:
            st.divider()
            df_ops_f['Tipo Día'] = pd.Categorical(df_ops_f['Tipo Día'], categories=['L', 'S', 'D/F'], ordered=True)
            res = df_ops_f.groupby("Tipo Día").agg({"Odómetro [km]":"sum", "Tren-Km [km]":"sum", "UMR [%]":"mean", "E_Total":"sum"}).reset_index()
            st.write("#### Consumo Eléctrico vs Operación")
            st.table(res.style.format("{:,.1f}"))

    with tabs[2]: # Trenes
        if not df_tr_f.empty:
            st.write("#### Kilometraje Diario Diario")
            st.dataframe(df_tr_f.pivot_table(index="Tren", columns=df_tr_f["Fecha"].dt.day, values="Valor", aggfunc='sum').fillna(0).style.format("{:,.1f}"), use_container_width=True)
        if not df_tr_a_f.empty:
            st.write("#### Lectura Odómetro Acumulado")
            st.dataframe(df_tr_a_f.pivot_table(index="Tren", columns=df_tr_a_f["Fecha"].dt.day, values="Valor", aggfunc='max').fillna(0).style.format("{:,.0f}"), use_container_width=True)

    st.sidebar.download_button("📥 Exportar Informe", to_excel_consolidado(df_ops_f, df_tr_f, df_tr_a_f, df_seat_f, pd.DataFrame(), pd.DataFrame()), "Reporte_SGE_EFE.xlsx")
else:
    st.info("👋 Sube los archivos para activar el diseño del Dashboard.")
