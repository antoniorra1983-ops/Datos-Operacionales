import streamlit as st
import pandas as pd
import numpy as np
import re
import holidays
from io import BytesIO
from datetime import datetime

# --- 1. CONFIGURACIÓN ---
st.set_page_config(page_title="EFE Valparaíso - Dashboard SGE", layout="wide", page_icon="🚆")

# Configuración de feriados de Chile
chile_holidays = holidays.Chile()

# Estilo para métricas
st.markdown("""
    <style>
    .stMetric { 
        background-color: #ffffff; 
        padding: 20px; 
        border-radius: 10px; 
        border-left: 5px solid #005195; 
        box-shadow: 0 2px 4px rgba(0,0,0,0.05); 
    }
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

def to_excel_consolidado(df_ops, df_trenes, df_seat, df_prmte, df_factura):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        if not df_ops.empty: df_ops.to_excel(writer, index=False, sheet_name='Datos_Operacionales')
        if not df_trenes.empty: df_trenes.to_excel(writer, index=False, sheet_name='Detalle_Kilometraje')
        if not df_seat.empty: df_seat.to_excel(writer, index=False, sheet_name='Energia_SEAT')
        if not df_prmte.empty: df_prmte.to_excel(writer, index=False, sheet_name='Medidas_PRMTE')
        if not df_factura.empty: df_factura.to_excel(writer, index=False, sheet_name='Consumo_Factura')
    return output.getvalue()

def color_umr(val):
    if val > 96.4: return 'color: green; font-weight: bold;'
    elif val < 96.4: return 'color: red; font-weight: bold;'
    return 'color: black;'

# --- 3. SIDEBAR (FILTROS Y CARGADORES) ---
MESES_NOMBRES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
MESES_NUM_MAP = {nombre: i+1 for i, nombre in enumerate(MESES_NOMBRES)}

with st.sidebar:
    st.header("📂 Gestión de Archivos")
    f_umr_list = st.file_uploader("1. UMR / Odómetros", type=["xlsx"], accept_multiple_files=True)
    f_seat_list = st.file_uploader("2. Energía SEAT (Resumen)", type=["xlsx"], accept_multiple_files=True)
    f_prmte_list = st.file_uploader("3. Medidas PRMTE (15 min)", type=["xlsx"], accept_multiple_files=True)
    f_factura_list = st.file_uploader("4. Consumo Factura (Horario)", type=["xlsx"], accept_multiple_files=True)
    
    st.divider()
    f_anio_list = st.multiselect("Años", [2024, 2025, 2026], default=[2025, 2026])
    f_mes_list = st.multiselect("Meses", MESES_NOMBRES, default=[MESES_NOMBRES[datetime.now().month - 1]])
    meses_num = [MESES_NUM_MAP[m] for m in f_mes_list]
    f_dias = st.multiselect("Días", list(range(1, 32)), default=list(range(1, 32)))

# --- 4. PROCESAMIENTO ---
all_data_ops, all_data_trenes, all_data_seat, all_data_prmte, all_data_factura = [], [], [], [], []

# [Lógica UMR / Odómetros / Energía SEAT se mantiene idéntica...]
if f_umr_list:
    for f in f_umr_list:
        try:
            xl = pd.ExcelFile(f)
            sn_res = next((s for s in xl.sheet_names if 'UMR' in s.upper() and 'RESUMEN' in s.upper()), None)
            if sn_res:
                df_raw_res = pd.read_excel(f, sheet_name=sn_res, header=None)
                hdr_row = next((i for i in range(min(50, len(df_raw_res))) if 'ODO' in " ".join(df_raw_res.iloc[i].astype(str)).upper()), None)
                if hdr_row is not None:
                    cols_clean = [re.sub(r'[^A-Z]', '', str(c).upper().replace('Ó','O')) for c in df_raw_res.iloc[hdr_row]]
                    idx_fch = next((i for i, c in enumerate(cols_clean) if 'FECHA' in c), None)
                    idx_odo = next((i for i, c in enumerate(cols_clean) if 'ODO' in c and 'ACUM' not in str(df_raw_res.iloc[hdr_row, i]).upper()), None)
                    idx_tkm = next((i for i, c in enumerate(cols_clean) if 'TREN' in c and 'KM' in c), None)
                    if None not in [idx_fch, idx_odo, idx_tkm]:
                        df_ext = df_raw_res.iloc[hdr_row + 1:].copy()
                        df_ext['_dt'] = pd.to_datetime(df_ext.iloc[:, idx_fch], errors='coerce')
                        mask = (df_ext['_dt'].dt.day.isin(f_dias)) & (df_ext['_dt'].dt.month.isin(meses_num)) & (df_ext['_dt'].dt.year.isin(f_anio_list))
                        for _, row in df_ext[mask].iterrows():
                            fch = row.iloc[idx_fch]
                            t_dia = "D/F" if (fch in chile_holidays or fch.strftime('%A') == 'Sunday') else ("S" if fch.strftime('%A') == 'Saturday' else "L")
                            o, t = parse_latam_number(row.iloc[idx_odo]), parse_latam_number(row.iloc[idx_tkm])
                            all_data_ops.append({"Fecha": fch.strftime('%d/%m/%Y'), "Tipo Día": t_dia, "N° Semana": fch.isocalendar()[1], "Odómetro [km]": o, "Tren-Km [km]": t, "UMR [%]": (t/o*100 if o>0 else 0), "Timestamp": fch})

            sn_trenes = next((s for s in xl.sheet_names if 'ODO' in s.upper() and 'KIL' in s.upper()), None)
            if sn_trenes:
                df_raw_tr = pd.read_excel(f, sheet_name=sn_trenes, header=None)
                date_row_idx, col_to_date = None, {}
                for i in range(min(50, len(df_raw_tr)-2)):
                    for j in range(1, len(df_raw_tr.columns)):
                        parsed_date = pd.to_datetime(df_raw_tr.iloc[i, j], errors='coerce')
                        if pd.notna(parsed_date) and parsed_date.year in f_anio_list:
                            if 'KILO' in str(df_raw_tr.iloc[i+1, j]).upper() and 'DIARIO' in str(df_raw_tr.iloc[i+2, j]).upper():
                                date_row_idx, col_to_date[j] = i, parsed_date
                if date_row_idx is not None:
                    for i in range(date_row_idx+3, len(df_raw_tr)):
                        cell_a = str(df_raw_tr.iloc[i, 0]).strip().upper()
                        if re.match(r'^(M0[1-9]|M1[0-9]|M2[0-7]|XM2[8-9]|XM3[0-5])$', cell_a):
                            for col_idx, col_date in col_to_date.items():
                                if col_date.day in f_dias and col_date.month in meses_num and col_date.year in f_anio_list:
                                    all_data_trenes.append({"Tren": cell_a, "Fecha": col_date.strftime('%d/%m/%Y'), "Día": col_date.day, "Kilometraje Diario [km]": parse_latam_number(df_raw_tr.iloc[i, col_idx]), "Timestamp": col_date})
        except: continue

if f_seat_list:
    for f in f_seat_list:
        try:
            xl = pd.ExcelFile(f)
            sn_seat = next((s for s in xl.sheet_names if 'SEAT' in s.upper() and 'SER' in s.upper()), None)
            if sn_seat:
                df_raw_s = pd.read_excel(f, sheet_name=sn_seat, header=None)
                for i in range(len(df_raw_s)):
                    fch_s = pd.to_datetime(df_raw_s.iloc[i, 1], errors='coerce')
                    if pd.notna(fch_s) and fch_s.year in f_anio_list and fch_s.month in meses_num and fch_s.day in f_dias:
                        tot, trac, kv12 = parse_latam_number(df_raw_s.iloc[i, 3]), parse_latam_number(df_raw_s.iloc[i, 5]), parse_latam_number(df_raw_s.iloc[i, 7])
                        all_data_seat.append({"Fecha": fch_s.strftime('%d/%m/%Y'), "Total [kWh]": tot, "Tracción [kWh]": trac, "12 KV [kWh]": kv12, "% Tracción": (trac/tot*100 if tot>0 else 0), "% 12 KV": (kv12/tot*100 if tot>0 else 0), "Timestamp": fch_s})
        except: continue

# --- PROCESAR MEDIDAS PRMTE (NUEVO) ---
if f_prmte_list:
    for f in f_prmte_list:
        try:
            xl = pd.ExcelFile(f)
            sn_prmte = next((s for s in xl.sheet_names if 'MEDIDAS' in s.upper() or 'PRMTE' in s.upper()), xl.sheet_names[0])
            df_p = pd.read_excel(f, sheet_name=sn_prmte, header=None)
            h_row = next(i for i, row in df_p.iterrows() if 'AÑO' in [str(c).upper() for c in row])
            df_p.columns = [str(c).strip() for c in df_p.iloc[h_row]]
            df_p = df_p.iloc[h_row+1:].copy()
            df_p['Timestamp'] = pd.to_datetime(df_p[['AÑO', 'MES', 'DIA']].rename(columns={'AÑO':'year', 'MES':'month', 'DIA':'day'}))
            
            mask = (df_p['Timestamp'].dt.year.isin(f_anio_list)) & (df_p['Timestamp'].dt.month.isin(meses_num)) & (df_p['Timestamp'].dt.day.isin(f_dias))
            df_p = df_p[mask]
            
            for d, group in df_p.groupby(df_p['Timestamp'].dt.date):
                retiro = group['Retiro_Energia_Activa (kWhD)'].apply(parse_latam_number).sum()
                all_data_prmte.append({"Fecha": d.strftime('%d/%m/%Y'), "Retiro Energía (PRMTE) [kWh]": retiro, "Timestamp": pd.Timestamp(d)})
        except: continue

# --- PROCESAR CONSUMO FACTURA (NUEVO) ---
if f_factura_list:
    for f in f_factura_list:
        try:
            df_f = pd.read_excel(f)
            df_f['Timestamp'] = pd.to_datetime(df_f.iloc[:, 0], errors='coerce')
            df_f = df_f.dropna(subset=['Timestamp'])
            mask = (df_f['Timestamp'].dt.year.isin(f_anio_list)) & (df_f['Timestamp'].dt.month.isin(meses_num)) & (df_f['Timestamp'].dt.day.isin(f_dias))
            df_f = df_f[mask]
            
            for d, group in df_f.groupby(df_f['Timestamp'].dt.date):
                factura = group.iloc[:, 1].apply(parse_latam_number).sum()
                all_data_factura.append({"Fecha": d.strftime('%d/%m/%Y'), "Consumo Factura [kWh]": abs(factura), "Timestamp": pd.Timestamp(d)})
        except: continue

# --- 5. RENDERIZADO ---
if any([all_data_ops, all_data_seat, all_data_prmte, all_data_factura]):
    df_ops_final = pd.DataFrame(all_data_ops).drop_duplicates(subset=['Fecha'], keep='last').sort_values("Timestamp") if all_data_ops else pd.DataFrame()
    df_trenes_final = pd.DataFrame(all_data_trenes).sort_values(["Timestamp", "Tren"]) if all_data_trenes else pd.DataFrame()
    df_seat_final = pd.DataFrame(all_data_seat).drop_duplicates(subset=['Fecha']).sort_values("Timestamp") if all_data_seat else pd.DataFrame()
    df_prmte_final = pd.DataFrame(all_data_prmte).drop_duplicates(subset=['Fecha']).sort_values("Timestamp") if all_data_prmte else pd.DataFrame()
    df_factura_final = pd.DataFrame(all_data_factura).drop_duplicates(subset=['Fecha']).sort_values("Timestamp") if all_data_factura else pd.DataFrame()

    tab_resumen, tab_datos, tab_trenes, tab_seat, tab_prmte, tab_factura = st.tabs(["📊 Resumen", "📑 Datos operacionales", "📑 Odómetro por Tren", "⚡ Energía SEAT", "📈 Medidas PRMTE", "💰 Facturación"])
    
    with tab_resumen:
        if not df_ops_final.empty:
            st.subheader("Indicadores Globales")
            c1, c2, c3 = st.columns(3)
            tot_o, tot_t = df_ops_final["Odómetro [km]"].sum(), df_ops_final["Tren-Km [km]"].sum()
            c1.metric("Odómetro Total", f"{tot_o:,.1f} km")
            c2.metric("Tren-Km Total", f"{tot_t:,.1f} km")
            c3.metric("UMR Global", f"{(tot_t / tot_o * 100 if tot_o > 0 else 0):.2f} %")
            st.divider()
            df_ops_final['Tipo Día'] = pd.Categorical(df_ops_final['Tipo Día'], categories=['L', 'S', 'D/F'], ordered=True)
            resumen_tipo = df_ops_final.groupby("Tipo Día").agg({"Odómetro [km]": "sum", "Tren-Km [km]": "sum", "UMR [%]": "mean"}).reset_index()
            st.table(resumen_tipo.style.format({"Odómetro [km]": "{:,.1f}", "Tren-Km [km]": "{:,.1f}", "UMR [%]": "{:.2f}%"}))

    with tab_datos:
        if not df_ops_final.empty:
            cols_v = ["Fecha", "Tipo Día", "N° Semana", "Odómetro [km]", "Tren-Km [km]", "UMR [%]"]
            st.dataframe(df_ops_final[cols_v].style.format({"Odómetro [km]": "{:,.1f}", "Tren-Km [km]": "{:,.1f}", "UMR [%]": "{:.2f}%"}).applymap(color_umr, subset=['UMR [%]']), use_container_width=True)

    with tab_trenes:
        if not df_trenes_final.empty:
            st.write("### 📅 Matriz de Kilometraje Diario")
            df_pivot = df_trenes_final.pivot_table(index="Tren", columns="Día", values="Kilometraje Diario [km]", aggfunc='sum').fillna(0)
            st.dataframe(df_pivot.style.format("{:,.1f}"), use_container_width=True)

    with tab_seat:
        if not df_seat_final.empty:
            st.subheader("Consumo de Energía SEAT")
            e1, e2, e3 = st.columns(3)
            g_tot, g_trac, g_kv12 = df_seat_final['Total [kWh]'].sum(), df_seat_final['Tracción [kWh]'].sum(), df_seat_final['12 KV [kWh]'].sum()
            e1.metric("Consumo Total", f"{g_tot:,.0f} kWh")
            e2.metric("Tracción", f"{g_trac:,.0f} kWh", f"{(g_trac/g_tot*100 if g_tot>0 else 0):.1f}%")
            e3.metric("12 KV", f"{g_kv12:,.0f} kWh", f"{(g_kv12/g_tot*100 if g_tot>0 else 0):.1f}%")
            st.dataframe(df_seat_final[["Fecha", "Total [kWh]", "Tracción [kWh]", "% Tracción", "12 KV [kWh]", "% 12 KV"]].style.format({"Total [kWh]": "{:,.0f}", "Tracción [kWh]": "{:,.0f}", "12 KV [kWh]": "{:,.0f}", "% Tracción": "{:.2f}%", "% 12 KV": "{:.2f}%"}), use_container_width=True)

    with tab_prmte:
        if not df_prmte_final.empty:
            st.subheader("Medidas Oficiales (PRMTE)")
            st.metric("Retiro Total Mes", f"{df_prmte_final['Retiro Energía (PRMTE) [kWh]'].sum():,.1f} kWh")
            st.dataframe(df_prmte_final[["Fecha", "Retiro Energía (PRMTE) [kWh]"]].style.format({"Retiro Energía (PRMTE) [kWh]": "{:,.1f}"}), use_container_width=True)

    with tab_factura:
        if not df_factura_final.empty:
            st.subheader("Consumo según Factura")
            st.metric("Total Facturado Mes", f"{df_factura_final['Consumo Factura [kWh]'].sum():,.1f} kWh")
            st.dataframe(df_factura_final[["Fecha", "Consumo Factura [kWh]"]].style.format({"Consumo Factura [kWh]": "{:,.1f}"}), use_container_width=True)

    st.sidebar.download_button("📥 Reporte Consolidado", to_excel_consolidado(df_ops_final, df_trenes_final, df_seat_final, df_prmte_final, df_factura_final), "Reporte_SGE_EFE.xlsx")
else:
    st.info("👋 Sube los archivos para comenzar el análisis.")
