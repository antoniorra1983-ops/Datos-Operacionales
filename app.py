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

def to_excel_consolidado(df_ops, df_trenes, df_seat):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        if not df_ops.empty:
            df_ops.to_excel(writer, index=False, sheet_name='Datos_Operacionales')
        if not df_trenes.empty:
            df_trenes.to_excel(writer, index=False, sheet_name='Detalle_Kilometraje')
        if not df_seat.empty:
            df_seat.to_excel(writer, index=False, sheet_name='Energia_SEAT')
    return output.getvalue()

def color_umr(val):
    if val > 96.4: return 'color: green; font-weight: bold;'
    elif val < 96.4: return 'color: red; font-weight: bold;'
    return 'color: black;'

# --- 3. SIDEBAR (FILTROS) ---
MESES_NOMBRES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
MESES_NUM_MAP = {nombre: i+1 for i, nombre in enumerate(MESES_NOMBRES)}

with st.sidebar:
    st.header("📂 Gestión de Archivos")
    f_umr_list = st.file_uploader("Subir archivos UMR / Odómetros", type=["xlsx"], accept_multiple_files=True)
    f_seat_list = st.file_uploader("Subir archivos Energía SEAT", type=["xlsx"], accept_multiple_files=True)
    
    st.divider()
    f_anio_list = st.multiselect("Años", [2024, 2025, 2026], default=[2025, 2026])
    f_mes_list = st.multiselect("Meses", MESES_NOMBRES, default=[MESES_NOMBRES[datetime.now().month - 1]])
    meses_num = [MESES_NUM_MAP[m] for m in f_mes_list]
    f_dias = st.multiselect("Días", list(range(1, 32)), default=list(range(1, 32)))

# --- 4. PROCESAMIENTO ---
all_data_ops = []
all_data_trenes = []
all_data_seat = []

# Procesar archivos UMR
if f_umr_list:
    for f in f_umr_list:
        try:
            xl = pd.ExcelFile(f)
            sn_resumen = next((s for s in xl.sheet_names if 'UMR' in s.upper() and 'RESUMEN' in s.upper()), None)
            if sn_resumen:
                df_raw_res = pd.read_excel(f, sheet_name=sn_resumen, header=None)
                hdr_row = None
                for i in range(min(100, len(df_raw_res))):
                    fila_txt = " ".join(df_raw_res.iloc[i].astype(str)).upper()
                    if ('ODO' in fila_txt or 'FECHA' in fila_txt) and 'TREN' in fila_txt:
                        hdr_row = i; break
                if hdr_row is not None:
                    cols_orig = df_raw_res.iloc[hdr_row].astype(str).tolist()
                    cols_clean = [re.sub(r'[^A-Z]', '', c.upper().replace('Ó','O')) for c in cols_orig]
                    idx_fch = next((i for i, c in enumerate(cols_clean) if 'FECHA' in c), None)
                    idx_odo = next((i for i, c in enumerate(cols_clean) if 'ODO' in c and 'ACUM' not in cols_orig[i].upper()), None)
                    idx_tkm = next((i for i, c in enumerate(cols_clean) if 'TREN' in c and 'KM' in c), None)
                    if None not in [idx_fch, idx_odo, idx_tkm]:
                        df_ext = df_raw_res.iloc[hdr_row + 1:].copy()
                        df_ext['_dt'] = pd.to_datetime(df_ext.iloc[:, idx_fch], errors='coerce')
                        mask = (df_ext['_dt'].dt.day.isin(f_dias)) & (df_ext['_dt'].dt.month.isin(meses_num)) & (df_ext['_dt'].dt.year.isin(f_anio_list))
                        for _, row in df_ext[mask].iterrows():
                            fch = row.iloc[idx_fch]
                            if not isinstance(fch, (datetime, pd.Timestamp)): continue
                            t_dia = "D/F" if (fch in chile_holidays or fch.strftime('%A') == 'Sunday') else ("S" if fch.strftime('%A') == 'Saturday' else "L")
                            o, t = parse_latam_number(row.iloc[idx_odo]), parse_latam_number(row.iloc[idx_tkm])
                            all_data_ops.append({
                                "Fecha": fch.strftime('%d/%m/%Y'), "Tipo Día": t_dia, "N° Semana": fch.isocalendar()[1],
                                "Odómetro [km]": o, "Tren-Km [km]": t, "UMR [%]": (t / o * 100) if o > 0 else 0,
                                "Timestamp": fch, "Archivo": f.name
                            })
            sn_trenes = next((s for s in xl.sheet_names if 'ODO' in s.upper() and 'KIL' in s.upper()), None)
            if sn_trenes:
                df_raw_tr = pd.read_excel(f, sheet_name=sn_trenes, header=None)
                date_row_idx = None
                col_to_date = {} 
                for i in range(min(50, len(df_raw_tr) - 2)):
                    for j in range(1, len(df_raw_tr.columns)):
                        val = df_raw_tr.iloc[i, j]
                        parsed_date = pd.to_datetime(val, errors='coerce')
                        if pd.notna(parsed_date) and parsed_date.year in f_anio_list:
                            if 'KILO' in str(df_raw_tr.iloc[i+1, j]).upper() and 'DIARIO' in str(df_raw_tr.iloc[i+2, j]).upper():
                                date_row_idx = i
                                col_to_date[j] = parsed_date
                if date_row_idx is not None:
                    for i in range(date_row_idx + 3, len(df_raw_tr)):
                        cell_a = str(df_raw_tr.iloc[i, 0]).strip().upper()
                        if re.match(r'^(M0[1-9]|M1[0-9]|M2[0-7]|XM2[8-9]|XM3[0-5])$', cell_a):
                            for col_idx, col_date in col_to_date.items():
                                if col_date.day in f_dias and col_date.month in meses_num and col_date.year in f_anio_list:
                                    val_diario = parse_latam_number(df_raw_tr.iloc[i, col_idx])
                                    all_data_trenes.append({
                                        "Tren": cell_a, "Fecha": col_date.strftime('%d/%m/%Y'),
                                        "Día": col_date.day, "Kilometraje Diario [km]": val_diario, "Timestamp": col_date
                                    })
        except: continue

# Procesar archivos Energía SEAT
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
                        tot = parse_latam_number(df_raw_s.iloc[i, 3])
                        trac = parse_latam_number(df_raw_s.iloc[i, 5])
                        kv12 = parse_latam_number(df_raw_s.iloc[i, 7])
                        all_data_seat.append({
                            "Fecha": fch_s.strftime('%d/%m/%Y'),
                            "Total [kWh]": tot,
                            "Tracción [kWh]": trac,
                            "12 KV [kWh]": kv12,
                            "% Tracción": (trac / tot * 100) if tot > 0 else 0,
                            "% 12 KV": (kv12 / tot * 100) if tot > 0 else 0,
                            "Timestamp": fch_s
                        })
        except: continue

# --- 5. RENDERIZADO ---
if all_data_ops or all_data_seat:
    df_ops_final = pd.DataFrame(all_data_ops).drop_duplicates(subset=['Fecha'], keep='last').sort_values("Timestamp") if all_data_ops else pd.DataFrame()
    df_trenes_final = pd.DataFrame(all_data_trenes).sort_values(["Timestamp", "Tren"]) if all_data_trenes else pd.DataFrame()
    df_seat_final = pd.DataFrame(all_data_seat).drop_duplicates(subset=['Fecha']).sort_values("Timestamp") if all_data_seat else pd.DataFrame()

    tab_resumen, tab_datos, tab_trenes, tab_seat = st.tabs(["📊 Resumen", "📑 Datos operacionales", "📑 Odómetro por Tren", "⚡ Energía SEAT"])
    
    with tab_resumen:
        if not df_ops_final.empty:
            st.subheader("Indicadores Globales")
            c1, c2, c3 = st.columns(3)
            tot_o, tot_t = df_ops_final["Odómetro [km]"].sum(), df_ops_final["Tren-Km [km]"].sum()
            c1.metric("Odómetro Total", f"{tot_o:,.1f} km")
            c2.metric("Tren-Km Total", f"{tot_t:,.1f} km")
            c3.metric("UMR Global", f"{(tot_t / tot_o * 100 if tot_o > 0 else 0):.2f} %")
            st.divider()
            st.write("### Resumen por Tipo de Jornada")
            df_ops_final['Tipo Día'] = pd.Categorical(df_ops_final['Tipo Día'], categories=['L', 'S', 'D/F'], ordered=True)
            resumen_tipo = df_ops_final.groupby("Tipo Día").agg({"Odómetro [km]": "sum", "Tren-Km [km]": "sum", "UMR [%]": "mean"}).reset_index()
            st.table(resumen_tipo.style.format({"Odómetro [km]": "{:,.1f}", "Tren-Km [km]": "{:,.1f}", "UMR [%]": "{:.2f}%"}))
        else:
            st.info("Cargue archivos UMR para ver el resumen.")

    with tab_datos:
        if not df_ops_final.empty:
            st.subheader("Detalle Cronológico Operacional")
            cols_v = ["Fecha", "Tipo Día", "N° Semana", "Odómetro [km]", "Tren-Km [km]", "UMR [%]"]
            st.dataframe(df_ops_final[cols_v].style.format({"Odómetro [km]": "{:,.1f}", "Tren-Km [km]": "{:,.1f}", "UMR [%]": "{:.2f}%"}).applymap(color_umr, subset=['UMR [%]']), use_container_width=True)

    with tab_trenes:
        if not df_trenes_final.empty:
            st.write("### 📏 Kilometraje Acumulado por Unidad")
            res_acum = df_trenes_final.groupby("Tren")["Kilometraje Diario [km]"].sum().reset_index()
            st.dataframe(res_acum.style.format({"Kilometraje Diario [km]": "{:,.1f}"}), use_container_width=True)
            st.divider()
            st.write("### 📅 Matriz de Kilometraje Diario")
            df_pivot = df_trenes_final.pivot_table(index="Tren", columns="Día", values="Kilometraje Diario [km]", aggfunc='sum').fillna(0)
            st.dataframe(df_pivot.style.format("{:,.1f}"), use_container_width=True)
        else:
            st.info("No se encontraron datos individuales de trenes.")

    with tab_seat:
        if not df_seat_final.empty:
            st.subheader("Consumo de Energía SEAT")
            
            # Cálculos globales para métricas
            g_tot = df_seat_final['Total [kWh]'].sum()
            g_trac = df_seat_final['Tracción [kWh]'].sum()
            g_kv12 = df_seat_final['12 KV [kWh]'].sum()
            
            e1, e2, e3 = st.columns(3)
            e1.metric("Consumo Total", f"{g_tot:,.0f} kWh")
            e2.metric("Tracción", f"{g_trac:,.0f} kWh", f"{(g_trac/g_tot*100 if g_tot>0 else 0):.1f}% del total")
            e3.metric("12 KV", f"{g_kv12:,.0f} kWh", f"{(g_kv12/g_tot*100 if g_tot>0 else 0):.1f}% del total")
            
            st.divider()
            st.dataframe(
                df_seat_final[["Fecha", "Total [kWh]", "Tracción [kWh]", "% Tracción", "12 KV [kWh]", "% 12 KV"]].style.format({
                    "Total [kWh]": "{:,.0f}", 
                    "Tracción [kWh]": "{:,.0f}", 
                    "12 KV [kWh]": "{:,.0f}",
                    "% Tracción": "{:.2f}%",
                    "% 12 KV": "{:.2f}%"
                }), use_container_width=True
            )
        else:
            st.info("Suba archivos de energía SEAT para ver el análisis.")

    # Botón de descarga
    st.sidebar.download_button("📥 Descargar Reporte Completo", to_excel_consolidado(df_ops_final, df_trenes_final, df_seat_final), "Reporte_SGE_Consolidado.xlsx")
else:
    st.info("👋 Sube los archivos UMR o Energía SEAT para comenzar.")
