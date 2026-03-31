import streamlit as st
import pandas as pd
import numpy as np
import re
import holidays
from io import BytesIO
from datetime import datetime

# --- 1. CONFIGURACIÓN ---
st.set_page_config(page_title="EFE Valparaíso - Dashboard SGE", layout="wide", page_icon="🚆")
chile_holidays = holidays.Chile()

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

def color_umr(val):
    if val > 96.4: return 'color: green; font-weight: bold;'
    elif val < 96.4: return 'color: red; font-weight: bold;'
    return 'color: black;'

# --- 3. SIDEBAR ---
st.sidebar.header("📂 Gestión de Archivos")
f_umr_list = st.sidebar.file_uploader("Subir archivos UMR / Odómetros", type=["xlsx"], accept_multiple_files=True)
f_seat_list = st.sidebar.file_uploader("Subir archivos Energía SEAT", type=["xlsx"], accept_multiple_files=True)

MESES_NOMBRES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
st.sidebar.divider()
st.sidebar.subheader("📅 Filtros de Visualización")
# Aumentamos el rango de años por si hay datos históricos
sel_anios = st.sidebar.multiselect("Años", [2022, 2023, 2024, 2025, 2026], default=[2025, 2026])
sel_meses = st.sidebar.multiselect("Meses", MESES_NOMBRES, default=MESES_NOMBRES) # Por defecto todos para evitar tabla vacía
sel_meses_num = [MESES_NOMBRES.index(m) + 1 for m in sel_meses]

# --- 4. PROCESAMIENTO ---
all_resumen_raw = []
all_trenes_raw = []

if f_umr_list:
    for f in f_umr_list:
        try:
            xl = pd.ExcelFile(f)
            # MEJORA: Buscador de hojas más flexible (UMR o RESUMEN)
            sn_res = next((s for s in xl.sheet_names if 'UMR' in s.upper() or 'RESUMEN' in s.upper()), None)
            
            if sn_res:
                df_raw = pd.read_excel(f, sheet_name=sn_res, header=None)
                # Buscamos la fila de cabecera
                hdr_row = None
                for i in range(min(50, len(df_raw))):
                    txt = " ".join(df_raw.iloc[i].astype(str)).upper()
                    if 'FECHA' in txt and ('ODO' in txt or 'TREN' in txt):
                        hdr_row = i
                        break
                
                if hdr_row is not None:
                    cols = [str(c).upper() for c in df_raw.iloc[hdr_row]]
                    idx_fch = next((i for i, c in enumerate(cols) if 'FECHA' in c), None)
                    idx_odo = next((i for i, c in enumerate(cols) if 'ODO' in c and 'ACUM' not in c), None)
                    idx_tkm = next((i for i, c in enumerate(cols) if 'TREN' in c and 'KM' in c), None)
                    
                    if idx_fch is not None:
                        df_ext = df_raw.iloc[hdr_row+1:].copy()
                        df_ext['_dt'] = pd.to_datetime(df_ext.iloc[:, idx_fch], errors='coerce')
                        
                        # Procesar todas las filas y filtrar después
                        for _, row in df_ext.dropna(subset=['_dt']).iterrows():
                            fch = row.iloc[idx_fch]
                            if fch.year in sel_anios and fch.month in sel_meses_num:
                                o, t = parse_latam_number(row.iloc[idx_odo]), parse_latam_number(row.iloc[idx_tkm])
                                all_resumen_raw.append({
                                    "Fecha_DT": fch, "Fecha": fch.strftime('%d/%m/%Y'),
                                    "Año": fch.year, "Mes": fch.month,
                                    "Odómetro [km]": o, "Tren-Km [km]": t, "UMR [%]": (t/o*100 if o>0 else 0)
                                })
            else:
                st.sidebar.error(f"No se halló hoja UMR en {f.name}")
        except Exception as e:
            st.sidebar.error(f"Error en {f.name}: {e}")

# (El resto del procesamiento de Energía y Trenes se mantiene igual...)

# --- 5. RENDERIZADO ---
if all_resumen_raw:
    df_res = pd.DataFrame(all_resumen_raw).drop_duplicates(subset=['Fecha']).sort_values("Fecha_DT")
    
    t_res, t_datos, t_trenes, t_seat = st.tabs(["📊 Resumen", "📑 Datos operacionales", "📑 Odómetro por Tren", "⚡ Energía SEAT"])

    with t_res:
        # Lógica de Tipo Día con orden L -> S -> D/F
        def get_tipo_dia(fch):
            nom_dia = fch.strftime('%A')
            if fch in chile_holidays or nom_dia == 'Sunday': return "D/F"
            return "S" if nom_dia == 'Saturday' else "L"
        
        df_res['Tipo Día'] = df_res['Fecha_DT'].apply(get_tipo_dia)
        df_res['Tipo Día'] = pd.Categorical(df_res['Tipo Día'], categories=['L', 'S', 'D/F'], ordered=True)
        
        c1, c2, c3 = st.columns(3)
        tot_o, tot_t = df_res["Odómetro [km]"].sum(), df_res["Tren-Km [km]"].sum()
        c1.metric("Odómetro Total", f"{tot_o:,.1f} km")
        c2.metric("Tren-Km Total", f"{tot_t:,.1f} km")
        c3.metric("UMR Global", f"{(tot_t/tot_o*100 if tot_o>0 else 0):.2f} %")
        
        st.write("### Resumen por Jornada")
        res_t = df_res.groupby("Tipo Día").agg({"Odómetro [km]": "sum", "Tren-Km [km]": "sum", "UMR [%]": "mean"}).reset_index()
        st.table(res_t.style.format({"Odómetro [km]": "{:,.1f}", "Tren-Km [km]": "{:,.1f}", "UMR [%]": "{:.2f}%"}))
    
    # ... (Resto de pestañas)
else:
    if f_umr_list:
        st.warning("⚠️ Archivos cargados pero no hay datos para los años/meses seleccionados en el Sidebar.")
        st.info("Revisa que los filtros de 'Años' y 'Meses' incluyan las fechas de tus archivos.")
