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

# --- 3. PROCESAMIENTO (SIDEBAR SIEMPRE VISIBLE) ---
st.sidebar.header("📂 Gestión de Archivos")
f_umr_list = st.sidebar.file_uploader("Subir archivos Excel (UMR/Odómetro)", type=["xlsx"], accept_multiple_files=True)

all_resumen_raw = []
all_trenes_raw = []

if f_umr_list:
    for f in f_umr_list:
        try:
            xl = pd.ExcelFile(f)
            
            # --- PARTE A: UMR RESUMEN ---
            sn_res = next((s for s in xl.sheet_names if 'UMR' in s.upper() and 'RESUMEN' in s.upper()), None)
            if sn_res:
                df_raw = pd.read_excel(f, sheet_name=sn_res, header=None)
                # Buscamos la fila de títulos (ODO)
                hdr_row = next((i for i in range(min(50, len(df_raw))) if 'ODO' in " ".join(df_raw.iloc[i].astype(str)).upper()), None)
                if hdr_row is not None:
                    cols = df_raw.iloc[hdr_row].astype(str).tolist()
                    idx_fch = next((i for i, c in enumerate(cols) if 'FECHA' in c.upper()), None)
                    idx_odo = next((i for i, c in enumerate(cols) if 'ODO' in c.upper() and 'ACUM' not in c.upper()), None)
                    idx_tkm = next((i for i, c in enumerate(cols) if 'TREN' in c.upper() and 'KM' in c.upper() and 'ACUM' not in c.upper()), None)
                    
                    if idx_fch is not None:
                        df_ext = df_raw.iloc[hdr_row+1:].copy()
                        df_ext['_dt'] = pd.to_datetime(df_ext.iloc[:, idx_fch], errors='coerce')
                        for _, row in df_ext.dropna(subset=['_dt']).iterrows():
                            fch = row.iloc[idx_fch]
                            o, t = parse_latam_number(row.iloc[idx_odo]), parse_latam_number(row.iloc[idx_tkm])
                            all_resumen_raw.append({
                                "Fecha_DT": fch, "Fecha": fch.strftime('%d/%m/%Y'),
                                "Año": fch.year, "Mes": fch.month, "Día": fch.day,
                                "Odómetro [km]": o, "Tren-Km [km]": t, "UMR [%]": (t/o*100 if o>0 else 0)
                            })

            # --- PARTE B: ODOMETRO POR TREN (ESTRUCTURA DE TU IMAGEN) ---
            sn_tren = next((s for s in xl.sheet_names if 'ODO' in s.upper() and 'KIL' in s.upper()), None)
            if sn_tren:
                df_tr_raw = pd.read_excel(f, sheet_name=sn_tren, header=None)
                # Buscamos la celda que dice "Tren" (según tu imagen es la A5)
                row_tren = next((i for i in range(min(50, len(df_tr_raw))) if 'TREN' in str(df_tr_raw.iloc[i, 0]).upper()), None)
                
                if row_tren is not None:
                    # Según tu imagen, las fechas están 3 o 4 filas más arriba del primer tren
                    col_map = {}
                    for i_search in range(row_tren):
                        for c_idx, val in enumerate(df_tr_raw.iloc[i_search]):
                            f_parsed = pd.to_datetime(val, errors='coerce')
                            if pd.notna(f_parsed) and f_parsed.year > 2000:
                                col_map[c_idx] = f_parsed
                    
                    for r_idx in range(row_tren + 1, len(df_tr_raw)):
                        nombre = str(df_tr_raw.iloc[r_idx, 0]).strip().upper()
                        if re.match(r'^(M\d|XM\d)', nombre):
                            for c_idx, f_dt in col_map.items():
                                val_km = parse_latam_number(df_tr_raw.iloc[r_idx, c_idx])
                                all_trenes_raw.append({
                                    "Tren": nombre, "Fecha_DT": f_dt, "Kilometraje": val_km,
                                    "Año": f_dt.year, "Mes": f_dt.month, "Día": f_dt.day
                                })
        except Exception as e:
            st.sidebar.error(f"Error en {f.name}: {e}")

# --- 4. RENDERIZADO DE FILTROS Y TABLAS ---
if all_resumen_raw or all_trenes_raw:
    df_res_base = pd.DataFrame(all_resumen_raw).drop_duplicates(subset=['Fecha']) if all_resumen_raw else pd.DataFrame()
    df_tr_base = pd.DataFrame(all_trenes_raw) if all_trenes_raw else pd.DataFrame()

    # Combinar años y meses de ambas fuentes
    anios_totales = sorted(list(set(df_res_base['Año'].unique()) | set(df_tr_base['Año'].unique())))
    
    st.sidebar.divider()
    st.sidebar.subheader("📅 Filtros de Periodo")
    sel_anios = st.sidebar.multiselect("Seleccionar Año", anios_totales, default=anios_totales)
    
    MESES_NOMBRES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    sel_meses_nombres = st.sidebar.multiselect("Seleccionar Mes", MESES_NOMBRES, default=MESES_NOMBRES)
    sel_meses_num = [MESES_NOMBRES.index(m) + 1 for m in sel_meses_nombres]

    # Aplicar Filtros
    df_res = df_res_base[df_res_base['Año'].isin(sel_anios) & df_res_base['Mes'].isin(sel_meses_num)].sort_values("Fecha_DT") if not df_res_base.empty else pd.DataFrame()
    df_tr = df_tr_base[df_tr_base['Año'].isin(sel_anios) & df_tr_base['Mes'].isin(sel_meses_num)] if not df_tr_base.empty else pd.DataFrame()

    # --- TABS ---
    t_res, t_datos, t_trenes = st.tabs(["📊 Resumen", "📑 Datos operacionales", "📑 Odómetro por Tren"])

    with t_res:
        if not df_res.empty:
            c1, c2, c3 = st.columns(3)
            tot_o, tot_t = df_res["Odómetro [km]"].sum(), df_res["Tren-Km [km]"].sum()
            c1.metric("Odómetro Total", f"{tot_o:,.1f} km")
            c2.metric("Tren-Km Total", f"{tot_t:,.1f} km")
            c3.metric("UMR Global", f"{(tot_t/tot_o*100 if tot_o>0 else 0):.2f} %")
        else:
            st.info("No hay datos de Resumen para el periodo seleccionado.")

    with t_datos:
        if not df_res.empty:
            # Lógica Tipo Día
            def get_tipo_dia(fch):
                nom_dia = fch.strftime('%A')
                if fch in chile_holidays or nom_dia == 'Sunday': return "D/F"
                return "S" if nom_dia == 'Saturday' else "L"
            
            df_res['Tipo Día'] = df_res['Fecha_DT'].apply(get_tipo_dia)
            df_res['N° Semana'] = df_res['Fecha_DT'].dt.isocalendar().week
            
            st.dataframe(df_res[["Fecha", "Tipo Día", "N° Semana", "Odómetro [km]", "Tren-Km [km]", "UMR [%]"]]
                         .style.format({"Odómetro [km]": "{:,.1f}", "Tren-Km [km]": "{:,.1f}", "UMR [%]": "{:.2f}%"})
                         .applymap(color_umr, subset=['UMR [%]']), use_container_width=True)

    with t_trenes:
        if not df_tr.empty:
            pivot_tr = df_tr.pivot_table(index="Tren", columns="Día", values="Kilometraje", aggfunc='sum').fillna(0)
            st.write("### Kilometraje Diario por Unidad (M/XM)")
            st.dataframe(pivot_tr.style.format("{:,.1f}"), use_container_width=True)
        else:
            st.warning("No se encontraron datos individuales de trenes. Revisa la hoja Odómetro-Kilometraje.")
else:
    st.info("👋 Sube tus archivos Excel para activar los filtros y ver el análisis.")
