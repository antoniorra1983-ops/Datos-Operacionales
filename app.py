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

def color_umr(val):
    if val > 96.4: return 'color: green; font-weight: bold;'
    elif val < 96.4: return 'color: red; font-weight: bold;'
    return 'color: black;'

# --- 3. SIDEBAR (FILTROS) ---
st.sidebar.header("📂 Gestión de Archivos")
f_umr_list = st.sidebar.file_uploader("Subir archivos Excel (UMR/Odómetro)", type=["xlsx"], accept_multiple_files=True)

# Definimos periodos por defecto por si aún no hay archivos
MESES_NOMBRES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
st.sidebar.divider()
st.sidebar.subheader("📅 Filtros de Periodo")
sel_anios = st.sidebar.multiselect("Seleccionar Año", [2024, 2025, 2026], default=[2025, 2026])
sel_meses_nombres = st.sidebar.multiselect("Seleccionar Mes", MESES_NOMBRES, default=[MESES_NOMBRES[datetime.now().month - 1]])
sel_meses_num = [MESES_NOMBRES.index(m) + 1 for m in sel_meses_nombres]

# --- 4. PROCESAMIENTO ---
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
                            if fch.year in sel_anios and fch.month in sel_meses_num:
                                o, t = parse_latam_number(row.iloc[idx_odo]), parse_latam_number(row.iloc[idx_tkm])
                                all_resumen_raw.append({
                                    "Fecha_DT": fch, "Fecha": fch.strftime('%d/%m/%Y'),
                                    "Año": fch.year, "Mes": fch.month, "Día": fch.day,
                                    "Odómetro [km]": o, "Tren-Km [km]": t, "UMR [%]": (t/o*100 if o>0 else 0)
                                })

            # --- PARTE B: ODOMETRO POR TREN (TRIPLE ENCABEZADO) ---
            sn_tren = next((s for s in xl.sheet_names if 'ODO' in s.upper() and 'KIL' in s.upper()), None)
            if sn_tren:
                df_tr_raw = pd.read_excel(f, sheet_name=sn_tren, header=None)
                
                # Buscamos la fila de FECHA
                date_row = None
                col_map = {}
                for i in range(min(50, len(df_tr_raw))):
                    # Intentamos encontrar una fila que tenga al menos una fecha válida
                    row_content = df_tr_raw.iloc[i]
                    for c_idx, val in enumerate(row_content):
                        f_parsed = pd.to_datetime(val, errors='coerce')
                        if pd.notna(f_parsed) and f_parsed.year > 2000:
                            date_row = i
                            col_map[c_idx] = f_parsed
                    if date_row is not None: break

                if date_row is not None:
                    # Buscamos la fila donde empiezan los nombres de los trenes (Columna A)
                    # Bajamos desde date_row buscando "M01" o "Tren"
                    start_trains_row = next((r for r in range(date_row, len(df_tr_raw)) if re.match(r'^(M\d|XM\d|TREN)', str(df_tr_raw.iloc[r, 0]).upper())), date_row + 3)
                    
                    for r_idx in range(start_trains_row, len(df_tr_raw)):
                        nombre = str(df_tr_raw.iloc[r_idx, 0]).strip().upper()
                        if re.match(r'^(M\d{1,2}|XM\d{1,2})', nombre):
                            for c_idx, f_dt in col_map.items():
                                if f_dt.year in sel_anios and f_dt.month in sel_meses_num:
                                    val_km = parse_latam_number(df_tr_raw.iloc[r_idx, c_idx])
                                    all_trenes_raw.append({
                                        "Tren": nombre, "Fecha_DT": f_dt, "Kilometraje": val_km,
                                        "Año": f_dt.year, "Mes": f_dt.month, "Día": f_dt.day
                                    })
        except Exception as e:
            st.sidebar.warning(f"Aviso en {f.name}: {e}")

# --- 5. RENDERIZADO DE TABLAS ---
if all_resumen_raw or all_trenes_raw:
    df_res = pd.DataFrame(all_resumen_raw).drop_duplicates(subset=['Fecha']).sort_values("Fecha_DT") if all_resumen_raw else pd.DataFrame()
    df_tr = pd.DataFrame(all_trenes_raw) if all_trenes_raw else pd.DataFrame()

    t_res, t_datos, t_trenes = st.tabs(["📊 Resumen", "📑 Datos operacionales", "📑 Odómetro por Tren"])

    with t_res:
        if not df_res.empty:
            c1, c2, c3 = st.columns(3)
            tot_o, tot_t = df_res["Odómetro [km]"].sum(), df_res["Tren-Km [km]"].sum()
            c1.metric("Odómetro Total", f"{tot_o:,.1f} km")
            c2.metric("Tren-Km Total", f"{tot_t:,.1f} km")
            c3.metric("UMR Global", f"{(tot_t/tot_o*100 if tot_o>0 else 0):.2f} %")
        else:
            st.info("Selecciona los filtros correctos para ver el resumen.")

    with t_datos:
        if not df_res.empty:
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
            st.write("### 📏 Kilometraje Total Acumulado por Tren")
            res_sum = df_tr.groupby("Tren")["Kilometraje"].sum().reset_index().sort_values("Kilometraje", ascending=False)
            st.dataframe(res_sum.style.format({"Kilometraje": "{:,.1f}"}), use_container_width=True)
            
            st.divider()
            st.write("### 📅 Kilometrajes Diarios")
            pivot_diario = df_tr.pivot_table(index="Tren", columns="Día", values="Kilometraje", aggfunc='sum').fillna(0)
            st.dataframe(pivot_diario.style.format("{:,.1f}"), use_container_width=True)
        else:
            st.warning("No se encontraron datos de trenes. Verifica que los filtros de Año/Mes coincidan con el archivo.")
else:
    st.info("👋 Sube tus archivos Excel para activar el análisis.")
