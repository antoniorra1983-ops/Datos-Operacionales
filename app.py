import streamlit as st
import pandas as pd
import numpy as np
import re
import holidays
from io import BytesIO
from datetime import datetime

# --- 1. CONFIGURACIÓN ---
st.set_page_config(page_title="EFE Valparaíso - Auditoría UMR", layout="wide", page_icon="🚆")

DIAS_MAP = {'Monday': 'Lun', 'Tuesday': 'Mar', 'Wednesday': 'Mié', 'Thursday': 'Jue', 'Friday': 'Vie', 'Saturday': 'Sáb', 'Sunday': 'Dom'}
chile_holidays = holidays.Chile()

st.markdown("""
    <style>
    .stMetric { background-color: #ffffff; padding: 20px; border-radius: 10px; border-left: 5px solid #005195; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    </style>
    """, unsafe_allow_html=True)

# --- 2. FUNCIONES DE LIMPIEZA ---
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
MESES_NOMBRES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
MESES_NUM_MAP = {nombre: i+1 for i, nombre in enumerate(MESES_NOMBRES)}

with st.sidebar:
    st.header("📂 Carga de Datos")
    f_umr_list = st.file_uploader("Subir archivos UMR", type=["xlsx"], accept_multiple_files=True)
    st.divider()
    f_anio_list = st.multiselect("Años", [2024, 2025, 2026], default=[2025, 2026])
    f_mes_list = st.multiselect("Meses", MESES_NOMBRES, default=[MESES_NOMBRES[datetime.now().month - 1]])
    meses_seleccionados_num = [MESES_NUM_MAP[m] for m in f_mes_list]
    f_dias = st.multiselect("Días", list(range(1, 32)), default=list(range(1, 32)))

# --- 4. PROCESAMIENTO ROBUSTO ---
if f_umr_list:
    all_res_diario = []
    
    for f in f_umr_list:
        try:
            xl = pd.ExcelFile(f)
            sn_umr = next((s for s in xl.sheet_names if 'UMR' in s.upper() and 'RESUMEN' in s.upper()), None)
            
            if not sn_umr:
                st.warning(f"⚠️ {f.name}: No se encontró la hoja 'UMR Resumen'.")
                continue
            
            df_raw = pd.read_excel(f, sheet_name=sn_umr, header=None)
            
            # Buscador de cabeceras mejorado
            hdr_row = None
            for i in range(min(100, len(df_raw))):
                fila_scan = " ".join(df_raw.iloc[i].astype(str)).upper()
                if 'ODO' in fila_scan or 'FECHA' in fila_scan:
                    hdr_row = i
                    break
            
            if hdr_row is None:
                st.warning(f"⚠️ {f.name}: No se detectó la fila de títulos.")
                continue

            cols_str = [str(c).strip().upper() for c in df_raw.iloc[hdr_row]]
            
            # Identificación flexible de columnas
            def find_col(keywords, columns):
                for i, c in enumerate(columns):
                    if all(k in c for k in keywords): return i
                return None

            idx_fecha = find_col(['FECHA'], cols_str)
            idx_odo = find_col(['ODO'], [c if 'ACUM' not in c else '' for c in cols_str])
            idx_tkm = find_col(['TREN', 'KM'], [c if 'ACUM' not in c else '' for c in cols_str])

            if None in [idx_fecha, idx_odo, idx_tkm]:
                missing = [k for k, v in {"Fecha": idx_fecha, "Odo": idx_odo, "Tren-Km": idx_tkm}.items() if v is None]
                st.error(f"❌ {f.name}: Faltan columnas: {', '.join(missing)}")
                continue

            df_data = df_raw.iloc[hdr_row + 1:].copy()
            df_data['_dt'] = pd.to_datetime(df_data.iloc[:, idx_fecha], errors='coerce')
            
            mask = (
                (df_data['_dt'].dt.day.isin(f_dias)) & 
                (df_data['_dt'].dt.month.isin(meses_seleccionados_num)) & 
                (df_data['_dt'].dt.year.isin(f_anio_list))
            )
            row_found = df_data[mask]
            
            for _, row in row_found.iterrows():
                fecha_dt = row.iloc[idx_fecha]
                if not isinstance(fecha_dt, (datetime, pd.Timestamp)): continue
                
                v_odo = parse_latam_number(row.iloc[idx_odo])
                v_tkm = parse_latam_number(row.iloc[idx_tkm])
                v_umr_calc = (v_tkm / v_odo * 100) if v_odo > 0 else 0
                
                nombre_dia_en = fecha_dt.strftime('%A')
                dia_abr = DIAS_MAP.get(nombre_dia_en, nombre_dia_en[:3])
                es_festivo = "SÍ" if fecha_dt in chile_holidays else "NO"
                
                all_res_diario.append({
                    "Fecha": fecha_dt.strftime('%d/%m/%Y'),
                    "Día": dia_abr,
                    "Festivo": es_festivo,
                    "Timestamp": fecha_dt,
                    "Odómetro [km]": v_odo,
                    "Tren-Km [km]": v_tkm,
                    "UMR [%]": v_umr_calc,
                    "Archivo": f.name
                })
        except Exception as e:
            st.error(f"💥 Error crítico en {f.name}: {str(e)}")

    if all_res_diario:
        df_final = pd.DataFrame(all_res_diario).drop_duplicates(subset=['Fecha'], keep='last').sort_values("Timestamp")
        
        st.write(f"## 📊 Consolidado Operacional")
        
        c1, c2, c3 = st.columns(3)
        t_odo, t_tkm = df_final["Odómetro [km]"].sum(), df_final["Tren-Km [km]"].sum()
        c1.metric("Odómetro Total", f"{t_odo:,.1f} km")
        c2.metric("Tren-Km Total", f"{t_tkm:,.1f} km")
        c3.metric("UMR Global", f"{(t_tkm/t_odo*100 if t_odo>0 else 0):.2f} %")
        
        st.divider()
        st.dataframe(
            df_final[["Fecha", "Día", "Festivo", "Odómetro [km]", "Tren-Km [km]", "UMR [%]", "Archivo"]]
            .style.format({"Odómetro [km]": "{:,.1f}", "Tren-Km [km]": "{:,.1f}", "UMR [%]": "{:.2f}%"})
            .applymap(color_umr, subset=['UMR [%]']), 
            use_container_width=True
        )
    else:
        st.warning("⚠️ No se encontraron registros que coincidan con los filtros de Mes/Año seleccionados.")
