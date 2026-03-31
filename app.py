import streamlit as st
import pandas as pd
import numpy as np
import re
import holidays
from io import BytesIO
from datetime import datetime

# --- 1. CONFIGURACIÓN ---
st.set_page_config(page_title="EFE Valparaíso - Control Operacional UMR", layout="wide", page_icon="🚆")

# Configuración de feriados de Chile
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

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Reporte_EFE_SGE')
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
    f_umr_list = st.file_uploader("Subir archivos UMR", type=["xlsx"], accept_multiple_files=True)
    st.divider()
    f_anio_list = st.multiselect("Años", [2024, 2025, 2026], default=[2025, 2026])
    f_mes_list = st.multiselect("Meses", MESES_NOMBRES, default=MESES_NOMBRES)
    meses_seleccionados_num = [MESES_NUM_MAP[m] for m in f_mes_list]
    f_dias = st.multiselect("Días", list(range(1, 32)), default=list(range(1, 32)))

# --- 4. PROCESAMIENTO ---
if f_umr_list:
    all_res_diario = []
    
    for f in f_umr_list:
        try:
            xl = pd.ExcelFile(f)
            sn_umr = next((s for s in xl.sheet_names if 'UMR' in s.upper() and 'RESUMEN' in s.upper()), None)
            if not sn_umr: continue
            
            df_raw = pd.read_excel(f, sheet_name=sn_umr, header=None)
            
            # Buscador robusto de fila de títulos
            hdr_row = None
            for i in range(min(100, len(df_raw))):
                fila_txt = " ".join(df_raw.iloc[i].astype(str)).upper()
                if ('ODO' in fila_txt or 'FECHA' in fila_txt) and 'TREN' in fila_txt:
                    hdr_row = i
                    break
            
            if hdr_row is None: continue

            # Limpieza de nombres de columnas y detección por alias
            cols_orig = df_raw.iloc[hdr_row].astype(str).tolist()
            cols_clean = [re.sub(r'[^A-Z]', '', c.upper().replace('Ó','O').replace('Á','A')) for c in cols_orig]
            
            def encontrar_col(alias_list, lista_limpia, lista_orig):
                for i, col in enumerate(lista_limpia):
                    if 'ACUM' in lista_orig[i].upper(): continue
                    if any(a in col for a in alias_list): return i
                return None

            idx_fch = encontrar_col(['FECHA', 'FCH', 'DATE'], cols_clean, cols_orig)
            idx_odo = encontrar_col(['ODO', 'METRO', 'KM', 'KILO'], cols_clean, cols_orig)
            idx_tkm = encontrar_col(['TRENKM', 'TK', 'TRKM', 'KMTR'], cols_clean, cols_orig)

            # Fallback si el Odómetro no se encuentra por nombre
            if idx_odo is None and idx_fch is not None: idx_odo = idx_fch + 1

            if None in [idx_fch, idx_odo, idx_tkm]: continue

            df_data = df_raw.iloc[hdr_row + 1:].copy()
            df_data['_dt'] = pd.to_datetime(df_data.iloc[:, idx_fch], errors='coerce')
            
            mask = (
                (df_data['_dt'].dt.day.isin(f_dias)) & 
                (df_data['_dt'].dt.month.isin(meses_seleccionados_num)) & 
                (df_data['_dt'].dt.year.isin(f_anio_list))
            )
            row_found = df_data[mask]
            
            for _, row in row_found.iterrows():
                fecha_dt = row.iloc[idx_fch]
                if not isinstance(fecha_dt, (datetime, pd.Timestamp)): continue
                
                v_odo = parse_latam_number(row.iloc[idx_odo])
                v_tkm = parse_latam_number(row.iloc[idx_tkm])
                v_umr = (v_tkm / v_odo * 100) if v_odo > 0 else 0
                
                # --- NUEVA LÓGICA DE CLASIFICACIÓN DE DÍA ---
                es_festivo = fecha_dt in chile_holidays
                dia_semana_en = fecha_dt.strftime('%A') # Monday, Tuesday...
                
                if es_festivo or dia_semana_en == 'Sunday':
                    tipo_dia = "D/F" # Domingo o Festivo
                elif dia_semana_en == 'Saturday':
                    tipo_dia = "S"   # Sábado
                else:
                    tipo_dia = "L"   # Lunes a Viernes (Laboral)
                
                all_res_diario.append({
                    "Fecha": fecha_dt.strftime('%d/%m/%Y'),
                    "Tipo Día": tipo_dia,
                    "Timestamp": fecha_dt,
                    "Odómetro [km]": v_odo,
                    "Tren-Km [km]": v_tkm,
                    "UMR [%]": v_umr,
                    "Archivo": f.name
                })
        except: continue

    if all_res_diario:
        df_final = pd.DataFrame(all_res_diario).drop_duplicates(subset=['Fecha'], keep='last').sort_values("Timestamp")
        
        st.write(f"## 📊 Consolidado de Utilización de Flota (UMR)")
        
        # Métricas Globales
        c1, c2, c3 = st.columns(3)
        t_odo, t_tkm = df_final["Odómetro [km]"].sum(), df_final["Tren-Km [km]"].sum()
        u_global = (t_tkm / t_odo * 100) if t_odo > 0 else 0
        
        c1.metric("Odómetro Total", f"{t_odo:,.1f} km")
        c2.metric("Tren-Km Total", f"{t_tkm:,.1f} km")
        c3.metric("UMR Global", f"{u_global:.2f} %")
        
        st.divider()
        
        # Tabla con semáforo y nueva columna "Tipo Día"
        st.dataframe(
            df_final[["Fecha", "Tipo Día", "Odómetro [km]", "Tren-Km [km]", "UMR [%]", "Archivo"]]
            .style.format({"Odómetro [km]": "{:,.1f}", "Tren-Km [km]": "{:,.1f}", "UMR [%]": "{:.2f}%"})
            .applymap(color_umr, subset=['UMR [%]']), 
            use_container_width=True
        )
        
        st.download_button("📥 Descargar Reporte Consolidado", to_excel(df_final), "Consolidado_UMR_EFE.xlsx")
    else:
        st.warning("⚠️ No se encontraron registros. Verifica los filtros de Fecha y que los archivos tengan la hoja 'UMR Resumen'.")
