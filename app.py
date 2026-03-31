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

# Estilo para métricas en la pestaña Resumen
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

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Datos_Operacionales')
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
    f_mes_list = st.multiselect("Meses", MESES_NOMBRES, default=[MESES_NOMBRES[datetime.now().month - 1]])
    meses_num = [MESES_NUM_MAP[m] for m in f_mes_list]
    f_dias = st.multiselect("Días", list(range(1, 32)), default=list(range(1, 32)))

# --- 4. PROCESAMIENTO ---
if f_umr_list:
    all_data = []
    
    for f in f_umr_list:
        try:
            xl = pd.ExcelFile(f)
            sn = next((s for s in xl.sheet_names if 'UMR' in s.upper() and 'RESUMEN' in s.upper()), None)
            if not sn: continue
            
            df_raw = pd.read_excel(f, sheet_name=sn, header=None)
            
            # Buscador de fila de títulos
            hdr_row = None
            for i in range(min(100, len(df_raw))):
                fila_txt = " ".join(df_raw.iloc[i].astype(str)).upper()
                if ('ODO' in fila_txt or 'FECHA' in fila_txt) and 'TREN' in fila_txt:
                    hdr_row = i
                    break
            
            if hdr_row is None: continue

            cols_orig = df_raw.iloc[hdr_row].astype(str).tolist()
            cols_clean = [re.sub(r'[^A-Z]', '', c.upper().replace('Ó','O').replace('Á','A')) for c in cols_orig]
            
            def find_idx(aliases, clean_list, orig_list):
                for i, c in enumerate(clean_list):
                    if 'ACUM' in orig_list[i].upper(): continue
                    if any(a in c for a in aliases): return i
                return None

            idx_fch = find_idx(['FECHA', 'FCH', 'DATE'], cols_clean, cols_orig)
            idx_odo = find_idx(['ODO', 'METRO', 'KM', 'KILO'], cols_clean, cols_orig)
            idx_tkm = find_idx(['TRENKM', 'TK', 'TRKM', 'KMTR'], cols_clean, cols_orig)

            if idx_odo is None and idx_fch is not None: idx_odo = idx_fch + 1
            if None in [idx_fch, idx_odo, idx_tkm]: continue

            df_extracted = df_raw.iloc[hdr_row + 1:].copy()
            df_extracted['_dt'] = pd.to_datetime(df_extracted.iloc[:, idx_fch], errors='coerce')
            
            mask = (df_extracted['_dt'].dt.day.isin(f_dias)) & \
                   (df_extracted['_dt'].dt.month.isin(meses_num)) & \
                   (df_extracted['_dt'].dt.year.isin(f_anio_list))
            
            rows = df_extracted[mask]
            
            for _, row in rows.iterrows():
                fch = row.iloc[idx_fch]
                if not isinstance(fch, (datetime, pd.Timestamp)): continue
                
                # Lógica Tipo Día
                es_festivo = fch in chile_holidays
                nom_dia = fch.strftime('%A')
                if es_festivo or nom_dia == 'Sunday': t_dia = "D/F"
                elif nom_dia == 'Saturday': t_dia = "S"
                else: t_dia = "L"
                
                o, t = parse_latam_number(row.iloc[idx_odo]), parse_latam_number(row.iloc[idx_tkm])
                
                all_data.append({
                    "Fecha": fch.strftime('%d/%m/%Y'),
                    "Tipo Día": t_dia,
                    "N° Semana": fch.isocalendar()[1],
                    "Odómetro [km]": o,
                    "Tren-Km [km]": t,
                    "UMR [%]": (t / o * 100) if o > 0 else 0,
                    "Timestamp": fch,
                    "Archivo": f.name
                })
        except: continue

    if all_data:
        df_final = pd.DataFrame(all_data).drop_duplicates(subset=['Fecha'], keep='last').sort_values("Timestamp")
        
        # --- 5. CREACIÓN DE PESTAÑAS ---
        tab_resumen, tab_datos = st.tabs(["📊 Resumen", "📑 Datos operacionales"])
        
        with tab_resumen:
            st.subheader("Indicadores Globales")
            col1, col2, col3 = st.columns(3)
            tot_odo = df_final["Odómetro [km]"].sum()
            tot_tkm = df_final["Tren-Km [km]"].sum()
            umr_glob = (tot_tkm / tot_odo * 100) if tot_odo > 0 else 0
            
            col1.metric("Odómetro Total", f"{tot_odo:,.1f} km")
            col2.metric("Tren-Km Total", f"{tot_tkm:,.1f} km")
            col3.metric("UMR Global", f"{umr_glob:.2f} %")
            
            st.divider()
            # Mini resumen por tipo de día
            st.write("### Promedio UMR por Tipo de Día")
            resumen_dia = df_final.groupby("Tipo Día")["UMR [%]"].mean().reset_index()
            st.table(resumen_dia.style.format({"UMR [%]": "{:.2f}%"}))

        with tab_datos:
            st.subheader("Detalle Cronológico de Operación")
            
            # Reordenar columnas según pedido
            cols_vista = ["Fecha", "Tipo Día", "N° Semana", "Odómetro [km]", "Tren-Km [km]", "UMR [%]"]
            
            styled_df = df_final[cols_vista].style.format({
                "Odómetro [km]": "{:,.1f}",
                "Tren-Km [km]": "{:,.1f}",
                "UMR [%]": "{:.2f}%"
            }).applymap(color_umr, subset=['UMR [%]'])
            
            st.dataframe(styled_df, use_container_width=True)
            
            st.download_button("📥 Descargar Datos Operacionales (Excel)", 
                             to_excel(df_final[cols_vista]), 
                             "Datos_Operacionales_EFE.xlsx")
    else:
        st.warning("No se encontraron registros coincidentes.")
else:
    st.info("👋 Sube los archivos UMR para generar el reporte por pestañas.")
