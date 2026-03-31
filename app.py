import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
from datetime import datetime

# --- 1. CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="EFE Valparaíso - Reporte Multi-Periodo", layout="wide", page_icon="🚆")

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
    elif ',' in s:
        s = s.replace(',', '.')
    try: return float(s)
    except: return 0.0

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Reporte_Consolidado')
    return output.getvalue()

# --- 3. SIDEBAR (FILTROS MULTI-ELECCIÓN) ---
MESES_NOMBRES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
MESES_NUM_MAP = {nombre: i+1 for i, nombre in enumerate(MESES_NOMBRES)}

with st.sidebar:
    st.header("📂 Carga de Archivos")
    f_umr_list = st.file_uploader("Subir archivos Excel de UMR", type=["xlsx"], accept_multiple_files=True)
    
    st.divider()
    st.subheader("📅 Filtros de Periodo")
    # Multiselección de Años
    f_anio_list = st.multiselect("Seleccionar Años", [2024, 2025, 2026], default=[datetime.now().year])
    
    # Multiselección de Meses
    f_mes_list = st.multiselect("Seleccionar Meses", MESES_NOMBRES, default=[MESES_NOMBRES[datetime.now().month - 1]])
    meses_seleccionados_num = [MESES_NUM_MAP[m] for m in f_mes_list]
    
    # Multiselección de Días
    f_dias = st.multiselect("Seleccionar Días", list(range(1, 32)), default=list(range(1, 32)))

# --- 4. PROCESAMIENTO ---
if f_umr_list and f_anio_list and f_mes_list:
    all_res_diario = []
    
    with st.spinner("Consolidando periodos seleccionados..."):
        for f in f_umr_list:
            try:
                xl = pd.ExcelFile(f)
                sn_umr = next((s for s in xl.sheet_names if 'UMR' in s.upper() and 'RESUMEN' in s.upper()), None)
                
                if sn_umr:
                    df_raw = pd.read_excel(f, sheet_name=sn_umr, header=None)
                    
                    # Buscador de cabeceras
                    hdr_row = None
                    for i in range(min(100, len(df_raw))):
                        fila_scan = " ".join(df_raw.iloc[i].astype(str)).upper()
                        if 'ODO' in fila_scan or 'FECHA' in fila_scan:
                            hdr_row = i
                            break
                    
                    if hdr_row is not None:
                        cols_str = [str(c).strip().upper().replace('Ó','O').replace('Á','A') for c in df_raw.iloc[hdr_row]]
                        
                        idx_fecha = next(i for i, c in enumerate(cols_str) if 'FECHA' in c)
                        idx_odo   = next(i for i, c in enumerate(cols_str) if 'ODO' in c and 'ACUM' not in c)
                        idx_tkm   = next(i for i, c in enumerate(cols_str) if 'TREN' in c and 'KM' in c and 'ACUM' not in c)

                        df_data = df_raw.iloc[hdr_row + 1:].copy()
                        df_data['_dt'] = pd.to_datetime(df_data.iloc[:, idx_fecha], errors='coerce')
                        
                        # Filtro por listas de Año, Mes y Día
                        mask = (
                            (df_data['_dt'].dt.day.isin(f_dias)) & 
                            (df_data['_dt'].dt.month.isin(meses_seleccionados_num)) & 
                            (df_data['_dt'].dt.year.isin(f_anio_list))
                        )
                        row_found = df_data[mask]
                        
                        for _, row in row_found.iterrows():
                            v_odo = parse_latam_number(row.iloc[idx_odo])
                            v_tkm = parse_latam_number(row.iloc[idx_tkm])
                            # Ecuación: (Tren-Km / Odómetro) * 100
                            v_umr_calc = (v_tkm / v_odo * 100) if v_odo > 0 else 0
                            
                            all_res_diario.append({
                                "Fecha": row.iloc[idx_fecha].strftime('%d/%m/%Y') if isinstance(row.iloc[idx_fecha], datetime) else str(row.iloc[idx_fecha]),
                                "Timestamp": row.iloc[idx_fecha],
                                "Odómetro [km]": v_odo,
                                "Tren-Km [km]": v_tkm,
                                "UMR [%]": v_umr_calc,
                                "Archivo": f.name
                            })
            except Exception as e:
                st.error(f"Error en archivo {f.name}: {e}")

    if all_res_diario:
        df_final = pd.DataFrame(all_res_diario).drop_duplicates(subset=['Fecha'], keep='last').sort_values("Timestamp")
        
        # --- 5. RENDERIZADO ---
        st.write(f"## 📊 Análisis Consolidado")
        st.info(f"Periodos activos: {', '.join(f_mes_list)} de {', '.join(map(str, f_anio_list))}")
        
        c1, c2, c3 = st.columns(3)
        t_odo = df_final["Odómetro [km]"].sum()
        t_tkm = df_final["Tren-Km [km]"].sum()
        prom_umr = (t_tkm / t_odo * 100) if t_odo > 0 else 0
        
        c1.metric("Odómetro Total", f"{t_odo:,.1f} km")
        c2.metric("Tren-Km Total", f"{t_tkm:,.1f} km")
        c3.metric("UMR Global (Calculado)", f"{prom_umr:.2f} %")
        
        st.divider()
        st.subheader("Detalle por Jornada")
        st.dataframe(
            df_final[["Fecha", "Odómetro [km]", "Tren-Km [km]", "UMR [%]", "Archivo"]].style.format({
                "Odómetro [km]": "{:,.1f}",
                "Tren-Km [km]": "{:,.1f}",
                "UMR [%]": "{:.2f}%"
            }), use_container_width=True
        )
        st.download_button("📥 Descargar Reporte Completo (Excel)", to_excel(df_final), "Consolidado_EFE.xlsx")
    else:
        st.warning("No se encontraron registros para los criterios de búsqueda seleccionados.")
else:
    st.info("👋 Selecciona al menos un archivo, un mes y un año para generar el reporte.")
