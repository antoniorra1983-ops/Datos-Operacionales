import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
from datetime import datetime

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="EFE Valparaíso - Módulo UMR", layout="wide", page_icon="🚆")

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

# --- UTILIDADES ---
def parse_latam_number(val):
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    s = str(val).strip().replace(' ', '').replace('$', '')
    # Limpieza de caracteres no numéricos excepto separadores
    s = re.sub(r'[^\d.,-]', '', s)
    if not s: return 0.0
    # Manejo de formatos 1.234,56
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
        df.to_excel(writer, index=False, sheet_name='Resumen_UMR')
    return output.getvalue()

# --- SIDEBAR ---
with st.sidebar:
    st.header("📂 Carga de Archivo")
    f_umr = st.file_uploader("Subir Excel de Odómetros (UMR)", type=["xlsx"])
    st.divider()
    f_anio = st.selectbox("Año de Operación", [2025, 2026], index=1)
    meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    f_mes = st.selectbox("Mes de Operación", meses, index=datetime.now().month - 1)
    mes_num = meses.index(f_mes) + 1
    f_dias = st.multiselect("Días a visualizar", list(range(1, 32)), default=[datetime.now().day])

# --- LÓGICA UMR ---
if f_umr:
    try:
        xl = pd.ExcelFile(f_umr)
        # 1. Buscar la hoja correcta
        sn_umr = next((s for s in xl.sheet_names if 'UMR' in s.upper() and 'RESUMEN' in s.upper()), None)
        
        if sn_umr:
            df_raw = pd.read_excel(f_umr, sheet_name=sn_umr, header=None)
            
            # === MEJORA: SÚPER-BUSCADOR DE CABECERAS ===
            hdr_row = None
            for i in range(min(100, len(df_raw))):
                fila_scan = " ".join(df_raw.iloc[i].astype(str).upper())
                if 'ODO' in fila_scan or 'FECHA' in fila_scan or 'TREN' in fila_scan:
                    hdr_row = i
                    break
            
            if hdr_row is not None:
                cols_str = [str(c).strip().upper().replace('Ó','O').replace('Á','A').replace('É','E') 
                            for c in df_raw.iloc[hdr_row]]
                
                # Mapeo de columnas con tolerancia a cambios de nombre
                try:
                    idx_fecha = next(i for i, c in enumerate(cols_str) if 'FECHA' in c)
                    idx_odo   = next(i for i, c in enumerate(cols_str) if 'ODO' in c and 'ACUM' not in c)
                    idx_tkm   = next(i for i, c in enumerate(cols_str) if 'TREN' in c and 'KM' in c and 'ACUM' not in c)
                except StopIteration:
                    st.error(f"⚠️ No encontré las columnas exactas. Columnas detectadas: {cols_str}")
                    st.stop()

                # Extraer datos reales
                df_data = df_raw.iloc[hdr_row + 1:].copy()
                df_data['_dt'] = pd.to_datetime(df_data.iloc[:, idx_fecha], errors='coerce')
                
                res_diario = []
                for d in f_dias:
                    mask = (df_data['_dt'].dt.day == d) & (df_data['_dt'].dt.month == mes_num) & (df_data['_dt'].dt.year == f_anio)
                    row_ur = df_data[mask]
                    
                    if not row_ur.empty:
                        v_odo = parse_latam_number(row_ur.iloc[0, idx_odo])
                        v_tkm = parse_latam_number(row_ur.iloc[0, idx_tkm])
                        
                        # ECUACIÓN SOLICITADA: (Tren-Km / Odómetro) * 100
                        v_umr_calc = (v_tkm / v_odo * 100) if v_odo > 0 else 0
                        
                        res_diario.append({
                            "Fecha": f"{d:02d}/{mes_num:02d}/{f_anio}",
                            "Odómetro [km]": v_odo,
                            "Tren-Km [km]": v_tkm,
                            "UMR [%]": v_umr_calc
                        })

                df_final = pd.DataFrame(res_diario)

                if not df_final.empty:
                    st.write(f"## 📊 Análisis Operacional EFE - {f_mes} {f_anio}")
                    
                    # Métricas de la parte superior
                    c1, c2, c3 = st.columns(3)
                    t_odo = df_final["Odómetro [km]"].sum()
                    t_tkm = df_final["Tren-Km [km]"].sum()
                    prom_umr = (t_tkm / t_odo * 100) if t_odo > 0 else 0
                    
                    c1.metric("Odómetro Total", f"{t_odo:,.1f} km")
                    c2.metric("Tren-Km Total", f"{t_tkm:,.1f} km")
                    c3.metric("UMR Global (Ecuación)", f"{prom_umr:.2f} %")
                    
                    st.divider()
                    st.subheader("Desglose Diario")
                    st.dataframe(
                        df_final.style.format({
                            "Odómetro [km]": "{:,.1f}",
                            "Tren-Km [km]": "{:,.1f}",
                            "UMR [%]": "{:.2f}%"
                        }), 
                        use_container_width=True
                    )
                    
                    st.download_button("📥 Descargar Reporte UMR (Excel)", to_excel(df_final), f"Reporte_UMR_{f_mes}.xlsx")
                else:
                    st.warning(f"No hay datos registrados para los días seleccionados en {f_mes} {f_anio}.")
            else:
                st.error("No logré encontrar la fila de encabezados. Asegúrate de que existan las palabras 'Fecha' u 'Odómetro'.")
        else:
            st.error("No se encontró la hoja 'UMR Resumen'. Revisa que el nombre coincida exactamente.")
            
    except Exception as e:
        st.error(f"💥 Error al procesar el archivo: {e}")
else:
    st.info("👋 Por favor, sube el archivo Excel de Odómetros (UMR) para visualizar el análisis.")
