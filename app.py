import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
from datetime import datetime

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="EFE Valparaíso - Sistema SGE", layout="wide", page_icon="🚆")

# Estilo para las métricas
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

# --- FUNCIONES DE APOYO ---
def parse_latam_number(val):
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    s = str(val).strip().replace(' ', '')
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
        df.to_excel(writer, index=False, sheet_name='Resumen_UMR')
    return output.getvalue()

# --- SIDEBAR ---
with st.sidebar:
    st.header("📂 Carga de Datos")
    f_umr = st.file_uploader("Subir Excel de Odómetros (UMR)", type=["xlsx"])
    
    st.divider()
    st.subheader("📅 Configuración")
    f_anio = st.selectbox("Año", [2025, 2026], index=1)
    meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    f_mes = st.selectbox("Mes", meses, index=datetime.now().month - 1)
    mes_num = meses.index(f_mes) + 1
    
    f_dias = st.multiselect("Días a visualizar", list(range(1, 32)), default=[datetime.now().day])

# --- PROCESAMIENTO ---
if f_umr:
    try:
        xl = pd.ExcelFile(f_umr)
        # Buscar la hoja de resumen
        sn_umr = next((s for s in xl.sheet_names if 'UMR' in s.upper() and 'RESUMEN' in s.upper()), None)
        
        if sn_umr:
            df_raw = pd.read_excel(f_umr, sheet_name=sn_umr, header=None)
            
            # Buscamos la fila de los títulos
            hdr_row = next((i for i in range(min(15, len(df_raw))) if any('ODO' in str(v).upper() for v in df_raw.iloc[i])), None)
            
            if hdr_row is None:
                st.error("❌ No logré encontrar la fila con los títulos (Odómetro, Tren-Km). Revisa el formato del Excel.")
            else:
                cols = [str(c).strip().upper().replace('Ó','O').replace('Á','A') for c in df_raw.iloc[hdr_row]]
                
                # Identificamos columnas por nombre para que no falle si se mueven
                try:
                    idx_fecha = next(i for i, c in enumerate(cols) if 'FECHA' in c)
                    idx_odo   = next(i for i, c in enumerate(cols) if 'ODO' in c and 'ACUM' not in c)
                    idx_tkm   = next(i for i, c in enumerate(cols) if 'TREN' in c and 'KM' in c and 'ACUM' not in c)
                except StopIteration:
                    st.error(f"❌ Faltan columnas críticas en la hoja. Columnas encontradas: {cols}")
                    st.stop()

                df_data = df_raw.iloc[hdr_row + 1:].copy()
                df_data['_dt'] = pd.to_datetime(df_data.iloc[:, idx_fecha], errors='coerce')
                
                res_diario = []
                for d in f_dias:
                    mask = (df_data['_dt'].dt.day == d) & (df_data['_dt'].dt.month == mes_num) & (df_data['_dt'].dt.year == f_anio)
                    row_found = df_data[mask]
                    
                    if not row_found.empty:
                        val_odo = parse_latam_number(row_found.iloc[0, idx_odo])
                        val_tkm = parse_latam_number(row_found.iloc[0, idx_tkm])
                        umr_calc = (val_tkm / val_odo * 100) if val_odo > 0 else 0
                        
                        res_diario.append({
                            "Fecha": f"{d:02d}/{mes_num:02d}/{f_anio}",
                            "Odómetro [km]": val_odo,
                            "Tren-Km [km]": val_tkm,
                            "UMR [%]": umr_calc
                        })

                df_final = pd.DataFrame(res_diario)

                if not df_final.empty:
                    st.write(f"## 📊 Módulo UMR - {f_mes} {f_anio}")
                    c1, c2, c3 = st.columns(3)
                    t_odo = df_final["Odómetro [km]"].sum()
                    t_tkm = df_final["Tren-Km [km]"].sum()
                    c1.metric("Odómetro Total", f"{t_odo:,.1f} km")
                    c2.metric("Tren-Km Total", f"{t_tkm:,.1f} km")
                    c3.metric("UMR Promedio", f"{(t_tkm/t_odo*100 if t_odo>0 else 0):.2f} %")
                    
                    st.divider()
                    st.subheader("Desglose de Operación Diario")
                    st.dataframe(df_final.style.format({"Odómetro [km]": "{:,.1f}", "Tren-Km [km]": "{:,.1f}", "UMR [%]": "{:.2f}%"}), use_container_width=True)
                    st.download_button("📥 Descargar Reporte (Excel)", to_excel(df_final), f"Reporte_UMR_{f_mes}.xlsx")
                else:
                    st.warning("⚠️ No hay datos para los días seleccionados. Verifica que coincidan con el Excel.")
        else:
            st.error("❌ No encontré la hoja 'UMR Resumen'. Revisa el nombre en tu archivo Excel.")
            
    except Exception as e:
        st.error(f"💥 Ocurrió un error inesperado: {e}")
else:
    st.info("👋 Sube el archivo UMR en el panel de la izquierda para comenzar.")
