import streamlit as st
import pandas as pd
import numpy as np
import re
import holidays
from io import BytesIO
from datetime import datetime

# --- 1. CONFIGURACIÓN ---
st.set_page_config(page_title="EFE Valparaíso - Auditoría UMR", layout="wide", page_icon="🚆")

# Mapeo de días y feriados
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
    # Años y Meses multiselección
    f_anio_list = st.multiselect("Años", [2024, 2025, 2026], default=[2024, 2025, 2026])
    f_mes_list = st.multiselect("Meses", MESES_NOMBRES, default=MESES_NOMBRES)
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
            
            # Buscador dinámico de fila de títulos (Busca palabras clave en la fila)
            hdr_row = None
            for i in range(min(100, len(df_raw))):
                fila_txt = " ".join(df_raw.iloc[i].astype(str)).upper()
                if ('ODO' in fila_txt or 'FECHA' in fila_txt) and 'TREN' in fila_txt:
                    hdr_row = i
                    break
            
            if hdr_row is None:
                st.warning(f"⚠️ {f.name}: No se detectó la fila de títulos (Fecha/Odo/Tren).")
                continue

            # Limpiamos los nombres de las columnas para la búsqueda
            cols_orig = df_raw.iloc[hdr_row].astype(str).tolist()
            cols_clean = [re.sub(r'[^A-Z]', '', c.upper().replace('Ó','O').replace('Á','A')) for c in cols_orig]
            
            # --- BUSCADOR FLEXIBLE POR ALIAS ---
            def encontrar_columna(alias_list, lista_columnas, original_cols):
                for i, col_limpia in enumerate(lista_columnas):
                    # No buscamos si la columna original dice "ACUMULADO"
                    if 'ACUM' in original_cols[i].upper(): continue
                    if any(alias in col_limpia for alias in alias_list):
                        return i
                return None

            idx_fecha = encontrar_columna(['FECHA', 'FCH', 'DATE'], cols_clean, cols_orig)
            idx_odo   = encontrar_columna(['ODO', 'METRO', 'KILO', 'KM'], cols_clean, cols_orig)
            idx_tkm   = encontrar_columna(['TRENKM', 'TK', 'TRKM', 'KMTR'], cols_clean, cols_orig)

            # Si el ODO falló, intentamos una búsqueda más desesperada (la primera columna numérica después de fecha)
            if idx_odo is None and idx_fecha is not None:
                idx_odo = idx_fecha + 1 

            if None in [idx_fecha, idx_odo, idx_tkm]:
                missing = []
                if idx_fecha is None: missing.append("Fecha")
                if idx_odo is None: missing.append("Odómetro")
                if idx_tkm is None: missing.append("Tren-Km")
                st.error(f"❌ {f.name}: Faltan columnas críticas: {', '.join(missing)}. Detectadas: {cols_orig}")
                continue

            df_data = df_raw.iloc[hdr_row + 1:].copy()
            df_data['_dt'] = pd.to_datetime(df_data.iloc[:, idx_fecha], errors='coerce')
            
            # Filtro por selección del usuario
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
                
                # Ecuación solicitada: (Tren-Km / Odómetro) * 100
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
            st.error(f"💥 Error en {f.name}: {str(e)}")

    if all_res_diario:
        df_final = pd.DataFrame(all_res_diario).drop_duplicates(subset=['Fecha'], keep='last').sort_values("Timestamp")
        
        st.write(f"## 📊 Consolidado Operacional EFE Valparaíso")
        
        # Métricas Globales
        c1, c2, c3 = st.columns(3)
        t_odo, t_tkm = df_final["Odómetro [km]"].sum(), df_final["Tren-Km [km]"].sum()
        c1.metric("Odómetro Total", f"{t_odo:,.1f} km")
        c2.metric("Tren-Km Total", f"{t_tkm:,.1f} km")
        c3.metric("UMR Global (KPI)", f"{(t_tkm/t_odo*100 if t_odo>0 else 0):.2f} %")
        
        st.divider()
        
        # Tabla Principal con el semáforo del 96.4%
        st.dataframe(
            df_final[["Fecha", "Día", "Festivo", "Odómetro [km]", "Tren-Km [km]", "UMR [%]", "Archivo"]]
            .style.format({"Odómetro [km]": "{:,.1f}", "Tren-Km [km]": "{:,.1f}", "UMR [%]": "{:.2f}%"})
            .applymap(color_umr, subset=['UMR [%]']), 
            use_container_width=True
        )
    else:
        st.warning("⚠️ No se encontraron registros. Asegúrate de que el Año y Mes seleccionados en el filtro coincidan con tus archivos.")
