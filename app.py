import streamlit as st
import pandas as pd
import numpy as np
import re
import holidays
from io import BytesIO
from datetime import datetime

# --- 1. CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="EFE Valparaíso - Sistema de Gestión de Energía", layout="wide", page_icon="🚆")

# Feriados de Chile para clasificar D/F
chile_holidays = holidays.Chile()

# Estilo EFE para métricas
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

def to_excel(df_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in df_dict.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name[:31]) # Límite de Excel
    return output.getvalue()

def color_umr(val):
    """Semáforo: >96.4% Verde, <96.4% Rojo, ==96.4% Negro"""
    if val > 96.4: return 'color: green; font-weight: bold;'
    elif val < 96.4: return 'color: red; font-weight: bold;'
    return 'color: black;'

# --- 3. SIDEBAR (FILTROS) ---
MESES_NOMBRES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
MESES_NUM_MAP = {nombre: i+1 for i, nombre in enumerate(MESES_NOMBRES)}

with st.sidebar:
    st.header("📂 Gestión de Archivos")
    f_umr_list = st.file_uploader("Subir archivos Excel (UMR/Odómetro)", type=["xlsx"], accept_multiple_files=True)
    
    st.divider()
    st.subheader("📅 Filtros de Periodo")
    f_anio_list = st.multiselect("Seleccionar Años", [2024, 2025, 2026], default=[2025, 2026])
    f_mes_list = st.multiselect("Seleccionar Meses", MESES_NOMBRES, default=[MESES_NOMBRES[datetime.now().month - 1]])
    meses_num = [MESES_NUM_MAP[m] for m in f_mes_list]
    f_dias = st.multiselect("Seleccionar Días", list(range(1, 32)), default=list(range(1, 32)))

# --- 4. PROCESAMIENTO ---
if f_umr_list:
    all_resumen = []
    all_trenes = []
    
    for f in f_umr_list:
        try:
            xl = pd.ExcelFile(f)
            
            # --- HOJA: UMR RESUMEN ---
            sn_res = next((s for s in xl.sheet_names if 'UMR' in s.upper() and 'RESUMEN' in s.upper()), None)
            if sn_res:
                df_raw = pd.read_excel(f, sheet_name=sn_res, header=None)
                hdr_row = None
                for i in range(min(50, len(df_raw))):
                    fila_txt = " ".join(df_raw.iloc[i].astype(str)).upper()
                    if ('ODO' in fila_txt or 'FECHA' in fila_txt) and 'TREN' in fila_txt:
                        hdr_row = i; break
                
                if hdr_row is not None:
                    cols_orig = df_raw.iloc[hdr_row].astype(str).tolist()
                    cols_clean = [re.sub(r'[^A-Z]', '', c.upper().replace('Ó','O')) for c in cols_orig]
                    
                    def find_idx(aliases, clean_list, orig_list):
                        for i, c in enumerate(clean_list):
                            if 'ACUM' in orig_list[i].upper(): continue
                            if any(a in c for a in aliases): return i
                        return None

                    idx_fch = find_idx(['FECHA'], cols_clean, cols_orig)
                    idx_odo = find_idx(['ODO', 'METRO', 'KM'], cols_clean, cols_orig)
                    idx_tkm = find_idx(['TRENKM', 'TK', 'TRKM'], cols_clean, cols_orig)

                    if idx_fch is not None:
                        df_ext = df_raw.iloc[hdr_row+1:].copy()
                        df_ext['_dt'] = pd.to_datetime(df_ext.iloc[:, idx_fch], errors='coerce')
                        mask = (df_ext['_dt'].dt.day.isin(f_dias)) & (df_ext['_dt'].dt.month.isin(meses_num)) & (df_ext['_dt'].dt.year.isin(f_anio_list))
                        
                        for _, row in df_ext[mask].iterrows():
                            fch = row.iloc[idx_fch]
                            if not isinstance(fch, (datetime, pd.Timestamp)): continue
                            
                            # Lógica Tipo Día: L, S, D/F
                            nom_dia = fch.strftime('%A')
                            es_fest = fch in chile_holidays
                            if es_fest or nom_dia == 'Sunday': t_dia = "D/F"
                            elif nom_dia == 'Saturday': t_dia = "S"
                            else: t_dia = "L"
                            
                            odo_val = parse_latam_number(row.iloc[idx_odo])
                            tkm_val = parse_latam_number(row.iloc[idx_tkm])
                            
                            all_resumen.append({
                                "Fecha": fch.strftime('%d/%m/%Y'),
                                "Tipo Día": t_dia,
                                "N° Semana": fch.isocalendar()[1],
                                "Odómetro [km]": odo_val,
                                "Tren-Km [km]": tkm_val,
                                "UMR [%]": (tkm_val / odo_val * 100) if odo_val > 0 else 0,
                                "Timestamp": fch,
                                "Archivo": f.name
                            })

            # --- HOJA: ODOMETRO-KILOMETRAJE ---
            sn_tren = next((s for s in xl.sheet_names if 'ODO' in s.upper() and 'KIL' in s.upper()), None)
            if sn_tren:
                df_tr_raw = pd.read_excel(f, sheet_name=sn_tren, header=None)
                start_row = None
                for i in range(min(100, len(df_tr_raw))):
                    if 'KILOMETRAJE DIARIO RECORRIDO' in str(df_tr_raw.iloc[i, 0]).upper():
                        start_row = i; break
                
                if start_row is not None:
                    dias_row = start_row + 1
                    col_map = {}
                    for c_idx, val in enumerate(df_tr_raw.iloc[dias_row]):
                        try:
                            d_val = int(float(str(val)))
                            if 1 <= d_val <= 31: col_map[d_val] = c_idx
                        except: continue
                    
                    for r_idx in range(start_row + 2, len(df_tr_raw)):
                        nombre_tren = str(df_tr_raw.iloc[r_idx, 0]).strip().upper()
                        # Solo trenes M01-27 y XM28-35
                        if re.match(r'^(M\d{2}|XM\d{2})', nombre_tren):
                            for d_sel in f_dias:
                                if d_sel in col_map:
                                    val_km = parse_latam_number(df_tr_raw.iloc[r_idx, col_map[d_sel]])
                                    if val_km > 0:
                                        all_trenes.append({
                                            "Tren": nombre_tren,
                                            "Día": d_sel,
                                            "Kilometraje": val_km,
                                            "Archivo": f.name
                                        })
        except Exception as e:
            st.error(f"Error procesando {f.name}: {e}")

    # --- 5. RENDERIZADO DE PESTAÑAS ---
    if all_resumen:
        df_res = pd.DataFrame(all_resumen).drop_duplicates(subset=['Fecha'], keep='last').sort_values("Timestamp")
        
        t_res, t_datos, t_trenes = st.tabs(["📊 Resumen", "📑 Datos operacionales", "📑 Odómetro por Tren"])
        
        with t_res:
            col1, col2, col3 = st.columns(3)
            sum_odo = df_res["Odómetro [km]"].sum()
            sum_tkm = df_res["Tren-Km [km]"].sum()
            col1.metric("Odómetro Total", f"{sum_odo:,.1f} km")
            col2.metric("Tren-Km Total", f"{sum_tkm:,.1f} km")
            col3.metric("UMR Global", f"{(sum_tkm/sum_odo*100 if sum_odo>0 else 0):.2f} %")
            
            st.divider()
            st.write("### Desempeño promedio por tipo de jornada")
            res_tipo = df_res.groupby("Tipo Día")["UMR [%]"].mean().reset_index()
            st.table(res_tipo.style.format({"UMR [%]": "{:.2f}%"}))

        with t_datos:
            st.write("### Detalle Cronológico Operacional")
            # Orden de columnas solicitado
            cols_op = ["Fecha", "Tipo Día", "N° Semana", "Odómetro [km]", "Tren-Km [km]", "UMR [%]"]
            styled_op = df_res[cols_op].style.format({
                "Odómetro [km]": "{:,.1f}", "Tren-Km [km]": "{:,.1f}", "UMR [%]": "{:.2f}%"
            }).applymap(color_umr, subset=['UMR [%]'])
            
            st.dataframe(styled_op, use_container_width=True)

        with t_trenes:
            if all_trenes:
                df_tr = pd.DataFrame(all_trenes)
                # Matriz de Trenes vs Días
                pivot_trenes = df_tr.pivot_table(index="Tren", columns="Día", values="Kilometraje", aggfunc='sum').fillna(0)
                st.write("### Kilometraje Diario por Unidad (M / XM)")
                st.dataframe(pivot_trenes.style.format("{:,.1f}"), use_container_width=True)
            else:
                st.warning("No se encontraron datos individuales de trenes en la hoja 'Odometro-Kilometraje'.")
        
        # Botón de descarga consolidado
        st.download_button("📥 Descargar Reporte Completo (Excel)", 
                         to_excel({"Resumen": df_res, "Odometro_por_Tren": pd.DataFrame(all_trenes)}), 
                         "Reporte_SGE_EFE.xlsx")
    else:
        st.warning("No se encontraron registros válidos para los filtros seleccionados.")
else:
    st.info("👋 Sube los archivos Excel para comenzar el análisis del SGE.")
