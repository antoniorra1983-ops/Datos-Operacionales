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

def to_excel(df_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in df_dict.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    return output.getvalue()

def color_umr(val):
    if val > 96.4: return 'color: green; font-weight: bold;'
    elif val < 96.4: return 'color: red; font-weight: bold;'
    return 'color: black;'

# --- 3. SIDEBAR ---
MESES_NOMBRES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
MESES_NUM_MAP = {nombre: i+1 for i, nombre in enumerate(MESES_NOMBRES)}

with st.sidebar:
    st.header("📂 Gestión de Archivos")
    f_umr_list = st.file_uploader("Subir archivos Excel (UMR/Odómetro)", type=["xlsx"], accept_multiple_files=True)
    st.divider()
    f_anio_list = st.multiselect("Años", [2024, 2025, 2026], default=[2025, 2026])
    f_mes_list = st.multiselect("Meses", MESES_NOMBRES, default=[MESES_NOMBRES[datetime.now().month - 1]])
    meses_num = [MESES_NUM_MAP[m] for m in f_mes_list]
    f_dias = st.multiselect("Días", list(range(1, 32)), default=list(range(1, 32)))

# --- 4. PROCESAMIENTO ---
if f_umr_list:
    all_resumen = []
    all_trenes = []
    
    for f in f_umr_list:
        try:
            xl = pd.ExcelFile(f)
            
            # --- PARTE A: UMR RESUMEN ---
            sn_res = next((s for s in xl.sheet_names if 'UMR' in s.upper() and 'RESUMEN' in s.upper()), None)
            if sn_res:
                df_raw = pd.read_excel(f, sheet_name=sn_res, header=None)
                hdr_row = next((i for i in range(min(50, len(df_raw))) if 'ODO' in " ".join(df_raw.iloc[i].astype(str)).upper()), None)
                if hdr_row is not None:
                    cols_orig = df_raw.iloc[hdr_row].astype(str).tolist()
                    def find_idx(aliases, orig_list):
                        for i, c in enumerate(orig_list):
                            c_up = c.upper().replace('Ó','O')
                            if 'ACUM' in c_up: continue
                            if any(a in c_up for a in aliases): return i
                        return None
                    idx_fch, idx_odo, idx_tkm = find_idx(['FECHA'], cols_orig), find_idx(['ODO', 'METRO'], cols_orig), find_idx(['TRENKM', 'TK', 'TRKM'], cols_orig)
                    if idx_fch is not None:
                        df_ext = df_raw.iloc[hdr_row+1:].copy()
                        df_ext['_dt'] = pd.to_datetime(df_ext.iloc[:, idx_fch], errors='coerce')
                        mask = (df_ext['_dt'].dt.day.isin(f_dias)) & (df_ext['_dt'].dt.month.isin(meses_num)) & (df_ext['_dt'].dt.year.isin(f_anio_list))
                        for _, row in df_ext[mask].iterrows():
                            fch = row.iloc[idx_fch]
                            if not isinstance(fch, (datetime, pd.Timestamp)): continue
                            nom_dia = fch.strftime('%A')
                            t_dia = "D/F" if (fch in chile_holidays or nom_dia == 'Sunday') else ("S" if nom_dia == 'Saturday' else "L")
                            o, t = parse_latam_number(row.iloc[idx_odo]), parse_latam_number(row.iloc[idx_tkm])
                            all_resumen.append({"Fecha": fch.strftime('%d/%m/%Y'), "Tipo Día": t_dia, "N° Semana": fch.isocalendar()[1], "Odómetro [km]": o, "Tren-Km [km]": t, "UMR [%]": (t/o*100 if o>0 else 0), "Timestamp": fch})

            # --- PARTE B: ODOMETRO-KILOMETRAJE (BASADO EN TU IMAGEN) ---
            sn_tren = next((s for s in xl.sheet_names if 'ODO' in s.upper() and 'KIL' in s.upper()), None)
            if sn_tren:
                df_tr_raw = pd.read_excel(f, sheet_name=sn_tren, header=None)
                # Buscamos la fila que dice "Tren" o "M01"
                row_tren = next((i for i in range(min(50, len(df_tr_raw))) if 'TREN' in str(df_tr_raw.iloc[i, 0]).upper() or 'M01' in str(df_tr_raw.iloc[i, 0]).upper()), None)
                
                if row_tren is not None:
                    # Las fechas suelen estar 1 o 2 filas arriba de los trenes
                    fila_fechas = row_tren - 1 if row_tren > 0 else 0
                    col_map = {} # {dia_num: col_index}
                    for c_idx, val in enumerate(df_tr_raw.iloc[fila_fechas]):
                        try:
                            # Intentamos extraer el día de formatos como "01-04-2025" o simplemente "1"
                            if isinstance(val, (datetime, pd.Timestamp)): d_val = val.day
                            else:
                                match = re.search(r'(\d{1,2})', str(val))
                                d_val = int(match.group(1)) if match else None
                            if d_val and 1 <= d_val <= 31: col_map[d_val] = c_idx
                        except: continue
                    
                    # Extraer kilometraje por tren
                    for r_idx in range(row_tren, len(df_tr_raw)):
                        nombre = str(df_tr_raw.iloc[r_idx, 0]).strip().upper()
                        if re.match(r'^(M\d{1,2}|XM\d{1,2})', nombre):
                            for d_sel in f_dias:
                                if d_sel in col_map:
                                    val_km = parse_latam_number(df_tr_raw.iloc[r_idx, col_map[d_sel]])
                                    if val_km >= 0: # Incluimos 0 para ver si el tren estuvo parado
                                        all_trenes.append({"Tren": nombre, "Día": d_sel, "Kilometraje": val_km})
        except Exception as e:
            st.error(f"Error en {f.name}: {e}")

    # --- 5. RENDERIZADO ---
    if all_resumen:
        df_res = pd.DataFrame(all_resumen).drop_duplicates(subset=['Fecha'], keep='last').sort_values("Timestamp")
        t_res, t_datos, t_trenes = st.tabs(["📊 Resumen", "📑 Datos operacionales", "📑 Odómetro por Tren"])
        
        with t_res:
            c1, c2, c3 = st.columns(3)
            tot_o, tot_t = df_res["Odómetro [km]"].sum(), df_res["Tren-Km [km]"].sum()
            c1.metric("Odómetro Total", f"{tot_o:,.1f} km"); c2.metric("Tren-Km Total", f"{tot_t:,.1f} km"); c3.metric("UMR Global", f"{(tot_t/tot_o*100 if tot_o>0 else 0):.2f} %")
            st.divider()
            st.write("### Promedio UMR por Tipo de Jornada")
            st.table(df_res.groupby("Tipo Día")["UMR [%]"].mean().reset_index().style.format({"UMR [%]": "{:.2f}%"}))

        with t_datos:
            st.dataframe(df_res[["Fecha", "Tipo Día", "N° Semana", "Odómetro [km]", "Tren-Km [km]", "UMR [%]"]].style.format({"Odómetro [km]": "{:,.1f}", "Tren-Km [km]": "{:,.1f}", "UMR [%]": "{:.2f}%"}).applymap(color_umr, subset=['UMR [%]']), use_container_width=True)

        with t_trenes:
            if all_trenes:
                df_tr_final = pd.DataFrame(all_trenes).pivot_table(index="Tren", columns="Día", values="Kilometraje", aggfunc='sum').fillna(0)
                st.write("### Matriz de Kilometraje por Unidad (M/XM)")
                st.dataframe(df_tr_final.style.format("{:,.1f}"), use_container_width=True)
            else:
                st.warning("⚠️ No se pudieron extraer datos de la hoja Odómetro-Kilometraje. Revisa que el nombre de los trenes esté en la columna A.")
        
        st.download_button("📥 Descargar Reporte (Excel)", to_excel({"Resumen": df_res, "Odometro_por_Tren": pd.DataFrame(all_trenes)}), "Reporte_SGE_EFE.xlsx")
    else:
        st.warning("No se encontraron registros válidos.")
