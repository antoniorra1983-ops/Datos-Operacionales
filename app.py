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

# --- 3. PROCESAMIENTO INICIAL (SIN FILTROS) ---
all_resumen_raw = []
all_trenes_raw = []

with st.sidebar:
    st.header("📂 Gestión de Archivos")
    f_umr_list = st.file_uploader("Subir archivos Excel (UMR/Odómetro)", type=["xlsx"], accept_multiple_files=True)
    st.divider()

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
                        
                        for _, row in df_ext.dropna(subset=['_dt']).iterrows():
                            fch = row.iloc[idx_fch]
                            nom_dia = fch.strftime('%A')
                            t_dia = "D/F" if (fch in chile_holidays or nom_dia == 'Sunday') else ("S" if nom_dia == 'Saturday' else "L")
                            o, t = parse_latam_number(row.iloc[idx_odo]), parse_latam_number(row.iloc[idx_tkm])
                            all_resumen_raw.append({
                                "Fecha_DT": fch, "Fecha": fch.strftime('%d/%m/%Y'), "Tipo Día": t_dia, 
                                "N° Semana": fch.isocalendar()[1], "Odómetro [km]": o, 
                                "Tren-Km [km]": t, "UMR [%]": (t/o*100 if o>0 else 0), "Año": fch.year, "Mes": fch.month
                            })

            # --- PARTE B: ODOMETRO POR TREN (MEJORADO) ---
            sn_tren = next((s for s in xl.sheet_names if 'ODO' in s.upper() and 'KIL' in s.upper()), None)
            if sn_tren:
                df_tr_raw = pd.read_excel(f, sheet_name=sn_tren, header=None)
                row_tren = next((i for i in range(min(50, len(df_tr_raw))) if 'TREN' in str(df_tr_raw.iloc[i, 0]).upper() or 'M01' in str(df_tr_raw.iloc[i, 0]).upper()), None)
                if row_tren is not None:
                    fila_fechas = row_tren - 1
                    col_map = {} # {col_idx: datetime_obj}
                    for c_idx, val in enumerate(df_tr_raw.iloc[fila_fechas]):
                        f_parsed = pd.to_datetime(val, errors='coerce')
                        if pd.notna(f_parsed): col_map[c_idx] = f_parsed
                    
                    for r_idx in range(row_tren, len(df_tr_raw)):
                        nombre = str(df_tr_raw.iloc[r_idx, 0]).strip().upper()
                        if re.match(r'^(M\d{1,2}|XM\d{1,2})', nombre):
                            for c_idx, f_dt in col_map.items():
                                val_km = parse_latam_number(df_tr_raw.iloc[r_idx, c_idx])
                                all_trenes_raw.append({
                                    "Tren": nombre, "Fecha": f_dt, "Kilometraje": val_km, 
                                    "Año": f_dt.year, "Mes": f_dt.month, "Día": f_dt.day
                                })
        except Exception as e:
            st.error(f"Error en {f.name}: {e}")

# --- 4. FILTROS DINÁMICOS (POST-PROCESAMIENTO) ---
if all_resumen_raw:
    df_base_res = pd.DataFrame(all_resumen_raw).drop_duplicates(subset=['Fecha'], keep='last')
    df_base_tr = pd.DataFrame(all_trenes_raw)

    with st.sidebar:
        st.subheader("📅 Filtros Activos")
        anios_disp = sorted(df_base_res['Año'].unique())
        f_anios = st.multiselect("Filtrar por Año", anios_disp, default=anios_disp)
        
        meses_disp = sorted(df_base_res['Mes'].unique())
        MESES_NOMBRES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        f_meses_nombres = st.multiselect("Filtrar por Mes", MESES_NOMBRES, default=[MESES_NOMBRES[m-1] for m in meses_disp])
        f_meses_num = [MESES_NOMBRES.index(m)+1 for m in f_meses_nombres]

    # Aplicar Filtros
    df_res = df_base_res[df_base_res['Año'].isin(f_anios) & df_base_res['Mes'].isin(f_meses_num)].sort_values("Fecha_DT")
    df_tr = df_base_tr[df_base_tr['Año'].isin(f_anios) & df_base_tr['Mes'].isin(f_meses_num)]

    # --- 5. RENDERIZADO ---
    if not df_res.empty:
        t_res, t_datos, t_trenes = st.tabs(["📊 Resumen", "📑 Datos operacionales", "📑 Odómetro por Tren"])
        
        with t_res:
            c1, c2, c3 = st.columns(3)
            tot_o, tot_t = df_res["Odómetro [km]"].sum(), df_res["Tren-Km [km]"].sum()
            c1.metric("Odómetro Total", f"{tot_o:,.1f} km"); c2.metric("Tren-Km Total", f"{tot_t:,.1f} km"); c3.metric("UMR Global", f"{(tot_t/tot_o*100 if tot_o>0 else 0):.2f} %")
            st.divider()
            st.write("### Promedio UMR por Tipo de Jornada")
            st.table(df_res.groupby("Tipo Día")["UMR [%]"].mean().reset_index().style.format({"UMR [%]": "{:.2f}%"}))

        with t_datos:
            cols_op = ["Fecha", "Tipo Día", "N° Semana", "Odómetro [km]", "Tren-Km [km]", "UMR [%]"]
            st.dataframe(df_res[cols_op].style.format({"Odómetro [km]": "{:,.1f}", "Tren-Km [km]": "{:,.1f}", "UMR [%]": "{:.2f}%"}).applymap(color_umr, subset=['UMR [%]']), use_container_width=True)

        with t_trenes:
            if not df_tr.empty:
                pivot_tr = df_tr.pivot_table(index="Tren", columns="Día", values="Kilometraje", aggfunc='sum').fillna(0)
                st.write(f"### Kilometraje por Unidad - {f_meses_nombres[0] if f_meses_nombres else ''}")
                st.dataframe(pivot_tr.style.format("{:,.1f}"), use_container_width=True)
            else:
                st.warning("No hay datos de trenes para el periodo seleccionado.")
        
        st.download_button("📥 Descargar Reporte (Excel)", to_excel({"Resumen": df_res, "Odometro_por_Tren": df_tr}), "Reporte_SGE_EFE.xlsx")
    else:
        st.warning("No hay datos que coincidan con los filtros seleccionados.")
else:
    if f_umr_list:
        st.error("❌ No se detectaron datos válidos en los archivos. Revisa que las hojas se llamen 'UMR Resumen' y 'Odometro-Kilometraje'.")
    else:
        st.info("👋 Sube los archivos Excel para comenzar el análisis.")
