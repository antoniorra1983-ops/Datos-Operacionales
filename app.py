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

def to_excel_consolidado(df_ops, df_trenes):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_ops.to_excel(writer, index=False, sheet_name='Datos_Operacionales')
        df_trenes.to_excel(writer, index=False, sheet_name='Detalle_Kilometraje_Trenes')
    return output.getvalue()

def color_umr(val):
    if val > 96.4: return 'color: green; font-weight: bold;'
    elif val < 96.4: return 'color: red; font-weight: bold;'
    return 'color: black;'

# --- 3. SIDEBAR (FILTROS DINÁMICOS) ---
st.sidebar.header("📂 Gestión de Archivos")
f_umr_list = st.sidebar.file_uploader("Subir archivos Excel (UMR/Odómetro)", type=["xlsx"], accept_multiple_files=True)

all_resumen_raw = []
all_trenes_raw = []

if f_umr_list:
    for f in f_umr_list:
        try:
            xl = pd.ExcelFile(f)
            
            # --- SECCIÓN A: HOJA RESUMEN UMR ---
            sn_resumen = next((s for s in xl.sheet_names if 'UMR' in s.upper() and 'RESUMEN' in s.upper()), None)
            if sn_resumen:
                df_raw_res = pd.read_excel(f, sheet_name=sn_resumen, header=None)
                hdr_row = next((i for i in range(min(50, len(df_raw_res))) if 'ODO' in " ".join(df_raw_res.iloc[i].astype(str)).upper()), None)
                if hdr_row is not None:
                    cols = df_raw_res.iloc[hdr_row].astype(str).tolist()
                    idx_fch = next((i for i, c in enumerate(cols) if 'FECHA' in c.upper()), None)
                    idx_odo = next((i for i, c in enumerate(cols) if 'ODO' in c.upper() and 'ACUM' not in c.upper()), None)
                    idx_tkm = next((i for i, c in enumerate(cols) if 'TREN' in c.upper() and 'KM' in c.upper() and 'ACUM' not in c.upper()), None)
                    
                    if idx_fch is not None:
                        df_ext = df_raw_res.iloc[hdr_row+1:].copy()
                        df_ext['_dt'] = pd.to_datetime(df_ext.iloc[:, idx_fch], errors='coerce')
                        for _, row in df_ext.dropna(subset=['_dt']).iterrows():
                            fch = row.iloc[idx_fch]
                            all_resumen_raw.append({
                                "Fecha_DT": fch, "Fecha": fch.strftime('%d/%m/%Y'), "Año": fch.year, "Mes": fch.month, "Día": fch.day,
                                "Odómetro [km]": parse_latam_number(row.iloc[idx_odo]), "Tren-Km [km]": parse_latam_number(row.iloc[idx_tkm])
                            })

            # --- SECCIÓN B: HOJA ODOMETRO-KILOMETRAJE (M MOTORIZADO) ---
            sn_trenes = next((s for s in xl.sheet_names if 'ODO' in s.upper() and 'KIL' in s.upper()), None)
            if sn_trenes:
                df_tr_raw = pd.read_excel(f, sheet_name=sn_trenes, header=None)
                
                # 1. Buscar fila de FECHA y DIARIO (Triple encabezado)
                date_row = None
                diario_row = None
                col_to_date = {}

                for i in range(min(50, len(df_tr_raw))):
                    fila_txt = " ".join(df_tr_raw.iloc[i].astype(str)).upper()
                    # Si encontramos una fila con fechas
                    if any(pd.to_datetime(val, errors='coerce').year > 2000 for val in df_tr_raw.iloc[i]):
                        date_row = i
                        for j, val in enumerate(df_tr_raw.iloc[i]):
                            d_parsed = pd.to_datetime(val, errors='coerce')
                            if pd.notna(d_parsed): col_to_date[j] = d_parsed
                        
                        # Verificamos si 2 filas abajo dice "DIARIO" (según tu descripción)
                        if i+2 < len(df_tr_raw) and 'DIARIO' in " ".join(df_tr_raw.iloc[i+2].astype(str)).upper():
                            diario_row = i+2
                            break

                # 2. Extraer datos si encontramos la estructura
                if date_row is not None:
                    # El inicio de los trenes suele ser después de los encabezados (diario_row + 1)
                    start_trains = (diario_row + 1) if diario_row else (date_row + 2)
                    for i in range(start_trains, len(df_tr_raw)):
                        nombre_tren = str(df_tr_raw.iloc[i, 0]).strip().upper()
                        # Filtro Flota EFE Valpo
                        if re.match(r'^(M\d{1,2}|XM\d{1,2})', nombre_tren):
                            for col_idx, col_date in col_to_date.items():
                                val_km = parse_latam_number(df_tr_raw.iloc[i, col_idx])
                                all_trenes_raw.append({
                                    "Tren": nombre_tren, "Fecha_DT": col_date, "Kilometraje": val_km,
                                    "Año": col_date.year, "Mes": col_date.month, "Día": col_date.day
                                })
        except: continue

# --- 4. RENDERIZADO Y FILTROS ---
if all_resumen_raw or all_trenes_raw:
    df_res_base = pd.DataFrame(all_resumen_raw).drop_duplicates(subset=['Fecha']) if all_resumen_raw else pd.DataFrame()
    df_tr_base = pd.DataFrame(all_trenes_raw) if all_trenes_raw else pd.DataFrame()

    anios_totales = sorted(list(set(df_res_base['Año'].unique() if not df_res_base.empty else []) | set(df_tr_base['Año'].unique() if not df_tr_base.empty else [])))
    
    st.sidebar.divider()
    sel_anios = st.sidebar.multiselect("Año", anios_totales, default=anios_totales)
    MESES_NOMBRES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    sel_meses = st.sidebar.multiselect("Mes", MESES_NOMBRES, default=MESES_NOMBRES)
    sel_meses_num = [MESES_NOMBRES.index(m) + 1 for m in sel_meses]

    df_res = df_res_base[df_res_base['Año'].isin(sel_anios) & df_res_base['Mes'].isin(sel_meses_num)].sort_values("Fecha_DT") if not df_res_base.empty else pd.DataFrame()
    df_tr = df_tr_base[df_tr_base['Año'].isin(sel_anios) & df_tr_base['Mes'].isin(sel_meses_num)] if not df_tr_base.empty else pd.DataFrame()

    tab1, tab2, tab3 = st.tabs(["📊 Resumen", "📑 Datos operacionales", "📑 Odómetro por Tren"])

    with tab1:
        if not df_res.empty:
            c1, c2, c3 = st.columns(3)
            t_o, t_t = df_res["Odómetro [km]"].sum(), df_res["Tren-Km [km]"].sum()
            c1.metric("Odómetro Total", f"{t_o:,.1f} km"); c2.metric("Tren-Km Total", f"{t_t:,.1f} km"); c3.metric("UMR Global", f"{(t_t/t_o*100 if t_o>0 else 0):.2f} %")
        else: st.info("Suba archivos para ver el resumen.")

    with tab2:
        if not df_res.empty:
            def get_tipo_dia(fch):
                if fch in chile_holidays or fch.strftime('%A') == 'Sunday': return "D/F"
                return "S" if fch.strftime('%A') == 'Saturday' else "L"
            df_res['Tipo Día'] = df_res['Fecha_DT'].apply(get_tipo_dia)
            df_res['N° Semana'] = df_res['Fecha_DT'].dt.isocalendar().week
            st.dataframe(df_res[["Fecha", "Tipo Día", "N° Semana", "Odómetro [km]", "Tren-Km [km]", "UMR [%]"]]
                         .style.format({"Odómetro [km]": "{:,.1f}", "Tren-Km [km]": "{:,.1f}", "UMR [%]": "{:.2f}%"})
                         .applymap(color_umr, subset=['UMR [%]']), use_container_width=True)

    with tab3:
        if not df_tr.empty:
            st.write("### 📏 Kilometraje Total por Tren (Acumulado en periodo)")
            resumen_total_tren = df_tr.groupby("Tren")["Kilometraje"].sum().reset_index().sort_values("Kilometraje", ascending=False)
            st.dataframe(resumen_total_tren.style.format({"Kilometraje": "{:,.1f}"}), use_container_width=True)
            
            st.divider()
            st.write("### 📅 Kilometrajes Diarios (Detalle por unidad)")
            pivot_diario = df_tr.pivot_table(index="Tren", columns="Día", values="Kilometraje", aggfunc='sum').fillna(0)
            st.dataframe(pivot_diario.style.format("{:,.1f}"), use_container_width=True)
        else:
            st.warning("No se encontraron datos en 'Odometro-Kilometraje'. Verifique la estructura: Fecha > Kilometraje > Diario.")

    if not df_res.empty or not df_tr.empty:
        st.sidebar.download_button("📥 Descargar Reporte Completo", to_excel_consolidado(df_res, df_tr), "Reporte_SGE_EFE.xlsx")
else:
    st.info("👋 Por favor, sube los archivos UMR para comenzar el análisis técnico.")
