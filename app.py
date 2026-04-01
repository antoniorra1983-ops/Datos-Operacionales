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

def to_excel_consolidado(df_ops, df_trenes, df_seat, df_prmte_d, df_prmte_15, df_fact_h, df_fact_d):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        if not df_ops.empty: df_ops.to_excel(writer, index=False, sheet_name='Datos_Operacionales')
        if not df_trenes.empty: df_trenes.to_excel(writer, index=False, sheet_name='Detalle_Kilometraje')
        if not df_seat.empty: df_seat.to_excel(writer, index=False, sheet_name='Energia_SEAT')
        if not df_prmte_d.empty: df_prmte_d.to_excel(writer, index=False, sheet_name='PRMTE_Diario')
        if not df_prmte_15.empty: df_prmte_15.to_excel(writer, index=False, sheet_name='PRMTE_15min')
        if not df_fact_h.empty: df_fact_h.to_excel(writer, index=False, sheet_name='Factura_Horaria')
        if not df_fact_d.empty: df_fact_d.to_excel(writer, index=False, sheet_name='Factura_Diaria_Analitica')
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
    f_umr_list = st.file_uploader("1. UMR / Odómetros", type=["xlsx"], accept_multiple_files=True)
    f_seat_list = st.file_uploader("2. Energía SEAT", type=["xlsx"], accept_multiple_files=True)
    f_billing_list = st.file_uploader("3. Facturación y PRMTE", type=["xlsx"], accept_multiple_files=True)
    
    st.divider()
    f_anio_list = st.multiselect("Años", [2024, 2025, 2026], default=[2025, 2026])
    # Cambiamos el default para que se vean todos los datos cargados de inmediato
    f_mes_list = st.multiselect("Meses", MESES_NOMBRES, default=MESES_NOMBRES)
    meses_num = [MESES_NUM_MAP[m] for m in f_mes_list]
    f_dias = st.multiselect("Días", list(range(1, 32)), default=list(range(1, 32)))

# --- 4. PROCESAMIENTO ---
all_data_ops, all_data_trenes, all_data_seat, all_data_prmte_15, all_data_factura_h = [], [], [], [], []
todos_archivos = (f_umr_list or []) + (f_seat_list or []) + (f_billing_list or [])

for f in todos_archivos:
    try:
        xl = pd.ExcelFile(f)
        for sn in xl.sheet_names:
            sn_up = sn.upper()
            
            # A. UMR / RESUMEN
            if any(k in sn_up for k in ['UMR', 'RESUMEN']):
                df_r = pd.read_excel(f, sheet_name=sn, header=None)
                h_r = next((i for i in range(min(100, len(df_r))) if any(k in " ".join(df_r.iloc[i].astype(str)).upper() for k in ['ODO', 'FECHA', 'TREN'])), None)
                if h_r is not None:
                    cols_raw = df_r.iloc[h_r].astype(str).tolist()
                    cols_clean = [re.sub(r'[^A-Z]', '', c.upper().replace('Ó','O')) for c in cols_raw]
                    idx_f, idx_o, idx_t = next((i for i, c in enumerate(cols_clean) if 'FECHA' in c), None), next((i for i, c in enumerate(cols_clean) if 'ODO' in c and 'ACUM' not in cols_raw[i].upper()), None), next((i for i, c in enumerate(cols_clean) if 'TREN' in c and 'KM' in c), None)
                    if None not in [idx_f, idx_o, idx_t]:
                        df_e = df_r.iloc[h_r+1:].copy()
                        df_e['_dt'] = pd.to_datetime(df_e.iloc[:, idx_f], errors='coerce')
                        mask = (df_e['_dt'].dt.day.isin(f_dias)) & (df_e['_dt'].dt.month.isin(meses_num)) & (df_e['_dt'].dt.year.isin(f_anio_list))
                        for _, row in df_e[mask].dropna(subset=['_dt']).iterrows():
                            fch = row.iloc[idx_f]
                            t_dia = "D/F" if (fch in chile_holidays or fch.strftime('%A') == 'Sunday') else ("S" if fch.strftime('%A') == 'Saturday' else "L")
                            o, tk = parse_latam_number(row.iloc[idx_o]), parse_latam_number(row.iloc[idx_t])
                            if o > 0: all_data_ops.append({"Fecha": fch.strftime('%d/%m/%Y'), "Tipo Día": t_dia, "N° Semana": fch.isocalendar()[1], "Odómetro [km]": o, "Tren-Km [km]": tk, "UMR [%]": (tk/o*100), "Timestamp": fch})

            # B. ODÓMETRO POR TREN
            if 'ODO' in sn_up and 'KIL' in sn_up:
                df_tr = pd.read_excel(f, sheet_name=sn, header=None)
                d_idx, c_map = None, {}
                for i in range(min(80, len(df_tr)-2)):
                    for j in range(1, len(df_tr.columns)):
                        p_d = pd.to_datetime(df_tr.iloc[i, j], errors='coerce')
                        if pd.notna(p_d) and p_d.year in f_anio_list:
                            if any(k in str(df_tr.iloc[i+1, j]).upper() for k in ['KILO', 'DIARIO']):
                                d_idx, c_map[j] = i, p_d
                if d_idx is not None:
                    for i in range(d_idx+3, len(df_tr)):
                        n_tr = str(df_tr.iloc[i, 0]).strip().upper()
                        if re.match(r'^(M|XM)', n_tr):
                            for c_idx, c_fch in c_map.items():
                                if c_fch.day in f_dias and c_fch.month in meses_num:
                                    all_data_trenes.append({"Tren": n_tr, "Fecha": c_fch.strftime('%d/%m/%Y'), "Día": c_fch.day, "Kilometraje Diario [km]": parse_latam_number(df_tr.iloc[i, c_idx]), "Timestamp": c_fch})

            # C. ENERGÍA SEAT
            if 'SEAT' in sn_up and 'SER' in sn_up:
                df_s = pd.read_excel(f, sheet_name=sn, header=None)
                for i in range(len(df_s)):
                    fs = pd.to_datetime(df_s.iloc[i, 1], errors='coerce')
                    if pd.notna(fs) and fs.year in f_anio_list and fs.month in meses_num:
                        tot, tra, k12 = parse_latam_number(df_s.iloc[i, 3]), parse_latam_number(df_s.iloc[i, 5]), parse_latam_number(df_s.iloc[i, 7])
                        all_data_seat.append({"Fecha": fs.strftime('%d/%m/%Y'), "Total [kWh]": tot, "Tracción [kWh]": tra, "12 KV [kWh]": k12, "% Tracción": (tra/tot*100 if tot>0 else 0), "% 12 KV": (k12/tot*100 if tot>0 else 0), "Timestamp": fs})

            # D. PRMTE (15 MINUTOS Y DIARIO)
            if any(k in sn_up for k in ['PRMTE', 'MEDIDAS']):
                df_p = pd.read_excel(f, sheet_name=sn, header=None)
                h_idx = next((i for i, r in df_p.iterrows() if 'AÑO' in [str(c).upper() for c in r]), None)
                if h_idx is not None:
                    df_p.columns = [str(c).strip() for c in df_p.iloc[h_idx]]; df_p = df_p.iloc[h_idx+1:].copy()
                    # Reconstrucción de tiempo para archivos con Año, Mes, Día separados
                    df_p['Timestamp'] = pd.to_datetime(df_p[['AÑO', 'MES', 'DIA', 'HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'})) + pd.to_timedelta(df_p['INICIO INTERVALO'].astype(int), unit='m')
                    mask_p = (df_p['Timestamp'].dt.month.isin(meses_num)) & (df_p['Timestamp'].dt.year.isin(f_anio_list))
                    for _, r in df_p[mask_p].iterrows():
                        all_data_prmte_15.append({"Fecha y Hora": r['Timestamp'].strftime('%d/%m/%Y %H:%M'), "Fecha": r['Timestamp'].strftime('%d/%m/%Y'), "Energía PRMTE [kWh]": parse_latam_number(r['Retiro_Energia_Activa (kWhD)']), "Timestamp": r['Timestamp']})

            # E. FACTURA (HORARIA)
            if any(k in sn_up for k in ['FACTURA', 'CONSUMO']):
                df_f = pd.read_excel(f, sheet_name=sn)
                df_f['Timestamp'] = pd.to_datetime(df_f.iloc[:, 0], errors='coerce')
                df_f = df_f.dropna(subset=['Timestamp'])
                mask_f = (df_f['Timestamp'].dt.month.isin(meses_num)) & (df_f['Timestamp'].dt.year.isin(f_anio_list))
                for _, r in df_f[mask_f].iterrows():
                    all_data_factura_h.append({
                        "Fecha y Hora": r['Timestamp'].strftime('%d/%m/%Y %H:%M'), 
                        "Fecha": r['Timestamp'].strftime('%d/%m/%Y'), 
                        "Consumo Horario [kWh]": abs(parse_latam_number(r.iloc[1])), 
                        "Timestamp": r['Timestamp']
                    })
    except Exception as e:
        st.sidebar.error(f"Error procesando {f.name}: {e}")

# --- 5. RENDERIZADO (BLINDAJE DE VARIABLES) ---
df_ops, df_tr, df_seat, df_prmte_15, df_prmte_d, df_fact_h, df_fact_d = [pd.DataFrame()] * 7

if any([all_data_ops, all_data_seat, all_data_factura_h, all_data_prmte_15]):
    if all_data_ops: df_ops = pd.DataFrame(all_data_ops).drop_duplicates(subset=['Fecha']).sort_values("Timestamp")
    if all_data_trenes: df_tr = pd.DataFrame(all_data_trenes).sort_values(["Timestamp", "Tren"])
    if all_data_seat: df_seat = pd.DataFrame(all_data_seat).drop_duplicates(subset=['Fecha']).sort_values("Timestamp")
    
    if all_data_prmte_15:
        df_prmte_15 = pd.DataFrame(all_data_prmte_15).sort_values("Timestamp")
        df_prmte_d = df_prmte_15.groupby("Fecha")["Energía PRMTE [kWh]"].sum().reset_index().rename(columns={"Energía PRMTE [kWh]":"Total PRMTE Diario [kWh]"})

    if all_data_factura_h:
        df_fact_h = pd.DataFrame(all_data_factura_h).sort_values("Timestamp")
        df_fact_d = df_fact_h.groupby("Fecha")["Consumo Horario [kWh]"].sum().reset_index().rename(columns={"Consumo Horario [kWh]":"Total Facturado [kWh]"})
        if not df_seat.empty:
            df_fact_d = pd.merge(df_fact_d, df_seat[["Fecha", "% Tracción", "% 12 KV"]], on="Fecha", how="left")
            df_fact_d["Energía Tracción Facturación [kWh]"] = df_fact_d["Total Facturado [kWh]"] * (df_fact_d["% Tracción"] / 100)
            df_fact_d["Energía 12kV Facturación [kWh]"] = df_fact_d["Total Facturado [kWh]"] * (df_fact_d["% 12 KV"] / 100)

    tabs = st.tabs(["📊 Resumen", "📑 Datos operacionales", "📑 Odómetro por Tren", "⚡ Energía SEAT", "📈 Medidas PRMTE", "💰 Facturación"])
    
    with tabs[0]: # Resumen
        if not df_ops.empty:
            st.subheader("Indicadores Globales")
            c1, c2, c3 = st.columns(3)
            to, tk = df_ops["Odómetro [km]"].sum(), df_ops["Tren-Km [km]"].sum()
            c1.metric("Odómetro Total", f"{to:,.1f} km"); c2.metric("Tren-Km Total", f"{tk:,.1f} km"); c3.metric("UMR Global", f"{(tk/to*100 if to>0 else 0):.2f} %")
            st.divider()
            df_ops['Tipo Día'] = pd.Categorical(df_ops['Tipo Día'], categories=['L', 'S', 'D/F'], ordered=True)
            res_t = df_ops.groupby("Tipo Día").agg({"Odómetro [km]":"sum", "Tren-Km [km]":"sum", "UMR [%]":"mean"}).reset_index()
            st.table(res_t.style.format({"Odómetro [km]":"{:,.1f}", "Tren-Km [km]":"{:,.1f}", "UMR [%]":"{:.2f}%"}))

    with tabs[4]: # PRMTE
        if not df_prmte_15.empty:
            st.subheader("Análisis PRMTE")
            st.write("#### 📅 Resumen Diario")
            st.dataframe(df_prmte_d.style.format({"Total PRMTE Diario [kWh]": "{:,.1f}"}), use_container_width=True)
            st.divider()
            st.write("#### 🕒 Detalle cada 15 Minutos")
            st.dataframe(df_prmte_15[["Fecha y Hora", "Energía PRMTE [kWh]"]].style.format({"Energía PRMTE [kWh]": "{:,.2f}"}), use_container_width=True)

    with tabs[5]: # FACTURACIÓN
        if not df_fact_h.empty:
            st.write("#### 📅 Resumen Diario (Proporcional SEAT)")
            cols_f = ["Fecha", "Total Facturado [kWh]", "Energía Tracción Facturación [kWh]", "Energía 12kV Facturación [kWh]"]
            st.dataframe(df_fact_d[[c for c in cols_f if c in df_fact_d.columns]].style.format({
                "Total Facturado [kWh]":"{:,.1f}", "Energía Tracción Facturación [kWh]":"{:,.1f}", "Energía 12kV Facturación [kWh]":"{:,.1f}"
            }), use_container_width=True)
            st.divider()
            st.write("#### 🕒 Detalle de Energía por Hora")
            st.dataframe(df_fact_h[["Fecha y Hora", "Consumo Horario [kWh]"]].style.format({"Consumo Horario [kWh]": "{:,.2f}"}), use_container_width=True)

    st.sidebar.download_button("📥 Descargar Reporte", to_excel_consolidado(df_ops, df_tr, df_seat, df_prmte_d, df_prmte_15, df_fact_h, df_fact_d), "Reporte_SGE_EFE.xlsx")
else:
    st.info("👋 Sube los archivos para comenzar el análisis.")
