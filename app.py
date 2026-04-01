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

def get_tipo_dia(fch):
    if fch in chile_holidays or fch.weekday() == 6: return "D/F"
    if fch.weekday() == 5: return "S"
    return "L"

def to_excel_consolidado(df_ops, df_tr, df_seat, df_prm_d, df_prm_15, df_fact_h, df_fact_d):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dfs = {
            'Operaciones': df_ops, 'Kilometrajes': df_tr, 'Energia_SEAT': df_seat,
            'PRMTE_Diario': df_prm_d, 'PRMTE_15min': df_prm_15, 
            'Factura_Hora': df_fact_h, 'Factura_Dia': df_fact_d
        }
        for name, df in dfs.items():
            if not df.empty: df.to_excel(writer, index=False, sheet_name=name)
    return output.getvalue()

def color_umr(val):
    if val > 96.4: return 'color: green; font-weight: bold;'
    elif val < 96.4: return 'color: red; font-weight: bold;'
    return 'color: black;'

# --- 3. SIDEBAR (FILTROS GLOBALES) ---
MESES_NOMBRES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
MESES_NUM_MAP = {nombre: i+1 for i, nombre in enumerate(MESES_NOMBRES)}

with st.sidebar:
    st.header("📂 Gestión de Archivos")
    f_umr = st.file_uploader("1. UMR / Odómetros", type=["xlsx"], accept_multiple_files=True)
    f_seat_files = st.file_uploader("2. Energía SEAT", type=["xlsx"], accept_multiple_files=True)
    f_bill_files = st.file_uploader("3. Facturación y PRMTE", type=["xlsx"], accept_multiple_files=True)
    
    st.divider()
    st.header("🎯 Filtros Globales")
    f_anio_list = st.multiselect("Años", [2024, 2025, 2026], default=[2025, 2026])
    f_mes_list = st.multiselect("Meses", MESES_NOMBRES, default=MESES_NOMBRES)
    meses_num = [MESES_NUM_MAP[m] for m in f_mes_list]
    f_dias = st.multiselect("Días", list(range(1, 32)), default=list(range(1, 32)))

# --- 4. PROCESAMIENTO ---
all_ops, all_tr, all_seat, all_prmte_15, all_fact_h = [], [], [], [], []
todos_archivos = (f_umr or []) + (f_seat_files or []) + (f_bill_files or [])

for f in todos_archivos:
    try:
        xl = pd.ExcelFile(f)
        for sn in xl.sheet_names:
            sn_up = sn.upper()
            
            # A. UMR / OPERACIONES
            if any(k in sn_up for k in ['UMR', 'RESUMEN']):
                df_raw = pd.read_excel(f, sheet_name=sn, header=None)
                h_r = next((i for i in range(min(100, len(df_raw))) if any(k in str(df_raw.iloc[i]).upper() for k in ['ODO', 'FECHA', 'TREN'])), None)
                if h_r is not None:
                    df_p = pd.read_excel(f, sheet_name=sn, header=h_r)
                    df_p.columns = [re.sub(r'[^A-Z]', '', str(c).upper().replace('Ó','O')) for c in df_p.columns]
                    idx_f = next((c for c in df_p.columns if 'FECHA' in c), None)
                    idx_o = next((c for c in df_p.columns if 'ODO' in c and 'ACUM' not in c), None)
                    idx_t = next((c for c in df_p.columns if 'TREN' in c and 'KM' in c), None)
                    if idx_f and idx_o:
                        df_p['_dt'] = pd.to_datetime(df_p[idx_f], errors='coerce')
                        mask = (df_p['_dt'].dt.year.isin(f_anio_list)) & (df_p['_dt'].dt.month.isin(meses_num)) & (df_p['_dt'].dt.day.isin(f_dias))
                        for _, r in df_p[mask].dropna(subset=['_dt']).iterrows():
                            o, tk = parse_latam_number(r[idx_o]), parse_latam_number(r[idx_t])
                            if o > 0:
                                t_dia = get_tipo_dia(r['_dt'])
                                all_ops.append({"Fecha": r['_dt'], "Tipo Día": t_dia, "N° Semana": r['_dt'].isocalendar()[1], "Odómetro [km]": o, "Tren-Km [km]": tk, "UMR [%]": (tk/o*100), "Timestamp": r['_dt']})

            # B. ODÓMETRO POR TREN
            if 'ODO' in sn_up and 'KIL' in sn_up:
                df_tr_raw = pd.read_excel(f, sheet_name=sn, header=None)
                for i in range(min(100, len(df_tr_raw)-2)):
                    for j in range(1, len(df_tr_raw.columns)):
                        p_d = pd.to_datetime(df_tr_raw.iloc[i, j], errors='coerce')
                        if pd.notna(p_d) and p_d.year in f_anio_list and p_d.month in meses_num:
                            for k in range(i+3, len(df_tr_raw)):
                                n_tr = str(df_tr_raw.iloc[k, 0]).strip().upper()
                                if re.match(r'^(M|XM)', n_tr):
                                    all_tr.append({"Tren": n_tr, "Fecha": p_d, "Día": p_d.day, "Kilometraje [km]": parse_latam_number(df_tr_raw.iloc[k, j]), "Timestamp": p_d})

            # C. ENERGÍA SEAT
            if 'SEAT' in sn_up and 'SER' in sn_up:
                df_s = pd.read_excel(f, sheet_name=sn, header=None)
                for i in range(len(df_s)):
                    fs = pd.to_datetime(df_s.iloc[i, 1], errors='coerce')
                    if pd.notna(fs) and fs.year in f_anio_list and fs.month in meses_num:
                        tot, tra, k12 = parse_latam_number(df_s.iloc[i, 3]), parse_latam_number(df_s.iloc[i, 5]), parse_latam_number(df_s.iloc[i, 7])
                        all_seat.append({"Fecha": fs, "Tipo Día": get_tipo_dia(fs), "Total [kWh]": tot, "Tracción [kWh]": tra, "12 KV [kWh]": k12, "% Tracción": (tra/tot*100 if tot>0 else 0), "% 12 KV": (k12/tot*100 if tot>0 else 0), "Timestamp": fs})

            # D. PRMTE
            if any(k in sn_up for k in ['PRMTE', 'MEDIDAS']):
                df_prm = pd.read_excel(f, sheet_name=sn, header=None)
                h_idx = next((i for i in range(len(df_prm)) if 'AÑO' in str(df_prm.iloc[i]).upper()), None)
                if h_idx is not None:
                    df_prm_data = pd.read_excel(f, sheet_name=sn, header=h_idx)
                    df_prm_data['Timestamp'] = pd.to_datetime(df_prm_data[['AÑO', 'MES', 'DIA', 'HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'})) + pd.to_timedelta(df_prm_data['INICIO INTERVALO'].astype(int), unit='m')
                    mask_p = (df_prm_data['Timestamp'].dt.year.isin(f_anio_list)) & (df_prm_data['Timestamp'].dt.month.isin(meses_num))
                    for _, r in df_prm_data[mask_p].iterrows():
                        all_prmte_15.append({"Fecha y Hora": r['Timestamp'].strftime('%d/%m/%Y %H:%M'), "Fecha": r['Timestamp'].date(), "Energía PRMTE [kWh]": parse_latam_number(r['Retiro_Energia_Activa (kWhD)']), "Timestamp": r['Timestamp']})

            # E. FACTURA
            if any(k in sn_up for k in ['FACTURA', 'CONSUMO']):
                df_f = pd.read_excel(f, sheet_name=sn)
                df_f.columns = ['FechaHora', 'Valor']
                df_f['Timestamp'] = pd.to_datetime(df_f['FechaHora'], errors='coerce')
                mask_f = (df_f['Timestamp'].dt.year.isin(f_anio_list)) & (df_f['Timestamp'].dt.month.isin(meses_num))
                for _, r in df_f[mask_f].dropna(subset=['Timestamp']).iterrows():
                    all_fact_h.append({"Fecha y Hora": r['Timestamp'].strftime('%d/%m/%Y %H:%M'), "Fecha": r['Timestamp'].date(), "Consumo Horario [kWh]": abs(parse_latam_number(r['Valor'])), "Timestamp": r['Timestamp']})
    except: continue

# --- 5. RENDERIZADO (POST-PROCESAMIENTO) ---
df_ops, df_tr, df_seat, df_prm_15, df_prm_d, df_fact_h, df_fact_d = [pd.DataFrame()] * 7

if any([all_ops, all_tr, all_seat, all_prmte_15, all_fact_h]):
    if all_ops: df_ops = pd.DataFrame(all_ops).drop_duplicates(subset=['Fecha']).sort_values("Timestamp")
    if all_tr: df_tr = pd.DataFrame(all_tr).sort_values(["Timestamp", "Tren"])
    if all_seat: df_seat = pd.DataFrame(all_seat).drop_duplicates(subset=['Fecha']).sort_values("Timestamp")
    
    # Procesar PRMTE con Proporcional SEAT
    if all_prmte_15:
        df_prm_15 = pd.DataFrame(all_prmte_15).sort_values("Timestamp")
        df_prm_d = df_prm_15.groupby("Fecha")["Energía PRMTE [kWh]"].sum().reset_index().rename(columns={"Energía PRMTE [kWh]":"Total Diario PRMTE [kWh]"})
        if not df_seat.empty:
            df_prm_d['Fecha_dt'] = pd.to_datetime(df_prm_d['Fecha'])
            df_seat['Fecha_dt'] = pd.to_datetime(df_seat['Fecha'])
            df_prm_d = pd.merge(df_prm_d, df_seat[["Fecha_dt", "% Tracción", "% 12 KV"]], on="Fecha_dt", how="left")
            df_prm_d["Energía Tracción PRMTE [kWh]"] = df_prm_d["Total Diario PRMTE [kWh]"] * (df_prm_d["% Tracción"] / 100)
            df_prm_d["Energía 12kV PRMTE [kWh]"] = df_prm_d["Total Diario PRMTE [kWh]"] * (df_prm_d["% 12 KV"] / 100)
            df_prm_d['Tipo Día'] = df_prm_d['Fecha_dt'].apply(get_tipo_dia)

    # Procesar FACTURA con Proporcional SEAT
    if all_fact_h:
        df_fact_h = pd.DataFrame(all_fact_h).sort_values("Timestamp")
        df_fact_d = df_fact_h.groupby("Fecha")["Consumo Horario [kWh]"].sum().reset_index().rename(columns={"Consumo Horario [kWh]":"Total Factura [kWh]"})
        if not df_seat.empty:
            df_fact_d['Fecha_dt'] = pd.to_datetime(df_fact_d['Fecha'])
            df_seat['Fecha_dt'] = pd.to_datetime(df_seat['Fecha'])
            df_fact_d = pd.merge(df_fact_d, df_seat[["Fecha_dt", "% Tracción", "% 12 KV"]], on="Fecha_dt", how="left")
            df_fact_d["Energía Tracción Facturación [kWh]"] = df_fact_d["Total Factura [kWh]"] * (df_fact_d["% Tracción"] / 100)
            df_fact_d["Energía 12kV Facturación [kWh]"] = df_fact_d["Total Factura [kWh]"] * (df_fact_d["% 12 KV"] / 100)
            df_fact_d['Tipo Día'] = df_fact_d['Fecha_dt'].apply(get_tipo_dia)

    tabs = st.tabs(["📊 Resumen", "📑 Datos operacionales", "📑 Odómetro por Tren", "⚡ Energía SEAT", "📈 Medidas PRMTE", "💰 Facturación"])
    
    with tabs[0]: # PESTAÑA RESUMEN (NUEVA LÓGICA)
        if not df_ops.empty:
            st.subheader("Indicadores Globales")
            to, tk = df_ops["Odómetro [km]"].sum(), df_ops["Tren-Km [km]"].sum()
            c1, c2, c3 = st.columns(3)
            c1.metric("Odómetro Total", f"{to:,.1f} km"); c2.metric("Tren-Km Total", f"{tk:,.1f} km"); c3.metric("UMR Global", f"{(tk/to*100 if to>0 else 0):.2f} %")
            
            st.divider()
            st.write("### Desempeño Operacional y Energético por Jornada")
            
            # Base Operacional
            df_ops['Tipo Día'] = pd.Categorical(df_ops['Tipo Día'], categories=['L', 'S', 'D/F'], ordered=True)
            res_final = df_ops.groupby("Tipo Día").agg({"Odómetro [km]":"sum", "Tren-Km [km]":"sum", "UMR [%]":"mean"}).reset_index()
            
            # Jerarquía de Energía para el Resumen
            energy_label = ""
            if not df_fact_d.empty:
                e_sum = df_fact_d.groupby("Tipo Día")[["Total Factura [kWh]", "Energía Tracción Facturación [kWh]", "Energía 12kV Facturación [kWh]"]].sum().reset_index()
                res_final = pd.merge(res_final, e_sum, on="Tipo Día", how="left")
                energy_label = "(Basado en Facturación)"
                col_rename = {"Total Factura [kWh]": "Energía Total [kWh]", "Energía Tracción Facturación [kWh]": "Tracción [kWh]", "Energía 12kV Facturación [kWh]": "12 kV [kWh]"}
            elif not df_prm_d.empty:
                e_sum = df_prm_d.groupby("Tipo Día")[["Total Diario PRMTE [kWh]", "Energía Tracción PRMTE [kWh]", "Energía 12kV PRMTE [kWh]"]].sum().reset_index()
                res_final = pd.merge(res_final, e_sum, on="Tipo Día", how="left")
                energy_label = "(Basado en PRMTE)"
                col_rename = {"Total Diario PRMTE [kWh]": "Energía Total [kWh]", "Energía Tracción PRMTE [kWh]": "Tracción [kWh]", "Energía 12kV PRMTE [kWh]": "12 kV [kWh]"}
            elif not df_seat.empty:
                e_sum = df_seat.groupby("Tipo Día")[["Total [kWh]", "Tracción [kWh]", "12 KV [kWh]"]].sum().reset_index()
                res_final = pd.merge(res_final, e_sum, on="Tipo Día", how="left")
                energy_label = "(Basado en SEAT)"
                col_rename = {"Total [kWh]": "Energía Total [kWh]", "Tracción [kWh]": "Tracción [kWh]", "12 KV [kWh]": "12 kV [kWh]"}
            else:
                col_rename = {}

            if col_rename:
                res_final = res_final.rename(columns=col_rename)
                st.info(f"Visualizando datos de energía provenientes de: **{energy_label}**")
            
            st.table(res_final.style.format({
                "Odómetro [km]":"{:,.1f}", "Tren-Km [km]":"{:,.1f}", "UMR [%]":"{:.2f}%",
                "Energía Total [kWh]":"{:,.0f}", "Tracción [kWh]":"{:,.0f}", "12 kV [kWh]":"{:,.0f}"
            }))

    with tabs[4]: # MEDIDAS PRMTE
        if not df_prm_15.empty:
            st.write("#### 📅 Resumen Diario PRMTE (Proporcional SEAT)")
            st.dataframe(df_prm_d.style.format({
                "Total Diario PRMTE [kWh]": "{:,.1f}", "Energía Tracción PRMTE [kWh]": "{:,.1f}", "Energía 12kV PRMTE [kWh]": "{:,.1f}",
                "% Tracción": "{:.2f}%", "% 12 KV": "{:.2f}%"
            }), use_container_width=True)
            st.divider()
            st.write("#### 🕒 Detalle cada 15 Minutos")
            st.dataframe(df_prm_15[["Fecha y Hora", "Energía PRMTE [kWh]"]].style.format({"Energía PRMTE [kWh]": "{:,.2f}"}), use_container_width=True)

    with tabs[5]: # FACTURACIÓN
        if not df_fact_h.empty:
            st.write("#### 📅 Distribución Proporcional Facturación (kWh)"); 
            st.dataframe(df_fact_d[["Fecha", "Total Factura [kWh]", "Energía Tracción Facturación [kWh]", "Energía 12kV Facturación [kWh]"]].style.format({
                "Total Factura [kWh]":"{:,.1f}", "Energía Tracción Facturación [kWh]":"{:,.1f}", "Energía 12kV Facturación [kWh]":"{:,.1f}"
            }), use_container_width=True)
            st.divider()
            st.write("#### 🕒 Detalle Horario")
            st.dataframe(df_fact_h[["Fecha y Hora", "Consumo Horario [kWh]"]].style.format({"Consumo Horario [kWh]": "{:,.2f}"}), use_container_width=True)

    st.sidebar.download_button("📥 Reporte Consolidado", to_excel_consolidado(df_ops, df_tr, df_seat, df_prm_d, df_prm_15, df_fact_h, df_fact_d), "Reporte_SGE_EFE.xlsx")
else:
    st.info("👋 Sube los archivos para comenzar el análisis.")
