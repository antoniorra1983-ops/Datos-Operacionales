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
            'Datos_Operacionales': df_ops, 'Kilometraje_Trenes': df_tr, 
            'Energia_SEAT': df_seat, 'PRMTE_Diario': df_prm_d, 
            'PRMTE_15min': df_prm_15, 'Factura_Detalle': df_fact_h, 'Factura_Diaria': df_fact_d
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
                    idx_f, idx_o, idx_t = next((c for c in df_p.columns if 'FECHA' in c), None), next((c for c in df_p.columns if 'ODO' in c and 'ACUM' not in c), None), next((c for c in df_p.columns if 'TREN' in c and 'KM' in c), None)
                    if idx_f and idx_o:
                        df_p['_dt'] = pd.to_datetime(df_p[idx_f], errors='coerce')
                        mask = (df_p['_dt'].dt.year.isin(f_anio_list)) & (df_p['_dt'].dt.month.isin(meses_num)) & (df_p['_dt'].dt.day.isin(f_dias))
                        for _, r in df_p[mask].dropna(subset=['_dt']).iterrows():
                            o, tk = parse_latam_number(r[idx_o]), parse_latam_number(r[idx_t])
                            if o > 0:
                                all_ops.append({"Fecha": r['_dt'], "Tipo Día": get_tipo_dia(r['_dt']), "N° Semana": r['_dt'].isocalendar()[1], "Odómetro [km]": o, "Tren-Km [km]": tk, "UMR [%]": (tk/o*100), "Timestamp": r['_dt']})

            # B. ODÓMETRO POR TREN (RECUPERADO)
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

            # C. ENERGÍA SEAT (RECUPERADO)
            if 'SEAT' in sn_up and 'SER' in sn_up:
                df_s = pd.read_excel(f, sheet_name=sn, header=None)
                for i in range(len(df_s)):
                    fs = pd.to_datetime(df_s.iloc[i, 1], errors='coerce')
                    if pd.notna(fs) and fs.year in f_anio_list and fs.month in meses_num:
                        tot, tra, k12 = parse_latam_number(df_s.iloc[i, 3]), parse_latam_number(df_s.iloc[i, 5]), parse_latam_number(df_s.iloc[i, 7])
                        all_seat.append({"Fecha": fs, "Total [kWh]": tot, "Tracción [kWh]": tra, "12 KV [kWh]": k12, "% Tracción": (tra/tot*100 if tot>0 else 0), "% 12 KV": (k12/tot*100 if tot>0 else 0), "Timestamp": fs})

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

# --- 5. LÓGICA DE JERARQUÍA Y RENDERIZADO ---
df_ops, df_tr, df_seat, df_prm_d, df_fact_d = [pd.DataFrame()] * 5

if any([all_ops, all_tr, all_seat, all_prmte_15, all_fact_h]):
    # Preparar DataFrames base
    if all_ops: df_ops = pd.DataFrame(all_ops).drop_duplicates(subset=['Fecha']).sort_values("Timestamp")
    if all_tr: df_tr = pd.DataFrame(all_tr).sort_values(["Timestamp", "Tren"])
    if all_seat: df_seat = pd.DataFrame(all_seat).drop_duplicates(subset=['Fecha']).sort_values("Timestamp")
    
    # Consolidar Energía Diaria por fuente para la jerarquía
    df_energy_master = pd.DataFrame()
    
    # 1. SEAT (Base)
    if not df_seat.empty:
        df_energy_master = df_seat[["Fecha", "Total [kWh]", "Tracción [kWh]", "12 KV [kWh]"]].copy()
        df_energy_master.columns = ["Fecha", "E_Total", "E_Trac", "E_12kV"]
        df_energy_master["Fuente"] = "SEAT"

    # 2. PRMTE (Prioridad 2)
    if all_prmte_15:
        df_prm_15 = pd.DataFrame(all_prmte_15)
        df_prm_d = df_prm_15.groupby("Fecha")["Energía PRMTE [kWh]"].sum().reset_index()
        df_prm_d["Fecha"] = pd.to_datetime(df_prm_d["Fecha"])
        if not df_seat.empty:
            df_prm_d = pd.merge(df_prm_d, df_seat[["Fecha", "% Tracción", "% 12 KV"]], on="Fecha", how="left")
            df_prm_d["E_Trac"] = df_prm_d["Energía PRMTE [kWh]"] * (df_prm_d["% Tracción"] / 100)
            df_prm_d["E_12kV"] = df_prm_d["Energía PRMTE [kWh]"] * (df_prm_d["% 12 KV"] / 100)
            df_prm_d = df_prm_d.rename(columns={"Energía PRMTE [kWh]": "E_Total"})[["Fecha", "E_Total", "E_Trac", "E_12kV"]]
            df_prm_d["Fuente"] = "PRMTE"
            # Actualizar master con PRMTE donde exista
            df_energy_master = pd.concat([df_energy_master, df_prm_d]).drop_duplicates(subset=["Fecha"], keep="last")

    # 3. FACTURA (Prioridad 1 - Máxima)
    if all_fact_h:
        df_f_h = pd.DataFrame(all_fact_h)
        df_fact_d = df_f_h.groupby("Fecha")["Consumo Horario [kWh]"].sum().reset_index()
        df_fact_d["Fecha"] = pd.to_datetime(df_fact_d["Fecha"])
        if not df_seat.empty:
            df_fact_d = pd.merge(df_fact_d, df_seat[["Fecha", "% Tracción", "% 12 KV"]], on="Fecha", how="left")
            df_fact_d["E_Trac"] = df_fact_d["Consumo Horario [kWh]"] * (df_fact_d["% Tracción"] / 100)
            df_fact_d["E_12kV"] = df_fact_d["Consumo Horario [kWh]"] * (df_fact_d["% 12 KV"] / 100)
            df_fact_d = df_fact_d.rename(columns={"Consumo Horario [kWh]": "E_Total"})[["Fecha", "E_Total", "E_Trac", "E_12kV"]]
            df_fact_d["Fuente"] = "Factura"
            # Actualizar master con Factura donde exista
            df_energy_master = pd.concat([df_energy_master, df_fact_d]).drop_duplicates(subset=["Fecha"], keep="last")

    # Integrar Energía en Datos Operacionales
    if not df_ops.empty and not df_energy_master.empty:
        df_ops = pd.merge(df_ops, df_energy_master, on="Fecha", how="left")

    # --- PESTAÑAS ---
    t_res, t_ops, t_tr, t_seat, t_prmte, t_fact = st.tabs(["📊 Resumen", "📑 Datos operacionales", "📑 Odómetro por Tren", "⚡ Energía SEAT", "📈 Medidas PRMTE", "💰 Facturación"])
    
    with t_res:
        if not df_ops.empty:
            to, tk = df_ops["Odómetro [km]"].sum(), df_ops["Tren-Km [km]"].sum()
            c1, c2, c3 = st.columns(3)
            c1.metric("Odómetro Total", f"{to:,.1f} km"); c2.metric("Tren-Km Total", f"{tk:,.1f} km"); c3.metric("UMR Global", f"{(tk/to*100 if to>0 else 0):.2f} %")
            st.divider()
            df_ops['Tipo Día'] = pd.Categorical(df_ops['Tipo Día'], categories=['L', 'S', 'D/F'], ordered=True)
            res = df_ops.groupby("Tipo Día").agg({"Odómetro [km]":"sum", "Tren-Km [km]":"sum", "UMR [%]":"mean", "E_Total":"sum", "E_Trac":"sum", "E_12kV":"sum"}).reset_index()
            res = res.rename(columns={"E_Total":"Energía Total [kWh]", "E_Trac":"Tracción [kWh]", "E_12kV":"12 kV [kWh]"})
            st.table(res.style.format({"Odómetro [km]":"{:,.1f}", "Tren-Km [km]":"{:,.1f}", "UMR [%]":"{:.2f}%", "Energía Total [kWh]":"{:,.0f}", "Tracción [kWh]":"{:,.0f}", "12 kV [kWh]":"{:,.0f}"}))

    with t_ops:
        if not df_ops.empty:
            cols_v = ["Fecha", "Tipo Día", "N° Semana", "Odómetro [km]", "Tren-Km [km]", "UMR [%]", "E_Total", "E_Trac", "E_12kV", "Fuente"]
            st.dataframe(df_ops[cols_v].rename(columns={"E_Total":"Energía [kWh]", "E_Trac":"Tracción [kWh]", "E_12kV":"12 kV [kWh]"}).style.format({
                "Odómetro [km]":"{:,.1f}", "Tren-Km [km]":"{:,.1f}", "UMR [%]":"{:.2f}%", "Energía [kWh]":"{:,.0f}", "Tracción [kWh]":"{:,.0f}", "12 kV [kWh]":"{:,.0f}"
            }).applymap(color_umr, subset=['UMR [%]']), use_container_width=True)

    with t_tr:
        if not df_tr.empty: st.dataframe(df_tr.pivot_table(index="Tren", columns="Día", values="Kilometraje [km]", aggfunc='sum').fillna(0).style.format("{:,.1f}"), use_container_width=True)

    with t_seat:
        if not df_seat.empty: st.dataframe(df_seat[["Fecha", "Total [kWh]", "Tracción [kWh]", "% Tracción", "12 KV [kWh]", "% 12 KV"]].style.format({"Total [kWh]":"{:,.0f}", "% Tracción":"{:.2f}%", "% 12 KV":"{:.2f}%"}), use_container_width=True)

    with t_prmte:
        if not df_prm_d.empty:
            st.write("#### 📅 Resumen Diario PRMTE"); st.dataframe(df_prm_d.style.format("{:,.1f}"), use_container_width=True)
            st.write("#### 🕒 Detalle cada 15 Minutos"); st.dataframe(pd.DataFrame(all_prmte_15)[["Fecha y Hora", "Energía PRMTE [kWh]"]].style.format("{:,.2f}"), use_container_width=True)

    with t_fact:
        if not df_fact_d.empty:
            st.write("#### 📅 Resumen Diario Facturación"); st.dataframe(df_fact_d.style.format("{:,.1f}"), use_container_width=True)

    st.sidebar.download_button("📥 Descargar Reporte", to_excel_consolidado(df_ops, df_tr, df_seat, df_prm_d, pd.DataFrame(all_prmte_15), pd.DataFrame(all_fact_h), df_fact_d), "Reporte_EFE_Consolidado.xlsx")
else:
    st.info("👋 Sube los archivos para comenzar el análisis.")
