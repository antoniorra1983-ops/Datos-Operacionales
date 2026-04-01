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

def to_excel_consolidado(df_ops, df_tr, df_seat, df_prm_d, df_prm_15, df_fact_h, df_fact_d):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dict_dfs = {
            'Operaciones': df_ops, 'Kilometrajes': df_tr, 'Energia_SEAT': df_seat,
            'PRMTE_Diario': df_prm_d, 'PRMTE_15min': df_prm_15, 
            'Factura_Hora': df_fact_h, 'Factura_Dia': df_fact_d
        }
        for name, df in dict_dfs.items():
            if not df.empty: df.to_excel(writer, index=False, sheet_name=name)
    return output.getvalue()

# --- 3. PROCESAMIENTO INICIAL ---
# Inicialización de DataFrames maestros (vacíos por defecto para evitar NameError)
df_ops_master = pd.DataFrame()
df_tr_master = pd.DataFrame()
df_seat_master = pd.DataFrame()
df_prmte_master = pd.DataFrame()
df_fact_master = pd.DataFrame()

with st.sidebar:
    st.header("📂 Carga de Archivos")
    f_list = st.file_uploader("Subir archivos (UMR, SEAT, PRMTE, Factura)", type=["xlsx"], accept_multiple_files=True)

if f_list:
    ops_l, tr_l, seat_l, prm_l, fact_l = [], [], [], [], []
    for f in f_list:
        try:
            xl = pd.ExcelFile(f)
            for sn in xl.sheet_names:
                sn_up = sn.upper()
                # A. UMR / OPERACIONES
                if any(k in sn_up for k in ['UMR', 'RESUMEN']):
                    df_tmp = pd.read_excel(f, sheet_name=sn, header=None)
                    h_r = next((i for i in range(min(50, len(df_tmp))) if any(k in str(df_tmp.iloc[i]).upper() for k in ['ODO', 'FECHA'])), None)
                    if h_r is not None:
                        df_p = pd.read_excel(f, sheet_name=sn, header=h_r)
                        df_p.columns = [re.sub(r'[^A-Z]', '', str(c).upper().replace('Ó','O')) for c in df_p.columns]
                        idx_f = next((c for c in df_p.columns if 'FECHA' in c), None)
                        idx_o = next((c for c in df_p.columns if 'ODO' in c and 'ACUM' not in c), None)
                        idx_t = next((c for c in df_p.columns if 'TREN' in c and 'KM' in c), None)
                        if idx_f and idx_o:
                            df_p['Timestamp'] = pd.to_datetime(df_p[idx_f], errors='coerce')
                            for _, r in df_p.dropna(subset=['Timestamp']).iterrows():
                                o, tk = parse_latam_number(r[idx_o]), parse_latam_number(r[idx_t])
                                if o > 0:
                                    t_dia = "D/F" if (r['Timestamp'] in chile_holidays or r['Timestamp'].weekday() == 6) else ("S" if r['Timestamp'].weekday() == 5 else "L")
                                    ops_l.append({"Fecha": r['Timestamp'], "Tipo Día": t_dia, "Odómetro [km]": o, "Tren-Km [km]": tk, "UMR [%]": (tk/o*100)})

                # B. ODÓMETRO POR TREN
                if 'ODO' in sn_up and 'KIL' in sn_up:
                    df_raw = pd.read_excel(f, sheet_name=sn, header=None)
                    for i in range(min(40, len(df_raw)-2)):
                        for j in range(1, len(df_raw.columns)):
                            p_d = pd.to_datetime(df_raw.iloc[i, j], errors='coerce')
                            if pd.notna(p_d):
                                for k in range(i+3, len(df_raw)):
                                    n_tr = str(df_raw.iloc[k, 0]).strip().upper()
                                    if re.match(r'^(M|XM)', n_tr):
                                        tr_l.append({"Tren": n_tr, "Fecha": p_d, "Kilometraje [km]": parse_latam_number(df_raw.iloc[k, j])})

                # C. ENERGÍA SEAT
                if 'SEAT' in sn_up and 'SER' in sn_up:
                    df_s = pd.read_excel(f, sheet_name=sn, header=None)
                    for i in range(len(df_s)):
                        fs = pd.to_datetime(df_s.iloc[i, 1], errors='coerce')
                        if pd.notna(fs):
                            tot, tra, k12 = parse_latam_number(df_s.iloc[i, 3]), parse_latam_number(df_s.iloc[i, 5]), parse_latam_number(df_s.iloc[i, 7])
                            seat_l.append({"Fecha": fs, "Total [kWh]": tot, "Tracción [kWh]": tra, "12 KV [kWh]": k12, "% Tracción": (tra/tot*100 if tot>0 else 0), "% 12 KV": (k12/tot*100 if tot>0 else 0)})

                # D. PRMTE (15 MIN)
                if any(k in sn_up for k in ['PRMTE', 'MEDIDAS']):
                    df_prm = pd.read_excel(f, sheet_name=sn, header=None)
                    h_idx = next((i for i in range(len(df_prm)) if 'AÑO' in str(df_prm.iloc[i]).upper()), None)
                    if h_idx is not None:
                        df_prm_data = pd.read_excel(f, sheet_name=sn, header=h_idx)
                        df_prm_data['Timestamp'] = pd.to_datetime(df_prm_data[['AÑO', 'MES', 'DIA', 'HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'})) + pd.to_timedelta(df_prm_data['INICIO INTERVALO'].astype(int), unit='m')
                        for _, r in df_prm_data.iterrows():
                            p_val = parse_latam_number(r['Retiro_Energia_Activa (kWhD)'])
                            if p_val > 0: p_l.append({"Timestamp": r['Timestamp'], "Fecha": r['Timestamp'].date(), "Energía PRMTE [kWh]": p_val})

                # E. FACTURA
                if any(k in sn_up for k in ['FACTURA', 'CONSUMO']):
                    df_f = pd.read_excel(f, sheet_name=sn)
                    df_f.columns = ['FechaHora', 'Valor']
                    df_f['Timestamp'] = pd.to_datetime(df_f['FechaHora'], errors='coerce')
                    for _, r in df_f.dropna(subset=['Timestamp']).iterrows():
                        fact_l.append({"Timestamp": r['Timestamp'], "Fecha": r['Timestamp'].date(), "Consumo Horario [kWh]": abs(parse_latam_number(r['Valor']))})
        except: continue

    df_ops_master = pd.DataFrame(ops_l)
    df_tr_master = pd.DataFrame(tr_l)
    df_seat_master = pd.DataFrame(seat_l)
    df_prmte_master = pd.DataFrame(p_l)
    df_fact_master = pd.DataFrame(fact_l)

# --- 4. RENDERIZADO CON FILTROS INDEPENDIENTES ---
tabs = st.tabs(["📊 Resumen", "📑 Datos operacionales", "📑 Odómetro por Tren", "⚡ Energía SEAT", "📈 Medidas PRMTE", "💰 Facturación"])

# Definimos variables de salida para el Excel (evita NameError)
df_fact_d_final = pd.DataFrame()
df_prmte_d_final = pd.DataFrame()

with tabs[0]: # RESUMEN
    if not df_ops_master.empty:
        r_f = st.date_input("Filtrar Resumen", [df_ops_master['Fecha'].min(), df_ops_master['Fecha'].max()], key="res_f")
        if len(r_f) == 2:
            df = df_ops_master[(df_ops_master['Fecha'].dt.date >= r_f[0]) & (df_ops_master['Fecha'].dt.date <= r_f[1])]
            to, tk = df["Odómetro [km]"].sum(), df["Tren-Km [km]"].sum()
            c1, c2, c3 = st.columns(3)
            c1.metric("Odómetro Total", f"{to:,.0f} km"); c2.metric("Tren-Km Total", f"{tk:,.0f} km"); c3.metric("UMR Global", f"{(tk/to*100 if to>0 else 0):.2f} %")
            st.table(df.groupby("Tipo Día").agg({"Odómetro [km]":"sum", "Tren-Km [km]":"sum", "UMR [%]":"mean"}).reindex(['L','S','D/F']).style.format("{:,.1f}"))

with tabs[2]: # ODÓMETRO POR TREN
    if not df_tr_master.empty:
        tr_f = st.date_input("Filtrar Odómetros", [df_tr_master['Fecha'].min(), df_tr_master['Fecha'].max()], key="tr_f")
        if len(tr_f) == 2:
            df = df_tr_master[(df_tr_master['Fecha'].dt.date >= tr_f[0]) & (df_tr_master['Fecha'].dt.date <= tr_f[1])]
            df['Día'] = df['Fecha'].dt.day
            st.dataframe(df.pivot_table(index="Tren", columns="Día", values="Kilometraje [km]", aggfunc='sum').fillna(0).style.format("{:,.1f}"), use_container_width=True)

with tabs[3]: # SEAT
    if not df_seat_master.empty:
        st.dataframe(df_seat_master.style.format({"Total [kWh]":"{:,.0f}", "% Tracción":"{:.2f}%"}), use_container_width=True)

with tabs[4]: # PRMTE
    if not df_prmte_master.empty:
        p_f = st.date_input("Filtrar PRMTE", [df_prmte_master['Timestamp'].min().date(), df_prmte_master['Timestamp'].max().date()], key="prm_f")
        if len(p_f) == 2:
            df_sub = df_prmte_master[(df_prmte_master['Timestamp'].dt.date >= p_f[0]) & (df_prmte_master['Timestamp'].dt.date <= p_f[1])]
            df_prmte_d_final = df_sub.groupby("Fecha")["Energía PRMTE [kWh]"].sum().reset_index()
            st.write("#### 📅 Consumo Diario PRMTE")
            st.dataframe(df_prmte_d_final.style.format("{:,.1f}"), use_container_width=True)
            st.write("#### 🕒 Detalle cada 15 Minutos")
            st.dataframe(df_sub[["Timestamp", "Energía PRMTE [kWh]"]].style.format("{:,.2f}"), use_container_width=True)

with tabs[5]: # FACTURACIÓN
    if not df_fact_master.empty:
        f_f = st.date_input("Filtrar Facturación", [df_fact_master['Timestamp'].min().date(), df_fact_master['Timestamp'].max().date()], key="fact_f")
        if len(f_f) == 2:
            df_sub = df_fact_master[(df_fact_master['Timestamp'].dt.date >= f_f[0]) & (df_fact_master['Timestamp'].dt.date <= f_f[1])]
            df_fact_d_final = df_sub.groupby("Fecha")["Consumo Horario [kWh]"].sum().reset_index().rename(columns={"Consumo Horario [kWh]":"Factura Total [kWh]"})
            if not df_seat_master.empty:
                df_fact_d_final = pd.merge(df_fact_d_final, df_seat_master[["Fecha", "% Tracción", "% 12 KV"]], on="Fecha", how="left")
                df_fact_d_final["Energía Tracción Facturación"] = df_fact_d_final["Factura Total [kWh]"] * (df_fact_d_final["% Tracción"]/100)
                df_fact_d_final["Energía 12kV Facturación"] = df_fact_d_final["Factura Total [kWh]"] * (df_fact_d_final["% 12 KV"]/100)
            st.write("#### 📅 Resumen Diario Facturación Proporcional")
            st.dataframe(df_fact_d_final.style.format("{:,.1f}"), use_container_width=True)
            st.write("#### 🕒 Detalle Horario")
            st.dataframe(df_sub[["Timestamp", "Consumo Horario [kWh]"]].style.format("{:,.2f}"), use_container_width=True)

# --- BOTÓN DE DESCARGA CON VARIABLES PROTEGIDAS ---
if f_list:
    st.sidebar.download_button(
        "📥 Descargar Reporte EFE", 
        to_excel_consolidado(df_ops_master, df_tr_master, df_seat_master, df_prmte_d_final, df_prmte_master, df_fact_master, df_fact_d_final), 
        "Reporte_Consolidado.xlsx"
    )
