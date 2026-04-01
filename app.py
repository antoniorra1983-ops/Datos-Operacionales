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
        for df, name in zip([df_ops, df_trenes, df_seat, df_prmte_d, df_prmte_15, df_fact_h, df_fact_d], 
                            ['Ops', 'Kms', 'SEAT', 'PRMTE_D', 'PRMTE_15', 'Factura_H', 'Factura_D']):
            if not df.empty: df.to_excel(writer, index=False, sheet_name=name)
    return output.getvalue()

# --- 3. PROCESAMIENTO DE ARCHIVOS ---
with st.sidebar:
    st.header("📂 Carga de Archivos")
    f_list = st.file_uploader("Subir todos los archivos (UMR, SEAT, PRMTE, Factura)", type=["xlsx"], accept_multiple_files=True)

all_ops, all_tr, all_seat, all_prmte, all_fact = [], [], [], [], []

if f_list:
    for f in f_list:
        try:
            xl = pd.ExcelFile(f)
            for sn in xl.sheet_names:
                sn_up = sn.upper()
                df_raw = pd.read_excel(f, sheet_name=sn, header=None)
                
                # A. UMR / OPERACIONES
                if any(k in sn_up for k in ['UMR', 'RESUMEN']):
                    h_r = next((i for i in range(len(df_raw)) if any(k in str(df_raw.iloc[i]).upper() for k in ['ODO', 'FECHA'])), None)
                    if h_r is not None:
                        df_p = pd.read_excel(f, sheet_name=sn, header=h_r)
                        df_p.columns = [re.sub(r'[^A-Z]', '', str(c).upper().replace('Ó','O')) for c in df_p.columns]
                        idx_f = next((c for c in df_p.columns if 'FECHA' in c), None)
                        idx_o = next((c for c in df_p.columns if 'ODO' in c and 'ACUM' not in c), None)
                        idx_t = next((c for c in df_p.columns if 'TREN' in c and 'KM' in c), None)
                        if idx_f and idx_o:
                            df_p['_dt'] = pd.to_datetime(df_p[idx_f], errors='coerce')
                            for _, r in df_p.dropna(subset=['_dt']).iterrows():
                                o, tk = parse_latam_number(r[idx_o]), parse_latam_number(r[idx_t])
                                if o > 0:
                                    t_dia = "D/F" if (r['_dt'] in chile_holidays or r['_dt'].weekday() == 6) else ("S" if r['_dt'].weekday() == 5 else "L")
                                    all_ops.append({"Fecha": r['_dt'], "Tipo Día": t_dia, "N° Semana": r['_dt'].isocalendar()[1], "Odómetro [km]": o, "Tren-Km [km]": tk, "UMR [%]": (tk/o*100)})

                # B. ODÓMETRO POR TREN
                if 'ODO' in sn_up and 'KIL' in sn_up:
                    for i in range(min(50, len(df_raw)-2)):
                        for j in range(1, len(df_raw.columns)):
                            p_d = pd.to_datetime(df_raw.iloc[i, j], errors='coerce')
                            if pd.notna(p_d):
                                for k in range(i+3, len(df_raw)):
                                    n_tr = str(df_raw.iloc[k, 0]).strip().upper()
                                    if re.match(r'^(M|XM)', n_tr):
                                        all_tr.append({"Tren": n_tr, "Fecha": p_d, "Kilometraje Diario [km]": parse_latam_number(df_raw.iloc[k, j])})

                # C. ENERGÍA SEAT
                if 'SEAT' in sn_up and 'SER' in sn_up:
                    for i in range(len(df_raw)):
                        fs = pd.to_datetime(df_raw.iloc[i, 1], errors='coerce')
                        if pd.notna(fs):
                            tot, tra, k12 = parse_latam_number(df_raw.iloc[i, 3]), parse_latam_number(df_raw.iloc[i, 5]), parse_latam_number(df_raw.iloc[i, 7])
                            all_seat.append({"Fecha": fs, "Total [kWh]": tot, "Tracción [kWh]": tra, "12 KV [kWh]": k12, "% Tracción": (tra/tot*100 if tot>0 else 0), "% 12 KV": (k12/tot*100 if tot>0 else 0)})

                # D. PRMTE (15 MIN)
                if any(k in sn_up for k in ['PRMTE', 'MEDIDAS']):
                    h_idx = next((i for i in range(len(df_raw)) if 'AÑO' in str(df_raw.iloc[i]).upper()), None)
                    if h_idx is not None:
                        df_prm = pd.read_excel(f, sheet_name=sn, header=h_idx)
                        df_prm['Timestamp'] = pd.to_datetime(df_prm[['AÑO', 'MES', 'DIA', 'HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'})) + pd.to_timedelta(df_prm['INICIO INTERVALO'].astype(int), unit='m')
                        for _, r in df_prm.iterrows():
                            all_prmte.append({"Timestamp": r['Timestamp'], "Fecha": r['Timestamp'].date(), "Energía PRMTE [kWh]": parse_latam_number(r['Retiro_Energia_Activa (kWhD)'])})

                # E. FACTURA
                if any(k in sn_up for k in ['FACTURA', 'CONSUMO']):
                    df_f = pd.read_excel(f, sheet_name=sn)
                    df_f['Timestamp'] = pd.to_datetime(df_f.iloc[:, 0], errors='coerce')
                    for _, r in df_f.dropna(subset=['Timestamp']).iterrows():
                        all_fact.append({"Timestamp": r['Timestamp'], "Fecha": r['Timestamp'].date(), "Consumo Horario [kWh]": abs(parse_latam_number(r.iloc[1]))})
        except: continue

# --- 4. RENDERIZADO CON FILTROS INDEPENDIENTES ---
df_ops_full = pd.DataFrame(all_ops)
df_tr_full = pd.DataFrame(all_tr)
df_seat_full = pd.DataFrame(all_seat)
df_prmte_full = pd.DataFrame(all_prmte)
df_fact_full = pd.DataFrame(all_fact)

tabs = st.tabs(["📊 Resumen", "📑 Datos operacionales", "📑 Odómetro por Tren", "⚡ Energía SEAT", "📈 Medidas PRMTE", "💰 Facturación"])

with tabs[0]: # RESUMEN
    if not df_ops_full.empty:
        r_dates = st.date_input("Filtro Resumen", [df_ops_full['Fecha'].min(), df_ops_full['Fecha'].max()], key="f_res")
        if len(r_dates) == 2:
            df = df_ops_full[(df_ops_full['Fecha'].dt.date >= r_dates[0]) & (df_ops_full['Fecha'].dt.date <= r_dates[1])]
            c1, c2, c3 = st.columns(3)
            to, tk = df["Odómetro [km]"].sum(), df["Tren-Km [km]"].sum()
            c1.metric("Odómetro Total", f"{to:,.1f} km"); c2.metric("Tren-Km Total", f"{tk:,.1f} km"); c3.metric("UMR Global", f"{(tk/to*100 if to>0 else 0):.2f} %")
            st.table(df.groupby("Tipo Día").agg({"Odómetro [km]":"sum", "Tren-Km [km]":"sum", "UMR [%]":"mean"}).reindex(['L', 'S', 'D/F']).style.format("{:,.1f}"))

with tabs[3]: # SEAT
    if not df_seat_full.empty:
        s_dates = st.date_input("Filtro SEAT", [df_seat_full['Fecha'].min(), df_seat_full['Fecha'].max()], key="f_seat")
        if len(s_dates) == 2:
            df = df_seat_full[(df_seat_full['Fecha'].dt.date >= s_dates[0]) & (df_seat_full['Fecha'].dt.date <= s_dates[1])]
            st.dataframe(df.style.format({"Total [kWh]":"{:,.0f}", "% Tracción":"{:.2f}%", "% 12 KV":"{:.2f}%"}), use_container_width=True)

with tabs[4]: # PRMTE
    if not df_prmte_full.empty:
        p_dates = st.date_input("Filtro PRMTE", [df_prmte_full['Timestamp'].min().date(), df_prmte_full['Timestamp'].max().date()], key="f_prmte")
        if len(p_dates) == 2:
            df_sub = df_prmte_full[(df_prmte_full['Timestamp'].dt.date >= p_dates[0]) & (df_prmte_full['Timestamp'].dt.date <= p_dates[1])]
            df_d = df_sub.groupby("Fecha")["Energía PRMTE [kWh]"].sum().reset_index()
            st.write("#### Resumen Diario PRMTE"); st.dataframe(df_d.style.format("{:,.1f}"), use_container_width=True)
            st.write("#### Detalle 15 Minutos"); st.dataframe(df_sub.style.format("{:,.2f}"), use_container_width=True)

with tabs[5]: # FACTURACIÓN
    if not df_fact_full.empty:
        f_dates = st.date_input("Filtro Facturación", [df_fact_full['Timestamp'].min().date(), df_fact_full['Timestamp'].max().date()], key="f_fact")
        if len(f_dates) == 2:
            df_sub = df_fact_full[(df_fact_full['Timestamp'].dt.date >= f_dates[0]) & (df_fact_full['Timestamp'].dt.date <= f_dates[1])]
            df_d = df_sub.groupby("Fecha")["Consumo Horario [kWh]"].sum().reset_index().rename(columns={"Consumo Horario [kWh]":"Total Facturado [kWh]"})
            if not df_seat_full.empty:
                df_d = pd.merge(df_d, df_seat_full[["Fecha", "% Tracción", "% 12 KV"]], left_on="Fecha", right_on=df_seat_full["Fecha"].dt.date, how="left")
                df_d["Tracción Facturada [kWh]"] = df_d["Total Facturado [kWh]"] * (df_d["% Tracción"] / 100)
                df_d["12kV Facturada [kWh]"] = df_d["Total Facturado [kWh]"] * (df_d["% 12 KV"] / 100)
            st.write("#### Distribución Proporcional"); st.dataframe(df_d.style.format("{:,.1f}"), use_container_width=True)

if f_list:
    st.sidebar.download_button("📥 Descargar Excel Consolidado", to_excel_consolidado(df_ops_full, df_tr_full, df_seat_full, df_d, df_prmte_full, df_fact_full, df_d), "Reporte_EFE.xlsx")
