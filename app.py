import streamlit as st
import pandas as pd
import numpy as np
import re
import holidays
from io import BytesIO
from datetime import datetime, date

# --- 1. CONFIGURACIÓN Y ESTILO ---
st.set_page_config(page_title="EFE SGE - Control Energético", layout="wide", page_icon="🚆")
chile_holidays = holidays.Chile()

# Estilo CSS para tarjetas y métricas
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stMetric { 
        background-color: #ffffff; padding: 20px; border-radius: 12px; 
        border-top: 4px solid #005195; box-shadow: 0 4px 6px rgba(0,0,0,0.05); 
    }
    div[data-testid="stExpander"] { border: none; box-shadow: none; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. FUNCIONES TÉCNICAS ---
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

def to_excel_consolidado(df_ops, df_tr, df_tr_acum, df_seat, df_prm_d, df_fact_d):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dict_dfs = {
            'Ops_Energia': df_ops, 'Kms_Diarios': df_tr, 
            'Lectura_Odometros': df_tr_acum, 'SEAT_Interno': df_seat,
            'PRMTE_Consolidado': df_prm_d, 'Factura_Consolidada': df_fact_d
        }
        for name, df in dict_dfs.items():
            if not df.empty: df.to_excel(writer, index=False, sheet_name=name)
    return output.getvalue()

# --- 3. SIDEBAR: FILTRO CALENDARIO ---
with st.sidebar:
    st.image("https://www.efe.cl/wp-content/themes/efe/img/logo-efe.svg", width=150)
    st.header("📅 Rango de Análisis")
    
    # Selector de calendario tipo rango
    today = date.today()
    start_of_month = today.replace(day=1)
    date_range = st.date_input(
        "Selecciona el período",
        value=(start_of_month, today),
        help="Elige la fecha de inicio y fin en el calendario"
    )
    
    if len(date_range) == 2:
        start_date, end_date = date_range
    else:
        start_date = end_date = date_range[0]

    st.divider()
    st.header("📂 Carga de Datos")
    f_umr = st.file_uploader("1. UMR / Odómetros", type=["xlsx"], accept_multiple_files=True)
    f_seat_files = st.file_uploader("2. Energía SEAT", type=["xlsx"], accept_multiple_files=True)
    f_bill_files = st.file_uploader("3. Facturación y PRMTE", type=["xlsx"], accept_multiple_files=True)

# --- 4. MOTOR DE PROCESAMIENTO ---
all_ops, all_tr, all_tr_acum, all_seat, all_prmte_15, all_fact_h = [], [], [], [], [], []
todos = (f_umr or []) + (f_seat_files or []) + (f_bill_files or [])

for f in todos:
    try:
        xl = pd.ExcelFile(f)
        for sn in xl.sheet_names:
            sn_up = sn.upper()
            
            # UMR / OPS
            if any(k in sn_up for k in ['UMR', 'RESUMEN']):
                df_raw = pd.read_excel(f, sheet_name=sn, header=None)
                h_r = next((i for i in range(min(100, len(df_raw))) if any(k in str(df_raw.iloc[i]).upper() for k in ['ODO', 'FECHA'])), None)
                if h_r is not None:
                    df_p = pd.read_excel(f, sheet_name=sn, header=h_r)
                    df_p.columns = [re.sub(r'[^A-Z]', '', str(c).upper().replace('Ó','O')) for c in df_p.columns]
                    idx_f, idx_o, idx_t = next((c for c in df_p.columns if 'FECHA' in c), None), next((c for c in df_p.columns if 'ODO' in c and 'ACUM' not in c), None), next((c for c in df_p.columns if 'TREN' in c and 'KM' in c), None)
                    if idx_f:
                        df_p['_dt'] = pd.to_datetime(df_p[idx_f], errors='coerce')
                        mask = (df_p['_dt'].dt.date >= start_date) & (df_p['_dt'].dt.date <= end_date)
                        for _, r in df_p[mask].dropna(subset=['_dt']).iterrows():
                            o, tk = parse_latam_number(r[idx_o]), parse_latam_number(r[idx_t])
                            if o > 0: all_ops.append({"Fecha": r['_dt'].normalize(), "Tipo Día": get_tipo_dia(r['_dt']), "Odómetro [km]": o, "Tren-Km [km]": tk, "UMR [%]": (tk/o*100)})

            # TRENES (TABLAS DIARIA Y ACUMULADA)
            if 'ODO' in sn_up and 'KIL' in sn_up:
                df_tr_raw = pd.read_excel(f, sheet_name=sn, header=None)
                h_found = []
                for i in range(len(df_tr_raw)-2):
                    for j in range(1, len(df_tr_raw.columns)):
                        dt_v = pd.to_datetime(df_tr_raw.iloc[i, j], errors='coerce')
                        if pd.notna(dt_v) and start_date <= dt_v.date() <= end_date:
                            if i not in [h[0] for h in h_found]: h_found.append((i, dt_v))
                
                for idx, (r_idx, s_dt) in enumerate(h_found):
                    ctx = str(df_tr_raw.iloc[r_idx:r_idx+3, 0:3]).upper()
                    is_acum = any(k in ctx for k in ['ACUM', 'LECTURA', 'ODO'])
                    c_map = {j: pd.to_datetime(df_tr_raw.iloc[r_idx, j]).normalize() for j in range(1, len(df_tr_raw.columns)) if pd.notna(pd.to_datetime(df_tr_raw.iloc[r_idx, j], errors='coerce'))}
                    for k in range(r_idx+3, min(r_idx+40, len(df_tr_raw))):
                        n_tr = str(df_tr_raw.iloc[k, 0]).strip().upper()
                        if re.match(r'^(M|XM)', n_tr):
                            for c_idx, c_fch in c_map.items():
                                if start_date <= c_fch.date() <= end_date:
                                    val = parse_latam_number(df_tr_raw.iloc[k, c_idx])
                                    dp = {"Tren": n_tr, "Fecha": c_fch, "Día": c_fch.day, "Valor": val}
                                    if is_acum or idx > 0: all_tr_acum.append(dp)
                                    else: all_tr.append(dp)

            # SEAT
            if 'SEAT' in sn_up and 'SER' in sn_up:
                df_s = pd.read_excel(f, sheet_name=sn, header=None)
                for i in range(len(df_s)):
                    fs = pd.to_datetime(df_s.iloc[i, 1], errors='coerce')
                    if pd.notna(fs) and start_date <= fs.date() <= end_date:
                        tot, tra, k12 = parse_latam_number(df_s.iloc[i, 3]), parse_latam_number(df_s.iloc[i, 5]), parse_latam_number(df_s.iloc[i, 7])
                        all_seat.append({"Fecha": fs.normalize(), "Total [kWh]": tot, "Tracción [kWh]": tra, "12 KV [kWh]": k12, "% Tracción": (tra/tot*100 if tot>0 else 0), "% 12 KV": (k12/tot*100 if tot>0 else 0)})

            # PRMTE / FACTURA (Normalizados por rango de calendario)
            if any(k in sn_up for k in ['PRMTE', 'MEDIDAS']):
                df_prm = pd.read_excel(f, sheet_name=sn, header=None)
                h_i = next((i for i in range(len(df_prm)) if 'AÑO' in str(df_prm.iloc[i]).upper()), None)
                if h_i is not None:
                    df_p_d = pd.read_excel(f, sheet_name=sn, header=h_i)
                    df_p_d['TS'] = pd.to_datetime(df_p_d[['AÑO', 'MES', 'DIA', 'HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'})) + pd.to_timedelta(df_p_d['INICIO INTERVALO'].astype(int), unit='m')
                    for _, r in df_p_d[(df_p_d['TS'].dt.date >= start_date) & (df_p_d['TS'].dt.date <= end_date)].iterrows():
                        all_prmte_15.append({"F_H": r['TS'].strftime('%d/%m/%Y %H:%M'), "Fecha": r['TS'].normalize(), "kWh": parse_latam_number(r['Retiro_Energia_Activa (kWhD)'])})

            if any(k in sn_up for k in ['FACTURA', 'CONSUMO']):
                df_f = pd.read_excel(f, sheet_name=sn); df_f.columns = ['FH', 'V']
                df_f['TS'] = pd.to_datetime(df_f['FH'], errors='coerce')
                for _, r in df_f[(df_f['TS'].dt.date >= start_date) & (df_f['TS'].dt.date <= end_date)].dropna(subset=['TS']).iterrows():
                    all_fact_h.append({"F_H": r['TS'].strftime('%d/%m/%Y %H:%M'), "Fecha": r['TS'].normalize(), "kWh": abs(parse_latam_number(r['V']))})
    except: continue

# --- 5. LÓGICA DE JERARQUÍA Y RENDERIZADO ---
df_ops, df_tr, df_tr_acum, df_seat, df_e_master = [pd.DataFrame()] * 5

if any([all_ops, all_tr, all_tr_acum, all_seat, all_prmte_15, all_fact_h]):
    # Consolidación
    if all_ops: df_ops = pd.DataFrame(all_ops).drop_duplicates(subset=['Fecha']).sort_values("Fecha")
    if all_tr: df_tr = pd.DataFrame(all_tr).sort_values(["Fecha", "Tren"])
    if all_tr_acum: df_tr_acum = pd.DataFrame(all_tr_acum).sort_values(["Fecha", "Tren"])
    if all_seat: df_seat = pd.DataFrame(all_seat).drop_duplicates(subset=['Fecha']).sort_values("Fecha")
    
    # Jerarquía Energética
    if not df_seat.empty:
        df_e_master = df_seat[["Fecha", "Total [kWh]", "Tracción [kWh]", "12 KV [kWh]"]].copy().rename(columns={"Total [kWh]":"E_Total", "Tracción [kWh]":"E_Tr", "12 KV [kWh]":"E_12"})
        df_e_master["Fuente"] = "SEAT"

    if all_prmte_15:
        df_p_d = pd.DataFrame(all_prmte_15).groupby("Fecha")["kWh"].sum().reset_index()
        if not df_seat.empty:
            df_p_d = pd.merge(df_p_d, df_seat[["Fecha", "% Tracción", "% 12 KV"]], on="Fecha", how="left").fillna(0)
            df_p_d["E_Tr"], df_p_d["E_12"] = df_p_d["kWh"]*(df_p_d["% Tracción"]/100), df_p_d["kWh"]*(df_p_d["% 12 KV"]/100)
            df_p_p = df_p_d.rename(columns={"kWh":"E_Total"})[["Fecha","E_Total","E_Tr","E_12"]]; df_p_p["Fuente"] = "PRMTE"
            df_e_master = pd.concat([df_e_master, df_p_p]).drop_duplicates(subset=["Fecha"], keep="last")

    if all_fact_h:
        df_f_d = pd.DataFrame(all_fact_h).groupby("Fecha")["kWh"].sum().reset_index()
        if not df_seat.empty:
            df_f_d = pd.merge(df_f_d, df_seat[["Fecha", "% Tracción", "% 12 KV"]], on="Fecha", how="left").fillna(0)
            df_f_d["E_Tr"], df_f_d["E_12"] = df_f_d["kWh"]*(df_f_d["% Tracción"]/100), df_f_d["kWh"]*(df_f_d["% 12 KV"]/100)
            df_f_f = df_f_d.rename(columns={"kWh":"E_Total"})[["Fecha","E_Total","E_Tr","E_12"]]; df_f_f["Fuente"] = "Factura"
            df_e_master = pd.concat([df_e_master, df_f_f]).drop_duplicates(subset=["Fecha"], keep="last")

    if not df_ops.empty and not df_e_master.empty: df_ops = pd.merge(df_ops, df_e_master, on="Fecha", how="left")

    # --- PESTAÑAS ---
    tabs = st.tabs(["📊 Resumen", "📑 Datos Operacionales", "📑 Odómetros Tren", "⚡ Energía SEAT", "📈 PRMTE", "💰 Facturación"])
    
    with tabs[0]: # Resumen Estético
        c1, c2, c3 = st.columns(3)
        to, tk = df_ops["Odómetro [km]"].sum() if not df_ops.empty else 0, df_ops["Tren-Km [km]"].sum() if not df_ops.empty else 0
        c1.metric("Odómetro Total", f"{to:,.1f} km"); c2.metric("Tren-Km Total", f"{tk:,.1f} km"); c3.metric("UMR Global", f"{(tk/to*100 if to>0 else 0):.2f} %")
        st.divider()
        if not df_ops.empty:
            df_ops['Tipo Día'] = pd.Categorical(df_ops['Tipo Día'], categories=['L', 'S', 'D/F'], ordered=True)
            res = df_ops.groupby("Tipo Día").agg({"Odómetro [km]":"sum", "Tren-Km [km]":"sum", "UMR [%]":"mean", "E_Total":"sum", "E_Tr":"sum", "E_12":"sum"}).reset_index()
            st.table(res.rename(columns={"E_Total":"Energía [kWh]"}).style.format("{:,.0f}", subset=["Odómetro [km]","Tren-Km [km]","Energía [kWh]","E_Tr","E_12"]))

    with tabs[1]: # Datos Ops
        if not df_ops.empty: st.dataframe(df_ops.style.format({"Odómetro [km]":"{:,.1f}", "E_Total":"{:,.0f}"}), use_container_width=True)

    with tabs[2]: # Odómetros (Doble Tabla)
        if not df_tr.empty:
            st.write("### 🚗 Kilometraje Diario [km]")
            st.dataframe(df_tr.pivot_table(index="Tren", columns=df_tr["Fecha"].dt.day, values="Valor", aggfunc='sum').fillna(0).style.format("{:,.1f}"), use_container_width=True)
        if not df_tr_acum.empty:
            st.divider(); st.write("### 📈 Lectura de Odómetro / Acumulado [km]")
            st.dataframe(df_tr_acum.pivot_table(index="Tren", columns=df_tr_acum["Fecha"].dt.day, values="Valor", aggfunc='max').fillna(0).style.format("{:,.0f}"), use_container_width=True)

    with tabs[3]: # SEAT
        if not df_seat.empty: st.dataframe(df_seat.style.format({"Total [kWh]":"{:,.0f}", "% Tracción":"{:.2f}%"}), use_container_width=True)

    with tabs[4]: # PRMTE
        if not all_prmte_15: st.info("Sube datos PRMTE")
        else: st.dataframe(pd.DataFrame(all_prmte_15).style.format({"kWh":"{:,.2f}"}), use_container_width=True)

    with tabs[5]: # Factura
        if not all_fact_h: st.info("Sube datos Facturación")
        else: st.dataframe(pd.DataFrame(all_fact_h).style.format({"kWh":"{:,.2f}"}), use_container_width=True)

    st.sidebar.download_button("📥 Reporte Final", to_excel_consolidado(df_ops, df_tr, df_tr_acum, df_seat, pd.DataFrame(), pd.DataFrame()), "Reporte_EFE_SGE.xlsx")
else:
    st.info("👋 Sube los archivos y selecciona el rango en el calendario.")
