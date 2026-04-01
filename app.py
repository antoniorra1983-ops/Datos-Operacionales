import streamlit as st
import pandas as pd
import numpy as np
import re
import holidays
from io import BytesIO
from datetime import datetime, date

# --- 1. CONFIGURACIÓN Y UI PREMIUM ---
st.set_page_config(page_title="EFE SGE - Control Total", layout="wide", page_icon="🚆")
chile_holidays = holidays.Chile()

st.markdown("""
    <style>
    .stApp { background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%); background-attachment: fixed; }
    [data-testid="stSidebar"] { background-color: #005195 !important; color: white !important; }
    .stTable, .stDataFrame, div[data-testid="stMetric"] {
        background-color: rgba(255, 255, 255, 0.8) !important;
        backdrop-filter: blur(10px); border-radius: 15px !important;
        padding: 15px; box-shadow: 0 8px 32px 0 rgba(31, 38, 135, 0.1);
    }
    [data-testid="stMetricValue"] { color: #005195 !important; font-weight: bold; }
    h1, h2, h3 { color: #003366; font-family: 'Segoe UI', sans-serif; }
    .stMultiSelect label, .stDateInput label { color: white !important; }
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
        dict_dfs = {'Ops_Energia': df_ops, 'Kms_Diarios': df_tr, 'Lectura_Odometros': df_tr_acum, 
                    'SEAT': df_seat, 'PRMTE': df_prm_d, 'Factura': df_fact_d}
        for name, df in dict_dfs.items():
            if not df.empty: df.to_excel(writer, index=False, sheet_name=name)
    return output.getvalue()

# --- 3. SIDEBAR: CARGA Y FILTROS ---
with st.sidebar:
    st.image("https://www.efe.cl/wp-content/themes/efe/img/logo-efe.svg", width=120)
    st.header("📂 Carga de Datos")
    f_umr = st.file_uploader("1. UMR / Odómetros", type=["xlsx"], accept_multiple_files=True)
    f_seat_files = st.file_uploader("2. Energía SEAT", type=["xlsx"], accept_multiple_files=True)
    f_bill_files = st.file_uploader("3. Facturación / PRMTE", type=["xlsx"], accept_multiple_files=True)
    
    st.divider()
    st.header("🔍 Filtros de Búsqueda")
    # Calendario base
    today = date.today()
    date_range = st.date_input("Rango Calendario", value=(today.replace(day=1, month=1), today))
    start_date, end_date = (date_range[0], date_range[1]) if len(date_range) == 2 else (date_range[0], date_range[0])

    # Filtros granulares (se activan si hay datos)
    f_year = st.multiselect("Año", [2024, 2025, 2026], default=[2025, 2026])
    f_month = st.multiselect("Mes", list(range(1, 13)), default=list(range(1, 13)))
    f_jornada = st.multiselect("Jornada", ["L", "S", "D/F"], default=["L", "S", "D/F"])
    f_week = st.multiselect("N° Semana", list(range(1, 54)))
    f_day = st.multiselect("Día del Mes", list(range(1, 32)))

# --- 4. MOTOR DE PROCESAMIENTO ---
all_ops, all_tr, all_tr_acum, all_seat, all_prmte_15, all_fact_h = [], [], [], [], [], []
todos = (f_umr or []) + (f_seat_files or []) + (f_bill_files or [])

for f in todos:
    try:
        xl = pd.ExcelFile(f)
        for sn in xl.sheet_names:
            sn_up = sn.upper()
            
            # A. UMR / OPS
            if any(k in sn_up for k in ['UMR', 'RESUMEN']):
                df_raw = pd.read_excel(f, sheet_name=sn, header=None)
                h_r = next((i for i in range(min(100, len(df_raw))) if any(k in str(df_raw.iloc[i]).upper() for k in ['ODO', 'FECHA'])), None)
                if h_r is not None:
                    df_p = pd.read_excel(f, sheet_name=sn, header=h_r)
                    df_p.columns = [re.sub(r'[^A-Z]', '', str(c).upper().replace('Ó','O')) for c in df_p.columns]
                    idx_f, idx_o, idx_t = next((c for c in df_p.columns if 'FECHA' in c), None), next((c for c in df_p.columns if 'ODO' in c and 'ACUM' not in c), None), next((c for c in df_p.columns if 'TREN' in c and 'KM' in c), None)
                    if idx_f:
                        df_p['_dt'] = pd.to_datetime(df_p[idx_f], errors='coerce')
                        for _, r in df_p.dropna(subset=['_dt']).iterrows():
                            all_ops.append({"Fecha": r['_dt'].normalize(), "Tipo Día": get_tipo_dia(r['_dt']), "Odómetro [km]": parse_latam_number(r[idx_o]), "Tren-Km [km]": parse_latam_number(r[idx_t])})

            # B. TRENES (DOBLE TABLA)
            if 'ODO' in sn_up and 'KIL' in sn_up:
                df_tr_raw = pd.read_excel(f, sheet_name=sn, header=None)
                h_found = []
                for i in range(len(df_tr_raw)-2):
                    for j in range(1, len(df_tr_raw.columns)):
                        dt_v = pd.to_datetime(df_tr_raw.iloc[i, j], errors='coerce')
                        if pd.notna(dt_v):
                            if i not in [h[0] for h in h_found]: h_found.append((i, dt_v))
                for idx, (r_idx, s_dt) in enumerate(h_found):
                    ctx = str(df_tr_raw.iloc[r_idx:r_idx+3, 0:3]).upper()
                    is_acum = any(k in ctx for k in ['ACUM', 'LECTURA', 'ODO'])
                    c_map = {j: pd.to_datetime(df_tr_raw.iloc[r_idx, j]).normalize() for j in range(1, len(df_tr_raw.columns)) if pd.notna(pd.to_datetime(df_tr_raw.iloc[r_idx, j], errors='coerce'))}
                    for k in range(r_idx+3, min(r_idx+40, len(df_tr_raw))):
                        n_tr = str(df_tr_raw.iloc[k, 0]).strip().upper()
                        if re.match(r'^(M|XM)', n_tr):
                            for c_idx, c_fch in c_map.items():
                                val = parse_latam_number(df_tr_raw.iloc[k, c_idx])
                                dp = {"Tren": n_tr, "Fecha": c_fch, "Valor": val}
                                if is_acum or idx > 0: all_tr_acum.append(dp)
                                else: all_tr.append(dp)

            # C. SEAT / PRMTE / FACTURA (Procesamiento base)
            if 'SEAT' in sn_up and 'SER' in sn_up:
                df_s = pd.read_excel(f, sheet_name=sn, header=None)
                for i in range(len(df_s)):
                    fs = pd.to_datetime(df_s.iloc[i, 1], errors='coerce')
                    if pd.notna(fs):
                        all_seat.append({"Fecha": fs.normalize(), "Total [kWh]": parse_latam_number(df_s.iloc[i, 3]), "Tracción [kWh]": parse_latam_number(df_s.iloc[i, 5]), "12 KV [kWh]": parse_latam_number(df_s.iloc[i, 7]), "% Tracción": (parse_latam_number(df_s.iloc[i, 5])/parse_latam_number(df_s.iloc[i, 3])*100 if parse_latam_number(df_s.iloc[i, 3])>0 else 0), "% 12 KV": (parse_latam_number(df_s.iloc[i, 7])/parse_latam_number(df_s.iloc[i, 3])*100 if parse_latam_number(df_s.iloc[i, 3])>0 else 0)})

            if any(k in sn_up for k in ['PRMTE', 'MEDIDAS']):
                df_prm = pd.read_excel(f, sheet_name=sn, header=None)
                h_i = next((i for i in range(len(df_prm)) if 'AÑO' in str(df_prm.iloc[i]).upper()), None)
                if h_i is not None:
                    df_p_d = pd.read_excel(f, sheet_name=sn, header=h_i)
                    df_p_d['TS'] = pd.to_datetime(df_p_d[['AÑO', 'MES', 'DIA', 'HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'})) + pd.to_timedelta(df_p_d['INICIO INTERVALO'].astype(int), unit='m')
                    for _, r in df_p_d.iterrows():
                        all_prmte_15.append({"F_H": r['TS'], "Fecha": r['TS'].normalize(), "kWh": parse_latam_number(r['Retiro_Energia_Activa (kWhD)'])})

            if any(k in sn_up for k in ['FACTURA', 'CONSUMO']):
                df_f = pd.read_excel(f, sheet_name=sn); df_f.columns = ['FH', 'V']
                df_f['TS'] = pd.to_datetime(df_f['FH'], errors='coerce')
                for _, r in df_f.dropna(subset=['TS']).iterrows():
                    all_fact_h.append({"F_H": r['TS'], "Fecha": r['TS'].normalize(), "kWh": abs(parse_latam_number(r['V']))})
    except: continue

# --- 5. ENSAMBLAJE Y FILTRADO DINÁMICO ---
def apply_filters(df):
    if df.empty: return df
    df['Año'] = df['Fecha'].dt.year
    df['Mes'] = df['Fecha'].dt.month
    df['Día'] = df['Fecha'].dt.day
    df['Semana'] = df['Fecha'].dt.isocalendar().week
    if 'Tipo Día' not in df.columns: df['Tipo Día'] = df['Fecha'].apply(get_tipo_dia)
    
    mask = (df['Fecha'].dt.date >= start_date) & (df['Fecha'].dt.date <= end_date) & \
           (df['Año'].isin(f_year)) & (df['Mes'].isin(f_month)) & (df['Tipo Día'].isin(f_jornada))
    if f_week: mask &= (df['Semana'].isin(f_week))
    if f_day: mask &= (df['Día'].isin(f_day))
    return df[mask]

# Convertir y aplicar filtros a TODO
df_ops = apply_filters(pd.DataFrame(all_ops).drop_duplicates(subset=['Fecha']))
df_tr = apply_filters(pd.DataFrame(all_tr))
df_tr_a = apply_filters(pd.DataFrame(all_tr_acum))
df_seat = apply_filters(pd.DataFrame(all_seat).drop_duplicates(subset=['Fecha']))
df_prm_15 = apply_filters(pd.DataFrame(all_prmte_15))
df_fact_h = apply_filters(pd.DataFrame(all_fact_h))

# --- 6. JERARQUÍA ENERGÉTICA ---
df_e_master = pd.DataFrame()
if not df_seat.empty:
    df_e_master = df_seat[["Fecha", "Total [kWh]", "Tracción [kWh]", "12 KV [kWh]"]].copy().rename(columns={"Total [kWh]":"E_Total", "Tracción [kWh]":"E_Tr", "12 KV [kWh]":"E_12"})
    df_e_master["Fuente"] = "SEAT"

# Proporcionalidad PRMTE y Factura
for df_src, label, col_in in zip([df_prm_15, df_fact_h], ["PRMTE", "Factura"], ["kWh", "kWh"]):
    if not df_src.empty:
        df_d = df_src.groupby("Fecha")[col_in].sum().reset_index()
        if not df_seat.empty:
            df_d = pd.merge(df_d, df_seat[["Fecha", "% Tracción", "% 12 KV"]], on="Fecha", how="left").fillna(0)
            df_d["E_Tr"], df_d["E_12"] = df_d[col_in]*(df_d["% Tracción"]/100), df_d[col_in]*(df_d["% 12 KV"]/100)
            df_p = df_d.rename(columns={col_in:"E_Total"})[["Fecha","E_Total","E_Tr","E_12"]]; df_p["Fuente"] = label
            df_e_master = pd.concat([df_e_master, df_p]).drop_duplicates(subset=["Fecha"], keep="last")

if not df_ops.empty and not df_e_master.empty:
    df_ops = pd.merge(df_ops, df_e_master, on="Fecha", how="left")

# --- 7. RENDERIZADO DE TABS ---
if any([not df_ops.empty, not df_tr.empty, not df_seat.empty]):
    tabs = st.tabs(["📊 Resumen", "📑 Datos Operacionales", "📑 Odómetros Tren", "⚡ Energía SEAT", "📈 PRMTE", "💰 Facturación"])
    
    with tabs[0]:
        c1, c2, c3 = st.columns(3)
        to, tk = df_ops["Odómetro [km]"].sum() if not df_ops.empty else 0, df_ops["Tren-Km [km]"].sum() if not df_ops.empty else 0
        c1.metric("Odómetro Total", f"{to:,.1f} km"); c2.metric("Tren-Km Total", f"{tk:,.1f} km"); c3.metric("UMR Global", f"{(tk/to*100 if to>0 else 0):.2f} %")
        st.divider()
        if not df_ops.empty:
            res = df_ops.groupby("Tipo Día", observed=True).agg({"Odómetro [km]":"sum", "Tren-Km [km]":"sum", "UMR [%]":"mean", "E_Total":"sum", "E_Tr":"sum", "E_12":"sum"}).reset_index()
            st.table(res.style.format("{:,.0f}", subset=["Odómetro [km]","Tren-Km [km]","E_Total","E_Tr","E_12"]))

    with tabs[1]:
        st.dataframe(df_ops.style.format({"Odómetro [km]":"{:,.1f}", "E_Total":"{:,.0f}"}), use_container_width=True)

    with tabs[2]:
        if not df_tr.empty:
            st.write("### 🚗 Kilometraje Diario [km]")
            st.dataframe(df_tr.pivot_table(index="Tren", columns=df_tr["Fecha"].dt.day, values="Valor", aggfunc='sum').fillna(0).style.format("{:,.1f}"), use_container_width=True)
        if not df_tr_a.empty:
            st.divider(); st.write("### 📈 Lectura de Odómetro / Acumulado [km]")
            st.dataframe(df_tr_a.pivot_table(index="Tren", columns=df_tr_a["Fecha"].dt.day, values="Valor", aggfunc='max').fillna(0).style.format("{:,.0f}"), use_container_width=True)

    with tabs[3]: st.dataframe(df_seat.style.format({"Total [kWh]":"{:,.0f}", "% Tracción":"{:.2f}%"}), use_container_width=True)
    with tabs[4]: st.dataframe(df_prm_15.style.format({"kWh":"{:,.2f}"}), use_container_width=True)
    with tabs[5]: st.dataframe(df_fact_h.style.format({"kWh":"{:,.2f}"}), use_container_width=True)

    st.sidebar.download_button("📥 Reporte Final", to_excel_consolidado(df_ops, df_tr, df_tr_a, df_seat, df_prm_15, df_fact_h), "Reporte_SGE_EFE.xlsx")
else:
    st.info("👋 Sube archivos y ajusta los filtros para visualizar los datos.")
