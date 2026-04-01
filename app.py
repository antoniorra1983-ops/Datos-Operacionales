import streamlit as st
import pandas as pd
import numpy as np
import re
import holidays
from io import BytesIO
from datetime import datetime, date

# --- 1. CONFIGURACIÓN Y UI PREMIUM (GLASSMORPHISM) ---
st.set_page_config(page_title="EFE SGE - Dashboard Profesional", layout="wide", page_icon="🚆")
chile_holidays = holidays.Chile()

st.markdown("""
    <style>
    .stApp {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        background-attachment: fixed;
    }
    [data-testid="stSidebar"] {
        background-color: #005195 !important;
        color: white !important;
    }
    [data-testid="stSidebar"] .stMarkdown p, [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2 {
        color: white !important;
    }
    .stTable, .stDataFrame, div[data-testid="stMetric"] {
        background-color: rgba(255, 255, 255, 0.8) !important;
        backdrop-filter: blur(10px);
        border-radius: 15px !important;
        padding: 15px;
        box-shadow: 0 8px 32px 0 rgba(31, 38, 135, 0.1);
        border: 1px solid rgba(255, 255, 255, 0.18);
    }
    [data-testid="stMetricValue"] { color: #005195 !important; font-weight: bold; }
    h1, h2, h3 { color: #003366; }
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

def to_excel_consolidado(df_ops, df_tr, df_tr_a, df_seat, df_prm, df_fact):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dict_dfs = {'Operaciones': df_ops, 'Trenes_Diario': df_tr, 'Trenes_Acumulado': df_tr_a, 'SEAT': df_seat, 'PRMTE': df_prm, 'Factura': df_fact}
        for name, df in dict_dfs.items():
            if not df.empty: df.to_excel(writer, index=False, sheet_name=name)
    return output.getvalue()

# --- 3. SIDEBAR CON CALENDARIO ---
with st.sidebar:
    st.title("EFE Valparaíso")
    st.divider()
    st.header("📅 Rango de Análisis")
    date_range = st.date_input("Seleccionar Rango", value=(date(2026, 3, 1), date.today()))
    
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_d, end_d = date_range
    else:
        start_d = end_d = (date_range[0] if isinstance(date_range, tuple) else date_range)

    st.header("📂 Carga de Archivos")
    f_umr = st.file_uploader("UMR / Odómetros", type=["xlsx"], accept_multiple_files=True)
    f_seat_f = st.file_uploader("Energía SEAT", type=["xlsx"], accept_multiple_files=True)
    f_prmte_f = st.file_uploader("Facturación / PRMTE", type=["xlsx"], accept_multiple_files=True)

# --- 4. MOTOR DE PROCESAMIENTO ---
# Inicialización de listas
a_ops, a_tr, a_tr_a, a_seat, a_prm_15, a_fact_h = [], [], [], [], [], []
todos = (f_umr or []) + (f_seat_f or []) + (f_prmte_f or [])

for f in todos:
    try:
        xl = pd.ExcelFile(f)
        for sn in xl.sheet_names:
            sn_up = sn.upper()
            
            # OPS / UMR
            if any(k in sn_up for k in ['UMR', 'RESUMEN']):
                df_raw = pd.read_excel(f, sheet_name=sn, header=None)
                h_i = next((i for i in range(min(100, len(df_raw))) if any(k in str(df_raw.iloc[i]).upper() for k in ['ODO', 'FECHA'])), None)
                if h_i is not None:
                    df_p = pd.read_excel(f, sheet_name=sn, header=h_i)
                    df_p.columns = [re.sub(r'[^A-Z]', '', str(c).upper()) for c in df_p.columns]
                    cf, co, ct = next((c for c in df_p.columns if 'FECHA' in c), None), next((c for c in df_p.columns if 'ODO' in c and 'ACUM' not in c), None), next((c for c in df_p.columns if 'TREN' in c and 'KM' in c), None)
                    if cf:
                        df_p['_dt'] = pd.to_datetime(df_p[cf], errors='coerce')
                        mask = (df_p['_dt'].dt.date >= start_d) & (df_p['_dt'].dt.date <= end_d)
                        for _, r in df_p[mask].dropna(subset=['_dt']).iterrows():
                            a_ops.append({"Fecha": r['_dt'].normalize(), "Tipo Día": get_tipo_dia(r['_dt']), "Odómetro [km]": parse_latam_number(r[co]), "Tren-Km [km]": parse_latam_number(r[ct]), "UMR [%]": (parse_latam_number(r[ct])/parse_latam_number(r[co])*100 if parse_latam_number(r[co])>0 else 0)})

            # TRENES (DOBLE TABLA)
            if 'ODO' in sn_up and 'KIL' in sn_up:
                df_tr_raw = pd.read_excel(f, sheet_name=sn, header=None)
                blocks = []
                for i in range(len(df_tr_raw)-2):
                    for j in range(1, len(df_tr_raw.columns)):
                        dv = pd.to_datetime(df_tr_raw.iloc[i, j], errors='coerce')
                        if pd.notna(dv) and start_d <= dv.date() <= end_d:
                            if i not in [b[0] for b in blocks]: blocks.append((i, dv))
                
                for idx, (ri, sd) in enumerate(blocks):
                    is_acum = any(k in str(df_tr_raw.iloc[ri:ri+3, 0:3]).upper() for k in ['ACUM', 'ODO', 'LECTURA'])
                    c_map = {j: pd.to_datetime(df_tr_raw.iloc[ri, j]).normalize() for j in range(1, len(df_tr_raw.columns)) if pd.notna(pd.to_datetime(df_tr_raw.iloc[ri, j], errors='coerce'))}
                    for k in range(ri+3, min(ri+45, len(df_tr_raw))):
                        tr_n = str(df_tr_raw.iloc[k, 0]).strip().upper()
                        if re.match(r'^(M|XM)', tr_n):
                            for ci, cf in c_map.items():
                                if start_d <= cf.date() <= end_d:
                                    val = parse_latam_number(df_tr_raw.iloc[k, ci])
                                    dp = {"Tren": tr_n, "Fecha": cf, "Día": cf.day, "Valor": val}
                                    if is_acum or idx > 0: a_tr_a.append(dp)
                                    else: a_tr.append(dp)

            # SEAT
            if 'SEAT' in sn_up and 'SER' in sn_up:
                df_s = pd.read_excel(f, sheet_name=sn, header=None)
                for i in range(len(df_s)):
                    fs = pd.to_datetime(df_s.iloc[i, 1], errors='coerce')
                    if pd.notna(fs) and start_d <= fs.date() <= end_d:
                        a_seat.append({"Fecha": fs.normalize(), "Total [kWh]": parse_latam_number(df_s.iloc[i, 3]), "Tracción [kWh]": parse_latam_number(df_s.iloc[i, 5]), "12 KV [kWh]": parse_latam_number(df_s.iloc[i, 7]), "% Tracción": (parse_latam_number(df_s.iloc[i, 5])/parse_latam_number(df_s.iloc[i, 3])*100 if parse_latam_number(df_s.iloc[i, 3])>0 else 0), "% 12 KV": (parse_latam_number(df_s.iloc[i, 7])/parse_latam_number(df_s.iloc[i, 3])*100 if parse_latam_number(df_s.iloc[i, 3])>0 else 0)})

            # PRMTE / FACTURA
            if any(k in sn_up for k in ['PRMTE', 'MEDIDAS']):
                df_prm = pd.read_excel(f, sheet_name=sn, header=None)
                h_i = next((i for i in range(len(df_prm)) if 'AÑO' in str(df_prm.iloc[i]).upper()), None)
                if h_i is not None:
                    df_pd = pd.read_excel(f, sheet_name=sn, header=h_i)
                    df_pd['TS'] = pd.to_datetime(df_pd[['AÑO', 'MES', 'DIA', 'HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'})) + pd.to_timedelta(df_pd['INICIO INTERVALO'].astype(int), unit='m')
                    for _, r in df_pd[(df_pd['TS'].dt.date >= start_d) & (df_pd['TS'].dt.date <= end_d)].iterrows():
                        a_prm_15.append({"F_H": r['TS'].strftime('%d/%m/%Y %H:%M'), "Fecha": r['TS'].normalize(), "kWh": parse_latam_number(r['Retiro_Energia_Activa (kWhD)'])})

            if any(k in sn_up for k in ['FACTURA', 'CONSUMO']):
                df_f = pd.read_excel(f, sheet_name=sn); df_f.columns = ['FH', 'V']
                df_f['TS'] = pd.to_datetime(df_f['FH'], errors='coerce')
                for _, r in df_f[(df_f['TS'].dt.date >= start_d) & (df_f['TS'].dt.date <= end_d)].dropna(subset=['TS']).iterrows():
                    a_fact_h.append({"F_H": r['TS'].strftime('%d/%m/%Y %H:%M'), "Fecha": r['TS'].normalize(), "kWh": abs(parse_latam_number(r['V']))})
    except: continue

# --- 5. LÓGICA DE JERARQUÍA Y RENDERIZADO ---
df_ops, df_tr, df_tr_a, df_seat, df_prm_d, df_fact_d = [pd.DataFrame()] * 6
df_energy_master = pd.DataFrame()

if any([a_ops, a_tr, a_tr_a, a_seat, a_prm_15, a_fact_h]):
    # Consolidación de DataFrames
    if a_ops: df_ops = pd.DataFrame(a_ops).drop_duplicates(subset=['Fecha']).sort_values("Fecha")
    if a_tr: df_tr = pd.DataFrame(a_tr).sort_values(["Fecha", "Tren"])
    if a_tr_a: df_tr_a = pd.DataFrame(a_tr_a).sort_values(["Fecha", "Tren"])
    if a_seat: df_seat = pd.DataFrame(a_seat).drop_duplicates(subset=['Fecha']).sort_values("Fecha")
    
    # Jerarquía: SEAT (Base) -> PRMTE -> Factura (Rey)
    if not df_seat.empty:
        df_energy_master = df_seat[["Fecha", "Total [kWh]", "Tracción [kWh]", "12 KV [kWh]"]].copy().rename(columns={"Total [kWh]":"E_Total", "Tracción [kWh]":"E_Tr", "12 KV [kWh]":"E_12"})
        df_energy_master["Fuente"] = "SEAT"

    if a_prm_15:
        df_prm_d = pd.DataFrame(a_prm_15).groupby("Fecha")["kWh"].sum().reset_index()
        if not df_seat.empty:
            df_prm_d = pd.merge(df_prm_d, df_seat[["Fecha", "% Tracción", "% 12 KV"]], on="Fecha", how="left").fillna(0)
            df_prm_d["E_Tr"], df_prm_d["E_12"] = df_prm_d["kWh"]*(df_prm_d["% Tracción"]/100), df_prm_d["kWh"]*(df_prm_d["% 12 KV"]/100)
            df_prm_p = df_prm_d.rename(columns={"kWh":"E_Total"})[["Fecha","E_Total","E_Tr","E_12"]]; df_prm_p["Fuente"] = "PRMTE"
            df_energy_master = pd.concat([df_energy_master, df_prm_p]).drop_duplicates(subset=["Fecha"], keep="last")

    if a_fact_h:
        df_fact_d = pd.DataFrame(a_fact_h).groupby("Fecha")["kWh"].sum().reset_index()
        if not df_seat.empty:
            df_fact_d = pd.merge(df_fact_d, df_seat[["Fecha", "% Tracción", "% 12 KV"]], on="Fecha", how="left").fillna(0)
            df_fact_d["E_Tr"], df_fact_d["E_12"] = df_fact_d["kWh"]*(df_fact_d["% Tracción"]/100), df_fact_d["kWh"]*(df_fact_d["% 12 KV"]/100)
            df_fact_f = df_fact_d.rename(columns={"kWh":"E_Total"})[["Fecha","E_Total","E_Tr","E_12"]]; df_fact_f["Fuente"] = "Factura"
            df_energy_master = pd.concat([df_energy_master, df_fact_f]).drop_duplicates(subset=["Fecha"], keep="last")

    if not df_ops.empty and not df_energy_master.empty:
        df_ops = pd.merge(df_ops, df_energy_master, on="Fecha", how="left")

    # --- PESTAÑAS ---
    tabs = st.tabs(["📊 Resumen", "📑 Operaciones", "📑 Trenes", "⚡ SEAT", "📈 PRMTE", "💰 Facturación"])
    
    with tabs[0]: # Resumen
        c1, c2, c3 = st.columns(3)
        to, tk = df_ops["Odómetro [km]"].sum() if not df_ops.empty else 0, df_ops["Tren-Km [km]"].sum() if not df_ops.empty else 0
        c1.metric("Odómetro Total", f"{to:,.1f} km"); c2.metric("Tren-Km Total", f"{tk:,.1f} km"); c3.metric("UMR Global", f"{(tk/to*100 if to>0 else 0):.2f} %")
        if not df_ops.empty:
            st.divider()
            df_ops['Tipo Día'] = pd.Categorical(df_ops['Tipo Día'], categories=['L', 'S', 'D/F'], ordered=True)
            res = df_ops.groupby("Tipo Día").agg({"Odómetro [km]":"sum", "Tren-Km [km]":"sum", "UMR [%]":"mean", "E_Total":"sum", "E_Tr":"sum", "E_12":"sum"}).reset_index()
            st.write("#### Consumo Eléctrico vs Operación")
            st.table(res.style.format({"Odómetro [km]":"{:,.0f}", "Tren-Km [km]":"{:,.0f}", "UMR [%]":"{:.2f}%", "E_Total":"{:,.0f}", "E_Tr":"{:,.0f}", "E_12":"{:,.0f}"}))

    with tabs[1]: # Operaciones
        if not df_ops.empty:
            st.dataframe(df_ops.style.format({"Odómetro [km]":"{:,.1f}", "Tren-Km [km]":"{:,.1f}", "UMR [%]":"{:.2f}%", "E_Total":"{:,.0f}", "E_Tr":"{:,.0f}", "E_12":"{:,.0f}"}), use_container_width=True)

    with tabs[2]: # Trenes (Doble Tabla)
        if not df_tr.empty:
            st.write("#### Kilometraje Diario [km]")
            st.dataframe(df_tr.pivot_table(index="Tren", columns=df_tr["Fecha"].dt.day, values="Valor", aggfunc='sum').fillna(0).style.format("{:,.1f}"), use_container_width=True)
        if not df_tr_a.empty:
            st.divider(); st.write("#### Lectura Odómetro Acumulado [km]")
            st.dataframe(df_tr_a.pivot_table(index="Tren", columns=df_tr_a["Fecha"].dt.day, values="Valor", aggfunc='max').fillna(0).style.format("{:,.0f}"), use_container_width=True)

    with tabs[3]: # SEAT
        if not df_seat.empty: st.dataframe(df_seat.style.format({"Total [kWh]":"{:,.0f}", "% Tracción":"{:.2f}%"}), use_container_width=True)

    with tabs[4]: # PRMTE
        if not df_prm_d.empty:
            st.write("#### Resumen Diario PRMTE"); st.dataframe(df_prm_d.style.format({"kWh":"{:,.1f}", "E_Tr":"{:,.1f}", "E_12":"{:,.1f}"}), use_container_width=True)
            st.write("#### Detalle 15 Minutos"); st.dataframe(pd.DataFrame(a_prm_15).style.format({"kWh":"{:,.2f}"}), use_container_width=True)

    with tabs[5]: # Factura
        if not df_fact_d.empty:
            st.write("#### Resumen Diario Factura"); st.dataframe(df_fact_d.style.format({"kWh":"{:,.1f}", "E_Tr":"{:,.1f}", "E_12":"{:,.1f}"}), use_container_width=True)
            st.write("#### Detalle Horario"); st.dataframe(pd.DataFrame(a_fact_h).style.format({"kWh":"{:,.2f}"}), use_container_width=True)

    st.sidebar.download_button("📥 Exportar Informe", to_excel_consolidado(df_ops, df_tr, df_tr_a, df_seat, df_prm_d, df_fact_d), "Reporte_SGE_EFE.xlsx")
else:
    st.info("👋 Sube los archivos para activar el Dashboard profesional.")
