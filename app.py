import streamlit as st
import pandas as pd
import numpy as np
import re
import holidays
from io import BytesIO
from datetime import datetime, date

# --- 1. CONFIGURACIÓN Y UI PREMIUM ---
st.set_page_config(page_title="EFE SGE - Control Energético", layout="wide", page_icon="🚆")
chile_holidays = holidays.Chile()

st.markdown("""
    <style>
    .stApp { background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%); background-attachment: fixed; }
    [data-testid="stSidebar"] { background-color: #005195 !important; color: white !important; }
    .stTable, .stDataFrame, div[data-testid="stMetric"] {
        background-color: rgba(255, 255, 255, 0.8) !important;
        backdrop-filter: blur(10px); border-radius: 12px !important;
        padding: 15px; box-shadow: 0 4px 15px 0 rgba(31, 38, 135, 0.05);
    }
    [data-testid="stMetricValue"] { color: #005195 !important; font-weight: bold; }
    h1, h2, h3, h4 { color: #003366; font-family: 'Segoe UI', sans-serif; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. FUNCIONES TÉCNICAS ---
def parse_latam_number(val):
    if pd.isna(val) or val == '': return 0.0
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

def to_excel_consolidado(df_ops, df_tr, df_tr_a, df_seat, df_prm_d, df_fact_d):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dict_dfs = {'Operaciones': df_ops, 'Trenes_Diario': df_tr, 'Trenes_Acumulado': df_tr_a, 
                    'SEAT': df_seat, 'PRMTE_Diario': df_prm_d, 'Factura_Diaria': df_fact_d}
        for name, df in dict_dfs.items():
            if not df.empty: df.to_excel(writer, index=False, sheet_name=name)
    return output.getvalue()

# --- 3. SIDEBAR: CARGA Y CALENDARIO GLOBAL ---
with st.sidebar:
    st.image("https://www.efe.cl/wp-content/themes/efe/img/logo-efe.svg", width=120)
    st.divider()
    st.header("📅 Calendario Global")
    
    # Selector de Rango tipo Calendario (Apunta a Marzo 2026 por defecto)
    date_range = st.date_input(
        "Seleccionar Período de Análisis", 
        value=(date(2026, 3, 1), date(2026, 3, 31))
    )
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_d, end_d = date_range
    else:
        start_d = end_d = (date_range[0] if isinstance(date_range, tuple) else date_range)

    st.header("📂 Carga de Archivos")
    f_umr = st.file_uploader("1. UMR / Odómetros", type=["xlsx"], accept_multiple_files=True)
    f_seat_f = st.file_uploader("2. Energía SEAT", type=["xlsx"], accept_multiple_files=True)
    f_prmte_f = st.file_uploader("3. Facturación / PRMTE", type=["xlsx"], accept_multiple_files=True)

# --- 4. MOTOR DE PROCESAMIENTO (DATOS RECUPERADOS) ---
a_ops, a_tr, a_tr_a, a_seat, a_prm_15, a_fact_h = [], [], [], [], [], []
todos = (f_umr or []) + (f_seat_f or []) + (f_prmte_f or [])

for f in todos:
    try:
        xl = pd.ExcelFile(f)
        for sn in xl.sheet_names:
            sn_up = sn.upper()
            
            # A. DATOS OPERACIONALES (UMR)
            if any(k in sn_up for k in ['UMR', 'RESUMEN', 'OPERACIONAL']):
                df_raw = pd.read_excel(f, sheet_name=sn, header=None)
                h_i = None
                for i in range(min(100, len(df_raw))):
                    if any(k in str(df_raw.iloc[i]).upper() for k in ['FECHA', 'DIA', 'DATE']):
                        h_i = i; break
                
                if h_i is not None:
                    df_p = pd.read_excel(f, sheet_name=sn, header=h_i)
                    cols = df_p.columns.astype(str).str.upper()
                    idx_f = next((i for i, c in enumerate(cols) if 'FECHA' in c or 'DIA' in c), None)
                    idx_o = next((i for i, c in enumerate(cols) if 'ODO' in c and 'ACUM' not in c), None)
                    idx_t = next((i for i, c in enumerate(cols) if 'TREN' in c and 'KM' in c), None)
                    
                    # Corrección vital: `is not None` evita perder datos si la columna es la 0
                    if idx_f is not None:
                        for _, r in df_p.iterrows():
                            fch = pd.to_datetime(r.iloc[idx_f], errors='coerce')
                            if pd.notna(fch) and start_d <= fch.date() <= end_d:
                                odo = parse_latam_number(r.iloc[idx_o]) if idx_o is not None else 0
                                tkm = parse_latam_number(r.iloc[idx_t]) if idx_t is not None else 0
                                if odo > 0:
                                    a_ops.append({"Fecha": fch.normalize(), "Tipo Día": get_tipo_dia(fch), "N° Semana": fch.isocalendar()[1], "Odómetro [km]": odo, "Tren-Km [km]": tkm, "UMR [%]": (tkm/odo*100)})

            # B. TRENES (DOBLE TABLA)
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

            # C. SEAT
            if 'SEAT' in sn_up and 'SER' in sn_up:
                df_s = pd.read_excel(f, sheet_name=sn, header=None)
                for i in range(len(df_s)):
                    fs = pd.to_datetime(df_s.iloc[i, 1], errors='coerce')
                    if pd.notna(fs) and start_d <= fs.date() <= end_d:
                        a_seat.append({"Fecha": fs.normalize(), "Total [kWh]": parse_latam_number(df_s.iloc[i, 3]), "Tracción [kWh]": parse_latam_number(df_s.iloc[i, 5]), "12 KV [kWh]": parse_latam_number(df_s.iloc[i, 7]), "% Tracción": (parse_latam_number(df_s.iloc[i, 5])/parse_latam_number(df_s.iloc[i, 3])*100 if parse_latam_number(df_s.iloc[i, 3])>0 else 0), "% 12 KV": (parse_latam_number(df_s.iloc[i, 7])/parse_latam_number(df_s.iloc[i, 3])*100 if parse_latam_number(df_s.iloc[i, 3])>0 else 0)})

            # D. PRMTE / FACTURA
            if any(k in sn_up for k in ['PRMTE', 'MEDIDAS']):
                df_prm = pd.read_excel(f, sheet_name=sn, header=None)
                h_i = next((i for i in range(len(df_prm)) if 'AÑO' in str(df_prm.iloc[i]).upper()), None)
                if h_i is not None:
                    df_pd = pd.read_excel(f, sheet_name=sn, header=h_i)
                    df_pd['TS'] = pd.to_datetime(df_pd[['AÑO', 'MES', 'DIA', 'HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'})) + pd.to_timedelta(df_pd['INICIO INTERVALO'].astype(int) if 'INICIO INTERVALO' in df_pd else 0, unit='m')
                    for _, r in df_pd[(df_pd['TS'].dt.date >= start_d) & (df_pd['TS'].dt.date <= end_d)].iterrows():
                        a_prm_15.append({"F_H": r['TS'].strftime('%d/%m/%Y %H:%M'), "Fecha": r['TS'].normalize(), "kWh": parse_latam_number(r.get('Retiro_Energia_Activa (kWhD)', 0))})

            if any(k in sn_up for k in ['FACTURA', 'CONSUMO']):
                df_f = pd.read_excel(f, sheet_name=sn); df_f.columns = ['FH', 'V', 'T'] if len(df_f.columns)>2 else ['FH', 'V']
                df_f['TS'] = pd.to_datetime(df_f['FH'], errors='coerce')
                for _, r in df_f[(df_f['TS'].dt.date >= start_d) & (df_f['TS'].dt.date <= end_d)].dropna(subset=['TS']).iterrows():
                    a_fact_h.append({"F_H": r['TS'].strftime('%d/%m/%Y %H:%M'), "Fecha": r['TS'].normalize(), "kWh": abs(parse_latam_number(r['V']))})
    except: continue

# --- 5. ENSAMBLAJE DE JERARQUÍA ---
df_ops, df_tr, df_tr_a, df_seat, df_prm_15, df_fact_h = [pd.DataFrame()] * 6
df_prm_d, df_fact_d, df_energy_master = [pd.DataFrame()] * 3

if any([a_ops, a_tr, a_tr_a, a_seat, a_prm_15, a_fact_h]):
    if a_ops: df_ops = pd.DataFrame(a_ops).drop_duplicates(subset=['Fecha']).sort_values("Fecha")
    if a_tr: df_tr = pd.DataFrame(a_tr).sort_values(["Fecha", "Tren"])
    if a_tr_a: df_tr_a = pd.DataFrame(a_tr_a).sort_values(["Fecha", "Tren"])
    if a_seat: df_seat = pd.DataFrame(a_seat).drop_duplicates(subset=['Fecha']).sort_values("Fecha")
    if a_prm_15: df_prm_15 = pd.DataFrame(a_prm_15)
    if a_fact_h: df_fact_h = pd.DataFrame(a_fact_h)

    # Lógica Jerarquía Energética (Factura > PRMTE > SEAT)
    if not df_seat.empty:
        df_energy_master = df_seat[["Fecha", "Total [kWh]", "Tracción [kWh]", "12 KV [kWh]"]].copy().rename(columns={"Total [kWh]":"E_Total", "Tracción [kWh]":"E_Tr", "12 KV [kWh]":"E_12"})
        df_energy_master["Fuente"] = "SEAT"

    if not df_prm_15.empty:
        df_prm_d = df_prm_15.groupby("Fecha")["kWh"].sum().reset_index()
        if not df_seat.empty:
            df_prm_d = pd.merge(df_prm_d, df_seat[["Fecha", "% Tracción", "% 12 KV"]], on="Fecha", how="left").fillna(0)
            df_prm_d["E_Tr"], df_prm_d["E_12"] = df_prm_d["kWh"]*(df_prm_d["% Tracción"]/100), df_prm_d["kWh"]*(df_prm_d["% 12 KV"]/100)
            df_prm_p = df_prm_d.rename(columns={"kWh":"E_Total"})[["Fecha","E_Total","E_Tr","E_12"]]; df_prm_p["Fuente"] = "PRMTE"
            df_energy_master = pd.concat([df_energy_master, df_prm_p]).drop_duplicates(subset=["Fecha"], keep="last")

    if not df_fact_h.empty:
        df_fact_d = df_fact_h.groupby("Fecha")["kWh"].sum().reset_index()
        if not df_seat.empty:
            df_fact_d = pd.merge(df_fact_d, df_seat[["Fecha", "% Tracción", "% 12 KV"]], on="Fecha", how="left").fillna(0)
            df_fact_d["E_Tr"], df_fact_d["E_12"] = df_fact_d["kWh"]*(df_fact_d["% Tracción"]/100), df_fact_d["kWh"]*(df_fact_d["% 12 KV"]/100)
            df_fact_f = df_fact_d.rename(columns={"kWh":"E_Total"})[["Fecha","E_Total","E_Tr","E_12"]]; df_fact_f["Fuente"] = "Factura"
            df_energy_master = pd.concat([df_energy_master, df_fact_f]).drop_duplicates(subset=["Fecha"], keep="last")

    if not df_ops.empty and not df_energy_master.empty:
        df_ops = pd.merge(df_ops, df_energy_master, on="Fecha", how="left")

# --- 6. RENDERIZADO DE TABS CON FILTROS INDIVIDUALES ---
if not df_ops.empty or not df_tr.empty or not df_seat.empty:
    tabs = st.tabs(["📊 Resumen", "📑 Operaciones", "📑 Trenes", "⚡ SEAT", "📈 PRMTE", "💰 Facturación"])
    
    with tabs[0]: # RESUMEN
        st.write("#### 🔍 Filtro Resumen")
        jornadas_disponibles = df_ops['Tipo Día'].unique().tolist() if not df_ops.empty else ["L", "S", "D/F"]
        f_res_jornada = st.multiselect("Filtrar por Jornada:", jornadas_disponibles, default=jornadas_disponibles, key="f_res_jor")
        
        df_res_filtered = df_ops[df_ops['Tipo Día'].isin(f_res_jornada)] if not df_ops.empty else df_ops
        
        if not df_res_filtered.empty:
            c1, c2, c3 = st.columns(3)
            to, tk = df_res_filtered["Odómetro [km]"].sum(), df_res_filtered["Tren-Km [km]"].sum()
            c1.metric("Odómetro Total", f"{to:,.1f} km"); c2.metric("Tren-Km Total", f"{tk:,.1f} km"); c3.metric("UMR Global", f"{(tk/to*100 if to>0 else 0):.2f} %")
            st.divider()
            
            df_res_filtered['Tipo Día'] = pd.Categorical(df_res_filtered['Tipo Día'], categories=['L', 'S', 'D/F'], ordered=True)
            res = df_res_filtered.groupby("Tipo Día", observed=True).agg({"Odómetro [km]":"sum", "Tren-Km [km]":"sum", "UMR [%]":"mean", "E_Total":"sum", "E_Tr":"sum", "E_12":"sum"}).reset_index()
            st.write("#### Balance Consumo vs Operación")
            st.table(res.style.format("{:,.1f}", subset=["Odómetro [km]", "Tren-Km [km]"]).format("{:,.2f}%", subset=["UMR [%]"]).format("{:,.0f}", subset=["E_Total", "E_Tr", "E_12"]))

    with tabs[1]: # OPERACIONES
        st.write("#### 🔍 Filtros Operacionales")
        c1, c2 = st.columns(2)
        semanas_disp = sorted(df_ops['N° Semana'].unique().tolist()) if not df_ops.empty else []
        f_ops_sem = c1.multiselect("Filtrar por N° Semana:", semanas_disp, key="f_ops_sem")
        f_ops_jor = c2.multiselect("Filtrar por Jornada:", jornadas_disponibles, default=jornadas_disponibles, key="f_ops_jor")
        
        df_ops_filtered = df_ops.copy()
        if f_ops_sem: df_ops_filtered = df_ops_filtered[df_ops_filtered['N° Semana'].isin(f_ops_sem)]
        if f_ops_jor: df_ops_filtered = df_ops_filtered[df_ops_filtered['Tipo Día'].isin(f_ops_jor)]
        
        if not df_ops_filtered.empty:
            st.dataframe(df_ops_filtered.style.format({"Fecha": lambda x: x.strftime('%d/%m/%Y'), "Odómetro [km]":"{:,.1f}", "Tren-Km [km]":"{:,.1f}", "UMR [%]":"{:.2f}%", "E_Total":"{:,.0f}", "E_Tr":"{:,.0f}", "E_12":"{:,.0f}"}), use_container_width=True)

    with tabs[2]: # TRENES
        st.write("#### 🔍 Filtro de Flota")
        trenes_disp = sorted(pd.concat([df_tr['Tren'] if not df_tr.empty else pd.Series(), df_tr_a['Tren'] if not df_tr_a.empty else pd.Series()]).unique().tolist())
        f_tr_tren = st.multiselect("Seleccionar Tren(es):", trenes_disp, key="f_tr_tren")
        
        df_tr_filt = df_tr[df_tr['Tren'].isin(f_tr_tren)] if f_tr_tren and not df_tr.empty else df_tr
        df_tra_filt = df_tr_a[df_tr_a['Tren'].isin(f_tr_tren)] if f_tr_tren and not df_tr_a.empty else df_tr_a

        if not df_tr_filt.empty:
            st.write("#### Kilometraje Diario [km]")
            st.dataframe(df_tr_filt.pivot_table(index="Tren", columns=df_tr_filt["Fecha"].dt.day, values="Valor", aggfunc='sum').fillna(0).style.format("{:,.1f}"), use_container_width=True)
        if not df_tra_filt.empty:
            st.divider(); st.write("#### Lectura Odómetro Acumulado [km]")
            st.dataframe(df_tra_filt.pivot_table(index="Tren", columns=df_tra_filt["Fecha"].dt.day, values="Valor", aggfunc='max').fillna(0).style.format("{:,.0f}"), use_container_width=True)

    with tabs[3]: # SEAT
        st.write("#### ⚡ Registro Consolidado SEAT")
        if not df_seat.empty: st.dataframe(df_seat.style.format({"Fecha": lambda x: x.strftime('%d/%m/%Y'), "Total [kWh]":"{:,.0f}", "Tracción [kWh]":"{:,.0f}", "12 KV [kWh]":"{:,.0f}", "% Tracción":"{:.2f}%", "% 12 KV":"{:.2f}%"}), use_container_width=True)

    with tabs[4]: # PRMTE
        if not df_prm_d.empty:
            st.write("#### 📅 Resumen Diario PRMTE (Distribuido)"); st.dataframe(df_prm_d.style.format({"Fecha": lambda x: x.strftime('%d/%m/%Y'), "kWh":"{:,.1f}", "E_Tr":"{:,.1f}", "E_12":"{:,.1f}"}), use_container_width=True)
            st.write("#### 🕒 Detalle 15 Minutos"); st.dataframe(df_prm_15.style.format({"kWh":"{:,.2f}"}), use_container_width=True)

    with tabs[5]: # FACTURA
        if not df_fact_d.empty:
            st.write("#### 📅 Resumen Diario Facturación"); st.dataframe(df_fact_d.style.format({"Fecha": lambda x: x.strftime('%d/%m/%Y'), "kWh":"{:,.1f}", "E_Tr":"{:,.1f}", "E_12":"{:,.1f}"}), use_container_width=True)
            st.write("#### 🕒 Detalle Horario"); st.dataframe(df_fact_h.style.format({"kWh":"{:,.2f}"}), use_container_width=True)

    st.sidebar.download_button("📥 Descargar Excel", to_excel_consolidado(df_ops, df_tr, df_tr_a, df_seat, df_prm_d, df_fact_d), "Reporte_SGE_EFE.xlsx")
else:
    st.info("👋 Sube los archivos (UMR, SEAT, PRMTE/Factura) en el panel lateral para comenzar.")
