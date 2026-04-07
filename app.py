import streamlit as st
import pandas as pd
import numpy as np
import re
import holidays
from io import BytesIO
from datetime import datetime, date, timedelta, time
import plotly.graph_objects as go
import plotly.express as px

# --- 0. SEGURIDAD DE COLUMNAS (Evita error de PyArrow) ---
def make_columns_unique(df):
    if not isinstance(df, pd.DataFrame) or df.empty:
        return df
    cols = pd.Series(df.columns)
    for dup in cols[cols.duplicated()].unique():
        cols[cols == dup] = [f"{dup}_{i}" if i != 0 else dup for i in range(sum(cols == dup))]
    df.columns = cols
    return df

# --- 1. CONFIGURACIÓN Y ESTILOS ---
st.set_page_config(page_title="Gestión de Energía - Dashboard SGE", layout="wide", page_icon="🚆")
chile_holidays = holidays.Chile()
ORDEN_TIPO_DIA = ["L", "S", "D/F"]

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

def format_hm_short(minutos_float):
    if pd.isna(minutos_float): return "00:00"
    h, m = divmod(int(minutos_float), 60)
    return f"{h:02d}:{m:02d}"

# --- 3. MOTOR THDR (A1: FECHA | 3 FILAS VACÍAS | DATA FILA 7) ---
def convertir_a_minutos(val):
    if pd.isna(val) or str(val).strip() == "": return None
    try:
        if isinstance(val, (datetime, time)): return val.hour * 60 + val.minute + (val.second / 60.0)
        s_val = str(val).strip()
        m_ss = re.search(r'(\d{1,2}):(\d{2}):(\d{2})', s_val)
        if m_ss: return int(m_ss.group(1)) * 60 + int(m_ss.group(2)) + (int(m_ss.group(3)) / 60.0)
        m_mm = re.search(r'(\d{1,2}):(\d{2})', s_val)
        if m_mm: return int(m_mm.group(1)) * 60 + int(m_mm.group(2))
        return None
    except: return None

def procesar_thdr_eficiente(file, start_date, end_date):
    try:
        df_raw = pd.read_excel(file, header=None)
        # 1. Fecha en A1 (Celda 0,0) - Formato 10126 o 010126
        val_a1 = str(df_raw.iloc[0, 0]).strip().split('.')[0]
        if len(val_a1) == 5: d, m, a = int(val_a1[0]), int(val_a1[1:3]), 2000 + int(val_a1[3:])
        elif len(val_a1) == 6: d, m, a = int(val_a1[0:2]), int(val_a1[2:4]), 2000 + int(val_a1[4:])
        else: return pd.DataFrame()
        
        fecha_dt = pd.to_datetime(date(a, m, d)).normalize()
        if not (start_date <= fecha_dt.date() <= end_date): return pd.DataFrame()

        # 2. Cabeceras y Salto de Filas (Data en index 6)
        row_est = df_raw.iloc[0].copy(); row_est[0] = np.nan
        h1 = row_est.ffill().astype(str)
        h2 = df_raw.iloc[1].fillna('').astype(str)
        cols = [f"{st.strip()}_{tipo.strip()}" if (tipo and st != 'nan') else st.strip() for st, tipo in zip(h1, h2)]
        
        df = df_raw.iloc[6:].copy() # Salto de las 3 vacías (Fila 4, 5, 6)
        df.columns = cols
        df = make_columns_unique(df).dropna(how='all', axis=0)
        
        for col in df.columns:
            if any(k in col for k in ['Hora', 'Salida', 'Llegada']):
                df[f"{col}_min"] = df[col].apply(convertir_a_minutos)
        
        c_m2 = next((c for c in df.columns if 'Motriz 2' in str(c)), None)
        df['Unidad'] = df[c_m2].apply(lambda x: 'M' if parse_latam_number(x) > 0 else 'S') if c_m2 else 'S'
        df['Tren-Km'] = 43.13 * df['Unidad'].apply(lambda x: 2 if x == 'M' else 1)
        df['Fecha_Op'] = fecha_dt
        
        col_ref = next((c for c in df.columns if ('PUERTO' in c.upper() or 'LIMACHE' in c.upper()) and 'Salida' in c and '_min' in c), None)
        if col_ref: df['Hora_Ref_Min'] = df[col_ref]
        
        return df
    except: return pd.DataFrame()

# --- 4. INICIALIZACIÓN ---
df_ops, df_tr, df_seat = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
all_ops, all_tr, all_seat, all_comp_full, all_prmte_full, all_fact_full = [], [], [], [], [], []

# --- 5. SIDEBAR ---
with st.sidebar:
    st.header("📅 Filtro Global")
    dr = st.date_input("Rango", value=(date(2026, 1, 1), date(2026, 1, 31)))
    start_date, end_date = (dr[0], dr[1]) if isinstance(dr, tuple) and len(dr) == 2 else (dr, dr)
    st.divider()
    f_v1 = st.file_uploader("1. THDR Vía 1", accept_multiple_files=True)
    f_v2 = st.file_uploader("2. THDR Vía 2", accept_multiple_files=True)
    f_umr = st.file_uploader("3. UMR / Odómetros", accept_multiple_files=True)
    f_seat_files = st.file_uploader("4. Energía SEAT", accept_multiple_files=True)
    f_bill_files = st.file_uploader("5. Facturación y PRMTE", accept_multiple_files=True)

# --- 6. PROCESAMIENTO TOTAL ---
if any([f_v1, f_v2, f_umr, f_seat_files, f_bill_files]):
    # A. UMR / TRENES
    if f_umr:
        for f in f_umr:
            try:
                xl = pd.ExcelFile(f)
                for sn in xl.sheet_names:
                    df_raw = pd.read_excel(f, sheet_name=sn, header=None)
                    h_r = next((i for i in range(min(100, len(df_raw))) if any(k in str(df_raw.iloc[i].tolist()).upper() for k in ['FECHA', 'ODO', 'KILOM'])), None)
                    if h_r is not None:
                        df_p = pd.read_excel(f, sheet_name=sn, header=h_r)
                        df_p.columns = [str(c).upper().replace('Ó','O').strip() for c in df_p.columns]
                        c_f, c_o, c_t = next((c for c in df_p.columns if 'FECHA' in c), None), next((c for c in df_p.columns if 'ODO' in c), None), next((c for c in df_p.columns if 'KM' in c), None)
                        if c_f and c_o:
                            df_p['_dt'] = pd.to_datetime(df_p[c_f], errors='coerce').dt.normalize()
                            df_filt = df_p[(df_p['_dt'].dt.date >= start_date) & (df_p['_dt'].dt.date <= end_date)].dropna(subset=['_dt'])
                            for _, r in df_filt.iterrows():
                                all_ops.append({"Fecha": r['_dt'], "Tipo Día": get_tipo_dia(r['_dt']), "Odómetro [km]": parse_latam_number(r[c_o]), "Tren-Km [km]": parse_latam_number(r[c_t]) if c_t else 0.0})
                    if any(k in sn.upper() for k in ['KIL', 'ODO']):
                        for i in range(len(df_raw)-2):
                            for j in range(1, len(df_raw.columns)):
                                v_f = pd.to_datetime(df_raw.iloc[i, j], errors='coerce')
                                if pd.notna(v_f) and start_date <= v_f.date() <= end_date:
                                    for k in range(i+3, min(i+60, len(df_raw))):
                                        tren = str(df_raw.iloc[k, 0]).strip().upper()
                                        if re.match(r'^(M|XM)', tren): all_tr.append({"Tren": tren, "Fecha": v_f.normalize(), "Valor": parse_latam_number(df_raw.iloc[k, j])})
            except: pass

    # B. SEAT
    if f_seat_files:
        for f in f_seat_files:
            try:
                df_s = pd.read_excel(f, header=None)
                for i in range(len(df_s)):
                    fs = pd.to_datetime(df_s.iloc[i, 1], errors='coerce')
                    if pd.notna(fs) and start_date <= fs.date() <= end_date:
                        all_seat.append({"Fecha": fs.normalize(), "E_Total": parse_latam_number(df_s.iloc[i, 3]), "E_Tr": parse_latam_number(df_s.iloc[i, 5]), "E_12": parse_latam_number(df_s.iloc[i, 7])})
            except: pass

    # C. PRMTE / FACTURA (CON DETALLE 15MIN/HORA)
    if f_bill_files:
        for f in f_bill_files:
            try:
                xl = pd.ExcelFile(f)
                for sn in xl.sheet_names:
                    if 'FACT' in sn.upper():
                        df_f = pd.read_excel(f, sheet_name=sn); df_f.columns = ['F', 'V']; df_f['dt'] = pd.to_datetime(df_f['F'], errors='coerce')
                        for _, r in df_f.dropna(subset=['dt']).iterrows(): 
                            v = abs(parse_latam_number(r['V']))
                            all_fact_full.append({"Fecha": r['dt'].normalize(), "Hora": r['dt'].hour, "15min": (r['dt'].minute // 15) * 15, "Consumo [kWh]": v})
                    if 'PRMTE' in sn.upper():
                        df_pd_raw = pd.read_excel(f, sheet_name=sn, header=None); h = next((i for i in range(len(df_pd_raw)) if 'AÑO' in str(df_pd_raw.iloc[i]).upper()), 0)
                        df_pd = pd.read_excel(f, sheet_name=sn, header=h); df_pd['ts'] = pd.to_datetime(df_pd[['AÑO', 'MES', 'DIA', 'HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'}))
                        for _, r in df_pd.iterrows(): 
                            v = parse_latam_number(r.get('Retiro_Energia_Activa (kWhD)', 0))
                            all_prmte_full.append({"Fecha": r['ts'].normalize(), "Hora": r['ts'].hour, "15min": r['ts'].minute, "Consumo [kWh]": v})
            except: pass

    # --- JERARQUÍA Y CONSOLIDACIÓN IDE ---
    if all_ops:
        df_ops = pd.DataFrame(all_ops).groupby("Fecha").agg({"Odómetro [km]":"sum", "Tren-Km [km]":"sum", "Tipo Día":"first"}).reset_index()
        # Diarios para Jerarquía
        df_f_d = pd.DataFrame(all_fact_full).groupby("Fecha")["Consumo [kWh]"].sum().reset_index().rename(columns={"Consumo [kWh]": "E_Fact"}) if all_fact_full else pd.DataFrame(columns=["Fecha", "E_Fact"])
        df_p_d = pd.DataFrame(all_prmte_full).groupby("Fecha")["Consumo [kWh]"].sum().reset_index().rename(columns={"Consumo [kWh]": "E_Prmte"}) if all_prmte_full else pd.DataFrame(columns=["Fecha", "E_Prmte"])
        df_s_d = pd.DataFrame(all_seat).groupby("Fecha").agg({"E_Total":"sum", "E_Tr":"sum", "E_12":"sum"}).reset_index().rename(columns={"E_Total":"E_Seat_T", "E_Tr":"E_Seat_Tr", "E_12":"E_Seat_12"}) if all_seat else pd.DataFrame(columns=["Fecha", "E_Seat_T", "E_Seat_Tr", "E_Seat_12"])
        
        df_ops = df_ops.merge(df_f_d, on="Fecha", how="left").merge(df_p_d, on="Fecha", how="left").merge(df_s_d, on="Fecha", how="left").fillna(0)
        
        def aplicar_jerarquia(row):
            if row['E_Fact'] > 0: tot, src = row['E_Fact'], "Factura"
            elif row['E_Prmte'] > 0: tot, src = row['E_Prmte'], "PRMTE"
            elif row['E_Seat_T'] > 0: tot, src = row['E_Seat_T'], "SEAT"
            else: return 0, 0, 0, 0, 0, "N/A"
            
            r_tr = row['E_Seat_Tr']/row['E_Seat_T'] if row['E_Seat_T'] > 0 else 0.85
            r_12 = row['E_Seat_12']/row['E_Seat_T'] if row['E_Seat_T'] > 0 else 0.15
            return tot, tot*r_tr, tot*r_12, r_tr*100, r_12*100, src

        df_ops[['E_Total', 'E_Tr', 'E_12', '% Tracción', '% 12 kV', 'Fuente']] = df_ops.apply(aplicar_jerarquia, axis=1, result_type='expand')
        df_ops['IDE (kWh/km)'] = df_ops.apply(lambda r: r['E_Tr'] / r['Odómetro [km]'] if r['Odómetro [km]'] > 0 else 0, axis=1)

    if f_v1: df_thdr_v1 = pd.concat([procesar_thdr_eficiente(f, start_date, end_date) for f in f_v1], ignore_index=True)
    if f_v2: df_thdr_v2 = pd.concat([procesar_thdr_eficiente(f, start_date, end_date) for f in f_v2], ignore_index=True)

# --- 7. TABS ---
tabs = st.tabs(["📊 Resumen", "📑 Operaciones", "📑 Trenes", "⚡ Energía", "⚖️ Comparación hr", "📈 Regresión", "🚨 Atípicos", "📋 THDR"])

with tabs[0]: # RESUMEN
    if not df_ops.empty:
        c1, c2, c3 = st.columns(3); c1.metric("Odómetro", f"{df_ops['Odómetro [km]'].sum():,.1f} km"); c2.metric("Tren-Km", f"{df_ops['Tren-Km [km]'].sum():,.1f} km"); c3.metric("IDE Prom", f"{df_ops['IDE (kWh/km)'].mean():.4f}")
        st.plotly_chart(go.Figure(data=[go.Bar(x=df_ops['Fecha'], y=df_ops['Odómetro [km]'], marker_color="#005195")]), use_container_width=True)

with tabs[1]: # OPERACIONES (REPARADO)
    if not df_ops.empty:
        st.write("### 📑 Detalle Operacional e IDE (Jerarquía Activa)")
        st.dataframe(make_columns_unique(df_ops).style.format({'Odómetro [km]':"{:,.1f}", 'E_Total':"{:,.0f}", 'IDE (kWh/km)':"{:.4f}"}))

with tabs[2]: # TRENES
    if all_tr: st.dataframe(pd.DataFrame(all_tr).pivot_table(index="Tren", columns="Fecha", values="Valor", aggfunc='sum').fillna(0).style.format("{:,.1f}"))

with tabs[3]: # ⚡ ENERGÍA (CON TABLAS 15MIN Y HORA)
    e_tabs = st.tabs(["🔹 SEAT", "🔹 PRMTE", "🔹 Facturación"])
    with e_tabs[1]: # PRMTE
        if all_prmte_full:
            df_p = pd.DataFrame(all_prmte_full)
            c1, c2 = st.columns(2)
            with c1: st.write("#### 📅 Consumo por Día"); st.dataframe(df_p.groupby("Fecha")["Consumo [kWh]"].sum().reset_index().style.format("{:,.2f}"))
            with c2: st.write("#### ⏱️ Consumo por Hora"); st.dataframe(df_p.groupby(["Fecha", "Hora"])["Consumo [kWh]"].sum().reset_index().style.format("{:,.2f}"))
            st.write("#### ⏲️ Consumo cada 15 min")
            st.dataframe(df_p.groupby(["Fecha", "Hora", "15min"])["Consumo [kWh]"].sum().reset_index().style.format("{:,.2f}"))
    with e_tabs[2]: # FACTURACIÓN
        if all_fact_full:
            df_f = pd.DataFrame(all_fact_full)
            c1, c2 = st.columns(2)
            with c1: st.write("#### 📅 Consumo por Día"); st.dataframe(df_f.groupby("Fecha")["Consumo [kWh]"].sum().reset_index().style.format("{:,.2f}"))
            with c2: st.write("#### ⏱️ Consumo por Hora"); st.dataframe(df_f.groupby(["Fecha", "Hora"])["Consumo [kWh]"].sum().reset_index().style.format("{:,.2f}"))
            st.write("#### ⏲️ Consumo cada 15 min")
            st.dataframe(df_f.groupby(["Fecha", "Hora", "15min"])["Consumo [kWh]"].sum().reset_index().style.format("{:,.2f}"))

with tabs[7]: # THDR (REPARADO)
    st.header("📋 Análisis THDR")
    if not df_thdr_v1.empty or not df_thdr_v2.empty:
        c1, c2 = st.columns(2)
        with c1:
            st.write("#### Servicios por Hora")
            freq = []
            if not df_thdr_v1.empty: v1h = (df_thdr_v1['Hora_Ref_Min'] // 60).value_counts().reset_index(); v1h.columns=['Hora','Vía 1']; freq.append(v1h)
            if not df_thdr_v2.empty: v2h = (df_thdr_v2['Hora_Ref_Min'] // 60).value_counts().reset_index(); v2h.columns=['Hora','Vía 2']; freq.append(v2h)
            if freq:
                res = freq[0]
                if len(freq)>1: res = pd.merge(res, freq[1], on='Hora', how='outer').fillna(0)
                res['Hora'] = res['Hora'].apply(lambda x: f"{int(x):02d}:00")
                st.table(res.sort_values('Hora').set_index('Hora'))
        with c2:
            st.write("#### Detalle")
            st.dataframe(make_columns_unique(pd.concat([df_thdr_v1, df_thdr_v2])).head(50))
    else:
        st.error("Sube la THDR y verifica que el rango incluya la fecha del archivo (Celda A1).")
