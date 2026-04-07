import streamlit as st
import pandas as pd
import numpy as np
import re
import holidays
from io import BytesIO
from datetime import datetime, date, timedelta, time
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import tempfile
import os

# --- 1. CONFIGURACIÓN Y ESTILOS ---
st.set_page_config(page_title="Gestión de Energía - Dashboard SGE", layout="wide", page_icon="🚆")
chile_holidays = holidays.Chile()

ORDEN_TIPO_DIA = ["L", "S", "D/F"]

st.markdown("""
    <style>
    .stMetric { background-color: #ffffff; padding: 20px; border-radius: 10px; border-left: 5px solid #005195; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    </style>
    """, unsafe_allow_html=True)

# --- 2. FUNCIONES DE PROCESAMIENTO Y EXPORTACIÓN ---
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

def convertir_a_minutos(val):
    if pd.isna(val) or str(val).strip() == "": return None
    try:
        if isinstance(val, (datetime, time)): return val.hour * 60 + val.minute + (val.second / 60.0)
        s = str(val).strip()
        m_ss = re.search(r'(\d{1,2}):(\d{2}):(\d{2})', s)
        if m_ss: return int(m_ss.group(1)) * 60 + int(m_ss.group(2)) + (int(m_ss.group(3)) / 60.0)
        m_mm = re.search(r'(\d{1,2}):(\d{2})', s)
        if m_mm: return int(m_mm.group(1)) * 60 + int(m_mm.group(2))
        return None
    except: return None

def format_hms(minutos_float, con_signo=False):
    if pd.isna(minutos_float) or minutos_float == 0: return "00:00:00"
    signo = ("+" if minutos_float > 0 else "-" if minutos_float < 0 else "") if con_signo else ""
    total_segundos = int(round(abs(minutos_float) * 60))
    h, r = divmod(total_segundos, 3600); m, s = divmod(r, 60)
    return f"{signo}{h:02d}:{m:02d}:{s:02d}"

DISTANCIAS = {"PU-LI": 43.13, "LI-PU": 43.13, "PU-SA": 29.11, "SA-PU": 29.11, "EB-PU": 25.40, "PU-EB": 25.40, "VM-LI": 34.03, "LI-VM": 34.03, "VM-PU": 9.10, "PU-VM": 9.10}

@st.cache_data
def leer_fecha_archivo(file):
    try:
        df = pd.read_excel(file, nrows=1, header=None)
        val = str(df.iloc[0, 0]).split('.')[0].strip().zfill(6)
        return (int(val[0:2]), int(val[2:4]), 2000 + int(val[4:6]))
    except: return None

# --- 3. PROCESAMIENTO THDR (DINÁMICO PARA SERVICIOS CORTOS) ---
def procesar_thdr_avanzado(file):
    try:
        df_raw = pd.read_excel(file, header=None)
        h0 = df_raw.iloc[0].ffill().astype(str)
        h1 = df_raw.iloc[1].fillna('').astype(str)
        
        column_names_raw = []
        for i in range(len(h0)):
            base, sub = h0[i].strip(), h1[i].strip()
            column_names_raw.append(f"{base} ({sub})" if sub in ['Hora Llegada', 'Hora Salida'] else base)
        
        column_names, counts = [], {}
        for name in column_names_raw:
            if name in counts:
                counts[name] += 1
                column_names.append(f"{name}_{counts[name]}")
            else:
                counts[name] = 0
                column_names.append(name)
        
        df = df_raw.iloc[2:].copy()
        df.columns = column_names
        
        cols_salida = [c for c in df.columns if '(Hora Salida)' in c]
        cols_llegada = [c for c in df.columns if '(Hora Llegada)' in c]
        
        def get_journey(row):
            s_v, s_n, e_v, e_n = None, None, None, None
            for c in cols_salida:
                v = convertir_a_minutos(row[c])
                if v is not None: s_v, s_n = v, c.split('(')[0].strip(); break
            for c in reversed(cols_llegada):
                v = convertir_a_minutos(row[c])
                if v is not None: e_v, e_n = v, c.split('(')[0].strip(); break
            return pd.Series([s_v, s_n, e_v, e_n])

        df[['Hora_Salida_Real', 'Origen', 'Hora_Llegada_Real', 'Destino']] = df.apply(get_journey, axis=1)
        
        def buscar_columna(nombres_posibles):
            for col in df.columns:
                if any(posible.lower() in col.lower() for posible in nombres_posibles): return col
            return None
        
        c_servicio = buscar_columna(['Servicio', 'Serv', 'N° Servicio'])
        c_hora_prog = buscar_columna(['Hora_Prog', 'Hora Programada', 'Hora Prog', 'Prog'])
        c_motriz1 = buscar_columna(['Motriz 1', 'Motriz1', 'M1', 'Motor 1'])
        c_motriz2 = buscar_columna(['Motriz 2', 'Motriz2', 'M2', 'Motor 2'])
        
        df['Servicio'] = pd.to_numeric(df[c_servicio], errors='coerce').fillna(0).astype(int) if c_servicio else 0
        df['Hora_Prog'] = df[c_hora_prog] if c_hora_prog else '00:00:00'
        df['Motriz 1'] = pd.to_numeric(df[c_motriz1], errors='coerce').fillna(0).astype(int) if c_motriz1 else 0
        df['Motriz 2'] = pd.to_numeric(df[c_motriz2], errors='coerce').fillna(0).astype(int) if c_motriz2 else 0
        df['Unidad'] = df['Motriz 2'].apply(lambda x: 'M' if x > 0 else 'S')
        df['Min_Prog'] = df['Hora_Prog'].apply(convertir_a_minutos)
        df['Retraso'] = df['Hora_Salida_Real'] - df['Min_Prog']
        df['Puntual'] = (abs(df['Retraso'].fillna(999)) <= 5).astype(int)
        
        tdv = df['Hora_Llegada_Real'] - df['Hora_Salida_Real']
        df['TDV_Min'] = tdv.apply(lambda x: x if (x or 0) > 0 else (x + 1440 if pd.notna(x) else 0))
        
        def calc_km(r):
            o = str(r['Origen'])[:2].upper() if r['Origen'] else ""
            d = str(r['Destino'])[:2].upper() if r['Destino'] else ""
            map_est = {"PU":"PU", "VA":"PU", "LI":"LI", "VI":"VM", "EL":"EB"}
            key = f"{map_est.get(o, o)}-{map_est.get(d, d)}"
            return DISTANCIAS.get(key, 43.13) * (2 if r['Unidad'] == 'M' else 1)
        
        df['Tren-Km'] = df.apply(calc_km, axis=1)
        fch = leer_fecha_archivo(file)
        df['Fecha_Op'] = f"{fch[0]:02d}/{fch[1]:02d}/{fch[2]}" if fch else ''
        
        return df[df['Servicio'] > 0], df['Tren-Km'].sum(), df[df['TDV_Min'] > 0]['TDV_Min'].mean(), (df['Puntual'].sum() / len(df) * 100) if len(df) > 0 else 0
    except Exception as e:
        st.error(f"Error THDR: {e}"); return pd.DataFrame(), 0, 0, 0

# --- 4. INICIALIZACIÓN ---
df_ops = df_tr = df_tr_acum = df_seat = df_energy_master = df_p_d = df_f_d = df_thdr_v1 = df_thdr_v2 = pd.DataFrame()
all_ops, all_tr, all_tr_acum, all_seat, all_prmte_15, all_fact_h, all_comp_full = [], [], [], [], [], [], []

# --- 5. SIDEBAR ---
with st.sidebar:
    st.header("📅 Filtro Global")
    today = date.today()
    date_range = st.date_input("Selecciona el período", value=(today.replace(day=1), today))
    start_date, end_date = (date_range[0], date_range[1]) if isinstance(date_range, tuple) and len(date_range)==2 else (date_range, date_range)
    st.divider(); st.header("📂 Carga de Archivos")
    f_v1 = st.file_uploader("1. THDR Vía 1", accept_multiple_files=True)
    f_v2 = st.file_uploader("2. THDR Vía 2", accept_multiple_files=True)
    f_umr = st.file_uploader("3. UMR / Odómetros", accept_multiple_files=True)
    f_seat_f = st.file_uploader("4. Energía SEAT", accept_multiple_files=True)
    f_bill_f = st.file_uploader("5. Facturación y PRMTE", accept_multiple_files=True)

# --- 6. PROCESAMIENTO ---
if any([f_v1, f_v2, f_umr, f_seat_f, f_bill_f]):
    thdr_v1_list, thdr_v2_list = [], []
    todos = (f_v1 or []) + (f_v2 or []) + (f_umr or []) + (f_seat_f or []) + (f_bill_f or [])
    for f in todos:
        try:
            xl = pd.ExcelFile(f)
            for sn in xl.sheet_names:
                sn_up = sn.upper()
                if any(k in sn_up for k in ['UMR', 'RESUMEN']):
                    df_raw = pd.read_excel(f, sheet_name=sn, header=None)
                    h_r = next((i for i in range(min(100, len(df_raw))) if any(k in str(df_raw.iloc[i]).upper() for k in ['ODO', 'FECHA'])), None)
                    if h_r is not None:
                        df_p = pd.read_excel(f, sheet_name=sn, header=h_r)
                        df_p.columns = [re.sub(r'[^A-Z]', '', str(c).upper().replace('Ó','O')) for c in df_p.columns]
                        idx_f, idx_o, idx_t = next((c for c in df_p.columns if 'FECHA' in c), None), next((c for c in df_p.columns if 'ODO' in c and 'ACUM' not in c), None), next((c for c in df_p.columns if 'TREN' in c and 'KM' in c), None)
                        if idx_f and idx_o:
                            df_p['_dt'] = pd.to_datetime(df_p[idx_f], errors='coerce')
                            mask = (df_p['_dt'].dt.date >= start_date) & (df_p['_dt'].dt.date <= end_date)
                            for _, r in df_p[mask].dropna(subset=['_dt']).iterrows():
                                all_ops.append({"Fecha": r['_dt'].normalize(), "Tipo Día": get_tipo_dia(r['_dt']), "N° Semana": r['_dt'].isocalendar()[1], "Odómetro [km]": parse_latam_number(r[idx_o]), "Tren-Km [km]": parse_latam_number(r[idx_t]), "UMR [%]": (parse_latam_number(r[idx_t])/parse_latam_number(r[idx_o])*100 if parse_latam_number(r[idx_o])>0 else 0)})
                if 'ODO' in sn_up and 'KIL' in sn_up:
                    df_tr_raw = pd.read_excel(f, sheet_name=sn, header=None)
                    headers_found = []
                    for i in range(len(df_tr_raw)-2):
                        for j in range(1, len(df_tr_raw.columns)):
                            val = pd.to_datetime(df_tr_raw.iloc[i, j], errors='coerce')
                            if pd.notna(val) and start_date <= val.date() <= end_date:
                                if i not in [h[0] for h in headers_found]: headers_found.append((i, val))
                    for idx, (row_idx, s_dt) in enumerate(headers_found):
                        is_acum = any(k in str(df_tr_raw.iloc[row_idx:row_idx+3, 0:5]).upper() for k in ['ACUM', 'LECTURA', 'TOTAL'])
                        c_map = {j: pd.to_datetime(df_tr_raw.iloc[row_idx, j], errors='coerce') for j in range(1, len(df_tr_raw.columns)) if pd.notna(pd.to_datetime(df_tr_raw.iloc[row_idx, j], errors='coerce'))}
                        for k in range(row_idx+3, min(row_idx+40, len(df_tr_raw))):
                            n_tr = str(df_tr_raw.iloc[k, 0]).strip().upper()
                            if re.match(r'^(M|XM)', n_tr):
                                for c_idx, c_fch in c_map.items():
                                    val_km = parse_latam_number(df_tr_raw.iloc[k, c_idx])
                                    d_pt = {"Tren": n_tr, "Fecha": c_fch.normalize(), "Día": c_fch.day, "Valor": val_km}
                                    if is_acum or idx > 0: all_tr_acum.append(d_pt)
                                    else: all_tr.append(d_pt)
                if 'SEAT' in sn_up and 'SER' in sn_up:
                    df_s = pd.read_excel(f, sheet_name=sn, header=None)
                    for i in range(len(df_s)):
                        fs = pd.to_datetime(df_s.iloc[i, 1], errors='coerce')
                        if pd.notna(fs) and start_date <= fs.date() <= end_date:
                            tot, tra, k12 = parse_latam_number(df_s.iloc[i, 3]), parse_latam_number(df_s.iloc[i, 5]), parse_latam_number(df_s.iloc[i, 7])
                            all_seat.append({"Fecha": fs.normalize(), "E_Total": tot, "E_Tr": tra, "E_12": k12, "% Tracción": (tra/tot*100 if tot>0 else 0), "% 12 KV": (k12/tot*100 if tot>0 else 0)})
                if any(k in sn_up for k in ['PRMTE', 'MEDIDAS']):
                    df_pd = pd.read_excel(f, sheet_name=sn)
                    if 'AÑO' in str(df_pd.columns).upper():
                        df_pd['TS'] = pd.to_datetime(df_pd[['AÑO', 'MES', 'DIA', 'HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'}))
                        for _, r in df_pd.iterrows():
                            all_comp_full.append({"Fecha": r['TS'].normalize(), "Hora": r['TS'].hour, "Consumo": parse_latam_number(r.get('Retiro_Energia_Activa (kWhD)', 0)), "Fuente": "PRMTE"})
        except: continue
    
    if f_v1:
        for file in f_v1:
            df, _, _, _ = procesar_thdr_avanzado(file)
            if not df.empty: thdr_v1_list.append(df)
        if thdr_v1_list: df_thdr_v1 = pd.concat(thdr_v1_list, ignore_index=True)
    if f_v2:
        for file in f_v2:
            df, _, _, _ = procesar_thdr_avanzado(file)
            if not df.empty: thdr_v2_list.append(df)
        if thdr_v2_list: df_thdr_v2 = pd.concat(thdr_v2_list, ignore_index=True)

    if all_ops:
        df_ops = pd.DataFrame(all_ops).drop_duplicates(subset=['Fecha']).sort_values("Fecha")
        for c in ['E_Total', 'E_Tr', 'E_12']: df_ops[c] = 0.0
        if all_seat:
            df_s_df = pd.DataFrame(all_seat).drop_duplicates(subset=['Fecha'])
            df_ops = pd.merge(df_ops.drop(columns=['E_Total','E_Tr','E_12']), df_s_df, on="Fecha", how="left").fillna(0)
        df_ops['IDE (kWh/km)'] = df_ops.apply(lambda r: r['E_Tr']/r['Odómetro [km]'] if r['Odómetro [km]']>0 else 0, axis=1)
    if all_tr: df_tr = pd.DataFrame(all_tr).sort_values(["Fecha", "Tren"])
    if all_tr_acum: df_tr_acum = pd.DataFrame(all_tr_acum).sort_values(["Fecha", "Tren"])

# --- 7. DASHBOARD (RESTAURACIÓN TOTAL) ---
tabs = st.tabs(["📊 Resumen", "📑 Operaciones", "📑 Trenes", "⚡ Energía", "⚖️ Comparación Energía hr", "📈 Regresión Nocturna", "🚨 Datos Atípicos", "📋 THDR"])

# PESTAÑA RESUMEN
with tabs[0]:
    if not df_ops.empty:
        c1, c2, c3 = st.columns(3)
        anios, meses, semanas = sorted(df_ops['Fecha'].dt.year.unique()), sorted(df_ops['Fecha'].dt.month.unique()), sorted(df_ops['N° Semana'].unique())
        f_ano = c1.multiselect("Año", anios, default=anios, key="res_ano")
        f_mes = c2.multiselect("Mes", meses, default=meses, key="res_mes")
        f_sem = c3.multiselect("Semana", semanas, default=semanas, key="res_sem")
        unique_jor = df_ops['Tipo Día'].unique()
        f_jor = st.multiselect("Jornada", [d for d in ORDEN_TIPO_DIA if d in unique_jor], default=unique_jor, key="res_jor")
        
        df_res_f = df_ops[df_ops['Fecha'].dt.year.isin(f_ano) & df_ops['Fecha'].dt.month.isin(f_mes) & df_ops['N° Semana'].isin(f_sem) & df_ops['Tipo Día'].isin(f_jor)]
        
        if not df_res_f.empty:
            sub_tabs = st.tabs(["📅 Semanal", "📅 Mensual", "📅 Anual"])
            with sub_tabs[0]:
                to_val, tk_val = df_res_f["Odómetro [km]"].sum(), df_res_f["Tren-Km [km]"].sum()
                col1, col2, col3 = st.columns(3)
                col1.metric("Odómetro Total", f"{to_val:,.1f} km")
                col2.metric("Tren-Km Total", f"{tk_val:,.1f} km")
                col3.metric("UMR Global", f"{(tk_val/to_val*100):.2f}%" if to_val>0 else "0%")
                e_tr, e_12 = df_res_f['E_Tr'].sum(), df_res_f['E_12'].sum()
                if (e_tr + e_12) > 0:
                    st.info(f"⚡ Composición Energía: Tracción {e_tr/(e_tr+e_12)*100:.1f}% | 12kV {e_12/(e_tr+e_12)*100:.1f}%")
                st.plotly_chart(go.Figure(go.Scatter(x=df_res_f['Fecha'], y=df_res_f['Odómetro [km]'], name="Km")), use_container_width=True)

# PESTAÑA TRENES (RESTAURADA)
with tabs[2]:
    if not df_tr.empty:
        st.write("#### Kilometraje Diario [km]")
        piv = df_tr.pivot_table(index="Tren", columns=df_tr["Fecha"].dt.day, values="Valor", aggfunc='sum').fillna(0)
        st.dataframe(piv.style.format("{:,.1f}"), use_container_width=True)
    if not df_tr_acum.empty:
        st.divider(); st.write("#### Odómetro Acumulado [km]")
        piv_a = df_tr_acum.pivot_table(index="Tren", columns=df_tr_acum["Fecha"].dt.day, values="Valor", aggfunc='max').fillna(0)
        st.dataframe(piv_a.style.format("{:,.0f}"), use_container_width=True)

# PESTAÑA THDR (CORREGIDA)
with tabs[7]:
    st.write("### 📋 Datos THDR Dinámicos")
    c1, c2 = st.columns(2)
    with c1:
        st.write("#### Vía 1 (Detección de Origen/Destino)")
        if not df_thdr_v1.empty:
            cols = ['Fecha_Op', 'Servicio', 'Origen', 'Destino', 'Unidad', 'Tren-Km', 'Retraso']
            st.dataframe(df_thdr_v1[[c for c in cols if c in df_thdr_v1.columns]], use_container_width=True)
    with c2:
        st.write("#### Vía 2 (Detección de Origen/Destino)")
        if not df_thdr_v2.empty:
            cols = ['Fecha_Op', 'Servicio', 'Origen', 'Destino', 'Unidad', 'Tren-Km', 'Retraso']
            st.dataframe(df_thdr_v2[[c for c in cols if c in df_thdr_v2.columns]], use_container_width=True)

# PESTAÑAS ADICIONALES (RESTANTE)
with tabs[4]:
    if all_comp_full:
        df_c = pd.DataFrame(all_comp_full)
        piv_c = df_c.pivot_table(index="Hora", columns=df_c['Fecha'].dt.year, values="Consumo", aggfunc='median').fillna(0)
        st.line_chart(piv_c)

st.sidebar.download_button("📥 Reporte Completo", to_excel_consolidado(df_ops, df_tr, df_tr_acum, df_seat, pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()), "Reporte_EFE.xlsx")
