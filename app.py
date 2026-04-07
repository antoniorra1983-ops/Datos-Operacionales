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

# --- 1. CONFIGURACIÓN Y ESTILOS ---
st.set_page_config(page_title="Gestión de Energía - Dashboard SGE", layout="wide", page_icon="🚆")
chile_holidays = holidays.Chile()
ORDEN_TIPO_DIA = ["L", "S", "D/F"]

st.markdown("""
    <style>
    .stMetric { background-color: #ffffff; padding: 20px; border-radius: 10px; border-left: 5px solid #005195; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    </style>
    """, unsafe_allow_html=True)

# --- 2. FUNCIONES DE EXPORTACIÓN Y APOYO ---

def to_excel_consolidado(df_ops, df_tr, df_tr_acum, df_seat, df_p_d, df_p_15, df_fact_h, df_fact_d):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dict_dfs = {
            'Operaciones': df_ops, 'Kms_Diarios_Tren': df_tr, 
            'Odometros_Acum_Tren': df_tr_acum, 'SEAT': df_seat, 
            'PRMTE_D': df_p_d, 'PRMTE_15': df_p_15, 
            'Fact_H': df_fact_h, 'Fact_D': df_fact_d
        }
        for name, df in dict_dfs.items():
            if df is not None and not df.empty:
                df.to_excel(writer, index=False, sheet_name=name)
    return output.getvalue()

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

# --- 3. PROCESAMIENTO THDR DINÁMICO ---

DISTANCIAS = {"PU-LI": 43.13, "LI-PU": 43.13, "PU-SA": 29.11, "SA-PU": 29.11, "EB-PU": 25.40, "PU-EB": 25.40, "VM-LI": 34.03, "LI-VM": 34.03, "VM-PU": 9.10, "PU-VM": 9.10}

def procesar_thdr_avanzado(file):
    try:
        df_raw = pd.read_excel(file, header=None)
        h0 = df_raw.iloc[0].ffill().astype(str)
        h1 = df_raw.iloc[1].fillna('').astype(str)
        cols_raw = []
        for i in range(len(h0)):
            base, sub = h0[i].strip(), h1[i].strip()
            cols_raw.append(f"{base} ({sub})" if "Hora" in sub else base)
        
        final_cols, counts = [], {}
        for name in cols_raw:
            if name in counts: counts[name] += 1; final_cols.append(f"{name}_{counts[name]}")
            else: counts[name] = 0; final_cols.append(name)
        
        df = df_raw.iloc[2:].copy(); df.columns = final_cols
        c_sal = [c for c in df.columns if 'Hora Salida' in c]
        c_lle = [c for c in df.columns if 'Hora Llegada' in c]
        
        def detectar(row):
            iv, In, fv, fn = None, None, None, None
            for c in c_sal:
                v = convertir_a_minutos(row[c])
                if v is not None: iv, In = v, c.split('(')[0].strip(); break
            for c in reversed(c_lle):
                v = convertir_a_minutos(row[c])
                if v is not None: fv, fn = v, c.split('(')[0].strip(); break
            return pd.Series([iv, In, fv, fn])

        df[['T_Ini', 'Origen', 'T_Fin', 'Destino']] = df.apply(detectar, axis=1)
        
        def find_k(ks):
            for c in df.columns:
                if any(k.lower() in c.lower() for k in ks): return c
            return None

        cs, cp, cm1, cm2 = find_k(['Servicio', 'N°']), find_k(['Prog']), find_k(['Motriz 1', 'M1']), find_k(['Motriz 2', 'M2'])
        df['Servicio'] = pd.to_numeric(df[cs], errors='coerce').fillna(0).astype(int) if cs else 0
        df['Unidad'] = pd.to_numeric(df[cm2], errors='coerce').fillna(0).apply(lambda x: 'M' if x > 0 else 'S')
        df['Min_Prog'] = df[cp].apply(convertir_a_minutos) if cp else 0
        df['Retraso'] = df['T_Ini'] - df['Min_Prog']
        
        def calc_km(r):
            o, d = str(r['Origen'])[:2].upper(), str(r['Destino'])[:2].upper()
            map_e = {"PU":"PU", "VA":"PU", "LI":"LI", "VI":"VM", "EL":"EB"}
            k = f"{map_e.get(o,o)}-{map_e.get(d,d)}"
            return DISTANCIAS.get(k, 43.13) * (2 if r['Unidad'] == 'M' else 1)
            
        df['Tren-Km'] = df.apply(calc_km, axis=1)
        try:
            f_str = str(df_raw.iloc[0, 0]).split('.')[0].strip().zfill(6)
            df['Fecha_Op'] = f"{f_str[0:2]}/{f_str[2:4]}/20{f_str[4:6]}"
        except: df['Fecha_Op'] = ""
        
        return df[df['Servicio'] > 0]
    except Exception as e:
        st.error(f"Error procesando archivo: {e}"); return pd.DataFrame()

# --- 4. INICIALIZACIÓN ---
df_ops = df_tr = df_tr_acum = df_seat = df_p_d = df_thdr_v1 = df_thdr_v2 = pd.DataFrame()
all_ops, all_tr, all_tr_acum, all_seat, all_comp_full = [], [], [], [], []

# --- 5. SIDEBAR ---
with st.sidebar:
    st.header("📅 Período de Análisis")
    date_range = st.date_input("Rango", value=(date(2025, 1, 1), date(2026, 12, 31)))
    st.header("📂 Carga de Archivos")
    fv1 = st.file_uploader("1. THDR Vía 1", accept_multiple_files=True)
    fv2 = st.file_uploader("2. THDR Vía 2", accept_multiple_files=True)
    fumr = st.file_uploader("3. UMR / Odómetros", accept_multiple_files=True)
    fseat = st.file_uploader("4. Energía SEAT", accept_multiple_files=True)
    f_bill_f = st.file_uploader("5. Facturación y PRMTE", accept_multiple_files=True)

# --- 6. PROCESAMIENTO ---
if any([fv1, fv2, fumr, fseat, f_bill_f]):
    if fumr:
        for f in fumr:
            xl = pd.ExcelFile(f)
            for sn in xl.sheet_names:
                sn_up = sn.upper()
                if any(k in sn_up for k in ['UMR', 'RESUMEN']):
                    df_u = pd.read_excel(f, sheet_name=sn, header=None)
                    h_idx = next((i for i in range(min(50, len(df_u))) if 'ODO' in str(df_u.iloc[i]).upper()), None)
                    if h_idx is not None:
                        df_p = pd.read_excel(f, sheet_name=sn, header=h_idx)
                        df_p.columns = [re.sub(r'[^A-Z]', '', str(c).upper()) for c in df_p.columns]
                        if 'FECHA' in df_p.columns:
                            df_p['DT'] = pd.to_datetime(df_p['FECHA'], errors='coerce')
                            for _, r in df_p.dropna(subset=['DT']).iterrows():
                                all_ops.append({"Fecha": r['DT'].normalize(), "Tipo Día": get_tipo_dia(r['DT']), "N° Semana": r['DT'].isocalendar()[1], "Odómetro [km]": parse_latam_number(r.get('ODO',0)), "Tren-Km [km]": parse_latam_number(r.get('TRENKM',0))})
                if 'KIL' in sn_up and 'ODO' in sn_up:
                    df_tr_raw = pd.read_excel(f, sheet_name=sn, header=None)
                    for i in range(len(df_tr_raw)-2):
                        for j in range(1, len(df_tr_raw.columns)):
                            v_fch = pd.to_datetime(df_tr_raw.iloc[i,j], errors='coerce')
                            if pd.notna(v_fch):
                                for k in range(i+3, min(i+40, len(df_tr_raw))):
                                    nt = str(df_tr_raw.iloc[k,0]).strip().upper()
                                    if nt.startswith(('M','XM')):
                                        val = parse_latam_number(df_tr_raw.iloc[k,j])
                                        all_tr.append({"Tren": nt, "Fecha": v_fch.normalize(), "Valor": val})
                                        all_tr_acum.append({"Tren": nt, "Fecha": v_fch.normalize(), "Valor": val})

    if fseat:
        for f in fseat:
            df_s = pd.read_excel(f, header=None)
            for i in range(len(df_s)):
                dt = pd.to_datetime(df_s.iloc[i,1], errors='coerce')
                if pd.notna(dt):
                    all_seat.append({"Fecha": dt.normalize(), "E_Total": parse_latam_number(df_s.iloc[i,3]), "E_Tr": parse_latam_number(df_s.iloc[i,5]), "E_12": parse_latam_number(df_s.iloc[i,7])})
    
    if f_bill_f:
        for f in f_bill_f:
            xl_b = pd.ExcelFile(f)
            for sn in xl_b.sheet_names:
                df_b = pd.read_excel(f, sheet_name=sn)
                if 'AÑO' in str(df_b.columns).upper():
                    df_b['TS'] = pd.to_datetime(df_b[['AÑO', 'MES', 'DIA', 'HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'}))
                    for _, r in df_b.iterrows():
                        v_p = parse_latam_number(r.get('Retiro_Energia_Activa (kWhD)', 0))
                        all_comp_full.append({"Fecha": r['TS'].normalize(), "Hora": r['TS'].hour, "Consumo": v_p})

    if fv1: df_thdr_v1 = pd.concat([procesar_thdr_avanzado(f) for f in fv1])
    if fv2: df_thdr_v2 = pd.concat([procesar_thdr_avanzado(f) for f in fv2])

    df_ops = pd.DataFrame(all_ops).drop_duplicates(subset=['Fecha']) if all_ops else pd.DataFrame()
    if not df_ops.empty and all_seat:
        df_ops = pd.merge(df_ops, pd.DataFrame(all_seat), on="Fecha", how="left").fillna(0)
        df_ops['IDE'] = df_ops.apply(lambda r: r['E_Tr']/r['Odómetro [km]'] if r['Odómetro [km]']>0 else 0, axis=1)

# --- 7. DASHBOARD ---
tabs = st.tabs(["📊 Resumen", "📑 Operaciones", "📑 Trenes", "⚡ Energía", "⚖️ Comparativa", "📈 Regresión", "🚨 Atípicos", "📋 THDR"])

with tabs[0]: # PESTAÑA RESUMEN
    if not df_ops.empty:
        c1, c2, c3 = st.columns(3)
        anios, meses, semanas = sorted(df_ops['Fecha'].dt.year.unique()), sorted(df_ops['Fecha'].dt.month.unique()), sorted(df_ops['N° Semana'].unique())
        f_ano = c1.multiselect("Año", anios, default=anios)
        f_mes = c2.multiselect("Mes", meses, default=meses)
        f_sem = c3.multiselect("Semana", semanas, default=semanas)
        ujor = df_ops['Tipo Día'].unique()
        f_jor = st.multiselect("Jornada", [d for d in ORDEN_TIPO_DIA if d in ujor], default=ujor)
        
        df_f = df_ops[df_ops['Fecha'].dt.year.isin(f_ano) & df_ops['Fecha'].dt.month.isin(f_mes) & df_ops['N° Semana'].isin(f_sem) & df_ops['Tipo Día'].isin(f_jor)]
        
        if not df_f.empty:
            sub_tabs = st.tabs(["📅 Semanal", "📅 Mensual", "📅 Anual"])
            with sub_tabs[0]:
                m1, m2, m3 = st.columns(3)
                m1.metric("Odómetro [km]", f"{df_f['Odómetro [km]'].sum():,.1f}")
                m2.metric("Tren-Km [km]", f"{df_f['Tren-Km [km]'].sum():,.1f}")
                m3.metric("IDE [kWh/km]", f"{df_f['IDE'].mean():.4f}")
                et, e12 = df_f['E_Tr'].sum(), df_f['E_12'].sum()
                if (et+e12) > 0: st.info(f"⚡ Composición Energía: Tracción **{et/(et+e12)*100:.1f}%** | Otros 12kV **{e12/(et+e12)*100:.1f}%**")
                st.plotly_chart(go.Figure(go.Scatter(x=df_f['Fecha'], y=df_f['IDE'], name="IDE")), use_container_width=True)

with tabs[2]: # TRENES
    if all_tr:
        df_t = pd.DataFrame(all_tr)
        st.write("#### Kilometraje Diario por Unidad [km]")
        st.dataframe(df_t.pivot_table(index="Tren", columns=df_t["Fecha"].dt.day, values="Valor", aggfunc='sum').fillna(0))
    if all_tr_acum:
        st.divider(); st.write("#### Odómetro Acumulado por Unidad [km]")
        df_ta = pd.DataFrame(all_tr_acum)
        st.dataframe(df_ta.pivot_table(index="Tren", columns=df_ta["Fecha"].dt.day, values="Valor", aggfunc='max').fillna(0))

with tabs[4]: # COMPARATIVA HORARIA
    if all_comp_full:
        df_c = pd.DataFrame(all_comp_full)
        piv_c = df_c.pivot_table(index="Hora", columns=df_c['Fecha'].dt.year, values="Consumo", aggfunc='median').fillna(0)
        st.line_chart(piv_c)

with tabs[7]: # THDR
    st.write("### 📋 Tabla Horaria de Desempeño Real")
    col1, col2 = st.columns(2)
    with col1:
        st.write("#### Vía 1")
        if not df_thdr_v1.empty: st.dataframe(df_thdr_v1[['Fecha_Op', 'Servicio', 'Origen', 'Destino', 'Unidad', 'Tren-Km', 'Retraso']], use_container_width=True)
    with col2:
        st.write("#### Vía 2")
        if not df_thdr_v2.empty: st.dataframe(df_thdr_v2[['Fecha_Op', 'Servicio', 'Origen', 'Destino', 'Unidad', 'Tren-Km', 'Retraso']], use_container_width=True)

if not df_ops.empty:
    st.sidebar.download_button("📥 Reporte Completo", to_excel_consolidado(df_ops, pd.DataFrame(all_tr), pd.DataFrame(all_tr_acum), pd.DataFrame(all_seat), None, None, None, None), "Reporte_SGE_EFE.xlsx")
