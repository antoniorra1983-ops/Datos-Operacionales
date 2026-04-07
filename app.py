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
import traceback

# --- 0. FUNCIÓN DE SEGURIDAD PARA COLUMNAS DUPLICADAS (REPARADA) ---
def make_columns_unique(df):
    """
    Evita el ValueError de PyArrow. Recibe un DataFrame y devuelve un DataFrame con nombres únicos.
    """
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

# --- 2. FUNCIONES DE EXPORTACIÓN ---
def to_pptx(title_text, df=None, metrics_dict=None):
    prs = Presentation()
    slide_layout = prs.slide_layouts[5] 
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    title.text = f"EFE Valparaíso: {title_text}"
    y_cursor = Inches(1.5)
    if metrics_dict:
        txBox = slide.shapes.add_textbox(Inches(0.5), y_cursor, Inches(9), Inches(1))
        tf = txBox.text_frame
        for k, v in metrics_dict.items():
            p = tf.add_paragraph()
            p.text = f"• {k}: {v}"; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = RGBColor(0, 81, 149)
        y_cursor += Inches(1.2)
    if df is not None and not df.empty:
        df_display = df.head(12).reset_index(drop=True)
        rows, cols = df_display.shape
        table = slide.shapes.add_table(rows + 1, cols, Inches(0.5), y_cursor, Inches(9), Inches(3)).table
        for c, col_name in enumerate(df_display.columns):
            cell = table.cell(0, c); cell.text = str(col_name); cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(0, 81, 149) 
            p = cell.text_frame.paragraphs[0]; p.font.color.rgb = RGBColor(255, 255, 255); p.font.size = Pt(10); p.font.bold = True
        for r in range(rows):
            for c in range(cols):
                val = df_display.iloc[r, c]; formatted_val = str(val) if not isinstance(val, float) else f"{val:,.1f}"
                table.cell(r + 1, c).text = formatted_val; table.cell(r + 1, c).text_frame.paragraphs[0].font.size = Pt(9)
    binary_output = BytesIO(); prs.save(binary_output); return binary_output.getvalue()

def exportar_resumen_excel(metrics_dict, df_resumen_jornada, df_energia, df_datos_semanales=None):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_metrics = pd.DataFrame([metrics_dict]).T.reset_index(); df_metrics.columns = ['Métrica', 'Valor']
        df_metrics.to_excel(writer, sheet_name='Métricas', index=False)
        if df_resumen_jornada is not None: df_resumen_jornada.to_excel(writer, sheet_name='Resumen_Jornada', index=False)
        if df_energia is not None: df_energia.to_excel(writer, sheet_name='Energía_Prioridad', index=False)
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

def to_excel_consolidado(df_ops, df_tr, df_tr_acum, df_seat, df_p_d, df_p_15, df_fact_h, df_fact_d):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dfs = {'Operaciones': df_ops, 'Kms_Diarios_Tren': df_tr, 'Odometros_Acum_Tren': df_tr_acum, 'SEAT': df_seat, 'PRMTE_D': df_p_d, 'PRMTE_15': df_p_15, 'Fact_H': df_fact_h, 'Fact_D': df_fact_d}
        for name, df_to_save in dfs.items():
            if not df_to_save.empty: df_to_save.to_excel(writer, index=False, sheet_name=name)
    return output.getvalue()

# --- 3. FUNCIONES THDR ---
def convertir_a_minutos(val):
    if pd.isna(val) or str(val).strip() == "": return None
    try:
        if isinstance(val, (datetime, time)): return val.hour * 60 + val.minute + (val.second / 60.0)
        m_ss = re.search(r'(\d{1,2}):(\d{2}):(\d{2})', str(val))
        if m_ss: return int(m_ss.group(1)) * 60 + int(m_ss.group(2)) + (int(m_ss.group(3)) / 60.0)
        m_mm = re.search(r'(\d{1,2}):(\d{2})', str(val))
        if m_mm: return int(m_mm.group(1)) * 60 + int(m_mm.group(2))
        return None
    except: return None

def format_hms(minutos_float):
    if pd.isna(minutos_float) or minutos_float == 0: return "00:00:00"
    total_segundos = int(round(abs(minutos_float) * 60))
    h, r = divmod(total_segundos, 3600); m, s = divmod(r, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"

DISTANCIAS = {"PU-LI": 43.13, "LI-PU": 43.13, "PU-SA": 29.11, "SA-PU": 29.11, "EB-PU": 25.40, "PU-EB": 25.40, "VM-LI": 34.03, "LI-VM": 34.03, "VM-PU": 9.10,  "PU-VM": 9.10}

def procesar_thdr_avanzado(file, start_date=None, end_date=None):
    try:
        try: df_raw = pd.read_excel(file, header=None)
        except: df_raw = pd.read_excel(file, header=None, engine='xlrd')
        h0, h1 = df_raw.iloc[0].fillna('').astype(str), df_raw.iloc[1].fillna('').astype(str)
        cols = [f"{h0[i].strip()}_{h1[i].strip()}" if h1[i].strip() in ['Hora Llegada', 'Hora Salida'] else h0[i].strip() for i in range(len(h0))]
        df = df_raw.iloc[2:].copy(); df.columns = cols; df = make_columns_unique(df)
        
        c_serv = next((c for c in df.columns if 'Servicio' in str(c) or 'N° Serv' in str(c)), None)
        c_prog = next((c for c in df.columns if 'Prog' in str(c)), None)
        c_m1 = next((c for c in df.columns if 'Motriz 1' in str(c)), None)
        c_m2 = next((c for c in df.columns if 'Motriz 2' in str(c)), None)
        
        df['Servicio'] = df[c_serv] if c_serv else 0
        df['Hora_Prog'] = df[c_prog] if c_prog else '00:00:00'
        df['Motriz 1'] = pd.to_numeric(df[c_m1], errors='coerce').fillna(0).astype(int) if c_m1 else 0
        df['Motriz 2'] = pd.to_numeric(df[c_m2], errors='coerce').fillna(0).astype(int) if c_m2 else 0
        df['Unidad'] = df['Motriz 2'].apply(lambda x: 'M' if x > 0 else 'S')
        
        columnas_horas = {}
        for col in df.columns:
            if 'hora salida' in str(col).lower(): columnas_horas[f"{str(col).lower().replace('hora salida','').strip()}_salida"] = col
            elif 'hora llegada' in str(col).lower(): columnas_horas[f"{str(col).lower().replace('hora llegada','').strip()}_llegada"] = col
        
        for k, c in columnas_horas.items():
            df[f"{k}_min"] = df[c].apply(convertir_a_minutos)
            df[f"{k}_fmt"] = df[f"{k}_min"].apply(lambda x: format_hms(x) if pd.notna(x) else "")
            
        p_key, l_key = next((k for k in columnas_horas.keys() if 'puerto' in k and 'salida' in k), None), next((k for k in columnas_horas.keys() if 'limache' in k and 'llegada' in k), None)
        df['Hora_Salida_Real'] = df[f"{p_key}_min"] if p_key else None
        df['Hora_Llegada_Real'] = df[f"{l_key}_min"] if l_key else None
        df['Retraso'] = df['Hora_Salida_Real'] - df['Hora_Prog'].apply(convertir_a_minutos)
        df['TDV_Min'] = (df['Hora_Llegada_Real'] - df['Hora_Salida_Real']).apply(lambda x: x if (pd.notna(x) and x > 0) else (x + 1440 if pd.notna(x) else 0))
        df['Tipo_Rec'] = "PU-LI" if p_key and l_key else "OTRO"
        df['Tren-Km'] = df['Tipo_Rec'].map(DISTANCIAS).fillna(0) * df['Unidad'].apply(lambda x: 2 if x == 'M' else 1)
        
        patron = re.search(r'(\d{2})(\d{2})(\d{2})', file.name)
        df['Fecha_Op'] = pd.to_datetime(date(2000+int(patron.group(3)), int(patron.group(2)), int(patron.group(1)))) if patron else pd.to_datetime(date.today())
        return df[(df['Fecha_Op'].dt.date >= start_date) & (df['Fecha_Op'].dt.date <= end_date)] if start_date else df
    except: return pd.DataFrame()

# --- 4. INICIALIZACIÓN ---
df_ops, df_tr, df_tr_acum, df_seat, df_energy_master, df_p_d, df_f_d = [pd.DataFrame() for _ in range(7)]
df_thdr_v1, df_thdr_v2 = pd.DataFrame(), pd.DataFrame()
all_ops, all_tr, all_tr_acum, all_seat, all_comp_full, all_prmte_15, all_fact_h = [], [], [], [], [], [], []

# --- 5. SIDEBAR ---
with st.sidebar:
    st.header("📅 Filtro Global")
    dr = st.date_input("Período", value=(date.today().replace(day=1), date.today()))
    start_date, end_date = (dr[0], dr[1]) if isinstance(dr, tuple) and len(dr)==2 else (dr, dr)
    st.divider()
    f_v1 = st.file_uploader("1. THDR Vía 1", type=["xls", "xlsx"], accept_multiple_files=True)
    f_v2 = st.file_uploader("2. THDR Vía 2", type=["xls", "xlsx"], accept_multiple_files=True)
    f_umr = st.file_uploader("3. UMR / Odómetros", type=["xlsx"], accept_multiple_files=True)
    f_seat_files = st.file_uploader("4. Energía SEAT", type=["xlsx"], accept_multiple_files=True)
    f_bill_files = st.file_uploader("5. Facturación y PRMTE", type=["xlsx"], accept_multiple_files=True)

# --- 6. PROCESAMIENTO ---
if any([f_v1, f_v2, f_umr, f_seat_files, f_bill_files]):
    todos = (f_v1 or []) + (f_v2 or []) + (f_umr or []) + (f_seat_files or []) + (f_bill_files or [])
    for f in todos:
        try:
            xl = pd.ExcelFile(f)
            for sn in xl.sheet_names:
                sn_up = sn.upper()
                if any(k in sn_up for k in ['UMR', 'RESUMEN']):
                    df_raw = pd.read_excel(f, sheet_name=sn, header=None)
                    h_r = next((i for i in range(min(100, len(df_raw))) if any(k in str(df_raw.iloc[i].tolist()).upper() for k in ['ODO', 'FECHA'])), None)
                    if h_r is not None:
                        df_p = pd.read_excel(f, sheet_name=sn, header=h_r)
                        df_p.columns = [re.sub(r'[^A-Z]', '', str(c).upper()) for c in df_p.columns]
                        idx_f, idx_o, idx_t = next((c for c in df_p.columns if 'FECHA' in c), None), next((c for c in df_p.columns if 'ODO' in c), None), next((c for c in df_p.columns if 'TREN' in c and 'KM' in c), None)
                        if idx_f and idx_o:
                            df_p['_dt'] = pd.to_datetime(df_p[idx_f], errors='coerce')
                            for _, r in df_p[(df_p['_dt'].dt.date >= start_date) & (df_p['_dt'].dt.date <= end_date)].dropna(subset=['_dt']).iterrows():
                                all_ops.append({"Fecha": r['_dt'].normalize(), "Tipo Día": get_tipo_dia(r['_dt']), "N° Semana": r['_dt'].isocalendar()[1], "Odómetro [km]": parse_latam_number(r[idx_o]), "Tren-Km [km]": parse_latam_number(r[idx_t]), "UMR [%]": (parse_latam_number(r[idx_t])/parse_latam_number(r[idx_o])*100 if parse_latam_number(r[idx_o])>0 else 0)})

                if 'ODO' in sn_up and 'KIL' in sn_up:
                    df_tr_raw = pd.read_excel(f, sheet_name=sn, header=None)
                    for i in range(len(df_tr_raw)-2):
                        for j in range(1, len(df_tr_raw.columns)):
                            v = pd.to_datetime(df_tr_raw.iloc[i, j], errors='coerce')
                            if pd.notna(v) and start_date <= v.date() <= end_date:
                                for k in range(i+3, min(i+40, len(df_tr_raw))):
                                    n_tr = str(df_tr_raw.iloc[k, 0]).strip().upper()
                                    if re.match(r'^(M|XM)', n_tr): all_tr.append({"Tren": n_tr, "Fecha": v.normalize(), "Valor": parse_latam_number(df_tr_raw.iloc[k, j])})

                if 'SEAT' in sn_up and 'SER' in sn_up:
                    df_s = pd.read_excel(f, sheet_name=sn, header=None)
                    for i in range(len(df_s)):
                        fs = pd.to_datetime(df_s.iloc[i, 1], errors='coerce')
                        if pd.notna(fs) and start_date <= fs.date() <= end_date:
                            all_seat.append({"Fecha": fs.normalize(), "Total [kWh]": parse_latam_number(df_s.iloc[i, 3]), "Tracción [kWh]": parse_latam_number(df_s.iloc[i, 5]), "12 KV [kWh]": parse_latam_number(df_s.iloc[i, 7]), "% Tracción": (parse_latam_number(df_s.iloc[i, 5])/parse_latam_number(df_s.iloc[i, 3])*100 if parse_latam_number(df_s.iloc[i, 3])>0 else 0), "% 12 KV": (parse_latam_number(df_s.iloc[i, 7])/parse_latam_number(df_s.iloc[i, 3])*100 if parse_latam_number(df_s.iloc[i, 3])>0 else 0)})

                if any(k in sn_up for k in ['PRMTE', 'MEDIDAS']):
                    df_prm = pd.read_excel(f, sheet_name=sn, header=None)
                    h_idx = next((i for i in range(len(df_prm)) if 'AÑO' in str(df_prm.iloc[i]).upper()), None)
                    if h_idx is not None:
                        df_pd = pd.read_excel(f, sheet_name=sn, header=h_idx)
                        df_pd['Timestamp'] = pd.to_datetime(df_pd[['AÑO', 'MES', 'DIA', 'HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'})) + pd.to_timedelta(df_pd['INICIO INTERVALO'].astype(int), unit='m')
                        cols_e = [c for c in df_pd.columns if 'Retiro_Energia_Activa (kWhD)' in str(c)]
                        for _, r in df_pd.iterrows():
                            ts = r['Timestamp']; val_p = sum([parse_latam_number(r[col]) for col in cols_e])
                            all_comp_full.append({"Fecha": ts.normalize(), "Hora": ts.hour, "Consumo": val_p, "Fuente": "PRMTE", "Año": ts.year, "Tipo Día": get_tipo_dia(ts)})
                            if start_date <= ts.date() <= end_date: all_prmte_15.append({"Fecha": ts, "Energía PRMTE [kWh]": val_p})

                if any(k in sn_up for k in ['FACTURA', 'CONSUMO']):
                    df_f = pd.read_excel(f, sheet_name=sn); df_f.columns = ['FechaHora', 'Valor']
                    df_f['Timestamp'] = pd.to_datetime(df_f['FechaHora'], errors='coerce')
                    for _, r in df_f.dropna(subset=['Timestamp']).iterrows():
                        ts = r['Timestamp']; val_f = abs(parse_latam_number(r['Valor']))
                        all_comp_full.append({"Fecha": ts.normalize(), "Hora": ts.hour, "Consumo": val_f, "Fuente": "Factura", "Año": ts.year, "Tipo Día": get_tipo_dia(ts)})
                        if start_date <= ts.date() <= end_date: all_fact_h.append({"Fecha": ts, "Consumo": val_f})
        except: continue

    if f_v1: df_thdr_v1 = make_columns_unique(pd.concat([procesar_thdr_avanzado(f, start_date, end_date) for f in f_v1], ignore_index=True))
    if f_v2: df_thdr_v2 = make_columns_unique(pd.concat([procesar_thdr_avanzado(f, start_date, end_date) for f in f_v2], ignore_index=True))

    if all_ops:
        df_ops = pd.DataFrame(all_ops).drop_duplicates(subset=['Fecha']).sort_values("Fecha")
        if all_seat:
            df_seat = pd.DataFrame(all_seat).drop_duplicates(subset=['Fecha']).sort_values("Fecha")
            df_ops = pd.merge(df_ops, df_seat, on="Fecha", how="left")
            df_ops['IDE (kWh/km)'] = df_ops.apply(lambda r: r['Tracción [kWh]'] / r['Odómetro [km]'] if r['Odómetro [km]'] > 0 else 0, axis=1)

# --- 7. DASHBOARD TABS ---
tabs = st.tabs(["📊 Resumen", "📑 Operaciones", "📑 Trenes", "⚡ Energía", "⚖️ Comparación Energía hr", "📈 Regresión Nocturna", "🚨 Datos Atípicos", "📋 THDR"])

with tabs[0]: # PESTAÑA RESUMEN
    if not df_ops.empty:
        if 'filtros_compartidos' not in st.session_state: st.session_state.filtros_compartidos = {'anios': [], 'meses': [], 'jornadas': []}
        anios, meses = sorted(df_ops['Fecha'].dt.year.unique()), sorted(df_ops['Fecha'].dt.month.unique())
        c1, c2 = st.columns(2)
        f_a = c1.multiselect("Año", anios, default=[a for a in st.session_state.filtros_compartidos['anios'] if a in anios] or anios)
        f_m = c2.multiselect("Mes", meses, default=[m for m in st.session_state.filtros_compartidos['meses'] if m in meses] or meses)
        st.session_state.filtros_compartidos['anios'], st.session_state.filtros_compartidos['meses'] = f_a, f_m
        df_rf = df_ops[(df_ops['Fecha'].dt.year.isin(f_a)) & (df_ops['Fecha'].dt.month.isin(f_m))]
        if not df_rf.empty:
            m1, m2, m3 = st.columns(3)
            m1.metric("Odómetro", f"{df_rf['Odómetro [km]'].sum():,.1f} km")
            m2.metric("Tren-Km", f"{df_rf['Tren-Km [km]'].sum():,.1f} km")
            m3.metric("IDE Prom", f"{df_rf['IDE (kWh/km)'].mean():.4f}")
            st.plotly_chart(go.Figure(data=[go.Bar(x=df_rf['Fecha'], y=df_rf['Odómetro [km]'])]), use_container_width=True)
    else: st.info("Sube archivos para comenzar.")

with tabs[1]: # PESTAÑA OPERACIONES
    if not df_ops.empty: st.dataframe(make_columns_unique(df_ops).style.format({'Odómetro [km]': "{:,.1f}", 'IDE (kWh/km)': "{:.4f}"}))

with tabs[2]: # PESTAÑA TRENES
    if all_tr:
        df_tr_piv = pd.DataFrame(all_tr).pivot_table(index="Tren", columns="Fecha", values="Valor", aggfunc='sum').fillna(0)
        st.dataframe(make_columns_unique(df_tr_piv).style.format("{:,.1f}"))

with tabs[3]: # PESTAÑA ENERGÍA
    if not df_seat.empty: st.dataframe(make_columns_unique(df_seat).style.format({'Total [kWh]': "{:,.0f}", 'Tracción [kWh]': "{:,.0f}"}))

with tabs[4]: # COMPARACIÓN ENERGÍA HR
    if all_comp_full:
        df_c = pd.DataFrame(all_comp_full).groupby(['Año', 'Tipo Día', 'Hora'])['Consumo'].median().reset_index()
        fig = px.line(df_c, x='Hora', y='Consumo', color='Año', line_dash='Tipo Día', title="Mediana de Consumo Horario")
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(make_columns_unique(df_c.pivot_table(index="Hora", columns=["Año", "Tipo Día"], values="Consumo")).style.format("{:,.1f}"))

with tabs[5]: # REGRESIÓN NOCTURNA
    if all_comp_full:
        df_reg = pd.DataFrame(all_comp_full)[pd.DataFrame(all_comp_full)['Hora'] <= 5]
        Q1, Q3 = df_reg['Consumo'].quantile(0.25), df_reg['Consumo'].quantile(0.75); IQR = Q3 - Q1
        lim_sup, lim_inf = Q3 + 1.5*IQR, Q1 - 1.5*IQR
        df_norm = df_reg[(df_reg['Consumo'] >= lim_inf) & (df_reg['Consumo'] <= lim_sup)].copy()
        st.session_state.outliers = df_reg[(df_reg['Consumo'] < lim_inf) | (df_reg['Consumo'] > lim_sup)].copy()
        x, y = np.arange(len(df_norm)), df_norm['Consumo'].values
        m, n = np.polyfit(x, y, 1); y_pred = m * x + n
        r2 = 1 - (np.sum((y - y_pred)**2) / np.sum((y - np.mean(y))**2))
        st.plotly_chart(px.scatter(df_norm, x=df_norm.index, y='Consumo', trendline="ols"))
        st.write(f"**Ecuación:** $Consumo = {m:.4f}x + {n:.2f}$ | $R^2 = {r2:.4f}$")

with tabs[6]: # DATOS ATÍPICOS
    if 'outliers' in st.session_state and not st.session_state.outliers.empty:
        st.error(f"Se detectaron {len(st.session_state.outliers)} anomalías.")
        st.dataframe(make_columns_unique(st.session_state.outliers))
    else: st.success("No hay anomalías detectadas.")

with tabs[7]: # THDR
    st.header("📋 Datos THDR")
    if not df_thdr_v1.empty: st.subheader("Vía 1"); st.dataframe(make_columns_unique(df_thdr_v1))
    if not df_thdr_v2.empty: st.subheader("Vía 2"); st.dataframe(make_columns_unique(df_thdr_v2))

# REPORTE FINAL SIDEBAR
st.sidebar.download_button("📥 Excel Completo", to_excel_consolidado(df_ops, pd.DataFrame(all_tr), pd.DataFrame(), df_seat, pd.DataFrame(), pd.DataFrame(all_prmte_15), pd.DataFrame(all_fact_h), pd.DataFrame()), "EFE_Dashboard_SGE.xlsx")
