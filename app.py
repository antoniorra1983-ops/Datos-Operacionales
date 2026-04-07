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
            p.text = f"• {k}: {v}"
            p.font.size = Pt(16)
            p.font.bold = True
            p.font.color.rgb = RGBColor(0, 81, 149)
        y_cursor += Inches(1.2)
    if df is not None and not df.empty:
        df_display = df.head(12).reset_index(drop=True)
        rows, cols = df_display.shape
        table = slide.shapes.add_table(rows + 1, cols, Inches(0.5), y_cursor, Inches(9), Inches(3)).table
        for c, col_name in enumerate(df_display.columns):
            cell = table.cell(0, c)
            cell.text = str(col_name)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(0, 81, 149) 
            p = cell.text_frame.paragraphs[0]
            p.font.color.rgb = RGBColor(255, 255, 255)
            p.font.size = Pt(10)
            p.font.bold = True
        for r in range(rows):
            for c in range(cols):
                val = df_display.iloc[r, c]
                formatted_val = str(val) if not isinstance(val, float) else f"{val:,.1f}"
                table.cell(r + 1, c).text = formatted_val
                table.cell(r + 1, c).text_frame.paragraphs[0].font.size = Pt(9)
    binary_output = BytesIO()
    prs.save(binary_output)
    return binary_output.getvalue()

def exportar_resumen_excel(metrics_dict, df_resumen_jornada, df_energia, df_datos_semanales=None):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_metrics = pd.DataFrame([metrics_dict]).T.reset_index()
        df_metrics.columns = ['Métrica', 'Valor']
        df_metrics.to_excel(writer, sheet_name='Métricas', index=False)
        if df_resumen_jornada is not None and not df_resumen_jornada.empty:
            df_resumen_jornada.to_excel(writer, sheet_name='Resumen_Jornada', index=False)
        if df_energia is not None and not df_energia.empty:
            df_energia.to_excel(writer, sheet_name='Energía_Prioridad', index=False)
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
        dfs = {'Operaciones': df_ops, 'Kms_Diarios_Tren': df_tr, 'Odometros_Acum_Tren': df_tr_acum,
               'SEAT': df_seat, 'PRMTE_D': df_p_d, 'PRMTE_15': df_p_15, 'Fact_H': df_fact_h, 'Fact_D': df_fact_d}
        for name, df in dfs.items():
            if not df.empty: df.to_excel(writer, index=False, sheet_name=name)
    return output.getvalue()

# --- 3. FUNCIONES PARA PROCESAR THDR ---
def convertir_a_minutos(val):
    if pd.isna(val) or str(val).strip() == "": return None
    try:
        if isinstance(val, (datetime, time)): return val.hour * 60 + val.minute + (val.second / 60.0)
        s = str(val).strip()
        m = re.search(r'(\d{1,2}):(\d{2}):?(\d{2})?', s)
        if m:
            mins = int(m.group(1)) * 60 + int(m.group(2))
            if m.group(3): mins += int(m.group(3)) / 60.0
            return mins
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

def procesar_thdr_avanzado(file):
    try:
        df_raw = pd.read_excel(file, header=None)
        header0 = df_raw.iloc[0].ffill().astype(str) # CORREGIDO
        header1 = df_raw.iloc[1].fillna('').astype(str)
        df = df_raw.iloc[2:].copy()
        df.columns = [f"{a}_{b}".strip('_ ') for a, b in zip(header0, header1)]
        def find_c(ks):
            for c in df.columns:
                if any(k.lower() in c.lower() for k in ks): return c
            return None
        c_s, c_m1, c_m2, c_p = find_c(['Servicio', 'N°']), find_c(['Motriz 1', 'M1']), find_c(['Motriz 2', 'M2']), find_c(['Prog'])
        c_ps = find_c(['Puerto_Hora Salida', 'Puerto_Salida']) or find_c(['Puerto'])
        c_ll = find_c(['Limache_Hora Llegada', 'Limache_Llegada']) or find_c(['Limache'])
        df['Servicio'] = pd.to_numeric(df[c_s], errors='coerce').fillna(0).astype(int) if c_s else 0
        df['Motriz 1'] = pd.to_numeric(df[c_m1], errors='coerce').fillna(0).astype(int) if c_m1 else 0
        df['Motriz 2'] = pd.to_numeric(df[c_m2], errors='coerce').fillna(0).astype(int) if c_m2 else 0
        df['Unidad'] = df['Motriz 2'].apply(lambda x: 'M' if x > 0 else 'S')
        df['Min_Prog'] = df[c_p].apply(convertir_a_minutos) if c_p else 0
        df['Hora_Salida_Real'] = df[c_ps].apply(convertir_a_minutos) if c_ps else None
        df['Hora_Llegada_Real'] = df[c_ll].apply(convertir_a_minutos) if c_ll else None
        df['Retraso'] = df['Hora_Salida_Real'] - df['Min_Prog']
        df['Puntual'] = df['Retraso'].apply(lambda x: 1 if pd.notna(x) and abs(x) <= 5 else 0)
        df['TDV_Min'] = df.apply(lambda r: (r['Hora_Llegada_Real'] - r['Hora_Salida_Real'] + (1440 if (r['Hora_Llegada_Real'] or 0) < (r['Hora_Salida_Real'] or 0) else 0)) if pd.notna(r['Hora_Salida_Real']) else 0, axis=1)
        df['Tipo_Rec'] = "PU-LI"; df['Dist_Base'] = 43.13; df['Tren-Km'] = df['Dist_Base'] * df['Unidad'].apply(lambda x: 2 if x == 'M' else 1)
        fch = leer_fecha_archivo(file)
        df['Fecha_Op'] = f"{fch[0]:02d}/{fch[1]:02d}/{fch[2]}" if fch else ""
        return df[df['Servicio'] > 0], df['Tren-Km'].sum(), df[df['TDV_Min']>0]['TDV_Min'].mean(), (df['Puntual'].sum()/len(df)*100 if len(df)>0 else 0)
    except: return pd.DataFrame(), 0, 0, 0

# --- 4. INICIALIZACIÓN ---
df_ops = df_tr = df_tr_acum = df_seat = df_energy_master = df_p_d = df_f_d = df_thdr_v1 = df_thdr_v2 = pd.DataFrame()
all_ops, all_tr, all_tr_acum, all_seat, all_prmte_15, all_fact_h, all_comp_full = [], [], [], [], [], [], []

# --- 5. SIDEBAR ---
with st.sidebar:
    st.header("📅 Filtro Global")
    date_range = st.date_input("Período", value=(date(2025, 1, 1), date.today()))
    start_date, end_date = (date_range[0], date_range[1]) if len(date_range)==2 else (date_range, date_range)
    st.header("📂 Carga de Archivos")
    f_v1 = st.file_uploader("1. THDR Vía 1", accept_multiple_files=True)
    f_v2 = st.file_uploader("2. THDR Vía 2", accept_multiple_files=True)
    f_umr = st.file_uploader("3. UMR / Odómetros", accept_multiple_files=True)
    f_seat_files = st.file_uploader("4. Energía SEAT", accept_multiple_files=True)
    f_bill_files = st.file_uploader("5. Facturación y PRMTE", accept_multiple_files=True)

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
                    h_r = next((i for i in range(min(50, len(df_raw))) if any(k in str(df_raw.iloc[i]).upper() for k in ['ODO', 'FECHA'])), None)
                    if h_r is not None:
                        df_p = pd.read_excel(f, sheet_name=sn, header=h_r)
                        df_p.columns = [re.sub(r'[^A-Z]', '', str(c).upper().replace('Ó','O')) for c in df_p.columns]
                        if 'FECHA' in df_p.columns and 'ODO' in df_p.columns:
                            df_p['_dt'] = pd.to_datetime(df_p['FECHA'], errors='coerce')
                            for _, r in df_p.dropna(subset=['_dt']).iterrows():
                                if start_date <= r['_dt'].date() <= end_date:
                                    all_ops.append({"Fecha": r['_dt'].normalize(), "Tipo Día": get_tipo_dia(r['_dt']), "N° Semana": r['_dt'].isocalendar()[1], "Odómetro [km]": parse_latam_number(r.get('ODO',0)), "Tren-Km [km]": parse_latam_number(r.get('TRENKM',0)), "UMR [%]": (parse_latam_number(r.get('TRENKM',0))/parse_latam_number(r.get('ODO',1))*100)})
                if 'ODO' in sn_up and 'KIL' in sn_up:
                    df_tr_raw = pd.read_excel(f, sheet_name=sn, header=None)
                    for i in range(len(df_tr_raw)-2):
                        for j in range(1, len(df_tr_raw.columns)):
                            v = pd.to_datetime(df_tr_raw.iloc[i,j], errors='coerce')
                            if pd.notna(v) and start_date <= v.date() <= end_date:
                                for k in range(i+3, min(i+40, len(df_tr_raw))):
                                    n = str(df_tr_raw.iloc[k,0]).strip().upper()
                                    if n.startswith(('M','XM')):
                                        all_tr.append({"Tren":n, "Fecha":v.normalize(), "Día":v.day, "Valor":parse_latam_number(df_tr_raw.iloc[k,j])})
                if 'SEAT' in sn_up:
                    df_s = pd.read_excel(f, sheet_name=sn, header=None)
                    for i in range(len(df_s)):
                        fs = pd.to_datetime(df_s.iloc[i,1], errors='coerce')
                        if pd.notna(fs) and start_date <= fs.date() <= end_date:
                            tot, tra = parse_latam_number(df_s.iloc[i,3]), parse_latam_number(df_s.iloc[i,5])
                            all_seat.append({"Fecha":fs.normalize(), "Total [kWh]":tot, "Tracción [kWh]":tra, "12 KV [kWh]":parse_latam_number(df_s.iloc[i,7]), "% Tracción":(tra/tot*100 if tot>0 else 0), "% 12 KV":(parse_latam_number(df_s.iloc[i,7])/tot*100 if tot>0 else 0)})
                if any(k in sn_up for k in ['PRMTE', 'MEDIDAS']):
                    df_prm = pd.read_excel(f, sheet_name=sn, header=None)
                    h_idx = next((i for i in range(len(df_prm)) if 'AÑO' in str(df_prm.iloc[i]).upper()), None)
                    if h_idx is not None:
                        df_pd = pd.read_excel(f, sheet_name=sn, header=h_idx)
                        df_pd['Timestamp'] = pd.to_datetime(df_pd[['AÑO', 'MES', 'DIA', 'HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'})) + pd.to_timedelta(df_pd['INICIO INTERVALO'].astype(int), unit='m')
                        cols_e = [c for c in df_pd.columns if 'Retiro_Energia_Activa (kWhD)' in str(c)]
                        for _, r in df_pd.iterrows():
                            ts, val_p = r['Timestamp'], sum([parse_latam_number(r[col]) for col in cols_e])
                            all_comp_full.append({"Fecha":ts.normalize(), "Hora":ts.hour, "Consumo Horario [kWh]":val_p, "Fuente":"PRMTE"})
                            if start_date <= ts.date() <= end_date: all_prmte_15.append({"Fecha":ts.normalize(), "Energía PRMTE [kWh]":val_p})
                if any(k in sn_up for k in ['FACTURA', 'CONSUMO']):
                    df_f = pd.read_excel(f, sheet_name=sn); df_f.columns = ['FechaHora', 'Valor']
                    df_f['Timestamp'] = pd.to_datetime(df_f['FechaHora'], errors='coerce')
                    for _, r in df_f.dropna(subset=['Timestamp']).iterrows():
                        ts, val_f = r['Timestamp'], abs(parse_latam_number(r['Valor']))
                        all_comp_full.append({"Fecha":ts.normalize(), "Hora":ts.hour, "Consumo Horario [kWh]":val_f, "Fuente":"Factura"})
                        if start_date <= ts.date() <= end_date: all_fact_h.append({"Fecha":ts.normalize(), "Consumo [kWh]":val_f})
        except: continue
    if f_v1:
        l1 = [procesar_thdr_avanzado(f)[0] for f in f_v1]; df_thdr_v1 = pd.concat(l1) if l1 else pd.DataFrame()
    if f_v2:
        l2 = [procesar_thdr_avanzado(f)[0] for f in f_v2]; df_thdr_v2 = pd.concat(l2) if l2 else pd.DataFrame()
    if all_ops:
        df_ops = pd.DataFrame(all_ops).drop_duplicates(subset=['Fecha']).sort_values("Fecha")
        if all_seat:
            df_seat = pd.DataFrame(all_seat).drop_duplicates(subset=['Fecha'])
            df_ops = pd.merge(df_ops, df_seat.rename(columns={"Total [kWh]":"E_Total", "Tracción [kWh]":"E_Tr", "12 KV [kWh]":"E_12"}), on="Fecha", how="left")
            df_ops['IDE (kWh/km)'] = df_ops['E_Tr'] / df_ops['Odómetro [km]']
    if all_tr: df_tr = pd.DataFrame(all_tr).sort_values(["Fecha", "Tren"])
    if all_fact_h: df_f_d = pd.DataFrame(all_fact_h).groupby("Fecha")["Consumo [kWh]"].sum().reset_index()
    if all_prmte_15: df_p_d = pd.DataFrame(all_prmte_15).groupby("Fecha")["Energía PRMTE [kWh]"].sum().reset_index()

# --- 7. DASHBOARD ---
tabs = st.tabs(["📊 Resumen", "📑 Operaciones", "📑 Trenes", "⚡ Energía", "⚖️ Comparación Energía hr", "📈 Regresión Nocturna", "🚨 Datos Atípicos", "📋 THDR"])

with tabs[0]:
    if not df_ops.empty:
        st.write("### Indicadores Globales")
        c1, c2, c3 = st.columns(3)
        c1.metric("Odómetro Total", f"{df_ops['Odómetro [km]'].sum():,.1f} km")
        c2.metric("Tren-Km Total", f"{df_ops['Tren-Km [km]'].sum():,.1f} km")
        c3.metric("IDE Medio", f"{df_ops['IDE (kWh/km)'].mean():.4f}")
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_trace(go.Bar(x=df_ops['Fecha'], y=df_ops['Odómetro [km]'], name="Km"), secondary_y=False)
        fig.add_trace(go.Scatter(x=df_ops['Fecha'], y=df_ops['IDE (kWh/km)'], name="IDE"), secondary_y=True)
        st.plotly_chart(fig, use_container_width=True)
    else: st.info("Sube archivos para ver el resumen.")

with tabs[1]:
    if not df_ops.empty: st.dataframe(df_ops.style.format({'IDE (kWh/km)': '{:.4f}'}))
    else: st.info("Sin datos de operaciones.")

with tabs[2]:
    if not df_tr.empty:
        piv = df_tr.pivot_table(index="Tren", columns=df_tr["Fecha"].dt.day, values="Valor", aggfunc='sum')
        st.write("#### Kilometraje Diario por Tren")
        st.dataframe(piv.style.format("{:,.1f}"))
    else: st.info("Sin datos de trenes.")

with tabs[3]:
    sub_e = st.tabs(["⚡ SEAT", "📈 PRMTE", "💰 Factura"])
    with sub_e[0]: st.dataframe(df_seat)
    with sub_e[1]: st.dataframe(df_p_d)
    with sub_e[2]: st.dataframe(df_f_d)

with tabs[4]:
    if all_comp_full:
        df_c = pd.DataFrame(all_comp_full).groupby(['Fecha','Hora','Fuente'])['Consumo Horario [kWh]'].sum().reset_index()
        df_c['Año'] = df_c['Fecha'].dt.year
        piv_c = df_c.pivot_table(index="Hora", columns="Año", values="Consumo Horario [kWh]", aggfunc='median')
        st.write("#### Mediana de Consumo Horario por Año")
        st.line_chart(piv_c); st.dataframe(piv_c)

with tabs[5]:
    if all_comp_full:
        df_reg = pd.DataFrame(all_comp_full)
        df_reg = df_reg[df_reg['Hora'] <= 5]
        h_sel = st.selectbox("Selecciona Hora para Regresión", range(6))
        df_h = df_reg[df_reg['Hora'] == h_sel].sort_values("Fecha")
        if len(df_h) > 2:
            x = np.arange(len(df_h)); y = df_h['Consumo Horario [kWh]'].values
            m, b = np.polyfit(x, y, 1)
            st.markdown(f"**Ecuación:** $Consumo = {m:.4f}x + {b:.2f}$")
            fig_r = go.Figure()
            fig_r.add_trace(go.Scatter(x=df_h['Fecha'], y=y, name="Real"))
            fig_r.add_trace(go.Scatter(x=df_h['Fecha'], y=m*x+b, name="Tendencia"))
            st.plotly_chart(fig_r)

with tabs[6]:
    if all_comp_full:
        df_at = pd.DataFrame(all_comp_full)
        q1, q3 = df_at['Consumo Horario [kWh]'].quantile(0.25), df_at['Consumo Horario [kWh]'].quantile(0.75)
        iqr = q3 - q1
        out = df_at[(df_at['Consumo Horario [kWh]'] > q3+1.5*iqr) | (df_at['Consumo Horario [kWh]'] < q1-1.5*iqr)]
        st.write(f"Anomalías detectadas: {len(out)}")
        st.dataframe(out)

with tabs[7]:
    st.write("#### Vía 1"); st.dataframe(df_thdr_v1)
    st.write("#### Vía 2"); st.dataframe(df_thdr_v2)

st.sidebar.download_button("📥 Reporte Completo", to_excel_consolidado(df_ops, df_tr, pd.DataFrame(), df_seat, df_p_d, pd.DataFrame(), pd.DataFrame(), df_f_d), "Reporte_SGE.xlsx")
