import streamlit as st
import pandas as pd
import numpy as np
import re
import holidays
import requests
from io import BytesIO
from datetime import datetime, date, timedelta, time
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import traceback

# --- 1. CONFIGURACIÓN Y API CORREDOR TÉRMICO ---
st.set_page_config(page_title="Gestión de Energía - Dashboard SGE", layout="wide", page_icon="🚆")

API_KEY = "de25da707bfeb645ec2b488c4676af19" 

# Ciudades que componen el tramo operativo
CIUDADES_CORREDOR = {
    "Valparaíso (Puerto)": "Valparaiso,CL",
    "Viña del Mar": "Vina del Mar,CL",
    "Quilpué": "Quilpue,CL",
    "Villa Alemana": "Villa Alemana,CL",
    "Limache": "Limache,CL"
}

chile_holidays = holidays.Chile()
ORDEN_TIPO_DIA = ["L", "S", "D/F"]

st.markdown("""
    <style>
    .stMetric { background-color: #ffffff; padding: 20px; border-radius: 10px; border-left: 5px solid #005195; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    </style>
    """, unsafe_allow_html=True)

# --- 2. FUNCIONES DE APOYO Y CLIMA ---

@st.cache_data(ttl=3600)
def obtener_clima_corredor():
    resultados = {}
    for nombre, query in CIUDADES_CORREDOR.items():
        try:
            url = f"http://api.openweathermap.org/data/2.5/weather?q={query}&appid={API_KEY}&units=metric&lang=es"
            resp = requests.get(url, timeout=5)
            if resp.status_code == 200:
                d = resp.json()
                resultados[nombre] = {
                    "temp": d['main']['temp'],
                    "hum": d['main']['humidity'],
                    "desc": d['weather'][0]['description']
                }
        except:
            resultados[nombre] = None
    return resultados

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

# --- 3. FUNCIONES THDR (FIX AMBIGUOUS SERIES) ---
def convertir_a_minutos(val):
    if pd.isna(val) or str(val).strip() == "": return None
    try:
        if isinstance(val, (datetime, time)): return val.hour * 60 + val.minute + (val.second / 60.0)
        if isinstance(val, str):
            val = val.strip()
            m_ss = re.search(r'(\d{1,2}):(\d{2}):(\d{2})', val)
            if m_ss: return int(m_ss.group(1)) * 60 + int(m_ss.group(2)) + (int(m_ss.group(3)) / 60.0)
            m_mm = re.search(r'(\d{1,2}):(\d{2})', val)
            if m_mm: return int(m_mm.group(1)) * 60 + int(m_mm.group(2))
        return None
    except: return None

def format_hms(minutos_float, con_signo=False):
    if pd.isna(minutos_float) or minutos_float == 0: return "00:00:00"
    signo = ("+" if minutos_float > 0 else "-" if minutos_float < 0 else "") if con_signo else ""
    total_segundos = int(round(abs(minutos_float) * 60))
    h, r = divmod(total_segundos, 3600)
    m, s = divmod(r, 60)
    return f"{signo}{h:02d}:{m:02d}:{s:02d}"

def procesar_thdr_avanzado(file, start_date=None, end_date=None):
    try:
        try: df_raw = pd.read_excel(file, header=None, engine=None)
        except: df_raw = pd.read_excel(file, header=None, engine='xlrd')
        
        header0 = df_raw.iloc[0].fillna('').astype(str)
        header1 = df_raw.iloc[1].fillna('').astype(str)
        column_names = []
        for i in range(len(header0)):
            base, sub = header0[i].strip(), header1[i].strip()
            if sub in ['Hora Llegada', 'Hora Salida']: column_names.append(f"{base}_{sub}")
            else: column_names.append(base)
        df = df_raw.iloc[2:].copy()
        df.columns = column_names
        
        def buscar_columna(nombres_posibles):
            for col in df.columns:
                for posible in nombres_posibles:
                    if posible.lower() in col.lower(): return col
            return None
        
        c_serv = buscar_columna(['Servicio', 'Serv', 'N° Servicio'])
        c_prog = buscar_columna(['Hora_Prog', 'Hora Programada', 'Prog'])
        c_m1, c_m2 = buscar_columna(['Motriz 1']), buscar_columna(['Motriz 2'])
        
        df['Servicio'] = df[c_serv] if c_serv else 0
        df['Hora_Prog'] = df[c_prog] if c_prog else '00:00:00'
        df['Motriz 1'] = pd.to_numeric(df[c_m1], errors='coerce').fillna(0).astype(int) if c_m1 else 0
        df['Motriz 2'] = pd.to_numeric(df[c_m2], errors='coerce').fillna(0).astype(int) if c_m2 else 0
        df['Unidad'] = df['Motriz 2'].apply(lambda x: 'M' if x > 0 else 'S')
        
        columnas_horas = {}
        for col in df.columns:
            clow = col.lower()
            if 'hora salida' in clow:
                est = col.split('_')[0].split('Hora Salida')[0].strip()
                if not est: est = "Estacion"
                columnas_horas[f"{est}_salida"] = col
            elif 'hora llegada' in clow:
                est = col.split('_')[0].split('Hora Llegada')[0].strip()
                if not est: est = "Estacion"
                columnas_horas[f"{est}_llegada"] = col
        
        for key, col in columnas_horas.items():
            df[f"{key}_min"] = df[col].apply(convertir_a_minutos)
            df[f"{key}_fmt"] = df[f"{key}_min"].apply(lambda x: format_hms(x) if pd.notna(x) else "")
        
        p_key = next((k for k in columnas_horas.keys() if 'puerto' in k.lower() and 'salida' in k), None)
        l_key = next((k for k in columnas_horas.keys() if 'limache' in k.lower() and 'llegada' in k), None)
        
        df['Hora_Salida_Real'] = df[f"{p_key}_min"] if p_key else None
        df['Hora_Llegada_Real'] = df[f"{l_key}_min"] if l_key else None
        df['Min_Prog'] = df['Hora_Prog'].apply(convertir_a_minutos)
        df['Retraso'] = df['Hora_Salida_Real'] - df['Min_Prog']
        df['Puntual'] = (abs(df['Retraso']) <= 5).astype(int)
        
        if p_key and l_key:
            tdv = (df['Hora_Llegada_Real'] - df['Hora_Salida_Real']).apply(lambda x: x if x > 0 else x + 1440 if pd.notna(x) else 0)
            df['TDV_Min'] = tdv
        else: df['TDV_Min'] = 0
        
        fch_f = extraer_fecha_desde_nombre_archivo(file.name)
        df['Fecha_Op'] = pd.to_datetime(fch_f if fch_f is not None else date.today())
        
        if start_date is not None and end_date is not None:
            mask = (df['Fecha_Op'].dt.date >= start_date) & (df['Fecha_Op'].dt.date <= end_date)
            df = df[mask].copy()
            
        return df
    except Exception as e:
        st.error(f"Error procesando THDR {file.name}: {str(e)}")
        return pd.DataFrame()

# --- 4. INICIALIZACIÓN ---
df_ops, df_tr, df_tr_acum, df_seat, df_energy_master, df_p_d, df_f_d = [pd.DataFrame() for _ in range(7)]
df_thdr_v1, df_thdr_v2 = pd.DataFrame(), pd.DataFrame()
all_ops, all_tr, all_tr_acum, all_seat, all_prmte_15, all_fact_h, all_comp_full = [], [], [], [], [], [], []

# --- 5. SIDEBAR Y CLIMA CORREDOR ---
with st.sidebar:
    st.header("📅 Período de Análisis")
    date_range = st.date_input("Rango", value=(date.today().replace(day=1), date.today()))
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = date_range[0], date_range[1]
    else: start_date, end_date = date_range, date_range

    st.divider()
    st.header("📂 Carga de Archivos")
    f_v1 = st.file_uploader("1. THDR Vía 1", type=["xls", "xlsx"], accept_multiple_files=True)
    f_v2 = st.file_uploader("2. THDR Vía 2", type=["xls", "xlsx"], accept_multiple_files=True)
    f_umr = st.file_uploader("3. UMR / Odómetros", type=["xlsx"], accept_multiple_files=True)
    f_seat_files = st.file_uploader("4. Energía SEAT", type=["xlsx"], accept_multiple_files=True)
    f_bill_files = st.file_uploader("5. Facturación y PRMTE", type=["xlsx"], accept_multiple_files=True)
    
    st.divider()
    st.header("🌤️ Corredor Puerto-Limache")
    climas = obtener_clima_corredor()
    if climas:
        for loc, info in climas.items():
            if info: st.write(f"**{loc}:** {info['temp']}°C | {info['desc'].capitalize()}")
            else: st.write(f"**{loc}:** Sin datos")

# --- 6. PROCESAMIENTO GENERAL ---
if f_v1 or f_v2 or f_umr or f_seat_files or f_bill_files:
    todos = (f_v1 or []) + (f_v2 or []) + (f_umr or []) + (f_seat_files or []) + (f_bill_files or [])
    for f in todos:
        try:
            xl = pd.ExcelFile(f)
            for sn in xl.sheet_names:
                sn_up = sn.upper()
                if any(k in sn_up for k in ['UMR', 'RESUMEN']):
                    df_raw = pd.read_excel(f, sheet_name=sn, header=None)
                    # Arreglo de Truth Value Error en detección de fila de cabecera
                    h_r = None
                    for i in range(min(100, len(df_raw))):
                        row_str = str(df_raw.iloc[i]).upper()
                        if 'ODO' in row_str or 'FECHA' in row_str:
                            h_r = i
                            break
                    if h_r is not None:
                        df_p = pd.read_excel(f, sheet_name=sn, header=h_r)
                        df_p.columns = [re.sub(r'[^A-Z]', '', str(c).upper().replace('Ó','O')) for c in df_p.columns]
                        idx_f = next((c for c in df_p.columns if 'FECHA' in c), None)
                        idx_o = next((c for c in df_p.columns if 'ODO' in c and 'ACUM' not in c), None)
                        idx_t = next((c for c in df_p.columns if 'TREN' in c and 'KM' in c), None)
                        if idx_f and idx_o:
                            df_p['_dt'] = pd.to_datetime(df_p[idx_f], errors='coerce')
                            mask = (df_p['_dt'].dt.date >= start_date) & (df_p['_dt'].dt.date <= end_date)
                            for _, r in df_p[mask].dropna(subset=['_dt']).iterrows():
                                all_ops.append({
                                    "Fecha": r['_dt'].normalize(), "Tipo Día": get_tipo_dia(r['_dt']), 
                                    "N° Semana": r['_dt'].isocalendar()[1], "Odómetro [km]": parse_latam_number(r[idx_o]), 
                                    "Tren-Km [km]": parse_latam_number(r[idx_t]), 
                                    "UMR [%]": (parse_latam_number(r[idx_t])/parse_latam_number(r[idx_o])*100 if parse_latam_number(r[idx_o])>0 else 0)
                                })
                
                if 'ODO' in sn_up and 'KIL' in sn_up:
                    df_tr_raw = pd.read_excel(f, sheet_name=sn, header=None)
                    for i in range(len(df_tr_raw)-2):
                        for j in range(1, len(df_tr_raw.columns)):
                            val = pd.to_datetime(df_tr_raw.iloc[i, j], errors='coerce')
                            if pd.notna(val) and start_date <= val.date() <= end_date:
                                # Lógica para capturar km por tren
                                for k in range(i+3, min(i+40, len(df_tr_raw))):
                                    n_tr = str(df_tr_raw.iloc[k, 0]).strip().upper()
                                    if re.match(r'^(M|XM)', n_tr):
                                        val_km = parse_latam_number(df_tr_raw.iloc[k, j])
                                        all_tr.append({"Tren": n_tr, "Fecha": val.normalize(), "Valor": val_km})

                if 'SEAT' in sn_up and 'SER' in sn_up:
                    df_s = pd.read_excel(f, sheet_name=sn, header=None)
                    for i in range(len(df_s)):
                        fs = pd.to_datetime(df_s.iloc[i, 1], errors='coerce')
                        if pd.notna(fs) and start_date <= fs.date() <= end_date:
                            all_seat.append({
                                "Fecha": fs.normalize(), "Total [kWh]": parse_latam_number(df_s.iloc[i, 3]), 
                                "Tracción [kWh]": parse_latam_number(df_s.iloc[i, 5]), "12 KV [kWh]": parse_latam_number(df_s.iloc[i, 7])
                            })

                if any(k in sn_up for k in ['PRMTE', 'MEDIDAS']):
                    df_prm = pd.read_excel(f, sheet_name=sn, header=None)
                    h_idx = None
                    for i in range(len(df_prm)):
                        if 'AÑO' in str(df_prm.iloc[i]).upper():
                            h_idx = i
                            break
                    if h_idx is not None:
                        df_pd = pd.read_excel(f, sheet_name=sn, header=h_idx)
                        df_pd['Timestamp'] = pd.to_datetime(df_pd[['AÑO', 'MES', 'DIA', 'HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'})) + pd.to_timedelta(df_pd['INICIO INTERVALO'].astype(int), unit='m')
                        cols_e = [c for c in df_pd.columns if 'Retiro_Energia_Activa (kWhD)' in str(c)]
                        for _, r in df_pd.iterrows():
                            ts = r['Timestamp']
                            val_p = sum([parse_latam_number(r[col]) for col in cols_e])
                            all_comp_full.append({"Fecha": ts.normalize(), "Hora": ts.hour, "Consumo Horario [kWh]": val_p, "Fuente": "PRMTE"})
                            if start_date <= ts.date() <= end_date:
                                all_prmte_15.append({"Fecha y Hora": ts, "Fecha": ts.normalize(), "Energía PRMTE [kWh]": val_p})

                if any(k in sn_up for k in ['FACTURA', 'CONSUMO']):
                    df_f = pd.read_excel(f, sheet_name=sn)
                    df_f.columns = ['FechaHora', 'Valor']
                    df_f['Timestamp'] = pd.to_datetime(df_f['FechaHora'], errors='coerce')
                    for _, r in df_f.dropna(subset=['Timestamp']).iterrows():
                        ts = r['Timestamp']
                        val_f = abs(parse_latam_number(r['Valor']))
                        all_comp_full.append({"Fecha": ts.normalize(), "Hora": ts.hour, "Consumo Horario [kWh]": val_f, "Fuente": "Factura"})
                        if start_date <= ts.date() <= end_date:
                            all_fact_h.append({"Fecha y Hora": ts, "Fecha": ts.normalize(), "Consumo Horario [kWh]": val_f})
        except: continue

    if f_v1:
        th1 = [procesar_thdr_avanzado(f, start_date, end_date) for f in f_v1]
        df_thdr_v1 = pd.concat(th1, ignore_index=True) if th1 else pd.DataFrame()
    if f_v2:
        th2 = [procesar_thdr_avanzado(f, start_date, end_date) for f in f_v2]
        df_thdr_v2 = pd.concat(th2, ignore_index=True) if th2 else pd.DataFrame()

    if all_ops:
        df_ops = pd.DataFrame(all_ops).drop_duplicates(subset=['Fecha']).sort_values("Fecha")
        if all_seat:
            df_seat = pd.DataFrame(all_seat).drop_duplicates(subset=['Fecha']).sort_values("Fecha")
            df_ops = pd.merge(df_ops, df_seat, on="Fecha", how="left")
            df_ops['IDE (kWh/km)'] = df_ops.apply(lambda r: r['Tracción [kWh]'] / r['Odómetro [km]'] if r['Odómetro [km]'] > 0 else 0, axis=1)

# --- 7. DASHBOARD TABS ---
tabs = st.tabs(["📊 Resumen", "📑 Operaciones", "📑 Trenes", "⚡ Energía", "⚖️ Comparativa hr", "📈 Regresión", "🚨 Atípicos", "📋 THDR"])

with tabs[0]:
    if not df_ops.empty:
        col_m1, col_m2, col_m3 = st.columns(3)
        col_m1.metric("Odómetro Total", f"{df_ops['Odómetro [km]'].sum():,.1f} km")
        col_m2.metric("Tren-Km Total", f"{df_ops['Tren-Km [km]'].sum():,.1f} km")
        col_m3.metric("IDE Promedio", f"{df_ops['IDE (kWh/km)'].mean():.4f} kWh/km")
        if climas:
            st.write("#### 🚉 Perfil Térmico Corredor")
            c_cols = st.columns(len(climas))
            for i, (loc, info) in enumerate(climas.items()):
                if info: c_cols[i].metric(loc, f"{info['temp']}°C", info['desc'])
    else: st.info("Carga datos para visualizar.")

with tabs[1]:
    if not df_ops.empty:
        st.dataframe(df_ops.style.format({'Odómetro [km]': "{:,.1f}", 'IDE (kWh/km)': "{:.4f}"}))

with tabs[2]:
    if not df_tr.empty:
        piv = df_tr.pivot_table(index="Tren", columns=df_tr["Fecha"].dt.day, values="Valor", aggfunc='sum').fillna(0)
        st.dataframe(piv.style.format("{:,.1f}"))

with tabs[5]:
    if all_comp_full:
        df_reg = pd.DataFrame(all_comp_full).query("Hora <= 5")
        if len(df_reg) > 2:
            x, y = np.arange(len(df_reg)), df_reg['Consumo Horario [kWh]'].values
            m, n = np.polyfit(x, y, 1)
            st.markdown(f"**Línea de Base:** $kWh = {m:.4f}x + {n:.2f}$")
            fig_reg = go.Figure()
            fig_reg.add_trace(go.Scatter(x=x, y=y, mode='markers', name='Real'))
            fig_reg.add_trace(go.Scatter(x=x, y=m*x+n, mode='lines', name='Tendencia'))
            st.plotly_chart(fig_reg)

with tabs[7]:
    st.header("📋 Datos THDR - Secuencia de Estaciones")
    def mostrar_thdr_ordenada(df, titulo):
        if df.empty: return st.info(f"Sin datos para {titulo}")
        st.subheader(f"📍 {titulo}")
        cols_fmt = [c for c in df.columns if c.endswith('_fmt')]
        estaciones = []
        for c in cols_fmt:
            e = c.replace('_salida_fmt', '').replace('_llegada_fmt', '')
            if e not in estaciones: estaciones.append(e)
        
        f_cols = ['Fecha_Op', 'Servicio', 'Unidad']
        for e in estaciones:
            l, s = f"{e}_llegada_fmt", f"{e}_salida_fmt"
            if l in df.columns: f_cols.append(l)
            if s in df.columns: f_cols.append(s)
        
        existentes = [c for c in f_cols if c in df.columns]
        df_d = df[existentes].copy()
        st.dataframe(df_d.rename(columns={c: c.replace('_fmt','').replace('_',' ').title() for c in existentes}), use_container_width=True)

    mostrar_thdr_ordenada(df_thdr_v1, "Vía 1 (Puerto → Limache)")
    mostrar_thdr_ordenada(df_thdr_v2, "Vía 2 (Limache → Puerto)")

# --- 8. EXPORTACIÓN ---
st.sidebar.download_button("📥 Excel Completo", to_excel_consolidado(df_ops, df_tr, df_tr_acum, df_seat, df_p_d, pd.DataFrame(all_prmte_15), pd.DataFrame(all_fact_h), df_f_d), "Reporte_EFE_SGE.xlsx")
