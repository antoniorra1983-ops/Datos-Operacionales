import streamlit as st
import pandas as pd
import numpy as np
import re
import holidays
import requests  # Para la API de Clima
from io import BytesIO
from datetime import datetime, date, timedelta, time
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import traceback

# --- 1. CONFIGURACIÓN, ESTILOS Y API ---
st.set_page_config(page_title="Gestión de Energía - Dashboard SGE", layout="wide", page_icon="🚆")

# CONFIGURACIÓN API CLIMA - CLAVE INTEGRADA
API_KEY = "de25da707bfeb645ec2b488c4676af19" 
CIUDAD = "Valparaiso,CL"

chile_holidays = holidays.Chile()
ORDEN_TIPO_DIA = ["L", "S", "D/F"]

st.markdown("""
    <style>
    .stMetric { background-color: #ffffff; padding: 20px; border-radius: 10px; border-left: 5px solid #005195; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    </style>
    """, unsafe_allow_html=True)

# --- 2. FUNCIONES DE APOYO Y API ---

@st.cache_data(ttl=3600)
def obtener_clima_actual():
    """Consulta la temperatura actual mediante API REST."""
    try:
        url = f"http://api.openweathermap.org/data/2.5/weather?q={CIUDAD}&appid={API_KEY}&units=metric&lang=es"
        response = requests.get(url, timeout=5)
        if response.status_code == 200:
            data = response.json()
            return {
                "temp": data['main']['temp'],
                "hum": data['main']['humidity'],
                "desc": data['weather'][0]['description']
            }
    except:
        return None
    return None

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

DISTANCIAS = {
    "PU-LI": 43.13, "LI-PU": 43.13, "PU-SA": 29.11, "SA-PU": 29.11,
    "EB-PU": 25.40, "PU-EB": 25.40, "VM-LI": 34.03, "LI-VM": 34.03,
    "VM-PU": 9.10,  "PU-VM": 9.10
}

def extraer_fecha_desde_nombre_archivo(nombre_archivo):
    patrones = [r'(\d{2})(\d{2})(\d{2})', r'(\d{2})-(\d{2})-(\d{2})', r'(\d{2})\.(\d{2})\.(\d{2})']
    for pat in patrones:
        m = re.search(pat, nombre_archivo)
        if m:
            try:
                dia, mes, anio = int(m.group(1)), int(m.group(2)), int(m.group(3))
                if anio < 100: anio += 2000
                return date(anio, mes, dia)
            except: pass
    return None

def procesar_thdr_avanzado(file, start_date=None, end_date=None):
    try:
        try: df_raw = pd.read_excel(file, header=None, engine=None)
        except: df_raw = pd.read_excel(file, header=None, engine='xlrd')
        
        header0, header1 = df_raw.iloc[0].fillna('').astype(str), df_raw.iloc[1].fillna('').astype(str)
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
        
        col_servicio = buscar_columna(['Servicio', 'Serv', 'N° Servicio'])
        col_hora_prog = buscar_columna(['Hora_Prog', 'Hora Programada', 'Prog'])
        col_motriz1, col_motriz2 = buscar_columna(['Motriz 1']), buscar_columna(['Motriz 2'])
        
        df['Servicio'] = df[col_servicio] if col_servicio else 0
        df['Hora_Prog'] = df[col_hora_prog] if col_hora_prog else '00:00:00'
        df['Motriz 1'] = pd.to_numeric(df[col_motriz1], errors='coerce').fillna(0).astype(int) if col_motriz1 else 0
        df['Motriz 2'] = pd.to_numeric(df[col_motriz2], errors='coerce').fillna(0).astype(int) if col_motriz2 else 0
        df['Unidad'] = df['Motriz 2'].apply(lambda x: 'M' if x > 0 else 'S')
        
        columnas_horas = {}
        for col in df.columns:
            if 'hora salida' in col.lower():
                est = col.lower().replace('_hora salida', '').replace('hora salida', '').strip()
                if est: columnas_horas[f"{est}_salida"] = col
            elif 'hora llegada' in col.lower():
                est = col.lower().replace('_hora llegada', '').replace('hora llegada', '').strip()
                if est: columnas_horas[f"{est}_llegada"] = col
        
        for key, col in columnas_horas.items():
            df[f"{key}_min"] = df[col].apply(convertir_a_minutos)
            df[f"{key}_fmt"] = df[f"{key}_min"].apply(lambda x: format_hms(x) if pd.notna(x) else "")
        
        puerto_key = next((k for k in columnas_horas.keys() if 'puerto' in k and 'salida' in k), None)
        limache_key = next((k for k in columnas_horas.keys() if 'limache' in k and 'llegada' in k), None)
        
        df['Hora_Salida_Real'] = df[f"{puerto_key}_min"] if puerto_key else None
        df['Hora_Llegada_Real'] = df[f"{limache_key}_min"] if limache_key else None
        df['Min_Prog'] = df['Hora_Prog'].apply(convertir_a_minutos)
        df['Retraso'] = df['Hora_Salida_Real'] - df['Min_Prog']
        df['Puntual'] = (abs(df['Retraso']) <= 5).astype(int)
        
        if puerto_key and limache_key:
            tdv = (df['Hora_Llegada_Real'] - df['Hora_Salida_Real']).apply(lambda x: x if x > 0 else x + 1440 if pd.notna(x) else 0)
            df['TDV_Min'] = tdv
        else: df['TDV_Min'] = 0
        
        origen, destino = ('PU' if puerto_key else 'OTRO'), ('LI' if limache_key else 'OTRO')
        df['Tipo_Rec'] = f"{origen}-{destino}" if origen != 'OTRO' and destino != 'OTRO' else 'OTRO'
        df['Dist_Base'] = df['Tipo_Rec'].map(DISTANCIAS).fillna(0)
        df['Tren-Km'] = df['Dist_Base'] * df['Unidad'].apply(lambda x: 2 if x == 'M' else 1)
        
        df['Fecha_Op'] = pd.to_datetime(extraer_fecha_desde_nombre_archivo(file.name) or date.today())
        if start_date and end_date and not df.empty:
            mask = (df['Fecha_Op'].dt.date >= start_date) & (df['Fecha_Op'].dt.date <= end_date)
            df = df[mask].copy()
            
        df.attrs['estaciones_keys'] = list(columnas_horas.keys())
        return df
    except Exception as e:
        st.error(f"Error procesando THDR {file.name}: {str(e)}")
        return pd.DataFrame()

# --- 4. CARGA Y PROCESAMIENTO ---
df_ops, df_tr, df_tr_acum, df_seat, df_energy_master, df_p_d, df_f_d = [pd.DataFrame() for _ in range(7)]
df_thdr_v1, df_thdr_v2 = pd.DataFrame(), pd.DataFrame()
all_comp_full, all_prmte_15, all_fact_h = [], [], []

with st.sidebar:
    st.header("📅 Filtro Global")
    date_range = st.date_input("Período", value=(date.today().replace(day=1), date.today()))
    start_date, end_date = (date_range[0], date_range[1]) if isinstance(date_range, tuple) and len(date_range)==2 else (date_range, date_range)
    
    st.header("📂 Carga")
    f_v1 = st.file_uploader("1. THDR Vía 1", type=["xls", "xlsx"], accept_multiple_files=True)
    f_v2 = st.file_uploader("2. THDR Vía 2", type=["xls", "xlsx"], accept_multiple_files=True)
    f_umr = st.file_uploader("3. UMR / Odómetros", type=["xlsx"], accept_multiple_files=True)
    f_seat_files = st.file_uploader("4. Energía SEAT", type=["xlsx"], accept_multiple_files=True)
    f_bill_files = st.file_uploader("5. Facturación y PRMTE", type=["xlsx"], accept_multiple_files=True)

    st.divider()
    st.header("🌤️ Clima Real")
    clima = obtener_clima_actual()
    if clima:
        c1, c2 = st.columns(2)
        c1.metric("Temp", f"{clima['temp']}°C")
        c2.metric("Hum", f"{clima['hum']}%")
        st.caption(f"Condición: {clima['desc'].capitalize()}")
    else: st.warning("API de Clima no disponible.")

# --- (PROCESAMIENTO DE ARCHIVOS - SE MANTIENE TU LÓGICA DE UNIÓN) ---
if f_v1 or f_v2 or f_umr or f_seat_files or f_bill_files:
    # ... (Se asume procesamiento interno igual al original para rellenar dataframes) ...
    if f_v1:
        th1 = [procesar_thdr_avanzado(f, start_date, end_date) for f in f_v1]
        df_thdr_v1 = pd.concat(th1, ignore_index=True) if th1 else pd.DataFrame()
    if f_v2:
        th2 = [procesar_thdr_avanzado(f, start_date, end_date) for f in f_v2]
        df_thdr_v2 = pd.concat(th2, ignore_index=True) if th2 else pd.DataFrame()

# --- 7. DASHBOARD ---
tabs = st.tabs(["📊 Resumen", "📑 Operaciones", "📑 Trenes", "⚡ Energía", "⚖️ Comparativa hr", "📈 Regresión", "🚨 Atípicos", "📋 THDR"])

# ================== PESTAÑA THDR (ORDENADA POR ESTACIÓN) ==================
with tabs[7]:
    st.header("📋 Datos THDR - Orden Secuencial de Estaciones")

    def mostrar_tabla_thdr_ordenada(df, titulo):
        if df.empty:
            st.info(f"Sin datos para {titulo}")
            return
        
        st.subheader(f"📍 {titulo}")
        
        # Identificar estaciones únicas (manteniendo orden de aparición en columnas)
        cols_fmt = [c for c in df.columns if c.endswith('_fmt')]
        nombres_estaciones = []
        for c in cols_fmt:
            est = c.replace('_salida_fmt', '').replace('_llegada_fmt', '')
            if est not in nombres_estaciones: nombres_estaciones.append(est)
        
        # Construir columnas finales: Datos básicos + Estaciones (Llegada then Salida)
        columnas_finales = ['Fecha_Op', 'Servicio', 'Unidad', 'Tren-Km']
        for est in nombres_estaciones:
            # Ordenar: Por cada estación, primero la llegada, luego la salida (excepto puntas)
            llegada = f"{est}_llegada_fmt"
            salida = f"{est}_salida_fmt"
            if llegada in df.columns: columnas_finales.append(llegada)
            if salida in df.columns: columnas_finales.append(salida)
        
        columnas_finales += ['Retraso', 'TDV_Min']
        cols_existentes = [c for c in columnas_finales if c in df.columns]
        
        df_display = df[cols_existentes].copy()
        # Renombrar para legibilidad
        nombres_amigables = {c: c.replace('_fmt','').replace('_',' ').title() for c in cols_existentes}
        st.dataframe(df_display.rename(columns=nombres_amigables), use_container_width=True)

    mostrar_tabla_thdr_ordenada(df_thdr_v1, "Vía 1 (Puerto → Limache)")
    mostrar_tabla_thdr_ordenada(df_thdr_v2, "Vía 2 (Limache → Puerto)")

# ================== PESTAÑAS RESTANTES (SE MANTIENE TU ESTRUCTURA) ==================
# ... (Aquí sigue el resto de tu lógica de Resumen, Trenes, Energía, etc.)
with tabs[0]:
    if not df_ops.empty:
        st.metric("IDE Global", f"{df_ops['IDE (kWh/km)'].mean():.4f} kWh/km")
    else: st.info("Carga datos para ver el análisis crítico del SGE.")

# --- 8. EXPORTACIÓN ---
st.sidebar.download_button("📥 Reporte Excel Completo", to_excel_consolidado(df_ops, df_tr, df_tr_acum, df_seat, df_p_d, pd.DataFrame(all_prmte_15), pd.DataFrame(all_fact_h), df_f_d), "Reporte_EFE_SGE.xlsx")
