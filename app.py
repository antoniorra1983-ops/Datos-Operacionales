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

# --- 2. FUNCIONES DE PROCESAMIENTO Y EXPORTACIÓN (IDÉNTICAS A TU MASTER) ---
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
        if df_datos_semanales is not None and not df_datos_semanales.empty:
            df_datos_semanales.to_excel(writer, sheet_name='Datos_Semanales', index=False)
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

# --- 3. FUNCIONES PARA PROCESAR THDR (REPARADA PARA MOSTRAR TODO) ---
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

def procesar_thdr_avanzado(file, start_date=None, end_date=None):
    try:
        try: df_raw = pd.read_excel(file, header=None)
        except: df_raw = pd.read_excel(file, header=None, engine='xlrd')
        
        # Lógica de cabeceras dinámica para capturar TODAS las estaciones
        h0 = df_raw.iloc[0].ffill().fillna('').astype(str)
        h1 = df_raw.iloc[1].fillna('').astype(str)
        
        final_cols = []
        for i in range(len(h0)):
            base, sub = h0[i].strip(), h1[i].strip()
            if 'Hora' in sub: final_cols.append(f"{base}_{sub}")
            else: final_cols.append(base)
            
        df = df_raw.iloc[2:].copy(); df.columns = final_cols
        
        # Identificar columnas de estaciones para el cálculo de Tren-Km
        def detectar_extremos(row):
            times = []
            for col in df.columns:
                if 'Hora' in col:
                    val = convertir_a_minutos(row[col])
                    if val is not None: times.append((val, col.split('_')[0]))
            if not times: return pd.Series([None, "N/A", None, "N/A"])
            return pd.Series([times[0][0], times[0][1], times[-1][0], times[-1][1]])

        df[['H_Ini', 'Origen', 'H_Fin', 'Destino']] = df.apply(detectar_extremos, axis=1)
        
        # Buscar columnas técnicas
        c_serv = next((c for c in df.columns if any(k in c.lower() for k in ['servicio', 'n°'])), None)
        c_prog = next((c for c in df.columns if 'prog' in c.lower()), None)
        c_m2 = next((c for c in df.columns if any(k in c.lower() for k in ['motriz 2', 'm2'])), None)
        
        df['Servicio'] = pd.to_numeric(df[c_serv], errors='coerce').fillna(0).astype(int) if c_serv else 0
        df['Unidad'] = pd.to_numeric(df[c_m2], errors='coerce').fillna(0).apply(lambda x: 'M' if x > 0 else 'S')
        df['Min_Prog'] = df[c_prog].apply(convertir_a_minutos) if c_prog else 0
        df['Retraso'] = df['H_Ini'] - df['Min_Prog']
        
        def calc_km(r):
            o, d = str(r['Origen'])[:2].upper(), str(r['Destino'])[:2].upper()
            map_e = {"PU":"PU", "VA":"PU", "LI":"LI", "VI":"VM", "EL":"EB"}
            k = f"{map_e.get(o,o)}-{map_e.get(d,d)}"
            return DISTANCIAS.get(k, 43.13) * (2 if r['Unidad'] == 'M' else 1)
        
        df['Tren-Km'] = df.apply(calc_km, axis=1)
        
        # Fecha
        try:
            f_str = str(df_raw.iloc[0, 0]).split('.')[0].strip().zfill(6)
            df['Fecha_Op'] = pd.to_datetime(f"{f_str[0:2]}/{f_str[2:4]}/20{f_str[4:6]}", format='%d/%m/%Y')
        except: df['Fecha_Op'] = pd.NaT
        
        return df[df['Servicio'] > 0]
    except Exception as e:
        st.error(f"Error THDR {file.name}: {e}"); return pd.DataFrame()

# --- 4. INICIALIZACIÓN ---
df_ops = df_tr = df_tr_acum = df_seat = df_energy_master = df_p_d = df_f_d = df_thdr_v1 = df_thdr_v2 = pd.DataFrame()
all_comp_full = []; all_prmte_15 = []; all_fact_h = []

# --- 5. SIDEBAR (LOS 5 CARGADORES QUE PEDISTE) ---
with st.sidebar:
    st.header("📅 Filtro Global")
    date_range = st.date_input("Período", value=(date.today().replace(day=1), date.today()))
    start_date, end_date = (date_range[0], date_range[1]) if len(date_range)==2 else (date_range, date_range)
    st.divider()
    st.header("📂 Carga de Archivos")
    f_v1 = st.file_uploader("1. THDR Vía 1", type=["xls", "xlsx"], accept_multiple_files=True)
    f_v2 = st.file_uploader("2. THDR Vía 2", type=["xls", "xlsx"], accept_multiple_files=True)
    f_umr = st.file_uploader("3. UMR / Odómetros", type=["xlsx"], accept_multiple_files=True)
    f_seat_files = st.file_uploader("4. Energía SEAT", type=["xlsx"], accept_multiple_files=True)
    f_bill_files = st.file_uploader("5. Facturación y PRMTE", type=["xlsx"], accept_multiple_files=True)

# --- 6. PROCESAMIENTO GENERAL ---
if any([f_v1, f_v2, f_umr, f_seat_files, f_bill_files]):
    all_ops, all_tr, all_tr_acum, all_seat = [], [], [], []
    todos = (f_v1 or []) + (f_v2 or []) + (f_umr or []) + (f_seat_files or []) + (f_bill_files or [])

    for f in todos:
        try:
            xl = pd.ExcelFile(f)
            for sn in xl.sheet_names:
                sn_up = sn.upper()
                if any(k in sn_up for k in ['UMR', 'RESUMEN']):
                    df_u = pd.read_excel(f, sheet_name=sn, header=None)
                    h_idx = next((i for i in range(min(50, len(df_u))) if 'ODO' in str(df_u.iloc[i]).upper()), None)
                    if h_idx:
                        df_p = pd.read_excel(f, sheet_name=sn, header=h_idx)
                        df_p.columns = [re.sub(r'[^A-Z]', '', str(c).upper()) for c in df_p.columns]
                        df_p['_dt'] = pd.to_datetime(df_p.get('FECHA'), errors='coerce')
                        for _, r in df_p[(df_p['_dt'].dt.date >= start_date) & (df_p['_dt'].dt.date <= end_date)].iterrows():
                            all_ops.append({"Fecha": r['_dt'].normalize(), "Tipo Día": get_tipo_dia(r['_dt']), "Odómetro [km]": parse_latam_number(r.get('ODO')), "Tren-Km [km]": parse_latam_number(r.get('TRENKM'))})
                if 'ODO' in sn_up and 'KIL' in sn_up:
                    df_tr_raw = pd.read_excel(f, sheet_name=sn, header=None)
                    for i in range(len(df_tr_raw)-2):
                        for j in range(1, len(df_tr_raw.columns)):
                            v = pd.to_datetime(df_tr_raw.iloc[i,j], errors='coerce')
                            if pd.notna(v) and start_date <= v.date() <= end_date:
                                for k in range(i+3, min(i+40, len(df_tr_raw))):
                                    nt = str(df_tr_raw.iloc[k,0]).strip().upper()
                                    if nt.startswith(('M','XM')):
                                        d_pt = {"Tren": nt, "Fecha": v.normalize(), "Valor": parse_latam_number(df_tr_raw.iloc[k,j])}
                                        if any(k in str(df_tr_raw.iloc[i:i+3, 0]).upper() for k in ['ACUM', 'TOTAL']): all_tr_acum.append(d_pt)
                                        else: all_tr.append(d_pt)
                if 'SEAT' in sn_up:
                    df_s = pd.read_excel(f, header=None)
                    for i in range(len(df_s)):
                        dt = pd.to_datetime(df_s.iloc[i,1], errors='coerce')
                        if pd.notna(dt): all_seat.append({"Fecha": dt.normalize(), "E_Total": parse_latam_number(df_s.iloc[i,3]), "E_Tr": parse_latam_number(df_s.iloc[i,5]), "E_12": parse_latam_number(df_s.iloc[i,7])})
        except: continue

    if f_v1: df_thdr_v1 = pd.concat([procesar_thdr_avanzado(f, start_date, end_date) for f in f_v1], ignore_index=True)
    if f_v2: df_thdr_v2 = pd.concat([procesar_thdr_avanzado(f, start_date, end_date) for f in f_v2], ignore_index=True)

    if all_ops:
        df_ops = pd.DataFrame(all_ops).drop_duplicates(subset=['Fecha']).sort_values("Fecha")
        if all_seat:
            df_ops = pd.merge(df_ops, pd.DataFrame(all_seat), on="Fecha", how="left").fillna(0)
            df_ops['IDE (kWh/km)'] = df_ops.apply(lambda r: r['E_Tr']/r['Odómetro [km]'] if r['Odómetro [km]']>0 else 0, axis=1)

# --- 7. DASHBOARD (8 TABS EXACTAS) ---
tabs = st.tabs(["📊 Resumen", "📑 Operaciones", "📑 Trenes", "⚡ Energía", "⚖️ Comparación Energía hr", "📈 Regresión Nocturna", "🚨 Datos Atípicos", "📋 THDR"])

# PESTAÑA RESUMEN (MANTENIENDO TU LÓGICA DE SESSION STATE)
with tabs[0]:
    if not df_ops.empty:
        if 'filtros_compartidos' not in st.session_state: st.session_state.filtros_compartidos = {'anios': [], 'meses': []}
        c1, c2 = st.columns(2)
        f_ano = c1.multiselect("Año", sorted(df_ops['Fecha'].dt.year.unique()), default=sorted(df_ops['Fecha'].dt.year.unique()))
        f_mes = c2.multiselect("Mes", sorted(df_ops['Fecha'].dt.month.unique()), default=sorted(df_ops['Fecha'].dt.month.unique()))
        df_f = df_ops[df_ops['Fecha'].dt.year.isin(f_ano) & df_ops['Fecha'].dt.month.isin(f_mes)]
        if not df_f.empty:
            m1, m2, m3 = st.columns(3)
            m1.metric("Odómetro Total", f"{df_f['Odómetro [km]'].sum():,.1f} km")
            m2.metric("Tren-Km Total", f"{df_f['Tren-Km [km]'].sum():,.1f} km")
            m3.metric("IDE Promedio", f"{df_f['IDE (kWh/km)'].mean():.4f}")

# PESTAÑA TRENES
with tabs[2]:
    if not df_tr.empty:
        st.write("#### Kilometraje Diario [km]")
        st.dataframe(df_tr.pivot_table(index="Tren", columns=df_tr["Fecha"].dt.day, values="Valor", aggfunc='sum').fillna(0))

# PESTAÑA THDR (MEJORADA: MUESTRA TODAS LAS ESTACIONES DINÁMICAMENTE)
with tabs[7]:
    st.header("📋 Datos THDR - Vía 1 y Vía 2")
    
    def mostrar_tabla_thdr_dinamica(df, titulo, emoji):
        st.subheader(f"{emoji} {titulo}")
        if df.empty:
            st.info(f"No hay datos para {titulo}")
            return
        
        # Seleccionar todas las columnas originales que contienen "Hora" (Estaciones)
        cols_estaciones = [c for c in df.columns if 'Hora' in c]
        cols_tecnicas = ['Fecha_Op', 'Servicio', 'Unidad', 'Tren-Km', 'Retraso']
        
        df_display = df.copy()
        # Formatear solo las columnas de estaciones que tengan datos
        for col in cols_estaciones:
            df_display[col] = df_display[col].apply(lambda x: format_hms(x) if pd.notna(x) else "")
            
        final_cols = [c for c in cols_tecnicas if c in df_display.columns] + cols_estaciones
        st.dataframe(df_display[final_cols], use_container_width=True)

    mostrar_tabla_thdr_dinamica(df_thdr_v1, "Vía 1 (Puerto -> Limache)", "🟢")
    mostrar_tabla_thdr_dinamica(df_thdr_v2, "Vía 2 (Limache -> Puerto)", "🔵")

st.sidebar.download_button("📥 Reporte Excel Completo", to_excel_consolidado(df_ops, pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()), "Reporte_SGE_EFE.xlsx")
