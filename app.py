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

# --- 3. FUNCIONES PARA PROCESAR THDR ---
def convertir_a_minutos(val):
    if pd.isna(val) or str(val).strip() == "":
        return None
    try:
        if isinstance(val, (datetime, time)):
            return val.hour * 60 + val.minute + (val.second / 60.0)
        if isinstance(val, str):
            val = val.strip()
            m_ss = re.search(r'(\d{1,2}):(\d{2}):(\d{2})', val)
            if m_ss:
                return int(m_ss.group(1)) * 60 + int(m_ss.group(2)) + (int(m_ss.group(3)) / 60.0)
            m_mm = re.search(r'(\d{1,2}):(\d{2})', val)
            if m_mm:
                return int(m_mm.group(1)) * 60 + int(m_mm.group(2))
        return None
    except:
        return None

def format_hms(minutos_float, con_signo=False):
    if pd.isna(minutos_float) or minutos_float == 0:
        return "00:00:00"
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
    patrones = [
        r'(\d{2})(\d{2})(\d{2})',
        r'(\d{2})-(\d{2})-(\d{2})',
        r'(\d{2})\.(\d{2})\.(\d{2})'
    ]
    for pat in patrones:
        m = re.search(pat, nombre_archivo)
        if m:
            try:
                dia, mes, anio = int(m.group(1)), int(m.group(2)), int(m.group(3))
                if anio < 100:
                    anio += 2000
                return date(anio, mes, dia)
            except:
                pass
    return None

def procesar_thdr_avanzado(file, start_date=None, end_date=None):
    try:
        try:
            df_raw = pd.read_excel(file, header=None, engine=None)
        except Exception:
            df_raw = pd.read_excel(file, header=None, engine='xlrd')
        
        header0 = df_raw.iloc[0].fillna('').astype(str)
        header1 = df_raw.iloc[1].fillna('').astype(str)
        column_names = []
        for i in range(len(header0)):
            base = header0.iloc[i].strip()
            sub = header1.iloc[i].strip()
            if sub in ['Hora Llegada', 'Hora Salida']:
                column_names.append(f"{base}_{sub}")
            else:
                column_names.append(base)
        df = df_raw.iloc[2:].copy()
        df.columns = column_names
        
        def buscar_columna(nombres_posibles):
            for col in df.columns:
                for posible in nombres_posibles:
                    if posible.lower() in str(col).lower():
                        return col
            return None
        
        col_servicio = buscar_columna(['Servicio', 'Serv', 'N° Servicio'])
        col_hora_prog = buscar_columna(['Hora_Prog', 'Hora Programada', 'Hora Prog', 'Prog'])
        col_m1 = buscar_columna(['Motriz 1', 'Motriz1', 'M1', 'Motor 1'])
        col_m2 = buscar_columna(['Motriz 2', 'Motriz2', 'M2', 'Motor 2'])
        
        df['Servicio'] = df[col_servicio] if col_servicio is not None else 0
        df['Hora_Prog'] = df[col_hora_prog] if col_hora_prog is not None else '00:00:00'
        df['Motriz 1'] = pd.to_numeric(df[col_m1], errors='coerce').fillna(0).astype(int) if col_m1 is not None else 0
        df['Motriz 2'] = pd.to_numeric(df[col_m2], errors='coerce').fillna(0).astype(int) if col_m2 is not None else 0
        df['Unidad'] = df['Motriz 2'].apply(lambda x: 'M' if x > 0 else 'S')
        
        columnas_horas = {}
        for col in df.columns:
            if 'hora salida' in str(col).lower():
                nombre_est = str(col).lower().replace('_hora salida', '').replace('hora salida', '').strip()
                if nombre_est: columnas_horas[f"{nombre_est}_salida"] = col
            elif 'hora llegada' in str(col).lower():
                nombre_est = str(col).lower().replace('_hora llegada', '').replace('hora llegada', '').strip()
                if nombre_est: columnas_horas[f"{nombre_est}_llegada"] = col
        
        for key, col in columnas_horas.items():
            df[f"{key}_min"] = df[col].apply(convertir_a_minutos)
            df[f"{key}_fmt"] = df[f"{key}_min"].apply(lambda x: format_hms(x) if pd.notna(x) else "")
        
        puerto_key = next((k for k in columnas_horas.keys() if 'puerto' in k and 'salida' in k), None)
        limache_key = next((k for k in columnas_horas.keys() if 'limache' in k and 'llegada' in k), None)
        
        df['Hora_Salida_Real'] = df[f"{puerto_key}_min"] if puerto_key else None
        df['Hora_Llegada_Real'] = df[f"{limache_key}_min"] if limache_key else None
        df['Min_Prog'] = df['Hora_Prog'].apply(convertir_a_minutos)
        
        if 'Hora_Salida_Real' in df.columns and 'Min_Prog' in df.columns:
            df['Retraso'] = df['Hora_Salida_Real'] - df['Min_Prog']
        
        if puerto_key and limache_key:
            tdv = df['Hora_Llegada_Real'] - df['Hora_Salida_Real']
            df['TDV_Min'] = tdv.apply(lambda x: x if (pd.notna(x) and x > 0) else (x + 1440 if pd.notna(x) else 0))
        
        df['Tipo_Rec'] = "PU-LI" if (puerto_key and limache_key) else "OTRO"
        df['Dist_Base'] = df['Tipo_Rec'].map(DISTANCIAS).fillna(0)
        df['Tren-Km'] = df['Dist_Base'] * df['Unidad'].apply(lambda x: 2 if x == 'M' else 1)
        
        fecha_nombre = extraer_fecha_desde_nombre_archivo(file.name)
        df['Fecha_Op'] = pd.to_datetime(fecha_nombre if fecha_nombre else date.today())
        
        if start_date and end_date and not df.empty:
            mask = (df['Fecha_Op'].dt.date >= start_date) & (df['Fecha_Op'].dt.date <= end_date)
            df = df[mask].copy()
        
        df.attrs['estaciones_keys'] = list(columnas_horas.keys())
        return df
    except Exception as e:
        st.error(f"Error procesando THDR {file.name}: {str(e)}")
        return pd.DataFrame()

# --- 4. INICIALIZACIÓN ---
df_ops, df_tr, df_tr_acum, df_seat, df_p_d, df_f_d = [pd.DataFrame() for _ in range(6)]
df_thdr_v1, df_thdr_v2 = pd.DataFrame(), pd.DataFrame()
all_comp_full, all_prmte_15, all_fact_h = [], [], []

# --- 5. SIDEBAR ---
with st.sidebar:
    st.header("📅 Filtro Global")
    today = date.today()
    start_of_month = today.replace(day=1)
    date_range = st.date_input("Selecciona el período", value=(start_of_month, today))
    start_date, end_date = (date_range[0], date_range[1]) if isinstance(date_range, tuple) and len(date_range)==2 else (date_range, date_range)
    
    st.divider()
    st.header("📂 Carga de Archivos")
    f_v1 = st.file_uploader("1. THDR Vía 1", type=["xls", "xlsx"], accept_multiple_files=True)
    f_v2 = st.file_uploader("2. THDR Vía 2", type=["xls", "xlsx"], accept_multiple_files=True)
    f_umr = st.file_uploader("3. UMR / Odómetros", type=["xlsx"], accept_multiple_files=True)
    f_seat_files = st.file_uploader("4. Energía SEAT", type=["xlsx"], accept_multiple_files=True)
    f_bill_files = st.file_uploader("5. Facturación y PRMTE", type=["xlsx"], accept_multiple_files=True)

# --- 6. PROCESAMIENTO GENERAL ---
if f_v1 or f_v2 or f_umr or f_seat_files or f_bill_files:
    all_ops, all_tr, all_tr_acum, all_seat = [], [], [], []
    todos = (f_v1 or []) + (f_v2 or []) + (f_umr or []) + (f_seat_files or []) + (f_bill_files or [])

    for f in todos:
        try:
            xl = pd.ExcelFile(f)
            for sn in xl.sheet_names:
                sn_up = sn.upper()
                if any(k in sn_up for k in ['UMR', 'RESUMEN']):
                    df_raw = pd.read_excel(f, sheet_name=sn, header=None)
                    # FIX: Truth value of a Series is ambiguous (convertimos a lista para buscar)
                    h_r = None
                    for i in range(min(100, len(df_raw))):
                        row_values = [str(val).upper() for val in df_raw.iloc[i].tolist()]
                        if any('ODO' in val or 'FECHA' in val for val in row_values):
                            h_r = i; break
                    if h_r is not None:
                        df_p = pd.read_excel(f, sheet_name=sn, header=h_r)
                        df_p.columns = [re.sub(r'[^A-Z]', '', str(c).upper()) for c in df_p.columns]
                        idx_f = next((c for c in df_p.columns if 'FECHA' in c), None)
                        idx_o = next((c for c in df_p.columns if 'ODO' in c), None)
                        idx_t = next((c for c in df_p.columns if 'TREN' in c and 'KM' in c), None)
                        if idx_f and idx_o:
                            df_p['_dt'] = pd.to_datetime(df_p[idx_f], errors='coerce')
                            mask = (df_p['_dt'].dt.date >= start_date) & (df_p['_dt'].dt.date <= end_date)
                            for _, r in df_p[mask].dropna(subset=['_dt']).iterrows():
                                all_ops.append({"Fecha": r['_dt'].normalize(), "Tipo Día": get_tipo_dia(r['_dt']), "N° Semana": r['_dt'].isocalendar()[1], "Odómetro [km]": parse_latam_number(r[idx_o]), "Tren-Km [km]": parse_latam_number(r[idx_t]), "UMR [%]": (parse_latam_number(r[idx_t])/parse_latam_number(r[idx_o])*100 if parse_latam_number(r[idx_o])>0 else 0)})

                if 'ODO' in sn_up and 'KIL' in sn_up:
                    df_tr_raw = pd.read_excel(f, sheet_name=sn, header=None)
                    for i in range(len(df_tr_raw)-2):
                        for j in range(1, len(df_tr_raw.columns)):
                            v = pd.to_datetime(df_tr_raw.iloc[i, j], errors='coerce')
                            if pd.notna(v) and start_date <= v.date() <= end_date:
                                for k in range(i+3, min(i+40, len(df_tr_raw))):
                                    n_tr = str(df_tr_raw.iloc[k, 0]).strip().upper()
                                    if re.match(r'^(M|XM)', n_tr):
                                        all_tr.append({"Tren": n_tr, "Fecha": v.normalize(), "Valor": parse_latam_number(df_tr_raw.iloc[k, j])})

                if 'SEAT' in sn_up and 'SER' in sn_up:
                    df_s = pd.read_excel(f, sheet_name=sn, header=None)
                    for i in range(len(df_s)):
                        fs = pd.to_datetime(df_s.iloc[i, 1], errors='coerce')
                        if pd.notna(fs) and start_date <= fs.date() <= end_date:
                            all_seat.append({"Fecha": fs.normalize(), "Total [kWh]": parse_latam_number(df_s.iloc[i, 3]), "Tracción [kWh]": parse_latam_number(df_s.iloc[i, 5]), "12 KV [kWh]": parse_latam_number(df_s.iloc[i, 7])})
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

# --- 7. TABS DASHBOARD ---
tabs = st.tabs(["📊 Resumen", "📑 Operaciones", "📑 Trenes", "⚡ Energía", "⚖️ Comparación Energía hr", "📈 Regresión Nocturna", "🚨 Datos Atípicos", "📋 THDR"])

# ================== PESTAÑA RESUMEN ==================
with tabs[0]:
    if not df_ops.empty:
        if 'filtros_compartidos' not in st.session_state:
            st.session_state.filtros_compartidos = {'anios': [], 'meses': [], 'semanas': [], 'jornadas': []}
        
        def mostrar_filtros_compartidos(df):
            c1, c2, c3 = st.columns(3)
            anios = sorted(df['Fecha'].dt.year.unique())
            meses = sorted(df['Fecha'].dt.month.unique())
            
            # FIX: StreamlitAPIException (limpiar default si no está en opciones)
            def_anios = [a for a in st.session_state.filtros_compartidos['anios'] if a in anios]
            def_meses = [m for m in st.session_state.filtros_compartidos['meses'] if m in meses]
            
            f_ano = c1.multiselect("Año", anios, default=def_anios or anios, key="filtro_ano")
            f_mes = c2.multiselect("Mes", meses, default=def_meses or meses, key="filtro_mes")
            
            st.session_state.filtros_compartidos['anios'] = f_ano
            st.session_state.filtros_compartidos['meses'] = f_mes
            
            if 'N° Semana' in df.columns:
                semanas = sorted(df['N° Semana'].unique())
                def_semanas = [s for s in st.session_state.filtros_compartidos['semanas'] if s in semanas]
                f_sem = c3.multiselect("N° Semana", semanas, default=def_semanas or semanas, key="filtro_sem")
                st.session_state.filtros_compartidos['semanas'] = f_sem
            
            if 'Tipo Día' in df.columns:
                jornadas = [d for d in ORDEN_TIPO_DIA if d in df['Tipo Día'].unique()]
                def_jor = [j for j in st.session_state.filtros_compartidos['jornadas'] if j in jornadas]
                f_jor = st.multiselect("Jornada", jornadas, default=def_jor or jornadas, key="filtro_jor")
                st.session_state.filtros_compartidos['jornadas'] = f_jor

        mostrar_filtros_compartidos(df_ops)
        
        # Aplicar filtros
        mask = (df_ops['Fecha'].dt.year.isin(st.session_state.filtros_compartidos['anios'])) & \
               (df_ops['Fecha'].dt.month.isin(st.session_state.filtros_compartidos['meses']))
        df_res_f = df_ops[mask]

        if not df_res_f.empty:
            to_val = df_res_f["Odómetro [km]"].sum()
            tk_val = df_res_f["Tren-Km [km]"].sum()
            
            m1, m2, m3 = st.columns(3)
            m1.metric("Odómetro Total", f"{to_val:,.1f} km")
            m2.metric("Tren-Km Total", f"{tk_val:,.1f} km")
            m3.metric("IDE Promedio", f"{df_res_f['IDE (kWh/km)'].mean():.4f} kWh/km")
            
            fig = go.Figure()
            fig.add_trace(go.Bar(x=df_res_f['Fecha'], y=df_res_f['Odómetro [km]'], name='Odómetro'))
            st.plotly_chart(fig, use_container_width=True)
            
            # Exportación Resumen
            metrics = {"Odómetro": f"{to_val:,.1f}", "Tren-Km": f"{tk_val:,.1f}"}
            st.download_button("📈 Exportar Resumen XLSX", exportar_resumen_excel(metrics, df_res_f, None), "Resumen_EFE.xlsx")
    else: st.info("Sube archivos para ver el resumen.")

# ================== PESTAÑA OPERACIONES ==================
with tabs[1]:
    if not df_ops.empty:
        st.dataframe(df_ops.style.format({'Odómetro [km]': "{:,.1f}", 'Tren-Km [km]': "{:,.1f}", 'IDE (kWh/km)': "{:.4f}"}))
    else: st.info("Sin datos de operaciones.")

# ================== PESTAÑA TRENES ==================
with tabs[2]:
    if not df_tr.empty:
        piv = df_tr.pivot_table(index="Tren", columns=df_tr["Fecha"].dt.day, values="Valor", aggfunc='sum').fillna(0)
        st.write("### Kilometraje Diario")
        st.dataframe(piv.style.format("{:,.1f}"))
    else: st.info("Sin datos de trenes.")

# ================== PESTAÑA ENERGÍA ==================
with tabs[3]:
    if not df_seat.empty:
        st.dataframe(df_seat.style.format({'Total [kWh]': "{:,.0f}", 'Tracción [kWh]': "{:,.0f}"}))
    else: st.info("Sin datos SEAT.")

# ================== PESTAÑA COMPARATIVA ==================
with tabs[4]:
    st.info("Carga datos de PRMTE y Factura para comparar consumo horario.")

# ================== PESTAÑA REGRESIÓN ==================
with tabs[5]:
    if not df_ops.empty:
        x = np.arange(len(df_ops))
        y = df_ops['Odómetro [km]'].values
        m, n = np.polyfit(x, y, 1)
        st.write(f"Ecuación: y = {m:.2f}x + {n:.2f}")
        st.line_chart(df_ops.set_index('Fecha')['Odómetro [km]'])

# ================== PESTAÑA ATÍPICOS ==================
with tabs[6]:
    st.info("Análisis de outliers basado en consumo basal (00:00 - 05:00).")

# ================== PESTAÑA THDR ==================
with tabs[7]:
    st.header("📋 Datos THDR")
    if not df_thdr_v1.empty:
        st.subheader("Vía 1 (Puerto -> Limache)")
        st.dataframe(df_thdr_v1)
    if not df_thdr_v2.empty:
        st.subheader("Vía 2 (Limache -> Puerto)")
        st.dataframe(df_thdr_v2)

# --- 8. DESCARGA FINAL ---
st.sidebar.download_button("📥 Reporte Consolidado", to_excel_consolidado(df_ops, df_tr, df_tr_acum, df_seat, df_p_d, pd.DataFrame(all_prmte_15), pd.DataFrame(all_fact_h), df_f_d), "Reporte_SGE_EFE.xlsx")
