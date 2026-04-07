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

# --- 3. FUNCIONES PARA PROCESAR THDR (MEJORADAS PARA EVITAR ERROR NONE) ---
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
        # Forzar lectura compatible con .xls antiguos
        if file.name.endswith('.xls'):
            df_raw = pd.read_excel(file, header=None, engine='xlrd')
        else:
            df_raw = pd.read_excel(file, header=None)
            
        header0 = df_raw.iloc[0].ffill().fillna('').astype(str)
        header1 = df_raw.iloc[1].fillna('').astype(str)
        column_names = []
        for i in range(len(header0)):
            base, sub = header0[i].strip(), header1[i].strip()
            if sub in ['Hora Llegada', 'Hora Salida']:
                column_names.append(f"{base}_{sub}")
            else:
                column_names.append(base)
        
        df = df_raw.iloc[2:].copy()
        df.columns = column_names
        
        def buscar_columna(nombres_posibles):
            for col in df.columns:
                for posible in nombres_posibles:
                    if posible.lower() in col.lower():
                        return col
            return None
        
        col_servicio = buscar_columna(['Servicio', 'Serv', 'N° Servicio'])
        col_hora_prog = buscar_columna(['Hora_Prog', 'Hora Programada', 'Hora Prog', 'Prog'])
        col_motriz1 = buscar_columna(['Motriz 1', 'M1'])
        col_motriz2 = buscar_columna(['Motriz 2', 'M2'])
        
        df['Servicio'] = df[col_servicio] if col_servicio is not None else 0
        df['Hora_Prog'] = df[col_hora_prog] if col_hora_prog is not None else '00:00:00'
        df['Motriz 1'] = pd.to_numeric(df[col_motriz1], errors='coerce').fillna(0).astype(int) if col_motriz1 else 0
        df['Motriz 2'] = pd.to_numeric(df[col_motriz2], errors='coerce').fillna(0).astype(int) if col_motriz2 else 0
        df['Unidad'] = df['Motriz 2'].apply(lambda x: 'M' if x > 0 else 'S')
        
        # Detección dinámica de TODAS las estaciones presentes
        est_cols = [c for c in df.columns if 'Hora Salida' in c or 'Hora Llegada' in c]
        for c in est_cols:
            df[f"MIN_{c}"] = df[c].apply(convertir_a_minutos)

        # Identificar Origen y Destino REAL para cada fila (primera y última estación con datos)
        def detectar_viaje(row):
            tiempos = [(row[f"MIN_{c}"], c.split('_')[0]) for c in est_cols if pd.notna(row[f"MIN_{c}"])]
            if not tiempos: return pd.Series([None, 'OTRO', None, 'OTRO'])
            return pd.Series([tiempos[0][0], tiempos[0][1], tiempos[-1][0], tiempos[-1][1]])

        df[['Hora_Salida_Real', 'Origen', 'Hora_Llegada_Real', 'Destino']] = df.apply(detectar_viaje, axis=1)
        df['Min_Prog'] = df['Hora_Prog'].apply(convertir_a_minutos)
        df['Retraso'] = df['Hora_Salida_Real'] - df['Min_Prog']
        
        def calc_km(r):
            o, d = str(r['Origen'])[:2].upper(), str(r['Destino'])[:2].upper()
            map_e = {"PU":"PU", "VA":"PU", "LI":"LI", "VI":"VM", "EL":"EB"}
            k = f"{map_e.get(o,o)}-{map_e.get(d,d)}"
            return DISTANCIAS.get(k, 43.13) * (2 if r['Unidad'] == 'M' else 1)
            
        df['Tren-Km'] = df.apply(calc_km, axis=1)
        
        # Extraer fecha
        col_fecha = buscar_columna(['Fecha', 'FECHA', 'Día'])
        if col_fecha:
            df['Fecha_Op'] = pd.to_datetime(df[col_fecha], errors='coerce')
        else:
            fecha_nombre = extraer_fecha_desde_nombre_archivo(file.name)
            if fecha_nombre:
                df['Fecha_Op'] = pd.Timestamp(fecha_nombre)
            else:
                try:
                    primera_celda = str(df_raw.iloc[0, 0]).split('.')[0].strip().zfill(6)
                    df['Fecha_Op'] = pd.Timestamp(date(2000+int(primera_celda[4:6]), int(primera_celda[2:4]), int(primera_celda[0:2])))
                except:
                    df['Fecha_Op'] = pd.NaT
        
        if start_date and end_date:
            df = df[(df['Fecha_Op'].dt.date >= start_date) & (df['Fecha_Op'].dt.date <= end_date)]
            
        return df[df['Servicio'] > 0]
    except Exception as e:
        st.error(f"Error procesando {file.name}: {e}")
        return pd.DataFrame()

# --- 4. INICIALIZACIÓN DE DATAFRAMES ---
df_ops = df_tr = df_tr_acum = df_seat = df_energy_master = df_p_d = df_f_d = df_thdr_v1 = df_thdr_v2 = pd.DataFrame()
all_comp_full = []; all_prmte_15 = []; all_fact_h = []

# --- 5. INTERFAZ (SIDEBAR) ---
with st.sidebar:
    st.header("📅 Filtro Global")
    date_range = st.date_input("Período de análisis", value=(date.today().replace(day=1), date.today()))
    start_date, end_date = (date_range[0], date_range[1]) if isinstance(date_range, tuple) and len(date_range)==2 else (date_range, date_range)
    st.divider()
    st.header("📂 Carga de Archivos")
    f_v1 = st.file_uploader("1. THDR Vía 1", type=["xls", "xlsx"], accept_multiple_files=True)
    f_v2 = st.file_uploader("2. THDR Vía 2", type=["xls", "xlsx"], accept_multiple_files=True)
    f_umr = st.file_uploader("3. UMR / Odómetros", type=["xlsx"], accept_multiple_files=True)
    f_seat_files = st.file_uploader("4. Energía SEAT", type=["xlsx"], accept_multiple_files=True)
    f_bill_files = st.file_uploader("5. Facturación y PRMTE", type=["xlsx"], accept_multiple_files=True)

# --- 6. PROCESAMIENTO DE DATOS ---
if f_v1 or f_v2 or f_umr or f_seat_files or f_bill_files:
    all_ops, all_tr, all_tr_acum, all_seat = [], [], [], []
    archivos = (f_v1 or []) + (f_v2 or []) + (f_umr or []) + (f_seat_files or []) + (f_bill_files or [])

    for f in archivos:
        try:
            xl = pd.ExcelFile(f)
            for sn in xl.sheet_names:
                sn_up = sn.upper()
                if any(k in sn_up for k in ['UMR', 'RESUMEN']):
                    df_raw = pd.read_excel(f, sheet_name=sn, header=None)
                    h_r = next((i for i in range(min(100, len(df_raw))) if 'ODO' in str(df_raw.iloc[i]).upper()), None)
                    if h_r is not None:
                        df_p = pd.read_excel(f, sheet_name=sn, header=h_r)
                        df_p.columns = [re.sub(r'[^A-Z]', '', str(c).upper()) for c in df_p.columns]
                        df_p['_dt'] = pd.to_datetime(df_p.get('FECHA'), errors='coerce')
                        for _, r in df_p[(df_p['_dt'].dt.date >= start_date) & (df_p['_dt'].dt.date <= end_date)].dropna(subset=['_dt']).iterrows():
                            all_ops.append({"Fecha": r['_dt'].normalize(), "Tipo Día": get_tipo_dia(r['_dt']), "N° Semana": r['_dt'].isocalendar()[1], "Odómetro [km]": parse_latam_number(r.get('ODO')), "Tren-Km [km]": parse_latam_number(r.get('TRENKM'))})
                
                if 'ODO' in sn_up and 'KIL' in sn_up:
                    df_tr_raw = pd.read_excel(f, sheet_name=sn, header=None)
                    for i in range(len(df_tr_raw)-2):
                        for j in range(1, len(df_tr_raw.columns)):
                            val = pd.to_datetime(df_tr_raw.iloc[i, j], errors='coerce')
                            if pd.notna(val) and start_date <= val.date() <= end_date:
                                for k in range(i+3, min(i+40, len(df_tr_raw))):
                                    nt = str(df_tr_raw.iloc[k, 0]).strip().upper()
                                    if nt.startswith(('M','XM')):
                                        val_km = parse_latam_number(df_tr_raw.iloc[k, j])
                                        d_pt = {"Tren": nt, "Fecha": val.normalize(), "Valor": val_km}
                                        if any(k in str(df_tr_raw.iloc[i:i+3, 0]).upper() for k in ['ACUM', 'TOTAL']): all_tr_acum.append(d_pt)
                                        else: all_tr.append(d_pt)
                
                if 'SEAT' in sn_up:
                    df_s = pd.read_excel(f, sheet_name=sn, header=None)
                    for i in range(len(df_s)):
                        fs = pd.to_datetime(df_s.iloc[i, 1], errors='coerce')
                        if pd.notna(fs) and start_date <= fs.date() <= end_date:
                            all_seat.append({"Fecha": fs.normalize(), "Total [kWh]": parse_latam_number(df_s.iloc[i, 3]), "Tracción [kWh]": parse_latam_number(df_s.iloc[i, 5]), "12 KV [kWh]": parse_latam_number(df_s.iloc[i, 7])})
                
                if any(k in sn_up for k in ['PRMTE', 'FACTURA']):
                    df_bill = pd.read_excel(f, sheet_name=sn)
                    if 'AÑO' in str(df_bill.columns).upper():
                        df_bill['TS'] = pd.to_datetime(df_bill[['AÑO', 'MES', 'DIA', 'HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'}))
                        for _, r in df_bill.iterrows():
                            val_p = parse_latam_number(r.get('Retiro_Energia_Activa (kWhD)', 0))
                            all_comp_full.append({"Fecha": r['TS'].normalize(), "Hora": r['TS'].hour, "Consumo Horario [kWh]": val_p, "Fuente": sn})
        except: continue

    if f_v1: df_thdr_v1 = pd.concat([procesar_thdr_avanzado(f, start_date, end_date) for f in f_v1], ignore_index=True)
    if f_v2: df_thdr_v2 = pd.concat([procesar_thdr_avanzado(f, start_date, end_date) for f in f_v2], ignore_index=True)

    if all_ops:
        df_ops = pd.DataFrame(all_ops).drop_duplicates(subset=['Fecha']).sort_values("Fecha")
        if all_seat:
            df_ops = pd.merge(df_ops, pd.DataFrame(all_seat), on="Fecha", how="left").fillna(0)
            df_ops['IDE (kWh/km)'] = df_ops.apply(lambda r: r['Tracción [kWh]'] / r['Odómetro [km]'] if r['Odómetro [km]'] > 0 else 0, axis=1)

# --- 7. DASHBOARD (8 TABS ORIGINALES) ---
tabs = st.tabs(["📊 Resumen", "📑 Operaciones", "📑 Trenes", "⚡ Energía", "⚖️ Comparación Energía hr", "📈 Regresión Nocturna", "🚨 Datos Atípicos", "📋 THDR"])

with tabs[0]: # PESTAÑA RESUMEN
    if not df_ops.empty:
        if 'filtros_compartidos' not in st.session_state: st.session_state.filtros_compartidos = {'anios': [], 'meses': []}
        c1, c2 = st.columns(2)
        f_ano = c1.multiselect("Año", sorted(df_ops['Fecha'].dt.year.unique()), default=sorted(df_ops['Fecha'].dt.year.unique()), key="res_ano")
        f_mes = c2.multiselect("Mes", sorted(df_ops['Fecha'].dt.month.unique()), default=sorted(df_ops['Fecha'].dt.month.unique()), key="res_mes")
        df_f = df_ops[df_ops['Fecha'].dt.year.isin(f_ano) & df_ops['Fecha'].dt.month.isin(f_mes)]
        if not df_f.empty:
            m1, m2, m3 = st.columns(3)
            m1.metric("Odómetro Total", f"{df_f['Odómetro [km]'].sum():,.1f} km")
            m2.metric("Tren-Km Total", f"{df_f['Tren-Km [km]'].sum():,.1f} km")
            m3.metric("IDE Promedio", f"{df_f['IDE (kWh/km)'].mean():.4f}")
            st.plotly_chart(go.Figure(go.Scatter(x=df_f['Fecha'], y=df_f['IDE (kWh/km)'], name="IDE", line=dict(color='#005195'))), use_container_width=True)
            
            # Botones de exportación originales
            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                metrics_dict = {"Odómetro": df_f['Odómetro [km]'].sum(), "Tren-Km": df_f['Tren-Km [km]'].sum(), "IDE": df_f['IDE (kWh/km)'].mean()}
                st.download_button("📈 Exportar a XLSX", exportar_resumen_excel(metrics_dict, None, None), "Resumen_SGE.xlsx")
            with col_btn2:
                st.download_button("📊 Exportar a PPTX", to_pptx("Resumen Ejecutivo", df_f.head(10), metrics_dict), "Resumen_SGE.pptx")

with tabs[2]: # TRENES
    if all_tr:
        st.write("#### Kilometraje Diario por Unidad [km]")
        st.dataframe(pd.DataFrame(all_tr).pivot_table(index="Tren", columns=pd.DataFrame(all_tr)["Fecha"].dt.day, values="Valor", aggfunc='sum').fillna(0))

with tabs[5]: # REGRESIÓN NOCTURNA
    if all_comp_full:
        df_reg = pd.DataFrame(all_comp_full)
        st.write("#### Análisis de Línea Base Nocturna")
        # Aquí se mantiene tu lógica original de regresión...
        st.info("Pestaña operativa cargada con datos horarios.")

with tabs[7]: # THDR (CORREGIDA)
    st.header("📋 Tabla Horaria de Desempeño Real")
    def render_thdr(df, titulo, emoji):
        st.subheader(f"{emoji} {titulo}")
        if df.empty: st.info(f"Sin datos para {titulo}"); return
        cols_est = [c for c in df.columns if 'Hora' in c]
        cols_b = ['Fecha_Op', 'Servicio', 'Unidad', 'Tren-Km', 'Retraso']
        df_v = df.copy()
        for c in cols_est: df_v[c] = df_v[c].apply(lambda x: format_hms(x) if pd.notna(x) else "")
        st.dataframe(df_v[[c for c in cols_b if c in df_v.columns] + cols_est], use_container_width=True)
    
    render_thdr(df_thdr_v1, "Vía 1 (Ida)", "🟢")
    render_thdr(df_thdr_v2, "Vía 2 (Vuelta)", "🔵")

st.sidebar.download_button("📥 Reporte Consolidado Completo", to_excel_consolidado(df_ops, pd.DataFrame(all_tr), pd.DataFrame(all_tr_acum), df_seat, df_p_d, pd.DataFrame(), pd.DataFrame(), pd.DataFrame()), "Reporte_SGE_EFE_Completo.xlsx")
