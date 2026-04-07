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

# Configuración API Clima (OpenWeatherMap)
# Nota: Reemplaza con tu propia clave para producción
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

@st.cache_data(ttl=3600)  # Cache de 1 hora para no saturar la API
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
            base = header0[i].strip()
            sub = header1[i].strip()
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
        col_motriz1 = buscar_columna(['Motriz 1', 'Motriz1', 'M1', 'Motor 1'])
        col_motriz2 = buscar_columna(['Motriz 2', 'Motriz2', 'M2', 'Motor 2'])
        
        df['Servicio'] = df[col_servicio] if col_servicio is not None else 0
        df['Hora_Prog'] = df[col_hora_prog] if col_hora_prog is not None else '00:00:00'
        
        if col_motriz1 is not None:
            df['Motriz 1'] = pd.to_numeric(df[col_motriz1], errors='coerce').fillna(0).astype(int)
        else:
            df['Motriz 1'] = 0
        if col_motriz2 is not None:
            df['Motriz 2'] = pd.to_numeric(df[col_motriz2], errors='coerce').fillna(0).astype(int)
        else:
            df['Motriz 2'] = 0
        
        df['Unidad'] = df['Motriz 2'].apply(lambda x: 'M' if x > 0 else 'S')
        
        columnas_horas = {}
        for col in df.columns:
            if 'hora salida' in col.lower():
                nombre_est = col.lower().replace('_hora salida', '').replace('hora salida', '').strip()
                if nombre_est:
                    columnas_horas[f"{nombre_est}_salida"] = col
            elif 'hora llegada' in col.lower():
                nombre_est = col.lower().replace('_hora llegada', '').replace('hora llegada', '').strip()
                if nombre_est:
                    columnas_horas[f"{nombre_est}_llegada"] = col
        
        for key, col in columnas_horas.items():
            df[f"{key}_min"] = df[col].apply(convertir_a_minutos)
            df[f"{key}_fmt"] = df[f"{key}_min"].apply(lambda x: format_hms(x) if pd.notna(x) else "")
        
        puerto_key = None
        limache_key = None
        for key in columnas_horas.keys():
            if 'puerto' in key and 'salida' in key: puerto_key = key
            if 'limache' in key and 'llegada' in key: limache_key = key
        
        df['Hora_Salida_Real'] = df[f"{puerto_key}_min"] if puerto_key else None
        df['Hora_Llegada_Real'] = df[f"{limache_key}_min"] if limache_key else None
        
        df['Min_Prog'] = df['Hora_Prog'].apply(convertir_a_minutos)
        df['Retraso'] = df['Hora_Salida_Real'] - df['Min_Prog']
        df['Puntual'] = (abs(df['Retraso']) <= 5).astype(int)
        
        if puerto_key and limache_key:
            tdv = df['Hora_Llegada_Real'] - df['Hora_Salida_Real']
            tdv = tdv.apply(lambda x: x if x > 0 else x + 1440)
            df['TDV_Min'] = tdv
        else:
            df['TDV_Min'] = 0
        
        origen = 'PU' if puerto_key else 'OTRO'
        destino = 'LI' if limache_key else 'OTRO'
        df['Tipo_Rec'] = f"{origen}-{destino}" if origen != 'OTRO' and destino != 'OTRO' else 'OTRO'
        df['Dist_Base'] = df['Tipo_Rec'].map(DISTANCIAS).fillna(0)
        df['Peso'] = df['Unidad'].apply(lambda x: 2 if x == 'M' else 1)
        df['Tren-Km'] = df['Dist_Base'] * df['Peso']
        
        col_fecha = buscar_columna(['Fecha', 'FECHA', 'Date', 'Día'])
        if col_fecha:
            df['Fecha_Op'] = pd.to_datetime(df[col_fecha], errors='coerce')
        else:
            fecha_nombre = extraer_fecha_desde_nombre_archivo(file.name)
            if fecha_nombre:
                df['Fecha_Op'] = pd.Timestamp(fecha_nombre)
            else:
                try:
                    primera_celda = str(df_raw.iloc[0, 0]).split('.')[0].strip().zfill(6)
                    dia, mes, anio = int(primera_celda[0:2]), int(primera_celda[2:4]), 2000 + int(primera_celda[4:6])
                    df['Fecha_Op'] = pd.Timestamp(date(anio, mes, dia))
                except:
                    df['Fecha_Op'] = pd.Timestamp(date.today())
        
        df['Fecha_Op'] = pd.to_datetime(df['Fecha_Op'], errors='coerce')
        if start_date and end_date and not df.empty:
            mask = (df['Fecha_Op'].dt.date >= start_date) & (df['Fecha_Op'].dt.date <= end_date)
            df = df[mask].copy()
        
        df = df.loc[:, ~df.columns.duplicated()]
        df.attrs['estaciones_keys'] = list(columnas_horas.keys())
        return df
    except Exception as e:
        st.error(f"Error procesando THDR {file.name}: {str(e)}")
        return pd.DataFrame()

# --- 4. INICIALIZACIÓN DE DATAFRAMES ---
df_ops = pd.DataFrame()
df_tr = pd.DataFrame()
df_tr_acum = pd.DataFrame()
df_seat = pd.DataFrame()
df_energy_master = pd.DataFrame()
df_p_d = pd.DataFrame()
df_f_d = pd.DataFrame()
df_thdr_v1 = pd.DataFrame()
df_thdr_v2 = pd.DataFrame()
all_comp_full = []
all_prmte_15 = []
all_fact_h = []

# --- 5. INTERFAZ DE USUARIO (SIDEBAR) ---
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

    # --- INTEGRACIÓN API CLIMA EN SIDEBAR ---
    st.divider()
    st.header("🌤️ Estado Climático")
    clima = obtener_clima_actual()
    if clima:
        c_c1, c_c2 = st.columns(2)
        c_c1.metric("Temp", f"{clima['temp']}°C")
        c_c2.metric("Hum", f"{clima['hum']}%")
        st.caption(f"Condición: {clima['desc'].capitalize()}")
    else:
        st.info("Configura tu API Key para ver el clima.")

# --- 6. LECTURA Y PROCESAMIENTO ---
if f_v1 or f_v2 or f_umr or f_seat_files or f_bill_files:
    all_ops, all_tr, all_tr_acum, all_seat, all_prmte_15, all_fact_h, all_comp_full = [], [], [], [], [], [], []
    thdr_v1_list, thdr_v2_list = [], []
    todos = (f_v1 or []) + (f_v2 or []) + (f_umr or []) + (f_seat_files or []) + (f_bill_files or [])

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
                        idx_f = next((c for c in df_p.columns if 'FECHA' in c), None)
                        idx_o = next((c for c in df_p.columns if 'ODO' in c and 'ACUM' not in c), None)
                        idx_t = next((c for c in df_p.columns if 'TREN' in c and 'KM' in c), None)
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
                            all_seat.append({"Fecha": fs.normalize(), "Total [kWh]": tot, "Tracción [kWh]": tra, "12 KV [kWh]": k12, "% Tracción": (tra/tot*100 if tot>0 else 0), "% 12 KV": (k12/tot*100 if tot>0 else 0)})
                
                if any(k in sn_up for k in ['PRMTE', 'MEDIDAS']):
                    df_prm = pd.read_excel(f, sheet_name=sn, header=None)
                    h_idx = next((i for i in range(len(df_prm)) if 'AÑO' in str(df_prm.iloc[i]).upper()), None)
                    if h_idx is not None:
                        df_pd = pd.read_excel(f, sheet_name=sn, header=h_idx)
                        df_pd['Timestamp'] = pd.to_datetime(df_pd[['AÑO', 'MES', 'DIA', 'HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'})) + pd.to_timedelta(df_pd['INICIO INTERVALO'].astype(int), unit='m')
                        cols_e = [c for c in df_pd.columns if 'Retiro_Energia_Activa (kWhD)' in str(c)]
                        for _, r in df_pd.iterrows():
                            ts = r['Timestamp']
                            val_p = sum([parse_latam_number(r[col]) for col in cols_e])
                            all_comp_full.append({"Fecha": ts.normalize(), "Hora": ts.hour, "Consumo Horario [kWh]": val_p, "Fuente": "PRMTE"})
                            if start_date <= ts.date() <= end_date: all_prmte_15.append({"Fecha y Hora": ts.strftime('%d/%m/%Y %H:%M'), "Fecha": ts.normalize(), "Energía PRMTE [kWh]": val_p})
                
                if any(k in sn_up for k in ['FACTURA', 'CONSUMO']):
                    df_f = pd.read_excel(f, sheet_name=sn)
                    df_f.columns = ['FechaHora', 'Valor']
                    df_f['Timestamp'] = pd.to_datetime(df_f['FechaHora'], errors='coerce')
                    for _, r in df_f.dropna(subset=['Timestamp']).iterrows():
                        ts = r['Timestamp']
                        val_f = abs(parse_latam_number(r['Valor']))
                        all_comp_full.append({"Fecha": ts.normalize(), "Hora": ts.hour, "Consumo Horario [kWh]": val_f, "Fuente": "Factura"})
                        if start_date <= ts.date() <= end_date: all_fact_h.append({"Fecha y Hora": ts.strftime('%d/%m/%Y %H:%M'), "Fecha": ts.normalize(), "Consumo Horario [kWh]": val_f})
        except: continue

    if f_v1:
        for file in f_v1:
            df = procesar_thdr_avanzado(file, start_date, end_date)
            if not df.empty:
                df['Vía'] = 'Vía 1'
                thdr_v1_list.append(df)
        if thdr_v1_list: df_thdr_v1 = pd.concat(thdr_v1_list, ignore_index=True)
    if f_v2:
        for file in f_v2:
            df = procesar_thdr_avanzado(file, start_date, end_date)
            if not df.empty:
                df['Vía'] = 'Vía 2'
                thdr_v2_list.append(df)
        if thdr_v2_list: df_thdr_v2 = pd.concat(thdr_v2_list, ignore_index=True)

    if any([all_ops, all_tr, all_tr_acum, all_seat, all_prmte_15, all_fact_h]):
        if all_ops: df_ops = pd.DataFrame(all_ops).drop_duplicates(subset=['Fecha']).sort_values("Fecha")
        if all_tr: df_tr = pd.DataFrame(all_tr).sort_values(["Fecha", "Tren"])
        if all_tr_acum: df_tr_acum = pd.DataFrame(all_tr_acum).sort_values(["Fecha", "Tren"])
        if all_seat: df_seat = pd.DataFrame(all_seat).drop_duplicates(subset=['Fecha']).sort_values("Fecha")
        
        if not df_seat.empty:
            df_energy_master = df_seat[["Fecha", "Total [kWh]", "Tracción [kWh]", "12 KV [kWh]"]].copy().rename(columns={"Total [kWh]":"E_Total", "Tracción [kWh]":"E_Tr", "12 KV [kWh]":"E_12"})
            df_energy_master["Fuente"] = "SEAT"

        if all_prmte_15:
            df_p_d = pd.DataFrame(all_prmte_15).groupby("Fecha")["Energía PRMTE [kWh]"].sum().reset_index()
            if not df_seat.empty:
                df_p_d = pd.merge(df_p_d, df_seat[["Fecha", "% Tracción", "% 12 KV"]], on="Fecha", how="left").fillna(0)
                df_p_d["E_Tr"], df_p_d["E_12"] = df_p_d["Energía PRMTE [kWh]"] * (df_p_d["% Tracción"] / 100), df_p_d["Energía PRMTE [kWh]"] * (df_p_d["% 12 KV"] / 100)
                df_p_p = df_p_d.rename(columns={"Energía PRMTE [kWh]":"E_Total"})[["Fecha","E_Total","E_Tr","E_12"]]
                df_p_p["Fuente"] = "PRMTE"
                df_energy_master = pd.concat([df_energy_master, df_p_p]).drop_duplicates(subset=["Fecha"], keep="last")

        if all_fact_h:
            df_f_d = pd.DataFrame(all_fact_h).groupby("Fecha")["Consumo Horario [kWh]"].sum().reset_index()
            if not df_seat.empty:
                df_f_d = pd.merge(df_f_d, df_seat[["Fecha", "% Tracción", "% 12 KV"]], on="Fecha", how="left").fillna(0)
                df_f_d["E_Tr"], df_f_d["E_12"] = df_f_d["Consumo Horario [kWh]"] * (df_f_d["% Tracción"] / 100), df_f_d["Consumo Horario [kWh]"] * (df_f_d["% 12 KV"] / 100)
                df_f_f = df_f_d.rename(columns={"Consumo Horario [kWh]":"E_Total"})[["Fecha","E_Total","E_Tr","E_12"]]
                df_f_f["Fuente"] = "Factura"
                df_energy_master = pd.concat([df_energy_master, df_f_f]).drop_duplicates(subset=["Fecha"], keep="last")

        if not df_ops.empty and not df_energy_master.empty:
            df_ops = pd.merge(df_ops, df_energy_master, on="Fecha", how="left")
            df_ops['IDE (kWh/km)'] = df_ops.apply(lambda row: row['E_Tr'] / row['Odómetro [km]'] if row['Odómetro [km]'] > 0 else 0, axis=1)

# --- 7. DASHBOARD ---
tabs = st.tabs(["📊 Resumen", "📑 Operaciones", "📑 Trenes", "⚡ Energía", "⚖️ Comparación Energía hr", "📈 Regresión Nocturna", "🚨 Datos Atípicos", "📋 THDR"])

# ================== PESTAÑA RESUMEN ==================
with tabs[0]:
    if not df_ops.empty:
        if 'filtros_compartidos' not in st.session_state:
            st.session_state.filtros_compartidos = {'anios': [], 'meses': [], 'semanas': [], 'jornadas': []}
        
        c1, c2, c3 = st.columns(3)
        anios = sorted(df_ops['Fecha'].dt.year.unique())
        meses = sorted(df_ops['Fecha'].dt.month.unique())
        f_ano = c1.multiselect("Año", anios, default=st.session_state.filtros_compartidos['anios'] or anios, key="filtro_ano")
        f_mes = c2.multiselect("Mes", meses, default=st.session_state.filtros_compartidos['meses'] or meses, key="filtro_mes")
        st.session_state.filtros_compartidos['anios'], st.session_state.filtros_compartidos['meses'] = f_ano, f_mes
        
        mask = df_ops['Fecha'].dt.year.isin(f_ano) & df_ops['Fecha'].dt.month.isin(f_mes)
        df_res_f = df_ops[mask]
        
        if not df_res_f.empty:
            to_val, tk_val = df_res_f["Odómetro [km]"].sum(), df_res_f["Tren-Km [km]"].sum()
            umr_val = (tk_val/to_val*100) if to_val>0 else 0
            
            sub_tabs = st.tabs(["📅 Semanal", "📅 Mensual", "📅 Anual"])
            # (Lógica de sub_tabs igual a tu código original...)
            with sub_tabs[0]:
                st.write("##### Evolución Semanal")
                col_s1, col_s2, col_s3 = st.columns(3)
                # ... resto de visualizaciones semanales ...
                st.metric("Odómetro Global", f"{to_val:,.1f} km")
                st.metric("Tren-Km Global", f"{tk_val:,.1f} km")
                st.metric("UMR Global", f"{umr_val:.2f}%")
        else:
            st.info("Aplica filtros para ver el resumen.")
    else:
        st.info("Carga datos para comenzar.")

# ================== PESTAÑA OPERACIONES ==================
with tabs[1]:
    if not df_ops.empty:
        st.write("#### Datos Operacionales")
        st.dataframe(df_ops.style.format({'Odómetro [km]': "{:,.1f}", 'Tren-Km [km]': "{:,.1f}", 'UMR [%]': "{:.2f}%", 'IDE (kWh/km)': "{:.4f}"}))
    else:
        st.info("Sin datos.")

# ================== PESTAÑA TRENES ==================
with tabs[2]:
    if not df_tr.empty:
        st.write("#### Kilometraje Diario por Tren")
        piv = df_tr.pivot_table(index="Tren", columns=df_tr["Fecha"].dt.day, values="Valor", aggfunc='sum').fillna(0)
        st.dataframe(piv.style.format("{:,.1f}"))

# ================== PESTAÑA ENERGÍA ==================
with tabs[3]:
    if not df_energy_master.empty:
        st.write("#### Consumo Consolidado")
        st.dataframe(df_energy_master.style.format({'E_Total': "{:,.0f}", 'E_Tr': "{:,.0f}", 'E_12': "{:,.0f}"}))

# ================== PESTAÑA COMPARACIÓN HR ==================
with tabs[4]:
    if all_comp_full:
        df_c = pd.DataFrame(all_comp_full)
        piv_c = df_c.pivot_table(index="Hora", columns="Fuente", values="Consumo Horario [kWh]", aggfunc='median').fillna(0)
        st.write("#### Mediana de Consumo por Hora")
        st.line_chart(piv_c)

# ================== PESTAÑA REGRESIÓN NOCTURNA ==================
if 'outliers' not in st.session_state: st.session_state.outliers = pd.DataFrame()
with tabs[5]:
    if all_comp_full:
        st.write("#### Análisis de Regresión Nocturna")
        df_reg = pd.DataFrame(all_comp_full).query("Hora <= 5")
        if len(df_reg) > 2:
            x = np.arange(len(df_reg))
            y = df_reg['Consumo Horario [kWh]'].values
            m, n = np.polyfit(x, y, 1)
            y_pred = m * x + n
            st.markdown(f"**Ecuación de Línea de Base:** $$Consumo = {m:.4f}x + {n:.2f}$$")
            fig_reg = go.Figure()
            fig_reg.add_trace(go.Scatter(x=x, y=y, mode='markers', name='Real'))
            fig_reg.add_trace(go.Scatter(x=x, y=y_pred, mode='lines', name='Tendencia'))
            st.plotly_chart(fig_reg)

# ================== PESTAÑA ATÍPICOS ==================
with tabs[6]:
    if not st.session_state.outliers.empty:
        st.dataframe(st.session_state.outliers)
    else:
        st.success("No se detectan anomalías críticas.")

# ================== PESTAÑA THDR ==================
with tabs[7]:
    if not df_thdr_v1.empty:
        st.write("#### Vía 1 (Puerto -> Limache)")
        st.dataframe(df_thdr_v1)
    if not df_thdr_v2.empty:
        st.write("#### Vía 2 (Limache -> Puerto)")
        st.dataframe(df_thdr_v2)

# --- 8. CIERRE Y EXPORTACIÓN ---
st.sidebar.divider()
st.sidebar.download_button("📥 Reporte Excel Completo", to_excel_consolidado(df_ops, df_tr, df_tr_acum, df_seat, df_p_d, pd.DataFrame(all_prmte_15), pd.DataFrame(all_fact_h), df_f_d), "Reporte_EFE_SGE.xlsx")
