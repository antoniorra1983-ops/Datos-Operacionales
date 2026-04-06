import streamlit as st
import pandas as pd
import numpy as np
import re
import holidays
from io import BytesIO
from datetime import datetime, date, timedelta
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

def clasificar_flota_func(motriz):
    try:
        num = int(float(motriz))
        if 1 <= num <= 27:
            return "XT-100"
        if 28 <= num <= 35:
            return "XT-M"
        if 101 <= num <= 110:
            return "SFE (Chino)"
        return "OTRO"
    except:
        return "S/I"

DISTANCIAS = {
    "PU-LI": 43.13, "LI-PU": 43.13, "PU-SA": 29.11, "SA-PU": 29.11,
    "EB-PU": 25.40, "PU-EB": 25.40, "VM-LI": 34.03, "LI-VM": 34.03,
    "VM-PU": 9.10,  "PU-VM": 9.10
}

@st.cache_data
def leer_fecha_archivo(file):
    try:
        df = pd.read_excel(file, nrows=1, header=None)
        val = str(df.iloc[0, 0]).split('.')[0].strip().zfill(6)
        return (int(val[0:2]), int(val[2:4]), 2000 + int(val[4:6]))
    except:
        return None

def procesar_thdr_avanzado(file):
    try:
        df_raw = pd.read_excel(file, header=None)
        est_h = df_raw.iloc[0].ffill().values
        df = df_raw.iloc[2:].copy()
        df = df[pd.to_numeric(df.iloc[:, 1], errors='coerce').notna()]
        
        # Mapeo de estaciones
        columnas_base = ["Recorrido", "Servicio", "Hora_Prog", "Motriz 1", "Motriz 2", "Unidad"]
        estaciones_raw = [str(est_h[i]) if pd.notna(est_h[i]) else f"Col_{i}" for i in range(6, len(df_raw.columns))]
        nombres_finales, conteos = list(columnas_base), {}
        for nombre in estaciones_raw:
            if nombre not in conteos:
                conteos[nombre] = 0
                nombres_finales.append(nombre)
            else:
                conteos[nombre] += 1
                nombres_finales.append(f"{nombre}.{conteos[nombre]}")
        df.columns = nombres_finales
        est_cols = nombres_finales[6:]

        def get_trip(row):
            h_reales = row[est_cols].apply(convertir_a_minutos).dropna()
            if len(h_reales) < 2:
                return "OTRO", "OTRO", None, 0, 0
            def cod(n_est):
                n = str(n_est).upper()
                return ("PU" if "PUERTO" in n else
                        "LI" if "LIMACHE" in n else
                        "VM" if "VIÑA" in n or "VINA" in n else
                        "EB" if "BELLOTO" in n else
                        n[:2])
            t_s, t_l = h_reales.iloc[0], h_reales.iloc[-1]
            if t_l < t_s:
                t_l += 1440
            return (f"{cod(h_reales.index[0])}-{cod(h_reales.index[-1])}",
                    cod(h_reales.index[0]),
                    t_s,
                    (t_l - t_s),
                    int(t_s // 60) % 24)

        stats = df.apply(get_trip, axis=1)
        df['Tipo_Rec'], df['Origen'], df['Min_S_Real'], df['TDV_Min'], df['Hora_Salida'] = (
            [x[0] for x in stats], [x[1] for x in stats], [x[2] for x in stats],
            [x[3] for x in stats], [x[4] for x in stats]
        )
        df['Min_Prog'] = df['Hora_Prog'].apply(convertir_a_minutos)
        df['Retraso'] = df['Min_S_Real'] - df['Min_Prog']
        df['Puntual'] = (abs(df['Retraso']) <= 5).astype(int)
        df['Dist_Base'] = df['Tipo_Rec'].map(DISTANCIAS).fillna(0)
        df['Peso'] = df['Unidad'].apply(lambda x: 2 if str(x).strip().upper() == 'M' else 1)
        df['Tren-Km'] = df['Dist_Base'] * df['Peso']
        df['Flota'] = df['Motriz 1'].apply(clasificar_flota_func)
        fecha_info = leer_fecha_archivo(file)
        if fecha_info:
            df['Fecha_Op'] = f"{fecha_info[0]:02d}/{fecha_info[1]:02d}/{fecha_info[2]}"
        return df, df['Tren-Km'].sum(), df[df['TDV_Min'] > 0]['TDV_Min'].mean(), (df['Puntual'].sum() / len(df) * 100) if len(df) > 0 else 0
    except Exception:
        return pd.DataFrame(), 0, 0, 0

# --- 4. INICIALIZACIÓN DE DATAFRAMES VACÍOS (siempre existirán) ---
df_ops = pd.DataFrame()
df_tr = pd.DataFrame()
df_tr_acum = pd.DataFrame()
df_seat = pd.DataFrame()
df_energy_master = pd.DataFrame()
df_p_d = pd.DataFrame()
df_f_d = pd.DataFrame()
df_total = pd.DataFrame()
all_comp_full = []
all_prmte_15 = []
all_fact_h = []

# --- 5. INTERFAZ DE USUARIO (SIDEBAR) ---
with st.sidebar:
    st.header("📅 Filtro Global")
    today = date.today()
    start_of_month = today.replace(day=1) if today.day > 1 else (today.replace(month=today.month-1, day=1) if today.month>1 else today.replace(year=today.year-1, month=12, day=1))
    date_range = st.date_input("Selecciona el período", value=(start_of_month, today))
    start_date, end_date = (date_range[0], date_range[1]) if isinstance(date_range, tuple) and len(date_range)==2 else (date_range, date_range)
    st.divider()
    st.header("📂 Carga de Archivos")
    f_v1 = st.file_uploader("1. THDR Vía 1", type=["xls", "xlsx"], accept_multiple_files=True)
    f_v2 = st.file_uploader("2. THDR Vía 2", type=["xls", "xlsx"], accept_multiple_files=True)
    f_umr = st.file_uploader("3. UMR / Odómetros", type=["xlsx"], accept_multiple_files=True)
    f_seat_files = st.file_uploader("4. Energía SEAT", type=["xlsx"], accept_multiple_files=True)
    f_bill_files = st.file_uploader("5. Facturación y PRMTE", type=["xlsx"], accept_multiple_files=True)

# --- 6. LECTURA Y PROCESAMIENTO DE DATOS (solo si hay archivos) ---
if f_v1 or f_v2 or f_umr or f_seat_files or f_bill_files:
    all_ops, all_tr, all_tr_acum, all_seat, all_prmte_15, all_fact_h, all_comp_full = [], [], [], [], [], [], []
    todos = (f_v1 or []) + (f_v2 or []) + (f_umr or []) + (f_seat_files or []) + (f_bill_files or [])

    for f in todos:
        try:
            xl = pd.ExcelFile(f)
            for sn in xl.sheet_names:
                sn_up = sn.upper()
                # --- UMR / RESUMEN ---
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
                # --- ODOMETRO/KILOMETRAJE (detalle de trenes) ---
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
                # --- SEAT (subestaciones) ---
                if 'SEAT' in sn_up and 'SER' in sn_up:
                    df_s = pd.read_excel(f, sheet_name=sn, header=None)
                    for i in range(len(df_s)):
                        fs = pd.to_datetime(df_s.iloc[i, 1], errors='coerce')
                        if pd.notna(fs) and start_date <= fs.date() <= end_date:
                            tot, tra, k12 = parse_latam_number(df_s.iloc[i, 3]), parse_latam_number(df_s.iloc[i, 5]), parse_latam_number(df_s.iloc[i, 7])
                            all_seat.append({"Fecha": fs.normalize(), "Total [kWh]": tot, "Tracción [kWh]": tra, "12 KV [kWh]": k12, "% Tracción": (tra/tot*100 if tot>0 else 0), "% 12 KV": (k12/tot*100 if tot>0 else 0)})
                # --- PRMTE / MEDIDAS (perfil horario) ---
                if any(k in sn_up for k in ['PRMTE', 'MEDIDAS']):
                    df_prm = pd.read_excel(f, sheet_name=sn, header=None)
                    h_idx = next((i for i in range(len(df_prm)) if 'AÑO' in str(df_prm.iloc[i]).upper()), None)
                    if h_idx is not None:
                        df_pd = pd.read_excel(f, sheet_name=sn, header=h_idx)
                        df_pd['Timestamp'] = pd.to_datetime(df_pd[['AÑO', 'MES', 'DIA', 'HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'})) + pd.to_timedelta(df_pd['INICIO INTERVALO'].astype(int), unit='m')
                        cols_e = [c for c in df_pd.columns if 'Retiro_Energia_Activa (kWhD)' in str(c)]
                        for _, r in df_pd.iterrows():
                            ts, val_p = r['Timestamp'], sum([parse_latam_number(r[col]) for col in cols_e])
                            all_comp_full.append({"Fecha": ts.normalize(), "Hora": ts.hour, "Consumo Horario [kWh]": val_p, "Fuente": "PRMTE"})
                            if start_date <= ts.date() <= end_date: all_prmte_15.append({"Fecha y Hora": ts.strftime('%d/%m/%Y %H:%M'), "Fecha": ts.normalize(), "Energía PRMTE [kWh]": val_p})
                # --- FACTURA / CONSUMO ---
                if any(k in sn_up for k in ['FACTURA', 'CONSUMO']):
                    df_f = pd.read_excel(f, sheet_name=sn); df_f.columns = ['FechaHora', 'Valor']
                    df_f['Timestamp'] = pd.to_datetime(df_f['FechaHora'], errors='coerce')
                    for _, r in df_f.dropna(subset=['Timestamp']).iterrows():
                        ts, val_f = r['Timestamp'], abs(parse_latam_number(r['Valor']))
                        all_comp_full.append({"Fecha": ts.normalize(), "Hora": ts.hour, "Consumo Horario [kWh]": val_f, "Fuente": "Factura"})
                        if start_date <= ts.date() <= end_date: all_fact_h.append({"Fecha y Hora": ts.strftime('%d/%m/%Y %H:%M'), "Fecha": ts.normalize(), "Consumo Horario [kWh]": val_f})
        except: continue

    # --- 7. JERARQUÍA Y PRE-FILTRADO ---
    if any([all_ops, all_tr, all_tr_acum, all_seat, all_prmte_15, all_fact_h]):
        if all_ops: df_ops = pd.DataFrame(all_ops).drop_duplicates(subset=['Fecha']).sort_values("Fecha")
        if all_tr: df_tr = pd.DataFrame(all_tr).sort_values(["Fecha", "Tren"])
        if all_tr_acum: df_tr_acum = pd.DataFrame(all_tr_acum).sort_values(["Fecha", "Tren"])
        if all_seat: df_seat = pd.DataFrame(all_seat).drop_duplicates(subset=['Fecha']).sort_values("Fecha")
        
        # Energía priorizada: Factura > PRMTE > SEAT
        if not df_seat.empty:
            df_energy_master = df_seat[["Fecha", "Total [kWh]", "Tracción [kWh]", "12 KV [kWh]"]].copy().rename(columns={"Total [kWh]":"E_Total", "Tracción [kWh]":"E_Tr", "12 KV [kWh]":"E_12"})
            df_energy_master["Fuente"] = "SEAT"

        if all_prmte_15:
            df_p_d = pd.DataFrame(all_prmte_15).groupby("Fecha")["Energía PRMTE [kWh]"].sum().reset_index()
            if not df_seat.empty:
                df_p_d = pd.merge(df_p_d, df_seat[["Fecha", "% Tracción", "% 12 KV"]], on="Fecha", how="left").fillna(0)
                df_p_d["E_Tr"], df_p_d["E_12"] = df_p_d["Energía PRMTE [kWh]"]*(df_p_d["% Tracción"]/100), df_p_d["Energía PRMTE [kWh]"]*(df_p_d["% 12 KV"]/100)
                df_p_p = df_p_d.rename(columns={"Energía PRMTE [kWh]":"E_Total"})[["Fecha","E_Total","E_Tr","E_12"]]; df_p_p["Fuente"] = "PRMTE"
                df_energy_master = pd.concat([df_energy_master, df_p_p]).drop_duplicates(subset=["Fecha"], keep="last")

        if all_fact_h:
            df_f_d = pd.DataFrame(all_fact_h).groupby("Fecha")["Consumo Horario [kWh]"].sum().reset_index()
            if not df_seat.empty:
                df_f_d = pd.merge(df_f_d, df_seat[["Fecha", "% Tracción", "% 12 KV"]], on="Fecha", how="left").fillna(0)
                df_f_d["E_Tr"], df_f_d["E_12"] = df_f_d["Consumo Horario [kWh]"]*(df_f_d["% Tracción"]/100), df_f_d["Consumo Horario [kWh]"]*(df_f_d["% 12 KV"]/100)
                df_f_f = df_f_d.rename(columns={"Consumo Horario [kWh]":"E_Total"})[["Fecha","E_Total","E_Tr","E_12"]]; df_f_f["Fuente"] = "Factura"
                df_energy_master = pd.concat([df_energy_master, df_f_f]).drop_duplicates(subset=["Fecha"], keep="last")

        if not df_ops.empty and not df_energy_master.empty:
            df_ops = pd.merge(df_ops, df_energy_master, on="Fecha", how="left")
            df_ops['IDE (kWh/km)'] = df_ops.apply(lambda row: row['E_Tr'] / row['Odómetro [km]'] if row['Odómetro [km]'] > 0 else 0, axis=1)
        
        # Asignar df_total (para la pestaña THDR y otros)
        df_total = df_ops.copy()

# --- 8. DASHBOARD (siempre visible) ---
tabs = st.tabs(["📊 Resumen", "📑 Operaciones", "📑 Trenes", "⚡ Energía", "⚖️ Comparación Energía hr", "📈 Regresión Nocturna", "🚨 Datos Atípicos", "📋 THDR"])

with tabs[0]: # Resumen
    if not df_ops.empty:
        # Filtros compartidos (definidos dentro de este bloque)
        if 'filtros_compartidos' not in st.session_state:
            st.session_state.filtros_compartidos = {'anios': [], 'meses': [], 'semanas': [], 'jornadas': []}
        def mostrar_filtros_compartidos(df):
            if df.empty: return
            c1, c2, c3 = st.columns(3)
            anios = sorted(df['Fecha'].dt.year.unique())
            meses = sorted(df['Fecha'].dt.month.unique())
            f_ano = c1.multiselect("Año", anios, default=st.session_state.filtros_compartidos['anios'] or anios, key="filtro_ano")
            f_mes = c2.multiselect("Mes", meses, default=st.session_state.filtros_compartidos['meses'] or meses, key="filtro_mes")
            st.session_state.filtros_compartidos['anios'] = f_ano
            st.session_state.filtros_compartidos['meses'] = f_mes
            if 'N° Semana' in df.columns:
                semanas = sorted(df['N° Semana'].unique())
                f_sem = c3.multiselect("N° Semana", semanas, default=st.session_state.filtros_compartidos['semanas'] or semanas, key="filtro_sem")
                st.session_state.filtros_compartidos['semanas'] = f_sem
            if 'Tipo Día' in df.columns:
                unique_jor = df['Tipo Día'].unique()
                ordered_jor = [d for d in ORDEN_TIPO_DIA if d in unique_jor]
                f_jor = st.multiselect("Jornada", ordered_jor, default=st.session_state.filtros_compartidos['jornadas'] or ordered_jor, key="filtro_jor")
                st.session_state.filtros_compartidos['jornadas'] = f_jor
        def aplicar_filtros_compartidos(df):
            if df.empty: return df
            mask = pd.Series(True, index=df.index)
            if st.session_state.filtros_compartidos['anios']:
                mask &= df['Fecha'].dt.year.isin(st.session_state.filtros_compartidos['anios'])
            if st.session_state.filtros_compartidos['meses']:
                mask &= df['Fecha'].dt.month.isin(st.session_state.filtros_compartidos['meses'])
            if 'N° Semana' in df.columns and st.session_state.filtros_compartidos['semanas']:
                mask &= df['N° Semana'].isin(st.session_state.filtros_compartidos['semanas'])
            if 'Tipo Día' in df.columns and st.session_state.filtros_compartidos['jornadas']:
                mask &= df['Tipo Día'].isin(st.session_state.filtros_compartidos['jornadas'])
            return df[mask]
        mostrar_filtros_compartidos(df_ops)
        df_res_f = aplicar_filtros_compartidos(df_ops)
        if not df_res_f.empty:
            to_val = df_res_f["Odómetro [km]"].sum()
            tk_val = df_res_f["Tren-Km [km]"].sum()
            umr_val = (tk_val/to_val*100) if to_val>0 else 0
            
            # Sub-pestañas: Semanal, Mensual, Anual
            sub_tabs = st.tabs(["📅 Semanal", "📅 Mensual", "📅 Anual"])
            
            with sub_tabs[0]:  # Semanal
                st.write("##### Evolución Semanal")
                col_s1, col_s2, col_s3 = st.columns(3)
                anios_sem = sorted(df_res_f['Fecha'].dt.year.unique())
                f_ano_sem = col_s1.selectbox("Año (semana)", anios_sem, key="sem_ano")
                semanas_df = df_res_f[df_res_f['Fecha'].dt.year == f_ano_sem]['N° Semana'].unique()
                semanas_ord = sorted(semanas_df)
                f_semana = col_s2.selectbox("N° Semana", semanas_ord, key="sem_num")
                tipos_sem = df_res_f['Tipo Día'].unique()
                orden_tipos_sem = [d for d in ORDEN_TIPO_DIA if d in tipos_sem]
                f_tipo_sem = col_s3.multiselect("Tipo Día (semana)", orden_tipos_sem, default=orden_tipos_sem, key="sem_tipo")
                
                mask_sem = (df_res_f['Fecha'].dt.year == f_ano_sem) & (df_res_f['N° Semana'] == f_semana)
                if f_tipo_sem:
                    mask_sem &= df_res_f['Tipo Día'].isin(f_tipo_sem)
                df_semana = df_res_f[mask_sem].sort_values('Fecha')
                if not df_semana.empty:
                    to_val_sem = df_semana["Odómetro [km]"].sum()
                    tk_val_sem = df_semana["Tren-Km [km]"].sum()
                    umr_val_sem = (tk_val_sem/to_val_sem*100) if to_val_sem>0 else 0
                    col_m1, col_m2, col_m3 = st.columns(3)
                    col_m1.metric("Odómetro", f"{to_val_sem:,.1f} km")
                    col_m2.metric("Tren-Km", f"{tk_val_sem:,.1f} km")
                    col_m3.metric("UMR", f"{umr_val_sem:.2f} %")
                    
                    fig = make_subplots(specs=[[{"secondary_y": True}]])
                    fig.add_trace(go.Bar(x=df_semana['Fecha'].dt.strftime('%d/%m'), y=df_semana['Odómetro [km]'] / 1000, name='Odómetro (miles km)', marker_color='#005195'), secondary_y=False)
                    fig.add_trace(go.Bar(x=df_semana['Fecha'].dt.strftime('%d/%m'), y=df_semana['Tren-Km [km]'] / 1000, name='Tren-Km (miles km)', marker_color='#4CAF50'), secondary_y=False)
                    fig.add_trace(go.Scatter(x=df_semana['Fecha'].dt.strftime('%d/%m'), y=df_semana['UMR [%]'], name='UMR (%)', mode='lines+markers', line=dict(color='#FF5733', width=3), marker=dict(size=8)), secondary_y=True)
                    fig.update_layout(title=f"Semana {f_semana} - {f_ano_sem}", xaxis_title="Día del mes", barmode='group', legend_title="Métrica", height=400)
                    fig.update_yaxes(title_text="Kilómetros (miles)", secondary_y=False)
                    fig.update_yaxes(title_text="UMR (%)", secondary_y=True, range=[0, 100])
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # Energía priorizada (semana)
                    st.markdown("#### ⚡ Energía (prioridad: Factura > PRMTE > SEAT)")
                    energia_fechas_sem = []
                    for fecha in df_semana['Fecha'].unique():
                        if not df_f_d.empty and fecha in df_f_d['Fecha'].values:
                            row = df_f_d[df_f_d['Fecha'] == fecha].iloc[0]
                            energia_fechas_sem.append({'Fecha': fecha, 'E_Total': row['Consumo Horario [kWh]'], 'E_Tr': row['E_Tr'], 'E_12': row['E_12'], 'Fuente': 'Factura'})
                        elif not df_p_d.empty and fecha in df_p_d['Fecha'].values:
                            row = df_p_d[df_p_d['Fecha'] == fecha].iloc[0]
                            energia_fechas_sem.append({'Fecha': fecha, 'E_Total': row['Energía PRMTE [kWh]'], 'E_Tr': row['E_Tr'], 'E_12': row['E_12'], 'Fuente': 'PRMTE'})
                        elif not df_seat.empty and fecha in df_seat['Fecha'].values:
                            row = df_seat[df_seat['Fecha'] == fecha].iloc[0]
                            energia_fechas_sem.append({'Fecha': fecha, 'E_Total': row['Total [kWh]'], 'E_Tr': row['Tracción [kWh]'], 'E_12': row['12 KV [kWh]'], 'Fuente': 'SEAT'})
                        else:
                            energia_fechas_sem.append({'Fecha': fecha, 'E_Total': 0, 'E_Tr': 0, 'E_12': 0, 'Fuente': 'Sin datos'})
                    df_energia_sem = pd.DataFrame(energia_fechas_sem)
                    total_energia_sem = df_energia_sem['E_Total'].sum()
                    total_traccion_sem = df_energia_sem['E_Tr'].sum()
                    total_12kv_sem = df_energia_sem['E_12'].sum()
                    fuente_sem = df_energia_sem['Fuente'].iloc[0] if not df_energia_sem.empty else "Sin datos"
                    col_e1, col_e2, col_e3, col_e4 = st.columns(4)
                    col_e1.metric("Energía Total", f"{total_energia_sem:,.0f} kWh")
                    col_e2.metric("Energía Tracción", f"{total_traccion_sem:,.0f} kWh")
                    col_e3.metric("Energía 12 kV", f"{total_12kv_sem:,.0f} kWh")
                    col_e4.metric("Fuente principal", fuente_sem)
                    if total_energia_sem > 0:
                        st.caption(f"⚡ Composición: Tracción {total_traccion_sem/total_energia_sem*100:.1f}% | 12 kV {total_12kv_sem/total_energia_sem*100:.1f}%")
                    
                    # Resumen por jornada (semana)
                    res_j_sem = df_semana.groupby("Tipo Día", observed=True).agg({"Odómetro [km]":"sum", "Tren-Km [km]":"sum", "UMR [%]":"mean"}).reset_index()
                    res_j_sem['Tipo Día'] = pd.Categorical(res_j_sem['Tipo Día'], categories=ORDEN_TIPO_DIA, ordered=True)
                    res_j_sem = res_j_sem.sort_values('Tipo Día').reset_index(drop=True)
                    st.write("#### Resumen por Jornada (semana)")
                    st.table(res_j_sem.style.format({"Odómetro [km]":"{:,.1f}", "Tren-Km [km]":"{:,.1f}", "UMR [%]":"{:.2f}%"}))
                else:
                    st.info("No hay datos para la semana seleccionada.")
            
            with sub_tabs[1]:  # Mensual
                st.write("##### Evolución Mensual")
                col_m1, col_m2, col_m3 = st.columns(3)
                anios_mes = sorted(df_res_f['Fecha'].dt.year.unique())
                f_ano_mes = col_m1.selectbox("Año (mensual)", anios_mes, key="mes_ano")
                meses_mes = sorted(df_res_f[df_res_f['Fecha'].dt.year == f_ano_mes]['Fecha'].dt.month.unique())
                f_mes_mes = col_m2.selectbox("Mes", meses_mes, format_func=lambda x: f"{x:02d}", key="mes_num")
                tipos_mes = df_res_f['Tipo Día'].unique()
                orden_tipos_mes = [d for d in ORDEN_TIPO_DIA if d in tipos_mes]
                f_tipo_mes = col_m3.multiselect("Tipo Día (mensual)", orden_tipos_mes, default=orden_tipos_mes, key="mes_tipo")
                
                mask_mes = (df_res_f['Fecha'].dt.year == f_ano_mes) & (df_res_f['Fecha'].dt.month == f_mes_mes)
                if f_tipo_mes:
                    mask_mes &= df_res_f['Tipo Día'].isin(f_tipo_mes)
                df_mes = df_res_f[mask_mes].sort_values('Fecha')
                if not df_mes.empty:
                    to_val_mes = df_mes["Odómetro [km]"].sum()
                    tk_val_mes = df_mes["Tren-Km [km]"].sum()
                    umr_val_mes = (tk_val_mes/to_val_mes*100) if to_val_mes>0 else 0
                    col_met1, col_met2, col_met3 = st.columns(3)
                    col_met1.metric("Odómetro", f"{to_val_mes:,.1f} km")
                    col_met2.metric("Tren-Km", f"{tk_val_mes:,.1f} km")
                    col_met3.metric("UMR", f"{umr_val_mes:.2f} %")
                    
                    fechas_str = df_mes['Fecha'].dt.strftime('%d/%m')
                    odometro_miles = df_mes['Odómetro [km]'] / 1000
                    trenkm_miles = df_mes['Tren-Km [km]'] / 1000
                    umr = df_mes['UMR [%]']
                    fig = make_subplots(specs=[[{"secondary_y": True}]])
                    fig.add_trace(go.Bar(x=fechas_str, y=odometro_miles, name='Odómetro (miles km)', marker_color='#005195'), secondary_y=False)
                    fig.add_trace(go.Bar(x=fechas_str, y=trenkm_miles, name='Tren-Km (miles km)', marker_color='#4CAF50'), secondary_y=False)
                    fig.add_trace(go.Scatter(x=fechas_str, y=umr, name='UMR (%)', mode='lines+markers', line=dict(color='#FF5733', width=3), marker=dict(size=8)), secondary_y=True)
                    fig.update_layout(title=f"Evolución Diaria - {f_ano_mes}-{f_mes_mes:02d}", xaxis_title="Día del mes", barmode='group', legend_title="Métrica", height=400)
                    fig.update_yaxes(title_text="Kilómetros (miles)", secondary_y=False)
                    fig.update_yaxes(title_text="UMR (%)", secondary_y=True, range=[0, 100])
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # Energía priorizada (mes)
                    st.markdown("#### ⚡ Energía (prioridad: Factura > PRMTE > SEAT)")
                    energia_fechas_mes = []
                    for fecha in df_mes['Fecha'].unique():
                        if not df_f_d.empty and fecha in df_f_d['Fecha'].values:
                            row = df_f_d[df_f_d['Fecha'] == fecha].iloc[0]
                            energia_fechas_mes.append({'Fecha': fecha, 'E_Total': row['Consumo Horario [kWh]'], 'E_Tr': row['E_Tr'], 'E_12': row['E_12'], 'Fuente': 'Factura'})
                        elif not df_p_d.empty and fecha in df_p_d['Fecha'].values:
                            row = df_p_d[df_p_d['Fecha'] == fecha].iloc[0]
                            energia_fechas_mes.append({'Fecha': fecha, 'E_Total': row['Energía PRMTE [kWh]'], 'E_Tr': row['E_Tr'], 'E_12': row['E_12'], 'Fuente': 'PRMTE'})
                        elif not df_seat.empty and fecha in df_seat['Fecha'].values:
                            row = df_seat[df_seat['Fecha'] == fecha].iloc[0]
                            energia_fechas_mes.append({'Fecha': fecha, 'E_Total': row['Total [kWh]'], 'E_Tr': row['Tracción [kWh]'], 'E_12': row['12 KV [kWh]'], 'Fuente': 'SEAT'})
                        else:
                            energia_fechas_mes.append({'Fecha': fecha, 'E_Total': 0, 'E_Tr': 0, 'E_12': 0, 'Fuente': 'Sin datos'})
                    df_energia_mes = pd.DataFrame(energia_fechas_mes)
                    total_energia_mes = df_energia_mes['E_Total'].sum()
                    total_traccion_mes = df_energia_mes['E_Tr'].sum()
                    total_12kv_mes = df_energia_mes['E_12'].sum()
                    fuente_mes = df_energia_mes['Fuente'].iloc[0] if not df_energia_mes.empty else "Sin datos"
                    col_e1, col_e2, col_e3, col_e4 = st.columns(4)
                    col_e1.metric("Energía Total", f"{total_energia_mes:,.0f} kWh")
                    col_e2.metric("Energía Tracción", f"{total_traccion_mes:,.0f} kWh")
                    col_e3.metric("Energía 12 kV", f"{total_12kv_mes:,.0f} kWh")
                    col_e4.metric("Fuente principal", fuente_mes)
                    if total_energia_mes > 0:
                        st.caption(f"⚡ Composición: Tracción {total_traccion_mes/total_energia_mes*100:.1f}% | 12 kV {total_12kv_mes/total_energia_mes*100:.1f}%")
                    
                    # Resumen por jornada (mes)
                    res_j_mes = df_mes.groupby("Tipo Día", observed=True).agg({"Odómetro [km]":"sum", "Tren-Km [km]":"sum", "UMR [%]":"mean"}).reset_index()
                    res_j_mes['Tipo Día'] = pd.Categorical(res_j_mes['Tipo Día'], categories=ORDEN_TIPO_DIA, ordered=True)
                    res_j_mes = res_j_mes.sort_values('Tipo Día').reset_index(drop=True)
                    st.write("#### Resumen por Jornada (mes)")
                    st.table(res_j_mes.style.format({"Odómetro [km]":"{:,.1f}", "Tren-Km [km]":"{:,.1f}", "UMR [%]":"{:.2f}%"}))
                else:
                    st.info("No hay datos para el mes seleccionado.")
            
            with sub_tabs[2]:  # Anual
                st.write("##### Evolución Anual")
                col_a1, col_a2 = st.columns(2)
                anios_anual = sorted(df_res_f['Fecha'].dt.year.unique())
                f_ano_anual = col_a1.selectbox("Año (anual)", anios_anual, key="anual_ano")
                tipos_anual = df_res_f['Tipo Día'].unique()
                orden_tipos_anual = [d for d in ORDEN_TIPO_DIA if d in tipos_anual]
                f_tipo_anual = col_a2.multiselect("Tipo Día (anual)", orden_tipos_anual, default=orden_tipos_anual, key="anual_tipo")
                
                mask_anual = (df_res_f['Fecha'].dt.year == f_ano_anual)
                if f_tipo_anual:
                    mask_anual &= df_res_f['Tipo Día'].isin(f_tipo_anual)
                df_anual = df_res_f[mask_anual].copy()
                if not df_anual.empty:
                    df_anual['Mes'] = df_anual['Fecha'].dt.month
                    df_mensual = df_anual.groupby('Mes').agg({
                        'Odómetro [km]': 'sum',
                        'Tren-Km [km]': 'sum',
                        'UMR [%]': 'mean'
                    }).reset_index()
                    to_val_anio = df_mensual['Odómetro [km]'].sum()
                    tk_val_anio = df_mensual['Tren-Km [km]'].sum()
                    umr_val_anio = (tk_val_anio/to_val_anio*100) if to_val_anio>0 else 0
                    col_met1, col_met2, col_met3 = st.columns(3)
                    col_met1.metric("Odómetro", f"{to_val_anio:,.1f} km")
                    col_met2.metric("Tren-Km", f"{tk_val_anio:,.1f} km")
                    col_met3.metric("UMR", f"{umr_val_anio:.2f} %")
                    
                    odometro_miles = df_mensual['Odómetro [km]'] / 1000
                    trenkm_miles = df_mensual['Tren-Km [km]'] / 1000
                    umr = df_mensual['UMR [%]']
                    fig = make_subplots(specs=[[{"secondary_y": True}]])
                    fig.add_trace(go.Bar(x=df_mensual['Mes'], y=odometro_miles, name='Odómetro (miles km)', marker_color='#005195'), secondary_y=False)
                    fig.add_trace(go.Bar(x=df_mensual['Mes'], y=trenkm_miles, name='Tren-Km (miles km)', marker_color='#4CAF50'), secondary_y=False)
                    fig.add_trace(go.Scatter(x=df_mensual['Mes'], y=umr, name='UMR (%)', mode='lines+markers', line=dict(color='#FF5733', width=3), marker=dict(size=8)), secondary_y=True)
                    fig.update_layout(title=f"Evolución Mensual - {f_ano_anual}", xaxis_title="Mes", barmode='group', legend_title="Métrica", height=400)
                    fig.update_yaxes(title_text="Kilómetros (miles)", secondary_y=False)
                    fig.update_yaxes(title_text="UMR (%)", secondary_y=True, range=[0, 100])
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # Energía priorizada (anual)
                    st.markdown("#### ⚡ Energía (prioridad: Factura > PRMTE > SEAT)")
                    energia_fechas_anio = []
                    for fecha in df_anual['Fecha'].unique():
                        if not df_f_d.empty and fecha in df_f_d['Fecha'].values:
                            row = df_f_d[df_f_d['Fecha'] == fecha].iloc[0]
                            energia_fechas_anio.append({'Fecha': fecha, 'E_Total': row['Consumo Horario [kWh]'], 'E_Tr': row['E_Tr'], 'E_12': row['E_12'], 'Fuente': 'Factura'})
                        elif not df_p_d.empty and fecha in df_p_d['Fecha'].values:
                            row = df_p_d[df_p_d['Fecha'] == fecha].iloc[0]
                            energia_fechas_anio.append({'Fecha': fecha, 'E_Total': row['Energía PRMTE [kWh]'], 'E_Tr': row['E_Tr'], 'E_12': row['E_12'], 'Fuente': 'PRMTE'})
                        elif not df_seat.empty and fecha in df_seat['Fecha'].values:
                            row = df_seat[df_seat['Fecha'] == fecha].iloc[0]
                            energia_fechas_anio.append({'Fecha': fecha, 'E_Total': row['Total [kWh]'], 'E_Tr': row['Tracción [kWh]'], 'E_12': row['12 KV [kWh]'], 'Fuente': 'SEAT'})
                        else:
                            energia_fechas_anio.append({'Fecha': fecha, 'E_Total': 0, 'E_Tr': 0, 'E_12': 0, 'Fuente': 'Sin datos'})
                    df_energia_anio = pd.DataFrame(energia_fechas_anio)
                    total_energia_anio = df_energia_anio['E_Total'].sum()
                    total_traccion_anio = df_energia_anio['E_Tr'].sum()
                    total_12kv_anio = df_energia_anio['E_12'].sum()
                    fuente_anio = df_energia_anio['Fuente'].iloc[0] if not df_energia_anio.empty else "Sin datos"
                    col_e1, col_e2, col_e3, col_e4 = st.columns(4)
                    col_e1.metric("Energía Total", f"{total_energia_anio:,.0f} kWh")
                    col_e2.metric("Energía Tracción", f"{total_traccion_anio:,.0f} kWh")
                    col_e3.metric("Energía 12 kV", f"{total_12kv_anio:,.0f} kWh")
                    col_e4.metric("Fuente principal", fuente_anio)
                    if total_energia_anio > 0:
                        st.caption(f"⚡ Composición: Tracción {total_traccion_anio/total_energia_anio*100:.1f}% | 12 kV {total_12kv_anio/total_energia_anio*100:.1f}%")
                    
                    # Resumen por jornada (anual)
                    res_j_anio = df_anual.groupby("Tipo Día", observed=True).agg({"Odómetro [km]":"sum", "Tren-Km [km]":"sum", "UMR [%]":"mean"}).reset_index()
                    res_j_anio['Tipo Día'] = pd.Categorical(res_j_anio['Tipo Día'], categories=ORDEN_TIPO_DIA, ordered=True)
                    res_j_anio = res_j_anio.sort_values('Tipo Día').reset_index(drop=True)
                    st.write("#### Resumen por Jornada (año)")
                    st.table(res_j_anio.style.format({"Odómetro [km]":"{:,.1f}", "Tren-Km [km]":"{:,.1f}", "UMR [%]":"{:.2f}%"}))
                else:
                    st.info("No hay datos para el año seleccionado.")
            
            # Botones de exportación (comunes)
            st.write("---")
            st.write("#### 📥 Exportar pestaña Resumen")
            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                if st.button("🖨️ Imprimir / Guardar como PDF", use_container_width=True):
                    st.markdown("""
                        <script>
                            window.print();
                        </script>
                    """, unsafe_allow_html=True)
                    st.info("Haz clic derecho y selecciona 'Guardar como PDF' en el diálogo de impresión.")
            with col_btn2:
                if st.button("📈 Exportar a XLSX", use_container_width=True):
                    metrics_dict = {
                        "Odómetro Total (km)": to_val,
                        "Tren-Km Total (km)": tk_val,
                        "UMR Global (%)": umr_val,
                        "Energía Total (kWh)": df_res_f['E_Total'].sum() if 'E_Total' in df_res_f else 0,
                        "Energía Tracción (kWh)": df_res_f['E_Tr'].sum() if 'E_Tr' in df_res_f else 0,
                        "Energía 12 kV (kWh)": df_res_f['E_12'].sum() if 'E_12' in df_res_f else 0,
                        "Fuente principal": "Factura" if not df_f_d.empty else ("PRMTE" if not df_p_d.empty else "SEAT")
                    }
                    res_j_total = df_res_f.groupby("Tipo Día", observed=True).agg({"Odómetro [km]":"sum", "Tren-Km [km]":"sum", "UMR [%]":"mean"}).reset_index()
                    res_j_total['Tipo Día'] = pd.Categorical(res_j_total['Tipo Día'], categories=ORDEN_TIPO_DIA, ordered=True)
                    res_j_total = res_j_total.sort_values('Tipo Día').reset_index(drop=True)
                    # Construir df_energia_prioridad para el período
                    energia_fechas_total = []
                    for fecha in df_res_f['Fecha'].unique():
                        if not df_f_d.empty and fecha in df_f_d['Fecha'].values:
                            row = df_f_d[df_f_d['Fecha'] == fecha].iloc[0]
                            energia_fechas_total.append({'Fecha': fecha, 'E_Total': row['Consumo Horario [kWh]'], 'E_Tr': row['E_Tr'], 'E_12': row['E_12'], 'Fuente': 'Factura'})
                        elif not df_p_d.empty and fecha in df_p_d['Fecha'].values:
                            row = df_p_d[df_p_d['Fecha'] == fecha].iloc[0]
                            energia_fechas_total.append({'Fecha': fecha, 'E_Total': row['Energía PRMTE [kWh]'], 'E_Tr': row['E_Tr'], 'E_12': row['E_12'], 'Fuente': 'PRMTE'})
                        elif not df_seat.empty and fecha in df_seat['Fecha'].values:
                            row = df_seat[df_seat['Fecha'] == fecha].iloc[0]
                            energia_fechas_total.append({'Fecha': fecha, 'E_Total': row['Total [kWh]'], 'E_Tr': row['Tracción [kWh]'], 'E_12': row['12 KV [kWh]'], 'Fuente': 'SEAT'})
                        else:
                            energia_fechas_total.append({'Fecha': fecha, 'E_Total': 0, 'E_Tr': 0, 'E_12': 0, 'Fuente': 'Sin datos'})
                    df_energia_total = pd.DataFrame(energia_fechas_total)
                    excel_data = exportar_resumen_excel(
                        metrics_dict=metrics_dict,
                        df_resumen_jornada=res_j_total,
                        df_energia=df_energia_total,
                        df_datos_semanales=None
                    )
                    st.download_button("⬇️ Descargar XLSX", excel_data, "Resumen_EFE.xlsx", use_container_width=True)
        else:
            st.info("No hay datos con los filtros seleccionados.")
    else:
        st.info("No hay datos de operaciones cargados.")

with tabs[1]: # Operaciones
    if not df_ops.empty:
        st.write("#### Filtros de Operaciones")
        col_o1, col_o2, col_o3 = st.columns(3)
        anios_op = sorted(df_ops['Fecha'].dt.year.unique())
        f_ano_op = col_o1.multiselect("Año", anios_op, default=anios_op, key="op_ano")
        meses_op = sorted(df_ops['Fecha'].dt.month.unique())
        f_mes_op = col_o2.multiselect("Mes", meses_op, default=meses_op, key="op_mes")
        tipos_op = df_ops['Tipo Día'].unique()
        orden_tipos_op = [d for d in ORDEN_TIPO_DIA if d in tipos_op]
        f_tipo_op = col_o3.multiselect("Tipo Día", orden_tipos_op, default=orden_tipos_op, key="op_tipo")
        
        mask_op = (df_ops['Fecha'].dt.year.isin(f_ano_op)) & (df_ops['Fecha'].dt.month.isin(f_mes_op))
        if f_tipo_op:
            mask_op &= df_ops['Tipo Día'].isin(f_tipo_op)
        df_ops_f = df_ops[mask_op].copy()
        
        for col in ['E_Total', 'E_Tr', 'E_12', 'Fuente']:
            if col not in df_ops_f.columns:
                df_ops_f[col] = 0
        if 'IDE (kWh/km)' not in df_ops_f.columns:
            df_ops_f['IDE (kWh/km)'] = 0
        
        columnas_mostrar = [
            'Fecha', 'Tipo Día', 'N° Semana', 
            'Odómetro [km]', 'Tren-Km [km]', 'UMR [%]', 
            'E_Total', 'E_Tr', 'E_12', 'IDE (kWh/km)', 'Fuente'
        ]
        
        st.dataframe(df_ops_f[columnas_mostrar].style.format({
            'Odómetro [km]': "{:,.1f}",
            'Tren-Km [km]': "{:,.1f}",
            'UMR [%]': "{:.2f}%",
            'E_Total': "{:,.0f}",
            'E_Tr': "{:,.0f}",
            'E_12': "{:,.0f}",
            'IDE (kWh/km)': "{:.4f}"
        }), use_container_width=True)
        
        st.download_button("📥 Descargar Operaciones (PPTX)", to_pptx("Datos Operacionales", df_ops_f[columnas_mostrar]), "EFE_Operaciones.pptx")
    else:
        st.info("No hay datos de operaciones para mostrar.")

with tabs[2]: # Trenes
    if not df_tr.empty or not df_tr_acum.empty:
        st.write("#### Filtros Trenes")
        df_tr_comb = pd.concat([df_tr, df_tr_acum])
        c1, c2 = st.columns(2)
        meses_tr = sorted(df_tr_comb['Fecha'].dt.month.unique())
        trenes_tr = sorted(df_tr_comb['Tren'].unique())
        f_mes_tr = c1.multiselect("Mes", meses_tr, default=meses_tr, key="tr_m")
        f_tren_tr = c2.multiselect("Tren(es)", trenes_tr, key="tr_t")
        
        if not df_tr.empty:
            st.write("### 🚗 Kilometraje Diario [km]")
            df_tr_f = df_tr[df_tr['Fecha'].dt.month.isin(f_mes_tr)]
            if f_tren_tr: df_tr_f = df_tr_f[df_tr_f['Tren'].isin(f_tren_tr)]
            if not df_tr_f.empty:
                piv_diario = df_tr_f.pivot_table(index="Tren", columns=df_tr_f["Fecha"].dt.day, values="Valor", aggfunc='sum').fillna(0)
                st.dataframe(piv_diario.style.format("{:,.1f}"), use_container_width=True)
                st.download_button("📥 Descargar Kilometraje (PPTX)", to_pptx("Kilometraje Diario Trenes", piv_diario.reset_index()), "EFE_Kilometraje.pptx")
        
        if not df_tr_acum.empty:
            st.divider()
            st.write("### 📈 Lectura de Odómetro / Acumulado [km]")
            df_tra_f = df_tr_acum[df_tr_acum['Fecha'].dt.month.isin(f_mes_tr)]
            if f_tren_tr: df_tra_f = df_tra_f[df_tra_f['Tren'].isin(f_tren_tr)]
            if not df_tra_f.empty:
                piv_acum = df_tra_f.pivot_table(index="Tren", columns=df_tra_f["Fecha"].dt.day, values="Valor", aggfunc='max').fillna(0)
                st.dataframe(piv_acum.style.format("{:,.0f}"), use_container_width=True)
                st.download_button("📥 Descargar Acumulados (PPTX)", to_pptx("Odómetro Acumulado", piv_acum.reset_index()), "EFE_Acumulados.pptx")
    else:
        st.info("No hay datos de trenes cargados.")

with tabs[3]: # Energía
    st.write("#### ⚡ Módulo de Medición")
    sub_e = st.tabs(["⚡ SEAT", "📈 PRMTE", "💰 Facturación"])
    with sub_e[0]:
        if not df_seat.empty:
            c1, c2 = st.columns(2)
            anios_s = sorted(df_seat['Fecha'].dt.year.unique())
            meses_s = sorted(df_seat['Fecha'].dt.month.unique())
            f_ano_s = c1.multiselect("Año SEAT", anios_s, default=anios_s, key="seat_a")
            f_mes_s = c2.multiselect("Mes SEAT", meses_s, default=meses_s, key="seat_m")
            mask = df_seat['Fecha'].dt.year.isin(f_ano_s) & df_seat['Fecha'].dt.month.isin(f_mes_s)
            df_s_f = df_seat[mask]
            st.dataframe(df_s_f, use_container_width=True)
            st.download_button("📥 Descargar SEAT (PPTX)", to_pptx("Energía SEAT", df_s_f), "EFE_SEAT.pptx")
        else:
            st.info("No hay datos SEAT cargados.")
    with sub_e[1]:
        if not df_p_d.empty:
            c1, c2 = st.columns(2)
            anios_p = sorted(df_p_d['Fecha'].dt.year.unique())
            meses_p = sorted(df_p_d['Fecha'].dt.month.unique())
            f_ano_p = c1.multiselect("Año PRMTE", anios_p, default=anios_p, key="prm_a")
            f_mes_p = c2.multiselect("Mes PRMTE", meses_p, default=meses_p, key="prm_m")
            mask = df_p_d['Fecha'].dt.year.isin(f_ano_p) & df_p_d['Fecha'].dt.month.isin(f_mes_p)
            df_p_f = df_p_d[mask]
            st.dataframe(df_p_f, use_container_width=True)
            st.download_button("📥 Descargar PRMTE (PPTX)", to_pptx("Medidas PRMTE", df_p_f), "EFE_PRMTE.pptx")
        else:
            st.info("No hay datos PRMTE cargados.")
    with sub_e[2]:
        if not df_f_d.empty:
            c1, c2 = st.columns(2)
            anios_f = sorted(df_f_d['Fecha'].dt.year.unique())
            meses_f = sorted(df_f_d['Fecha'].dt.month.unique())
            f_ano_f = c1.multiselect("Año Factura", anios_f, default=anios_f, key="fact_a")
            f_mes_f = c2.multiselect("Mes Factura", meses_f, default=meses_f, key="fact_m")
            mask = df_f_d['Fecha'].dt.year.isin(f_ano_f) & df_f_d['Fecha'].dt.month.isin(f_mes_f)
            df_f_f = df_f_d[mask]
            st.dataframe(df_f_f, use_container_width=True)
            st.download_button("📥 Descargar Facturación (PPTX)", to_pptx("Facturación", df_f_f), "EFE_Facturacion.pptx")
        else:
            st.info("No hay datos de facturación cargados.")

with tabs[4]: # Comparación hr
    if all_comp_full:
        df_c = pd.DataFrame(all_comp_full).groupby(['Fecha','Hora','Fuente'])['Consumo Horario [kWh]'].sum().reset_index()
        fechas_f = df_c[df_c['Fuente']=='Factura']['Fecha'].unique()
        df_cf = df_c[~((df_c['Fuente']=='PRMTE') & (df_c['Fecha'].isin(fechas_f)))].copy()
        df_cf['Año'], df_cf['Tipo Día'] = df_cf['Fecha'].dt.year, df_cf['Fecha'].apply(get_tipo_dia)
        st.write("#### Mediana de Consumo 2025 vs 2026")
        df_st = df_cf[df_cf['Año'].isin([2025, 2026])]
        if not df_st.empty:
            pivot_st = df_st.pivot_table(index="Hora", columns=["Año", "Tipo Día"], values="Consumo Horario [kWh]", aggfunc='median', observed=False).fillna(0)
            st.dataframe(pivot_st.style.format("{:,.1f}"), use_container_width=True)
            st.download_button("📥 Descargar Comparativa (PPTX)", to_pptx("Comparación Energía por hr", pivot_st.reset_index()), "EFE_Comparativa.pptx")
        else:
            st.info("No hay datos suficientes para la comparación.")
    else:
        st.info("No hay datos de consumo horario cargados.")

if 'outliers' not in st.session_state: st.session_state.outliers = pd.DataFrame()

with tabs[5]: # Regresión
    if all_comp_full:
        df_reg = pd.DataFrame(all_comp_full).groupby(['Fecha','Hora','Fuente'])['Consumo Horario [kWh]'].sum().reset_index()
        fechas_f = df_reg[df_reg['Fuente']=='Factura']['Fecha'].unique()
        df_reg = df_reg[~((df_reg['Fuente']=='PRMTE') & (df_reg['Fecha'].isin(fechas_f)))].copy()
        df_reg = df_reg[df_reg['Hora']<=5]
        df_reg['Año'], df_reg['Tipo Día'] = df_reg['Fecha'].dt.year, df_reg['Fecha'].apply(get_tipo_dia)
        
        c1, c2, c3 = st.columns(3)
        f_ra, f_rj, f_rh = c1.selectbox("Año", sorted(df_reg['Año'].unique()), key="reg_a"), c2.selectbox("Jornada", ['Total', 'L', 'S', 'D/F'], key="reg_j"), c3.selectbox("Hora", range(6), key="reg_h")
        
        df_pl = df_reg[(df_reg['Año']==f_ra) & (df_reg['Hora']==f_rh)]
        if f_rj != 'Total': df_pl = df_pl[df_pl['Tipo Día']==f_rj]
        df_pl = df_pl.sort_values('Fecha')

        if len(df_pl) > 1:
            Q1, Q3 = df_pl['Consumo Horario [kWh]'].quantile(0.25), df_pl['Consumo Horario [kWh]'].quantile(0.75)
            IQR = Q3 - Q1
            lim_sup, lim_inf = Q3 + 1.5*IQR, Q1 - 1.5*IQR
            df_norm = df_pl[(df_pl['Consumo Horario [kWh]']>=lim_inf) & (df_pl['Consumo Horario [kWh]']<=lim_sup)].copy()
            st.session_state.outliers = df_pl[(df_pl['Consumo Horario [kWh]']<lim_inf) | (df_pl['Consumo Horario [kWh]']>lim_sup)].copy()
            
            if len(df_norm) > 1:
                x, y = np.arange(len(df_norm)), df_norm['Consumo Horario [kWh]'].values
                m, n = np.polyfit(x, y, 1)
                r2 = 1 - (np.sum((y - (m*x+n))**2) / np.sum((y - np.mean(y))**2))
                st.line_chart(pd.DataFrame({'Real': y, 'Tendencia': m*x+n}, index=df_norm['Fecha'].dt.strftime('%d/%m')))
                st.markdown(f"**Ecuación:** $Consumo = {m:.4f}x + {n:.2f}$ | $R^2 = {r2:.4f}$")
                st.info(f"Instalación basal inicial: {n:.2f} kWh. Variación cronológica: {m:.4f} kWh por hora.")
                m_reg = {"Ecuación": f"Consumo = {m:.4f}x + {n:.2f}", "R2": f"{r2:.4f}", "Total Limpio": f"{y.sum():,.1f} kWh"}
                st.download_button("📥 Descargar Regresión (PPTX)", to_pptx(f"Regresión Nocturna - Hora {f_rh}", df_norm[['Fecha','Consumo Horario [kWh]']], m_reg), "EFE_Regresion.pptx")
            else:
                st.warning("No hay suficientes datos limpios para la regresión.")
        else:
            st.warning("Se necesitan al menos 2 puntos para la regresión.")
    else:
        st.info("No hay datos de consumo horario cargados para regresión.")

with tabs[6]: # Atípicos
    if not st.session_state.outliers.empty:
        st.error(f"Se detectaron {len(st.session_state.outliers)} anomalías.")
        st.dataframe(st.session_state.outliers, use_container_width=True)
        csv = st.session_state.outliers.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Descargar CSV", csv, "Anomalias.csv", "text/csv")
        st.download_button("📥 Descargar Atípicos (PPTX)", to_pptx("Datos Atípicos de Instalaciones", st.session_state.outliers), "EFE_Atipicos.pptx")
    else:
        st.success("No hay anomalías detectadas en la selección actual.")

with tabs[7]: # THDR
    st.header("📋 Datos THDR - Vía 1 y Vía 2")
    if not df_total.empty:
        st.write("#### Filtros rápidos")
        col1, col2 = st.columns(2)
        servicios = sorted(df_total['Servicio'].unique())
        servicio_sel = col1.multiselect("Filtrar por Servicio", servicios, default=[])
        motrices = sorted(df_total['Motriz 1'].unique())
        motriz_sel = col2.multiselect("Filtrar por Motriz 1", motrices, default=[])
        
        df_filtrado = df_total.copy()
        if servicio_sel:
            df_filtrado = df_filtrado[df_filtrado['Servicio'].isin(servicio_sel)]
        if motriz_sel:
            df_filtrado = df_filtrado[df_filtrado['Motriz 1'].isin(motriz_sel)]
        
        st.subheader("🟢 Vía 1")
        df_v1 = df_filtrado[df_filtrado['Vía'] == 'Vía 1'].copy()
        if not df_v1.empty:
            df_v1['Hora_Salida_Str'] = df_v1['Hora_Salida'].apply(lambda x: f"{int(x//60):02d}:{int(x%60):02d}" if pd.notna(x) else "")
            df_v1['Retraso_Str'] = df_v1['Retraso'].apply(lambda x: format_hms(x, con_signo=True))
            df_v1['Puntual_Str'] = df_v1['Puntual'].apply(lambda x: "Sí" if x == 1 else "No")
            df_v1_display = df_v1[['Fecha_Op', 'Servicio', 'Hora_Prog', 'Hora_Salida_Str', 'Motriz 1', 'Motriz 2', 'Unidad', 'Tipo_Rec', 'Tren-Km', 'Retraso_Str', 'Puntual_Str']]
            df_v1_display.columns = ['Fecha', 'Servicio', 'Hora Prog', 'Hora Real', 'Motriz 1', 'Motriz 2', 'Unidad', 'Recorrido', 'Tren-Km', 'Retraso (PS)', 'Puntual']
            st.dataframe(df_v1_display.style.format({'Tren-Km': "{:.1f}"}), use_container_width=True)
        else:
            st.info("No hay datos para Vía 1 con los filtros seleccionados.")
        
        st.subheader("🔵 Vía 2")
        df_v2 = df_filtrado[df_filtrado['Vía'] == 'Vía 2'].copy()
        if not df_v2.empty:
            df_v2['Hora_Salida_Str'] = df_v2['Hora_Salida'].apply(lambda x: f"{int(x//60):02d}:{int(x%60):02d}" if pd.notna(x) else "")
            df_v2['Retraso_Str'] = df_v2['Retraso'].apply(lambda x: format_hms(x, con_signo=True))
            df_v2['Puntual_Str'] = df_v2['Puntual'].apply(lambda x: "Sí" if x == 1 else "No")
            df_v2_display = df_v2[['Fecha_Op', 'Servicio', 'Hora_Prog', 'Hora_Salida_Str', 'Motriz 1', 'Motriz 2', 'Unidad', 'Tipo_Rec', 'Tren-Km', 'Retraso_Str', 'Puntual_Str']]
            df_v2_display.columns = ['Fecha', 'Servicio', 'Hora Prog', 'Hora Real', 'Motriz 1', 'Motriz 2', 'Unidad', 'Recorrido', 'Tren-Km', 'Retraso (PS)', 'Puntual']
            st.dataframe(df_v2_display.style.format({'Tren-Km': "{:.1f}"}), use_container_width=True)
        else:
            st.info("No hay datos para Vía 2 con los filtros seleccionados.")
    else:
        st.info("No hay datos de THDR cargados. Sube archivos de THDR Vía 1 y/o Vía 2.")

# --- 9. DESCARGA DE REPORTE EXCEL COMPLETO ---
st.sidebar.download_button("📥 Reporte Excel Completo", to_excel_consolidado(df_ops, df_tr, df_tr_acum, df_seat, df_p_d, pd.DataFrame(all_prmte_15), pd.DataFrame(all_fact_h), df_f_d), "Reporte_EFE_SGE.xlsx")
