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
        if len(val) == 6 and val.isdigit():
            return (int(val[0:2]), int(val[2:4]), 2000 + int(val[4:6]))
        # Intento por nombre de archivo
        match = re.search(r'(\d{2})[.-](\d{2})[.-](\d{2})', file.name)
        if match:
            return (int(match.group(1)), int(match.group(2)), 2000 + int(match.group(3)))
    except:
        pass
    return None

def procesar_thdr_avanzado(file):
    try:
        df_raw = pd.read_excel(file, header=None)
        est_h = df_raw.iloc[0].ffill().values
        df = df_raw.iloc[2:].copy()
        df = df[pd.to_numeric(df.iloc[:, 1], errors='coerce').notna()]

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
        df.columns = nombres_finales[:len(df.columns)]
        est_cols = df.columns[6:]

        def get_trip(row):
            h_reales = row[est_cols].apply(convertir_a_minutos).dropna()
            if len(h_reales) < 2:
                return "OTRO", "OTRO", None, 0, 0
            def cod(n_est):
                n = str(n_est).upper()
                if "PUERTO" in n: return "PU"
                if "LIMACHE" in n: return "LI"
                if "VIÑA" in n or "VINA" in n: return "VM"
                if "BELLOTO" in n: return "EB"
                return n[:2]
            t_s, t_l = h_reales.iloc[0], h_reales.iloc[-1]
            if t_l < t_s: t_l += 1440
            return (f"{cod(h_reales.index[0])}-{cod(h_reales.index[-1])}",
                    cod(h_reales.index[0]), t_s, (t_l - t_s), int(t_s // 60) % 24)

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
        f_info = leer_fecha_archivo(file)
        df['Fecha_Op'] = f"{f_info[0]:02d}/{f_info[1]:02d}/{f_info[2]}" if f_info else None
        return df, df['Tren-Km'].sum(), df[df['TDV_Min'] > 0]['TDV_Min'].mean(), (df['Puntual'].sum() / len(df) * 100) if len(df) > 0 else 0
    except Exception as e:
        st.error(f"Error procesando THDR: {e}")
        return pd.DataFrame(), 0, 0, 0

# --- 4. INICIALIZACIÓN DE DATAFRAMES VACÍOS ---
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

# --- 6. LECTURA Y PROCESAMIENTO DE DATOS ---
if f_v1 or f_v2 or f_umr or f_seat_files or f_bill_files:
    all_ops, all_tr, all_tr_acum, all_seat, all_prmte_15, all_fact_h, all_comp_full = [], [], [], [], [], [], []
    thdr_v1_list, thdr_v2_list = [], []
    
    if f_v1:
        for file in f_v1:
            df, _, _, _ = procesar_thdr_avanzado(file)
            if not df.empty: df['Vía'] = 'Vía 1'; thdr_v1_list.append(df)
        if thdr_v1_list: df_thdr_v1 = pd.concat(thdr_v1_list, ignore_index=True)
    
    if f_v2:
        for file in f_v2:
            df, _, _, _ = procesar_thdr_avanzado(file)
            if not df.empty: df['Vía'] = 'Vía 2'; thdr_v2_list.append(df)
        if thdr_v2_list: df_thdr_v2 = pd.concat(thdr_v2_list, ignore_index=True)

    todos = (f_umr or []) + (f_seat_files or []) + (f_bill_files or [])
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
                        idx_f, idx_o, idx_t = next((c for c in df_p.columns if 'FECHA' in c), None), next((c for c in df_p.columns if 'ODO' in c and 'ACUM' not in c), None), next((c for c in df_p.columns if 'TREN' in c and 'KM' in c), None)
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
                            ts, val_p = r['Timestamp'], sum([parse_latam_number(r[col]) for col in cols_e])
                            all_comp_full.append({"Fecha": ts.normalize(), "Hora": ts.hour, "Consumo Horario [kWh]": val_p, "Fuente": "PRMTE"})
                            if start_date <= ts.date() <= end_date: all_prmte_15.append({"Fecha y Hora": ts.strftime('%d/%m/%Y %H:%M'), "Fecha": ts.normalize(), "Energía PRMTE [kWh]": val_p})

                if any(k in sn_up for k in ['FACTURA', 'CONSUMO']):
                    df_f = pd.read_excel(f, sheet_name=sn); df_f.columns = ['FechaHora', 'Valor']
                    df_f['Timestamp'] = pd.to_datetime(df_f['FechaHora'], errors='coerce')
                    for _, r in df_f.dropna(subset=['Timestamp']).iterrows():
                        ts, val_f = r['Timestamp'], abs(parse_latam_number(r['Valor']))
                        all_comp_full.append({"Fecha": ts.normalize(), "Hora": ts.hour, "Consumo Horario [kWh]": val_f, "Fuente": "Factura"})
                        if start_date <= ts.date() <= end_date: all_fact_h.append({"Fecha y Hora": ts.strftime('%d/%m/%Y %H:%M'), "Fecha": ts.normalize(), "Consumo Horario [kWh]": val_f})
        except: continue

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

# --- 7. DASHBOARD ---
tabs = st.tabs(["📊 Resumen", "📑 Operaciones", "📑 Trenes", "⚡ Energía", "⚖️ Comparación Energía hr", "📈 Regresión Nocturna", "🚨 Datos Atípicos", "📋 THDR"])

# ================== PESTAÑA RESUMEN ==================
with tabs[0]:
    if not df_ops.empty:
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
            if st.session_state.filtros_compartidos['anios']: mask &= df['Fecha'].dt.year.isin(st.session_state.filtros_compartidos['anios'])
            if st.session_state.filtros_compartidos['meses']: mask &= df['Fecha'].dt.month.isin(st.session_state.filtros_compartidos['meses'])
            if 'N° Semana' in df.columns and st.session_state.filtros_compartidos['semanas']: mask &= df['N° Semana'].isin(st.session_state.filtros_compartidos['semanas'])
            if 'Tipo Día' in df.columns and st.session_state.filtros_compartidos['jornadas']: mask &= df['Tipo Día'].isin(st.session_state.filtros_compartidos['jornadas'])
            return df[mask]

        mostrar_filtros_compartidos(df_ops)
        df_res_f = aplicar_filtros_compartidos(df_ops)
        if not df_res_f.empty:
            to_val, tk_val = df_res_f["Odómetro [km]"].sum(), df_res_f["Tren-Km [km]"].sum()
            umr_val = (tk_val/to_val*100) if to_val>0 else 0
            
            sub_tabs = st.tabs(["📅 Semanal", "📅 Mensual", "📅 Anual"])
            with sub_tabs[0]:
                st.write("##### Evolución Semanal")
                col_s1, col_s2, col_s3 = st.columns(3)
                anios_sem = sorted(df_res_f['Fecha'].dt.year.unique())
                f_ano_sem = col_s1.selectbox("Año (semana)", anios_sem, key="sem_ano")
                semanas_df = df_res_f[df_res_f['Fecha'].dt.year == f_ano_sem]['N° Semana'].unique()
                f_semana = col_s2.selectbox("N° Semana", sorted(semanas_df), key="sem_num")
                tipos_sem = df_res_f['Tipo Día'].unique()
                f_tipo_sem = col_s3.multiselect("Tipo Día (semana)", [d for d in ORDEN_TIPO_DIA if d in tipos_sem], default=[d for d in ORDEN_TIPO_DIA if d in tipos_sem], key="sem_tipo")
                df_semana = df_res_f[(df_res_f['Fecha'].dt.year == f_ano_sem) & (df_res_f['N° Semana'] == f_semana)]
                if f_tipo_sem: df_semana = df_semana[df_semana['Tipo Día'].isin(f_tipo_sem)]
                
                if not df_semana.empty:
                    col_m1, col_m2, col_m3 = st.columns(3)
                    col_m1.metric("Odómetro", f"{df_semana['Odómetro [km]'].sum():,.1f} km")
                    col_m2.metric("Tren-Km", f"{df_semana['Tren-Km [km]'].sum():,.1f} km")
                    col_m3.metric("UMR", f"{(df_semana['Tren-Km [km]'].sum()/df_semana['Odómetro [km]'].sum()*100 if df_semana['Odómetro [km]'].sum()>0 else 0):.2f} %")
                    fig = make_subplots(specs=[[{"secondary_y": True}]])
                    fig.add_trace(go.Bar(x=df_semana['Fecha'].dt.strftime('%d/%m'), y=df_semana['Odómetro [km]']/1000, name='Odómetro (kkm)', marker_color='#005195'), secondary_y=False)
                    fig.add_trace(go.Bar(x=df_semana['Fecha'].dt.strftime('%d/%m'), y=df_semana['Tren-Km [km]']/1000, name='Tren-Km (kkm)', marker_color='#4CAF50'), secondary_y=False)
                    fig.add_trace(go.Scatter(x=df_semana['Fecha'].dt.strftime('%d/%m'), y=df_semana['UMR [%]'], name='UMR (%)', mode='lines+markers', line=dict(color='#FF5733')), secondary_y=True)
                    st.plotly_chart(fig, use_container_width=True)

            with sub_tabs[1]:
                st.write("##### Evolución Mensual")
                col_m1, col_m2, col_m3 = st.columns(3)
                anios_mes = sorted(df_res_f['Fecha'].dt.year.unique())
                f_ano_mes = col_m1.selectbox("Año (mensual)", anios_mes, key="mes_ano")
                meses_mes = sorted(df_res_f[df_res_f['Fecha'].dt.year == f_ano_mes]['Fecha'].dt.month.unique())
                f_mes_mes = col_m2.selectbox("Mes", meses_mes, format_func=lambda x: f"{x:02d}", key="mes_num")
                tipos_mes = df_res_f['Tipo Día'].unique()
                f_tipo_mes = col_m3.multiselect("Tipo Día (mensual)", [d for d in ORDEN_TIPO_DIA if d in tipos_mes], default=[d for d in ORDEN_TIPO_DIA if d in tipos_mes], key="mes_tipo")
                df_mes = df_res_f[(df_res_f['Fecha'].dt.year == f_ano_mes) & (df_res_f['Fecha'].dt.month == f_mes_mes)]
                if f_tipo_mes: df_mes = df_mes[df_mes['Tipo Día'].isin(f_tipo_mes)]
                
                if not df_mes.empty:
                    st.columns(3)[0].metric("Odómetro Mes", f"{df_mes['Odómetro [km]'].sum():,.1f} km")
                    fig = make_subplots(specs=[[{"secondary_y": True}]])
                    fig.add_trace(go.Bar(x=df_mes['Fecha'].dt.strftime('%d/%m'), y=df_mes['Odómetro [km]']/1000, name='Odómetro', marker_color='#005195'), secondary_y=False)
                    fig.add_trace(go.Scatter(x=df_mes['Fecha'].dt.strftime('%d/%m'), y=df_mes['UMR [%]'], name='UMR (%)', mode='lines+markers', line=dict(color='#FF5733')), secondary_y=True)
                    st.plotly_chart(fig, use_container_width=True)

            with sub_tabs[2]:
                st.write("##### Evolución Anual")
                col_a1, col_a2 = st.columns(2)
                anios_anual = sorted(df_res_f['Fecha'].dt.year.unique())
                f_ano_anual = col_a1.selectbox("Año (anual)", anios_anual, key="anual_ano")
                df_anual = df_res_f[df_res_f['Fecha'].dt.year == f_ano_anual].copy()
                if not df_anual.empty:
                    df_mensual = df_anual.groupby(df_anual['Fecha'].dt.month).agg({'Odómetro [km]': 'sum', 'Tren-Km [km]': 'sum', 'UMR [%]': 'mean'}).reset_index()
                    fig = make_subplots(specs=[[{"secondary_y": True}]])
                    fig.add_trace(go.Bar(x=df_mensual['Fecha'], y=df_mensual['Odómetro [km]']/1000, name='Odómetro (kkm)'), secondary_y=False)
                    fig.add_trace(go.Scatter(x=df_mensual['Fecha'], y=df_mensual['UMR [%]'], name='UMR (%)', mode='lines+markers'), secondary_y=True)
                    st.plotly_chart(fig, use_container_width=True)

            st.write("---")
            st.write("#### 📥 Exportar pestaña Resumen")
            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                if st.button("🖨️ Imprimir / PDF", use_container_width=True): st.markdown("<script>window.print();</script>", unsafe_allow_html=True)
            with col_btn2:
                if st.button("📈 Exportar a XLSX", use_container_width=True):
                    m_dict = {"Odómetro": to_val, "Tren-Km": tk_val, "UMR": umr_val}
                    st.download_button("⬇️ Descargar", exportar_resumen_excel(m_dict, df_res_f, df_res_f), "Resumen_EFE.xlsx")
        else:
            st.info("No hay datos con los filtros seleccionados.")
    else:
        st.info("No hay datos de operaciones cargados.")

# ================== PESTAÑA OPERACIONES ==================
with tabs[1]:
    if not df_ops.empty:
        st.write("#### Detalle de Operaciones")
        st.dataframe(df_ops.style.format({'Odómetro [km]': "{:,.1f}", 'Tren-Km [km]': "{:,.1f}", 'UMR [%]': "{:.2f}%", 'E_Total': "{:,.0f}", 'IDE (kWh/km)': "{:.4f}"}), use_container_width=True)
        st.download_button("📥 PPTX", to_pptx("Datos Operacionales", df_ops), "Operaciones.pptx")
    else:
        st.info("No hay datos.")

# ================== PESTAÑA TRENES ==================
with tabs[2]:
    if not df_tr.empty or not df_tr_acum.empty:
        if not df_tr.empty:
            st.write("### 🚗 Kilometraje Diario")
            piv_diario = df_tr.pivot_table(index="Tren", columns=df_tr["Fecha"].dt.day, values="Valor", aggfunc='sum').fillna(0)
            st.dataframe(piv_diario.style.format("{:,.1f}"), use_container_width=True)
        if not df_tr_acum.empty:
            st.write("### 📈 Lectura de Odómetro")
            piv_acum = df_tr_acum.pivot_table(index="Tren", columns=df_tr_acum["Fecha"].dt.day, values="Valor", aggfunc='max').fillna(0)
            st.dataframe(piv_acum.style.format("{:,.0f}"), use_container_width=True)

# ================== PESTAÑA ENERGÍA ==================
with tabs[3]:
    sub_e = st.tabs(["⚡ SEAT", "📈 PRMTE", "💰 Facturación"])
    with sub_e[0]:
        if not df_seat.empty: st.dataframe(df_seat)
    with sub_e[1]:
        if not df_p_d.empty: st.dataframe(df_p_d)
    with sub_e[2]:
        if not df_f_d.empty: st.dataframe(df_f_d)

# ================== PESTAÑA COMPARACIÓN HR ==================
with tabs[4]:
    if all_comp_full:
        df_c = pd.DataFrame(all_comp_full).groupby(['Fecha','Hora','Fuente'])['Consumo Horario [kWh]'].sum().reset_index()
        pivot_st = df_c.pivot_table(index="Hora", columns=df_c['Fecha'].dt.year, values="Consumo Horario [kWh]", aggfunc='median').fillna(0)
        st.write("#### Mediana de Consumo Horario")
        st.line_chart(pivot_st)
        st.dataframe(pivot_st.style.format("{:,.1f}"))

# ================== PESTAÑA REGRESIÓN ==================
if 'outliers' not in st.session_state: st.session_state.outliers = pd.DataFrame()
with tabs[5]:
    if all_comp_full:
        df_reg = pd.DataFrame(all_comp_full).groupby(['Fecha','Hora','Fuente'])['Consumo Horario [kWh]'].sum().reset_index()
        df_reg = df_reg[df_reg['Hora']<=5]
        df_pl = df_reg.sort_values('Fecha')
        if len(df_pl) > 2:
            x, y = np.arange(len(df_pl)), df_pl['Consumo Horario [kWh]'].values
            m, n = np.polyfit(x, y, 1)
            st.line_chart(pd.DataFrame({'Real': y, 'Tendencia': m*x+n}))
            st.markdown(f"**Ecuación:** $y = {m:.4f}x + {n:.2f}$")

# ================== PESTAÑA DATOS ATÍPICOS ==================
with tabs[6]:
    if not st.session_state.outliers.empty:
        st.dataframe(st.session_state.outliers)
    else:
        st.success("No se detectan anomalías.")

# ================== PESTAÑA THDR (REVISADA Y COMPLETA) ==================
with tabs[7]:
    st.header("📋 Datos THDR - Vía 1 y Vía 2")
    
    def mostrar_tabla_thdr_v2(df, titulo, color_emoji):
        st.subheader(f"{color_emoji} {titulo}")
        if df.empty:
            st.info(f"No hay datos para {titulo}. Sube archivos en la barra lateral.")
            return
        
        df_calc = df.copy()
        # Formatear tiempos
        df_calc['H_Real_Salida'] = df_calc['Min_S_Real'].apply(lambda x: format_hms(x) if pd.notna(x) else "")
        df_calc['H_Programada'] = df_calc['Min_Prog'].apply(lambda x: format_hms(x) if pd.notna(x) else "")
        df_calc['Dif_Puntualidad'] = df_calc['Retraso'].apply(lambda x: format_hms(x, con_signo=True) if pd.notna(x) else "")
        df_calc['Llegada_Terminal'] = (df_calc['Min_S_Real'] + df_calc['TDV_Min']).apply(lambda x: format_hms(x) if pd.notna(x) else "")
        
        df_display = df_calc[['Fecha_Op', 'Servicio', 'H_Programada', 'H_Real_Salida', 'Dif_Puntualidad', 'Llegada_Terminal', 'Motriz 1', 'Tipo_Rec', 'Tren-Km']]
        st.dataframe(df_display, use_container_width=True)
    
    mostrar_tabla_thdr_v2(df_thdr_v1, "Vía 1", "🟢")
    mostrar_tabla_thdr_v2(df_thdr_v2, "Vía 2", "🔵")

# --- 8. DESCARGA DE REPORTE EXCEL COMPLETO ---
st.sidebar.download_button("📥 Reporte Excel Completo", to_excel_consolidado(df_ops, df_tr, df_tr_acum, df_seat, df_p_d, pd.DataFrame(all_prmte_15), pd.DataFrame(all_fact_h), df_f_d), "Reporte_EFE_SGE.xlsx")
