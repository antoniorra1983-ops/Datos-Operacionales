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

# --- 2. FUNCIONES DE APOYO Y EXPORTACIÓN ---
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

def format_hms(m, con_signo=False):
    if pd.isna(m) or m == 0: return "00:00:00"
    signo = ("+" if m > 0 else "-" if m < 0 else "") if con_signo else ""
    sec = int(round(abs(m) * 60))
    h, r = divmod(sec, 3600); mi, se = divmod(r, 60)
    return f"{signo}{h:02d}:{mi:02d}:{se:02d}"

def to_pptx(title_text, df=None, metrics_dict=None):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = f"EFE Valparaíso: {title_text}"
    y_cursor = Inches(1.5)
    if metrics_dict:
        txBox = slide.shapes.add_textbox(Inches(0.5), y_cursor, Inches(9), Inches(1))
        tf = txBox.text_frame
        for k, v in metrics_dict.items():
            p = tf.add_paragraph()
            p.text = f"• {k}: {v}"; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = RGBColor(0, 81, 149)
        y_cursor += Inches(1.2)
    if df is not None and not df.empty:
        df_d = df.head(12).reset_index(drop=True)
        rows, cols = df_d.shape
        table = slide.shapes.add_table(rows + 1, cols, Inches(0.5), y_cursor, Inches(9), Inches(3)).table
        for c, col in enumerate(df_d.columns):
            cell = table.cell(0, c); cell.text = str(col); cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(0, 81, 149)
            p = cell.text_frame.paragraphs[0]; p.font.color.rgb = RGBColor(255, 255, 255); p.font.size = Pt(10); p.font.bold = True
        for r in range(rows):
            for c in range(cols):
                val = df_d.iloc[r, c]
                table.cell(r + 1, c).text = f"{val:,.1f}" if isinstance(val, float) else str(val)
                table.cell(r + 1, c).text_frame.paragraphs[0].font.size = Pt(9)
    out = BytesIO(); prs.save(out); return out.getvalue()

def to_excel_consolidado(df_ops, df_tr, df_tr_acum, df_seat, df_p_d, df_fact_d):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dfs = {'Operaciones': df_ops, 'Kms_Diarios_Tren': df_tr, 'Odometros_Acum_Tren': df_tr_acum, 'SEAT': df_seat, 'PRMTE_D': df_p_d, 'Factura_D': df_fact_d}
        for name, df in dfs.items():
            if not df.empty: df.to_excel(writer, index=False, sheet_name=name)
    return output.getvalue()

# --- 3. PROCESAMIENTO THDR (DINÁMICO PARA SERVICIOS CORTOS) ---
DISTANCIAS = {"PU-LI": 43.13, "LI-PU": 43.13, "PU-SA": 29.11, "SA-PU": 29.11, "EB-PU": 25.40, "PU-EB": 25.40, "VM-LI": 34.03, "LI-VM": 34.03, "VM-PU": 9.10, "PU-VM": 9.10}

def procesar_thdr_avanzado(file):
    try:
        df_raw = pd.read_excel(file, header=None)
        h0 = df_raw.iloc[0].ffill().astype(str)
        h1 = df_raw.iloc[1].fillna('').astype(str)
        
        column_names_raw = []
        for i in range(len(h0)):
            base, sub = h0[i].strip(), h1[i].strip()
            column_names_raw.append(f"{base} ({sub})" if sub in ['Hora Llegada', 'Hora Salida'] else base)
        
        # Deduplicar columnas para PyArrow
        final_cols, counts = [], {}
        for name in column_names_raw:
            if name in counts:
                counts[name] += 1
                final_cols.append(f"{name}_{counts[name]}")
            else:
                counts[name] = 0
                final_cols.append(name)
        
        df = df_raw.iloc[2:].copy()
        df.columns = final_cols
        
        # Detección dinámica de Origen/Destino para servicios que no son de Puerto
        cols_salida = [c for c in df.columns if '(Hora Salida)' in c]
        cols_llegada = [c for c in df.columns if '(Hora Llegada)' in c]
        
        def get_journey_data(row):
            start_val, start_name, end_val, end_name = None, None, None, None
            for c in cols_salida:
                v = convertir_a_minutos(row[c])
                if v is not None: start_val, start_name = v, c.split('(')[0].strip(); break
            for c in reversed(cols_llegada):
                v = convertir_a_minutos(row[c])
                if v is not None: end_val, end_name = v, c.split('(')[0].strip(); break
            return pd.Series([start_val, start_name, end_val, end_name])

        df[['Hora_Salida_Real', 'Origen', 'Hora_Llegada_Real', 'Destino']] = df.apply(get_journey_data, axis=1)
        
        def find_c(ks):
            for c in df.columns:
                if any(k.lower() in c.lower() for k in ks): return c
            return None

        c_s, c_m1, c_m2, c_p = find_c(['Servicio', 'N°']), find_c(['Motriz 1', 'M1']), find_c(['Motriz 2', 'M2']), find_c(['Prog'])
        df['Servicio'] = pd.to_numeric(df[c_s], errors='coerce').fillna(0).astype(int) if c_s else 0
        df['Motriz 1'] = pd.to_numeric(df[c_m1], errors='coerce').fillna(0).astype(int) if c_m1 else 0
        df['Motriz 2'] = pd.to_numeric(df[c_m2], errors='coerce').fillna(0).astype(int) if c_m2 else 0
        df['Unidad'] = df['Motriz 2'].apply(lambda x: 'M' if x > 0 else 'S')
        df['Min_Prog'] = df[c_p].apply(convertir_a_minutos) if c_p else 0
        df['Retraso'] = df['Hora_Salida_Real'] - df['Min_Prog']
        df['Puntual'] = df['Retraso'].apply(lambda x: 1 if pd.notna(x) and abs(x) <= 5 else 0)
        df['TDV_Min'] = df.apply(lambda r: (r['Hora_Llegada_Real'] - r['Hora_Salida_Real'] + (1440 if (r['Hora_Llegada_Real'] or 0) < (r['Hora_Salida_Real'] or 0) else 0)) if pd.notna(r['Hora_Salida_Real']) and pd.notna(r['Hora_Llegada_Real']) else 0, axis=1)
        
        def calc_tren_km(r):
            o = str(r['Origen'])[:2].upper() if r['Origen'] else ""
            d = str(r['Destino'])[:2].upper() if r['Destino'] else ""
            map_est = {"PU":"PU", "VA":"PU", "LI":"LI", "VI":"VM", "EL":"EB"}
            key = f"{map_est.get(o, o)}-{map_est.get(d, d)}"
            return DISTANCIAS.get(key, 43.13) * (2 if r['Unidad'] == 'M' else 1)
        
        df['Tren-Km'] = df.apply(calc_tren_km, axis=1)
        try:
            val_f = str(df_raw.iloc[0, 0]).split('.')[0].strip().zfill(6)
            df['Fecha_Op'] = f"{val_f[0:2]}/{val_f[2:4]}/20{val_f[4:6]}"
        except: df['Fecha_Op'] = ""
        return df[df['Servicio'] > 0], df['Tren-Km'].sum()
    except Exception as e:
        st.error(f"Error THDR: {e}"); return pd.DataFrame(), 0

# --- 4. INICIALIZACIÓN ---
df_ops = df_tr = df_tr_acum = df_seat = df_p_d = df_f_d = df_thdr_v1 = df_thdr_v2 = pd.DataFrame()
all_ops, all_tr, all_tr_acum, all_seat, all_comp_full = [], [], [], [], []

# --- 5. SIDEBAR ---
with st.sidebar:
    st.header("📅 Filtro Global")
    date_range = st.date_input("Período", value=(date(2025, 1, 1), date.today()))
    start_date, end_date = (date_range[0], date_range[1]) if len(date_range)==2 else (date_range, date_range)
    st.header("📂 Carga de Archivos")
    f_v1 = st.file_uploader("1. THDR Vía 1", accept_multiple_files=True)
    f_v2 = st.file_uploader("2. THDR Vía 2", accept_multiple_files=True)
    f_umr = st.file_uploader("3. UMR / Odómetros", accept_multiple_files=True)
    f_seat_f = st.file_uploader("4. Energía SEAT", accept_multiple_files=True)
    f_bill_f = st.file_uploader("5. Facturación y PRMTE", accept_multiple_files=True)

# --- 6. PROCESAMIENTO ---
if any([f_v1, f_v2, f_umr, f_seat_f, f_bill_f]):
    if f_umr:
        for f in f_umr:
            xl = pd.ExcelFile(f)
            for sn in xl.sheet_names:
                sn_up = sn.upper()
                if any(k in sn_up for k in ['UMR', 'RESUMEN']):
                    df_u = pd.read_excel(f, sheet_name=sn, header=None)
                    h_idx = next((i for i in range(min(50, len(df_u))) if 'ODO' in str(df_u.iloc[i]).upper()), None)
                    if h_idx is not None:
                        df_p = pd.read_excel(f, sheet_name=sn, header=h_idx)
                        df_p.columns = [re.sub(r'[^A-Z]', '', str(c).upper()) for c in df_p.columns]
                        if 'FECHA' in df_p.columns:
                            df_p['DT'] = pd.to_datetime(df_p['FECHA'], errors='coerce')
                            for _, r in df_p.dropna(subset=['DT']).iterrows():
                                if start_date <= r['DT'].date() <= end_date:
                                    all_ops.append({"Fecha": r['DT'].normalize(), "Tipo Día": get_tipo_dia(r['DT']), "N° Semana": r['DT'].isocalendar()[1], "Odómetro [km]": parse_latam_number(r.get('ODO',0)), "Tren-Km [km]": parse_latam_number(r.get('TRENKM',0)), "UMR [%]": (parse_latam_number(r.get('TRENKM',0))/parse_latam_number(r.get('ODO',1))*100)})
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
                                        all_tr_acum.append({"Tren":n, "Fecha":v.normalize(), "Día":v.day, "Valor":parse_latam_number(df_tr_raw.iloc[k,j])})

    if f_seat_f:
        for f in f_seat_f:
            df_s = pd.read_excel(f, header=None)
            for i in range(len(df_s)):
                dt = pd.to_datetime(df_s.iloc[i,1], errors='coerce')
                if pd.notna(dt) and start_date <= dt.date() <= end_date:
                    all_seat.append({"Fecha": dt.normalize(), "E_Total": parse_latam_number(df_s.iloc[i,3]), "E_Tr": parse_latam_number(df_s.iloc[i,5]), "E_12": parse_latam_number(df_s.iloc[i,7])})

    if f_bill_f:
        for f in f_bill_f:
            xl_b = pd.ExcelFile(f)
            for sn in xl_b.sheet_names:
                df_b = pd.read_excel(f, sheet_name=sn)
                if 'AÑO' in str(df_b.columns).upper():
                    df_b['TS'] = pd.to_datetime(df_b[['AÑO', 'MES', 'DIA', 'HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'}))
                    for _, r in df_b.iterrows():
                        v = parse_latam_number(r.get('Retiro_Energia_Activa (kWhD)', 0))
                        all_comp_full.append({"Fecha": r['TS'].normalize(), "Hora": r['TS'].hour, "Consumo": v, "Fuente": "PRMTE"})

    if f_v1:
        l1 = [procesar_thdr_avanzado(f)[0] for f in f_v1]; df_thdr_v1 = pd.concat(l1) if l1 else pd.DataFrame()
    if f_v2:
        l2 = [procesar_thdr_avanzado(f)[0] for f in f_v2]; df_thdr_v2 = pd.concat(l2) if l2 else pd.DataFrame()

    if all_ops:
        df_ops = pd.DataFrame(all_ops).drop_duplicates(subset=['Fecha']).sort_values("Fecha")
        for c in ['E_Total', 'E_Tr', 'E_12']: df_ops[c] = 0.0
        if all_seat:
            df_s_df = pd.DataFrame(all_seat).drop_duplicates(subset=['Fecha'])
            df_ops = pd.merge(df_ops.drop(columns=['E_Total','E_Tr','E_12']), df_s_df, on="Fecha", how="left").fillna(0)
        df_ops['IDE (kWh/km)'] = df_ops.apply(lambda r: r['E_Tr']/r['Odómetro [km]'] if r['Odómetro [km]']>0 else 0.0, axis=1)

# --- 7. DASHBOARD ---
tabs = st.tabs(["📊 Resumen", "📑 Operaciones", "📑 Trenes", "⚡ Energía", "⚖️ Comparación Energía hr", "📈 Regresión Nocturna", "🚨 Datos Atípicos", "📋 THDR"])

with tabs[0]: # PESTAÑA RESUMEN (RESTAURADA)
    if not df_ops.empty:
        c_f1, c_f2, c_f3 = st.columns(3)
        anios, meses, semanas = sorted(df_ops['Fecha'].dt.year.unique()), sorted(df_ops['Fecha'].dt.month.unique()), sorted(df_ops['N° Semana'].unique())
        f_ano = c_f1.multiselect("Año", anios, default=anios)
        f_mes = c_f2.multiselect("Mes", meses, default=meses)
        f_sem = c_f3.multiselect("Semana", semanas, default=semanas)
        df_res_f = df_ops[df_ops['Fecha'].dt.year.isin(f_ano) & df_ops['Fecha'].dt.month.isin(f_mes) & df_ops['N° Semana'].isin(f_sem)]
        
        if not df_res_f.empty:
            c1, c2, c3 = st.columns(3)
            c1.metric("Odómetro Total", f"{df_res_f['Odómetro [km]'].sum():,.1f} km")
            c2.metric("Tren-Km Total", f"{df_res_f['Tren-Km [km]'].sum():,.1f} km")
            c3.metric("IDE Medio", f"{df_res_f['IDE (kWh/km)'].mean():.4f}")
            e_tr, e_12 = df_res_f['E_Tr'].sum(), df_res_f['E_12'].sum()
            if (e_tr + e_12) > 0:
                st.info(f"⚡ Composición: Tracción {e_tr/(e_tr+e_12)*100:.1f}% | 12kV {e_12/(e_tr+e_12)*100:.1f}%")
            st.plotly_chart(go.Figure(go.Scatter(x=df_res_f['Fecha'], y=df_res_f['Odómetro [km]'], name="Km")), use_container_width=True)

with tabs[1]: # OPERACIONES
    if not df_ops.empty: st.dataframe(df_ops.style.format({'IDE (kWh/km)': '{:.4f}', 'Odómetro [km]': '{:,.1f}'}), use_container_width=True)

with tabs[2]: # TRENES (DIARIO + ACUMULADO)
    if not df_tr.empty:
        st.write("#### Kilometraje Diario por Tren")
        piv = df_tr.pivot_table(index="Tren", columns=df_tr["Fecha"].dt.day, values="Valor", aggfunc='sum').fillna(0)
        st.dataframe(piv.style.format("{:,.1f}"), use_container_width=True)
    if not df_tr_acum.empty:
        st.write("#### Lectura de Odómetro Acumulado")
        piv_a = df_tr_acum.pivot_table(index="Tren", columns=df_tr_acum["Fecha"].dt.day, values="Valor", aggfunc='max').fillna(0)
        st.dataframe(piv_a.style.format("{:,.0f}"), use_container_width=True)

with tabs[3]: # ENERGÍA
    if not df_ops.empty: st.dataframe(df_ops[['Fecha', 'E_Total', 'E_Tr', 'E_12']].style.format("{:,.1f}"), use_container_width=True)

with tabs[4]: # COMPARACIÓN
    if all_comp_full:
        df_c = pd.DataFrame(all_comp_full)
        piv_c = df_c.pivot_table(index="Hora", columns=df_c['Fecha'].dt.year, values="Consumo", aggfunc='median').fillna(0)
        st.line_chart(piv_c)

with tabs[5]: # REGRESIÓN
    if all_comp_full:
        df_r = pd.DataFrame(all_comp_full); df_r = df_r[df_r['Hora'] <= 5]
        h_sel = st.selectbox("Hora Regresión", sorted(df_r['Hora'].unique()))
        df_h = df_r[df_r['Hora'] == h_sel].sort_values("Fecha")
        if len(df_h) > 2:
            x, y = np.arange(len(df_h)), df_h['Consumo'].values
            m, b = np.polyfit(x, y, 1)
            st.markdown(f"**Ecuación:** $Consumo = {m:.4f}x + {b:.2f}$")
            st.plotly_chart(go.Figure([go.Scatter(x=df_h['Fecha'], y=y, name="Real"), go.Scatter(x=df_h['Fecha'], y=m*x+b, name="Tendencia")]))

with tabs[7]: # THDR (COLUMNAS CLARAS Y DINÁMICAS)
    st.write("### 📋 Tabla Horaria de Desempeño Real")
    c1, c2 = st.columns(2)
    with c1:
        st.write("#### Vía 1 (Detección de Origen/Destino)")
        if not df_thdr_v1.empty:
            cols_show = ['Fecha_Op', 'Servicio', 'Origen', 'Destino', 'Unidad', 'Tren-Km', 'Retraso']
            st.dataframe(df_thdr_v1[[c for c in cols_show if c in df_thdr_v1.columns]], use_container_width=True)
    with c2:
        st.write("#### Vía 2 (Detección de Origen/Destino)")
        if not df_thdr_v2.empty:
            cols_show = ['Fecha_Op', 'Servicio', 'Origen', 'Destino', 'Unidad', 'Tren-Km', 'Retraso']
            st.dataframe(df_thdr_v2[[c for c in cols_show if c in df_thdr_v2.columns]], use_container_width=True)

st.sidebar.divider()
st.sidebar.download_button("📥 Reporte Excel Completo", to_excel_consolidado(df_ops, df_tr, df_tr_acum, df_seat, df_p_d, df_f_d), "Reporte_SGE.xlsx")
