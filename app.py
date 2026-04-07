
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

def exportar_resumen_excel(metrics_dict, df_resumen_jornada, df_energia):
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

def procesar_thdr_avanzado(file):
    try:
        df_raw = pd.read_excel(file, header=None)
        h0 = df_raw.iloc[0].ffill().astype(str)
        h1 = df_raw.iloc[1].fillna('').astype(str)
        
        # Deduplicación de columnas para evitar error de PyArrow
        raw_cols = [f"{a}_{b}".strip('_ ') for a, b in zip(h0, h1)]
        final_cols = []
        counts = {}
        for col in raw_cols:
            if col in counts:
                counts[col] += 1
                final_cols.append(f"{col}_{counts[col]}")
            else:
                counts[col] = 0
                final_cols.append(col)
        
        df = df_raw.iloc[2:].copy()
        df.columns = final_cols
        
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
        
        try:
            val_f = str(df_raw.iloc[0, 0]).split('.')[0].strip().zfill(6)
            df['Fecha_Op'] = f"{val_f[0:2]}/{val_f[2:4]}/20{val_f[4:6]}"
        except: df['Fecha_Op'] = ""
            
        return df[df['Servicio'] > 0], df['Tren-Km'].sum(), df[df['TDV_Min']>0]['TDV_Min'].mean(), (df['Puntual'].sum()/len(df)*100 if len(df)>0 else 0)
    except: return pd.DataFrame(), 0, 0, 0

# --- 4. INICIALIZACIÓN ---
if 'outliers' not in st.session_state: st.session_state.outliers = pd.DataFrame()
df_ops = df_tr = df_tr_acum = df_seat = df_energy_master = df_p_d = df_f_d = df_thdr_v1 = df_thdr_v2 = pd.DataFrame()
all_ops, all_tr, all_tr_acum, all_seat, all_prmte_15, all_fact_h, all_comp_full = [], [], [], [], [], [], []

# --- 5. SIDEBAR ---
with st.sidebar:
    st.header("📅 Filtro Global")
    date_range = st.date_input("Período", value=(date(2025, 1, 1), date.today()))
    start_date, end_date = (date_range[0], date_range[1]) if isinstance(date_range, tuple) and len(date_range)==2 else (date_range, date_range)
    st.header("📂 Carga de Archivos")
    f_v1 = st.file_uploader("1. THDR Vía 1", accept_multiple_files=True)
    f_v2 = st.file_uploader("2. THDR Vía 2", accept_multiple_files=True)
    f_umr = st.file_uploader("3. UMR / Odómetros", accept_multiple_files=True)
    f_seat_f = st.file_uploader("4. Energía SEAT", accept_multiple_files=True)
    f_bill_f = st.file_uploader("5. Facturación y PRMTE", accept_multiple_files=True)

# --- 6. PROCESAMIENTO ---
if any([f_v1, f_v2, f_umr, f_seat_f, f_bill_f]):
    todos = (f_v1 or []) + (f_v2 or []) + (f_umr or []) + (f_seat_f or []) + (f_bill_f or [])
    for f in todos:
        try:
            xl = pd.ExcelFile(f)
            for sn in xl.sheet_names:
                sn_up = sn.upper()
                # UMR
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
                # Trenes
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
                # SEAT
                if 'SEAT' in sn_up:
                    df_s = pd.read_excel(f, sheet_name=sn, header=None)
                    for i in range(len(df_s)):
                        dt = pd.to_datetime(df_s.iloc[i,1], errors='coerce')
                        if pd.notna(dt) and start_date <= dt.date() <= end_date:
                            tot, tra = parse_latam_number(df_s.iloc[i,3]), parse_latam_number(df_s.iloc[i,5])
                            all_seat.append({"Fecha": dt.normalize(), "E_Total": tot, "E_Tr": tra, "E_12": parse_latam_number(df_s.iloc[i,7])})
                # PRMTE
                if any(k in sn_up for k in ['PRMTE', 'MEDIDAS']):
                    df_pd = pd.read_excel(f, sheet_name=sn, header=0) # Asumimos header en 0 o similar
                    # Lógica de PRMTE para comparación horaria
                    if 'AÑO' in str(df_pd.columns).upper():
                        df_pd['Timestamp'] = pd.to_datetime(df_pd[['AÑO', 'MES', 'DIA', 'HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'}))
                        cols_e = [c for c in df_pd.columns if 'Retiro' in str(c)]
                        for _, r in df_pd.iterrows():
                            ts, val_p = r['Timestamp'], sum([parse_latam_number(r[col]) for col in cols_e])
                            all_comp_full.append({"Fecha":ts.normalize(), "Hora":ts.hour, "Consumo Horario [kWh]":val_p, "Fuente":"PRMTE"})
                            if start_date <= ts.date() <= end_date: all_prmte_15.append({"Fecha":ts.normalize(), "Energía PRMTE [kWh]":val_p})
        except: continue
    
    if f_v1:
        l1 = [procesar_thdr_avanzado(f)[0] for f in f_v1]; df_thdr_v1 = pd.concat(l1) if l1 else pd.DataFrame()
    if f_v2:
        l2 = [procesar_thdr_avanzado(f)[0] for f in f_v2]; df_thdr_v2 = pd.concat(l2) if l2 else pd.DataFrame()

    if all_ops:
        df_ops = pd.DataFrame(all_ops).drop_duplicates(subset=['Fecha']).sort_values("Fecha")
        if all_seat:
            df_seat = pd.DataFrame(all_seat).drop_duplicates(subset=['Fecha'])
            df_ops = pd.merge(df_ops, df_seat, on="Fecha", how="left").fillna(0)
            df_ops['IDE (kWh/km)'] = df_ops.apply(lambda r: r['E_Tr']/r['Odómetro [km]'] if r['Odómetro [km]']>0 else 0, axis=1)
    if all_tr: df_tr = pd.DataFrame(all_tr).sort_values(["Fecha", "Tren"])
    if all_prmte_15: df_p_d = pd.DataFrame(all_prmte_15).groupby("Fecha")["Energía PRMTE [kWh]"].sum().reset_index()

# --- 7. DASHBOARD ---
tabs = st.tabs(["📊 Resumen", "📑 Operaciones", "📑 Trenes", "⚡ Energía", "⚖️ Comparación Energía hr", "📈 Regresión Nocturna", "🚨 Datos Atípicos", "📋 THDR"])

with tabs[0]: # RESUMEN
    if not df_ops.empty:
        st.subheader("Estado del Sistema de Gestión")
        c1, c2, c3 = st.columns(3)
        c1.metric("Odómetro Total", f"{df_ops['Odómetro [km]'].sum():,.0f} km")
        c2.metric("Tren-Km Total", f"{df_ops['Tren-Km [km]'].sum():,.0f} km")
        c3.metric("IDE Medio", f"{df_ops['IDE (kWh/km)'].mean():.4f} kWh/km")
        fig_res = make_subplots(specs=[[{"secondary_y": True}]])
        fig_res.add_trace(go.Bar(x=df_ops['Fecha'], y=df_ops['Odómetro [km]'], name="Odómetro"), secondary_y=False)
        fig_res.add_trace(go.Scatter(x=df_ops['Fecha'], y=df_ops['IDE (kWh/km)'], name="IDE", line=dict(color='red')), secondary_y=True)
        st.plotly_chart(fig_res, use_container_width=True)
        
        st.write("#### 📥 Exportar Resumen")
        m_dict = {"Odómetro Total": df_ops['Odómetro [km]'].sum(), "IDE Promedio": df_ops['IDE (kWh/km)'].mean()}
        exc_res = exportar_resumen_excel(m_dict, df_ops.groupby("Tipo Día").sum().reset_index(), df_ops[['Fecha', 'E_Total']])
        st.download_button("Descargar Excel", exc_res, "Resumen_EFE.xlsx")
    else: st.info("Sube archivos para activar el resumen.")

with tabs[1]: # OPERACIONES
    if not df_ops.empty:
        st.dataframe(df_ops.style.format({'IDE (kWh/km)': '{:.4f}', 'Odómetro [km]': '{:,.1f}'}), use_container_width=True)
        st.download_button("Descargar PPTX", to_pptx("Datos Operacionales", df_ops), "EFE_Ops.pptx")
    else: st.info("Sin datos.")

with tabs[2]: # TRENES
    if not df_tr.empty:
        piv = df_tr.pivot_table(index="Tren", columns=df_tr["Fecha"].dt.day, values="Valor", aggfunc='sum').fillna(0)
        st.write("#### Kilometraje Diario por Unidad")
        st.dataframe(piv.style.format("{:,.1f}"), use_container_width=True)
    else: st.info("Sin datos de trenes.")

with tabs[3]: # ENERGIA
    if not df_ops.empty:
        st.write("#### Consumo Consolidado")
        st.dataframe(df_ops[['Fecha', 'E_Total', 'E_Tr', 'E_12']].style.format("{:,.1f}"), use_container_width=True)
    else: st.info("Sin datos de energía.")

with tabs[4]: # COMPARACION HR
    if all_comp_full:
        df_c = pd.DataFrame(all_comp_full).groupby(['Fecha','Hora','Fuente'])['Consumo Horario [kWh]'].sum().reset_index()
        df_c['Año'] = df_c['Fecha'].dt.year
        piv_c = df_c.pivot_table(index="Hora", columns="Año", values="Consumo Horario [kWh]", aggfunc='median').fillna(0)
        st.write("#### Mediana de Consumo Horario por Año")
        st.line_chart(piv_c)
        st.dataframe(piv_c.style.format("{:,.1f}"), use_container_width=True)
    else: st.info("Sin datos horarios.")

with tabs[5]: # REGRESION
    if all_comp_full:
        df_reg = pd.DataFrame(all_comp_full)
        df_reg = df_reg[df_reg['Hora'] <= 5] # Madrugada
        h_sel = st.selectbox("Hora para Regresión Nocturna", sorted(df_reg['Hora'].unique()))
        df_h = df_reg[df_reg['Hora'] == h_sel].sort_values("Fecha")
        if len(df_h) > 2:
            x = np.arange(len(df_h)); y = df_h['Consumo Horario [kWh]'].values
            m, b = np.polyfit(x, y, 1)
            st.markdown(f"**Ecuación:** $Consumo = {m:.4f}x + {b:.2f}$")
            fig_reg = go.Figure()
            fig_reg.add_trace(go.Scatter(x=df_h['Fecha'], y=y, name="Real"))
            fig_reg.add_trace(go.Scatter(x=df_h['Fecha'], y=m*x+b, name="Tendencia", line=dict(dash='dash')))
            st.plotly_chart(fig_reg)
    else: st.info("Carga datos de PRMTE o Factura para regresión.")

with tabs[6]: # ATIPICOS
    if all_comp_full:
        df_at = pd.DataFrame(all_comp_full)
        q1, q3 = df_at['Consumo Horario [kWh]'].quantile(0.25), df_at['Consumo Horario [kWh]'].quantile(0.75)
        iqr = q3 - q1
        outliers = df_at[(df_at['Consumo Horario [kWh]'] > q3 + 1.5*iqr) | (df_at['Consumo Horario [kWh]'] < q1 - 1.5*iqr)]
        st.error(f"Se detectaron {len(outliers)} anomalías.")
        st.dataframe(outliers, use_container_width=True)
    else: st.info("Sin datos para análisis de atípicos.")

with tabs[7]: # THDR
    st.write("#### Vía 1")
    if not df_thdr_v1.empty: st.dataframe(df_thdr_v1, use_container_width=True)
    else: st.info("No hay datos Vía 1.")
    st.write("#### Vía 2")
    if not df_thdr_v2.empty: st.dataframe(df_thdr_v2, use_container_width=True)
    else: st.info("No hay datos Vía 2.")

# --- 8. DESCARGA REPORTE GLOBAL ---
st.sidebar.divider()
st.sidebar.download_button("📥 Reporte Excel Completo", to_excel_consolidado(df_ops, df_tr, pd.DataFrame(), df_seat, df_p_d, pd.DataFrame(), pd.DataFrame(), pd.DataFrame()), "Reporte_SGE_EFE.xlsx")
