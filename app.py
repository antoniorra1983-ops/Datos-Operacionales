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

# --- 2. FUNCIONES DE PROCESAMIENTO ---

def parse_latam_number(val):
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    s = re.sub(r'[^\d.,-]', '', str(val).replace(' ', '').replace('$', ''))
    if not s: return 0.0
    if ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'): s = s.replace('.', '').replace(',', '.')
        else: s = s.replace(',', '')
    elif ',' in s: s = s.replace(',', '.')
    try: return float(s)
    except: return 0.0

def get_tipo_dia(fch):
    if fch in chile_holidays or fch.weekday() == 6: return "D/F"
    return "S" if fch.weekday() == 5 else "L"

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

def format_hms(m, con_signo=False):
    if pd.isna(m) or m == 0: return "00:00:00"
    signo = ("+" if m > 0 else "-" if m < 0 else "") if con_signo else ""
    sec = int(round(abs(m) * 60))
    h, r = divmod(sec, 3600); mi, se = divmod(r, 60)
    return f"{signo}{h:02d}:{mi:02d}:{se:02d}"

# --- 3. EXPORTACIÓN ---

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

# --- 4. PROCESAMIENTO THDR (CORREGIDO PARA EVITAR DUPLICADOS) ---

def procesar_thdr_avanzado(file):
    try:
        df_raw = pd.read_excel(file, header=None)
        # Fix Pandas 3.0: ffill() en lugar de method='ffill'
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
        df['Tipo_Rec'] = "PU-LI"
        df['Dist_Base'] = 43.13
        df['Tren-Km'] = df['Dist_Base'] * df['Unidad'].apply(lambda x: 2 if x == 'M' else 1)
        
        try:
            val_f = str(df_raw.iloc[0, 0]).split('.')[0].strip().zfill(6)
            df['Fecha_Op'] = f"{val_f[0:2]}/{val_f[2:4]}/20{val_f[4:6]}"
        except: df['Fecha_Op'] = ""
            
        return df[df['Servicio'] > 0], df['Tren-Km'].sum()
    except Exception as e:
        st.error(f"Error procesando THDR: {e}")
        return pd.DataFrame(), 0

# --- 5. INICIALIZACIÓN ---
if 'outliers' not in st.session_state: st.session_state.outliers = pd.DataFrame()

df_ops = df_tr = df_seat = df_p_d = df_f_d = df_thdr_v1 = df_thdr_v2 = pd.DataFrame()
all_ops, all_tr, all_seat, all_comp_full, all_prmte_15 = [], [], [], [], []

# --- 6. SIDEBAR ---
with st.sidebar:
    st.header("📅 Filtro Global")
    # Rango amplio por defecto
    date_range = st.date_input("Selecciona el período", value=(date(2025, 1, 1), date.today()))
    start_date, end_date = (date_range[0], date_range[1]) if len(date_range)==2 else (date_range, date_range)
    
    st.header("📂 Carga de Archivos")
    f_v1 = st.file_uploader("1. THDR Vía 1", accept_multiple_files=True)
    f_v2 = st.file_uploader("2. THDR Vía 2", accept_multiple_files=True)
    f_umr = st.file_uploader("3. UMR / Odómetros", accept_multiple_files=True)
    f_seat_f = st.file_uploader("4. Energía SEAT", accept_multiple_files=True)
    f_bill_f = st.file_uploader("5. Facturación y PRMTE", accept_multiple_files=True)

# --- 7. PROCESAMIENTO PRINCIPAL ---

if any([f_v1, f_v2, f_umr, f_seat_f, f_bill_f]):
    # Procesar UMR / Operaciones
    if f_umr:
        for f in f_umr:
            xl = pd.ExcelFile(f)
            for sn in xl.sheet_names:
                if any(k in sn.upper() for k in ['UMR', 'RESUMEN']):
                    df_u = pd.read_excel(f, sheet_name=sn, header=None)
                    h_idx = next((i for i in range(min(50, len(df_u))) if any(k in str(df_u.iloc[i]).upper() for k in ['ODO', 'FECHA'])), None)
                    if h_idx is not None:
                        df_p = pd.read_excel(f, sheet_name=sn, header=h_idx)
                        df_p.columns = [re.sub(r'[^A-Z]', '', str(c).upper()) for c in df_p.columns]
                        if 'FECHA' in df_p.columns:
                            df_p['DT'] = pd.to_datetime(df_p['FECHA'], errors='coerce')
                            for _, r in df_p.dropna(subset=['DT']).iterrows():
                                if start_date <= r['DT'].date() <= end_date:
                                    all_ops.append({
                                        "Fecha": r['DT'].normalize(), "Tipo Día": get_tipo_dia(r['DT']),
                                        "Odómetro [km]": parse_latam_number(r.get('ODO',0)),
                                        "Tren-Km [km]": parse_latam_number(r.get('TRENKM',0))
                                    })
                # Kms Diarios Trenes
                if 'ODO' in sn.upper() and 'KIL' in sn.upper():
                    df_tr_raw = pd.read_excel(f, sheet_name=sn, header=None)
                    for i in range(len(df_tr_raw)-2):
                        for j in range(1, len(df_tr_raw.columns)):
                            v = pd.to_datetime(df_tr_raw.iloc[i,j], errors='coerce')
                            if pd.notna(v) and start_date <= v.date() <= end_date:
                                for k in range(i+3, min(i+40, len(df_tr_raw))):
                                    n = str(df_tr_raw.iloc[k,0]).strip().upper()
                                    if n.startswith(('M','XM')):
                                        all_tr.append({"Tren":n, "Fecha":v.normalize(), "Valor":parse_latam_number(df_tr_raw.iloc[k,j])})

    # Procesar Energía SEAT
    if f_seat_f:
        for f in f_seat_f:
            df_s = pd.read_excel(f, header=None)
            for i in range(len(df_s)):
                dt = pd.to_datetime(df_s.iloc[i,1], errors='coerce')
                if pd.notna(dt) and start_date <= dt.date() <= end_date:
                    all_seat.append({
                        "Fecha": dt.normalize(), "E_Total": parse_latam_number(df_s.iloc[i,3]),
                        "E_Tr": parse_latam_number(df_s.iloc[i,5]), "E_12": parse_latam_number(df_s.iloc[i,7])
                    })

    # Procesar PRMTE / Facturas
    if f_bill_f:
        for f in f_bill_f:
            xl_b = pd.ExcelFile(f)
            for sn in xl_b.sheet_names:
                df_b = pd.read_excel(f, sheet_name=sn)
                if 'AÑO' in str(df_b.columns).upper():
                    # Consolidado horario para regresión y comparativa
                    df_b['TS'] = pd.to_datetime(df_b[['AÑO', 'MES', 'DIA', 'HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'}))
                    for _, r in df_b.iterrows():
                        v = parse_latam_number(r.get('Retiro_Energia_Activa (kWhD)', 0))
                        all_comp_full.append({"Fecha": r['TS'].normalize(), "Hora": r['TS'].hour, "Consumo": v, "Fuente": "PRMTE"})

    # Procesar THDR
    if f_v1:
        l1 = [procesar_thdr_avanzado(f)[0] for f in f_v1]; df_thdr_v1 = pd.concat(l1) if l1 else pd.DataFrame()
    if f_v2:
        l2 = [procesar_thdr_avanzado(f)[0] for f in f_v2]; df_thdr_v2 = pd.concat(l2) if l2 else pd.DataFrame()

    # Consolidación de tablas para el Dashboard
    if all_ops:
        df_ops = pd.DataFrame(all_ops).drop_duplicates(subset=['Fecha']).sort_values("Fecha")
        if all_seat:
            df_ops = pd.merge(df_ops, pd.DataFrame(all_seat).drop_duplicates(subset=['Fecha']), on="Fecha", how="left").fillna(0)
            df_ops['IDE (kWh/km)'] = df_ops.apply(lambda r: r['E_Tr']/r['Odómetro [km]'] if r['Odómetro [km]']>0 else 0, axis=1)

    if all_tr:
        df_tr = pd.DataFrame(all_tr).sort_values(["Fecha", "Tren"])

# --- 8. DASHBOARD ---

tabs = st.tabs(["📊 Resumen", "📑 Operaciones", "📑 Trenes", "⚡ Energía", "⚖️ Comparativa Hr", "📈 Regresión", "🚨 Atípicos", "📋 THDR"])

with tabs[0]: # RESUMEN
    if not df_ops.empty:
        c1, c2, c3 = st.columns(3)
        c1.metric("Odómetro Total", f"{df_ops['Odómetro [km]'].sum():,.0f} km")
        c2.metric("Tren-Km Total", f"{df_ops['Tren-Km [km]'].sum():,.0f} km")
        c3.metric("IDE Medio", f"{df_ops['IDE (kWh/km)'].mean():.4f} kWh/km")
        
        st.write("#### Kilometraje Diario")
        st.line_chart(df_ops.set_index("Fecha")["Odómetro [km]"])
    else:
        st.info("Sube archivos UMR para ver el resumen global.")

with tabs[1]: # OPERACIONES
    if not df_ops.empty:
        st.write("### Tabla de Operaciones")
        st.dataframe(df_ops.style.format({'IDE (kWh/km)': '{:.4f}', 'Odómetro [km]': '{:,.1f}'}), use_container_width=True)
    else:
        st.warning("No hay datos de operaciones cargados.")

with tabs[2]: # TRENES
    if not df_tr.empty:
        st.write("### Kilometraje por Tren")
        piv_tr = df_tr.pivot_table(index="Tren", columns=df_tr["Fecha"].dt.day, values="Valor", aggfunc='sum').fillna(0)
        st.dataframe(piv_tr.style.format("{:,.1f}"), use_container_width=True)
    else:
        st.info("Carga archivos UMR/Odómetros para ver el detalle por tren.")

with tabs[3]: # ENERGÍA
    if not df_ops.empty:
        st.write("### Consumo Energético Consolidado")
        st.dataframe(df_ops[['Fecha', 'E_Total', 'E_Tr', 'E_12']].style.format("{:,.1f}"), use_container_width=True)
    else:
        st.info("Carga archivos SEAT para visualizar el consumo.")

with tabs[4]: # COMPARATIVA
    if all_comp_full:
        df_c = pd.DataFrame(all_comp_full)
        piv_c = df_c.pivot_table(index="Hora", columns=df_c['Fecha'].dt.year, values="Consumo", aggfunc='median').fillna(0)
        st.write("#### Mediana de Consumo Horario por Año")
        st.line_chart(piv_c)
    else:
        st.info("Sube archivos PRMTE para comparar el consumo horario.")

with tabs[5]: # REGRESIÓN
    if all_comp_full:
        df_reg_all = pd.DataFrame(all_comp_full)
        df_reg_all = df_reg_all[df_reg_all['Hora'] <= 5] # Foco en madrugada
        h_sel = st.selectbox("Selecciona Hora para Línea de Base", sorted(df_reg_all['Hora'].unique()))
        df_h = df_reg_all[df_reg_all['Hora'] == h_sel].sort_values("Fecha")
        if len(df_h) > 2:
            x = np.arange(len(df_h)); y = df_h['Consumo'].values
            m, b = np.polyfit(x, y, 1)
            st.markdown(f"**Ecuación de Tendencia:** $Consumo = {m:.4f}x + {b:.2f}$")
            fig_r = go.Figure()
            fig_r.add_trace(go.Scatter(x=df_h['Fecha'], y=y, name="Real"))
            fig_r.add_trace(go.Scatter(x=df_h['Fecha'], y=m*x+b, name="Tendencia", line=dict(dash='dash')))
            st.plotly_chart(fig_r)

with tabs[6]: # ATÍPICOS
    if all_comp_full:
        df_at = pd.DataFrame(all_comp_full)
        q1, q3 = df_at['Consumo'].quantile(0.25), df_at['Consumo'].quantile(0.75)
        iqr = q3 - q1
        atipicos = df_at[(df_at['Consumo'] > q3 + 1.5*iqr) | (df_at['Consumo'] < q1 - 1.5*iqr)]
        st.error(f"Se detectaron {len(atipicos)} anomalías en el consumo horario.")
        st.dataframe(atipicos, use_container_width=True)

with tabs[7]: # THDR
    st.subheader("📋 Tabla Horaria de Desempeño Real")
    c1, c2 = st.columns(2)
    with c1:
        st.write("#### Vía 1")
        if not df_thdr_v1.empty: st.dataframe(df_thdr_v1, use_container_width=True)
        else: st.info("Vía 1 vacía.")
    with c2:
        st.write("#### Vía 2")
        if not df_thdr_v2.empty: st.dataframe(df_thdr_v2, use_container_width=True)
        else: st.info("Vía 2 vacía.")

st.sidebar.divider()
st.sidebar.write("Desarrollado para Gestión Energética EFE Valparaíso.")
