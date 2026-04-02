import streamlit as st
import pandas as pd
import numpy as np
import re
import holidays
from io import BytesIO
from datetime import datetime, date
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# --- 1. CONFIGURACIÓN Y CONSTANTES ---
st.set_page_config(page_title="SGE EFE Valparaíso - Dashboard Oficial", layout="wide", page_icon="🚆")
chile_holidays = holidays.Chile()
ORDEN_JORNADA = ['L', 'S', 'D/F']

st.markdown("""
    <style>
    .stMetric { background-color: #ffffff; padding: 20px; border-radius: 10px; border-left: 5px solid #005195; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    </style>
    """, unsafe_allow_html=True)

# --- 2. FUNCIONES DE APOYO ---

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
            p.text = f"• {k}: {v}"; p.font.size = Pt(14); p.font.bold = True
        y_cursor += Inches(1.2)
    if df is not None and not df.empty:
        df_export = df.reset_index() if hasattr(df, 'index') and (df.index.name or any(df.index.names)) else df
        rows, cols = df_export.shape
        table = slide.shapes.add_table(rows + 1, cols, Inches(0.1), y_cursor, Inches(9.8), Inches(4)).table
        for c, col_name in enumerate(df_export.columns):
            table.cell(0, c).text = str(col_name)
    buf = BytesIO(); prs.save(buf); return buf.getvalue()

# --- 3. SIDEBAR ---
with st.sidebar:
    st.header("📅 Filtro Global")
    date_range = st.date_input("Periodo", value=(date.today().replace(day=1), date.today()))
    start_date, end_date = (date_range[0], date_range[1]) if len(date_range)==2 else (date_range[0], date_range[0])
    st.header("📂 Archivos")
    f_umr = st.file_uploader("1. UMR", type=["xlsx"], accept_multiple_files=True)
    f_seat_f = st.file_uploader("2. SEAT", type=["xlsx"], accept_multiple_files=True)
    f_bill_f = st.file_uploader("3. Factura/PRMTE", type=["xlsx"], accept_multiple_files=True)

# --- 4. PROCESAMIENTO (REVISADO LÍNEA POR LÍNEA) ---
all_ops, all_tr, all_tr_acum, all_seat, all_prmte_15, all_fact_h, all_comp_full = [], [], [], [], [], [], []
archivos = (f_umr or []) + (f_seat_f or []) + (f_bill_f or [])

for f in archivos:
    try:
        xl = pd.ExcelFile(f)
        for sn in xl.sheet_names:
            su = sn.upper()
            if any(k in su for k in ['UMR', 'RESUMEN']):
                df = pd.read_excel(f, sheet_name=sn, header=None)
                h = next((i for i in range(min(50, len(df))) if any(k in str(df.iloc[i]).upper() for k in ['ODO', 'FECHA'])), None)
                if h is not None:
                    df_p = pd.read_excel(f, sheet_name=sn, header=h)
                    df_p.columns = [re.sub(r'[^A-Z]', '', str(c).upper().replace('Ó','O')) for c in df_p.columns]
                    if 'FECHA' in df_p.columns:
                        df_p['_dt'] = pd.to_datetime(df_p['FECHA'], errors='coerce')
                        mask = (df_p['_dt'].dt.date >= start_date) & (df_p['_dt'].dt.date <= end_date)
                        for _, r in df_p[mask].dropna(subset=['_dt']).iterrows():
                            all_ops.append({"Fecha": r['_dt'].normalize(), "Tipo Día": get_tipo_dia(r['_dt']), "N° Semana": r['_dt'].isocalendar()[1], "Odómetro [km]": parse_latam_number(r.get('ODO', 0)), "Tren-Km [km]": parse_latam_number(r.get('TRENKM', 0))})
            
            if 'ODO' in su and 'KIL' in su:
                df_r = pd.read_excel(f, sheet_name=sn, header=None)
                for i in range(len(df_r)-2):
                    for j in range(1, len(df_r.columns)):
                        dt = pd.to_datetime(df_r.iloc[i, j], errors='coerce')
                        if pd.notna(dt) and start_date <= dt.date() <= end_date:
                            for k in range(i+3, min(i+40, len(df_r))):
                                tr = str(df_r.iloc[k, 0]).strip().upper()
                                if re.match(r'^(M|XM)', tr):
                                    val = parse_latam_number(df_r.iloc[k, j])
                                    item = {"Tren": tr, "Fecha": dt.normalize(), "Valor": val}
                                    if 'ACUM' in su: all_tr_acum.append(item)
                                    else: all_tr.append(item)

            if 'SEAT' in su:
                df_s = pd.read_excel(f, sheet_name=sn, header=None)
                for i in range(len(df_s)):
                    fs = pd.to_datetime(df_s.iloc[i, 1], errors='coerce')
                    if pd.notna(fs) and start_date <= fs.date() <= end_date:
                        all_seat.append({"Fecha": fs.normalize(), "Total [kWh]": parse_latam_number(df_s.iloc[i, 3]), "Tracción [kWh]": parse_latam_number(df_s.iloc[i, 5]), "12 KV [kWh]": parse_latam_number(df_s.iloc[i, 7])})

            if any(k in su for k in ['PRMTE', 'FACTURA']):
                df_f = pd.read_excel(f, sheet_name=sn, header=None)
                hi = next((i for i in range(len(df_f)) if any(k in str(df_f.iloc[i]).upper() for k in ['AÑO', 'FECHAHORA'])), None)
                if hi is not None:
                    df_d = pd.read_excel(f, sheet_name=sn, header=hi)
                    if 'AÑO' in df_d.columns:
                        df_d['TS'] = pd.to_datetime(df_d[['AÑO','MES','DIA','HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'})) + pd.to_timedelta(df_d['INICIO INTERVALO'].astype(int), unit='m')
                        col_e = [c for c in df_d.columns if 'Energia_Activa' in str(c)][0]
                        for _, r in df_d.iterrows():
                            all_comp_full.append({"Fecha": r['TS'].normalize(), "Hora": r['TS'].hour, "Consumo": parse_latam_number(r[col_e]), "Fuente": "PRMTE"})
                            if start_date <= r['TS'].date() <= end_date: all_prmte_15.append({"Fecha": r['TS'].normalize(), "Energía PRMTE [kWh]": parse_latam_number(r[col_e])})
                    else:
                        df_d.columns = ['TS', 'Val']; df_d['TS'] = pd.to_datetime(df_d['TS'], errors='coerce')
                        for _, r in df_d.dropna(subset=['TS']).iterrows():
                            all_comp_full.append({"Fecha": r['TS'].normalize(), "Hora": r['TS'].hour, "Consumo": abs(parse_latam_number(r['Val'])), "Fuente": "Factura"})
                            if start_date <= r['TS'].date() <= end_date: all_fact_h.append({"Fecha": r['TS'].normalize(), "Consumo Horario [kWh]": abs(parse_latam_number(r['Val']))})
    except: continue

# --- 5. PREPARACIÓN DE DATAFRAMES ---
df_ops = pd.DataFrame(all_ops)
df_tr = pd.DataFrame(all_tr)
df_tr_a = pd.DataFrame(all_tr_acum)
df_seat = pd.DataFrame(all_seat)
df_prmte = pd.DataFrame(all_prmte_15)
df_fact = pd.DataFrame(all_fact_h)

# Triangulación de Energía para Resumen
df_m = pd.DataFrame()
if not df_seat.empty: df_m = df_seat.rename(columns={"Total [kWh]":"E_Total"})[["Fecha", "E_Total"]]
if not df_prmte.empty:
    p_sum = df_prmte.groupby("Fecha")["Energía PRMTE [kWh]"].sum().reset_index().rename(columns={"Energía PRMTE [kWh]":"E_Total"})
    df_m = pd.concat([df_m, p_sum]).drop_duplicates("Fecha", keep="last")
if not df_fact.empty:
    f_sum = df_fact.groupby("Fecha")["Consumo Horario [kWh]"].sum().reset_index().rename(columns={"Consumo Horario [kWh]":"E_Total"})
    df_m = pd.concat([df_m, f_sum]).drop_duplicates("Fecha", keep="last")
if not df_ops.empty and not df_m.empty: df_ops = pd.merge(df_ops, df_m, on="Fecha", how="left")

# --- 6. RENDERIZADO (TABS) ---
if not df_ops.empty or not df_seat.empty or all_comp_full:
    tabs = st.tabs(["📊 Resumen", "📑 Operaciones", "📑 Trenes", "⚡ Energía cruda", "⚖️ Comparación hr", "📈 Regresión", "🚨 Atípicos"])
    
    with tabs[0]: # --- RESUMEN ---
        st.header("📊 Resumen SGE")
        if not df_ops.empty:
            c1, c2 = st.columns(2)
            f_ano = c1.multiselect("Año", sorted(df_ops['Fecha'].dt.year.unique()), default=sorted(df_ops['Fecha'].dt.year.unique()))
            f_mes = c2.multiselect("Mes", sorted(df_ops['Fecha'].dt.month.unique()))
            mask = df_ops['Fecha'].dt.year.isin(f_ano)
            if f_mes: mask &= df_ops['Fecha'].dt.month.isin(f_mes)
            df_f = df_ops[mask].copy()
            if not df_f.empty:
                to, tk = df_f["Odómetro [km]"].sum(), df_f["Tren-Km [km]"].sum()
                st.columns(3)[0].metric("Odómetro", f"{to:,.1f}"); st.columns(3)[1].metric("Tren-Km", f"{tk:,.1f}"); st.columns(3)[2].metric("UMR", f"{(tk/to*100) if to>0 else 0:.2f}%")
                st.write("#### Jornada (Orden: L, S, D/F)")
                res_j = df_f.groupby("Tipo Día", observed=True).agg({"Odómetro [km]":"sum", "Tren-Km [km]":"sum"}).reindex(ORDEN_JORNADA).dropna(how='all')
                st.table(res_j.style.format("{:,.1f}"))

    with tabs[2]: # --- TRENES ---
        if not df_tr.empty:
            st.subheader("🚗 Kilometraje Diario")
            st.dataframe(df_tr.pivot_table(index="Tren", columns=df_tr["Fecha"].dt.day, values="Valor", aggfunc='sum'))
        if not df_tr_a.empty:
            st.subheader("📈 Odómetro Acumulado")
            st.dataframe(df_tr_a.pivot_table(index="Tren", columns=df_tr_a["Fecha"].dt.day, values="Valor", aggfunc='max'))

    with tabs[3]: # --- ENERGÍA CRUDA ---
        st.header("⚡ Energía (SEAT / PRMTE / Factura)")
        s1, s2, s3 = st.tabs(["⚡ SEAT", "📈 PRMTE", "💰 Factura"])
        s1.dataframe(df_seat); s2.dataframe(df_prmte); s3.dataframe(df_fact)

    with tabs[4]: # --- COMPARACIÓN HORARIA ---
        if all_comp_full:
            st.header("⚖️ Comparativa Horaria")
            df_c = pd.DataFrame(all_comp_full).groupby(['Fecha','Hora','Fuente'])['Consumo'].sum().reset_index()
            fechas_f = df_c[df_c['Fuente']=='Factura']['Fecha'].unique()
            df_cf = df_c[~((df_c['Fuente']=='PRMTE') & (df_c['Fecha'].isin(fechas_f)))].copy()
            df_cf['Año'] = df_cf['Fecha'].dt.year.astype(str)
            df_cf['Tipo Día'] = df_cf['Fecha'].apply(get_tipo_dia)
            pivot = df_cf.pivot_table(index="Hora", columns=["Año", "Tipo Día"], values="Consumo", aggfunc='median').fillna(0)
            for a in sorted(df_cf['Año'].unique()):
                pivot[(a, "Total Anual")] = df_cf[df_cf['Año'] == a].groupby("Hora")["Consumo"].median()
            new_c = []
            for a in sorted(df_cf['Año'].unique()):
                for j in ORDEN_JORNADA + ["Total Anual"]:
                    if (a, j) in pivot.columns: new_c.append((a, j))
            st.dataframe(pivot.reindex(columns=new_c).style.format("{:,.1f}"), use_container_width=True)

    with tabs[5]: # --- REGRESIÓN (00-05 AM) ---
        st.header("📈 Regresión 12kV (Baja Tensión)")
        if all_comp_full:
            df_reg = pd.DataFrame(all_comp_full).groupby(['Fecha','Hora'])['Consumo'].sum().reset_index()
            df_reg = df_reg[df_reg['Hora'] <= 5]
            df_reg['Año'] = df_reg['Fecha'].dt.year
            c1, c2 = st.columns(2)
            fa, fh = c1.selectbox("Año", sorted(df_reg['Año'].unique()), key="ra"), c2.selectbox("Hora", range(6), key="rh")
            df_p = df_reg[(df_reg['Año']==fa) & (df_reg['Hora']==fh)].sort_values("Fecha")
            if len(df_p) > 2:
                y = df_p['Consumo'].values; x = np.arange(len(y))
                m, n = np.polyfit(x, y, 1); r2 = np.corrcoef(x, y)[0,1]**2
                st.line_chart(pd.DataFrame({"Real":y, "Tendencia":m*x+n}, index=df_p['Fecha'].dt.strftime('%d/%m')))
                st.metric("Carga Basal (n)", f"{n:.2f} kWh")
                st.write(f"Ecuación: y = {m:.4f}x + {n:.2f} | R² = {r2:.4f}")
else:
    st.info("👋 Sube los archivos para activar el dashboard.")
