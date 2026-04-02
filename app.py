import streamlit as st
import pandas as pd
import numpy as np
import re
import holidays
from io import BytesIO
from datetime import datetime, date
# Librerías para PPTX
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# --- 1. CONFIGURACIÓN Y ESTILOS ---
st.set_page_config(page_title="Gestión de Energía - Dashboard SGE", layout="wide", page_icon="🚆")
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
            p.text = f"• {k}: {v}"; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = RGBColor(0, 81, 149)
        y_cursor += Inches(1.2)
    if df is not None and not df.empty:
        df_export = df.reset_index() if hasattr(df, 'index') and (df.index.name or any(df.index.names)) else df
        rows, cols = df_export.shape
        table = slide.shapes.add_table(rows + 1, cols, Inches(0.1), y_cursor, Inches(9.8), Inches(4.5)).table
        for c, col_name in enumerate(df_export.columns):
            cell = table.cell(0, c); cell.text = str(col_name); cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(0, 81, 149)
            cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255); cell.text_frame.paragraphs[0].font.size = Pt(9)
        for r in range(rows):
            for c in range(cols):
                val = df_export.iloc[r, c]
                table.cell(r + 1, c).text = str(val) if not isinstance(val, float) else f"{val:,.1f}"
    binary_output = BytesIO(); prs.save(binary_output); return binary_output.getvalue()

# --- 3. SIDEBAR Y CARGA ---
with st.sidebar:
    st.header("📅 Filtro Global")
    today = date.today()
    date_range = st.date_input("Selecciona el período", value=(today.replace(day=1), today))
    start_date, end_date = (date_range[0], date_range[1]) if isinstance(date_range, tuple) and len(date_range)==2 else (date_range, date_range)
    st.header("📂 Carga de Archivos")
    f_umr = st.file_uploader("1. UMR / Odómetros", type=["xlsx"], accept_multiple_files=True)
    f_seat_files = st.file_uploader("2. Energía SEAT", type=["xlsx"], accept_multiple_files=True)
    f_bill_files = st.file_uploader("3. Facturación y PRMTE", type=["xlsx"], accept_multiple_files=True)

# --- 4. MOTOR DE DATOS (ESTRUCTURA ORIGINAL) ---
all_ops, all_tr, all_tr_acum, all_seat, all_prmte_15, all_fact_h, all_comp_full = [], [], [], [], [], [], []
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
                            all_ops.append({"Fecha": r['_dt'].normalize(), "Tipo Día": get_tipo_dia(r['_dt']), "N° Semana": r['_dt'].isocalendar()[1], "Odómetro [km]": parse_latam_number(r[idx_o]), "Tren-Km [km]": parse_latam_number(r[idx_t])})
            
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
                                d_pt = {"Tren": n_tr, "Fecha": c_fch.normalize(), "Valor": val_km}
                                if is_acum or idx > 0: all_tr_acum.append(d_pt)
                                else: all_tr.append(d_pt)

            if 'SEAT' in sn_up and 'SER' in sn_up:
                df_s = pd.read_excel(f, sheet_name=sn, header=None)
                for i in range(len(df_s)):
                    fs = pd.to_datetime(df_s.iloc[i, 1], errors='coerce')
                    if pd.notna(fs) and start_date <= fs.date() <= end_date:
                        results_seat = {"Fecha": fs.normalize(), "Total [kWh]": parse_latam_number(df_s.iloc[i, 3]), "Tracción [kWh]": parse_latam_number(df_s.iloc[i, 5]), "12 KV [kWh]": parse_latam_number(df_s.iloc[i, 7])}
                        all_seat.append(results_seat)

            if any(k in sn_up for k in ['PRMTE', 'MEDIDAS']):
                df_prm = pd.read_excel(f, sheet_name=sn, header=None)
                h_idx = next((i for i in range(len(df_prm)) if 'AÑO' in str(df_prm.iloc[i]).upper()), None)
                if h_idx is not None:
                    df_pd = pd.read_excel(f, sheet_name=sn, header=h_idx)
                    df_pd['Timestamp'] = pd.to_datetime(df_pd[['AÑO', 'MES', 'DIA', 'HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'})) + pd.to_timedelta(df_pd['INICIO INTERVALO'].astype(int), unit='m')
                    cols_e = [c for c in df_pd.columns if 'Retiro_Energia_Activa (kWhD)' in str(c)]
                    for _, r in df_pd.iterrows():
                        ts, val_p = r['Timestamp'], sum([parse_latam_number(r[col]) for col in cols_e])
                        all_comp_full.append({"Fecha": ts.normalize(), "Hora": ts.hour, "Consumo": val_p, "Fuente": "PRMTE"})
                        if start_date <= ts.date() <= end_date: all_prmte_15.append({"Fecha": ts.normalize(), "Energía PRMTE [kWh]": val_p})

            if any(k in sn_up for k in ['FACTURA', 'CONSUMO']):
                df_f = pd.read_excel(f, sheet_name=sn); df_f.columns = ['FechaHora', 'Valor']
                df_f['Timestamp'] = pd.to_datetime(df_f['FechaHora'], errors='coerce')
                for _, r in df_f.dropna(subset=['Timestamp']).iterrows():
                    ts, val_f = r['Timestamp'], abs(parse_latam_number(r['Valor']))
                    all_comp_full.append({"Fecha": ts.normalize(), "Hora": ts.hour, "Consumo": val_f, "Fuente": "Factura"})
                    if start_date <= ts.date() <= end_date: all_fact_h.append({"Fecha": ts.normalize(), "Consumo Horario [kWh]": val_f})
    except: continue

# --- 5. TRIANGULACIÓN Y DATA PRE-RENDER ---
df_ops, df_tr, df_tr_acum, df_seat = pd.DataFrame(all_ops), pd.DataFrame(all_tr), pd.DataFrame(all_tr_acum), pd.DataFrame(all_seat)
df_p_d = pd.DataFrame(all_prmte_15).groupby("Fecha")["Energía PRMTE [kWh]"].sum().reset_index() if all_prmte_15 else pd.DataFrame()
df_f_d = pd.DataFrame(all_fact_h).groupby("Fecha")["Consumo Horario [kWh]"].sum().reset_index() if all_fact_h else pd.DataFrame()

df_energy_master = pd.DataFrame()
if not df_seat.empty:
    df_energy_master = df_seat[["Fecha", "Total [kWh]"]].copy().rename(columns={"Total [kWh]":"E_Total"})
if not df_p_d.empty:
    p_sum = df_p_d.rename(columns={"Energía PRMTE [kWh]":"E_Total"})[["Fecha", "E_Total"]]
    df_energy_master = pd.concat([df_energy_master, p_sum]).drop_duplicates(subset="Fecha", keep="last")
if not df_f_d.empty:
    f_sum = df_f_d.rename(columns={"Consumo Horario [kWh]":"E_Total"})[["Fecha", "E_Total"]]
    df_energy_master = pd.concat([df_energy_master, f_sum]).drop_duplicates(subset="Fecha", keep="last")

if not df_ops.empty and not df_energy_master.empty:
    df_ops = pd.merge(df_ops, df_energy_master, on="Fecha", how="left")

if 'outliers' not in st.session_state: st.session_state.outliers = pd.DataFrame()

# --- 6. RENDERIZADO DE TABS ---
if not df_ops.empty or not df_seat.empty or all_comp_full:
    tabs = st.tabs(["📊 Resumen", "📑 Operaciones", "📑 Trenes", "⚡ Energía cruda", "⚖️ Comparación Energía hr", "📈 Regresión Nocturna", "🚨 Datos Atípicos"])
    
    with tabs[0]: # --- RESUMEN ---
        st.header("📊 Resumen Operacional")
        c1, c2, c3 = st.columns(3)
        anios = sorted(df_ops['Fecha'].dt.year.unique())
        f_ano = c1.multiselect("Año", anios, default=anios, key="res_a")
        f_mes = c2.multiselect("Mes", sorted(df_ops['Fecha'].dt.month.unique()), key="res_m")
        mask = df_ops['Fecha'].dt.year.isin(f_ano)
        if f_mes: mask &= df_ops['Fecha'].dt.month.isin(f_mes)
        df_res_f = df_ops[mask].copy()
        
        if not df_res_f.empty:
            to_val, tk_val = df_res_f["Odómetro [km]"].sum(), df_res_f["Tren-Km [km]"].sum()
            st.columns(3)[0].metric("Odómetro Total", f"{to_val:,.1f} km")
            st.columns(3)[1].metric("Tren-Km Total", f"{tk_val:,.1f} km")
            st.columns(3)[2].metric("UMR Global", f"{(tk_val/to_val*100) if to_val>0 else 0:.2f} %")
            
            st.write("#### Resumen por Jornada")
            res_j = df_res_f.groupby("Tipo Día", observed=True).agg({"Odómetro [km]":"sum", "Tren-Km [km]":"sum"}).reindex(ORDEN_JORNADA).dropna(how='all').reset_index()
            st.table(res_j.style.format({"Odómetro [km]":"{:,.1f}", "Tren-Km [km]":"{:,.1f}"}))

    with tabs[1]: # --- OPERACIONES ---
        st.header("📑 Datos Operacionales")
        st.dataframe(df_ops, use_container_width=True)

    with tabs[2]: # --- TRENES ---
        st.header("🚆 Control de Kilometraje")
        if not df_tr.empty:
            st.subheader("🚗 Kilometraje Diario")
            piv_d = df_tr.pivot_table(index="Tren", columns=df_tr["Fecha"].dt.day, values="Valor", aggfunc='sum')
            st.dataframe(piv_d.style.format("{:,.1f}"))
        if not df_tr_acum.empty:
            st.subheader("📈 Odómetro Acumulado")
            piv_a = df_tr_acum.pivot_table(index="Tren", columns=df_tr_acum["Fecha"].dt.day, values="Valor", aggfunc='max')
            st.dataframe(piv_a.style.format("{:,.0f}"))

    with tabs[3]: # --- ENERGÍA CRUDA ---
        st.header("⚡ Datos Base de Energía")
        s1, s2, s3 = st.tabs(["⚡ SEAT", "📈 PRMTE", "💰 Facturación"])
        with s1: st.dataframe(df_seat)
        with s2: st.dataframe(pd.DataFrame(all_prmte_15))
        with s3: st.dataframe(pd.DataFrame(all_fact_h))

    with tabs[4]: # --- COMPARACIÓN HORARIA ---
        st.header("⚖️ Comparación Energía hr")
        if all_comp_full:
            df_c = pd.DataFrame(all_comp_full).groupby(['Fecha','Hora','Fuente'])['Consumo'].sum().reset_index()
            fechas_f = df_c[df_c['Fuente']=='Factura']['Fecha'].unique()
            df_cf = df_c[~((df_c['Fuente']=='PRMTE') & (df_c['Fecha'].isin(fechas_f)))].copy()
            df_cf['Año'] = df_cf['Fecha'].dt.year.astype(str)
            df_cf['Tipo Día'] = df_cf['Fecha'].apply(get_tipo_dia)
            
            pivot_st = df_cf.pivot_table(index="Hora", columns=["Año", "Tipo Día"], values="Consumo", aggfunc='median').fillna(0)
            
            # Hora Total por Año
            for anio in sorted(df_cf['Año'].unique()):
                pivot_st[(anio, "Total Anual")] = df_cf[df_cf['Año'] == anio].groupby("Hora")["Consumo"].median()
            
            # Forzar Orden L, S, D/F, Total
            new_cols = []
            for anio in sorted(df_cf['Año'].unique()):
                for jor in ORDEN_JORNADA + ["Total Anual"]:
                    if (anio, jor) in pivot_st.columns: new_cols.append((anio, jor))
            
            pivot_st = pivot_st.reindex(columns=new_columns)
            st.dataframe(pivot_st.style.format("{:,.1f}"), use_container_width=True)

    with tabs[5]: # --- REGRESIÓN NOCTURNA ---
        st.header("📈 Regresión 12kV (00:00 - 05:00)")
        if all_comp_full:
            df_reg = pd.DataFrame(all_comp_full).groupby(['Fecha','Hora'])['Consumo'].sum().reset_index()
            df_reg = df_reg[df_reg['Hora'] <= 5]
            df_reg['Año'] = df_reg['Fecha'].dt.year
            df_reg['Tipo Día'] = df_reg['Fecha'].apply(get_tipo_dia)
            
            c1, c2 = st.columns(2)
            f_ra = c1.selectbox("Año", sorted(df_reg['Año'].unique()))
            f_rh = c2.selectbox("Hora", range(6))
            
            df_pl = df_reg[(df_reg['Año']==f_ra) & (df_reg['Hora']==f_rh)].sort_values("Fecha")
            if len(df_pl) > 2:
                Q1, Q3 = df_pl['Consumo'].quantile(0.25), df_pl['Consumo'].quantile(0.75)
                IQR = Q3 - Q1
                df_norm = df_pl[(df_pl['Consumo'] >= Q1 - 1.5*IQR) & (df_pl['Consumo'] <= Q3 + 1.5*IQR)].copy()
                st.session_state.outliers = df_pl[(df_pl['Consumo'] < Q1 - 1.5*IQR) | (df_pl['Consumo'] > Q3 + 1.5*IQR)]
                
                x, y = np.arange(len(df_norm)), df_norm['Consumo'].values
                m, n = np.polyfit(x, y, 1)
                r2 = np.corrcoef(x, y)[0,1]**2
                
                st.line_chart(pd.DataFrame({"Real":y, "Tendencia":m*x+n}, index=df_norm['Fecha'].dt.strftime('%d/%m')))
                st.metric("Intercepto (Carga Basal n)", f"{n:.2f} kWh")
                st.markdown(f"**Ecuación:** $y = {m:.4f}x + {n:.2f}$ | $R^2 = {r2:.4f}$")

    with tabs[6]: # --- ATÍPICOS ---
        st.header("🚨 Datos Atípicos Detectados")
        st.dataframe(st.session_state.outliers, use_container_width=True)
else:
    st.info("👋 Sube los archivos en el panel lateral para comenzar el análisis.")
