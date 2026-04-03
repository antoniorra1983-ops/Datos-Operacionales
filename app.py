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

# --- 3. CARGA Y MOTOR DE DATOS ---
with st.sidebar:
    st.header("📅 Filtro Global")
    today = date.today()
    start_of_month = today.replace(day=1) if today.day > 1 else (today.replace(month=today.month-1, day=1) if today.month>1 else today.replace(year=today.year-1, month=12, day=1))
    date_range = st.date_input("Selecciona el período", value=(start_of_month, today))
    start_date, end_date = (date_range[0], date_range[1]) if isinstance(date_range, tuple) and len(date_range)==2 else (date_range, date_range)
    st.divider()
    st.header("📂 Carga de Archivos")
    f_umr = st.file_uploader("1. UMR / Odómetros", type=["xlsx"], accept_multiple_files=True)
    f_seat_files = st.file_uploader("2. Energía SEAT", type=["xlsx"], accept_multiple_files=True)
    f_bill_files = st.file_uploader("3. Facturación y PRMTE", type=["xlsx"], accept_multiple_files=True)

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

# --- 4. JERARQUÍA Y PRE-FILTRADO ---
df_ops, df_tr, df_tr_acum, df_seat, df_energy_master = [pd.DataFrame()] * 5
df_p_d, df_f_d = pd.DataFrame(), pd.DataFrame()

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

    # --- 5. RENDERIZADO DE PESTAÑAS ---
    tabs = st.tabs(["📊 Resumen", "📑 Operaciones", "📑 Trenes", "⚡ Energía", "⚖️ Comparación Energía hr", "📈 Regresión Nocturna", "🚨 Datos Atípicos"])
    
    with tabs[0]: # Resumen - con filtros independientes por sub-pestaña
        if not df_ops.empty:
            # Sub-pestañas
            sub_tabs = st.tabs(["📅 Semanal", "📅 Mensual", "📅 Anual"])
            
            with sub_tabs[0]:  # Semanal
                st.write("##### Filtros Semanales")
                # Obtener años y semanas disponibles en todo df_ops (sin filtrar por fecha global aún)
                años_sem = sorted(df_ops['Fecha'].dt.year.unique())
                semanas_disponibles = sorted(df_ops['N° Semana'].unique())
                col_f1, col_f2, col_f3 = st.columns(3)
                año_sel = col_f1.selectbox("Año", años_sem, key="sem_año")
                semana_sel = col_f2.selectbox("N° Semana", semanas_disponibles, key="sem_semana")
                tipo_dia_sel = col_f3.multiselect("Tipo Día", ORDEN_TIPO_DIA, default=ORDEN_TIPO_DIA, key="sem_tipo")
                # Aplicar filtros
                mask = (df_ops['Fecha'].dt.year == año_sel) & (df_ops['N° Semana'] == semana_sel)
                if tipo_dia_sel:
                    mask &= df_ops['Tipo Día'].isin(tipo_dia_sel)
                df_semana = df_ops[mask].copy()
                if not df_semana.empty:
                    # Métricas
                    to_val = df_semana["Odómetro [km]"].sum()
                    tk_val = df_semana["Tren-Km [km]"].sum()
                    umr_val = (tk_val/to_val*100) if to_val>0 else 0
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Odómetro", f"{to_val:,.1f} km")
                    col2.metric("Tren-Km", f"{tk_val:,.1f} km")
                    col3.metric("UMR", f"{umr_val:.2f} %")
                    # Gráfico diario
                    df_graf = df_semana.sort_values('Fecha')
                    fig = make_subplots(specs=[[{"secondary_y": True}]])
                    fig.add_trace(go.Bar(x=df_graf['Fecha'].dt.strftime('%d/%m'), y=df_graf['Odómetro [km]'] / 1000, name='Odómetro (miles km)', marker_color='#005195'), secondary_y=False)
                    fig.add_trace(go.Bar(x=df_graf['Fecha'].dt.strftime('%d/%m'), y=df_graf['Tren-Km [km]'] / 1000, name='Tren-Km (miles km)', marker_color='#4CAF50'), secondary_y=False)
                    fig.add_trace(go.Scatter(x=df_graf['Fecha'].dt.strftime('%d/%m'), y=df_graf['UMR [%]'], name='UMR (%)', mode='lines+markers', line=dict(color='#FF5733', width=3), marker=dict(size=8)), secondary_y=True)
                    fig.update_layout(title=f"Semana {semana_sel} - {año_sel}", xaxis_title="Día del mes", barmode='group', legend_title="Métrica", height=400, hovermode='x unified')
                    fig.update_yaxes(title_text="Kilómetros (miles)", secondary_y=False)
                    fig.update_yaxes(title_text="UMR (%)", secondary_y=True, range=[0, 100])
                    st.plotly_chart(fig, use_container_width=True)
                    # Energía priorizada para la semana
                    st.markdown("#### ⚡ Energía (prioridad: Factura > PRMTE > SEAT)")
                    energia_fechas = []
                    for fecha in df_semana['Fecha'].unique():
                        if not df_f_d.empty and fecha in df_f_d['Fecha'].values:
                            row = df_f_d[df_f_d['Fecha'] == fecha].iloc[0]
                            energia_fechas.append({'Fecha': fecha, 'E_Total': row['Consumo Horario [kWh]'] if 'Consumo Horario [kWh]' in row else 0, 'E_Tr': row['E_Tr'] if 'E_Tr' in row else 0, 'E_12': row['E_12'] if 'E_12' in row else 0, 'Fuente': 'Factura'})
                        elif not df_p_d.empty and fecha in df_p_d['Fecha'].values:
                            row = df_p_d[df_p_d['Fecha'] == fecha].iloc[0]
                            energia_fechas.append({'Fecha': fecha, 'E_Total': row['Energía PRMTE [kWh]'] if 'Energía PRMTE [kWh]' in row else 0, 'E_Tr': row['E_Tr'] if 'E_Tr' in row else 0, 'E_12': row['E_12'] if 'E_12' in row else 0, 'Fuente': 'PRMTE'})
                        elif not df_seat.empty and fecha in df_seat['Fecha'].values:
                            row = df_seat[df_seat['Fecha'] == fecha].iloc[0]
                            energia_fechas.append({'Fecha': fecha, 'E_Total': row['Total [kWh]'], 'E_Tr': row['Tracción [kWh]'], 'E_12': row['12 KV [kWh]'], 'Fuente': 'SEAT'})
                        else:
                            energia_fechas.append({'Fecha': fecha, 'E_Total': 0, 'E_Tr': 0, 'E_12': 0, 'Fuente': 'Sin datos'})
                    df_energia = pd.DataFrame(energia_fechas)
                    total_energia = df_energia['E_Total'].sum()
                    total_traccion = df_energia['E_Tr'].sum()
                    total_12kv = df_energia['E_12'].sum()
                    fuente_principal = df_energia['Fuente'].iloc[0] if not df_energia.empty else "Sin datos"
                    col_e1, col_e2, col_e3, col_e4 = st.columns(4)
                    col_e1.metric("Energía Total", f"{total_energia:,.0f} kWh")
                    col_e2.metric("Energía Tracción", f"{total_traccion:,.0f} kWh")
                    col_e3.metric("Energía 12 kV", f"{total_12kv:,.0f} kWh")
                    col_e4.metric("Fuente principal", fuente_principal)
                    if total_energia > 0:
                        st.caption(f"⚡ Composición: Tracción {total_traccion/total_energia*100:.1f}% | 12 kV {total_12kv/total_energia*100:.1f}%")
                    # Resumen por jornada
                    res_j = df_semana.groupby("Tipo Día", observed=True).agg({"Odómetro [km]":"sum", "Tren-Km [km]":"sum", "UMR [%]":"mean"}).reset_index()
                    res_j['Tipo Día'] = pd.Categorical(res_j['Tipo Día'], categories=ORDEN_TIPO_DIA, ordered=True)
                    res_j = res_j.sort_values('Tipo Día').reset_index(drop=True)
                    st.write("#### Resumen por Jornada")
                    st.table(res_j.style.format({"Odómetro [km]":"{:,.1f}", "Tren-Km [km]":"{:,.1f}", "UMR [%]":"{:.2f}%"}))
                else:
                    st.info("No hay datos para los filtros seleccionados.")
            
            with sub_tabs[1]:  # Mensual
                st.write("##### Filtros Mensuales")
                años_mes = sorted(df_ops['Fecha'].dt.year.unique())
                meses_numeros = sorted(df_ops['Fecha'].dt.month.unique())
                col_f1, col_f2, col_f3 = st.columns(3)
                año_mes_sel = col_f1.selectbox("Año", años_mes, key="mes_año")
                mes_sel = col_f2.selectbox("Mes", meses_numeros, format_func=lambda x: f"{x:02d}", key="mes_mes")
                tipo_dia_mes_sel = col_f3.multiselect("Tipo Día", ORDEN_TIPO_DIA, default=ORDEN_TIPO_DIA, key="mes_tipo")
                mask = (df_ops['Fecha'].dt.year == año_mes_sel) & (df_ops['Fecha'].dt.month == mes_sel)
                if tipo_dia_mes_sel:
                    mask &= df_ops['Tipo Día'].isin(tipo_dia_mes_sel)
                df_mes = df_ops[mask].copy()
                if not df_mes.empty:
                    # Métricas
                    to_val = df_mes["Odómetro [km]"].sum()
                    tk_val = df_mes["Tren-Km [km]"].sum()
                    umr_val = (tk_val/to_val*100) if to_val>0 else 0
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Odómetro", f"{to_val:,.1f} km")
                    col2.metric("Tren-Km", f"{tk_val:,.1f} km")
                    col3.metric("UMR", f"{umr_val:.2f} %")
                    # Gráfico diario
                    df_graf = df_mes.sort_values('Fecha')
                    fig = make_subplots(specs=[[{"secondary_y": True}]])
                    fig.add_trace(go.Bar(x=df_graf['Fecha'].dt.strftime('%d/%m'), y=df_graf['Odómetro [km]'] / 1000, name='Odómetro (miles km)', marker_color='#005195'), secondary_y=False)
                    fig.add_trace(go.Bar(x=df_graf['Fecha'].dt.strftime('%d/%m'), y=df_graf['Tren-Km [km]'] / 1000, name='Tren-Km (miles km)', marker_color='#4CAF50'), secondary_y=False)
                    fig.add_trace(go.Scatter(x=df_graf['Fecha'].dt.strftime('%d/%m'), y=df_graf['UMR [%]'], name='UMR (%)', mode='lines+markers', line=dict(color='#FF5733', width=3), marker=dict(size=8)), secondary_y=True)
                    fig.update_layout(title=f"Mes {mes_sel:02d} - {año_mes_sel}", xaxis_title="Día del mes", barmode='group', legend_title="Métrica", height=400, hovermode='x unified')
                    fig.update_yaxes(title_text="Kilómetros (miles)", secondary_y=False)
                    fig.update_yaxes(title_text="UMR (%)", secondary_y=True, range=[0, 100])
                    st.plotly_chart(fig, use_container_width=True)
                    # Energía priorizada
                    st.markdown("#### ⚡ Energía (prioridad: Factura > PRMTE > SEAT)")
                    energia_fechas = []
                    for fecha in df_mes['Fecha'].unique():
                        if not df_f_d.empty and fecha in df_f_d['Fecha'].values:
                            row = df_f_d[df_f_d['Fecha'] == fecha].iloc[0]
                            energia_fechas.append({'Fecha': fecha, 'E_Total': row['Consumo Horario [kWh]'] if 'Consumo Horario [kWh]' in row else 0, 'E_Tr': row['E_Tr'] if 'E_Tr' in row else 0, 'E_12': row['E_12'] if 'E_12' in row else 0, 'Fuente': 'Factura'})
                        elif not df_p_d.empty and fecha in df_p_d['Fecha'].values:
                            row = df_p_d[df_p_d['Fecha'] == fecha].iloc[0]
                            energia_fechas.append({'Fecha': fecha, 'E_Total': row['Energía PRMTE [kWh]'] if 'Energía PRMTE [kWh]' in row else 0, 'E_Tr': row['E_Tr'] if 'E_Tr' in row else 0, 'E_12': row['E_12'] if 'E_12' in row else 0, 'Fuente': 'PRMTE'})
                        elif not df_seat.empty and fecha in df_seat['Fecha'].values:
                            row = df_seat[df_seat['Fecha'] == fecha].iloc[0]
                            energia_fechas.append({'Fecha': fecha, 'E_Total': row['Total [kWh]'], 'E_Tr': row['Tracción [kWh]'], 'E_12': row['12 KV [kWh]'], 'Fuente': 'SEAT'})
                        else:
                            energia_fechas.append({'Fecha': fecha, 'E_Total': 0, 'E_Tr': 0, 'E_12': 0, 'Fuente': 'Sin datos'})
                    df_energia = pd.DataFrame(energia_fechas)
                    total_energia = df_energia['E_Total'].sum()
                    total_traccion = df_energia['E_Tr'].sum()
                    total_12kv = df_energia['E_12'].sum()
                    fuente_principal = df_energia['Fuente'].iloc[0] if not df_energia.empty else "Sin datos"
                    col_e1, col_e2, col_e3, col_e4 = st.columns(4)
                    col_e1.metric("Energía Total", f"{total_energia:,.0f} kWh")
                    col_e2.metric("Energía Tracción", f"{total_traccion:,.0f} kWh")
                    col_e3.metric("Energía 12 kV", f"{total_12kv:,.0f} kWh")
                    col_e4.metric("Fuente principal", fuente_principal)
                    if total_energia > 0:
                        st.caption(f"⚡ Composición: Tracción {total_traccion/total_energia*100:.1f}% | 12 kV {total_12kv/total_energia*100:.1f}%")
                    # Resumen por jornada
                    res_j = df_mes.groupby("Tipo Día", observed=True).agg({"Odómetro [km]":"sum", "Tren-Km [km]":"sum", "UMR [%]":"mean"}).reset_index()
                    res_j['Tipo Día'] = pd.Categorical(res_j['Tipo Día'], categories=ORDEN_TIPO_DIA, ordered=True)
                    res_j = res_j.sort_values('Tipo Día').reset_index(drop=True)
                    st.write("#### Resumen por Jornada")
                    st.table(res_j.style.format({"Odómetro [km]":"{:,.1f}", "Tren-Km [km]":"{:,.1f}", "UMR [%]":"{:.2f}%"}))
                else:
                    st.info("No hay datos para los filtros seleccionados.")
            
            with sub_tabs[2]:  # Anual
                st.write("##### Filtros Anuales")
                años_anio = sorted(df_ops['Fecha'].dt.year.unique())
                col_f1, col_f2 = st.columns(2)
                año_anio_sel = col_f1.selectbox("Año", años_anio, key="anio_año")
                tipo_dia_anio_sel = col_f2.multiselect("Tipo Día", ORDEN_TIPO_DIA, default=ORDEN_TIPO_DIA, key="anio_tipo")
                mask = (df_ops['Fecha'].dt.year == año_anio_sel)
                if tipo_dia_anio_sel:
                    mask &= df_ops['Tipo Día'].isin(tipo_dia_anio_sel)
                df_anio = df_ops[mask].copy()
                if not df_anio.empty:
                    # Agrupar por mes
                    df_mensual = df_anio.groupby(df_anio['Fecha'].dt.month).agg({
                        'Odómetro [km]': 'sum',
                        'Tren-Km [km]': 'sum',
                        'UMR [%]': 'mean'
                    }).reset_index()
                    df_mensual.columns = ['Mes', 'Odómetro [km]', 'Tren-Km [km]', 'UMR [%]']
                    # Métricas anuales
                    to_val = df_mensual['Odómetro [km]'].sum()
                    tk_val = df_mensual['Tren-Km [km]'].sum()
                    umr_val = (tk_val/to_val*100) if to_val>0 else 0
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Odómetro", f"{to_val:,.1f} km")
                    col2.metric("Tren-Km", f"{tk_val:,.1f} km")
                    col3.metric("UMR", f"{umr_val:.2f} %")
                    # Gráfico mensual
                    fig = make_subplots(specs=[[{"secondary_y": True}]])
                    fig.add_trace(go.Bar(x=df_mensual['Mes'].astype(str), y=df_mensual['Odómetro [km]'] / 1000, name='Odómetro (miles km)', marker_color='#005195'), secondary_y=False)
                    fig.add_trace(go.Bar(x=df_mensual['Mes'].astype(str), y=df_mensual['Tren-Km [km]'] / 1000, name='Tren-Km (miles km)', marker_color='#4CAF50'), secondary_y=False)
                    fig.add_trace(go.Scatter(x=df_mensual['Mes'].astype(str), y=df_mensual['UMR [%]'], name='UMR (%)', mode='lines+markers', line=dict(color='#FF5733', width=3), marker=dict(size=8)), secondary_y=True)
                    fig.update_layout(title=f"Año {año_anio_sel}", xaxis_title="Mes", barmode='group', legend_title="Métrica", height=400, hovermode='x unified')
                    fig.update_yaxes(title_text="Kilómetros (miles)", secondary_y=False)
                    fig.update_yaxes(title_text="UMR (%)", secondary_y=True, range=[0, 100])
                    st.plotly_chart(fig, use_container_width=True)
                    # Energía priorizada (anual)
                    st.markdown("#### ⚡ Energía (prioridad: Factura > PRMTE > SEAT)")
                    energia_fechas = []
                    for fecha in df_anio['Fecha'].unique():
                        if not df_f_d.empty and fecha in df_f_d['Fecha'].values:
                            row = df_f_d[df_f_d['Fecha'] == fecha].iloc[0]
                            energia_fechas.append({'Fecha': fecha, 'E_Total': row['Consumo Horario [kWh]'] if 'Consumo Horario [kWh]' in row else 0, 'E_Tr': row['E_Tr'] if 'E_Tr' in row else 0, 'E_12': row['E_12'] if 'E_12' in row else 0, 'Fuente': 'Factura'})
                        elif not df_p_d.empty and fecha in df_p_d['Fecha'].values:
                            row = df_p_d[df_p_d['Fecha'] == fecha].iloc[0]
                            energia_fechas.append({'Fecha': fecha, 'E_Total': row['Energía PRMTE [kWh]'] if 'Energía PRMTE [kWh]' in row else 0, 'E_Tr': row['E_Tr'] if 'E_Tr' in row else 0, 'E_12': row['E_12'] if 'E_12' in row else 0, 'Fuente': 'PRMTE'})
                        elif not df_seat.empty and fecha in df_seat['Fecha'].values:
                            row = df_seat[df_seat['Fecha'] == fecha].iloc[0]
                            energia_fechas.append({'Fecha': fecha, 'E_Total': row['Total [kWh]'], 'E_Tr': row['Tracción [kWh]'], 'E_12': row['12 KV [kWh]'], 'Fuente': 'SEAT'})
                        else:
                            energia_fechas.append({'Fecha': fecha, 'E_Total': 0, 'E_Tr': 0, 'E_12': 0, 'Fuente': 'Sin datos'})
                    df_energia = pd.DataFrame(energia_fechas)
                    total_energia = df_energia['E_Total'].sum()
                    total_traccion = df_energia['E_Tr'].sum()
                    total_12kv = df_energia['E_12'].sum()
                    fuente_principal = df_energia['Fuente'].iloc[0] if not df_energia.empty else "Sin datos"
                    col_e1, col_e2, col_e3, col_e4 = st.columns(4)
                    col_e1.metric("Energía Total", f"{total_energia:,.0f} kWh")
                    col_e2.metric("Energía Tracción", f"{total_traccion:,.0f} kWh")
                    col_e3.metric("Energía 12 kV", f"{total_12kv:,.0f} kWh")
                    col_e4.metric("Fuente principal", fuente_principal)
                    if total_energia > 0:
                        st.caption(f"⚡ Composición: Tracción {total_traccion/total_energia*100:.1f}% | 12 kV {total_12kv/total_energia*100:.1f}%")
                    # Resumen por jornada (anual)
                    res_j = df_anio.groupby("Tipo Día", observed=True).agg({"Odómetro [km]":"sum", "Tren-Km [km]":"sum", "UMR [%]":"mean"}).reset_index()
                    res_j['Tipo Día'] = pd.Categorical(res_j['Tipo Día'], categories=ORDEN_TIPO_DIA, ordered=True)
                    res_j = res_j.sort_values('Tipo Día').reset_index(drop=True)
                    st.write("#### Resumen por Jornada")
                    st.table(res_j.style.format({"Odómetro [km]":"{:,.1f}", "Tren-Km [km]":"{:,.1f}", "UMR [%]":"{:.2f}%"}))
                else:
                    st.info("No hay datos para los filtros seleccionados.")
            
            # --- Botones de exportación comunes para Resumen (opcional) ---
            st.write("---")
            st.write("#### 📥 Exportar datos actuales (según filtros de la pestaña activa)")
            # Nota: Esto exportaría los datos del último filtro seleccionado, pero puede complicarse. Por simplicidad, exportamos el DataFrame filtrado de la sub-pestaña actual.
            # Se puede omitir o dejar como estaba. Lo dejamos como botón genérico.
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
                    # Exportar el último DataFrame usado (se puede mejorar)
                    st.info("Para exportar datos específicos, use los botones dentro de cada sub-pestaña (no implementado).")
                    # Opcional: exportar todo el df_ops filtrado por fecha global
                    # Por ahora, no implementamos exportación compleja.
        else:
            st.info("No hay datos de operaciones cargados.")

    with tabs[1]: # Operaciones (sin filtros compartidos, usar filtros propios o ninguno)
        if not df_ops.empty:
            # Filtros simples para Operaciones (año, mes, semana, tipo día)
            st.write("#### Filtros de Operaciones")
            c1, c2, c3, c4 = st.columns(4)
            años_op = sorted(df_ops['Fecha'].dt.year.unique())
            meses_op = sorted(df_ops['Fecha'].dt.month.unique())
            semanas_op = sorted(df_ops['N° Semana'].unique())
            tipos_op = df_ops['Tipo Día'].unique()
            tipos_op_ord = [d for d in ORDEN_TIPO_DIA if d in tipos_op]
            f_ano_op = c1.multiselect("Año", años_op, default=años_op, key="op_a")
            f_mes_op = c2.multiselect("Mes", meses_op, default=meses_op, key="op_m")
            f_sem_op = c3.multiselect("N° Semana", semanas_op, default=semanas_op, key="op_s")
            f_tipo_op = c4.multiselect("Tipo Día", tipos_op_ord, default=tipos_op_ord, key="op_t")
            mask = (df_ops['Fecha'].dt.year.isin(f_ano_op)) & (df_ops['Fecha'].dt.month.isin(f_mes_op))
            if f_sem_op:
                mask &= df_ops['N° Semana'].isin(f_sem_op)
            if f_tipo_op:
                mask &= df_ops['Tipo Día'].isin(f_tipo_op)
            df_ops_f = df_ops[mask]
            st.dataframe(df_ops_f, use_container_width=True)
            st.download_button("📥 Descargar Operaciones (PPTX)", to_pptx("Datos Operacionales", df_ops_f), "EFE_Operaciones.pptx")
        else:
            st.info("No hay datos de operaciones para mostrar.")

    with tabs[2]: # Trenes (sin cambios)
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

    with tabs[3]: # Energía (sin cambios)
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

    with tabs[4]: # Comparación hr (sin cambios)
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

    with tabs[5]: # Regresión (sin cambios)
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

    with tabs[6]: # Atípicos (sin cambios)
        if not st.session_state.outliers.empty:
            st.error(f"Se detectaron {len(st.session_state.outliers)} anomalías.")
            st.dataframe(st.session_state.outliers, use_container_width=True)
            csv = st.session_state.outliers.to_csv(index=False).encode('utf-8')
            st.download_button("📥 Descargar CSV", csv, "Anomalias.csv", "text/csv")
            st.download_button("📥 Descargar Atípicos (PPTX)", to_pptx("Datos Atípicos de Instalaciones", st.session_state.outliers), "EFE_Atipicos.pptx")
        else:
            st.success("No hay anomalías detectadas en la selección actual.")

    st.sidebar.download_button("📥 Reporte Excel Completo", to_excel_consolidado(df_ops, df_tr, df_tr_acum, df_seat, df_p_d, pd.DataFrame(all_prmte_15), pd.DataFrame(all_fact_h), df_f_d), "Reporte_EFE_SGE.xlsx")
else:
    st.info("👋 Sube los archivos en el panel lateral para comenzar el análisis.")
