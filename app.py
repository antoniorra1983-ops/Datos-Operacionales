import streamlit as st
import pandas as pd
import numpy as np
import re
import holidays
from io import BytesIO
from datetime import datetime

# --- 1. CONFIGURACIÓN Y ESTILOS ---
st.set_page_config(page_title="Gestión de Energía - Dashboard SGE", layout="wide", page_icon="🚆")
chile_holidays = holidays.Chile()

st.markdown("""
    <style>
    .stMetric { background-color: #ffffff; padding: 20px; border-radius: 10px; border-left: 5px solid #005195; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    </style>
    """, unsafe_allow_html=True)

# --- 2. FUNCIONES DE PROCESAMIENTO ---
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

def to_excel_consolidado(df_ops, df_tr, df_tr_acum, df_seat, df_prm_d, df_prm_15, df_fact_h, df_fact_d):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dfs = {
            'Operaciones': df_ops, 'Kms_Diarios_Tren': df_tr, 'Odometros_Acum_Tren': df_tr_acum,
            'SEAT': df_seat, 'PRMTE_D': df_prm_d, 'PRMTE_15': df_prm_15, 
            'Fact_H': df_fact_h, 'Fact_D': df_fact_d
        }
        for name, df in dfs.items():
            if not df.empty: df.to_excel(writer, index=False, sheet_name=name)
    return output.getvalue()

# --- 3. CARGA Y FILTROS ---
with st.sidebar:
    st.header("📂 Carga de Archivos")
    f_umr = st.file_uploader("1. UMR / Odómetros", type=["xlsx"], accept_multiple_files=True)
    f_seat_files = st.file_uploader("2. Energía SEAT", type=["xlsx"], accept_multiple_files=True)
    f_bill_files = st.file_uploader("3. Facturación y PRMTE", type=["xlsx"], accept_multiple_files=True)
    st.divider()
    f_anio_list = st.multiselect("Años", [2024, 2025, 2026], default=[2025, 2026])
    meses_nombres = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    f_mes_list = st.multiselect("Meses", meses_nombres, default=meses_nombres)
    meses_num = [i+1 for i, m in enumerate(meses_nombres) if m in f_mes_list]

# --- 4. MOTOR DE DATOS ---
all_ops, all_tr, all_tr_acum, all_seat, all_prmte_15, all_fact_h = [], [], [], [], [], []
todos = (f_umr or []) + (f_seat_files or []) + (f_bill_files or [])

for f in todos:
    try:
        xl = pd.ExcelFile(f)
        for sn in xl.sheet_names:
            sn_up = sn.upper()
            
            # --- A. UMR / OPERACIONES ---
            if any(k in sn_up for k in ['UMR', 'RESUMEN']):
                df_raw = pd.read_excel(f, sheet_name=sn, header=None)
                h_r = next((i for i in range(min(100, len(df_raw))) if any(k in str(df_raw.iloc[i]).upper() for k in ['ODO', 'FECHA'])), None)
                if h_r is not None:
                    df_p = pd.read_excel(f, sheet_name=sn, header=h_r)
                    df_p.columns = [re.sub(r'[^A-Z]', '', str(c).upper().replace('Ó','O')) for c in df_p.columns]
                    idx_f, idx_o, idx_t = next((c for c in df_p.columns if 'FECHA' in c), None), next((c for c in df_p.columns if 'ODO' in c and 'ACUM' not in c), None), next((c for c in df_p.columns if 'TREN' in c and 'KM' in c), None)
                    if idx_f and idx_o:
                        df_p['_dt'] = pd.to_datetime(df_p[idx_f], errors='coerce')
                        mask = (df_p['_dt'].dt.year.isin(f_anio_list)) & (df_p['_dt'].dt.month.isin(meses_num))
                        for _, r in df_p[mask].dropna(subset=['_dt']).iterrows():
                            o, tk = parse_latam_number(r[idx_o]), parse_latam_number(r[idx_t])
                            if o > 0: all_ops.append({"Fecha": r['_dt'].normalize(), "Tipo Día": get_tipo_dia(r['_dt']), "N° Semana": r['_dt'].isocalendar()[1], "Odómetro [km]": o, "Tren-Km [km]": tk, "UMR [%]": (tk/o*100)})

            # --- B. ODÓMETRO POR TREN (TABLA DIARIA + TABLA ACUMULADA) ---
            if 'ODO' in sn_up and 'KIL' in sn_up:
                df_tr_raw = pd.read_excel(f, sheet_name=sn, header=None)
                headers_found = []
                # Buscamos todos los bloques de fechas en la hoja
                for i in range(len(df_tr_raw)-2):
                    for j in range(1, len(df_tr_raw.columns)):
                        val = pd.to_datetime(df_tr_raw.iloc[i, j], errors='coerce')
                        if pd.notna(val) and val.year in f_anio_list and val.month in meses_num:
                            if i not in [h[0] for h in headers_found]:
                                headers_found.append((i, val)) # Guardamos fila de inicio de tabla

                for idx, (row_idx, start_date) in enumerate(headers_found):
                    # Identificar si es tabla diaria o acumulada por texto cercano
                    context_text = str(df_tr_raw.iloc[row_idx:row_idx+3, 0:5]).upper()
                    is_acum = any(k in context_text for k in ['ACUM', 'LECTURA', 'TOTAL'])
                    
                    # Mapeo de columnas para este bloque
                    c_map = {}
                    for j in range(1, len(df_tr_raw.columns)):
                        dt = pd.to_datetime(df_tr_raw.iloc[row_idx, j], errors='coerce')
                        if pd.notna(dt): c_map[j] = dt

                    for k in range(row_idx+3, min(row_idx+40, len(df_tr_raw))):
                        n_tr = str(df_tr_raw.iloc[k, 0]).strip().upper()
                        if re.match(r'^(M|XM)', n_tr):
                            for c_idx, c_fch in c_map.items():
                                val_km = parse_latam_number(df_tr_raw.iloc[k, c_idx])
                                data_point = {"Tren": n_tr, "Fecha": c_fch.normalize(), "Día": c_fch.day, "Valor": val_km}
                                if is_acum or idx > 0: all_tr_acum.append(data_point)
                                else: all_tr.append(data_point)

            # --- C. ENERGÍA SEAT ---
            if 'SEAT' in sn_up and 'SER' in sn_up:
                df_s = pd.read_excel(f, sheet_name=sn, header=None)
                for i in range(len(df_s)):
                    fs = pd.to_datetime(df_s.iloc[i, 1], errors='coerce')
                    if pd.notna(fs) and fs.year in f_anio_list and fs.month in meses_num:
                        tot, tra, k12 = parse_latam_number(df_s.iloc[i, 3]), parse_latam_number(df_s.iloc[i, 5]), parse_latam_number(df_s.iloc[i, 7])
                        all_seat.append({"Fecha": fs.normalize(), "Total [kWh]": tot, "Tracción [kWh]": tra, "12 KV [kWh]": k12, "% Tracción": (tra/tot*100 if tot>0 else 0), "% 12 KV": (k12/tot*100 if tot>0 else 0)})

            # --- D. PRMTE / FACTURA (Lógica previa mantenida) ---
            if any(k in sn_up for k in ['PRMTE', 'MEDIDAS']):
                df_prm = pd.read_excel(f, sheet_name=sn, header=None)
                h_idx = next((i for i in range(len(df_prm)) if 'AÑO' in str(df_prm.iloc[i]).upper()), None)
                if h_idx is not None:
                    df_prm_d = pd.read_excel(f, sheet_name=sn, header=h_idx)
                    df_prm_d['Timestamp'] = pd.to_datetime(df_prm_d[['AÑO', 'MES', 'DIA', 'HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'})) + pd.to_timedelta(df_prm_d['INICIO INTERVALO'].astype(int), unit='m')
                    for _, r in df_prm_d[df_prm_d['Timestamp'].dt.month.isin(meses_num)].iterrows():
                        all_prmte_15.append({"Fecha y Hora": r['Timestamp'].strftime('%d/%m/%Y %H:%M'), "Fecha": r['Timestamp'].normalize(), "Energía PRMTE [kWh]": parse_latam_number(r['Retiro_Energia_Activa (kWhD)'])})

            if any(k in sn_up for k in ['FACTURA', 'CONSUMO']):
                df_f = pd.read_excel(f, sheet_name=sn); df_f.columns = ['FechaHora', 'Valor']
                df_f['Timestamp'] = pd.to_datetime(df_f['FechaHora'], errors='coerce')
                for _, r in df_f[df_f['Timestamp'].dt.month.isin(meses_num)].dropna(subset=['Timestamp']).iterrows():
                    all_fact_h.append({"Fecha y Hora": r['Timestamp'].strftime('%d/%m/%Y %H:%M'), "Fecha": r['Timestamp'].normalize(), "Consumo Horario [kWh]": abs(parse_latam_number(r['Valor']))})
    except: continue

# --- 5. JERARQUÍA Y RENDERIZADO ---
df_ops, df_tr, df_tr_acum, df_seat, df_energy_master = [pd.DataFrame()] * 5

if any([all_ops, all_tr, all_tr_acum, all_seat, all_prmte_15, all_fact_h]):
    if all_ops: df_ops = pd.DataFrame(all_ops).drop_duplicates(subset=['Fecha']).sort_values("Fecha")
    if all_tr: df_tr = pd.DataFrame(all_tr).sort_values(["Fecha", "Tren"])
    if all_tr_acum: df_tr_acum = pd.DataFrame(all_tr_acum).sort_values(["Fecha", "Tren"])
    if all_seat: df_seat = pd.DataFrame(all_seat).drop_duplicates(subset=['Fecha']).sort_values("Fecha")
    
    # Lógica de Jerarquía de Energía (Mantenida)
    if not df_seat.empty:
        df_energy_master = df_seat[["Fecha", "Total [kWh]", "Tracción [kWh]", "12 KV [kWh]"]].copy().rename(columns={"Total [kWh]":"E_Total", "Tracción [kWh]":"E_Tr", "12 KV [kWh]":"E_12"})
        df_energy_master["Fuente"] = "SEAT"

    if all_prmte_15:
        df_p_d = pd.DataFrame(all_prmte_15).groupby("Fecha")["Energía PRMTE [kWh]"].sum().reset_index()
        if not df_seat.empty:
            df_p_d = pd.merge(df_p_d, df_seat[["Fecha", "% Tracción", "% 12 KV"]], on="Fecha", how="left")
            df_p_d["E_Tr"], df_p_d["E_12"] = df_p_d["Energía PRMTE [kWh]"]*(df_p_d["% Tracción"]/100), df_p_d["Energía PRMTE [kWh]"]*(df_p_d["% 12 KV"]/100)
            df_p_p = df_p_d.rename(columns={"Energía PRMTE [kWh]":"E_Total"})[["Fecha","E_Total","E_Tr","E_12"]]; df_p_p["Fuente"] = "PRMTE"
            df_energy_master = pd.concat([df_energy_master, df_p_p]).drop_duplicates(subset=["Fecha"], keep="last")

    if all_fact_h:
        df_f_d = pd.DataFrame(all_fact_h).groupby("Fecha")["Consumo Horario [kWh]"].sum().reset_index()
        if not df_seat.empty:
            df_f_d = pd.merge(df_f_d, df_seat[["Fecha", "% Tracción", "% 12 KV"]], on="Fecha", how="left")
            df_f_d["E_Tr"], df_f_d["E_12"] = df_f_d["Consumo Horario [kWh]"]*(df_f_d["% Tracción"]/100), df_f_d["Consumo Horario [kWh]"]*(df_f_d["% 12 KV"]/100)
            df_f_f = df_f_d.rename(columns={"Consumo Horario [kWh]":"E_Total"})[["Fecha","E_Total","E_Tr","E_12"]]; df_f_f["Fuente"] = "Factura"
            df_energy_master = pd.concat([df_energy_master, df_f_f]).drop_duplicates(subset=["Fecha"], keep="last")

    if not df_ops.empty and not df_energy_master.empty:
        df_ops = pd.merge(df_ops, df_energy_master, on="Fecha", how="left")

    tabs = st.tabs(["📊 Resumen", "📑 Datos operacionales", "📑 Odómetro por Tren", "⚡ Energía SEAT", "📈 Medidas PRMTE", "💰 Facturación"])
    
    with tabs[0]: # Resumen
        if not df_ops.empty:
            res = df_ops.groupby("Tipo Dia" if "Tipo Dia" in df_ops else "Tipo Día").agg({"Odómetro [km]":"sum", "Tren-Km [km]":"sum", "UMR [%]":"mean", "E_Total":"sum", "E_Tr":"sum", "E_12":"sum"}).reset_index()
            st.table(res.style.format({"Odómetro [km]":"{:,.1f}", "E_Total":"{:,.0f}"}))

    with tabs[2]: # Odómetro por Tren
        if not df_tr.empty:
            st.write("### 🚗 Kilometraje Diario [km]")
            st.dataframe(df_tr.pivot_table(index="Tren", columns=df_tr["Fecha"].dt.day, values="Valor", aggfunc='sum').fillna(0).style.format("{:,.1f}"), use_container_width=True)
        
        if not df_tr_acum.empty:
            st.divider()
            st.write("### 📈 Lectura de Odómetro / Acumulado [km]")
            st.dataframe(df_tr_acum.pivot_table(index="Tren", columns=df_tr_acum["Fecha"].dt.day, values="Valor", aggfunc='max').fillna(0).style.format("{:,.0f}"), use_container_width=True)

    st.sidebar.download_button("📥 Descargar Reporte", to_excel_consolidado(df_ops, df_tr, df_tr_acum, df_seat, pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()), "Reporte_EFE_SGE.xlsx")
