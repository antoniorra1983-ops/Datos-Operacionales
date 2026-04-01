import streamlit as st
import pandas as pd
import numpy as np
import re
import holidays
from io import BytesIO
from datetime import datetime, date

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
    st.header("📅 Filtro Global")
    today = date.today()
    if today.day == 1:
        start_of_month = today.replace(month=today.month-1 if today.month > 1 else 12, year=today.year if today.month > 1 else today.year-1, day=1)
    else:
        start_of_month = today.replace(day=1)
        
    date_range = st.date_input("Selecciona el período de análisis", value=(start_of_month, today))
    
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = date_range
    else:
        start_date = end_date = (date_range[0] if isinstance(date_range, tuple) else date_range)

    st.divider()
    st.header("📂 Carga de Archivos")
    f_umr = st.file_uploader("1. UMR / Odómetros", type=["xlsx"], accept_multiple_files=True)
    f_seat_files = st.file_uploader("2. Energía SEAT", type=["xlsx"], accept_multiple_files=True)
    f_bill_files = st.file_uploader("3. Facturación y PRMTE", type=["xlsx"], accept_multiple_files=True)

# --- 4. MOTOR DE DATOS ---
all_ops, all_tr, all_tr_acum, all_seat, all_prmte_15, all_fact_h = [], [], [], [], [], []
all_comp_full = [] 

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
                            o, tk = parse_latam_number(r[idx_o]), parse_latam_number(r[idx_t])
                            if o > 0: all_ops.append({"Fecha": r['_dt'].normalize(), "Tipo Día": get_tipo_dia(r['_dt']), "N° Semana": r['_dt'].isocalendar()[1], "Odómetro [km]": o, "Tren-Km [km]": tk, "UMR [%]": (tk/o*100)})

            if 'ODO' in sn_up and 'KIL' in sn_up:
                df_tr_raw = pd.read_excel(f, sheet_name=sn, header=None)
                headers_found = []
                for i in range(len(df_tr_raw)-2):
                    for j in range(1, len(df_tr_raw.columns)):
                        val = pd.to_datetime(df_tr_raw.iloc[i, j], errors='coerce')
                        if pd.notna(val) and start_date <= val.date() <= end_date:
                            if i not in [h[0] for h in headers_found]:
                                headers_found.append((i, val))

                for idx, (row_idx, s_dt) in enumerate(headers_found):
                    context_text = str(df_tr_raw.iloc[row_idx:row_idx+3, 0:5]).upper()
                    is_acum = any(k in context_text for k in ['ACUM', 'LECTURA', 'TOTAL'])
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
                    df_prm_d = pd.read_excel(f, sheet_name=sn, header=h_idx)
                    df_prm_d['Timestamp'] = pd.to_datetime(df_prm_d[['AÑO', 'MES', 'DIA', 'HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'})) + pd.to_timedelta(df_prm_d['INICIO INTERVALO'].astype(int), unit='m')
                    cols_energia = [c for c in df_prm_d.columns if 'Retiro_Energia_Activa (kWhD)' in str(c)]
                    for _, r in df_prm_d.iterrows():
                        ts = r['Timestamp']
                        suma_prmte = sum([parse_latam_number(r[col]) for col in cols_energia])
                        all_comp_full.append({"Fecha": ts.normalize(), "Hora": ts.hour, "Consumo Horario [kWh]": suma_prmte, "Fuente": "PRMTE"})
                        if start_date <= ts.date() <= end_date:
                            all_prmte_15.append({"Fecha y Hora": ts.strftime('%d/%m/%Y %H:%M'), "Fecha": ts.normalize(), "Energía PRMTE [kWh]": suma_prmte})

            if any(k in sn_up for k in ['FACTURA', 'CONSUMO']):
                df_f = pd.read_excel(f, sheet_name=sn); df_f.columns = ['FechaHora', 'Valor']
                df_f['Timestamp'] = pd.to_datetime(df_f['FechaHora'], errors='coerce')
                for _, r in df_f.dropna(subset=['Timestamp']).iterrows():
                    ts = r['Timestamp']
                    val = abs(parse_latam_number(r['Valor']))
                    all_comp_full.append({"Fecha": ts.normalize(), "Hora": ts.hour, "Consumo Horario [kWh]": val, "Fuente": "Factura"})
                    if start_date <= ts.date() <= end_date:
                        all_fact_h.append({"Fecha y Hora": ts.strftime('%d/%m/%Y %H:%M'), "Fecha": ts.normalize(), "Consumo Horario [kWh]": val})
    except: continue

# --- 5. JERARQUÍA Y PRE-FILTRADO ---
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

    def get_filtros(df, prefijo):
        if df.empty: return df
        c1, c2, c3 = st.columns(3)
        anios, meses = sorted(df['Fecha'].dt.year.unique()), sorted(df['Fecha'].dt.month.unique())
        f_ano, f_mes = c1.multiselect(f"Año", anios, default=anios, key=f"{prefijo}_a"), c2.multiselect(f"Mes", meses, default=meses, key=f"{prefijo}_m")
        mask = df['Fecha'].dt.year.isin(f_ano) & df['Fecha'].dt.month.isin(f_mes)
        if 'N° Semana' in df.columns:
            semanas = sorted(df[mask]['N° Semana'].unique()) if not df[mask].empty else sorted(df['N° Semana'].unique())
            f_sem = c3.multiselect("N° Semana", semanas, key=f"{prefijo}_s")
            if f_sem: mask &= df['N° Semana'].isin(f_sem)
        if 'Tipo Día' in df.columns:
            jornadas = df[mask]['Tipo Día'].unique() if not df[mask].empty else df['Tipo Día'].unique()
            f_jor = st.multiselect("Jornada", jornadas, default=jornadas, key=f"{prefijo}_j")
            if f_jor: mask &= df['Tipo Día'].isin(f_jor)
        return df[mask]

    # --- 6. RENDERIZADO DE PESTAÑAS ---
    tabs = st.tabs(["📊 Resumen", "📑 Datos operacionales", "📑 Odómetro por Tren", "⚡ Energía SEAT", "📈 Medidas PRMTE", "💰 Facturación", "⚖️ Comparación Energía por hr.", "📈 Regresión Nocturna", "🚨 Datos Atípicos"])
    
    with tabs[0]: # Resumen
        if not df_ops.empty:
            st.write("#### Filtros Resumen")
            df_res_f = get_filtros(df_ops, "res")
            if not df_res_f.empty:
                st.write("#### 📈 Indicadores Globales (Período Filtrado)")
                to_val, tk_val = df_res_f["Odómetro [km]"].sum(), df_res_f["Tren-Km [km]"].sum()
                umr_val = (tk_val / to_val * 100) if to_val > 0 else 0
                c1, c2, c3 = st.columns(3)
                c1.metric("Odómetro Total", f"{to_val:,.1f} km"); c2.metric("Tren-Km Total", f"{tk_val:,.1f} km"); c3.metric("UMR Global", f"{umr_val:.2f} %")
                if "E_Total" in df_res_f.columns:
                    e_tot, e_tr, e_12 = df_res_f["E_Total"].sum(), df_res_f["E_Tr"].sum() if "E_Tr" in df_res_f.columns else 0, df_res_f["E_12"].sum() if "E_12" in df_res_f.columns else 0
                    c4, c5, c6 = st.columns(3)
                    c4.metric("Energía Total", f"{e_tot:,.0f} kWh"); c5.metric("Tracción", f"{e_tr:,.0f} kWh"); c6.metric("12 kV", f"{e_12:,.0f} kWh")
                st.divider()
                df_res_f['Tipo Día'] = pd.Categorical(df_res_f['Tipo Día'], categories=['L', 'S', 'D/F'], ordered=True)
                agg_cols = {"Odómetro [km]":"sum", "Tren-Km [km]":"sum", "UMR [%]":"mean"}
                fmt_dict = {"Odómetro [km]":"{:,.1f}", "Tren-Km [km]":"{:,.1f}", "UMR [%]":"{:.2f}%"}
                for col in ["E_Total", "E_Tr", "E_12"]:
                    if col in df_res_f.columns: agg_cols[col] = "sum"; fmt_dict[col] = "{:,.0f}"
                res = df_res_f.groupby("Tipo Día", observed=True).agg(agg_cols).reset_index()
                st.write("#### Tabla Resumen por Tipo de Día"); st.table(res.style.format(fmt_dict))

    with tabs[1]: # Operaciones
        if not df_ops.empty:
            st.write("#### Filtros Operacionales")
            df_ops_f = get_filtros(df_ops, "ops")
            fmt_dict = {"Odómetro [km]":"{:,.1f}", "Tren-Km [km]":"{:,.1f}", "UMR [%]":"{:.2f}%"}
            for col in ["E_Total", "E_Tr", "E_12"]:
                if col in df_ops_f.columns: fmt_dict[col] = "{:,.0f}"
            st.dataframe(df_ops_f.style.format(fmt_dict), use_container_width=True)

    with tabs[2]: # Trenes
        if not df_tr.empty or not df_tr_acum.empty:
            st.write("#### Filtros Trenes")
            df_tr_comb = pd.concat([df_tr, df_tr_acum])
            if not df_tr_comb.empty:
                c1, c2 = st.columns(2)
                meses_tr, trenes = sorted(df_tr_comb['Fecha'].dt.month.unique()), sorted(df_tr_comb['Tren'].unique())
                f_mes_tr, f_tren = c1.multiselect("Mes", meses_tr, default=meses_tr, key="tr_m"), c2.multiselect("Tren(es)", trenes, key="tr_t")
                if not df_tr.empty:
                    df_tr_f = df_tr[df_tr['Fecha'].dt.month.isin(f_mes_tr)]
                    if f_tren: df_tr_f = df_tr_f[df_tr_f['Tren'].isin(f_tren)]
                    st.write("### 🚗 Kilometraje Diario [km]"); st.dataframe(df_tr_f.pivot_table(index="Tren", columns=df_tr_f["Fecha"].dt.day, values="Valor", aggfunc='sum').fillna(0).style.format("{:,.1f}"), use_container_width=True)
                if not df_tr_acum.empty:
                    df_tra_f = df_tr_acum[df_tr_acum['Fecha'].dt.month.isin(f_mes_tr)]
                    if f_tren: df_tra_f = df_tra_f[df_tra_f['Tren'].isin(f_tren)]
                    st.divider(); st.write("### 📈 Lectura de Odómetro / Acumulado [km]"); st.dataframe(df_tra_f.pivot_table(index="Tren", columns=df_tra_f["Fecha"].dt.day, values="Valor", aggfunc='max').fillna(0).style.format("{:,.0f}"), use_container_width=True)

    with tabs[3]: # SEAT
        if not df_seat.empty:
            st.write("#### Filtros Energía")
            df_seat_f = get_filtros(df_seat, "seat")
            st.dataframe(df_seat_f.style.format({"Total [kWh]":"{:,.0f}", "Tracción [kWh]":"{:,.0f}", "12 KV [kWh]":"{:,.0f}", "% Tracción":"{:.2f}%", "% 12 KV":"{:.2f}%"}), use_container_width=True)

    with tabs[4]: # PRMTE
        if not df_p_d.empty:
            st.write("#### 📅 Resumen Diario PRMTE"); df_prm_f = get_filtros(df_p_d, "prm")
            fmt_dict = {"Energía PRMTE [kWh]":"{:,.1f}"}
            for col in ["E_Tr", "E_12"]:
                if col in df_prm_f.columns: fmt_dict[col] = "{:,.1f}"
            st.dataframe(df_prm_f.style.format(fmt_dict), use_container_width=True)
            st.write("#### 🕒 Detalle 15 Minutos"); st.dataframe(pd.DataFrame(all_prmte_15).style.format({"Energía PRMTE [kWh]":"{:,.2f}"}), use_container_width=True)

    with tabs[5]: # FACTURA
        if not df_f_d.empty:
            st.write("#### 📅 Resumen Diario Facturación"); df_fact_f = get_filtros(df_f_d, "fact")
            fmt_dict = {"Consumo Horario [kWh]":"{:,.1f}"}
            for col in ["E_Tr", "E_12"]:
                if col in df_fact_f.columns: fmt_dict[col] = "{:,.1f}"
            st.dataframe(df_fact_f.style.format(fmt_dict), use_container_width=True)
            st.write("#### 🕒 Detalle Horario"); st.dataframe(pd.DataFrame(all_fact_h).style.format({"Consumo Horario [kWh]":"{:,.2f}"}), use_container_width=True)

    with tabs[6]: # Comparación por hr.
        if all_comp_full:
            df_comp = pd.DataFrame(all_comp_full).groupby(['Fecha', 'Hora', 'Fuente'])['Consumo Horario [kWh]'].sum().reset_index()
            fechas_con_factura = df_comp[df_comp['Fuente'] == 'Factura']['Fecha'].unique()
            df_comp_final = df_comp[~((df_comp['Fuente'] == 'PRMTE') & (df_comp['Fecha'].isin(fechas_con_factura)))].copy()
            df_comp_final['Fecha_str'] = df_comp_final['Fecha'].dt.strftime('%d/%m/%y')
            fechas_disp = sorted(df_comp_final['Fecha'].dt.date.unique())
            f_comp_fechas = st.multiselect("Selecciona fechas (xx/xx/xx):", fechas_disp, default=fechas_disp[:min(5, len(fechas_disp))], key="comp_f")
            if f_comp_fechas:
                df_comp_f = df_comp_final[df_comp_final['Fecha'].dt.date.isin(f_comp_fechas)]
                pivot_compare = df_comp_f.pivot_table(index="Hora", columns="Fecha_str", values="Consumo Horario [kWh]", aggfunc='sum').fillna(0).reindex(range(24)).fillna(0)
                cols_ordenadas = sorted(pivot_compare.columns, key=lambda x: datetime.strptime(x, '%d/%m/%y'))
                st.dataframe(pivot_compare[cols_ordenadas].style.format("{:,.1f}"), use_container_width=True)
                st.line_chart(pivot_compare[cols_ordenadas])
            st.divider(); st.write("#### 📊 Análisis Estadístico: Mediana 2025 vs 2026")
            df_comp_final['Año'], df_comp_final['Tipo Día'] = df_comp_final['Fecha'].dt.year, df_comp_final['Fecha'].apply(get_tipo_dia)
            df_stats = df_comp_final[df_comp_final['Año'].isin([2025, 2026])].copy()
            if not df_stats.empty:
                df_stats['Tipo Día'] = pd.Categorical(df_stats['Tipo Día'], categories=['L', 'S', 'D/F'], ordered=True)
                pivot_jor = df_stats.pivot_table(index="Hora", columns=["Año", "Tipo Día"], values="Consumo Horario [kWh]", aggfunc='median', observed=False).fillna(0)
                pivot_tot = df_stats.pivot_table(index="Hora", columns=["Año"], values="Consumo Horario [kWh]", aggfunc='median').fillna(0)
                frames = []
                for anio in sorted(df_stats['Año'].unique()):
                    temp = pd.DataFrame(index=range(24))
                    temp[f"{anio} - Mediana Total"] = pivot_tot.get(anio, pd.Series(0, index=range(24)))
                    for jor in ['L', 'S', 'D/F']: temp[f"{anio} - Mediana {jor}"] = pivot_jor[(anio, jor)] if (anio, jor) in pivot_jor.columns else 0
                    frames.append(temp)
                st.dataframe(pd.concat(frames, axis=1).fillna(0).style.format("{:,.1f}"), use_container_width=True)

    # --- VARIABLES GLOBALES PARA OUTLIERS ---
    df_outliers_global = pd.DataFrame()
    df_normal_global = pd.DataFrame()

    with tabs[7]: # REGRESIÓN NOCTURNA
        st.write("#### 📈 Regresión Lineal del Consumo Basal (00:00 - 05:00 hrs)")
        st.info("💡 Análisis de tendencia limpio: El sistema excluye automáticamente los datos anómalos (outliers) utilizando el Método del Rango Intercuartílico (IQR).")
        
        if all_comp_full:
            df_reg = pd.DataFrame(all_comp_full)
            df_reg = df_reg.groupby(['Fecha', 'Hora', 'Fuente'])['Consumo Horario [kWh]'].sum().reset_index()
            fechas_f = df_reg[df_reg['Fuente'] == 'Factura']['Fecha'].unique()
            df_reg = df_reg[~((df_reg['Fuente'] == 'PRMTE') & (df_reg['Fecha'].isin(fechas_f)))].copy()
            
            df_reg = df_reg[df_reg['Hora'] <= 5]
            df_reg['Año'] = df_reg['Fecha'].dt.year
            df_reg['Tipo Día'] = df_reg['Fecha'].apply(get_tipo_dia)
            
            c1, c2, c3 = st.columns(3)
            f_reg_anio = c1.selectbox("Año", sorted(df_reg['Año'].unique()), index=len(df_reg['Año'].unique())-1)
            f_reg_jor = c2.selectbox("Tipo de Jornada", ['L', 'S', 'D/F'])
            f_reg_hora = c3.selectbox("Hora específica", range(6))
            
            df_plot = df_reg[(df_reg['Año'] == f_reg_anio) & (df_reg['Tipo Día'] == f_reg_jor) & (df_reg['Hora'] == f_reg_hora)].sort_values('Fecha')
            
            if len(df_plot) > 1:
                Q1 = df_plot['Consumo Horario [kWh]'].quantile(0.25)
                Q3 = df_plot['Consumo Horario [kWh]'].quantile(0.75)
                IQR = Q3 - Q1
                
                if IQR == 0 and Q1 == Q3:
                    df_normal = df_plot.copy()
                    df_outliers = pd.DataFrame()
                else:
                    limite_superior = Q3 + (1.5 * IQR)
                    limite_inferior = Q1 - (1.5 * IQR)
                    
                    df_normal = df_plot[(df_plot['Consumo Horario [kWh]'] >= limite_inferior) & (df_plot['Consumo Horario [kWh]'] <= limite_superior)].copy()
                    df_outliers = df_plot[(df_plot['Consumo Horario [kWh]'] < limite_inferior) | (df_plot['Consumo Horario [kWh]'] > limite_superior)].copy()
                
                df_outliers_global = df_outliers
                df_normal_global = df_normal
                
                if len(df_normal) > 1:
                    x = np.arange(len(df_normal))
                    y = df_normal['Consumo Horario [kWh]'].values
                    
                    m, n = np.polyfit(x, y, 1)
                    y_pred = m * x + n
                    r_squared = 1 - (np.sum((y - y_pred)**2) / np.sum((y - np.mean(y))**2))
                    
                    # --- NUEVO CÁLCULO: TOTAL DEL PERÍODO LIMPIO ---
                    total_consumo = np.sum(y)
                    
                    c1, c2 = st.columns([2, 1])
                    chart_data = pd.DataFrame({'Días': df_normal['Fecha'].dt.strftime('%d/%m'), 'Consumo Real (Limpio)': y, 'Tendencia': y_pred}).set_index('Días')
                    c1.line_chart(chart_data)
                    
                    with c2:
                        st.metric("Consumo Total Acumulado", f"{total_consumo:,.2f} kWh", help="Suma de la energía en esta hora para el período filtrado (sin outliers)")
                        st.metric("Pendiente (m)", f"{m:.4f}", help="Incremento/decremento de kWh por día")
                        st.metric("Consumo Inicial (n)", f"{n:.2f} kWh")
                        st.metric("Coeficiente R²", f"{r_squared:.4f}")
                        
                        if not df_outliers.empty:
                            st.error(f"⚠️ {len(df_outliers)} datos atípicos excluidos.")
                    
                    # --- NUEVO RELATO DINÁMICO ---
                    # Lógica para la tendencia
                    if m > 0.5:
                        tendencia_txt = "al alza (posible alerta de equipos quedando encendidos de forma progresiva)"
                        icono_tendencia = "📈"
                    elif m < -0.5:
                        tendencia_txt = "a la baja (indicador positivo de mejora en los apagados de equipos)"
                        icono_tendencia = "📉"
                    else:
                        tendencia_txt = "estable (fluctuaciones menores propias de la operación base)"
                        icono_tendencia = "➖"
                        
                    # Lógica para la confiabilidad (ajustada a la realidad de 12kV nocturno)
                    if r_squared < 0.3:
                        confianza_txt = "aleatorio y sin una tendencia cronológica fuerte. **Esto es completamente normal para un consumo basal de baja tensión/12kV**, ya que indica que la energía responde a eventos puntuales (clima, mantenimientos en maestranza) y no a una degradación sistemática de la red a través de los días."
                    elif r_squared < 0.7:
                        confianza_txt = "de variabilidad moderada, mostrando cierta correlación con el paso de los días."
                    else:
                        confianza_txt = "altamente predecible y fuertemente correlacionado con el paso del tiempo."

                    st.markdown(f"#### {icono_tendencia} Análisis Experto de Desempeño")
                    st.info(f"""
                    Durante el período analizado, la instalación partió con un consumo inactivo base estimado de **{n:.2f} kWh**. 
                    
                    Al observar la evolución a través de los días, la tendencia de este consumo se muestra **{tendencia_txt}**, variando a un ritmo de **{m:.4f} kWh** diarios.
                    
                    Desde la óptica estadística del Sistema de Gestión ($R^2 = {r_squared:.4f}$), este comportamiento es **{confianza_txt}**. En total, durante este bloque horario y período específico, se han consumido **{total_consumo:,.2f} kWh** (excluyendo datos anómalos).
                    """)
                    
                    st.write(f"**Ecuación Matemática:** $Consumo = {m:.4f} \cdot x + {n:.2f}$")
                    st.dataframe(df_normal[['Fecha', 'Consumo Horario [kWh]']].style.format({"Consumo Horario [kWh]":"{:,.2f}"}), use_container_width=True)
                else:
                    st.warning("Al excluir los datos atípicos, no quedaron suficientes registros para trazar la regresión.")
            else:
                st.warning("No hay suficientes datos para calcular la regresión en la selección actual.")
        else:
            st.warning("Sube archivos de energía para procesar la regresión.")

    with tabs[8]: # DATOS ATÍPICOS
        st.write("#### 🚨 Registro de Datos Atípicos (Outliers)")
        st.info("💡 Aquí se muestran los registros que fueron excluidos automáticamente de la **Regresión Nocturna** mediante el método matemático del Rango Intercuartílico (IQR).")
        
        if not df_outliers_global.empty:
            st.error(f"Se detectaron **{len(df_outliers_global)}** registros anómalos para el filtro seleccionado.")
            st.dataframe(
                df_outliers_global[['Fecha', 'Hora', 'Consumo Horario [kWh]', 'Fuente', 'Tipo Día']].style.format({
                    "Fecha": lambda x: x.strftime('%d/%m/%Y'),
                    "Consumo Horario [kWh]": "{:,.2f}"
                }), 
                use_container_width=True
            )
            st.write("#### 📊 Análisis Causa Raíz")
            st.write("Si estos valores son inusualmente altos, te sugiero revisar las bitácoras operacionales para identificar:")
            st.markdown("""
            * Trabajos de mantenimiento pesado en las maestranzas durante la madrugada.
            * Climatización o sistemas auxiliares de trenes que no fueron apagados según el protocolo.
            * Errores puntuales en la lectura de los medidores PRMTE.
            """)
        else:
            if not df_normal_global.empty:
                st.success("✅ Excelente. No se detectaron anomalías matemáticas extremas. El consumo nocturno está dentro de los márgenes del IQR.")
            else:
                st.write("Selecciona parámetros en la pestaña 'Regresión Nocturna' para evaluar anomalías.")

    st.sidebar.download_button("📥 Descargar Reporte", to_excel_consolidado(df_ops, df_tr, df_tr_acum, df_seat, df_p_d, pd.DataFrame(all_prmte_15), pd.DataFrame(all_fact_h), df_f_d), "Reporte_EFE_SGE.xlsx")
else:
    st.info("👋 Sube los archivos en el panel lateral para comenzar el análisis.")
