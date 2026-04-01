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
    # Selector de calendario tipo rango
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
all_comp_full = [] # Memoria paralela unificada para la pestaña de Comparación por hora

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
                        mask = (df_p['_dt'].dt.date >= start_date) & (df_p['_dt'].dt.date <= end_date)
                        for _, r in df_p[mask].dropna(subset=['_dt']).iterrows():
                            o, tk = parse_latam_number(r[idx_o]), parse_latam_number(r[idx_t])
                            if o > 0: all_ops.append({"Fecha": r['_dt'].normalize(), "Tipo Día": get_tipo_dia(r['_dt']), "N° Semana": r['_dt'].isocalendar()[1], "Odómetro [km]": o, "Tren-Km [km]": tk, "UMR [%]": (tk/o*100)})

            # --- B. ODÓMETRO POR TREN (TABLA DIARIA + TABLA ACUMULADA) ---
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

            # --- C. ENERGÍA SEAT ---
            if 'SEAT' in sn_up and 'SER' in sn_up:
                df_s = pd.read_excel(f, sheet_name=sn, header=None)
                for i in range(len(df_s)):
                    fs = pd.to_datetime(df_s.iloc[i, 1], errors='coerce')
                    if pd.notna(fs) and start_date <= fs.date() <= end_date:
                        tot, tra, k12 = parse_latam_number(df_s.iloc[i, 3]), parse_latam_number(df_s.iloc[i, 5]), parse_latam_number(df_s.iloc[i, 7])
                        all_seat.append({"Fecha": fs.normalize(), "Total [kWh]": tot, "Tracción [kWh]": tra, "12 KV [kWh]": k12, "% Tracción": (tra/tot*100 if tot>0 else 0), "% 12 KV": (k12/tot*100 if tot>0 else 0)})

            # --- D. PRMTE / FACTURA ---
            if any(k in sn_up for k in ['PRMTE', 'MEDIDAS']):
                df_prm = pd.read_excel(f, sheet_name=sn, header=None)
                h_idx = next((i for i in range(len(df_prm)) if 'AÑO' in str(df_prm.iloc[i]).upper()), None)
                if h_idx is not None:
                    df_prm_d = pd.read_excel(f, sheet_name=sn, header=h_idx)
                    df_prm_d['Timestamp'] = pd.to_datetime(df_prm_d[['AÑO', 'MES', 'DIA', 'HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'})) + pd.to_timedelta(df_prm_d['INICIO INTERVALO'].astype(int), unit='m')
                    
                    # Suma dinámica de columnas "Retiro_Energia_Activa"
                    cols_energia = [c for c in df_prm_d.columns if 'Retiro_Energia_Activa (kWhD)' in str(c)]
                    
                    # Iteramos sin filtro para cargar la memoria histórica
                    for _, r in df_prm_d.iterrows():
                        ts = r['Timestamp']
                        suma_prmte = sum([parse_latam_number(r[col]) for col in cols_energia])
                        
                        # 1. Guardar en memoria de comparación anual
                        all_comp_full.append({
                            "Fecha": ts.normalize(),
                            "Hora": ts.hour,
                            "Consumo Horario [kWh]": suma_prmte,
                            "Fuente": "PRMTE"
                        })
                        
                        # 2. Guardar en memoria de pestañas principales si entra en el calendario
                        if start_date <= ts.date() <= end_date:
                            all_prmte_15.append({"Fecha y Hora": ts.strftime('%d/%m/%Y %H:%M'), "Fecha": ts.normalize(), "Energía PRMTE [kWh]": suma_prmte})

            if any(k in sn_up for k in ['FACTURA', 'CONSUMO']):
                df_f = pd.read_excel(f, sheet_name=sn); df_f.columns = ['FechaHora', 'Valor']
                df_f['Timestamp'] = pd.to_datetime(df_f['FechaHora'], errors='coerce')
                
                # Iteramos sin filtro para cargar la memoria histórica
                for _, r in df_f.dropna(subset=['Timestamp']).iterrows():
                    ts = r['Timestamp']
                    val = abs(parse_latam_number(r['Valor']))
                    
                    # 1. Guardar en memoria de comparación anual
                    all_comp_full.append({
                        "Fecha": ts.normalize(),
                        "Hora": ts.hour,
                        "Consumo Horario [kWh]": val,
                        "Fuente": "Factura"
                    })
                    
                    # 2. Guardar en memoria de pestañas principales si entra en el calendario
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
    
    # Lógica de Jerarquía de Energía
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
        anios = sorted(df['Fecha'].dt.year.unique())
        meses = sorted(df['Fecha'].dt.month.unique())
        
        f_ano = c1.multiselect(f"Año", anios, default=anios, key=f"{prefijo}_a")
        f_mes = c2.multiselect(f"Mes", meses, default=meses, key=f"{prefijo}_m")
        
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
    tabs = st.tabs(["📊 Resumen", "📑 Datos operacionales", "📑 Odómetro por Tren", "⚡ Energía SEAT", "📈 Medidas PRMTE", "💰 Facturación", "⚖️ Comparación Energía por hr."])
    
    with tabs[0]: # Resumen
        if not df_ops.empty:
            st.write("#### Filtros Resumen")
            df_res_f = get_filtros(df_ops, "res")
            if not df_res_f.empty:
                st.write("#### 📈 Indicadores Globales (Período Filtrado)")
                to_val = df_res_f["Odómetro [km]"].sum()
                tk_val = df_res_f["Tren-Km [km]"].sum()
                umr_val = (tk_val / to_val * 100) if to_val > 0 else 0
                
                c1, c2, c3 = st.columns(3)
                c1.metric("Odómetro Total", f"{to_val:,.1f} km")
                c2.metric("Tren-Km Total", f"{tk_val:,.1f} km")
                c3.metric("UMR Global", f"{umr_val:.2f} %")
                
                if "E_Total" in df_res_f.columns:
                    e_tot = df_res_f["E_Total"].sum()
                    e_tr = df_res_f["E_Tr"].sum() if "E_Tr" in df_res_f.columns else 0
                    e_12 = df_res_f["E_12"].sum() if "E_12" in df_res_f.columns else 0
                    
                    c4, c5, c6 = st.columns(3)
                    c4.metric("Energía Total", f"{e_tot:,.0f} kWh")
                    c5.metric("Tracción", f"{e_tr:,.0f} kWh")
                    c6.metric("12 kV", f"{e_12:,.0f} kWh")
                
                st.divider()

                df_res_f['Tipo Día'] = pd.Categorical(df_res_f['Tipo Día'], categories=['L', 'S', 'D/F'], ordered=True)
                agg_cols = {"Odómetro [km]":"sum", "Tren-Km [km]":"sum", "UMR [%]":"mean"}
                fmt_dict = {"Odómetro [km]":"{:,.1f}", "Tren-Km [km]":"{:,.1f}", "UMR [%]":"{:.2f}%"}
                for col in ["E_Total", "E_Tr", "E_12"]:
                    if col in df_res_f.columns:
                        agg_cols[col] = "sum"
                        fmt_dict[col] = "{:,.0f}"
                        
                res = df_res_f.groupby("Tipo Día", observed=True).agg(agg_cols).reset_index()
                st.write("#### Tabla Resumen por Tipo de Día")
                st.table(res.style.format(fmt_dict))

    with tabs[1]: # Datos Operacionales
        if not df_ops.empty:
            st.write("#### Filtros Operacionales")
            df_ops_f = get_filtros(df_ops, "ops")
            fmt_dict = {"Odómetro [km]":"{:,.1f}", "Tren-Km [km]":"{:,.1f}", "UMR [%]":"{:.2f}%"}
            for col in ["E_Total", "E_Tr", "E_12"]:
                if col in df_ops_f.columns: fmt_dict[col] = "{:,.0f}"
            st.dataframe(df_ops_f.style.format(fmt_dict), use_container_width=True)

    with tabs[2]: # Odómetro por Tren
        if not df_tr.empty or not df_tr_acum.empty:
            st.write("#### Filtros Trenes")
            df_tr_comb = pd.concat([df_tr, df_tr_acum])
            if not df_tr_comb.empty:
                c1, c2 = st.columns(2)
                meses_tr = sorted(df_tr_comb['Fecha'].dt.month.unique())
                trenes = sorted(df_tr_comb['Tren'].unique())
                f_mes_tr = c1.multiselect("Mes", meses_tr, default=meses_tr, key="tr_m")
                f_tren = c2.multiselect("Tren(es)", trenes, key="tr_t")
                
                if not df_tr.empty:
                    df_tr_f = df_tr[df_tr['Fecha'].dt.month.isin(f_mes_tr)]
                    if f_tren: df_tr_f = df_tr_f[df_tr_f['Tren'].isin(f_tren)]
                    st.write("### 🚗 Kilometraje Diario [km]")
                    st.dataframe(df_tr_f.pivot_table(index="Tren", columns=df_tr_f["Fecha"].dt.day, values="Valor", aggfunc='sum').fillna(0).style.format("{:,.1f}"), use_container_width=True)
                
                if not df_tr_acum.empty:
                    df_tra_f = df_tr_acum[df_tr_acum['Fecha'].dt.month.isin(f_mes_tr)]
                    if f_tren: df_tra_f = df_tra_f[df_tra_f['Tren'].isin(f_tren)]
                    st.divider()
                    st.write("### 📈 Lectura de Odómetro / Acumulado [km]")
                    st.dataframe(df_tra_f.pivot_table(index="Tren", columns=df_tra_f["Fecha"].dt.day, values="Valor", aggfunc='max').fillna(0).style.format("{:,.0f}"), use_container_width=True)

    with tabs[3]: # SEAT
        if not df_seat.empty:
            st.write("#### Filtros Energía")
            df_seat_f = get_filtros(df_seat, "seat")
            st.dataframe(df_seat_f.style.format({"Total [kWh]":"{:,.0f}", "Tracción [kWh]":"{:,.0f}", "12 KV [kWh]":"{:,.0f}", "% Tracción":"{:.2f}%", "% 12 KV":"{:.2f}%"}), use_container_width=True)

    with tabs[4]: # PRMTE
        if not df_p_d.empty:
            st.write("#### 📅 Resumen Diario PRMTE")
            df_prm_f = get_filtros(df_p_d, "prm")
            
            fmt_dict = {"Energía PRMTE [kWh]":"{:,.1f}"}
            for col in ["E_Tr", "E_12"]:
                if col in df_prm_f.columns: fmt_dict[col] = "{:,.1f}"
                
            st.dataframe(df_prm_f.style.format(fmt_dict), use_container_width=True)
            st.write("#### 🕒 Detalle 15 Minutos")
            st.dataframe(pd.DataFrame(all_prmte_15).style.format({"Energía PRMTE [kWh]":"{:,.2f}"}), use_container_width=True)

    with tabs[5]: # FACTURA
        if not df_f_d.empty:
            st.write("#### 📅 Resumen Diario Facturación")
            df_fact_f = get_filtros(df_f_d, "fact")
            
            fmt_dict = {"Consumo Horario [kWh]":"{:,.1f}"}
            for col in ["E_Tr", "E_12"]:
                if col in df_fact_f.columns: fmt_dict[col] = "{:,.1f}"
                
            st.dataframe(df_fact_f.style.format(fmt_dict), use_container_width=True)
            st.write("#### 🕒 Detalle Horario")
            st.dataframe(pd.DataFrame(all_fact_h).style.format({"Consumo Horario [kWh]":"{:,.2f}"}), use_container_width=True)

    with tabs[6]: # COMPARACIÓN ENERGÍA POR HR. (NUEVA PESTAÑA)
        st.write("#### ⚖️ Matriz Comparativa por Fechas Específicas")
        st.info("💡 Esta pestaña jerarquiza automáticamente: Utiliza los datos de Factura si existen para la fecha seleccionada; si no, usa el consolidado por hora del PRMTE.")
        
        if all_comp_full:
            # 1. Creamos el DataFrame histórico general
            df_comp = pd.DataFrame(all_comp_full)
            
            # 2. Agrupamos por Fecha, Hora y Fuente para sumar bloques de 15 min (si es PRMTE) a 1 hora
            df_comp = df_comp.groupby(['Fecha', 'Hora', 'Fuente'])['Consumo Horario [kWh]'].sum().reset_index()
            
            # 3. Jerarquía (Factura manda sobre PRMTE por día)
            fechas_con_factura = df_comp[df_comp['Fuente'] == 'Factura']['Fecha'].unique()
            # Si para un día hay Factura, descartamos cualquier dato de PRMTE de ese mismo día
            mask_descarta_prmte = (df_comp['Fuente'] == 'PRMTE') & (df_comp['Fecha'].isin(fechas_con_factura))
            df_comp_final = df_comp[~mask_descarta_prmte].copy()
            
            # Formatear la fecha como solicitaste: xx/xx/xx
            df_comp_final['Fecha_str'] = df_comp_final['Fecha'].dt.strftime('%d/%m/%y')
            fechas_disp = sorted(df_comp_final['Fecha'].dt.date.unique())
            
            # Selector múltiple de fechas a comparar
            f_comp_fechas = st.multiselect(
                "Selecciona las fechas a comparar (xx/xx/xx):", 
                fechas_disp, 
                default=fechas_disp[:min(5, len(fechas_disp))], 
                key="comp_f"
            )
            
            if f_comp_fechas:
                df_comp_f = df_comp_final[df_comp_final['Fecha'].dt.date.isin(f_comp_fechas)]
                
                # 4. Tabla Pivote exacta de la imagen (Filas: Horas 0-23, Columnas: xx/xx/xx)
                pivot_compare = df_comp_f.pivot_table(
                    index="Hora", 
                    columns="Fecha_str", 
                    values="Consumo Horario [kWh]", 
                    aggfunc='sum'
                ).fillna(0)
                
                # Rellenar con ceros las horas faltantes para siempre mostrar de 0 a 23
                pivot_compare = pivot_compare.reindex(range(24)).fillna(0)
                
                # Ordenar columnas de fechas cronológicamente de izquierda a derecha
                cols_ordenadas = sorted(pivot_compare.columns, key=lambda x: datetime.strptime(x, '%d/%m/%y'))
                pivot_compare = pivot_compare[cols_ordenadas]
                
                st.write("#### 📋 Detalle de Consumo por Hora [kWh]")
                st.dataframe(pivot_compare.style.format("{:,.1f}"), use_container_width=True)
                
                st.write("#### 📉 Curva de Carga Horaria")
                st.line_chart(pivot_compare)
            else:
                st.warning("Selecciona al menos una fecha para generar la matriz comparativa.")
        else:
            st.warning("Sube archivos de Facturación o PRMTE para habilitar esta comparativa.")

    st.sidebar.download_button("📥 Descargar Reporte", to_excel_consolidado(df_ops, df_tr, df_tr_acum, df_seat, df_p_d, pd.DataFrame(all_prmte_15), pd.DataFrame(all_fact_h), df_f_d), "Reporte_EFE_SGE.xlsx")
else:
    st.info("Sube los archivos y asegúrate de que el filtro de calendario abarque las fechas de los datos.")
