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

# --- 1. CONFIGURACIÓN Y UTILIDADES ---
st.set_page_config(page_title="Gestión de Energía - Dashboard SGE", layout="wide", page_icon="🚆")
chile_holidays = holidays.Chile()

def aplicar_estilos():
    st.markdown("""
        <style>
        .stMetric { background-color: #ffffff; padding: 20px; border-radius: 10px; border-left: 5px solid #005195; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
        </style>
        """, unsafe_allow_html=True)

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

# --- 2. FUNCIONES DE EXPORTACIÓN ---

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
        df_display = df.head(12).reset_index(drop=True)
        rows, cols = df_display.shape
        table = slide.shapes.add_table(rows + 1, cols, Inches(0.5), y_cursor, Inches(9), Inches(3)).table
        for c, col_name in enumerate(df_display.columns):
            cell = table.cell(0, c); cell.text = str(col_name); cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(0, 81, 149)
            p = cell.text_frame.paragraphs[0]; p.font.color.rgb = RGBColor(255, 255, 255); p.font.size = Pt(10)
        for r in range(rows):
            for c in range(cols):
                val = df_display.iloc[r, c]
                table.cell(r + 1, c).text = str(val) if not isinstance(val, float) else f"{val:,.1f}"
    binary_output = BytesIO(); prs.save(binary_output); return binary_output.getvalue()

def to_excel_consolidado(df_ops, df_tr, df_tr_acum, df_seat, df_p_d, df_p_15, df_fact_h, df_fact_d):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dfs = {'Operaciones': df_ops, 'Kms_Diarios_Tren': df_tr, 'Odometros_Acum_Tren': df_tr_acum,
               'SEAT': df_seat, 'PRMTE_D': df_p_d, 'PRMTE_15': df_p_15, 'Fact_H': df_fact_h, 'Fact_D': df_fact_d}
        for name, df in dfs.items():
            if not df.empty: df.to_excel(writer, index=False, sheet_name=name)
    return output.getvalue()

# --- 3. MOTOR DE DATOS (RESTAURADO 1:1) ---

def procesar_todo(todos, start_date, end_date):
    results = {"all_ops":[], "all_tr":[], "all_tr_acum":[], "all_seat":[], "all_prmte_15":[], "all_fact_h":[], "all_comp_full":[]}
    for f in todos:
        try:
            xl = pd.ExcelFile(f)
            for sn in xl.sheet_names:
                sn_up = sn.upper()
                # Lógica UMR / Resumen
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
                                results["all_ops"].append({"Fecha": r['_dt'].normalize(), "Tipo Día": get_tipo_dia(r['_dt']), "N° Semana": r['_dt'].isocalendar()[1], "Odómetro [km]": parse_latam_number(r[idx_o]), "Tren-Km [km]": parse_latam_number(r[idx_t]), "UMR [%]": (parse_latam_number(r[idx_t])/parse_latam_number(r[idx_o])*100 if parse_latam_number(r[idx_o])>0 else 0)})
                # Lógica ODO / Kilometraje
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
                                    if is_acum or idx > 0: results["all_tr_acum"].append(d_pt)
                                    else: results["all_tr"].append(d_pt)
                # Lógica SEAT
                if 'SEAT' in sn_up and 'SER' in sn_up:
                    df_s = pd.read_excel(f, sheet_name=sn, header=None)
                    for i in range(len(df_s)):
                        fs = pd.to_datetime(df_s.iloc[i, 1], errors='coerce')
                        if pd.notna(fs) and start_date <= fs.date() <= end_date:
                            tot, tra, k12 = parse_latam_number(df_s.iloc[i, 3]), parse_latam_number(df_s.iloc[i, 5]), parse_latam_number(df_s.iloc[i, 7])
                            results["all_seat"].append({"Fecha": fs.normalize(), "Total [kWh]": tot, "Tracción [kWh]": tra, "12 KV [kWh]": k12, "% Tracción": (tra/tot*100 if tot>0 else 0), "% 12 KV": (k12/tot*100 if tot>0 else 0)})
                # Lógica PRMTE
                if any(k in sn_up for k in ['PRMTE', 'MEDIDAS']):
                    df_prm = pd.read_excel(f, sheet_name=sn, header=None)
                    h_idx = next((i for i in range(len(df_prm)) if 'AÑO' in str(df_prm.iloc[i]).upper()), None)
                    if h_idx is not None:
                        df_pd = pd.read_excel(f, sheet_name=sn, header=h_idx)
                        df_pd['Timestamp'] = pd.to_datetime(df_pd[['AÑO', 'MES', 'DIA', 'HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'})) + pd.to_timedelta(df_pd['INICIO INTERVALO'].astype(int), unit='m')
                        cols_e = [c for c in df_pd.columns if 'Retiro_Energia_Activa (kWhD)' in str(c)]
                        for _, r in df_pd.iterrows():
                            ts, val_p = r['Timestamp'], sum([parse_latam_number(r[col]) for col in cols_e])
                            results["all_comp_full"].append({"Fecha": ts.normalize(), "Hora": ts.hour, "Consumo Horario [kWh]": val_p, "Fuente": "PRMTE"})
                            if start_date <= ts.date() <= end_date: results["all_prmte_15"].append({"Fecha y Hora": ts.strftime('%d/%m/%Y %H:%M'), "Fecha": ts.normalize(), "Energía PRMTE [kWh]": val_p})
                # Lógica Factura
                if any(k in sn_up for k in ['FACTURA', 'CONSUMO']):
                    df_f = pd.read_excel(f, sheet_name=sn); df_f.columns = ['FechaHora', 'Valor']
                    df_f['Timestamp'] = pd.to_datetime(df_f['FechaHora'], errors='coerce')
                    for _, r in df_f.dropna(subset=['Timestamp']).iterrows():
                        ts, val_f = r['Timestamp'], abs(parse_latam_number(r['Valor']))
                        results["all_comp_full"].append({"Fecha": ts.normalize(), "Hora": ts.hour, "Consumo Horario [kWh]": val_f, "Fuente": "Factura"})
                        if start_date <= ts.date() <= end_date: results["all_fact_h"].append({"Fecha y Hora": ts.strftime('%d/%m/%Y %H:%M'), "Fecha": ts.normalize(), "Consumo Horario [kWh]": val_f})
        except: continue
    return results

# --- 4. RENDERIZADO DE PESTAÑAS (RESTAURADO 1:1) ---

def render_resumen(df_ops):
    st.header("📊 Resumen Operacional")
    if not df_ops.empty:
        df_res_f = get_filtros(df_ops, "res")
        if not df_res_f.empty:
            to_val, tk_val = df_res_f["Odómetro [km]"].sum(), df_res_f["Tren-Km [km]"].sum()
            umr_val = (tk_val/to_val*100) if to_val>0 else 0
            c1, c2, c3 = st.columns(3)
            c1.metric("Odómetro Total", f"{to_val:,.1f} km"); c2.metric("Tren-Km Total", f"{tk_val:,.1f} km"); c3.metric("UMR Global", f"{umr_val:.2f} %")
            e_tot = df_res_f["E_Total"].sum() if "E_Total" in df_res_f.columns else 0
            st.metric("Energía Total", f"{e_tot:,.0f} kWh")
            st.write("#### Resumen por Jornada")
            res_j = df_res_f.groupby("Tipo Día", observed=True).agg({"Odómetro [km]":"sum", "Tren-Km [km]":"sum", "UMR [%]":"mean"}).reset_index()
            st.table(res_j.style.format({"Odómetro [km]":"{:,.1f}", "Tren-Km [km]":"{:,.1f}", "UMR [%]":"{:.2f}%"}))
            m_res = {"Odómetro": f"{to_val:,.1f} km", "Tren-Km": f"{tk_val:,.1f} km", "UMR": f"{umr_val:.2f}%", "Energía Total": f"{e_tot:,.0f} kWh"}
            st.download_button("📥 Descargar Resumen (PPTX)", to_pptx("Resumen Operacional", res_j, m_res), "EFE_Resumen.pptx")

def render_trenes(df_tr, df_tr_acum):
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
        if not df_tr_acum.empty:
            st.divider(); st.write("### 📈 Lectura de Odómetro / Acumulado [km]")
            df_tra_f = df_tr_acum[df_tr_acum['Fecha'].dt.month.isin(f_mes_tr)]
            if f_tren_tr: df_tra_f = df_tra_f[df_tra_f['Tren'].isin(f_tren_tr)]
            if not df_tra_f.empty:
                piv_acum = df_tra_f.pivot_table(index="Tren", columns=df_tra_f["Fecha"].dt.day, values="Valor", aggfunc='max').fillna(0)
                st.dataframe(piv_acum.style.format("{:,.0f}"), use_container_width=True)

def render_regresion(all_comp_full):
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
                m, n = np.polyfit(x, y, 1); r2 = 1 - (np.sum((y - (m*x+n))**2) / np.sum((y - np.mean(y))**2))
                st.line_chart(pd.DataFrame({'Real': y, 'Tendencia': m*x+n}, index=df_norm['Fecha'].dt.strftime('%d/%m')))
                st.markdown(f"**Ecuación:** $Consumo = {m:.4f}x + {n:.2f}$ | $R^2 = {r2:.4f}$")

def get_filtros(df, prefijo):
    if df.empty: return df
    c1, c2, c3 = st.columns(3)
    anios, meses = sorted(df['Fecha'].dt.year.unique()), sorted(df['Fecha'].dt.month.unique())
    f_ano = c1.multiselect(f"Año", anios, default=anios, key=f"{prefijo}_a")
    f_mes = c2.multiselect(f"Mes", meses, default=meses, key=f"{prefijo}_m")
    mask = df['Fecha'].dt.year.isin(f_ano) & df['Fecha'].dt.month.isin(f_mes)
    if 'N° Semana' in df.columns:
        f_sem = c3.multiselect("N° Semana", sorted(df[mask]['N° Semana'].unique()) if not df[mask].empty else [], key=f"{prefijo}_s")
        if f_sem: mask &= df['N° Semana'].isin(f_sem)
    if 'Tipo Día' in df.columns:
        f_jor = st.multiselect("Jornada", df[mask]['Tipo Día'].unique() if not df[mask].empty else [], default=df[mask]['Tipo Día'].unique() if not df[mask].empty else [], key=f"{prefijo}_j")
        if f_jor: mask &= df['Tipo Día'].isin(f_jor)
    return df[mask]

# --- 5. ORQUESTADOR (MAIN) ---

def main():
    aplicar_estilos()
    with st.sidebar:
        st.header("📅 Filtro Global")
        today = date.today()
        date_range = st.date_input("Selecciona el período", value=(today.replace(day=1), today))
        start_date, end_date = (date_range[0], date_range[1]) if isinstance(date_range, tuple) and len(date_range)==2 else (date_range, date_range)
        st.divider(); st.header("📂 Carga de Archivos")
        f_umr = st.file_uploader("1. UMR / Odómetros", type=["xlsx"], accept_multiple_files=True)
        f_seat_files = st.file_uploader("2. Energía SEAT", type=["xlsx"], accept_multiple_files=True)
        f_bill_files = st.file_uploader("3. Facturación y PRMTE", type=["xlsx"], accept_multiple_files=True)

    if 'outliers' not in st.session_state: st.session_state.outliers = pd.DataFrame()
    
    # Procesar datos
    r = procesar_todo((f_umr or []) + (f_seat_files or []) + (f_bill_files or []), start_date, end_date)
    
    # Construir DataFrames (Triangulación)
    df_ops, df_tr, df_tr_acum, df_seat = pd.DataFrame(r["all_ops"]), pd.DataFrame(r["all_tr"]), pd.DataFrame(r["all_tr_acum"]), pd.DataFrame(r["all_seat"])
    df_p_d = pd.DataFrame(r["all_prmte_15"]).groupby("Fecha")["Energía PRMTE [kWh]"].sum().reset_index() if r["all_prmte_15"] else pd.DataFrame()
    df_f_d = pd.DataFrame(r["all_fact_h"]).groupby("Fecha")["Consumo Horario [kWh]"].sum().reset_index() if r["all_fact_h"] else pd.DataFrame()
    
    # Lógica Master Energía
    df_energy_master = pd.DataFrame()
    if not df_seat.empty:
        df_energy_master = df_seat[["Fecha", "Total [kWh]", "Tracción [kWh]", "12 KV [kWh]"]].copy().rename(columns={"Total [kWh]":"E_Total", "Tracción [kWh]":"E_Tr", "12 KV [kWh]":"E_12"})
        df_energy_master["Fuente"] = "SEAT"
        if not df_p_d.empty:
            df_p_d = pd.merge(df_p_d, df_seat[["Fecha", "% Tracción", "% 12 KV"]], on="Fecha", how="left").fillna(0)
            df_p_d["E_Tr"], df_p_d["E_12"] = df_p_d["Energía PRMTE [kWh]"]*(df_p_d["% Tracción"]/100), df_p_d["Energía PRMTE [kWh]"]*(df_p_d["% 12 KV"]/100)
            df_energy_master = pd.concat([df_energy_master, df_p_d.rename(columns={"Energía PRMTE [kWh]":"E_Total"})[["Fecha","E_Total","E_Tr","E_12"]]]).drop_duplicates(subset=["Fecha"], keep="last")
        if not df_f_d.empty:
            df_f_d = pd.merge(df_f_d, df_seat[["Fecha", "% Tracción", "% 12 KV"]], on="Fecha", how="left").fillna(0)
            df_f_d["E_Tr"], df_f_d["E_12"] = df_f_d["Consumo Horario [kWh]"]*(df_f_d["% Tracción"]/100), df_f_d["Consumo Horario [kWh]"]*(df_f_d["% 12 KV"]/100)
            df_energy_master = pd.concat([df_energy_master, df_f_d.rename(columns={"Consumo Horario [kWh]":"E_Total"})[["Fecha","E_Total","E_Tr","E_12"]]]).drop_duplicates(subset=["Fecha"], keep="last")
    
    if not df_ops.empty and not df_energy_master.empty:
        df_ops = pd.merge(df_ops, df_energy_master, on="Fecha", how="left")

    # Tabs
    tabs = st.tabs(["📊 Resumen", "📑 Operaciones", "📑 Trenes", "⚡ Energía", "⚖️ Comparación Energía hr", "📈 Regresión Nocturna", "🚨 Datos Atípicos"])
    with tabs[0]: render_resumen(df_ops)
    with tabs[1]: st.dataframe(get_filtros(df_ops, "ops"))
    with tabs[2]: render_trenes(df_tr, df_tr_acum)
    with tabs[3]: 
        sub_e = st.tabs(["⚡ SEAT", "📈 PRMTE", "💰 Facturación"])
        with sub_e[0]: st.dataframe(df_seat)
        with sub_e[1]: st.dataframe(df_p_d)
        with sub_e[2]: st.dataframe(df_f_d)
    with tabs[5]: render_regresion(r["all_comp_full"])
    with tabs[6]: st.dataframe(st.session_state.outliers)

if __name__ == "__main__":
    main()
