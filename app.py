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

# --- 1. CONFIGURACIÓN Y ESTILOS ---
st.set_page_config(page_title="Dashboard SGE - EFE Valparaíso", layout="wide", page_icon="🚆")
chile_holidays = holidays.Chile()

# Orden operativo oficial de EFE Valparaíso
ORDEN_JORNADA = ['L', 'S', 'D/F']

def aplicar_estilos():
    st.markdown("""
        <style>
        .stMetric { background-color: #ffffff; padding: 20px; border-radius: 10px; border-left: 5px solid #005195; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
        .main { background-color: #f4f7f9; }
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

# --- 2. FUNCIONES DE APOYO (FILTROS Y EXPORTACIÓN) ---

def get_filtros(df, prefijo):
    if df.empty: return df
    # Asegurar que el tipo de día sea categórico para mantener el orden L, S, D/F
    if 'Tipo Día' in df.columns:
        df['Tipo Día'] = pd.Categorical(df['Tipo Día'], categories=ORDEN_JORNADA, ordered=True)
    
    c1, c2, c3 = st.columns(3)
    f_ano = c1.multiselect("Año", sorted(df['Fecha'].dt.year.unique()), key=f"{prefijo}_a")
    f_mes = c2.multiselect("Mes", sorted(df['Fecha'].dt.month.unique()), key=f"{prefijo}_m")
    
    mask = pd.Series([True] * len(df))
    if f_ano: mask &= df['Fecha'].dt.year.isin(f_ano)
    if f_mes: mask &= df['Fecha'].dt.month.isin(f_mes)
    
    if 'Tipo Día' in df.columns:
        f_jor = st.multiselect("Jornada", ORDEN_JORNADA, default=ORDEN_JORNADA, key=f"{prefijo}_j")
        if f_jor: mask &= df['Tipo Día'].isin(f_jor)
    
    return df[mask]

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
                table.cell(r + 1, c).text_frame.paragraphs[0].font.size = Pt(8)
    binary_output = BytesIO(); prs.save(binary_output); return binary_output.getvalue()

# --- 3. MOTOR DE PROCESAMIENTO ---

def procesar_todo(todos, start_date, end_date):
    results = {"all_ops":[], "all_tr":[], "all_tr_acum":[], "all_seat":[], "all_prmte_15":[], "all_fact_h":[], "all_comp_full":[]}
    for f in todos:
        try:
            xl = pd.ExcelFile(f)
            for sn in xl.sheet_names:
                sn_up = sn.upper()
                # UMR / Resumen
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
                                results["all_ops"].append({"Fecha": r['_dt'].normalize(), "Tipo Día": get_tipo_dia(r['_dt']), "Odómetro [km]": parse_latam_number(r[idx_o]), "Tren-Km [km]": parse_latam_number(r[idx_t])})
                # Trenes
                if 'ODO' in sn_up and 'KIL' in sn_up:
                    df_tr_raw = pd.read_excel(f, sheet_name=sn, header=None)
                    for i in range(len(df_tr_raw)-2):
                        for j in range(1, len(df_tr_raw.columns)):
                            val_dt = pd.to_datetime(df_tr_raw.iloc[i, j], errors='coerce')
                            if pd.notna(val_dt) and start_date <= val_dt.date() <= end_date:
                                for k in range(i+3, min(i+40, len(df_tr_raw))):
                                    tr = str(df_tr_raw.iloc[k, 0]).strip().upper()
                                    if re.match(r'^(M|XM)', tr):
                                        val_km = parse_latam_number(df_tr_raw.iloc[k, j])
                                        d_pt = {"Tren": tr, "Fecha": val_dt.normalize(), "Valor": val_km}
                                        if any(k in sn_up for k in ['ACUM', 'LECTURA']): results["all_tr_acum"].append(d_pt)
                                        else: results["all_tr"].append(d_pt)
                # SEAT
                if 'SEAT' in sn_up and 'SER' in sn_up:
                    df_s = pd.read_excel(f, sheet_name=sn, header=None)
                    for i in range(len(df_s)):
                        fs = pd.to_datetime(df_s.iloc[i, 1], errors='coerce')
                        if pd.notna(fs) and start_date <= fs.date() <= end_date:
                            results["all_seat"].append({"Fecha": fs.normalize(), "Total [kWh]": parse_latam_number(df_s.iloc[i, 3]), "Tracción [kWh]": parse_latam_number(df_s.iloc[i, 5]), "12 KV [kWh]": parse_latam_number(df_s.iloc[i, 7])})
                # PRMTE / Factura
                if any(k in sn_up for k in ['PRMTE', 'FACTURA']):
                    df_f = pd.read_excel(f, sheet_name=sn, header=None)
                    h_idx = next((i for i in range(len(df_f)) if any(k in str(df_f.iloc[i]).upper() for k in ['AÑO', 'FECHAHORA'])), None)
                    if h_idx is not None:
                        df_d = pd.read_excel(f, sheet_name=sn, header=h_idx)
                        if 'AÑO' in df_d.columns: # Caso PRMTE
                            df_d['TS'] = pd.to_datetime(df_d[['AÑO','MES','DIA','HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'})) + pd.to_timedelta(df_d['INICIO INTERVALO'].astype(int), unit='m')
                            col_e = [c for c in df_d.columns if 'Energia_Activa' in str(c)][0]
                            for _, r in df_d.iterrows():
                                results["all_comp_full"].append({"Fecha": r['TS'].normalize(), "Hora": r['TS'].hour, "Consumo Horario [kWh]": parse_latam_number(r[col_e]), "Fuente": "PRMTE"})
                                if start_date <= r['TS'].date() <= end_date: results["all_prmte_15"].append({"Fecha": r['TS'].normalize(), "Energía PRMTE [kWh]": parse_latam_number(r[col_e])})
                        else: # Caso Factura
                            df_d.columns = ['TS', 'Val']
                            df_d['TS'] = pd.to_datetime(df_d['TS'], errors='coerce')
                            for _, r in df_d.dropna(subset=['TS']).iterrows():
                                results["all_comp_full"].append({"Fecha": r['TS'].normalize(), "Hora": r['TS'].hour, "Consumo Horario [kWh]": abs(parse_latam_number(r['Val'])), "Fuente": "Factura"})
                                if start_date <= r['TS'].date() <= end_date: results["all_fact_h"].append({"Fecha": r['TS'].normalize(), "Consumo Horario [kWh]": abs(parse_latam_number(r['Val']))})
        except: continue
    return results

# --- 4. RENDERIZADO DE PESTAÑAS ---

def render_resumen(df_ops):
    st.header("📊 Resumen Operacional")
    if not df_ops.empty:
        df_res_f = get_filtros(df_ops, "res")
        if not df_res_f.empty:
            to_val, tk_val = df_res_f["Odómetro [km]"].sum(), df_res_f["Tren-Km [km]"].sum()
            c1, c2, c3 = st.columns(3)
            c1.metric("Odómetro Total", f"{to_val:,.1f} km")
            c2.metric("Tren-Km Total", f"{tk_val:,.1f} km")
            c3.metric("UMR Global", f"{(tk_val/to_val*100) if to_val>0 else 0:.2f} %")
            
            e_total = df_res_f["E_Total"].sum() if "E_Total" in df_res_f.columns else 0
            st.metric("Energía Total (Dato Triangulado)", f"{e_total:,.0f} kWh")
            
            st.write("#### Desempeño por Jornada")
            # Forzar orden L, S, D/F en la tabla de resumen
            res_j = df_res_f.groupby("Tipo Día", observed=True).agg({"Odómetro [km]":"sum", "Tren-Km [km]":"sum"}).reindex(ORDEN_JORNADA).dropna(how='all').reset_index()
            st.table(res_j.style.format({"Odómetro [km]":"{:,.1f}", "Tren-Km [km]":"{:,.1f}"}))

def render_comparacion_horaria(all_comp_full):
    st.header("⚖️ Comparación de Energía Horaria")
    if all_comp_full:
        df_c = pd.DataFrame(all_comp_full).groupby(['Fecha','Hora','Fuente'])['Consumo Horario [kWh]'].sum().reset_index()
        fechas_f = df_c[df_c['Fuente']=='Factura']['Fecha'].unique()
        df_cf = df_c[~((df_c['Fuente']=='PRMTE') & (df_c['Fecha'].isin(fechas_f)))].copy()
        df_cf['Año'] = df_cf['Fecha'].dt.year.astype(str)
        df_cf['Tipo Día'] = df_cf['Fecha'].apply(get_tipo_dia)
        
        pivot = df_cf.pivot_table(index="Hora", columns=["Año", "Tipo Día"], values="Consumo Horario [kWh]", aggfunc='median').fillna(0)
        
        # Agregar TOTAL ANUAL por año
        for anio in sorted(df_cf['Año'].unique()):
            pivot[(anio, "Total Anual")] = df_cf[df_cf['Año'] == anio].groupby("Hora")["Consumo Horario [kWh]"].median()
        
        # FORZAR ORDEN: Laboral -> Sábado -> D/F -> Total Anual
        new_cols = []
        for anio in sorted(df_cf['Año'].unique()):
            for jor in ORDEN_JORNADA + ["Total Anual"]:
                if (anio, jor) in pivot.columns: new_cols.append((anio, jor))
        
        pivot = pivot.reindex(columns=new_cols)
        st.dataframe(pivot.style.format("{:,.1f}"), use_container_width=True)

# --- 5. MAIN ---

def main():
    aplicar_estilos()
    with st.sidebar:
        st.header("📅 Filtro Global")
        today = date.today()
        date_range = st.date_input("Periodo", value=(today.replace(day=1), today))
        sd, ed = (date_range[0], date_range[1]) if len(date_range)==2 else (date_range[0], date_range[0])
        st.header("📂 Carga de Archivos")
        f_umr = st.file_uploader("1. UMR / Odómetros", type=["xlsx"], accept_multiple_files=True)
        f_seat = st.file_uploader("2. Energía SEAT", type=["xlsx"], accept_multiple_files=True)
        f_bill = st.file_uploader("3. Facturación y PRMTE", type=["xlsx"], accept_multiple_files=True)

    r = procesar_todo((f_umr or []) + (f_seat or []) + (f_bill or []), sd, ed)
    
    # Construcción de DFs
    df_ops = pd.DataFrame(r["all_ops"])
    df_tr = pd.DataFrame(r["all_tr"])
    df_tr_a = pd.DataFrame(r["all_tr_acum"])
    df_seat = pd.DataFrame(r["all_seat"])
    df_prmte = pd.DataFrame(r["all_prmte_15"])
    df_fact = pd.DataFrame(r["all_fact_h"])
    
    # Motor de Triangulación de Energía para el Resumen
    df_en_master = pd.DataFrame()
    if not df_seat.empty:
        df_en_master = df_seat.rename(columns={"Total [kWh]":"E_Total"})[["Fecha", "E_Total"]]
    
    # Si hay PRMTE o Factura, ellos mandan sobre el SEAT en el Resumen
    if not df_prmte.empty:
        p_res = df_prmte.groupby("Fecha")["Energía PRMTE [kWh]"].sum().reset_index().rename(columns={"Energía PRMTE [kWh]":"E_Total"})
        df_en_master = pd.concat([df_en_master, p_res]).drop_duplicates(subset="Fecha", keep="last")
    if not df_fact.empty:
        f_res = df_fact.groupby("Fecha")["Consumo Horario [kWh]"].sum().reset_index().rename(columns={"Consumo Horario [kWh]":"E_Total"})
        df_en_master = pd.concat([df_en_master, f_res]).drop_duplicates(subset="Fecha", keep="last")
    
    if not df_ops.empty and not df_en_master.empty:
        df_ops = pd.merge(df_ops, df_en_master, on="Fecha", how="left")

    # TABS
    tabs = st.tabs(["📊 Resumen", "📑 Operaciones", "📑 Trenes", "⚡ Energía cruda", "⚖️ Comparación Energía hr", "📈 Regresión Nocturna", "🚨 Atípicos"])
    
    with tabs[0]: render_resumen(df_ops)
    with tabs[1]: st.dataframe(get_filtros(df_ops, "ops"))
    with tabs[2]:
        if not df_tr.empty:
            st.subheader("🚗 Kilometraje Diario")
            st.dataframe(df_tr.pivot_table(index="Tren", columns=df_tr["Fecha"].dt.day, values="Valor", aggfunc='sum'))
        if not df_tr_a.empty:
            st.subheader("📈 Odómetro Acumulado")
            st.dataframe(df_tr_a.pivot_table(index="Tren", columns=df_tr_a["Fecha"].dt.day, values="Valor", aggfunc='max'))
    
    with tabs[3]: # RESTAURADAS PESTAÑAS DE ENERGÍA CRUDA
        sub_e = st.tabs(["⚡ SEAT", "📈 PRMTE", "💰 Facturación"])
        with sub_e[0]: st.dataframe(df_seat)
        with sub_e[1]: st.dataframe(df_prmte)
        with sub_e[2]: st.dataframe(df_fact)
        
    with tabs[4]: render_comparacion_horaria(r["all_comp_full"])
    
    with tabs[5]: # Regresión Nocturna
        if r["all_comp_full"]:
            df_reg = pd.DataFrame(r["all_comp_full"]).groupby(['Fecha','Hora'])['Consumo Horario [kWh]'].sum().reset_index()
            df_reg = df_reg[df_reg['Hora'] <= 5]
            df_reg['Año'] = df_reg['Fecha'].dt.year
            df_reg['Tipo Día'] = df_reg['Fecha'].apply(get_tipo_dia)
            f_ra = st.selectbox("Año", sorted(df_reg['Año'].unique()), key="ra_reg")
            f_rh = st.selectbox("Hora", range(6), key="rh_reg")
            df_pl = df_reg[(df_reg['Año']==f_ra) & (df_reg['Hora']==f_rh)].sort_values("Fecha")
            if len(df_pl) > 2:
                x = np.arange(len(df_pl))
                y = df_pl['Consumo Horario [kWh]'].values
                m, n = np.polyfit(x, y, 1)
                st.line_chart(pd.DataFrame({"Real":y, "Tendencia":m*x+n}, index=df_pl['Fecha'].dt.strftime('%d/%m')))
                st.metric("Intercepto (n)", f"{n:.2f} kWh")

if __name__ == "__main__":
    main()
