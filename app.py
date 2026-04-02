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

# =================================================================
# 1. CONFIGURACIÓN, ESTILOS Y UTILIDADES
# =================================================================
st.set_page_config(page_title="Dashboard SGE - EFE Valparaíso", layout="wide", page_icon="🚆")
chile_holidays = holidays.Chile()

def aplicar_estilos():
    st.markdown("""
        <style>
        .stMetric { background-color: #ffffff; padding: 20px; border-radius: 10px; border-left: 5px solid #005195; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
        .main { background-color: #f4f7f9; }
        .stTabs [data-baseweb="tab-list"] { gap: 10px; }
        .stTabs [data-baseweb="tab"] { background-color: #f1f3f5; border-radius: 5px; }
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

# =================================================================
# 2. MÓDULOS DE EXPORTACIÓN (PPTX / EXCEL)
# =================================================================

def gen_pptx(title_text, df=None, metrics=None):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = f"EFE Valparaíso: {title_text}"
    y_pos = Inches(1.5)
    
    if metrics:
        tx = slide.shapes.add_textbox(Inches(0.5), y_pos, Inches(9), Inches(1)).text_frame
        for k, v in metrics.items():
            p = tx.add_paragraph()
            p.text = f"• {k}: {v}"; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = RGBColor(0, 81, 149)
        y_pos += Inches(1.2)

    if df is not None and not df.empty:
        df_d = df.head(10).reset_index(drop=True)
        rows, cols = df_d.shape
        table = slide.shapes.add_table(rows + 1, cols, Inches(0.5), y_pos, Inches(9), Inches(3)).table
        for c, col in enumerate(df_d.columns):
            cell = table.cell(0, c); cell.text = str(col); cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(0, 81, 149)
            cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
        for r in range(rows):
            for c in range(cols):
                v = df_d.iloc[r, c]
                table.cell(r+1, c).text = str(v) if not isinstance(v, float) else f"{v:,.1f}"
    
    buf = BytesIO(); prs.save(buf); return buf.getvalue()

# =================================================================
# 3. MOTOR DE PROCESAMIENTO (DATA ENGINE)
# =================================================================

def engine_procesamiento(f_umr, f_seat, f_bill, start_date, end_date):
    results = {
        "all_ops": [], "all_tr": [], "all_tr_acum": [], 
        "all_seat": [], "all_prmte_15": [], "all_fact_h": [], "all_comp_full": []
    }
    archivos = (f_umr or []) + (f_seat or []) + (f_bill or [])
    
    for f in archivos:
        try:
            xl = pd.ExcelFile(f)
            for sn in xl.sheet_names:
                sn_up = sn.upper()
                # Lógica UMR / Operaciones
                if any(k in sn_up for k in ['UMR', 'RESUMEN']):
                    df = pd.read_excel(f, sheet_name=sn, header=None)
                    h = next((i for i in range(min(50, len(df))) if any(k in str(df.iloc[i]).upper() for k in ['ODO', 'FECHA'])), None)
                    if h is not None:
                        df_p = pd.read_excel(f, sheet_name=sn, header=h)
                        df_p.columns = [re.sub(r'[^A-Z]', '', str(c).upper().replace('Ó','O')) for c in df_p.columns]
                        if 'FECHA' in df_p.columns:
                            df_p['_dt'] = pd.to_datetime(df_p['FECHA'], errors='coerce')
                            mask = (df_p['_dt'].dt.date >= start_date) & (df_p['_dt'].dt.date <= end_date)
                            for _, r in df_p[mask].dropna(subset=['_dt']).iterrows():
                                results["all_ops"].append({
                                    "Fecha": r['_dt'].normalize(), "Tipo Día": get_tipo_dia(r['_dt']),
                                    "Odómetro [km]": parse_latam_number(r.get('ODO', 0)),
                                    "Tren-Km [km]": parse_latam_number(r.get('TRENKM', 0))
                                })
                # Lógica Trenes Individuales
                if 'ODO' in sn_up and 'KIL' in sn_up:
                    df_raw = pd.read_excel(f, sheet_name=sn, header=None)
                    for i in range(len(df_raw)-2):
                        for j in range(1, len(df_raw.columns)):
                            dt = pd.to_datetime(df_raw.iloc[i, j], errors='coerce')
                            if pd.notna(dt) and start_date <= dt.date() <= end_date:
                                for k in range(i+3, min(i+40, len(df_raw))):
                                    tr = str(df_raw.iloc[k, 0]).strip().upper()
                                    if re.match(r'^(M|XM)', tr):
                                        val = parse_latam_number(df_raw.iloc[k, j])
                                        item = {"Tren": tr, "Fecha": dt.normalize(), "Valor": val}
                                        if 'ACUM' in sn_up: results["all_tr_acum"].append(item)
                                        else: results["all_tr"].append(item)
                # Lógica Energía SEAT
                if 'SEAT' in sn_up:
                    df_s = pd.read_excel(f, sheet_name=sn, header=None)
                    for i in range(len(df_s)):
                        fs = pd.to_datetime(df_s.iloc[i, 1], errors='coerce')
                        if pd.notna(fs) and start_date <= fs.date() <= end_date:
                            results["all_seat"].append({
                                "Fecha": fs.normalize(), "Total [kWh]": parse_latam_number(df_s.iloc[i, 3]),
                                "Tracción [kWh]": parse_latam_number(df_s.iloc[i, 5]), "12 KV [kWh]": parse_latam_number(df_s.iloc[i, 7])
                            })
        except Exception as e: st.sidebar.error(f"Error procesando {f.name}: {e}")
    
    return results

# =================================================================
# 4. COMPONENTES DE INTERFAZ (RENDERERS)
# =================================================================

def render_resumen(df_ops):
    st.header("📊 Resumen Ejecutivo")
    if not df_ops.empty:
        c1, c2, c3 = st.columns(3)
        to, tk = df_ops["Odómetro [km]"].sum(), df_ops["Tren-Km [km]"].sum()
        c1.metric("Odómetro Total", f"{to:,.1f} km")
        c2.metric("Tren-Km Total", f"{tk:,.1f} km")
        c3.metric("UMR Global", f"{(tk/to*100) if to>0 else 0:.2f}%")
        
        st.subheader("Consumos por Jornada")
        res_j = df_ops.groupby("Tipo Día").agg({"Odómetro [km]":"sum", "Tren-Km [km]":"sum"}).reset_index()
        st.table(res_j.style.format({"Odómetro [km]":"{:,.1f}", "Tren-Km [km]":"{:,.1f}"}))

def render_trenes(df_tr, df_tr_acum):
    st.header("🚆 Detalle de Flota")
    if not df_tr.empty:
        st.subheader("Kilometraje Diario")
        piv = df_tr.pivot_table(index="Tren", columns=df_tr["Fecha"].dt.day, values="Valor", aggfunc='sum').fillna(0)
        st.dataframe(piv.style.format("{:,.1f}"))
    if not df_tr_acum.empty:
        st.subheader("Odómetro Acumulado")
        piv_a = df_tr_acum.pivot_table(index="Tren", columns=df_tr_acum["Fecha"].dt.day, values="Valor", aggfunc='max').fillna(0)
        st.dataframe(piv_a.style.format("{:,.0f}"))

def render_regresion(all_comp_full):
    st.header("📈 Regresión Nocturna (ISO 50001)")
    # Aquí pegamos toda tu lógica de outliers y polyfit
    if all_comp_full:
        df_reg = pd.DataFrame(all_comp_full).groupby(['Fecha','Hora'])['Consumo Horario [kWh]'].sum().reset_index()
        df_reg = df_reg[df_reg['Hora'] <= 5]
        
        h_sel = st.selectbox("Hora de Análisis", range(6))
        df_h = df_reg[df_reg['Hora'] == h_sel].sort_values('Fecha')
        
        if len(df_h) > 5:
            # Detección Outliers
            q1, q3 = df_h['Consumo Horario [kWh]'].quantile([0.25, 0.75])
            iqr = q3 - q1
            df_limpio = df_h[(df_h['Consumo Horario [kWh]'] >= q1 - 1.5*iqr) & (df_h['Consumo Horario [kWh]'] <= q3 + 1.5*iqr)]
            
            x = np.arange(len(df_limpio))
            y = df_limpio['Consumo Horario [kWh]'].values
            m, n = np.polyfit(x, y, 1)
            
            st.line_chart(pd.DataFrame({'Real': y, 'Tendencia': m*x + n}))
            st.success(f"Ecuación: y = {m:.4f}x + {n:.2f} | R²: {np.corrcoef(x, y)[0,1]**2:.4f}")
        else:
            st.info("Insuficientes datos para regresión.")

# =================================================================
# 5. ORQUESTADOR PRINCIPAL
# =================================================================

def main():
    aplicar_estilos()
    
    with st.sidebar:
        st.title("⚙️ Configuración")
        rango = st.date_input("Periodo", [date.today().replace(day=1), date.today()])
        st.divider()
        f_umr = st.file_uploader("UMR/Odómetros", accept_multiple_files=True)
        f_seat = st.file_uploader("Energía SEAT", accept_multiple_files=True)
        f_bill = st.file_uploader("Facturas/PRMTE", accept_multiple_files=True)

    if f_umr or f_seat or f_bill:
        sd, ed = (rango[0], rango[1]) if len(rango)==2 else (rango[0], rango[0])
        
        # 1. Procesar Datos
        res = engine_procesamiento(f_umr, f_seat, f_bill, sd, ed)
        
        # 2. Crear DataFrames
        df_ops = pd.DataFrame(res["all_ops"])
        df_tr = pd.DataFrame(res["all_tr"])
        df_tr_a = pd.DataFrame(res["all_tr_acum"])
        
        # 3. Dibujar Tabs
        t_res, t_tr, t_en, t_reg, t_at = st.tabs(["📊 Resumen", "🚆 Trenes", "⚡ Energía", "📈 Regresión", "🚨 Atípicos"])
        
        with t_res: render_resumen(df_ops)
        with t_tr: render_trenes(df_tr, df_tr_a)
        with t_reg: render_regresion(res["all_comp_full"])
        with t_at:
            st.header("🚨 Datos Atípicos")
            if 'outliers' in st.session_state: st.dataframe(st.session_state.outliers)
            else: st.success("Sin anomalías detectadas.")
            
    else:
        st.info("👋 Bienvenida/o. Por favor, carga los archivos Excel en el panel izquierdo para iniciar el análisis del SGE.")

if __name__ == "__main__":
    main()
