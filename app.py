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

def aplicar_estilos():
    st.markdown("""
        <style>
        .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; border-left: 5px solid #005195; }
        .main { background-color: #f8f9fa; }
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

# --- 2. FILTROS Y EXPORTACIÓN ---

def get_filtros(df, prefijo):
    if df.empty: return df
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

def to_pptx(title_text, df=None, metrics=None):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = f"EFE: {title_text}"
    y = Inches(1.5)
    if metrics:
        tx = slide.shapes.add_textbox(Inches(0.5), y, Inches(9), Inches(1)).text_frame
        for k, v in metrics.items():
            p = tx.add_paragraph(); p.text = f"• {k}: {v}"; p.font.bold = True; p.font.color.rgb = RGBColor(0, 81, 149)
        y += Inches(1.2)
    if df is not None and not df.empty:
        df_e = df.reset_index() if hasattr(df, 'index') and (df.index.name or any(df.index.names)) else df
        rows, cols = df_e.shape
        table = slide.shapes.add_table(rows+1, cols, Inches(0.1), y, Inches(9.8), Inches(4)).table
        for c, col in enumerate(df_e.columns):
            cell = table.cell(0, c); cell.text = str(col); cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(0, 81, 149)
    buf = BytesIO(); prs.save(buf); return buf.getvalue()

# --- 3. MOTOR DE DATOS (AUDITADO) ---

def procesar_todo(archivos, sd, ed):
    res = {"ops":[], "tr":[], "tr_a":[], "seat":[], "prmte":[], "fact":[], "comp":[]}
    for f in archivos:
        try:
            xl = pd.ExcelFile(f)
            for sn in xl.sheet_names:
                su = sn.upper()
                # UMR
                if any(k in su for k in ['UMR', 'RESUMEN']):
                    df = pd.read_excel(f, sheet_name=sn, header=None)
                    h = next((i for i in range(min(50, len(df))) if any(k in str(df.iloc[i]).upper() for k in ['ODO', 'FECHA'])), None)
                    if h is not None:
                        df_p = pd.read_excel(f, sheet_name=sn, header=h)
                        df_p.columns = [re.sub(r'[^A-Z]', '', str(c).upper().replace('Ó','O')) for c in df_p.columns]
                        if 'FECHA' in df_p.columns:
                            df_p['_dt'] = pd.to_datetime(df_p['FECHA'], errors='coerce')
                            m = (df_p['_dt'].dt.date >= sd) & (df_p['_dt'].dt.date <= ed)
                            for _, r in df_p[m].dropna(subset=['_dt']).iterrows():
                                res["ops"].append({"Fecha": r['_dt'].normalize(), "Tipo Día": get_tipo_dia(r['_dt']), "Odómetro [km]": parse_latam_number(r.get('ODO',0)), "Tren-Km [km]": parse_latam_number(r.get('TRENKM',0))})
                # Trenes
                if 'ODO' in su and 'KIL' in su:
                    df_r = pd.read_excel(f, sheet_name=sn, header=None)
                    for i in range(len(df_r)-2):
                        for j in range(1, len(df_r.columns)):
                            dt = pd.to_datetime(df_r.iloc[i, j], errors='coerce')
                            if pd.notna(dt) and sd <= dt.date() <= ed:
                                for k in range(i+3, min(i+40, len(df_r))):
                                    tr = str(df_r.iloc[k, 0]).strip().upper()
                                    if re.match(r'^(M|XM)', tr):
                                        val = parse_latam_number(df_r.iloc[k, j])
                                        d = {"Tren": tr, "Fecha": dt.normalize(), "Valor": val}
                                        if 'ACUM' in su: res["tr_a"].append(d)
                                        else: res["tr"].append(d)
                # SEAT
                if 'SEAT' in su:
                    df_s = pd.read_excel(f, sheet_name=sn, header=None)
                    for i in range(len(df_s)):
                        fs = pd.to_datetime(df_s.iloc[i, 1], errors='coerce')
                        if pd.notna(fs) and sd <= fs.date() <= ed:
                            res["seat"].append({"Fecha": fs.normalize(), "Total [kWh]": parse_latam_number(df_s.iloc[i,3]), "12 KV [kWh]": parse_latam_number(df_s.iloc[i,7])})
                # PRMTE / Factura
                if any(k in su for k in ['PRMTE', 'FACTURA']):
                    df_f = pd.read_excel(f, sheet_name=sn, header=None)
                    hi = next((i for i in range(len(df_f)) if any(k in str(df_f.iloc[i]).upper() for k in ['AÑO', 'FECHAHORA'])), None)
                    if hi is not None:
                        df_d = pd.read_excel(f, sheet_name=sn, header=hi)
                        if 'AÑO' in df_d.columns: # PRMTE
                            df_d['TS'] = pd.to_datetime(df_d[['AÑO','MES','DIA','HORA']].astype(int).rename(columns={'AÑO':'year','MES':'month','DIA':'day','HORA':'hour'})) + pd.to_timedelta(df_d['INICIO INTERVALO'].astype(int), unit='m')
                            ce = [c for c in df_d.columns if 'Energia_Activa' in str(c)][0]
                            for _, r in df_d.iterrows():
                                res["comp"].append({"Fecha": r['TS'].normalize(), "Hora": r['TS'].hour, "Consumo": parse_latam_number(r[ce]), "Fuente": "PRMTE"})
                                if sd <= r['TS'].date() <= ed: res["prmte"].append({"Fecha": r['TS'].normalize(), "Valor": parse_latam_number(r[ce])})
                        else: # Factura
                            df_d.columns = ['TS', 'Val']; df_d['TS'] = pd.to_datetime(df_d['TS'], errors='coerce')
                            for _, r in df_d.dropna(subset=['TS']).iterrows():
                                res["comp"].append({"Fecha": r['TS'].normalize(), "Hora": r['TS'].hour, "Consumo": abs(parse_latam_number(r['Val'])), "Fuente": "Factura"})
                                if sd <= r['TS'].date() <= ed: res["fact"].append({"Fecha": r['TS'].normalize(), "Valor": abs(parse_latam_number(r['Val']))})
        except: continue
    return res

# --- 4. MAIN Y RENDERIZADO ---

def main():
    aplicar_estilos()
    with st.sidebar:
        st.header("📂 Carga EFE SGE")
        r_f = st.date_input("Periodo", [date.today().replace(day=1), date.today()])
        sd, ed = (r_f[0], r_f[1]) if len(r_f)==2 else (r_f[0], r_f[0])
        files = st.file_uploader("Subir Excels", accept_multiple_files=True)

    if files:
        r = procesar_todo(files, sd, ed)
        df_ops, df_tr, df_tr_a, df_seat = pd.DataFrame(r["ops"]), pd.DataFrame(r["tr"]), pd.DataFrame(r["tr_a"]), pd.DataFrame(r["seat"])
        df_prmte, df_fact = pd.DataFrame(r["prmte"]), pd.DataFrame(r["fact"])
        
        # Triangulación para Resumen
        df_m = pd.DataFrame()
        if not df_seat.empty: df_m = df_seat.rename(columns={"Total [kWh]":"E_Total"})[["Fecha", "E_Total"]]
        if not df_prmte.empty: 
            p_sum = df_prmte.groupby("Fecha")["Valor"].sum().reset_index().rename(columns={"Valor":"E_Total"})
            df_m = pd.concat([df_m, p_sum]).drop_duplicates("Fecha", keep="last")
        if not df_fact.empty:
            f_sum = df_fact.groupby("Fecha")["Valor"].sum().reset_index().rename(columns={"Valor":"E_Total"})
            df_m = pd.concat([df_m, f_sum]).drop_duplicates("Fecha", keep="last")
        if not df_ops.empty and not df_m.empty: df_ops = pd.merge(df_ops, df_m, on="Fecha", how="left")

        tabs = st.tabs(["📊 Resumen", "📑 Operaciones", "📑 Trenes", "⚡ Energía cruda", "⚖️ Comparación hr", "📈 Regresión", "🚨 Atípicos"])
        
        with tabs[0]: # RESUMEN
            st.header("📊 Resumen SGE")
            if not df_ops.empty:
                df_f = get_filtros(df_ops, "res")
                if not df_f.empty:
                    c1, c2, c3 = st.columns(3)
                    to, tk = df_f["Odómetro [km]"].sum(), df_f["Tren-Km [km]"].sum()
                    c1.metric("Odómetro", f"{to:,.1f}"); c2.metric("Tren-Km", f"{tk:,.1f}"); c3.metric("UMR", f"{(tk/to*100) if to>0 else 0:.2f}%")
                    st.write("#### Jornada (Orden EFE)")
                    res_j = df_f.groupby("Tipo Día", observed=True).agg({"Odómetro [km]":"sum", "Tren-Km [km]":"sum"}).reindex(ORDEN_JORNADA).dropna(how='all')
                    st.table(res_j.style.format("{:,.1f}"))

        with tabs[2]: # TRENES
            if not df_tr.empty:
                st.subheader("Kilometraje Diario")
                st.dataframe(df_tr.pivot_table(index="Tren", columns=df_tr["Fecha"].dt.day, values="Valor", aggfunc='sum'))
            if not df_tr_a.empty:
                st.subheader("Odómetro Acumulado")
                st.dataframe(df_tr_a.pivot_table(index="Tren", columns=df_tr_a["Fecha"].dt.day, values="Valor", aggfunc='max'))

        with tabs[3]: # ENERGÍA CRUDA (RESTAURADA)
            s1, s2, s3 = st.tabs(["⚡ SEAT", "📈 PRMTE", "💰 Factura"])
            s1.dataframe(df_seat); s2.dataframe(df_prmte); s3.dataframe(df_fact)

        with tabs[4]: # COMPARACIÓN (CON TOTAL ANUAL Y ORDEN L,S,D/F)
            if r["comp"]:
                st.header("⚖️ Comparativa Horaria")
                df_c = pd.DataFrame(r["comp"]).groupby(['Fecha','Hora','Fuente'])['Consumo'].sum().reset_index()
                f_f = df_c[df_c['Fuente']=='Factura']['Fecha'].unique()
                df_cf = df_c[~((df_c['Fuente']=='PRMTE') & (df_c['Fecha'].isin(f_f)))].copy()
                df_cf['Año'] = df_cf['Fecha'].dt.year.astype(str)
                df_cf['Tipo Día'] = df_cf['Fecha'].apply(get_tipo_dia)
                pivot = df_cf.pivot_table(index="Hora", columns=["Año", "Tipo Día"], values="Consumo", aggfunc='median').fillna(0)
                for a in sorted(df_cf['Año'].unique()):
                    pivot[(a, "Total Anual")] = df_cf[df_cf['Año'] == a].groupby("Hora")["Consumo"].median()
                nc = []
                for a in sorted(df_cf['Año'].unique()):
                    for j in ORDEN_JORNADA + ["Total Anual"]:
                        if (a, j) in pivot.columns: nc.append((a, j))
                st.dataframe(pivot.reindex(columns=nc).style.format("{:,.1f}"), use_container_width=True)

        with tabs[5]: # REGRESIÓN (00-05 AM)
            if r["comp"]:
                st.header("📈 Regresión 12kV (Sin Trenes)")
                df_reg = pd.DataFrame(r["comp"])
                df_reg = df_reg[df_reg['Hora'] <= 5]
                df_reg['Año'] = df_reg['Fecha'].dt.year
                ay = sorted(df_reg['Año'].unique())
                c1, c2 = st.columns(2)
                fa, fh = c1.selectbox("Año", ay), c2.selectbox("Hora", range(6))
                df_p = df_reg[(df_reg['Año']==fa) & (df_reg['Hora']==fh)].sort_values("Fecha")
                if len(df_p) > 2:
                    y = df_p['Consumo'].values; x = np.arange(len(y))
                    m, n = np.polyfit(x, y, 1)
                    st.line_chart(pd.DataFrame({"Real":y, "Tendencia":m*x+n}, index=df_p['Fecha'].dt.strftime('%d/%m')))
                    st.latex(f"y = {m:.4f}x + {n:.2f}")

if __name__ == "__main__":
    main()
