import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime

# --- 1. CONFIGURACIÓN E INTERFAZ ---
st.set_page_config(page_title="SGE - EFE Valparaíso", layout="wide", page_icon="⚡")

# Estilos para mejorar la visualización de métricas y tablas
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    div[data-testid="metric-container"] {
        background-color: #ffffff;
        border: 1px solid #dee2e6;
        padding: 15px;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        background-color: #f1f3f5;
        border-radius: 5px 5px 0 0;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. CARGA Y PROCESAMIENTO (LOGICA DE NEGOCIO) ---
with st.sidebar:
    st.image("https://www.efe.cl/wp-content/uploads/2021/04/logo-efe-valparaiso.png", width=200) # Opcional si tienes el link
    st.header("📥 Carga de Datos")
    
    file_prmte = st.file_uploader("Perfil PRMTE (CSV)", type=["csv"])
    file_seat = st.file_uploader("Datos SEAT (Excel)", type=["xlsx"])
    file_km = st.file_uploader("Kilometraje UMR (Excel)", type=["xlsx"])
    
    st.divider()
    st.subheader("⚙️ Parámetros del Sistema")
    pax_xtrapolis = st.number_input("Pasajeros por Xtrapolis", value=406)
    factor_bt = st.slider("Deducción Baja Tensión (%)", 0.0, 10.0, 1.5) / 100

@st.cache_data
def procesar_todo(f_prmte, f_seat, f_km):
    data = {"prmte": None, "seat": None, "km": None, "comparador": None}
    
    # Procesar PRMTE
    if f_prmte:
        data["prmte"] = pd.read_csv(f_prmte, sep=';', decimal=',')
    
    # Procesar SEAT con Lógica de Consenso
    if f_seat:
        df_s = pd.read_excel(f_seat)
        # Lógica de Consenso: Si ID y Fecha coinciden pero el valor es distinto -> 0
        if 'ID' in df_s.columns and 'Valor' in df_s.columns:
            # Identificar duplicados inconsistentes
            duplicados = df_s.duplicated(subset=['ID', 'Fecha'], keep=False)
            df_s['Consenso'] = df_s['Valor']
            mask_inconsistente = duplicados & (df_s.groupby(['ID', 'Fecha'])['Valor'].transform('nunique') > 1)
            df_s.loc[mask_inconsistente, 'Consenso'] = 0
            data["seat"] = df_s

    # Procesar Kilometraje
    if f_km:
        data["km"] = pd.read_excel(f_km)
        
    return data

datos = procesar_todo(file_prmte, file_seat, file_km)

# --- 3. ESTRUCTURA DE 8 PESTAÑAS ---
st.title("⚡ Gestión Energética EFE Valparaíso")

t_kpi, t_iso, t_motrices, t_carga, t_malla, t_reporte, t_seat, t_comp = st.tabs([
    "📊 KPIs", "📜 ISO 50001", "🚆 Motrices", "🔌 Carga", 
    "🧬 Malla", "📅 Reporte Diario", "📉 SEAT Detalle", "⚖️ Comparador"
])

# --- PESTAÑA 1: KPIs ---
with t_kpi:
    st.header("Indicadores Principales")
    c1, c2, c3, c4 = st.columns(4)
    
    # Valores dinámicos si hay datos cargados
    e_total = datos["prmte"]["Energia"].sum() if datos["prmte"] is not None else 0
    km_total = datos["km"]["Distancia"].sum() if datos["km"] is not None else 1 # Evitar div/0
    
    c1.metric("Energía Total (PRMTE)", f"{e_total:,.0f} kWh")
    c2.metric("IDE Real", f"{(e_total/km_total):.2f} kWh/km")
    c3.metric("Cumplimiento Meta", "96.5%", "-0.5%")
    c4.metric("Trenes en Operación", "24/27")

# --- PESTAÑA 2: ISO 50001 ---
with t_iso:
    st.header("Cumplimiento SGE e ISO 50001")
    col_iso1, col_iso2 = st.columns(2)
    with col_iso1:
        st.subheader("Estado de Decretos")
        st.write("- **Decreto N°1 (2026):** En revisión de cumplimiento.")
        st.write("- **Decretos 2025:** Implementados en línea base.")
    with col_iso2:
        st.info("💡 Sugerencia: Actualizar la revisión energética trimestral según los consumos de SEAT validados.")

# --- PESTAÑA 3: MOTRICES / ODÓMETROS ---
with t_motrices:
    st.header("Control de Kilometraje y Odómetros")
    if datos["km"] is not None:
        col_km1, col_km2 = st.columns([2, 1])
        with col_km1:
            st.subheader("Kilometraje Diario por Motriz")
            st.dataframe(datos["km"], use_container_width=True)
        with col_km2:
            st.subheader("Conciliación UMR")
            st.warning("Revisar discrepancias > 2% en odómetros acumulados.")
    else:
        st.info("Cargue el archivo de Kilometraje UMR para visualizar las tablas.")

# --- PESTAÑA 4: CARGA ---
with t_carga:
    st.header("Análisis de Demanda y Carga")
    # Gráfico simple de ejemplo
    chart_data = pd.DataFrame(np.random.randn(24, 1), columns=['Demanda [kW]'])
    st.line_chart(chart_data)

# --- PESTAÑA 5: MALLA ---
with t_malla:
    st.header("Distribución de Energía por Malla")
    st.table(pd.DataFrame({
        "Subestación": ["Quilpué", "El Salto", "Limache"],
        "Consumo (kWh)": [450000, 520000, 480000],
        "Eficiencia": ["Óptima", "Alerta", "Óptima"]
    }))

# --- PESTAÑA 6: REPORTE DIARIO ---
with t_reporte:
    st.header("Generador de Reportes Oficiales")
    st.subheader("Resumen de Operación")
    if st.button("🚀 Generar Reporte PowerPoint"):
        st.success("Reporte generado con éxito. (Simulado)")
        # Aquí iría la integración con python-pptx

# --- PESTAÑA 7: SEAT DETALLE ---
with t_seat:
    st.header("Validación de Datos SEAT")
    if datos["seat"] is not None:
        st.write("Registros procesados con validación de consenso:")
        # Resaltar en rojo los valores que el consenso marcó como 0
        df_styled = datos["seat"].style.apply(lambda x: ['background-color: #ffcccc' if v == 0 else '' for v in x], subset=['Consenso'])
        st.dataframe(df_styled, use_container_width=True)
    else:
        st.info("Suba el archivo Excel de SEAT para aplicar la lógica de consenso.")

# --- PESTAÑA 8: COMPARADOR ---
with t_comp:
    st.header("Triangulación de Energía")
    st.markdown("""
    | Prioridad | Fuente | Uso |
    | :--- | :--- | :--- |
    | **1°** | **Factura** | Cierre mensual legal |
    | **2°** | **PRMTE** | Desglose por intervalos |
    | **3°** | **SEAT** | Respaldo y auditoría interna |
    """)
    
    if datos["prmte"] is not None:
        st.subheader("Cálculo de Energía Tracción")
        e_prmte = datos["prmte"]["Energia"].sum()
        deduccion = e_prmte * factor_bt
        e_traccion = e_prmte - deduccion
        
        st.metric("Energía PRMTE", f"{e_prmte:,.0f} kWh")
        st.metric("Deducción BT", f"-{deduccion:,.0f} kWh", delta=f"-{factor_bt*100}%")
        st.metric("Total Tracción Final", f"{e_traccion:,.0f} kWh", delta_color="off")
