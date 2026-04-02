import streamlit as st
import pandas as pd
import numpy as np

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Gestión Energética - EFE Valparaíso", layout="wide")

# --- ESTILOS PERSONALIZADOS (CSS) ---
st.markdown("""
    <style>
    .main { background-color: #f5f7f9; }
    .stTabs [data-baseweb="tab-list"] { gap: 10px; }
    .stTabs [data-baseweb="tab"] {
        background-color: #e1e4e8;
        border-radius: 4px 4px 0px 0px;
        padding: 10px 20px;
    }
    .stTabs [aria-selected="true"] { background-color: #00548b !important; color: white !important; }
    </style>
    """, unsafe_allow_html=True)

# --- TÍTULO Y FILTROS FLOTANTES (POPOVER) ---
st.title("⚡ Sistema de Gestión de Energía (SGE) - EFE Valparaíso")

with st.popover("🔍 Filtros de Visualización"):
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        fecha_sel = st.date_input("Seleccionar Mes/Año")
    with col_f2:
        tipo_tren = st.multiselect("Modelo de Tren", ["Xtrapolis", "Modular"], default=["Xtrapolis"])

# --- LÓGICA DE NAVEGACIÓN (8 PESTAÑAS) ---
tabs = st.tabs([
    "📊 KPIs", 
    "📜 ISO 50001", 
    "🚆 Motrices / Odómetros", 
    "🔌 Carga", 
    "🧬 Malla", 
    "📅 Reporte Diario", 
    "📉 SEAT Detalle", 
    "⚖️ Comparador (Factura vs SEAT)"
])

# --- 1. PESTAÑA KPIs ---
with tabs[0]:
    st.header("Indicadores Clave de Desempeño")
    st.info("Resumen ejecutivo del desempeño energético mensual.")
    # Ejemplo de fórmula LaTeX para CEE
    st.latex(r"CEE = \frac{\text{Energía Tracción [kWh]}}{\text{Kilometraje Total [km]}}")
    
    col1, col2, col3 = st.columns(3)
    col1.metric("IDE Real", "X.XX kWh/km", "-2%")
    col2.metric("Consumo Total", "XXXX MWh", "5%")
    col3.metric("Cumplimiento Meta", "98%", "1%")

# --- 2. PESTAÑA ISO 50001 ---
with tabs[1]:
    st.header("Gestión de Energía ISO 50001")
    st.write("Seguimiento de planes de acción y brechas del SGE.")
    # Aquí iría la lógica de los decretos 2025/2026 mencionados en tu contexto

# --- 3. PESTAÑA MOTRICES / ODÓMETROS ---
with tabs[2]:
    st.header("Control de Kilometraje por Motriz")
    
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        st.subheader("Kilometraje Diario [km]")
        # Placeholder para tabla de KM diario
        st.dataframe(pd.DataFrame({"Motriz": ["M1", "M2"], "KM": [150, 145]}))
        
    with col_m2:
        st.subheader("Lectura de Odómetro / Acumulado [km]")
        st.caption("Uso exclusivo para conciliación con Kilometraje Oficial (UMR)")
        # Placeholder para tabla de acumulados
        st.dataframe(pd.DataFrame({"Motriz": ["M1", "M2"], "Acumulado UMR": [50200, 48900]}))

# --- 4. PESTAÑA CARGA ---
with tabs[3]:
    st.header("Gestión de Carga Eléctrica")
    st.write("Análisis de demanda máxima y perfiles de carga en subestaciones.")

# --- 5. PESTAÑA MALLA ---
with tabs[4]:
    st.header("Operación de Malla Ferroviaria")
    st.write("Visualización de consumos segmentados por tramos de vía.")

# --- 6. PESTAÑA REPORTE DIARIO ---
with tabs[5]:
    st.header("Generación de Reportes")
    if st.button("Exportar a PowerPoint"):
        st.success("Generando presentación con indicadores mensuales...")
        # La lógica de generación de PPT va aquí

# --- 7. PESTAÑA SEAT DETALLE ---
with tabs[6]:
    st.header("Detalle de Energía SEAT")
    st.subheader("Consenso de Datos")
    st.warning("Nota: Si existen registros duplicados con valores distintos, el sistema asignará 0 para revisión manual.")
    
    # Lógica de Consenso simulada
    # df_seat['Energia'] = np.where(df_seat.duplicated(subset=['Fecha', 'Subestacion'], keep=False) & 
    #                              (df_seat['Valor'] != df_seat['Valor_Ref']), 0, df_seat['Valor'])

# --- 8. PESTAÑA COMPARADOR ---
with tabs[7]:
    st.header("Conciliación: Factura vs PRMTE vs SEAT")
    
    st.markdown("""
    **Jerarquía de Datos (Triangulación):**
    1. **Facturación:** Valor oficial para cierre mensual.
    2. **PRMTE:** Respaldo técnico de perfiles.
    3. **SEAT:** Datos internos de telemetría.
    """)
    
    # Lógica de deducción de Baja Tensión
    st.write("Cálculo de Energía Tracción (Deducción de consumos de servicios auxiliares/Baja Tensión).")
    
    # Tabla comparativa placeholder
    st.table(pd.DataFrame({
        "Fuente": ["Factura", "PRMTE", "SEAT"],
        "Energía [kWh]": [1250000, 1248000, 1230000],
        "Diferencia %": ["0.0%", "-0.16%", "-1.6%"]
    }))

# --- PIE DE PÁGINA ---
st.divider()
st.caption("EFE Valparaíso - Sistema de Monitoreo de Eficiencia Energética v2.0")
