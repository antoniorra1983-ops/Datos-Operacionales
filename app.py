import streamlit as st
from config.settings import PAGE_CONFIG
from utils.file_manager import combinar_fuentes
from core.data_loader import procesar_thdr_eficiente, procesar_carga_pasajeros
from core.anomaly_engine import analizar_eficiencia_energia

# --- 1. CONFIGURACIÓN INICIAL ---
# Esta configuración viene de tu archivo config/settings.py
st.set_page_config(**PAGE_CONFIG)

def main():
    st.title("🚆 Sistema de Gestión de Energía (SGE)")
    st.markdown("---")

    # --- 2. SIDEBAR: CARGA DE DATOS ---
    with st.sidebar:
        st.header("Configuración de Carga")
        uploaded_files = st.file_uploader("Subir Excels de Operación", accept_multiple_files=True)
        rango_fechas = st.date_input("Rango de fechas de análisis", [])

    # --- 3. LÓGICA DE ORQUESTACIÓN ---
    if not uploaded_files or not rango_fechas:
        st.info("Por favor, cargue los archivos y seleccione el rango de fechas para iniciar.")
        return

    # Procesamiento usando el motor centralizado (core/data_loader.py)
    with st.spinner("Procesando datos operativos..."):
        # Llama a la función del motor de carga
        df_ops, diag = procesar_thdr_eficiente(uploaded_files[0], rango_fechas[0], rango_fechas[1])
        
        if df_ops.empty:
            st.error(f"Error en procesamiento: {diag.get('error', 'Desconocido')}")
            return

        # Aplicamos el motor de anomalías (core/anomaly_engine.py)
        # Esto limpia y diagnostica los datos automáticamente
        df_ops = analizar_eficiencia_energia(df_ops)

    # --- 4. UI: DESPLIEGUE ---
    st.success("Análisis completado exitosamente")
    
    # Visualización basada en pestañas
    tab1, tab2 = st.tabs(["📊 Resumen de Datos", "🩺 Detección de Anomalías"])
    
    with tab1:
        st.write("Vista previa de los datos procesados:")
        st.dataframe(df_ops.head())
        
    with tab2:
        st.write("Registros identificados como anomalías (Z-Score > 2.5):")
        anomalias = df_ops[df_ops['Es_Anomalia'] == True]
        if not anomalias.empty:
            st.dataframe(anomalias)
        else:
            st.write("No se detectaron anomalías con el umbral configurado.")

if __name__ == "__main__":
    main()
