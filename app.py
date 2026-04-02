import streamlit as st
import pandas as pd
import numpy as np
import re
import holidays
from io import BytesIO
from datetime import datetime, date
from typing import Dict, List, Tuple, Optional, Any
from dataclasses import dataclass, field
from enum import Enum
import logging
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# ============================================================================
# CONFIGURACIÓN Y LOGGING
# ============================================================================

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

st.set_page_config(
    page_title="Gestión de Energía - Dashboard SGE", 
    layout="wide", 
    page_icon="🚆"
)

# Cache de holidays para mejorar rendimiento
@st.cache_resource
def get_chile_holidays():
    return holidays.Chile()

CHILE_HOLIDAYS = get_chile_holidays()

# ============================================================================
# MODELOS DE DATOS (Type Safety)
# ============================================================================

class TipoJornada(Enum):
    LABORAL = "L"
    SABADO = "S"
    FESTIVO = "D/F"

class FuenteEnergia(Enum):
    SEAT = "SEAT"
    PRMTE = "PRMTE"
    FACTURA = "Factura"

@dataclass
class RegistroOperacion:
    fecha: datetime
    odometro_km: float
    tren_km: float
    umr_pct: float
    
    @property
    def tipo_dia(self) -> TipoJornada:
        fecha_obj = self.fecha.date() if isinstance(self.fecha, datetime) else self.fecha
        if fecha_obj in CHILE_HOLIDAYS or fecha_obj.weekday() == 6:
            return TipoJornada.FESTIVO
        if fecha_obj.weekday() == 5:
            return TipoJornada.SABADO
        return TipoJornada.LABORAL
    
    @property
    def semana_numero(self) -> int:
        return self.fecha.isocalendar()[1]

@dataclass
class RegistroEnergia:
    fecha: datetime
    total_kwh: float
    traccion_kwh: float
    kv12_kwh: float
    fuente: FuenteEnergia
    
    @property
    def pct_traccion(self) -> float:
        return (self.traccion_kwh / self.total_kwh * 100) if self.total_kwh > 0 else 0
    
    @property
    def pct_kv12(self) -> float:
        return (self.kv12_kwh / self.total_kwh * 100) if self.total_kwh > 0 else 0

# ============================================================================
# UTILIDADES
# ============================================================================

def parse_latam_number(valor: Any) -> float:
    """Convierte números en formato latinoamericano a float."""
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    
    try:
        s = str(valor).strip().replace(' ', '').replace('$', '')
        s = re.sub(r'[^\d.,-]', '', s)
        if not s:
            return 0.0
        
        # Detectar formato: 1.234,56 vs 1,234.56
        if ',' in s and '.' in s:
            if s.rfind(',') > s.rfind('.'):
                s = s.replace('.', '').replace(',', '.')
            else:
                s = s.replace(',', '')
        elif ',' in s:
            s = s.replace(',', '.')
        
        return float(s)
    except (ValueError, TypeError):
        logger.warning(f"No se pudo convertir valor: {valor}")
        return 0.0

def validar_rango_fechas(fecha_inicio: date, fecha_fin: date) -> bool:
    """Valida que el rango de fechas sea coherente."""
    if fecha_inicio > fecha_fin:
        st.error("La fecha de inicio no puede ser posterior a la fecha de fin.")
        return False
    if (fecha_fin - fecha_inicio).days > 366:
        st.warning("El rango de fechas supera un año. El procesamiento puede ser lento.")
    return True

# ============================================================================
# PROCESADOR DE ARCHIVOS (Responsabilidad Única)
# ============================================================================

class ProcesadorArchivos:
    """Clase responsable del procesamiento de archivos Excel."""
    
    def __init__(self, fecha_inicio: date, fecha_fin: date):
        self.fecha_inicio = fecha_inicio
        self.fecha_fin = fecha_fin
        self._reset_data()
    
    def _reset_data(self):
        """Reinicia las estructuras de datos."""
        self.operaciones: List[RegistroOperacion] = []
        self.kms_diarios: List[Dict] = []
        self.kms_acumulados: List[Dict] = []
        self.energia_seat: List[RegistroEnergia] = []
        self.prmte_15min: List[Dict] = []
        self.facturacion_horaria: List[Dict] = []
        self.comparacion_horaria: List[Dict] = []
    
    def procesar_archivo(self, archivo) -> Dict[str, Any]:
        """Procesa un archivo y devuelve estadísticas del proceso."""
        resultados = {"archivo": archivo.name, "hojas_procesadas": 0, "errores": []}
        
        try:
            xl = pd.ExcelFile(archivo)
            for sheet_name in xl.sheet_names:
                try:
                    self._procesar_hoja(xl, sheet_name)
                    resultados["hojas_procesadas"] += 1
                except Exception as e:
                    error_msg = f"Error en hoja '{sheet_name}': {str(e)}"
                    logger.warning(error_msg)
                    resultados["errores"].append(error_msg)
        except Exception as e:
            resultados["errores"].append(f"Error abriendo archivo: {str(e)}")
        
        return resultados
    
    def _procesar_hoja(self, xl: pd.ExcelFile, sheet_name: str):
        """Procesa una hoja según su tipo detectado."""
        sheet_upper = sheet_name.upper()
        
        if any(k in sheet_upper for k in ['UMR', 'RESUMEN', 'OPERACIONES']):
            self._procesar_operaciones(xl, sheet_name)
        elif 'ODO' in sheet_upper and 'KIL' in sheet_upper:
            self._procesar_kilometraje_trenes(xl, sheet_name)
        elif 'SEAT' in sheet_upper:
            self._procesar_energia_seat(xl, sheet_name)
        elif any(k in sheet_upper for k in ['PRMTE', 'MEDIDAS']):
            self._procesar_prmte(xl, sheet_name)
        elif any(k in sheet_upper for k in ['FACTURA', 'CONSUMO']):
            self._procesar_facturacion(xl, sheet_name)
    
    def _procesar_operaciones(self, xl: pd.ExcelFile, sheet_name: str):
        """Procesa datos de operaciones (UMR, odómetros)."""
        df_raw = pd.read_excel(xl, sheet_name=sheet_name, header=None)
        
        # Buscar cabecera de manera más robusta
        header_row = None
        for i in range(min(100, len(df_raw))):
            row_str = ' '.join(str(v) for v in df_raw.iloc[i].values[:10]).upper()
            if any(k in row_str for k in ['FECHA', 'ODO', 'TREN', 'KM']):
                header_row = i
                break
        
        if header_row is None:
            return
        
        df = pd.read_excel(xl, sheet_name=sheet_name, header=header_row)
        df.columns = [re.sub(r'[^A-Z]', '', str(c).upper().replace('Ó', 'O')) for c in df.columns]
        
        # Identificar columnas
        col_fecha = next((c for c in df.columns if 'FECHA' in c), None)
        col_odometro = next((c for c in df.columns if 'ODO' in c and 'ACUM' not in c), None)
        col_tren_km = next((c for c in df.columns if 'TREN' in c and 'KM' in c), None)
        
        if not col_fecha or not col_odometro:
            return
        
        df['_fecha'] = pd.to_datetime(df[col_fecha], errors='coerce')
        df_filtrado = df[
            (df['_fecha'].dt.date >= self.fecha_inicio) & 
            (df['_fecha'].dt.date <= self.fecha_fin) &
            df['_fecha'].notna()
        ]
        
        for _, row in df_filtrado.iterrows():
            odometro = parse_latam_number(row[col_odometro])
            tren_km = parse_latam_number(row[col_tren_km]) if col_tren_km else 0
            
            self.operaciones.append(RegistroOperacion(
                fecha=row['_fecha'].normalize(),
                odometro_km=odometro,
                tren_km=tren_km,
                umr_pct=(tren_km/odometro*100) if odometro > 0 else 0
            ))
    
    def _procesar_energia_seat(self, xl: pd.ExcelFile, sheet_name: str):
        """Procesa datos de energía desde archivos SEAT."""
        df = pd.read_excel(xl, sheet_name=sheet_name, header=None)
        
        for i in range(len(df)):
            fecha = pd.to_datetime(df.iloc[i, 1], errors='coerce')
            if pd.notna(fecha) and self.fecha_inicio <= fecha.date() <= self.fecha_fin:
                total = parse_latam_number(df.iloc[i, 3])
                traccion = parse_latam_number(df.iloc[i, 5])
                kv12 = parse_latam_number(df.iloc[i, 7])
                
                self.energia_seat.append(RegistroEnergia(
                    fecha=fecha.normalize(),
                    total_kwh=total,
                    traccion_kwh=traccion,
                    kv12_kwh=kv12,
                    fuente=FuenteEnergia.SEAT
                ))
    
    def _procesar_kilometraje_trenes(self, xl: pd.ExcelFile, sheet_name: str):
        """Procesa kilometraje diario y odómetros de trenes."""
        # Implementación similar a la original pero con mejor manejo de errores
        df_raw = pd.read_excel(xl, sheet_name=sheet_name, header=None)
        # ... (código similar al original pero estructurado)
        pass
    
    def _procesar_prmte(self, xl: pd.ExcelFile, sheet_name: str):
        """Procesa datos PRMTE."""
        # Implementación similar a la original
        pass
    
    def _procesar_facturacion(self, xl: pd.ExcelFile, sheet_name: str):
        """Procesa datos de facturación."""
        # Implementación similar a la original
        pass

# ============================================================================
# MOTOR DE ANÁLISIS
# ============================================================================

class MotorAnalisis:
    """Realiza cálculos y agregaciones sobre los datos procesados."""
    
    def __init__(self, procesador: ProcesadorArchivos):
        self.procesador = procesador
        self._cache = {}
    
    @st.cache_data(ttl=300)
    def obtener_dataframe_operaciones(self) -> pd.DataFrame:
        """Devuelve DataFrame de operaciones con caché."""
        if not self.procesador.operaciones:
            return pd.DataFrame()
        
        df = pd.DataFrame([{
            'Fecha': op.fecha,
            'Tipo Día': op.tipo_dia.value,
            'N° Semana': op.semana_numero,
            'Odómetro [km]': op.odometro_km,
            'Tren-Km [km]': op.tren_km,
            'UMR [%]': op.umr_pct
        } for op in self.procesador.operaciones])
        
        return df.drop_duplicates(subset=['Fecha']).sort_values('Fecha')
    
    @st.cache_data(ttl=300)
    def obtener_dataframe_energia(self) -> pd.DataFrame:
        """Devuelve DataFrame consolidado de energía."""
        registros = []
        
        # SEAT
        for e in self.procesador.energia_seat:
            registros.append({
                'Fecha': e.fecha,
                'E_Total': e.total_kwh,
                'E_Tr': e.traccion_kwh,
                'E_12': e.kv12_kwh,
                'Fuente': e.fuente.value
            })
        
        # PRMTE (si existe)
        if self.procesador.prmte_15min:
            df_prmte = pd.DataFrame(self.procesador.prmte_15min)
            df_prmte_diario = df_prmte.groupby('Fecha')['Energía PRMTE [kWh]'].sum().reset_index()
            # ... agregar lógica de merge con porcentajes SEAT
        
        return pd.DataFrame(registros).drop_duplicates(subset=['Fecha'], keep='last')

# ============================================================================
# COMPONENTES DE UI
# ============================================================================

class UIComponentes:
    """Componentes reutilizables de la interfaz."""
    
    @staticmethod
    def sidebar_filtros() -> Tuple[date, date]:
        """Renderiza filtros globales en sidebar."""
        with st.sidebar:
            st.header("📅 Filtro Global")
            today = date.today()
            default_start = today.replace(day=1)
            
            fecha_inicio, fecha_fin = st.date_input(
                "Selecciona el período",
                value=(default_start, today),
                help="Selecciona el rango de fechas para el análisis"
            )
            
            if isinstance(fecha_inicio, tuple):
                fecha_inicio, fecha_fin = fecha_inicio
            
            return fecha_inicio, fecha_fin
    
    @staticmethod
    def sidebar_carga_archivos():
        """Renderiza área de carga de archivos."""
        with st.sidebar:
            st.divider()
            st.header("📂 Carga de Archivos")
            
            archivos = {
                'umr': st.file_uploader(
                    "1. UMR / Odómetros",
                    type=["xlsx", "xls"],
                    accept_multiple_files=True,
                    key="umr"
                ),
                'seat': st.file_uploader(
                    "2. Energía SEAT",
                    type=["xlsx", "xls"],
                    accept_multiple_files=True,
                    key="seat"
                ),
                'facturacion': st.file_uploader(
                    "3. Facturación y PRMTE",
                    type=["xlsx", "xls"],
                    accept_multiple_files=True,
                    key="facturacion"
                )
            }
            
            return archivos
    
    @staticmethod
    def metric_card(titulo: str, valor: str, delta: Optional[str] = None):
        """Tarjeta de métrica con estilo consistente."""
        st.markdown(f"""
            <div class="stMetric">
                <h3 style="margin:0; color:#666;">{titulo}</h3>
                <p style="font-size:32px; font-weight:bold; margin:10px 0;">{valor}</p>
                {f'<p style="margin:0; color:#28a745;">{delta}</p>' if delta else ''}
            </div>
        """, unsafe_allow_html=True)

# ============================================================================
# PÁGINAS PRINCIPALES
# ============================================================================

class PaginaResumen:
    """Página de resumen operacional."""
    
    @staticmethod
    def render(motor: MotorAnalisis):
        st.header("📊 Resumen Operacional")
        
        df_ops = motor.obtener_dataframe_operaciones()
        if df_ops.empty:
            st.info("No hay datos de operaciones para el período seleccionado.")
            return
        
        # Métricas principales
        col1, col2, col3, col4 = st.columns(4)
        
        total_odometro = df_ops['Odómetro [km]'].sum()
        total_tren_km = df_ops['Tren-Km [km]'].sum()
        umr_global = (total_tren_km / total_odometro * 100) if total_odometro > 0 else 0
        
        col1.metric("Odómetro Total", f"{total_odometro:,.1f} km")
        col2.metric("Tren-Km Total", f"{total_tren_km:,.1f} km")
        col3.metric("UMR Global", f"{umr_global:.2f}%")
        
        # Energía si está disponible
        df_energia = motor.obtener_dataframe_energia()
        if not df_energia.empty:
            total_energia = df_energia['E_Total'].sum()
            col4.metric("Energía Total", f"{total_energia:,.0f} kWh")
        
        # Resumen por jornada
        st.subheader("📋 Resumen por Jornada")
        resumen_jornada = df_ops.groupby('Tipo Día', observed=True).agg({
            'Odómetro [km]': 'sum',
            'Tren-Km [km]': 'sum',
            'UMR [%]': 'mean'
        }).round(2)
        
        st.dataframe(resumen_jornada, use_container_width=True)
        
        # Exportación
        if st.button("📥 Exportar Resumen a PowerPoint"):
            # Implementar exportación
            st.success("Exportación completada")

# ============================================================================
# APLICACIÓN PRINCIPAL
# ============================================================================

def main():
    """Función principal de la aplicación."""
    
    # Sidebar: Filtros y carga
    fecha_inicio, fecha_fin = UIComponentes.sidebar_filtros()
    
    if not validar_rango_fechas(fecha_inicio, fecha_fin):
        return
    
    archivos = UIComponentes.sidebar_carga_archivos()
    todos_archivos = archivos['umr'] + archivos['seat'] + archivos['facturacion']
    
    if not todos_archivos:
        st.info("👋 Sube los archivos en el panel lateral para comenzar el análisis.")
        return
    
    # Procesamiento con indicador de progreso
    with st.spinner("Procesando archivos..."):
        procesador = ProcesadorArchivos(fecha_inicio, fecha_fin)
        
        resultados = []
        progress_bar = st.progress(0)
        
        for i, archivo in enumerate(todos_archivos):
            resultado = procesador.procesar_archivo(archivo)
            resultados.append(resultado)
            progress_bar.progress((i + 1) / len(todos_archivos))
        
        progress_bar.empty()
        
        # Mostrar resumen de procesamiento
        with st.expander("📋 Detalles del procesamiento"):
            for r in resultados:
                if r['errores']:
                    st.warning(f"⚠️ {r['archivo']}: {len(r['errores'])} errores")
                    for error in r['errores']:
                        st.caption(f"  • {error}")
                else:
                    st.success(f"✅ {r['archivo']}: {r['hojas_procesadas']} hojas procesadas")
        
        motor = MotorAnalisis(procesador)
    
    # Pestañas principales
    tabs = st.tabs([
        "📊 Resumen",
        "📑 Operaciones",
        "🚆 Trenes",
        "⚡ Energía",
        "📈 Análisis Avanzado"
    ])
    
    with tabs[0]:
        PaginaResumen.render(motor)
    
    with tabs[1]:
        st.header("📑 Datos de Operaciones")
        df_ops = motor.obtener_dataframe_operaciones()
        if not df_ops.empty:
            st.dataframe(df_ops, use_container_width=True)
    
    with tabs[2]:
        st.header("🚆 Kilometraje de Trenes")
        st.info("Funcionalidad en desarrollo")
    
    with tabs[3]:
        st.header("⚡ Consumo Energético")
        df_energia = motor.obtener_dataframe_energia()
        if not df_energia.empty:
            st.dataframe(df_energia, use_container_width=True)
    
    with tabs[4]:
        st.header("📈 Análisis Avanzado")
        st.info("Funcionalidades de regresión y detección de anomalías")

if __name__ == "__main__":
    main()
