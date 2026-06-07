import holidays

# --- 1. CONFIGURACIÓN DE PÁGINA ---
PAGE_CONFIG = {
    "page_title": "Gestión de Energía - Dashboard SGE",
    "layout": "wide",
    "page_icon": "🚆"
}

# --- 2. FERIADOS OFICIALES ---
CHILE_HOLIDAYS = holidays.Chile()

# --- 3. CONSTANTES DE RED (INFRAESTRUCTURA) ---
ESTACIONES = [
    'Puerto','Bellavista','Francia','Baron','Portales','Recreo','Miramar',
    'Viña del Mar','Hospital','Chorrillos','El Salto','Valencia','Quilpue',
    'El Sol','El Belloto','Las Americas','La Concepcion','Villa Alemana',
    'Sargento Aldea','Peñablanca','Limache'
]

# --- 4. CARPETAS DE DATOS ---
DATA_DIRS = {
    "v1": "data/thdr_v1",
    "v2": "data/thdr_v2",
    "umr": "data/umr",
    "seat": "data/seat",
    "bill": "data/facturacion",
    "carga_v1": "data/carga_v1", 
    "carga_v2": "data/carga_v2"
}
