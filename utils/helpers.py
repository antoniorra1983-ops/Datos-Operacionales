import pandas as pd
import numpy as np
import re
from datetime import datetime
from config.settings import CHILE_HOLIDAYS

# --- FUNCIONES DE LIMPIEZA DE DATOS ---

def make_columns_unique(df):
    """Evita errores en Pandas cuando un Excel tiene columnas con el mismo nombre."""
    if not isinstance(df, pd.DataFrame) or df.empty: 
        return df
    cols = pd.Series(df.columns)
    for dup in cols[cols.duplicated()].unique():
        cols[cols == dup] = [f"{dup}_{i}" if i != 0 else dup for i in range(sum(cols == dup))]
    df.columns = cols
    return df

def parse_latam_number(val):
    """Convierte strings de formato latino ('1.500,50') a floats matemáticos."""
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    s = str(val).strip().replace(' ','').replace('$','')
    s = re.sub(r'[^\d.,-]','',s)
    if not s: return 0.0
    if ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'): s = s.replace('.','').replace(',','.')
        else: s = s.replace(',','')
    elif ',' in s: s = s.replace(',','.')
    try: return float(s)
    except ValueError: return 0.0

def _norm(s):
    """Normaliza strings: quita tildes y pasa a mayúsculas para cruces precisos."""
    return str(s).upper().translate(str.maketrans("ÁÉÍÓÚÜÑ", "AEIOUUN"))

# --- FUNCIONES DE MANEJO DE TIEMPO ---

def get_tipo_dia(fch):
    """Clasifica los días en L (Laboral), S (Sábado), D/F (Domingo o Festivo)."""
    if fch is None: return "N/A"
    if fch in CHILE_HOLIDAYS or fch.weekday() == 6: return "D/F"
    if fch.weekday() == 5: return "S"
    return "L"

def minutos_a_hhmmss(minutos_float):
    """Convierte un valor de minutos (float) al estándar ferroviario HH:MM:SS."""
    if pd.isna(minutos_float): return "00:00:00"
    sign = "-" if minutos_float < 0 else ""
    m_abs = abs(minutos_float)
    h = int(m_abs // 60)
    m = int(m_abs % 60)
    s = int(round((m_abs - int(m_abs)) * 60))
    return f"{sign}{h:02d}:{m:02d}:{s:02d}"
