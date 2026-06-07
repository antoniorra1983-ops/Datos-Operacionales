import pandas as pd
import numpy as np
import re
from datetime import datetime, date, time
from utils.helpers import make_columns_unique, parse_latam_number, _norm

def convertir_a_minutos(val):
    """Transforma horas (Ej: 14:30:00) a minutos decimales para cálculos matemáticos."""
    if pd.isna(val) or str(val).strip() == "": return None
    try:
        if isinstance(val, (datetime, time)): return val.hour*60 + val.minute + val.second/60.0
        sv = str(val).strip()
        m = re.search(r'(\d{1,2}):(\d{2}):(\d{2})', sv)
        if m: return int(m.group(1))*60 + int(m.group(2)) + int(m.group(3))/60.0
        m = re.search(r'(\d{1,2}):(\d{2})', sv)
        if m: return int(m.group(1))*60 + int(m.group(2))
        return None
    except Exception: return None

def parsear_fecha_nombre(nombre):
    """Adivina la fecha de la operación leyendo el nombre del archivo Excel."""
    nombre = str(nombre)
    patrones = [
        (r'(\d{2})[-_](\d{2})[-_](\d{4})', lambda m: date(int(m.group(3)), int(m.group(2)), int(m.group(1)))),
        (r'(\d{4})[-_](\d{2})[-_](\d{2})', lambda m: date(int(m.group(1)), int(m.group(2)), int(m.group(3)))),
        (r'(\d{8})', lambda m: date(int(m.group(1)[4:]), int(m.group(1)[2:4]), int(m.group(1)[:2]))),
        (r'(\d{6})', lambda m: date(2000+int(m.group(1)[4:]), int(m.group(1)[2:4]), int(m.group(1)[:2])))
    ]
    for patron, parser in patrones:
        m = re.search(patron, nombre)
        if m:
            try: return parser(m), f"Match ({m.group()})"
            except ValueError: pass
    return None, f"sin fecha en: '{nombre}'"

def procesar_thdr_eficiente(file, start_date, end_date):
    """Procesa el archivo THDR (Horarios) blindado contra errores de columnas."""
    nombre = getattr(file, 'name', str(file))
    diag = {"archivo": nombre, "fecha_parseada": None, "en_rango": None, "filas": 0, "error": None}
    try:
        fch_date, desc = parsear_fecha_nombre(nombre)
        diag["fecha_parseada"] = desc
        if fch_date is None:
            diag["error"] = "No se encontró fecha en el nombre"; return pd.DataFrame(), diag
        
        diag["en_rango"] = f"{start_date}≤{fch_date}≤{end_date}"
        if not (start_date <= fch_date <= end_date):
            diag["error"] = "Fecha fuera del rango"; return pd.DataFrame(), diag
            
        fch_dt = pd.to_datetime(fch_date).normalize()
        engine = "xlrd" if nombre.lower().endswith(".xls") else "openpyxl"
        df_raw = pd.read_excel(file, header=None, engine=engine)
        
        r0 = df_raw.iloc[0].copy(); r0[0] = np.nan
        h1 = r0.ffill().astype(str); h2 = df_raw.iloc[1].fillna('').astype(str)
        
        cols = []
        for stn, tip in zip(h1, h2):
            stn, tip = str(stn).strip(), str(tip).strip()
            if stn == 'nan' or stn == '': cols.append(tip if tip else '_vacio')
            else: cols.append(f"{stn}_{tip}" if tip else stn)
            
        df = df_raw.iloc[5:].copy().reset_index(drop=True)
        n = len(df.columns)
        cols_adj = cols[:n] if len(cols) >= n else cols + [f"_C{j}" for j in range(n-len(cols))]
        df.columns = cols_adj
        df = make_columns_unique(df).dropna(how='all', axis=0).reset_index(drop=True)
        
        for col in df.columns:
            if any(k in str(col) for k in ['Hora Llegada','Hora Salida','Hora Salida Programada']):
                df[f"{col}_min"] = df[col].apply(convertir_a_minutos)
                
        if 'Unidad' in df.columns:
            df['Unidad'] = df['Unidad'].fillna('S').replace('','S')
        else:
            c_m2 = next((c for c in df.columns if 'Motriz 2' in str(c)), None)
            df['Unidad'] = df[c_m2].apply(lambda x: 'M' if parse_latam_number(x)>0 else 'S') if c_m2 else 'S'
            
        df['Tren-Km'] = 43.13 * df['Unidad'].apply(lambda x: 2 if str(x).strip()=='M' else 1)
        df['Fecha_Op'] = fch_dt
        
        col_ref = next((c for c in df.columns if ('PUERTO' in str(c).upper() or 'LIMACHE' in str(c).upper()) and 'Salida' in str(c) and '_min' in str(c)), None)
        if col_ref: df['Hora_Ref_Min'] = df[col_ref]
        
        diag["filas"] = len(df)
        return df, diag
    except Exception as e:
        diag["error"] = str(e); return pd.DataFrame(), diag

def procesar_carga_pasajeros(f, start_date, end_date):
    """Procesa el archivo comercial de Carga Tren con escudo semántico."""
    try:
        is_csv = f.name.lower().endswith('.csv')
        if is_csv:
            try: df = pd.read_csv(f, header=None, encoding='utf-8')
            except UnicodeDecodeError: 
                f.seek(0); df = pd.read_csv(f, header=None, encoding='latin-1')
        else:
            eu = "xlrd" if f.name.lower().endswith(".xls") else "openpyxl"
            df = pd.read_excel(f, engine=eu, header=None)
            
        # Busca el ancla de la tabla
        h_idx = next((i for i in range(min(30, len(df))) if 'THDR' in str(df.iloc[i].values).upper() or 'VIAJE' in str(df.iloc[i].values).upper()), None)
        
        if h_idx is not None:
            f.seek(0)
            if is_csv:
                try: df = pd.read_csv(f, header=h_idx, encoding='utf-8')
                except UnicodeDecodeError:
                    f.seek(0); df = pd.read_csv(f, header=h_idx, encoding='latin-1')
            else:
                df = pd.read_excel(f, engine=eu, header=h_idx)
                
            df.columns = [str(c).strip() for c in df.columns]
            
            # Buscar fecha
            c_fecha = next((c for c in df.columns if 'FECHA' in str(c).upper() or 'DIA' in str(c).upper() or 'DATE' in str(c).upper()), None)
            if c_fecha:
                df['Fecha'] = pd.to_datetime(df[c_fecha], dayfirst=True, errors='coerce').dt.normalize()
                df = df.dropna(subset=['Fecha'])
                df = df[(df['Fecha'].dt.date >= start_date) & (df['Fecha'].dt.date <= end_date)]
            elif 'Fecha' in df.columns:
                df['Fecha'] = pd.to_datetime(df['Fecha'], dayfirst=True, errors='coerce').dt.normalize()
                df = df.dropna(subset=['Fecha'])
                df = df[(df['Fecha'].dt.date >= start_date) & (df['Fecha'].dt.date <= end_date)]
            
            if df.empty:
                return pd.DataFrame()
                
            # Busca Totales y Estaciones críticas
            c_tot = next((c for c in df.columns if 'TOTAL' in str(c).upper() and 'BORDO' in str(c).upper()), None)
            if c_tot:
                df['Total a Bordo'] = pd.to_numeric(df[c_tot], errors='coerce').fillna(0)
            elif 'Total a Bordo' in df.columns:
                df['Total a Bordo'] = pd.to_numeric(df['Total a Bordo'], errors='coerce').fillna(0)
                
            c_est_max = next((c for c in df.columns if 'ESTACI' in str(c).upper() and 'MAX' in _norm(c)), None)
            if c_est_max:
                df['Estación Máxima'] = df[c_est_max]
                
            return df
        return pd.DataFrame()
    except Exception:
        return pd.DataFrame()
