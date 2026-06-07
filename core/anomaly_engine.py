import pandas as pd
import numpy as np
from utils.helpers import _norm

def _robust_z_score(series):
    """Calcula el Z-Score robusto usando la Mediana y MAD (Median Absolute Deviation)."""
    median = series.median()
    mad = (series - median).abs().median()
    # Factor de consistencia 1.4826 para normalizar a la desviación estándar
    mad = mad * 1.4826
    if mad == 0: return pd.Series(0, index=series.index)
    return (series - median) / mad

def diagnosticar_anomalias(df, columna_objetivo, umbral=3.0):
    """
    Identifica filas anómalas basándose en la desviación robusta.
    Retorna un DataFrame con una columna 'Es_Anomalia' (Booleana).
    """
    if columna_objetivo not in df.columns:
        return df
    
    # Calcular Z-Score robusto
    z_scores = _robust_z_score(df[columna_objetivo])
    df['Z_Score_Robust'] = z_scores
    df['Es_Anomalia'] = df['Z_Score_Robust'].abs() > umbral
    
    return df

def perfil_horario_diario(df, col_valor, col_fecha='Fecha_Op'):
    """Genera el perfil promedio diario para comparar días futuros."""
    # Agrupa por día y calcula media (o mediana)
    perfil = df.groupby(col_fecha)[col_valor].mean().reset_index()
    return perfil

def analizar_eficiencia_energia(df_ops):
    """
    Análisis crítico para detectar si el consumo de energía es ineficiente 
    comparado con la mediana histórica.
    """
    if 'IDE (kWh/km)' not in df_ops.columns:
        return df_ops
    
    # Aplicar detección de anomalías al índice de eficiencia
    return diagnosticar_anomalias(df_ops, 'IDE (kWh/km)', umbral=2.5)

# Fuente recomendada: 'Robust Statistics' (Hampel, Ronchetti, Rousseeuw, Stahel)
# El uso de MAD es superior a la Desviación Estándar clásica en entornos ferroviarios
# porque los datos operativos (THDR) suelen tener colas pesadas o valores atípicos 
# por fallas técnicas que sesgarían una media simple.
