"""Simulación y contraste de energía de tracción para Datos Operacionales.

Motor usado: archivos oficiales del repositorio antoniorra1983-ops/Simulador,
versión fija comprobada por SHA de blob de Git. Solo se descarga al solicitar
una simulación; no cambia la carga del historial ni los datos fuente.
"""
from __future__ import annotations

from datetime import date
import hashlib
import importlib
import os
from pathlib import Path
import re
import sys
import tempfile
from typing import Callable

import numpy as np
import pandas as pd
import streamlit as st

ENGINE_REPO = 'antoniorra1983-ops/Simulador'
ENGINE_COMMIT = 'aa2beb2d46ba83c20beb65475d6645aa0b1fc055'
ENGINE_BLOBS = {
    'config.py': 'a51a4bfddb401c0bfab054b53bfd47e3f8ca9cac',
    'etl_parser.py': 'da892546cb8272c9ce8c4bd3960a18a58d103860',
    'motor_fisico.py': '8a845991a07f849cf5705e596b976a8559098f97',
}
ENGINE_PACKAGE = 'efe_motor_fisico_' + ENGINE_COMMIT[:10]


def _verificar_blob(data: bytes, sha: str) -> bool:
    """Hash Git SHA-1 del blob exacto, no hash SHA-1 simple de bytes."""
    return hashlib.sha1(b'blob ' + str(len(data)).encode() + b'\0' + data).hexdigest() == sha


def cargar_motor():
    """Instala bajo /tmp una copia íntegra y verificada del motor físico ya existente."""
    try:
        return importlib.import_module(f'{ENGINE_PACKAGE}.motor_fisico')
    except ModuleNotFoundError:
        pass
    base = Path(tempfile.gettempdir()) / 'efe_simulacion_motor_pin_v1'
    pkg = base / ENGINE_PACKAGE
    pkg.mkdir(parents=True, exist_ok=True)
    (pkg / '__init__.py').touch()
    import requests
    for nombre, blob_sha in ENGINE_BLOBS.items():
        destino = pkg / nombre
        # Los archivos transformados para import relativo llevan una cabecera,
        # así que no se confía en su hash: cada ejecución valida la copia original
        # descargada, almacenada aparte, antes de importarla.
        original = pkg / (nombre + '.original')
        data = original.read_bytes() if original.is_file() else b''
        if not _verificar_blob(data, blob_sha):
            url = f'https://raw.githubusercontent.com/{ENGINE_REPO}/{ENGINE_COMMIT}/{nombre}'
            resp = requests.get(url, timeout=60, headers={'User-Agent': 'Datos-Operacionales-Simulacion/1.0'})
            resp.raise_for_status()
            data = resp.content
            if len(data) > 3_000_000 or not _verificar_blob(data, blob_sha):
                raise ValueError(f'No se pudo verificar la versión de {nombre} del motor físico.')
            original.write_bytes(data)
        src = data.decode('utf-8')
        if nombre == 'motor_fisico.py':
            # Aislar los módulos del simulador de cualquier configuración local.
            src = re.sub(r'(?m)^(\s*)import config\s*$', r'\1from . import config', src)
            src = re.sub(r'(?m)^from etl_parser import ', 'from .etl_parser import ', src)
        destino.write_text(src, encoding='utf-8')
    if str(base) not in sys.path:
        sys.path.insert(0, str(base))
    return importlib.import_module(f'{ENGINE_PACKAGE}.motor_fisico')


def _nserie(v):
    """Convierte el identificador de motriz o tren sin confundir 1.0 con 10."""
    if v is None or (isinstance(v, (float, np.floating)) and not np.isfinite(v)):
        return None
    s = str(v).strip()
    if not s or s.casefold() in ('nan', 'none', 'sin asignar'):
        return None
    m = re.search(r'(\d+)(?:\.0+)?$', s)
    return int(m.group(1)) if m else None


def _texto_servicio(v):
    n = _nserie(v)
    return str(n) if n is not None else str(v).strip().upper()


def fechas_thdr(*tablas):
    f = set()
    for t in tablas:
        if isinstance(t, pd.DataFrame) and not t.empty and 'Fecha_Op' in t:
            f.update(pd.to_datetime(t['Fecha_Op'], errors='coerce').dropna().dt.date)
    return sorted(f)


def construir_viajes(via1, via2, fecha, estaciones, km_acum, get_col_thdr: Callable):
    """Transforma las THDR ya procesadas a las entradas reales del motor.

    Solo viajes con motriz válida y dos estaciones con horas se simulan. Nunca
    se inventa una motriz ni se supone que todos recorren Puerto-Limache.
    """
    fecha = pd.Timestamp(fecha).normalize()
    filas, descartes = [], []
    count = {1: 0, 2: 0}
    for via, fuente in ((1, via1), (2, via2)):
        if not isinstance(fuente, pd.DataFrame) or fuente.empty or 'Fecha_Op' not in fuente:
            continue
        dia = fuente.loc[pd.to_datetime(fuente['Fecha_Op'], errors='coerce').dt.normalize() == fecha]
        if dia.empty:
            continue
        # Se obtienen las columnas una sola vez, no una búsqueda por cada viaje.
        columnas = [(get_col_thdr(dia, est, 'LLEGADA'), get_col_thdr(dia, est, 'SALIDA'))
                    for est in estaciones]
        col1 = next((c for c in dia if str(c).strip().upper() == 'MOTRIZ 1'), None)
        col2 = next((c for c in dia if str(c).strip().upper() == 'MOTRIZ 2'), None)
        servicio = next((c for c in dia if str(c).strip().upper() == 'TREN'), None)
        viaje = next((c for c in dia if str(c).strip().upper() == 'VIAJE'), None)
        for idx, row in dia.iterrows():
            mot1 = _nserie(row.get(col1)) if col1 else None
            mot2 = _nserie(row.get(col2)) if col2 else None
            if mot1 in range(1, 28):
                tipo = 'XT-100'
            elif mot1 in range(28, 36):
                tipo = 'XT-M'
            elif mot1 in range(410, 415):
                tipo = 'SFE'
            else:
                descartes.append(f'Vía {via}, fila {idx}: motriz no reconocida ({mot1}).')
                continue
            doble = bool(mot2 is not None or str(row.get('Unidad', '')).strip().upper() == 'M')
            if tipo == 'SFE' and doble:
                descartes.append(f'Vía {via}, fila {idx}: SFE marcado como doble, revisar THDR.')
                continue
            if mot2 is not None and mot2 not in range(1, 36):
                descartes.append(f'Vía {via}, fila {idx}: segunda motriz no reconocida ({mot2}).')
                continue
            estaciones_con_hora = []
            recorrido_indices = range(len(estaciones)) if via == 1 else range(len(estaciones)-1, -1, -1)
            for pos in recorrido_indices:
                cl, cs = columnas[pos]
                tiempos = []
                for col in (cl, cs):
                    if col:
                        try:
                            minuto = float(row.get(col))
                            if np.isfinite(minuto) and 0 <= minuto < 2880:
                                tiempos.append(minuto)
                        except (ValueError, TypeError):
                            pass
                if tiempos:
                    estaciones_con_hora.append((pos, sorted(set(tiempos))))
            if len(estaciones_con_hora) < 2:
                descartes.append(f'Vía {via}, fila {idx}: menos de dos estaciones con horarios.')
                continue
            nodos = []
            ultimo = None
            ajuste = 0
            for pos, tiempos in estaciones_con_hora:
                for tiempo in tiempos:
                    t = tiempo + ajuste
                    if ultimo is not None and t < ultimo - 720:
                        ajuste += 1440
                        t = tiempo + ajuste
                    # Diferencias pequeñas de minutos entre estaciones pueden ser
                    # errores de orden de columnas: no fabricar una cronología.
                    if ultimo is not None and t < ultimo:
                        continue
                    km = float(km_acum[pos])
                    if nodos and nodos[-1] == (t, km):
                        continue
                    nodos.append((t, km))
                    ultimo = t
            orig, dest = float(km_acum[estaciones_con_hora[0][0]]), float(km_acum[estaciones_con_hora[-1][0]])
            if not nodos or len({p[1] for p in nodos}) < 2 or abs(dest-orig) < 0.4:
                descartes.append(f'Vía {via}, fila {idx}: recorrido inválido o incompleto.')
                continue
            if (via == 1 and dest <= orig) or (via == 2 and dest >= orig):
                descartes.append(f'Vía {via}, fila {idx}: sentido inconsistente.')
                continue
            id_servicio = _texto_servicio(row.get(servicio)) if servicio else ''
            id_viaje = _texto_servicio(row.get(viaje)) if viaje else ''
            filas.append({
                '_id': f'{fecha.date().isoformat()}_V{via}_{id_servicio}_{id_viaje}_{nodos[0][0]:.2f}_{idx}',
                'Fecha_str': fecha.date().isoformat(), 'Via': via,
                'num_servicio': id_servicio, 'nro_viaje': id_viaje,
                'motriz_num': f'{mot1}+{mot2}' if mot2 else str(mot1),
                'tipo_tren': tipo, 'doble': doble, 't_ini': float(nodos[0][0]),
                't_fin': float(nodos[-1][0]), 'km_orig': orig, 'km_dest': dest,
                'km_viaje': abs(dest-orig), 'nodos': nodos,
                'pax_d': {}, 'pax_abordo': 0, 'maniobra': None,
                'Fuente THDR': f'Vía {via}',
            })
            count[via] += 1
    return pd.DataFrame(filas), descartes, count


def incorporar_pasajeros(viajes, carga_v1, carga_v2, fecha, pax_estimados, columnas_pax):
    """Solo asigna una carga real si el servicio y el día son inequívocos."""
    d = viajes.copy()
    d['pax_abordo'] = int(pax_estimados)
    d['pax_d'] = [{} for _ in range(len(d))]
    encontrados = 0
    for via, carga in ((1, carga_v1), (2, carga_v2)):
        if not isinstance(carga, pd.DataFrame) or carga.empty or 'Fecha' not in carga:
            continue
        dia = carga.loc[pd.to_datetime(carga['Fecha'], errors='coerce').dt.normalize() == pd.Timestamp(fecha).normalize()].copy()
        if dia.empty:
            continue
        cc = [c for c in columnas_pax if c in dia]
        campos_servicio = [c for c in dia.columns if str(c).strip().upper() in ('TREN', 'SERVICIO', 'NRO THDR', 'NRO_THDR')]
        if not cc or not campos_servicio:
            continue
        mapa = {}
        for _, r in dia.iterrows():
            for c in campos_servicio:
                key = _texto_servicio(r.get(c))
                if key and key.lower() not in ('none', 'nan'):
                    mapa.setdefault(key, []).append(r)
        for i, v in d[d['Via'] == via].iterrows():
            cs = mapa.get(v['num_servicio'], [])
            if len(cs) != 1:
                continue  # evitar emparejamientos ambiguos
            perfil = {k: max(0, int(float(cs[0][k]))) for k in cc
                      if pd.notna(cs[0][k]) and pd.to_numeric(cs[0][k], errors='coerce') >= 0}
            if len(perfil) < 2:
                continue
            d.at[i, 'pax_d'] = perfil
            d.at[i, 'pax_abordo'] = max(perfil.values())
            encontrados += 1
    return d, encontrados


def ejecutar_motor(viajes, pct_traccion, usar_pendientes, usar_riel_mojado,
                   usar_regeneracion, receptividad_intervalos, estacion):
    motor = cargar_motor()
    captacion = {}
    if usar_regeneracion and receptividad_intervalos:
        captacion = motor.calcular_receptividad_por_headway(viajes)
    # El motor original tiene un gobernador que reemplaza la selección de
    # tracción por 75 % o 50 % cuando recibe Fecha_str. Para que el control
    # de esta pestaña sea funcional, entregamos Fecha_str vacío solo durante
    # el cálculo, y restauramos la fecha original en los resultados.
    entrada = viajes.copy()
    entrada['Fecha_str'] = ''
    salida = motor.calcular_termodinamica_flota_v111(
        entrada, float(pct_traccion), bool(usar_pendientes), bool(usar_riel_mojado),
        bool(usar_regeneracion), captacion, estacion_anio=estacion, prevenciones=[]
    )
    if isinstance(salida, pd.DataFrame) and len(salida) == len(viajes):
        salida['Fecha_str'] = viajes['Fecha_str'].to_numpy()
    return salida


def formatear(v, dec=1):
    try:
        if v is None or not np.isfinite(float(v)):
            return '—'
        s = f'{float(v):,.{dec}f}'
        return s.replace(',', 'X').replace('.', ',').replace('X', '.')
    except (TypeError, ValueError):
        return '—'


def render_simulacion(v1, v2, carga1, carga2, df_ops, estaciones, km_acum,
                      get_col_thdr, columnas_pax, clasificar_jornada=None):
    st.markdown('### ⚡ Simulación energética · THDR vs consumo real')
    st.caption('Modelo físico del repositorio Simulador, con las THDR y la energía del historial de Datos Operacionales. '
               'Los filtros globales delimitan las fechas que puedes simular.')
    st.info('La simulación representa el consumo estimado en los trenes; el consumo real corresponde al '
            'dato de tracción disponible en Datos Operacionales (Factura → PRMTE → SEAT). '
            'No son mediciones idénticas en el punto eléctrico y pueden diferir por pérdidas, operación en vacío y datos faltantes.')
    fechas = fechas_thdr(v1, v2)
    if not fechas:
        st.warning('No hay THDR disponibles con los filtros actuales. Amplía el período o revisa la carga del historial.')
        return
    fecha = st.selectbox('📅 Día a simular', fechas[::-1], format_func=lambda d: d.strftime('%d-%m-%Y'),
                         key='sim_energia_fecha')
    # Una fila por fecha de la misma fuente del Dashboard.
    real = pd.DataFrame()
    if isinstance(df_ops, pd.DataFrame) and not df_ops.empty and 'Fecha' in df_ops:
        real = df_ops[pd.to_datetime(df_ops['Fecha'], errors='coerce').dt.normalize() == pd.Timestamp(fecha)].copy()
    traccion_real = np.nan
    fuente_real = 'Sin medición válida'
    km_real = np.nan
    if not real.empty:
        tr = pd.to_numeric(real.get('E_Tr', pd.Series(dtype=float)), errors='coerce')
        if tr.notna().any() and tr.max() > 0:
            traccion_real = float(tr.sum(min_count=1))
            fuentes = real['Fuente'].astype(str).unique() if 'Fuente' in real else []
            fuente_real = ', '.join(fuentes) or 'Historial'
        if 'Odómetro [km]' in real:
            km = pd.to_numeric(real['Odómetro [km]'], errors='coerce')
            if km.notna().any() and km.sum() > 0:
                km_real = float(km.sum())
    col0, col1, col2 = st.columns(3)
    col0.metric('Tracción real registrada', formatear(traccion_real, 0)+' kWh')
    col1.metric('Fuente de energía', fuente_real)
    col2.metric('Odómetro UMR del día', formatear(km_real, 1)+' km')
    if not np.isfinite(traccion_real):
        st.warning('No hay energía de tracción válida para esta fecha; podrás simularla, pero no calcular el error frente al consumo real.')
    elif fuente_real != 'SEAT':
        st.caption('La energía de tracción procedente de Factura/PRMTE puede ser estimada '
                   'a partir de la proporción de tracción observada en SEAT; no equivale '
                   'necesariamente a una medición directa de tracción.')
    with st.expander('⚙️ Parámetros físicos de simulación', expanded=True):
        j = clasificar_jornada(fecha) if clasificar_jornada else ('L' if fecha.weekday() < 5 else 'S' if fecha.weekday() == 5 else 'D/F')
        c1, c2, c3 = st.columns(3)
        traccion = c1.slider('Tracción máxima (%)', min_value=30, max_value=100,
                            value=75 if j == 'L' else 50, step=5, key='sim_energia_pct')
        pax_estimados = c2.number_input('PAX estimados por tren sin carga', min_value=0,
                                       max_value=900, value=150, step=10, key='sim_energia_pax')
        riel = c3.checkbox('Riel mojado', value=False, key='sim_energia_riel')
        pendientes = st.checkbox('Considerar pendientes', value=True, key='sim_energia_pendientes')
        regeneracion = st.checkbox('Considerar regeneración', value=True, key='sim_energia_regen')
        headway = st.checkbox('Receptividad según intervalo de trenes', value=True,
                              disabled=not regeneracion, key='sim_energia_headway')
        st.caption('El modelo identifica tipo XT-100, XT-M y SFE mediante la motriz de cada THDR; '
                   'los horarios y paradas se obtienen de sus columnas por estación.')
    if st.button('▶️ Simular y comparar el consumo del día', type='primary', key='sim_energia_run',
                 use_container_width=True):
        viajes, descartes, cobertura = construir_viajes(v1, v2, fecha, estaciones, km_acum, get_col_thdr)
        if viajes.empty:
            st.error('No se encontraron viajes con horarios, recorrido y motriz válidos para esta fecha.')
            with st.expander('Ver viajes descartados'):
                st.write(descartes[:70])
            return
        viajes, encontrados = incorporar_pasajeros(viajes, carga1, carga2, fecha, pax_estimados, columnas_pax)
        mes = int(fecha.month)
        estacion = ('verano' if mes in (12,1,2) else 'otoño' if mes in (3,4,5)
                    else 'invierno' if mes in (6,7,8) else 'primavera')
        with st.spinner('Ejecutando el motor físico de Simulador con los servicios THDR...'):
            try:
                res = ejecutar_motor(viajes, traccion, pendientes, riel,
                                    regeneracion, headway, estacion)
            except Exception as exc:
                st.error(f'No se pudo completar la simulación: {type(exc).__name__}: {exc}')
                return
        esperado = ['kwh_viaje_trac','kwh_viaje_aux','kwh_viaje_regen','kwh_viaje_neto']
        faltantes = [c for c in esperado if c not in res]
        if faltantes:
            st.error('El motor no devolvió: '+', '.join(faltantes))
            return
        st.session_state['sim_energia_resultado'] = res
        st.session_state['sim_energia_meta'] = {
            'fecha': str(fecha), 'energia_real': traccion_real, 'fuente': fuente_real,
            'km_real': km_real, 'descartados': descartes, 'cobertura': cobertura,
            'pax_reales': encontrados, 'parametros': (traccion,pax_estimados,riel,pendientes,regeneracion,headway)
        }
        # Historial de simulaciones solo dentro de la sesión, no se mezclan
        # entre fechas ni se publican como consumo medido.
        neto = float(pd.to_numeric(res['kwh_viaje_neto'], errors='coerce').sum())
        hist = st.session_state.setdefault('sim_energia_historial', {})
        hist[str(fecha)] = {'Fecha': str(fecha), 'Real kWh': traccion_real,
                            'Simulado kWh': neto,
                            'Diferencia kWh': neto-traccion_real if np.isfinite(traccion_real) else np.nan,
                            'Diferencia %': 100*(neto-traccion_real)/traccion_real if np.isfinite(traccion_real) and traccion_real>0 else np.nan,
                            'Viajes': len(res)}
    res = st.session_state.get('sim_energia_resultado')
    meta = st.session_state.get('sim_energia_meta', {})
    if not isinstance(res, pd.DataFrame) or res.empty or meta.get('fecha') != str(fecha):
        return
    if tuple(meta.get('parametros', ())) != (traccion,pax_estimados,riel,pendientes,regeneracion,headway):
        st.info('Cambiaste los parámetros físicos. Pulsa «Simular y comparar» para recalcular '
                'antes de revisar los resultados.')
        return
    st.markdown('#### Resultado de la simulación')
    fields = {'kwh_viaje_trac':'Tracción bruta','kwh_viaje_aux':'Auxiliares del tren',
              'kwh_viaje_regen':'Regeneración','kwh_viaje_neto':'Consumo neto simulado'}
    totals = res[list(fields)].apply(pd.to_numeric, errors='coerce').sum()
    cx = st.columns(4)
    for co, (key, label) in zip(cx, fields.items()):
        co.metric(label, formatear(totals[key], 0)+' kWh')
    sim = float(totals['kwh_viaje_neto'])
    real_kwh = float(meta.get('energia_real', np.nan))
    if np.isfinite(real_kwh) and real_kwh>0:
        diff = sim-real_kwh
        error = diff/real_kwh*100
        c1,c2,c3 = st.columns(3)
        c1.metric('Diferencia simulada − real', formatear(diff, 0)+' kWh')
        c2.metric('Error relativo', formatear(error, 2)+' %')
        c3.metric('IDE simulado (km UMR)', formatear(sim/meta['km_real'],3)+' kWh/km'
                  if np.isfinite(meta['km_real']) and meta['km_real']>0 else '—')
        import plotly.graph_objects as go
        fig = go.Figure(go.Bar(x=['Real tracción','Neto simulado'],y=[real_kwh,sim],
                               marker_color=['#17bed2','#f2b658'],text=[formatear(real_kwh,0),formatear(sim,0)],
                               textposition='outside'))
        fig.update_layout(title='Comparación del mismo día',yaxis_title='Energía (kWh)',height=360,
                          margin=dict(l=20,r=15,t=55,b=35))
        st.plotly_chart(fig,use_container_width=True)
        st.caption('Error = (simulado − real) / real × 100. La diferencia incluye posibles pérdidas eléctricas, '
                   'motores no representados, itinerarios incompletos y supuestos de carga.')
    else:
        st.info('No se presenta porcentaje de error porque no hay tracción real comparable.')
    if res['Via'].nunique() < 2:
        st.warning('Falta al menos una vía. El consumo simulado NO representa el día completo.')
    st.caption(f"Servicios simulados: {len(res)} · vía 1: {meta['cobertura'][1]} · "
               f"vía 2: {meta['cobertura'][2]} · pasajeros reales asociados: "
               f"{meta['pax_reales']} · restantes con ocupación supuesta: "
               f"{len(res)-meta['pax_reales']} · descartados: {len(meta['descartados'])}.")
    for titulo, group in [('Vía', 'Via'), ('Tipo de tren', 'tipo_tren')]:
        st.markdown(f'#### Consumo por {titulo.lower()}')
        g = res.groupby(group,as_index=False)[list(fields)].sum()
        st.dataframe(g.rename(columns=fields),hide_index=True,use_container_width=True)
    st.markdown('#### Detalle estimado por viaje')
    cols = ['Fecha_str','Via','num_servicio','motriz_num','tipo_tren','doble','t_ini',
            'km_orig','km_dest','pax_abordo','tren_km',*fields]
    detalle = res[[c for c in cols if c in res.columns]].copy()
    st.dataframe(detalle, hide_index=True,use_container_width=True,height=360)
    st.download_button('⬇️ Descargar simulación del día (.csv)',
        detalle.to_csv(index=False,sep=';',decimal=',').encode('utf-8-sig'),
        file_name=f'simulacion_vs_real_{fecha}.csv',mime='text/csv',key='sim_energia_csv')
    if len(st.session_state.get('sim_energia_historial',{})) > 1:
        h = pd.DataFrame(st.session_state['sim_energia_historial'].values()).sort_values('Fecha')
        st.markdown('#### Evolución de los días simulados en esta sesión')
        st.line_chart(h.set_index('Fecha')[['Real kWh','Simulado kWh']])
        st.dataframe(h,hide_index=True,use_container_width=True)
        st.download_button('⬇️ Descargar comparación de días (.csv)',
                           h.to_csv(index=False,sep=';',decimal=',').encode('utf-8-sig'),
                           file_name='comparacion_simulacion_dias.csv',mime='text/csv',key='sim_energia_hist_csv')
    if meta['descartados']:
        with st.expander('⚠️ Revisar servicios descartados'):
            st.write(meta['descartados'][:150])
    st.caption(f'Versión del motor físico: {ENGINE_COMMIT[:12]} (Simulador en GitHub). '
               'Los resultados no modifican Datos Operacionales, la línea base ni las alertas.')
