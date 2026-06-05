# ... existing code ...
            # --- NUEVA ESTRUCTURA: SERVICIOS Y PAX LADO A LADO ---
            # Se aumentó ligeramente la proporción del gráfico para evitar compresión de títulos
            c_chart_s, c_card_s, c_chart_p, c_card_p = st.columns([2.5, 1, 2.5, 1]) 
            
            with c_chart_s:
                fig_serv = px.bar(df_resumen, x='Fecha', y='Servicios', 
                                  color_discrete_sequence=["#005195"],
                                  hover_data=hover_config, title="Servicios Programados")
                
                # Texto blanco, auto-rotación, y escudo anti-solapamiento
                fig_serv.update_traces(texttemplate='%{y:,.0f}', textposition='inside', insidetextanchor='middle', textfont=dict(color='white', size=13))
                fig_serv.update_layout(margin=dict(t=50, b=0, l=0, r=0), title=dict(font=dict(size=15), automargin=True),
                                       bargap=0.15, uniformtext=dict(minsize=9, mode='hide'))
                st.plotly_chart(fig_serv, use_container_width=True, config={'locale': 'es'})
                
            with c_card_s:
                st.markdown("<br><br>", unsafe_allow_html=True)
                # Formateo sin decimales para valores discretos, con separador de miles completo
                st.metric("Total Servicios", f"{int(df_resumen['Servicios'].sum()):,}")

            with c_chart_p:
                fig_pax = px.bar(df_resumen, x='Fecha', y='PAX', 
                                  color_discrete_sequence=["#E85500"], # Naranja para contrastar Demanda vs Oferta
                                  hover_data=hover_config, title="Pasajeros Transportados (PAX)")
                
                fig_pax.update_traces(texttemplate='%{y:,.0f}', textposition='inside', insidetextanchor='middle', textfont=dict(color='white', size=13))
                fig_pax.update_layout(margin=dict(t=50, b=0, l=0, r=0), title=dict(font=dict(size=15), automargin=True),
                                      bargap=0.15, uniformtext=dict(minsize=9, mode='hide'))
                st.plotly_chart(fig_pax, use_container_width=True, config={'locale': 'es'})
                
            with c_card_p:
                st.markdown("<br><br>", unsafe_allow_html=True)
                # Formateo sin decimales para valores discretos, con separador de miles completo
                st.metric("Total PAX", f"{int(df_resumen['PAX'].sum()):,}")

            st.divider() # Línea separadora para mantener orden visual
            
            # --- SEGUNDA ESTRUCTURA: KILOMETRAJE Y RENDIMIENTO (UMR) ---
            # Se eliminó el subtítulo a petición para mantener todo bajo DATOS OPERACIONALES
            c_chart_k, c_card_k, c_chart_u, c_card_u = st.columns([2.5, 1, 2.5, 1]) 
            
            with c_chart_k:
                # Gráfico Agrupado (Barmode='group') para comparar Odómetro vs Tren-Km
                fig_km = px.bar(df_resumen, x='Fecha', y=['Odómetro [km]', 'Tren-Km [km]'], 
                                barmode='group',
                                color_discrete_map={'Odómetro [km]': '#005195', 'Tren-Km [km]': '#66A5D9'}, # Tonos de azul
                                hover_data=hover_config, title="Kilometraje (Odómetro vs Tren-Km)")
                
                # Al ser barras dobles, aquí SÍ forzamos el vertical (-90) para que quepan los 2 decimales
                fig_km.update_traces(texttemplate='%{y:,.2f}', textposition='inside', insidetextanchor='middle', textangle=-90, textfont=dict(color='white', size=11))
                fig_km.update_layout(margin=dict(t=50, b=0, l=0, r=0), title=dict(font=dict(size=15), automargin=True),
                                     legend=dict(title="", orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                                     bargap=0.15, uniformtext=dict(minsize=8, mode='hide'))
                st.plotly_chart(fig_km, use_container_width=True, config={'locale': 'es'})
                
            with c_card_k:
                # Se apilan dos tarjetas para coincidir con las dos barras del gráfico
                st.markdown("<br>", unsafe_allow_html=True)
                # Formateo estandarizado a exactamente 2 decimales para variables continuas
                st.metric("Odómetro Total", f"{df_resumen['Odómetro [km]'].sum():,.2f} km")
                st.metric("Tren-Km Total", f"{df_resumen['Tren-Km [km]'].sum():,.2f} km")

            with c_chart_u:
                # Gráfico del Porcentaje UMR
                fig_umr = px.bar(df_resumen, x='Fecha', y='UMR (%)', 
                                  color_discrete_sequence=["#E85500"], 
                                  hover_data=hover_config, title="Tasa Acoplamiento (UMR %)")
                
                fig_umr.update_traces(texttemplate='%{y:,.2f}%', textposition='inside', insidetextanchor='middle', textfont=dict(color='white', size=13))
                fig_umr.update_layout(margin=dict(t=50, b=0, l=0, r=0), title=dict(font=dict(size=15), automargin=True),
                                      bargap=0.15, uniformtext=dict(minsize=9, mode='hide'))
                st.plotly_chart(fig_umr, use_container_width=True, config={'locale': 'es'})
                
            with c_card_u:
                st.markdown("<br><br>", unsafe_allow_html=True)
                # Programación defensiva y corrección matemática: Promedio ponderado real
                tot_tren_km = df_resumen['Tren-Km [km]'].sum()
                tot_odometro = df_resumen['Odómetro [km]'].sum()
                umr_global = (tot_tren_km / tot_odometro * 100) if tot_odometro > 0 else 0
                
                # Formateo estandarizado a exactamente 2 decimales
                st.metric("Tasa UMR Global", f"{umr_global:,.2f} %")
                
            st.divider() # Línea separadora
            
            # --- TERCERA ESTRUCTURA: ENERGÍA E IDE ---
            c_chart_e, c_card_e, c_chart_i, c_card_i = st.columns([2.5, 1, 2.5, 1]) 
            
            # Renombramos las columnas temporalmente para que la leyenda en Plotly salga perfecta en Español
            df_plot_ener = df_resumen.rename(columns={'E_Tr': 'Tracción', 'E_12': 'Baja Tensión'})
            
            with c_chart_e:
                fig_ener = px.bar(df_plot_ener, x='Fecha', y=['Tracción', 'Baja Tensión'], 
                                  barmode='stack',
                                  color_discrete_map={'Tracción': '#E85500', 'Baja Tensión': '#005195'},
                                  hover_data=hover_config, title="Consumo Energético (kWh)")
                
                # Gráfico apilado: forzamos vertical y tamaño mediano para que quepa en los bloques de energía
                fig_ener.update_traces(texttemplate='%{y:,.2f}', textposition='inside', insidetextanchor='middle', textangle=-90, textfont=dict(color='white', size=11)) 
                fig_ener.update_layout(margin=dict(t=50, b=0, l=0, r=0), title=dict(font=dict(size=15), automargin=True),
                                     legend=dict(title="", orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                                     bargap=0.15, uniformtext=dict(minsize=8, mode='hide'))
                st.plotly_chart(fig_ener, use_container_width=True, config={'locale': 'es'})
                
            with c_card_e:
                st.markdown("<br>", unsafe_allow_html=True)
                # Formateo estandarizado a exactamente 2 decimales para energía total
                st.metric("Total Tracción", f"{df_plot_ener['Tracción'].sum():,.2f} kWh")
                st.metric("Total Baja Tensión", f"{df_plot_ener['Baja Tensión'].sum():,.2f} kWh")

            with c_chart_i:
                fig_ide_bar = px.bar(df_resumen, x='Fecha', y='IDE (kWh/km)', 
                                  color_discrete_sequence=["#E85500"], 
                                  hover_data=hover_config, title="Desempeño Energético (IDE)")
                
                fig_ide_bar.update_traces(texttemplate='%{y:,.2f}', textposition='inside', insidetextanchor='middle', textfont=dict(color='white', size=13))
                fig_ide_bar.update_layout(margin=dict(t=50, b=0, l=0, r=0), title=dict(font=dict(size=15), automargin=True),
                                          bargap=0.15, uniformtext=dict(minsize=9, mode='hide'))
                st.plotly_chart(fig_ide_bar, use_container_width=True, config={'locale': 'es'})
                
            with c_card_i:
                st.markdown("<br><br>", unsafe_allow_html=True)
                # Cálculo matemáticamente correcto del IDE Global
                tot_traccion_real = df_resumen['E_Tr'].sum()
                ide_global = (tot_traccion_real / tot_odometro) if tot_odometro > 0 else 0
                
                # Estandarizado a exactamente 2 decimales para la tarjeta, respetando tu solicitud
                st.metric("IDE Global", f"{ide_global:,.2f} kWh/km")
            
    else: st.info("📂 Sube archivos desde el panel lateral para ver el resumen.")
# ... existing code ...
