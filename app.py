from pathlib import Path
from html import escape
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import zipfile
from urllib.error import URLError
from scripts.almacenamiento import leer, preparar, guardar, comparar, completar_historial, NUEVAS
from scripts.exportaciones import excel, pdf
from scripts.cmf_online import buscar_reportes, descargar_reporte, INDICE

ROOT = Path(__file__).resolve().parent
st.set_page_config(page_title='Radiografía bancaria | CMF', page_icon='📊', layout='wide')
css=ROOT/'estilos.css'
if css.exists():
    st.markdown('<style>'+css.read_text(encoding='utf-8')+'</style>', unsafe_allow_html=True)
else:
    st.warning('Falta estilos.css. Créalo junto a app.py para ver el nuevo diseño.')
COLORES=['#7944f5','#12b886','#ff9b20','#2f8dff','#ef5c9c','#a1a6ba']
ETIQUETAS={
 'colocaciones_mmclp':'Colocaciones totales (MM CLP)',
 'morosidad_90_pct':'Morosidad ≥90 días (%)',
 'provisiones_clasificacion_pct':'Provisiones por clasificación (%)',
 'eficiencia_gastos_ingresos_pct':'Gastos / ingresos operacionales (%)',
 'roe_pct':'ROE anualizado después de impuestos (%)',
 'roa_pct':'ROA anualizado después de impuestos (%)',
 'utilidad_acumulada_mmclp':'Utilidad consolidada acumulada (MM CLP)',
 'participacion_colocaciones_pct':'Participación en colocaciones (%)',
}
CORTAS=['Colocaciones · MM CLP','Morosidad ≥90 días','Provisiones','Gastos / ingresos',
        'ROE anualizado','ROA anualizado','Utilidad acumulada · MM CLP','Participación de mercado']
PAGINAS=['Resumen general','Comparar bancos','Evolución','Informes y descargas','Datos CMF','Definiciones']

def numero(valor, decimales=2):
    if pd.isna(valor): return 'No disponible'
    return f'{valor:,.{decimales}f}'.translate(str.maketrans({',':'.', '.':','}))

def nombre_mes(fecha):
    nombres=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
    d=pd.Timestamp(fecha)
    return f'{nombres[d.month-1]} {d.year}'

def cambiar_pagina(nombre):
    st.session_state['pagina']=nombre

def grafico(figura, altura=350):
    figura.update_layout(template='plotly_white',height=altura,
        font=dict(family='Arial, sans-serif',color='#666a82',size=12),
        colorway=COLORES,paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='rgba(0,0,0,0)',
        margin=dict(l=10,r=15,t=15,b=15),legend_title_text='')
    figura.update_xaxes(gridcolor='#f0eff6',zeroline=False)
    figura.update_yaxes(gridcolor='#f0eff6',zeroline=False)
    st.plotly_chart(figura,width='stretch',config={'displaylogo':False})

def tarjetas(actual, anterior, prefijo):
    for inicio in (0,4):
        for i,espacio in enumerate(st.columns(4),start=inicio):
            campo=list(ETIQUETAS)[i]
            with espacio,st.container(key=f'kpi_{["violet","green","orange","blue"][i%4]}_{prefijo}_{i}'):
                st.markdown(f'<div class="metric-symbol">{["↗","%","◷","≋","R","A","$","◉"][i]}</div>',unsafe_allow_html=True)
                valor=actual[campo]
                delta=None
                if anterior is not None and pd.notna(valor) and pd.notna(anterior[campo]):
                    base=anterior[campo]
                    if campo.endswith('mmclp'):
                        valida=base>0 if campo=='utilidad_acumulada_mmclp' else base!=0
                        if valida: delta=numero(100*(valor/base-1),3)+' % nominal'
                    else: delta=numero(valor-base,4)+' pp'
                texto=numero(valor,0) if campo.endswith('mmclp') else numero(valor)+' %' if pd.notna(valor) else 'No disponible'
                st.metric(CORTAS[i],texto,delta=delta,delta_color='off',help=ETIQUETAS[campo])

def mostrar_tabla(datos):
    vista=datos[['banco',*ETIQUETAS]].rename(columns={'banco':'Banco',**ETIQUETAS})
    formatos={ETIQUETAS[c]:(lambda x:numero(x,0)) if c.endswith('mmclp') else (lambda x:numero(x)+' %' if pd.notna(x) else 'No disponible') for c in ETIQUETAS}
    st.dataframe(vista.style.format(formatos),hide_index=True,width='stretch')

if 'migracion_v4' not in st.session_state:
    st.session_state['migracion_v4']=completar_historial(ROOT)
tabla=leer(ROOT)
with st.sidebar:
    st.markdown('<div class="brand"><div class="brand-symbol">↗</div><div><div class="brand-name">Radiografía</div><div class="brand-caption">Sistema bancario chileno</div></div></div>',unsafe_allow_html=True)
    if tabla.empty: st.session_state['pagina']='Datos CMF'
    pagina=st.radio('Navegación',PAGINAS,key='pagina',label_visibility='collapsed')
    st.divider()
    st.markdown('### Filtros del análisis')
    if not tabla.empty:
        meses=sorted(tabla.fecha.unique())
        mes=st.selectbox('Mes del reporte',meses,index=len(meses)-1,format_func=nombre_mes,key='mes')
        bancos=sorted(tabla.loc[tabla.tipo_entidad=='banco','banco'].unique())
        seleccion=st.multiselect('Bancos',bancos,default=bancos,key='bancos')
        indicador=st.selectbox('Indicador',list(ETIQUETAS),format_func=ETIQUETAS.get,key='indicador')
        anteriores=[m for m in meses if m<mes]
        referencia=st.selectbox('Comparar con',anteriores,index=len(anteriores)-1,format_func=nombre_mes,key='referencia') if anteriores else None
    st.markdown('<div class="sidebar-note"><strong>Datos públicos, análisis claro.</strong><br>Consulta cifras oficiales de la CMF, compara bancos y descarga tus informes.</div>',unsafe_allow_html=True)
    st.caption('Versión 4.1 · aplicación local')

izquierda,derecha=st.columns([3,1])
with izquierda:
    st.markdown('<div class="eyebrow">Análisis financiero · Chile</div>',unsafe_allow_html=True)
    st.title(pagina)
    st.caption('Radiografía del sistema bancario chileno · Balances consolidados')
with derecha:
    periodo=nombre_mes(mes) if not tabla.empty else 'Sin reportes guardados'
    st.markdown(f'<div class="period">▦ &nbsp; {escape(periodo)}<br><span class="source-pill">FUENTE: CMF</span></div>',unsafe_allow_html=True)

if 'guardado' in st.session_state: st.success(st.session_state.pop('guardado'))
if not tabla.empty and tabla[NUEVAS].isna().any().any():
    st.warning('Faltan ROE, ROA o utilidades en algunos meses. En Datos CMF puedes recuperar los indicadores o cargar los Excel correspondientes.')

if pagina=='Datos CMF':
    st.caption('Agrega reportes oficiales. Puedes cargar archivos Excel o consultar las publicaciones de la CMF.')
    if st.button('Completar indicadores desde archivos guardados'):
        st.session_state['migracion_v4']=completar_historial(ROOT)
        st.rerun()
    with st.expander('Detalle de recuperación de indicadores'):
        for mensaje in st.session_state['migracion_v4']: st.write(mensaje)
    archivos = st.file_uploader('Agregar reportes mensuales CMF', type=['xlsx'], accept_multiple_files=True)
    if archivos:
        try:
            lote = preparar([(a.name, a.getvalue()) for a in archivos])
            st.caption('Meses detectados: ' + ', '.join(nombre_mes(item[2].fecha.iloc[0]) for item in lote))
            reemplazar = st.checkbox('Reemplazar los meses que ya existen', value=False)
            if st.button('Guardar permanentemente', type='primary'):
                cantidad = guardar(ROOT, lote, reemplazar)
                st.session_state['guardado'] = f'{cantidad} reporte(s) guardado(s). Se conservarán al reiniciar.'
                st.rerun()
        except (ValueError, KeyError, OSError, zipfile.BadZipFile) as error:
            st.error(f'No se guardó: {error}')
    st.caption('Los datos permanecen en esta carpeta del proyecto, en este computador. No se sincronizan con otros equipos.')
    with st.expander('Buscar y descargar desde CMF'):
        st.caption('Consulta los reportes publicados. La CMF publica con desfase; no todos los meses calendario están disponibles todavía.')
        if st.button('Buscar reportes disponibles'):
            try:
                with st.spinner('Consultando el catálogo oficial...'):
                    st.session_state['catalogo_cmf']=buscar_reportes()
            except (URLError,ValueError,OSError) as error:
                st.warning(f'No se pudo consultar CMF: {error}. Puedes seguir usando los datos guardados o cargar Excel manualmente.')
        catalogo=st.session_state.get('catalogo_cmf',[])
        if catalogo:
            pendientes=[r for r in catalogo if r['fecha'] not in set(tabla.fecha)]
            st.caption(f'{len(catalogo)} publicaciones reconocidas; {len(pendientes)} meses sin guardar.')
            por_fecha={r['fecha']:r for r in pendientes}
            elegidos=st.multiselect('Meses para descargar',list(por_fecha),format_func=nombre_mes)
            if st.button('Descargar y guardar meses seleccionados',disabled=not elegidos):
                descargas=[]
                try:
                    progreso=st.progress(0)
                    for i,fecha in enumerate(elegidos):
                        nombre,contenido=descargar_reporte(por_fecha[fecha])
                        lote=preparar([(nombre,contenido)])
                        if lote[0][2].fecha.iloc[0] != fecha:
                            raise ValueError('La fecha del Excel no coincide con la publicación seleccionada.')
                        descargas.extend(lote)
                        progreso.progress((i+1)/len(elegidos))
                    cantidad=guardar(ROOT,descargas)
                    st.session_state['guardado']=f'{cantidad} reportes descargados de CMF y guardados.'
                    st.rerun()
                except (URLError,ValueError,KeyError,OSError,zipfile.BadZipFile) as error:
                    st.error(f'La descarga no se completó: {error}. No se incorporó el lote. Usa la carga manual si es necesario.')
        st.link_button('Abrir reportes en la CMF',INDICE)

    if not tabla.empty:
        with st.container(key='panel_archivos'):
            st.subheader('Reportes guardados')
            st.dataframe(tabla[['fecha','archivo_fuente']].drop_duplicates().sort_values('fecha',ascending=False),hide_index=True,width='stretch')
    st.stop()
if tabla.empty:
    st.info('Entra a Datos CMF y guarda un reporte para empezar.')
    st.stop()

sistema=tabla[(tabla.tipo_entidad=='sistema') & (tabla.fecha==mes)]
previo=tabla[(tabla.tipo_entidad=='sistema') & (tabla.fecha==referencia)] if referencia else pd.DataFrame()
comparacion=tabla[(tabla.tipo_entidad=='banco') & (tabla.fecha==mes) & tabla.banco.isin(seleccion)]

if pagina=='Resumen general':
    st.subheader('Una mirada al sistema bancario')
    if not sistema.empty:
        tarjetas(sistema.iloc[0],None if previo.empty else previo.iloc[0],'sistema')
    else: st.info('Este mes no tiene un total oficial del sistema.')
    st.caption('Tarjetas: sistema oficial completo. Montos en MM CLP. ROE y ROA anualizados; utilidad acumulada desde enero. Los filtros de bancos se usan en Comparar bancos y Evolución.')
    if referencia: st.caption(f'Variaciones frente a {nombre_mes(referencia)}. pp = puntos porcentuales. Variación de montos nominal, sin ajuste por inflación.')
    izq,der=st.columns(2)
    with izq,st.container(key='panel_distribucion'):
        st.subheader('Distribución de colocaciones')
        distribucion=tabla[(tabla.tipo_entidad=='banco') & (tabla.fecha==mes)].sort_values('colocaciones_mmclp',ascending=False)
        validos=distribucion.dropna(subset=['colocaciones_mmclp'])
        if not validos.empty and (validos.colocaciones_mmclp>=0).all() and validos.colocaciones_mmclp.sum()>0:
            nombres=list(validos.banco.iloc[:5])
            valores=list(validos.colocaciones_mmclp.iloc[:5])
            if len(validos)>5:
                nombres.append('Otros bancos'); valores.append(validos.colocaciones_mmclp.iloc[5:].sum())
            fig=go.Figure(go.Pie(labels=nombres,values=valores,hole=.73,sort=False,
                marker=dict(colors=COLORES,line=dict(color='white',width=3)),textinfo='none',
                hovertemplate='%{label}<br>%{percent}<extra></extra>'))
            fig.update_layout(legend=dict(orientation='h',y=-.08),annotations=[dict(text='Colocaciones<br>por banco',x=.5,y=.5,showarrow=False,font_size=15)])
            grafico(fig,370)
            st.caption('Todos los bancos del mes. El gráfico agrupa los restantes como Otros bancos.')
        else: st.info('No hay montos válidos para mostrar esta distribución.')
    with der,st.container(key='panel_tendencia'):
        st.subheader('Tendencia del sistema')
        historia=tabla[tabla.tipo_entidad=='sistema'].sort_values('fecha').copy()
        historia['fecha']=pd.to_datetime(historia.fecha)
        st.caption(ETIQUETAS[indicador])
        if len(historia)>1:
            fig=px.line(historia,x='fecha',y=indicador,markers=True,labels={'fecha':'',indicador:ETIQUETAS[indicador]})
            fig.update_traces(line=dict(color=COLORES[0],width=3),marker_size=8)
            fig.update_xaxes(tickmode='array',tickvals=pd.to_datetime(meses),ticktext=[nombre_mes(m) for m in meses])
            grafico(fig,335)
            st.caption('Cada punto es un cierre mensual reportado; no son observaciones diarias.')
        else: st.info('Agrega otro mes en Datos CMF para ver la evolución.')
    with st.container(key='panel_resumen_tabla'):
        st.subheader('Cifras por banco')
        st.caption(f'{nombre_mes(mes)} · {len(comparacion)} bancos seleccionados')
        if not comparacion.empty: mostrar_tabla(comparacion)
        else: st.info('Selecciona al menos un banco en los filtros laterales.')
    st.markdown('<div class="hero-banner"><strong>De los datos al informe.</strong><p>Descarga el historial para Power BI o prepara un informe en Excel y PDF.</p></div>',unsafe_allow_html=True)
    st.button('Ir a informes y descargas',on_click=cambiar_pagina,args=('Informes y descargas',),type='primary')

elif pagina=='Comparar bancos':
    st.caption('Elige bancos e indicador en la barra lateral. El orden del gráfico no es una clasificación general de solvencia.')
    if len(seleccion)==1 and not comparacion.empty:
        banco=seleccion[0]
        st.subheader(banco)
        base=tabla[(tabla.tipo_entidad=='banco') & (tabla.fecha==referencia) & (tabla.banco==banco)]
        tarjetas(comparacion.iloc[0],None if base.empty else base.iloc[0],'banco')
    with st.container(key='panel_comparacion'):
        st.subheader('Comparación por indicador')
        st.caption(ETIQUETAS[indicador])
        if comparacion.empty: st.info('Selecciona al menos un banco con datos en este período.')
        else:
            fig=px.bar(comparacion.sort_values(indicador),x=indicador,y='banco',orientation='h',labels={indicador:ETIQUETAS[indicador],'banco':''},color_discrete_sequence=[COLORES[0]])
            fig.update_traces(marker_cornerradius=5)
            grafico(fig,max(360,len(comparacion)*30))
            mostrar_tabla(comparacion)
            st.download_button('Descargar tabla CSV',comparacion.to_csv(index=False).encode('utf-8-sig'),f'bancos_{mes}.csv','text/csv')
    if referencia and seleccion:
        st.subheader(f'Cambios por banco · {nombre_mes(referencia)} a {nombre_mes(mes)}')
        cambios = comparar(tabla, mes, referencia, indicador, seleccion)
        es_monto = indicador.endswith('mmclp')
        nombres = {'banco':'Banco', 'referencia':nombre_mes(referencia), 'actual':nombre_mes(mes), 'cambio':'Cambio (MM CLP)' if es_monto else 'Cambio (puntos porcentuales)', 'variacion_pct':'Variación nominal (%)'}
        st.caption(ETIQUETAS[indicador])
        vista = cambios.rename(columns=nombres)
        st.dataframe(vista.style.format({nombres[c]:(lambda x: numero(x, 2)) for c in cambios.columns if c != 'banco'}, na_rep='No disponible'), hide_index=True, width='stretch')
        st.download_button('Descargar comparación CSV', cambios.to_csv(index=False).encode('utf-8-sig'), f'comparacion_{referencia}_{mes}.csv', 'text/csv')
        st.caption('Un banco sin dato en uno de los períodos aparece como No disponible. Si la base es cero, no se calcula variación porcentual. Para utilidades tampoco se calcula con base negativa; la diferencia monetaria sí se muestra. Las utilidades comparadas son acumuladas, no resultados mensuales independientes.')

elif pagina=='Evolución':
    with st.container(key='panel_evolucion'):
        st.subheader('Cómo cambian los indicadores')
        st.caption(ETIQUETAS[indicador])
        historico=tabla[(tabla.tipo_entidad=='banco') & tabla.banco.isin(seleccion)].copy()
        if len(meses)<2: st.info('Agrega otros reportes mensuales en Datos CMF para ver la evolución.')
        elif historico.empty: st.info('Selecciona al menos un banco en la barra lateral.')
        else:
            historico['fecha']=pd.to_datetime(historico.fecha)
            fig=px.line(historico.sort_values('fecha'),x='fecha',y=indicador,color='banco',markers=True,
                labels={indicador:ETIQUETAS[indicador],'fecha':'Mes reportado','banco':'Banco'},color_discrete_sequence=COLORES)
            fig.update_xaxes(tickmode='array',tickvals=pd.to_datetime(meses),ticktext=[nombre_mes(m) for m in meses])
            grafico(fig,470)
            st.caption('Cada punto es un dato mensual reportado. Selecciona pocos bancos para facilitar la lectura. Las utilidades son acumuladas desde enero de cada año.')

elif pagina=='Informes y descargas':
    st.subheader('Exportar informe')
    st.caption('Excel incluye el mes, la comparación, el historial de los bancos seleccionados y el sistema oficial por separado. El PDF resume el mes y la comparación con sus definiciones.')
    if seleccion:
        if st.button('Preparar Excel y PDF'):
            with st.spinner('Preparando archivos...'):
                st.session_state['exportaciones']={
                    'clave':(mes,referencia,indicador,tuple(seleccion),tabla.to_csv(index=False)),
                    'excel':excel(tabla,mes,referencia,indicador,seleccion),
                    'pdf':pdf(tabla,mes,referencia,indicador,seleccion),
                }
        resultado=st.session_state.get('exportaciones')
        clave=(mes,referencia,indicador,tuple(seleccion),tabla.to_csv(index=False))
        if resultado and resultado['clave']==clave:
            izquierda,derecha=st.columns(2)
            izquierda.download_button('Descargar Excel',resultado['excel'],f'radiografia_{mes}.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            derecha.download_button('Descargar PDF',resultado['pdf'],f'radiografia_{mes}.pdf','application/pdf')
        elif resultado:
            st.info('Cambiaste los filtros o los datos. Pulsa Preparar Excel y PDF para actualizar los archivos.')
    else:
        st.info('Selecciona al menos un banco para preparar los informes.')
    with st.expander('Datos guardados y descarga del historial', expanded=True):
        st.write(f'{len(meses)} meses guardados. {tabla[tabla.tipo_entidad == "banco"].banco.nunique()} bancos distintos en el historial.')
        st.download_button('Descargar historial CSV', tabla.to_csv(index=False).encode('utf-8-sig'), 'historial_bancario.csv', 'text/csv')
        st.write('Los nuevos originales se conservan en datos/archivados y los indicadores en datos/procesados/bancos.sqlite. Copia toda la carpeta para respaldar el proyecto. El historial descargado incluye las cargas guardadas; los CSV iniciales no se actualizan desde este botón.')

elif pagina=='Definiciones':
    with st.expander('Definiciones y límites de comparación', expanded=True):
        st.write('Colocaciones totales: costo amortizado y valor razonable. Morosidad: indicador oficial de 90 días o más sobre colocaciones totales. Provisiones: índice por clasificación sobre colocaciones a costo amortizado; no es un monto ni incluye automáticamente provisiones adicionales. Eficiencia: gastos / ingresos operacionales del período acumulado reportado; no representa solo el gasto de ese mes.')
        st.write('Los porcentajes conservan la escala 0–100 del reporte. Cero es un dato válido. El total del sistema se extrae de la CMF y no se obtiene promediando porcentajes de bancos. Los modelos de negocio difieren: estos gráficos no constituyen una clasificación general de solvencia.')
        st.write('Banco Bice absorbió Banco Security el 1 de noviembre de 2025, según la nota del reporte; considera esta ruptura al agregar historia.')
        st.write('ROE y ROA: resultado consolidado acumulado después de impuestos, anualizado, dividido por patrimonio o activos de cierre respectivamente. Se extraen los indicadores oficiales, sin recalcularlos. La utilidad/pérdida consolidada incluye el resultado consolidado; no se sustituye por el resultado atribuible a propietarios. Participación: colocaciones totales del banco / total oficial del sistema x 100. No cambia al filtrar bancos. Ningún indicador aislado constituye una clasificación general de calidad bancaria.')

st.markdown('<div class="footer">Radiografía bancaria · Fuente: CMF · Balances consolidados · Versión 4.1</div>',unsafe_allow_html=True)
