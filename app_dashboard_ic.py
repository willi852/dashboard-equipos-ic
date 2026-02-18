"""
Dashboard de Seguimiento - Equipos I&C
======================================
Sistema completo de seguimiento y análisis de avance para proyectos de Instrumentación y Control.

Autor: Dashboard I&C
Fecha: Febrero 2026
Versión: 1.0.8 - Fix: Equipos pendientes Suministro de Aire excluye N/A correctamente

USO:
    streamlit run app_dashboard_ic.py

DEPENDENCIAS:
    pip install streamlit pandas openpyxl plotly xlrd kaleido
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
from io import BytesIO

# ============================================================================
# CONFIGURACIÓN DE LA APLICACIÓN
# ============================================================================

st.set_page_config(
    page_title="Dashboard Equipos I&C",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# URL POR DEFECTO
URL_DEFECTO = "https://drive.google.com/uc?export=download&id=1x_uQhW4EKXiEgbLzZpF_InphP2oIItlu"

# Valores globales reutilizables
VALORES_NA = ['N/A', 'NA', 'n/a', 'na', 'N/a']
VALORES_OK = ['OK', 'SI', 'Completado', 'COMPLETADO', 'ok', 'X', 'x', 1, True]

# ============================================================================
# INICIALIZAR SESSION STATE
# ============================================================================

if 'filtros_inicializados' not in st.session_state:
    st.session_state.filtros_inicializados = False
    st.session_state.filtros = {}

# ============================================================================
# FUNCIONES AUXILIARES
# ============================================================================

def generar_excel_ejemplo():
    """Genera un archivo Excel de ejemplo para pruebas."""
    data = {
        'ITEM': [1, 2, 3, 4, 5, 6, 7, 8],
        'TAG': ['FT-001', 'PT-002', 'TT-003', 'LT-004', 'FV-005', 'PT-006', 'TT-007', 'LT-008'],
        'TIPO INSTRUMENTOS': [
            'Transmisor de Flujo', 'Transmisor de Presión', 'Transmisor de Temperatura',
            'Transmisor de Nivel', 'Válvula de Control', 'Transmisor de Presión',
            'Transmisor de Temperatura', 'Transmisor de Nivel'
        ],
        'AREA': ['Area 100', 'Area 100', 'Area 200', 'Area 200', 'Area 100', 'Area 300', 'Area 300', 'Area 200'],
        'SISTEMA GENERAL': ['Vapor', 'Vapor', 'Agua', 'Agua', 'Vapor', 'Combustible', 'Combustible', 'Agua'],
        'SISTEMA BMS/SMC/DCS': ['DCS', 'DCS', 'PLC', 'PLC', 'DCS', 'DCS', 'PLC', 'PLC'],
        'SISTEMA': ['Sistema A', 'Sistema A', 'Sistema B', 'Sistema B', 'Sistema A', 'Sistema C', 'Sistema C', 'Sistema B'],
        'SIGNAL ASSOCIATION': ['AI-001', 'AI-002', 'AI-003', 'AI-004', 'AO-001', 'AI-005', 'AI-006', 'AI-007'],
        'DESCRIPTION': [
            'Flujo de vapor principal', 'Presión de vapor', 'Temperatura agua',
            'Nivel tanque principal', 'Control flujo vapor', 'Presión combustible',
            'Temperatura combustible', 'Nivel tanque secundario'
        ],
        'SIGNAL': ['4-20mA'] * 8,
        'I/O': ['AI', 'AI', 'AI', 'AI', 'AO', 'AI', 'AI', 'AI'],
        'Hito': ['Hito 1', 'Hito 1', 'Hito 2', 'Hito 2', 'Hito 1', 'Hito 3', 'Hito 3', 'Hito 2'],
        'Prioridad': ['Alta', 'Media', 'Alta', 'Baja', 'Alta', 'Media', 'Baja', 'Alta'],
        'Pre Emsanblado':       ['OK', 'OK', 'Pendiente', 'OK', 'OK', 'OK', 'Pendiente', 'OK'],
        'A Instalar':           ['OK'] * 8,
        'Instalación':          ['OK', 'OK', 'Pendiente', 'OK', 'OK', 'OK', 'Pendiente', 'Pendiente'],
        'Canalización/Bandeja': ['OK', 'OK', 'OK', 'Pendiente', 'OK', 'OK', 'Pendiente', 'OK'],
        'Cableado':             ['OK', 'Pendiente', 'Pendiente', 'Pendiente', 'OK', 'Pendiente', 'Pendiente', 'Pendiente'],
        'Conexión Equipo':      ['OK', 'Pendiente', 'Pendiente', 'Pendiente', 'Pendiente', 'Pendiente', 'Pendiente', 'Pendiente'],
        'Conexión DCS':         ['Pendiente'] * 8,
        'Marquillado Equipo':   ['OK'] * 8,
        'Marquillado Cable':    ['OK', 'Pendiente', 'Pendiente', 'Pendiente', 'OK', 'Pendiente', 'Pendiente', 'Pendiente'],
        'Suiministro de Aire':  ['N/A', 'N/A', 'N/A', 'N/A', 'OK', 'N/A', 'N/A', 'N/A'],
        'Pre-Comisionamiento':  ['Pendiente'] * 8
    }
    df = pd.DataFrame(data)
    try:
        with pd.ExcelWriter('Equipos_IC_Ejemplo.xlsx', engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Equipos I&C', index=False)
        return True, "Equipos_IC_Ejemplo.xlsx"
    except Exception as e:
        return False, str(e)


def crear_excel_descarga(df, nombre_hoja="Datos"):
    """Crea un archivo Excel en memoria para descarga."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=nombre_hoja, index=False)
    buffer.seek(0)
    return buffer


def crear_tabla_imagen(df, titulo="Tabla", max_filas=50):
    """Crea una imagen de tabla usando Plotly."""
    df_display = df.head(max_filas).copy()
    columnas = list(df_display.columns)
    valores = [[str(v)[:50] + '...' if len(str(v)) > 50 else str(v) for v in df_display[col].tolist()] for col in columnas]

    fig = go.Figure(data=[go.Table(
        header=dict(
            values=[f'<b>{col}</b>' for col in columnas],
            fill_color='#1f77b4',
            font=dict(color='white', size=11, family='Arial'),
            align='center',
            height=30
        ),
        cells=dict(
            values=valores,
            fill_color=[['#f0f0f0', 'white'] * len(df_display)],
            font=dict(color='black', size=10, family='Arial'),
            align='left',
            height=25
        )
    )])

    altura = min(800 + (len(df_display) * 25), 4000)
    ancho  = max(1200, len(columnas) * 150)

    fig.update_layout(
        title=dict(
            text=f'<b>{titulo}</b><br><sub>{len(df)} equipos totales - Mostrando primeros {len(df_display)}</sub>',
            x=0.5, xanchor='center',
            font=dict(size=16, color='#1f77b4')
        ),
        height=altura, width=ancho,
        margin=dict(l=20, r=20, t=80, b=20)
    )
    return fig


# ============================================================================
# FUNCIONES PRINCIPALES
# ============================================================================

@st.cache_data(ttl=300)
def cargar_datos(url_excel):
    """Carga datos desde archivo Excel en la nube o local."""
    try:
        df = pd.read_excel(url_excel, sheet_name="Equipos I&C")
        if 'ITEM' in df.columns:
            df = df[df['ITEM'].notna()].copy()
            df = df[df['ITEM'].astype(str).str.strip() != ''].copy()
        return df
    except Exception as e:
        st.error(f"Error al cargar el archivo: {str(e)}")
        return None


def calcular_completados(df, actividad):
    """
    Calcula completados, pendientes, porcentaje y total aplicable.
    Para Suministro de Aire excluye celdas con N/A del conteo.

    Returns:
        tuple: (completados, pendientes, porcentaje, total_aplicable)
    """
    if 'suministro' in actividad.lower() or 'suiministro' in actividad.lower():
        # Solo equipos que SÍ aplican (sin N/A y sin vacíos)
        df_aplicable = df[
            ~df[actividad].isin(VALORES_NA) & df[actividad].notna()
        ].copy()
        total = len(df_aplicable)
        if total == 0:
            return 0, 0, 0, 0
        completados = df_aplicable[actividad].isin(VALORES_OK).sum()
        pendientes  = total - completados
        porcentaje  = (completados / total) * 100
        return completados, pendientes, porcentaje, total

    # Resto de actividades
    total = len(df)
    if total == 0:
        return 0, 0, 0, 0
    completados = df[actividad].isin(VALORES_OK).sum()
    pendientes  = total - completados
    porcentaje  = (completados / total) * 100
    return completados, pendientes, porcentaje, total


def es_suministro(nombre):
    return 'suministro' in nombre.lower() or 'suiministro' in nombre.lower()


# ============================================================================
# INTERFAZ PRINCIPAL
# ============================================================================

st.title("🏭 Dashboard de Seguimiento - Equipos I&C")
st.markdown("---")

# ============================================================================
# SIDEBAR - CONFIGURACIÓN
# ============================================================================

with st.sidebar:
    st.header("⚙️ Configuración")

    tab_url, tab_archivo, tab_ejemplo = st.tabs(["🌐 URL", "📁 Archivo", "🧪 Ejemplo"])

    with tab_url:
        url_excel = st.text_input(
            "URL del archivo Excel:",
            value=URL_DEFECTO,
            help="URL directa de descarga desde Google Drive"
        )
        usar_url = st.checkbox("Usar URL", value=True)
        if url_excel == URL_DEFECTO:
            st.success("✅ URL del proyecto cargada")

    with tab_archivo:
        archivo_subido = st.file_uploader(
            "Sube tu archivo Excel:",
            type=['xlsx', 'xls'],
            help="Selecciona el archivo Excel desde tu computadora"
        )

    with tab_ejemplo:
        st.info("💡 Genera un archivo Excel de ejemplo para probar la aplicación")
        if st.button("📊 Generar Archivo de Ejemplo", type="secondary"):
            exito, mensaje = generar_excel_ejemplo()
            if exito:
                st.success(f"✅ Archivo generado: {mensaje}")
            else:
                st.error(f"❌ Error: {mensaje}")

    if st.button("🔄 Cargar/Actualizar Datos", type="primary"):
        st.cache_data.clear()
        st.success("✅ Datos actualizados")

    st.markdown("---")

# ============================================================================
# CARGA DE DATOS
# ============================================================================

df = None

if archivo_subido is not None:
    df = cargar_datos(archivo_subido)
    fuente_datos = f"📁 Archivo: {archivo_subido.name}"
elif usar_url and url_excel:
    df = cargar_datos(url_excel)
    fuente_datos = "🌐 Google Drive" if "drive.google.com" in url_excel else "🌐 URL en la nube"
else:
    st.info("👈 Por favor, selecciona una fuente de datos en el panel lateral")

# ============================================================================
# PROCESAMIENTO Y VISUALIZACIÓN
# ============================================================================

if df is not None:

    actividades = [
        'A Instalar',
        'Instalación',
        'Canalización/Bandeja',
        'Cableado',
        'Conexión Equipo',
        'Conexión DCS',
        'Marquillado Equipo',
        'Marquillado Cable',
        'Suiministro de Aire',
        'Pre-Comisionamiento'
    ]
    actividades_existentes = [col for col in actividades if col in df.columns]

    # ========================================================================
    # SIDEBAR - FILTROS
    # ========================================================================

    with st.sidebar:
        st.header("🔍 Filtros")
        filtros_activos = {}

        if not st.session_state.filtros_inicializados:
            st.session_state.filtros = {}
            st.session_state.filtros_inicializados = True

        def multiselect_filtro(columna, label, opcion_todos):
            if columna in df.columns:
                opciones = [opcion_todos] + sorted(df[columna].dropna().unique().tolist())
                default  = st.session_state.filtros.get(columna, [opcion_todos])
                sel      = st.multiselect(label, opciones, default=default, key=f'filtro_{columna}')
                st.session_state.filtros[columna] = sel
                if opcion_todos not in sel:
                    filtros_activos[columna] = sel

        multiselect_filtro('AREA',              "Área:",               'Todas')
        multiselect_filtro('SISTEMA GENERAL',   "Sistema General:",    'Todos')
        multiselect_filtro('SISTEMA BMS/SMC/DCS',"Sistema BMS/SMC/DCS:",'Todos')
        multiselect_filtro('TIPO INSTRUMENTOS', "Tipo Instrumento:",   'Todos')
        multiselect_filtro('Prioridad',         "Prioridad:",          'Todas')
        multiselect_filtro('Hito',              "Hito:",               'Todos')

        if st.button("🔄 Resetear Filtros", type="secondary"):
            st.session_state.filtros = {}
            st.session_state.filtros_inicializados = False
            st.rerun()

    # Aplicar filtros
    df_filtrado = df.copy()
    for columna, valores in filtros_activos.items():
        df_filtrado = df_filtrado[df_filtrado[columna].isin(valores)]

    # ========================================================================
    # INFORMACIÓN GENERAL
    # ========================================================================

    col1, col2, col3, col4 = st.columns(4)
    col1.info(f"📋 **Total Equipos:** {len(df)}")
    col2.info(f"🔍 **Equipos Filtrados:** {len(df_filtrado)}")
    col3.success(f"✅ **Filtros Activos:** {len(filtros_activos)}")
    col4.info(f"📊 **Fuente:** {fuente_datos}")

    st.markdown("---")

    # ========================================================================
    # MÉTRICAS DE AVANCE
    # ========================================================================

    st.header("📊 Métricas de Avance General")

    # Panel especial para Suministro de Aire
    if 'Suiministro de Aire' in actividades_existentes:
        total_global  = len(df_filtrado)
        comp_sa, pend_sa, pct_sa, aplican_sa = calcular_completados(df_filtrado, 'Suiministro de Aire')
        na_sa = total_global - aplican_sa

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.info(f"**🔧 Suministro de Aire**

📊 {aplican_sa} equipos lo requieren")
        with c2:
            st.success(f"**✅ Completados**

{comp_sa} de {aplican_sa} → {pct_sa:.1f}%")
        with c3:
            st.warning(f"**⚠️ Pendientes**

{pend_sa} equipos")
        with c4:
            st.info(f"**⚪ No Aplica (N/A)**

{na_sa} equipos")

        st.markdown("---")

    metricas_cols = st.columns(5)

    for idx, actividad in enumerate(actividades_existentes):
        with metricas_cols[idx % 5]:
            total = len(df_filtrado)
            if total > 0:
                completados, pendientes, porcentaje, total_aplicable = calcular_completados(df_filtrado, actividad)
                etiqueta = actividad.replace('Suiministro', 'Suministro')

                if es_suministro(actividad):
                    na_count = total - total_aplicable
                    st.metric(label=etiqueta, value=f"{completados}/{total_aplicable}", delta=f"{porcentaje:.1f}%")
                    st.caption(f"✅ {total_aplicable} aplican | ⚪ {na_count} N/A")
                else:
                    st.metric(label=etiqueta, value=f"{completados}/{total}", delta=f"{porcentaje:.1f}%")
            else:
                st.metric(label=actividad, value="0/0", delta="0%")

    st.markdown("---")

    # ========================================================================
    # GRÁFICOS DE PROGRESO
    # ========================================================================

    st.header("📈 Progreso por Actividad")

    avance_data = []
    for actividad in actividades_existentes:
        if len(df_filtrado) > 0:
            completados, pendientes, porcentaje, _ = calcular_completados(df_filtrado, actividad)
            avance_data.append({
                'Actividad': actividad.replace('Suiministro', 'Suministro'),
                'Completados': completados,
                'Pendientes': pendientes,
                'Porcentaje': porcentaje
            })

    df_avance = pd.DataFrame(avance_data)

    col_g1, col_g2 = st.columns(2)

    with col_g1:
        fig_barras = go.Figure()
        fig_barras.add_trace(go.Bar(
            name='Completados', x=df_avance['Actividad'], y=df_avance['Completados'],
            marker_color='#10b981', text=df_avance['Completados'],
            textposition='inside', textfont=dict(size=12, color='white', family='Arial Black')
        ))
        fig_barras.add_trace(go.Bar(
            name='Pendientes', x=df_avance['Actividad'], y=df_avance['Pendientes'],
            marker_color='#ef4444', text=df_avance['Pendientes'],
            textposition='inside', textfont=dict(size=12, color='white', family='Arial Black')
        ))
        fig_barras.update_layout(
            barmode='stack', title='Estado de Actividades',
            xaxis_title='Actividad', yaxis_title='Cantidad',
            xaxis_tickangle=-45, height=400, showlegend=True
        )
        st.plotly_chart(fig_barras, use_container_width=True)

    with col_g2:
        fig_pct = go.Figure()
        fig_pct.add_trace(go.Bar(
            x=df_avance['Actividad'], y=df_avance['Porcentaje'],
            marker=dict(color=df_avance['Porcentaje'], colorscale='RdYlGn', cmin=0, cmax=100),
            text=[f"{p:.1f}%" for p in df_avance['Porcentaje']],
            textposition='outside', textfont=dict(size=11, color='black', family='Arial Black')
        ))
        fig_pct.update_layout(
            title='Porcentaje de Completitud por Actividad',
            xaxis_title='Actividad', yaxis_title='% Completado',
            xaxis_tickangle=-45, height=400,
            yaxis=dict(range=[0, max(df_avance['Porcentaje'].max() * 1.1, 10)])
        )
        st.plotly_chart(fig_pct, use_container_width=True)

    col_b1, col_b2, _ = st.columns([1, 1, 2])
    with col_b1:
        try:
            st.download_button(
                label="📸 Descargar Gráfico Barras (PNG)",
                data=fig_barras.to_image(format="png", width=1200, height=600),
                file_name=f"grafico_barras_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                mime="image/png"
            )
        except:
            st.info("ℹ️ PNG requiere configuración adicional")

    with col_b2:
        try:
            st.download_button(
                label="📸 Descargar Gráfico % (PNG)",
                data=fig_pct.to_image(format="png", width=1200, height=600),
                file_name=f"grafico_pct_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                mime="image/png"
            )
        except:
            st.info("ℹ️ PNG requiere configuración adicional")

    st.markdown("---")

    # ========================================================================
    # EQUIPOS PENDIENTES
    # ========================================================================

    st.header("⚠️ Equipos Pendientes por Actividad")

    actividades_seleccionadas = st.multiselect(
        "Selecciona una o más actividades para ver equipos pendientes:",
        actividades_existentes,
        default=[actividades_existentes[0]] if actividades_existentes else []
    )

    if actividades_seleccionadas:
        df_pend_work = df_filtrado.copy()

        for actividad in actividades_seleccionadas:
            if es_suministro(actividad):
                # ✅ CORRECTO:
                # Pendiente = (SÍ aplica: no es N/A y no está vacío)
                #             AND (NO está completado)
                mask_aplica     = ~df_pend_work[actividad].isin(VALORES_NA) & df_pend_work[actividad].notna()
                mask_incompleto = ~df_pend_work[actividad].isin(VALORES_OK)
                df_pend_work[f'{actividad}_Pendiente'] = mask_aplica & mask_incompleto
            else:
                # Otras actividades: pendiente = no tiene OK
                df_pend_work[f'{actividad}_Pendiente'] = ~df_pend_work[actividad].isin(VALORES_OK)

        condiciones  = [df_pend_work[f'{a}_Pendiente'] for a in actividades_seleccionadas]
        df_pendientes = df_pend_work[pd.concat(condiciones, axis=1).any(axis=1)]

        c_p1, c_p2, _ = st.columns([1, 1, 2])
        with c_p1:
            pct_pend = (len(df_pendientes) / len(df_filtrado) * 100) if len(df_filtrado) > 0 else 0
            st.metric("Equipos con Pendientes", len(df_pendientes), delta=f"{pct_pend:.1f}% del total")
        with c_p2:
            st.metric("Actividades Seleccionadas", len(actividades_seleccionadas))

        if len(df_pendientes) > 0:
            st.subheader(f"Listado de Equipos Pendientes ({len(actividades_seleccionadas)} actividad{'es' if len(actividades_seleccionadas) > 1 else ''})")

            cols_base    = ['ITEM', 'TAG', 'DESCRIPTION', 'AREA', 'SISTEMA GENERAL', 'TIPO INSTRUMENTOS', 'Prioridad']
            cols_base    = [c for c in cols_base if c in df_pendientes.columns]
            cols_mostrar = cols_base + actividades_seleccionadas

            st.dataframe(df_pendientes[cols_mostrar], use_container_width=True, height=400)
            st.info(f"💡 Muestra equipos con al menos una de las {len(actividades_seleccionadas)} actividades pendientes.")

            nombre_act = "_".join([a.replace('/', '-').replace(' ', '_') for a in actividades_seleccionadas[:3]])
            if len(actividades_seleccionadas) > 3:
                nombre_act += f"_y_{len(actividades_seleccionadas)-3}_mas"

            cd1, cd2, cd3 = st.columns(3)

            with cd1:
                st.download_button(
                    label="📥 Descargar Tabla (Excel)",
                    data=crear_excel_descarga(df_pendientes[cols_mostrar], "Pendientes"),
                    file_name=f"pendientes_{nombre_act}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            with cd2:
                try:
                    titulo_t = f"Equipos Pendientes - {', '.join(actividades_seleccionadas[:2])}"
                    if len(actividades_seleccionadas) > 2:
                        titulo_t += f" y {len(actividades_seleccionadas)-2} más"
                    fig_tabla = crear_tabla_imagen(df_pendientes[cols_mostrar], titulo=titulo_t, max_filas=50)
                    st.download_button(
                        label="📸 Descargar Tabla (PNG)",
                        data=fig_tabla.to_image(format="png", width=fig_tabla.layout.width, height=fig_tabla.layout.height),
                        file_name=f"tabla_pendientes_{nombre_act}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                        mime="image/png"
                    )
                except:
                    st.info("ℹ️ PNG requiere configuración adicional")

            with cd3:
                if len(df_pendientes) > 50:
                    st.warning(f"⚠️ PNG muestra solo las primeras 50 filas de {len(df_pendientes)}. Usa Excel para ver todos.")

            with st.expander("📊 Ver Resumen por Actividad"):
                resumen = []
                for actividad in actividades_seleccionadas:
                    pend_act = df_pendientes[df_pendientes[f'{actividad}_Pendiente']].shape[0]
                    resumen.append({
                        'Actividad': actividad.replace('Suiministro', 'Suministro'),
                        'Pendientes': pend_act,
                        'Porcentaje del total': f"{(pend_act / len(df_filtrado) * 100):.1f}%" if len(df_filtrado) > 0 else "0%"
                    })
                df_resumen = pd.DataFrame(resumen)
                st.dataframe(df_resumen, use_container_width=True)
                st.download_button(
                    label="📥 Descargar Resumen (Excel)",
                    data=crear_excel_descarga(df_resumen, "Resumen"),
                    file_name=f"resumen_pendientes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.success("✅ ¡No hay equipos pendientes para las actividades seleccionadas!")
    else:
        st.info("👆 Selecciona al menos una actividad para ver los equipos pendientes")

    st.markdown("---")

    # ========================================================================
    # ANÁLISIS MULTIDIMENSIONAL
    # ========================================================================

    st.header("🔍 Análisis Multidimensional")

    tab1, tab2, tab3, tab4 = st.tabs(["📍 Por Área", "⚙️ Por Sistema", "🔧 Por Tipo", "🎯 Por Prioridad"])

    with tab1:
        if 'AREA' in df_filtrado.columns:
            analisis = df_filtrado.groupby('AREA').size().reset_index(name='Cantidad')
            ca1, ca2 = st.columns([2, 1])
            with ca1:
                fig_area = px.pie(analisis, values='Cantidad', names='AREA',
                                  title='Distribución de Equipos por Área', hole=0.3)
                fig_area.update_traces(textposition='inside', textinfo='percent+label+value')
                st.plotly_chart(fig_area, use_container_width=True)
                try:
                    st.download_button("📸 Descargar (PNG)",
                        data=fig_area.to_image(format="png", width=1200, height=800),
                        file_name=f"area_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png", mime="image/png")
                except:
                    st.info("ℹ️ PNG requiere configuración adicional")
            with ca2:
                st.dataframe(analisis.sort_values('Cantidad', ascending=False), use_container_width=True, height=400)
                st.download_button("📥 Descargar (Excel)",
                    data=crear_excel_descarga(analisis, "Por Área"),
                    file_name=f"area_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with tab2:
        if 'SISTEMA GENERAL' in df_filtrado.columns:
            analisis = df_filtrado.groupby('SISTEMA GENERAL').size().reset_index(name='Cantidad')
            fig_sis = px.bar(analisis.sort_values('Cantidad', ascending=True),
                             x='Cantidad', y='SISTEMA GENERAL', title='Equipos por Sistema General',
                             orientation='h', color='Cantidad', color_continuous_scale='Blues', text='Cantidad')
            fig_sis.update_traces(textposition='outside')
            st.plotly_chart(fig_sis, use_container_width=True)
            cs1, cs2 = st.columns(2)
            with cs1:
                try:
                    st.download_button("📸 Descargar (PNG)",
                        data=fig_sis.to_image(format="png", width=1200, height=800),
                        file_name=f"sistema_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png", mime="image/png")
                except:
                    st.info("ℹ️ PNG requiere configuración adicional")
            with cs2:
                st.download_button("📥 Descargar (Excel)",
                    data=crear_excel_descarga(analisis, "Por Sistema"),
                    file_name=f"sistema_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.dataframe(analisis.sort_values('Cantidad', ascending=False), use_container_width=True)

    with tab3:
        if 'TIPO INSTRUMENTOS' in df_filtrado.columns:
            analisis = df_filtrado.groupby('TIPO INSTRUMENTOS').size().reset_index(name='Cantidad')
            fig_tipo = px.bar(analisis.sort_values('Cantidad', ascending=False),
                              x='TIPO INSTRUMENTOS', y='Cantidad', title='Equipos por Tipo de Instrumento',
                              color='Cantidad', color_continuous_scale='Viridis', text='Cantidad')
            fig_tipo.update_layout(xaxis_tickangle=-45)
            fig_tipo.update_traces(textposition='outside')
            st.plotly_chart(fig_tipo, use_container_width=True)
            ct1, ct2 = st.columns(2)
            with ct1:
                try:
                    st.download_button("📸 Descargar (PNG)",
                        data=fig_tipo.to_image(format="png", width=1200, height=800),
                        file_name=f"tipo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png", mime="image/png")
                except:
                    st.info("ℹ️ PNG requiere configuración adicional")
            with ct2:
                st.download_button("📥 Descargar (Excel)",
                    data=crear_excel_descarga(analisis, "Por Tipo"),
                    file_name=f"tipo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.dataframe(analisis.sort_values('Cantidad', ascending=False), use_container_width=True)

    with tab4:
        if 'Prioridad' in df_filtrado.columns:
            analisis = df_filtrado.groupby('Prioridad').size().reset_index(name='Cantidad')
            cp1, cp2 = st.columns([2, 1])
            with cp1:
                fig_pri = px.pie(analisis, values='Cantidad', names='Prioridad',
                                 title='Distribución por Prioridad', color='Prioridad',
                                 color_discrete_map={'Alta': '#ef4444', 'Media': '#f59e0b', 'Baja': '#10b981'})
                fig_pri.update_traces(textposition='inside', textinfo='percent+label+value')
                st.plotly_chart(fig_pri, use_container_width=True)
                try:
                    st.download_button("📸 Descargar (PNG)",
                        data=fig_pri.to_image(format="png", width=1200, height=800),
                        file_name=f"prioridad_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png", mime="image/png")
                except:
                    st.info("ℹ️ PNG requiere configuración adicional")
            with cp2:
                st.dataframe(analisis.sort_values('Cantidad', ascending=False), use_container_width=True, height=400)
                st.download_button("📥 Descargar (Excel)",
                    data=crear_excel_descarga(analisis, "Por Prioridad"),
                    file_name=f"prioridad_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.markdown("---")

    # ========================================================================
    # TABLA COMPLETA
    # ========================================================================

    st.header("📋 Tabla Completa de Equipos")

    co1, co2, co3 = st.columns([1, 2, 1])
    with co1:
        mostrar_todas = st.checkbox("Mostrar todas las columnas", value=False)
    with co2:
        buscar_tag = st.text_input("🔍 Buscar por TAG:", placeholder="Ingresa TAG...")
    with co3:
        st.metric("Total Mostrados", len(df_filtrado))

    df_mostrar = df_filtrado.copy()
    if buscar_tag and 'TAG' in df_mostrar.columns:
        df_mostrar = df_mostrar[df_mostrar['TAG'].str.contains(buscar_tag, case=False, na=False)]

    if not mostrar_todas:
        cols_def = ['ITEM', 'TAG', 'TIPO INSTRUMENTOS', 'AREA', 'SISTEMA GENERAL', 'DESCRIPTION', 'Prioridad'] + actividades_existentes
        cols_def  = [c for c in cols_def if c in df_mostrar.columns]
        df_mostrar = df_mostrar[cols_def]

    st.dataframe(df_mostrar, use_container_width=True, height=500)

    st.download_button(
        label="📥 Descargar Tabla Filtrada (Excel)",
        data=crear_excel_descarga(df_mostrar, "Equipos Filtrados"),
        file_name=f"equipos_ic_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # ========================================================================
    # PIE DE PÁGINA
    # ========================================================================

    st.markdown("---")
    pf1, pf2, pf3 = st.columns(3)
    pf1.info(f"📊 **Actualización:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    pf2.success("✅ **Datos cargados correctamente**")
    pf3.info("📈 **Versión:** 1.0.8")
