"""
Dashboard de Seguimiento - Equipos I&C
======================================
Sistema completo de seguimiento y análisis de avance para proyectos de Instrumentación y Control.

Autor: Dashboard I&C
Fecha: Febrero 2026
Versión: 1.4.2 - Fix: conteo correcto SA/Tubing (incluye equipos pendientes)

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

# URL POR DEFECTO (Google Sheets → URL de descarga)
URL_DEFECTO = "https://drive.google.com/uc?export=download&id=1x_uQhW4EKXiEgbLzZpF_InphP2oIItlu"

# ============================================================================
# INICIALIZAR SESSION STATE PARA FILTROS PERSISTENTES
# ============================================================================

if 'filtros_inicializados' not in st.session_state:
    st.session_state.filtros_inicializados = False
    st.session_state.filtros = {}

# ============================================================================
# FUNCIONES AUXILIARES
# ============================================================================

def verificar_dependencias():
    """Verifica que todas las dependencias estén instaladas."""
    dependencias = {
        'streamlit': 'streamlit',
        'pandas': 'pandas',
        'plotly': 'plotly',
        'openpyxl': 'openpyxl'
    }
    
    faltantes = []
    for modulo, paquete in dependencias.items():
        try:
            __import__(modulo)
        except ImportError:
            faltantes.append(paquete)
    
    if faltantes:
        st.error(f"❌ Faltan dependencias: {', '.join(faltantes)}")
        st.info(f"💡 Instala con: pip install {' '.join(faltantes)}")
        return False
    return True

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
        'Pre Emsanblado': ['OK', 'OK', 'Pendiente', 'OK', 'OK', 'OK', 'Pendiente', 'OK'],
        'A Instalar': ['OK'] * 8,
        'Instalación': ['OK', 'OK', 'Pendiente', 'OK', 'OK', 'OK', 'Pendiente', 'Pendiente'],
        'Canalización/Bandeja': ['OK', 'OK', 'OK', 'Pendiente', 'OK', 'OK', 'Pendiente', 'OK'],
        'Cableado': ['OK', 'Pendiente', 'Pendiente', 'Pendiente', 'OK', 'Pendiente', 'Pendiente', 'Pendiente'],
        'Conexión Equipo': ['OK', 'Pendiente', 'Pendiente', 'Pendiente', 'Pendiente', 'Pendiente', 'Pendiente', 'Pendiente'],
        'Conexión DCS': ['Pendiente'] * 8,
        'Marquillado Equipo': ['OK'] * 8,
        'Marquillado Cable': ['OK', 'Pendiente', 'Pendiente', 'Pendiente', 'OK', 'Pendiente', 'Pendiente', 'Pendiente'],
        'Suiministro de Aire/Tubing': ['N/A', 'N/A', 'N/A', 'N/A', 'OK', 'N/A', 'N/A', 'N/A'],
        'Pre-Comisionamiento': ['Pendiente'] * 8
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
    """Crea una imagen de una tabla usando Plotly."""
    df_display = df.head(max_filas).copy()
    
    columnas = list(df_display.columns)
    valores = [df_display[col].tolist() for col in columnas]
    
    for i, val_list in enumerate(valores):
        valores[i] = [str(v)[:50] + '...' if len(str(v)) > 50 else str(v) for v in val_list]
    
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
    ancho = max(1200, len(columnas) * 150)
    
    fig.update_layout(
        title=dict(
            text=f'<b>{titulo}</b><br><sub>{len(df)} equipos totales - Mostrando primeros {len(df_display)}</sub>',
            x=0.5,
            xanchor='center',
            font=dict(size=16, color='#1f77b4')
        ),
        height=altura,
        width=ancho,
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
    Calcula cuántos equipos tienen una actividad completada.
    Para 'Suministro de Aire', solo cuenta equipos que NO tienen N/A.
    
    Returns:
        tuple: (completados, pendientes, porcentaje, total_aplicable)
    """
    # Caso especial: Suministro de Aire - Excluir N/A
    if 'suministro' in actividad.lower() or 'suiministro' in actividad.lower():
        # Filtrar equipos donde NO sea N/A
        # Excluir solo los que tienen N/A o 'No aplica' — los blancos son PENDIENTES
        VALORES_NO_APLICA = ['N/A', 'NA', 'n/a', 'na', 'N/a',
                             'no', 'No', 'NO', 'no aplica', 'No Aplica', 'NO APLICA']
        df_aplicable = df[~df[actividad].astype(str).str.strip().isin(VALORES_NO_APLICA)].copy()
        total = len(df_aplicable)
        
        if total == 0:
            return 0, 0, 0, 0
        
        # Contar completados en los equipos aplicables
        completados = df_aplicable[actividad].isin([
            'OK', 'SI', 'Completado', 'COMPLETADO', 'ok', 'X', 'x', 1, True
        ]).sum()
        
        pendientes = total - completados
        porcentaje = (completados / total) * 100 if total > 0 else 0
        
        return completados, pendientes, porcentaje, total
    
    # Para otras actividades, contar normalmente
    total = len(df)
    if total == 0:
        return 0, 0, 0, 0
    
    completados = df[actividad].notna().sum()
    
    try:
        valores_completados = df[actividad].isin([
            'OK', 'SI', 'Completado', 'COMPLETADO', 'ok', 'X', 'x', 1, True
        ]).sum()
        if valores_completados > 0:
            completados = valores_completados
    except:
        pass
    
    pendientes = total - completados
    porcentaje = (completados / total) * 100 if total > 0 else 0
    return completados, pendientes, porcentaje, total

# ============================================================================
# INTERFAZ PRINCIPAL
# ============================================================================

st.title("🏭 Dashboard de Seguimiento - Equipos I&C")
st.markdown("---")

# ============================================================================
# SIDEBAR - CONFIGURACIÓN Y FILTROS
# ============================================================================

with st.sidebar:
    st.header("⚙️ Configuración")
    
    tab_url, tab_archivo, tab_ejemplo = st.tabs(["🌐 URL", "📁 Archivo", "🧪 Ejemplo"])
    
    with tab_url:
        url_excel = st.text_input(
            "URL del archivo Excel:",
            value=URL_DEFECTO,
            help="URL directa del archivo Excel en la nube (Google Drive configurado por defecto)"
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
                st.info("📂 Usa la pestaña 'Archivo' para cargarlo")
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
    st.markdown("""
    ### 📖 Opciones de Carga de Datos
    
    **1. URL desde la nube (Por defecto):**
    - Ya está configurada la URL del proyecto en Google Drive
    - Solo haz clic en "🔄 Cargar/Actualizar Datos"
    
    **2. Archivo local:**
    - Sube el archivo Excel desde tu computadora
    
    **3. Archivo de ejemplo:**
    - Genera un Excel de prueba para familiarizarte con el sistema
    """)
    
# ============================================================================
# PROCESAMIENTO Y VISUALIZACIÓN DE DATOS
# ============================================================================

if df is not None:
    if 'ITEM' not in df.columns:
        st.warning("⚠️ La columna ITEM no existe en el archivo. Se usará el conteo total de filas.")
    
    actividades = [
        'A Instalar',
        'Instalación',
        'Canalización/Bandeja',
        'Cableado',
        'Conexión Equipo',
        'Conexión DCS',
        'Marquillado Equipo',
        'Marquillado Cable',
        'Suiministro de Aire/Tubing',
        'Pre-Comisionamiento'
    ]
    
    actividades_existentes = [col for col in actividades if col in df.columns]
    
    # ========================================================================
    # FILTROS DINÁMICOS CON SESSION STATE
    # ========================================================================
    
    with st.sidebar:
        st.header("🔍 Filtros")
        
        filtros_activos = {}
        
        if not st.session_state.filtros_inicializados:
            st.session_state.filtros = {}
            st.session_state.filtros_inicializados = True
        
        if 'AREA' in df.columns:
            areas = ['Todas'] + sorted(df['AREA'].dropna().unique().tolist())
            default_areas = st.session_state.filtros.get('AREA', ['Todas'])
            area_sel = st.multiselect("Área:", areas, default=default_areas, key='filtro_area')
            st.session_state.filtros['AREA'] = area_sel
            if 'Todas' not in area_sel:
                filtros_activos['AREA'] = area_sel
        
        if 'SISTEMA GENERAL' in df.columns:
            sistemas_gen = ['Todos'] + sorted(df['SISTEMA GENERAL'].dropna().unique().tolist())
            default_sist_gen = st.session_state.filtros.get('SISTEMA GENERAL', ['Todos'])
            sistema_gen_sel = st.multiselect("Sistema General:", sistemas_gen, default=default_sist_gen, key='filtro_sist_gen')
            st.session_state.filtros['SISTEMA GENERAL'] = sistema_gen_sel
            if 'Todos' not in sistema_gen_sel:
                filtros_activos['SISTEMA GENERAL'] = sistema_gen_sel
        
        if 'SISTEMA BMS/SMC/DCS' in df.columns:
            sistemas_bms = ['Todos'] + sorted(df['SISTEMA BMS/SMC/DCS'].dropna().unique().tolist())
            default_sist_bms = st.session_state.filtros.get('SISTEMA BMS/SMC/DCS', ['Todos'])
            sistema_bms_sel = st.multiselect("Sistema BMS/SMC/DCS:", sistemas_bms, default=default_sist_bms, key='filtro_sist_bms')
            st.session_state.filtros['SISTEMA BMS/SMC/DCS'] = sistema_bms_sel
            if 'Todos' not in sistema_bms_sel:
                filtros_activos['SISTEMA BMS/SMC/DCS'] = sistema_bms_sel
        
        if 'TIPO INSTRUMENTOS' in df.columns:
            tipos = ['Todos'] + sorted(df['TIPO INSTRUMENTOS'].dropna().unique().tolist())
            default_tipos = st.session_state.filtros.get('TIPO INSTRUMENTOS', ['Todos'])
            tipo_sel = st.multiselect("Tipo Instrumento:", tipos, default=default_tipos, key='filtro_tipo')
            st.session_state.filtros['TIPO INSTRUMENTOS'] = tipo_sel
            if 'Todos' not in tipo_sel:
                filtros_activos['TIPO INSTRUMENTOS'] = tipo_sel
        
        if 'Prioridad' in df.columns:
            prioridades = ['Todas'] + sorted(df['Prioridad'].dropna().unique().tolist())
            default_prior = st.session_state.filtros.get('Prioridad', ['Todas'])
            prioridad_sel = st.multiselect("Prioridad:", prioridades, default=default_prior, key='filtro_prior')
            st.session_state.filtros['Prioridad'] = prioridad_sel
            if 'Todas' not in prioridad_sel:
                filtros_activos['Prioridad'] = prioridad_sel
        
        if 'Hito' in df.columns:
            hitos = ['Todos'] + sorted(df['Hito'].dropna().unique().tolist())
            default_hitos = st.session_state.filtros.get('Hito', ['Todos'])
            hito_sel = st.multiselect("Hito:", hitos, default=default_hitos, key='filtro_hito')
            st.session_state.filtros['Hito'] = hito_sel
            if 'Todos' not in hito_sel:
                filtros_activos['Hito'] = hito_sel
        
        if st.button("🔄 Resetear Filtros", type="secondary"):
            st.session_state.filtros = {}
            st.session_state.filtros_inicializados = False
            st.rerun()
    
    df_filtrado = df.copy()
    for columna, valores in filtros_activos.items():
        df_filtrado = df_filtrado[df_filtrado[columna].isin(valores)]
    
    # ========================================================================
    # INFORMACIÓN GENERAL
    # ========================================================================
    
    col_info1, col_info2, col_info3, col_info4 = st.columns(4)
    with col_info1:
        st.info(f"📋 **Total Equipos (ITEM):** {len(df)}")
    with col_info2:
        st.info(f"🔍 **Equipos Filtrados:** {len(df_filtrado)}")
    with col_info3:
        st.success(f"✅ **Filtros Activos:** {len(filtros_activos)}")
    with col_info4:
        st.info(f"📊 **Fuente:** {fuente_datos}")
    
    st.markdown("---")
    
    # ========================================================================
    # MÉTRICAS DE AVANCE CON INFORMACIÓN DE EQUIPOS APLICABLES
    # ========================================================================
    
    st.header("📊 Métricas de Avance General")
    
    # Panel informativo para Suministro de Aire - MUY VISIBLE
    if 'Suiministro de Aire/Tubing' in actividades_existentes:
        total_global = len(df_filtrado)
        completados_sa, pendientes_sa, porcentaje_sa, total_aplicable_sa = calcular_completados(
            df_filtrado, 'Suiministro de Aire/Tubing'
        )
        equipos_na_sa = total_global - total_aplicable_sa
        
        col_info_sa1, col_info_sa2, col_info_sa3, col_info_sa4 = st.columns(4)
        
        with col_info_sa1:
            st.info(f"""
**🔧 Suministro de Aire**  
📊 {total_aplicable_sa} equipos lo requieren
""")
        
        with col_info_sa2:
            st.success(f"""
**✅ Completados**  
{completados_sa} de {total_aplicable_sa} → {porcentaje_sa:.1f}%
""")
        
        with col_info_sa3:
            st.warning(f"""
**⚠️ Pendientes**  
{pendientes_sa} equipos
""")
        
        with col_info_sa4:
            st.info(f"""
**⚪ No Aplica (N/A)**  
{equipos_na_sa} equipos
""")
        
        st.markdown("---")
    
    metricas_cols = st.columns(5)
    
    for idx, actividad in enumerate(actividades_existentes):
        with metricas_cols[idx % 5]:
            total = len(df_filtrado)
            if total > 0:
                completados, pendientes, porcentaje, total_aplicable = calcular_completados(df_filtrado, actividad)
                
                # Etiqueta especial para Suministro de Aire
                if 'suministro' in actividad.lower() or 'suiministro' in actividad.lower():
                    equipos_na = total - total_aplicable
                    etiqueta = actividad.replace('Suiministro', 'Suministro')
                    
                    st.metric(
                        label=etiqueta,
                        value=f"{completados}/{total_aplicable}",
                        delta=f"{porcentaje:.1f}%"
                    )
                    st.caption(f"✅ {total_aplicable} aplican | ⚪ {equipos_na} N/A")
                else:
                    etiqueta = actividad.replace('Suiministro', 'Suministro')
                    
                    st.metric(
                        label=etiqueta,
                        value=f"{completados}/{total}",
                        delta=f"{porcentaje:.1f}%"
                    )
            else:
                st.metric(label=actividad, value="0/0", delta="0%")
    
    st.markdown("---")
    
    # ========================================================================
    # GRÁFICOS DE PROGRESO CON NÚMEROS VISIBLES
    # ========================================================================
    
    st.header("📈 Progreso por Actividad")
    
    col_graph1, col_graph2 = st.columns(2)
    
    with col_graph1:
        avance_data = []
        for actividad in actividades_existentes:
            total = len(df_filtrado)
            if total > 0:
                completados, pendientes, porcentaje, total_aplicable = calcular_completados(df_filtrado, actividad)
                avance_data.append({
                    'Actividad': actividad.replace('Suiministro', 'Suministro'),
                    'Completados': completados,
                    'Pendientes': pendientes,
                    'Porcentaje': porcentaje
                })
        
        df_avance = pd.DataFrame(avance_data)
        
        fig_barras = go.Figure()
        
        fig_barras.add_trace(go.Bar(
            name='Completados',
            x=df_avance['Actividad'],
            y=df_avance['Completados'],
            marker_color='#10b981',
            text=df_avance['Completados'],
            textposition='inside',
            textfont=dict(size=12, color='white', family='Arial Black')
        ))
        
        fig_barras.add_trace(go.Bar(
            name='Pendientes',
            x=df_avance['Actividad'],
            y=df_avance['Pendientes'],
            marker_color='#ef4444',
            text=df_avance['Pendientes'],
            textposition='inside',
            textfont=dict(size=12, color='white', family='Arial Black')
        ))
        
        fig_barras.update_layout(
            barmode='stack',
            title='Estado de Actividades',
            xaxis_title='Actividad',
            yaxis_title='Cantidad',
            xaxis_tickangle=-45,
            height=400,
            showlegend=True
        )
        
        st.plotly_chart(fig_barras, use_container_width=True)
    
    with col_graph2:
        fig_pct = go.Figure()
        
        fig_pct.add_trace(go.Bar(
            x=df_avance['Actividad'],
            y=df_avance['Porcentaje'],
            marker=dict(
                color=df_avance['Porcentaje'],
                colorscale='RdYlGn',
                cmin=0,
                cmax=100
            ),
            text=[f"{p:.1f}%" for p in df_avance['Porcentaje']],
            textposition='outside',
            textfont=dict(size=11, color='black', family='Arial Black')
        ))
        
        fig_pct.update_layout(
            title='Porcentaje de Completitud por Actividad',
            xaxis_title='Actividad',
            yaxis_title='% Completado',
            xaxis_tickangle=-45,
            height=400,
            yaxis=dict(range=[0, max(df_avance['Porcentaje'].max() * 1.1, 10)])
        )
        
        st.plotly_chart(fig_pct, use_container_width=True)
    
    col_btn_img1, col_btn_img2, col_btn_img3 = st.columns([1, 1, 2])
    with col_btn_img1:
        try:
            st.download_button(
                label="📸 Descargar Gráfico Barras (PNG)",
                data=fig_barras.to_image(format="png", width=1200, height=600),
                file_name=f"grafico_barras_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                mime="image/png"
            )
        except:
            st.info("ℹ️ PNG requiere configuración adicional")
    
    with col_btn_img2:
        try:
            st.download_button(
                label="📸 Descargar Gráfico % (PNG)",
                data=fig_pct.to_image(format="png", width=1200, height=600),
                file_name=f"grafico_porcentaje_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                mime="image/png"
            )
        except:
            st.info("ℹ️ PNG requiere configuración adicional")
    
    st.markdown("---")
    
    # ========================================================================
    # EQUIPOS PENDIENTES - SELECCIÓN MÚLTIPLE CON EXPORTACIÓN IMAGEN
    # ========================================================================
    
    st.header("⚠️ Equipos Pendientes por Actividad")
    
    actividades_seleccionadas = st.multiselect(
        "Selecciona una o más actividades para ver equipos pendientes:",
        actividades_existentes,
        default=[actividades_existentes[0]] if actividades_existentes else []
    )
    
    if actividades_seleccionadas:
        df_pendientes_combinado = df_filtrado.copy()
        
        for actividad in actividades_seleccionadas:
            if 'suministro' in actividad.lower() or 'suiministro' in actividad.lower():
                mask_aplicable = ~df_pendientes_combinado[actividad].isin(['N/A', 'NA', 'n/a', 'na', 'N/a'])
                mask_pendiente = ~df_pendientes_combinado[actividad].isin([
                    'OK', 'SI', 'Completado', 'COMPLETADO', 'ok', 'X', 'x', 1, True
                ]) | df_pendientes_combinado[actividad].isna()
                df_pendientes_combinado[f'{actividad}_Pendiente'] = mask_aplicable & mask_pendiente
            else:
                df_pendientes_combinado[f'{actividad}_Pendiente'] = ~df_pendientes_combinado[actividad].isin([
                    'OK', 'SI', 'Completado', 'COMPLETADO', 'ok', 'X', 'x', 1, True
                ]) | df_pendientes_combinado[actividad].isna()
        
        condiciones = [df_pendientes_combinado[f'{act}_Pendiente'] for act in actividades_seleccionadas]
        df_pendientes = df_pendientes_combinado[pd.concat(condiciones, axis=1).any(axis=1)]
        
        col_pend1, col_pend2, col_pend3 = st.columns([1, 1, 2])
        with col_pend1:
            st.metric(
                "Equipos con Pendientes",
                len(df_pendientes),
                delta=f"{(len(df_pendientes)/len(df_filtrado)*100):.1f}% del total" if len(df_filtrado) > 0 else "0%"
            )
        with col_pend2:
            st.metric(
                "Actividades Seleccionadas",
                len(actividades_seleccionadas)
            )
        
        if len(df_pendientes) > 0:
            st.subheader(f"Listado de Equipos con Pendientes ({len(actividades_seleccionadas)} actividad{'es' if len(actividades_seleccionadas) > 1 else ''})")
            
            columnas_base = ['ITEM', 'TAG', 'DESCRIPTION', 'AREA', 'SISTEMA GENERAL', 'TIPO INSTRUMENTOS', 'Prioridad']
            columnas_base = [col for col in columnas_base if col in df_pendientes.columns]
            columnas_mostrar = columnas_base + actividades_seleccionadas
            
            st.dataframe(
                df_pendientes[columnas_mostrar],
                use_container_width=True,
                height=400
            )
            
            st.info(f"💡 **Nota:** La tabla muestra todos los equipos que tienen al menos una de las {len(actividades_seleccionadas)} actividades pendientes.")
            
            col_desc1, col_desc2, col_desc3 = st.columns(3)
            
            actividades_nombre = "_".join([act.replace('/', '-').replace(' ', '_') for act in actividades_seleccionadas[:3]])
            if len(actividades_seleccionadas) > 3:
                actividades_nombre += f"_y_{len(actividades_seleccionadas)-3}_mas"
            
            with col_desc1:
                excel_buffer = crear_excel_descarga(df_pendientes[columnas_mostrar], "Pendientes")
                st.download_button(
                    label=f"📥 Descargar Tabla (Excel)",
                    data=excel_buffer,
                    file_name=f"pendientes_{actividades_nombre}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            with col_desc2:
                try:
                    titulo_tabla = f"Equipos Pendientes - {', '.join(actividades_seleccionadas[:2])}"
                    if len(actividades_seleccionadas) > 2:
                        titulo_tabla += f" y {len(actividades_seleccionadas)-2} más"
                    
                    fig_tabla = crear_tabla_imagen(
                        df_pendientes[columnas_mostrar],
                        titulo=titulo_tabla,
                        max_filas=50
                    )
                    
                    st.download_button(
                        label=f"📸 Descargar Tabla (PNG)",
                        data=fig_tabla.to_image(format="png", width=fig_tabla.layout.width, height=fig_tabla.layout.height),
                        file_name=f"tabla_pendientes_{actividades_nombre}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                        mime="image/png"
                    )
                except:
                    st.info("ℹ️ PNG requiere configuración adicional")
            
            with col_desc3:
                if len(df_pendientes) > 50:
                    st.warning(f"⚠️ La imagen PNG muestra solo las primeras 50 filas de {len(df_pendientes)}. Descarga Excel para ver todos.")
            
            with st.expander("📊 Ver Resumen por Actividad"):
                resumen_data = []
                for actividad in actividades_seleccionadas:
                    pendientes_act = df_pendientes[df_pendientes[f'{actividad}_Pendiente']].shape[0]
                    resumen_data.append({
                        'Actividad': actividad,
                        'Pendientes': pendientes_act,
                        'Porcentaje': f"{(pendientes_act/len(df_filtrado)*100):.1f}%" if len(df_filtrado) > 0 else "0%"
                    })
                
                df_resumen = pd.DataFrame(resumen_data)
                st.dataframe(df_resumen, use_container_width=True)
                
                excel_resumen = crear_excel_descarga(df_resumen, "Resumen")
                st.download_button(
                    label="📥 Descargar Resumen (Excel)",
                    data=excel_resumen,
                    file_name=f"resumen_pendientes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.success("✅ ¡No hay equipos pendientes para las actividades seleccionadas!")
    else:
        st.info("👆 Selecciona al menos una actividad para ver los equipos pendientes")
    
    st.markdown("---")
    
    # ========================================================================
    # ANÁLISIS MULTIDIMENSIONAL CON EXPORTACIÓN
    # ========================================================================
    
    st.header("🔍 Análisis Multidimensional")
    
    tab1, tab2, tab3, tab4 = st.tabs(["📍 Por Área", "⚙️ Por Sistema", "🔧 Por Tipo", "🎯 Por Prioridad"])
    
    with tab1:
        if 'AREA' in df_filtrado.columns:
            analisis_area = df_filtrado.groupby('AREA').size().reset_index(name='Cantidad')
            
            col_a1, col_a2 = st.columns([2, 1])
            with col_a1:
                fig_area = px.pie(
                    analisis_area,
                    values='Cantidad',
                    names='AREA',
                    title='Distribución de Equipos por Área',
                    hole=0.3
                )
                fig_area.update_traces(textposition='inside', textinfo='percent+label+value')
                st.plotly_chart(fig_area, use_container_width=True)
                
                try:
                    st.download_button(
                        label="📸 Descargar Gráfico (PNG)",
                        data=fig_area.to_image(format="png", width=1200, height=800),
                        file_name=f"analisis_area_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                        mime="image/png"
                    )
                except:
                    st.info("ℹ️ PNG requiere configuración adicional")
            
            with col_a2:
                st.dataframe(analisis_area.sort_values('Cantidad', ascending=False), 
                           use_container_width=True, height=400)
                
                excel_area = crear_excel_descarga(analisis_area, "Por Área")
                st.download_button(
                    label="📥 Descargar Tabla (Excel)",
                    data=excel_area,
                    file_name=f"tabla_por_area_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    
    with tab2:
        if 'SISTEMA GENERAL' in df_filtrado.columns:
            analisis_sistema = df_filtrado.groupby('SISTEMA GENERAL').size().reset_index(name='Cantidad')
            
            fig_sistema = px.bar(
                analisis_sistema.sort_values('Cantidad', ascending=True),
                x='Cantidad',
                y='SISTEMA GENERAL',
                title='Equipos por Sistema General',
                orientation='h',
                color='Cantidad',
                color_continuous_scale='Blues',
                text='Cantidad'
            )
            fig_sistema.update_traces(textposition='outside')
            st.plotly_chart(fig_sistema, use_container_width=True)
            
            col_s1, col_s2 = st.columns(2)
            with col_s1:
                try:
                    st.download_button(
                        label="📸 Descargar Gráfico (PNG)",
                        data=fig_sistema.to_image(format="png", width=1200, height=800),
                        file_name=f"analisis_sistema_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                        mime="image/png"
                    )
                except:
                    st.info("ℹ️ PNG requiere configuración adicional")
            with col_s2:
                excel_sistema = crear_excel_descarga(analisis_sistema, "Por Sistema")
                st.download_button(
                    label="📥 Descargar Tabla (Excel)",
                    data=excel_sistema,
                    file_name=f"tabla_sistema_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            st.dataframe(analisis_sistema.sort_values('Cantidad', ascending=False), use_container_width=True)
    
    with tab3:
        if 'TIPO INSTRUMENTOS' in df_filtrado.columns:
            analisis_tipo = df_filtrado.groupby('TIPO INSTRUMENTOS').size().reset_index(name='Cantidad')
            
            fig_tipo = px.bar(
                analisis_tipo.sort_values('Cantidad', ascending=False),
                x='TIPO INSTRUMENTOS',
                y='Cantidad',
                title='Equipos por Tipo de Instrumento',
                color='Cantidad',
                color_continuous_scale='Viridis',
                text='Cantidad'
            )
            fig_tipo.update_layout(xaxis_tickangle=-45)
            fig_tipo.update_traces(textposition='outside')
            st.plotly_chart(fig_tipo, use_container_width=True)
            
            col_t1, col_t2 = st.columns(2)
            with col_t1:
                try:
                    st.download_button(
                        label="📸 Descargar Gráfico (PNG)",
                        data=fig_tipo.to_image(format="png", width=1200, height=800),
                        file_name=f"analisis_tipo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                        mime="image/png"
                    )
                except:
                    st.info("ℹ️ PNG requiere configuración adicional")
            with col_t2:
                excel_tipo = crear_excel_descarga(analisis_tipo, "Por Tipo")
                st.download_button(
                    label="📥 Descargar Tabla (Excel)",
                    data=excel_tipo,
                    file_name=f"tabla_tipo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            st.dataframe(analisis_tipo.sort_values('Cantidad', ascending=False), use_container_width=True)
    
    with tab4:
        if 'Prioridad' in df_filtrado.columns:
            analisis_prioridad = df_filtrado.groupby('Prioridad').size().reset_index(name='Cantidad')
            
            col_p1, col_p2 = st.columns([2, 1])
            with col_p1:
                fig_prioridad = px.pie(
                    analisis_prioridad,
                    values='Cantidad',
                    names='Prioridad',
                    title='Distribución por Prioridad',
                    color='Prioridad',
                    color_discrete_map={'Alta': '#ef4444', 'Media': '#f59e0b', 'Baja': '#10b981'}
                )
                fig_prioridad.update_traces(textposition='inside', textinfo='percent+label+value')
                st.plotly_chart(fig_prioridad, use_container_width=True)
                
                try:
                    st.download_button(
                        label="📸 Descargar Gráfico (PNG)",
                        data=fig_prioridad.to_image(format="png", width=1200, height=800),
                        file_name=f"analisis_prioridad_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                        mime="image/png"
                    )
                except:
                    st.info("ℹ️ PNG requiere configuración adicional")
            
            with col_p2:
                st.dataframe(analisis_prioridad.sort_values('Cantidad', ascending=False),
                           use_container_width=True, height=400)
                
                excel_prioridad = crear_excel_descarga(analisis_prioridad, "Por Prioridad")
                st.download_button(
                    label="📥 Descargar Tabla (Excel)",
                    data=excel_prioridad,
                    file_name=f"tabla_prioridad_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    
    st.markdown("---")
    
    # ========================================================================
    # TABLA COMPLETA
    # ========================================================================
    
    st.header("📋 Tabla Completa de Equipos")
    
    col_opt1, col_opt2, col_opt3 = st.columns([1, 2, 1])
    with col_opt1:
        mostrar_todas_columnas = st.checkbox("Mostrar todas las columnas", value=False)
    with col_opt2:
        buscar_tag = st.text_input("🔍 Buscar por TAG:", placeholder="Ingresa TAG...")
    with col_opt3:
        st.metric("Total Mostrados", len(df_filtrado))
    
    df_mostrar = df_filtrado.copy()
    if buscar_tag and 'TAG' in df_mostrar.columns:
        df_mostrar = df_mostrar[df_mostrar['TAG'].str.contains(buscar_tag, case=False, na=False)]
    
    if not mostrar_todas_columnas:
        columnas_default = ['ITEM', 'TAG', 'TIPO INSTRUMENTOS', 'AREA', 'SISTEMA GENERAL', 
                           'DESCRIPTION', 'Prioridad'] + actividades_existentes
        columnas_default = [col for col in columnas_default if col in df_mostrar.columns]
        df_mostrar = df_mostrar[columnas_default]
    
    st.dataframe(df_mostrar, use_container_width=True, height=500)
    
    col_desc1, col_desc2 = st.columns(2)
    with col_desc1:
        excel_completo = crear_excel_descarga(df_mostrar, "Equipos Filtrados")
        st.download_button(
            label="📥 Descargar Tabla Filtrada (Excel)",
            data=excel_completo,
            file_name=f"equipos_ic_filtrados_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    # ========================================================================
    # PIE DE PÁGINA
    # ========================================================================
    
    st.markdown("---")
    col_footer1, col_footer2, col_footer3 = st.columns(3)
    with col_footer1:
        st.info(f"📊 **Última actualización:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    with col_footer2:
        st.success(f"✅ **Datos cargados correctamente**")
    with col_footer3:
        st.info(f"📈 **Versión:** 1.0.7")
