"""
Dashboard de Seguimiento - Equipos I&C
======================================
Sistema completo de seguimiento y análisis de avance para proyectos de Instrumentación y Control.

Autor: Dashboard I&C
Fecha: Febrero 2026
Versión: 1.9.2 - Hito S formateado (1.0→1, orden numérico en filtro)

USO:
    streamlit run app_dashboard_ic.py
    
DEPENDENCIAS:
    pip install streamlit pandas openpyxl plotly xlrd kaleido
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
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
        'PRO': ['Vapor', 'Vapor', 'Agua', 'Agua', 'Vapor', 'Combustible', 'Combustible', 'Agua'],
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
        'Categoria': ['Alta', 'Media', 'Alta', 'Baja', 'Alta', 'Media', 'Baja', 'Alta'],
        'Pre Emsanblado': ['OK', 'OK', 'Pendiente', 'OK', 'OK', 'OK', 'Pendiente', 'OK'],
        'A Instalar': ['OK'] * 8,
        'Instalación': ['OK', 'OK', 'Pendiente', 'OK', 'OK', 'OK', 'Pendiente', 'Pendiente'],
        'Canalización/Bandeja': ['OK', 'OK', 'OK', 'Pendiente', 'OK', 'OK', 'Pendiente', 'OK'],
        'Cableado': ['OK', 'Pendiente', 'Pendiente', 'Pendiente', 'OK', 'Pendiente', 'Pendiente', 'Pendiente'],
        'Conexión Equipo': ['OK', 'Pendiente', 'Pendiente', 'Pendiente', 'Pendiente', 'Pendiente', 'Pendiente', 'Pendiente'],
        'Conexión DCS': ['Pendiente'] * 8,
        'Marquillado Equipo': ['OK'] * 8,
        'Marquillado Cable': ['OK', 'Pendiente', 'Pendiente', 'Pendiente', 'OK', 'Pendiente', 'Pendiente', 'Pendiente'],
        'Suiministro de Aire': ['N/A', 'N/A', 'N/A', 'N/A', 'OK', 'N/A', 'N/A', 'N/A'],
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
        df = pd.read_excel(url_excel, sheet_name="Equipos I&C",
                           keep_default_na=False, na_values=[""])
        
        if 'ITEM' in df.columns:
            df = df[df['ITEM'].notna()].copy()
            df = df[df['ITEM'].astype(str).str.strip() != ''].copy()
        

        # Formatear columna Hito S: enteros como "1","2", decimales como "1.1"
        if 'Hito S' in df.columns:
            def _fmt_hito(x):
                try:
                    n = float(x)
                    if n == int(n):
                        return str(int(n))
                    return str(round(n, 4)).rstrip('0').rstrip('.')
                except (ValueError, TypeError):
                    return x
            df['Hito S'] = df['Hito S'].apply(_fmt_hito)
            # Ordenar de forma numérica en el filtro (guardar key de ordenamiento)
            def _sort_key(v):
                try:    return float(v)
                except: return float('inf')
            df['_hito_s_sort'] = df['Hito S'].apply(_sort_key)

        return df
    except Exception as e:
        st.error(f"Error al cargar el archivo: {str(e)}")
        return None


def _scurve(t):
    import math
    if t <= 0: return 0.0
    if t >= 1: return 1.0
    return 1.0 / (1.0 + math.exp(-12.0 * (t - 0.5)))

PESOS_CON_SA = {
    'A Instalar': 0,
    'Instalacion': 20.0, 'Instalación': 20.0,
    'Canalizacion/Bandeja': 30.0, 'Canalización/Bandeja': 30.0,
    'Cableado': 20.0,
    'Conexion Equipo': 2.5, 'Conexión Equipo': 2.5,
    'Conexion DCS': 2.5, 'Conexión DCS': 2.5,
    'Marquillado Equipo': 1.0, 'Marquillado Cable': 1.0,
    'Suiministro de Aire/Tubing': 20.0,
    'Pre-Comisionamiento': 3.0,
}

PESOS_SIN_SA = {
    'A Instalar': 0,
    'Instalacion': 20.0, 'Instalación': 20.0,
    'Canalizacion/Bandeja': 30.0, 'Canalización/Bandeja': 30.0,
    'Cableado': 20.0,
    'Conexion Equipo': 10.0, 'Conexión Equipo': 10.0,
    'Conexion DCS': 10.0, 'Conexión DCS': 10.0,
    'Marquillado Equipo': 2.5, 'Marquillado Cable': 2.5,
    'Suiministro de Aire/Tubing': 0.0,
    'Pre-Comisionamiento': 5.0,
}

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
        df_aplicable = df[~df[actividad].isin(['N/A', 'NA', 'n/a', 'na', 'N/a'])].copy()
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
        _COLS_F = [
            ("Hito",                "Hito",             "Todos"),
            ("PRO",                 "Sistema General",  "Todos"),
            ("AREA",                "Area",             "Todas"),
            ("SISTEMA BMS/SMC/DCS", "Sistema BMS/DCS",  "Todos"),
            ("TIPO INSTRUMENTOS",   "Tipo Instrumento", "Todos"),
            ("Hito S",              "Categoria",        "Todas"),
        ]
        _CA_F = [(c, l, t) for c, l, t in _COLS_F if c in df.columns]

        def _opts_f(col_obj):
            df_t = df.copy()
            for col, lbl, tod in _CA_F:
                if col == col_obj:
                    continue
                v = st.session_state.get("dyn_" + col, [tod])
                if v and tod not in v:
                    df_t = df_t[df_t[col].isin(v)]
            vals = df_t[col_obj].dropna().unique().tolist()
            # Orden numérico para Hito S
            if col_obj == "Hito S" and "_hito_s_sort" in df_t.columns:
                sort_map = df_t.drop_duplicates("Hito S").set_index("Hito S")["_hito_s_sort"].to_dict()
                vals = sorted(vals, key=lambda x: sort_map.get(x, float("inf")))
            else:
                vals = sorted(vals)
            return vals

        for _col_f, _lbl_f, _tod_f in _CA_F:
            _opts_v = _opts_f(_col_f)
            _cur_v  = st.session_state.get("dyn_" + _col_f, [_tod_f])
            _cur_v  = [x for x in _cur_v if x == _tod_f or x in _opts_v] or [_tod_f]
            st.multiselect(_lbl_f + ":", [_tod_f] + _opts_v,
                           default=_cur_v, key="dyn_" + _col_f)

        if st.button("Resetear Filtros", type="secondary"):
            for _col_f, _, _ in _CA_F:
                st.session_state.pop("dyn_" + _col_f, None)
            st.rerun()
    
    df_filtrado = df.copy()
    for _cf, _tf in [("Hito", "Todos"), ("PRO", "Todos"),
                     ("AREA", "Todas"), ("SISTEMA BMS/SMC/DCS", "Todos"),
                     ("TIPO INSTRUMENTOS", "Todos"), ("Hito S", "Todas")]:
        if _cf not in df.columns:
            continue
        _vf = st.session_state.get("dyn_" + _cf, [_tf])
        if _vf and _tf not in _vf:
            df_filtrado = df_filtrado[df_filtrado[_cf].isin(_vf)]
    filtros_activos = {}
    
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
    
    st.header("📊 Metricas de Avance General")
    if 'Suiministro de Aire/Tubing' in actividades_existentes:
        _tg2=len(df_filtrado)
        _cs2,_ps2,_pp2,_ta2=calcular_completados(df_filtrado,'Suiministro de Aire/Tubing')
        _na2=_tg2-_ta2
        _bb1,_bb2,_bb3,_bb4=st.columns(4)
        with _bb1: st.info(f"**🔧 Suministro de Aire/Tubing**\n\n{_ta2} equipos lo requieren")
        with _bb2: st.success(f"**✅ Completados**\n\n{_cs2} de {_ta2} — {_pp2:.1f}%")
        with _bb3: st.warning(f"**⚠️ Pendientes**\n\n{_ps2} equipos")
        with _bb4: st.info(f"**⚪ No Aplica**\n\n{_na2} equipos")
        st.markdown("---")
    if actividades_existentes and len(df_filtrado)>0:
        _sa3=next((a for a in actividades_existentes if "suiministro" in a.lower() or "suministro" in a.lower()),None)
        _ts3=calcular_completados(df_filtrado,_sa3)[3] if _sa3 else 0
        _sa3_ok=_sa3 is not None and _ts3>0
        _pw3=PESOS_CON_SA if _sa3_ok else PESOS_SIN_SA
        _modo3="Con SA/Tubing" if _sa3_ok else "Sin SA/Tubing"
        _spxp3=0.0; _tc3=0; _tp3=0; _det3=[]
        for _act3 in actividades_existentes:
            _c3,_p3,_pct3,_=calcular_completados(df_filtrado,_act3)
            _tc3+=_c3; _tp3+=_p3; _pe3=_pw3.get(_act3,0); _spxp3+=_pe3*_pct3
            _det3.append({"Actividad":_act3.replace("Suiministro","Suministro"),
                "Peso":_pe3,"Avance":round(_pct3,1),"Contribucion":round(_pe3*_pct3/100,2)})
        _pct_gen3=round(_spxp3/100,1)
        _cb3="#10b981" if _pct_gen3>=75 else ("#f59e0b" if _pct_gen3>=40 else "#ef4444")
        st.markdown(f'''<div style="background:linear-gradient(135deg,#1e293b,#0f172a);border-radius:12px;padding:20px 28px;margin-bottom:18px;border-left:6px solid {_cb3};display:flex;align-items:center;gap:32px;"><div style="flex:0 0 auto;"><div style="color:#94a3b8;font-size:12px;font-weight:600;letter-spacing:1px;text-transform:uppercase;">Avance General del Proyecto</div><div style="color:{_cb3};font-size:52px;font-weight:800;line-height:1.1;margin-top:4px;">{_pct_gen3}%</div><div style="color:#94a3b8;font-size:11px;margin-top:4px;">Pesos: <b style="color:#cbd5e1">{_modo3}</b></div></div><div style="flex:1;min-width:180px;"><div style="background:#334155;border-radius:8px;height:18px;overflow:hidden;"><div style="width:{_pct_gen3}%;height:100%;background:linear-gradient(90deg,{_cb3}99,{_cb3});border-radius:8px;"></div></div><div style="display:flex;justify-content:space-between;margin-top:10px;gap:12px;flex-wrap:wrap;"><div style="text-align:center;"><div style="color:#10b981;font-size:20px;font-weight:700;">{_tc3}</div><div style="color:#94a3b8;font-size:11px;">Completados</div></div><div style="text-align:center;"><div style="color:#ef4444;font-size:20px;font-weight:700;">{_tp3}</div><div style="color:#94a3b8;font-size:11px;">Pendientes</div></div><div style="text-align:center;"><div style="color:#60a5fa;font-size:20px;font-weight:700;">{len(actividades_existentes)}</div><div style="color:#94a3b8;font-size:11px;">Actividades</div></div><div style="text-align:center;"><div style="color:#f8fafc;font-size:20px;font-weight:700;">{len(df_filtrado)}</div><div style="color:#94a3b8;font-size:11px;">Equipos</div></div></div></div></div>''',unsafe_allow_html=True)
        with st.expander("📊 Ver detalle de pesos por actividad"):
            _dfd3=pd.DataFrame(_det3)
            def _cpct3(val):
                if val>=75: return "background-color:#d1fae5;color:#065f46;font-weight:600"
                elif val>=40: return "background-color:#fef3c7;color:#92400e;font-weight:600"
                return "background-color:#fee2e2;color:#991b1b;font-weight:600"
            st.dataframe(_dfd3[["Actividad","Peso","Avance","Contribucion"]].style.applymap(_cpct3,subset=["Avance"]).format({"Peso":"{:.1f}%","Avance":"{:.1f}%","Contribucion":"{:.2f}%"}),use_container_width=True)
            st.caption(f"Tabla: **{_modo3}** | Σ(Peso×Avance)/100={_spxp3:.1f}/100=**{_pct_gen3}%**")
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

    st.header("📅 Avance vs Meta")
    _col_ff,_col_info=st.columns([1,2])
    with _col_ff:
        _ff=st.date_input("Fecha límite de entrega",value=datetime(2026,12,31).date(),key="av_ff")
    if actividades_existentes and len(df_filtrado)>0:
        _sa5=next((a for a in actividades_existentes if "suiministro" in a.lower() or "suministro" in a.lower()),None)
        _ts5=calcular_completados(df_filtrado,_sa5)[3] if _sa5 else 0
        _pw5=PESOS_CON_SA if (_sa5 and _ts5>0) else PESOS_SIN_SA
        _sp5=sum(_pw5.get(a,0)*calcular_completados(df_filtrado,a)[2] for a in actividades_existentes)
        _pr5=round(_sp5/100,1)
    else:
        _pr5=0.0
    from datetime import date as _dt5
    _hoy5=_dt5.today(); _dias_rest=max(0,(_ff-_hoy5).days)
    _faltante=round(100.0-_pr5,1); _vencido=_hoy5>_ff
    _col_color="#10b981" if _pr5>=75 else ("#f59e0b" if _pr5>=40 else "#ef4444")
    if _pr5>=100: _falt_color="#10b981"; _falt_label="✅ completado!"
    elif _vencido: _falt_color="#ef4444"; _falt_label="🚨 vencido — falta completar"
    else: _falt_color="#10b981"; _falt_label="⏳ para completar"
    with _col_info:
        if _pr5>=100: st.success(f"**Fecha límite:** {_ff.strftime('%d/%m/%Y')}  |  ✅ Proyecto completado!")
        elif _vencido: st.error(f"**Fecha límite:** {_ff.strftime('%d/%m/%Y')}  |  🚨 Fecha vencida — {_faltante}% pendiente")
        else: st.info(f"**Fecha límite:** {_ff.strftime('%d/%m/%Y')}  |  ⏳ Faltan {_dias_rest} días para la entrega")
    st.markdown(f'''<div style="text-align:center;margin-bottom:-10px;">
        <span style="font-size:18px;font-weight:700;color:#f8fafc;">Avance Real del Proyecto</span><br>
        <span style="font-size:13px;color:#94a3b8;">vs Meta 100%</span></div>''',unsafe_allow_html=True)
    import plotly.graph_objects as go
    _fig5=go.Figure(go.Indicator(
        mode="gauge+number+delta",value=_pr5,
        delta={"reference":100,"valueformat":".1f","suffix":"%","relative":False,"font":{"size":20}},
        number={"suffix":"%","font":{"size":60,"color":_col_color}},
        gauge={"axis":{"range":[0,100],"ticksuffix":"%","tickcolor":"#94a3b8","tickfont":{"color":"#94a3b8","size":12}},
               "bar":{"color":_col_color,"thickness":0.28},"bgcolor":"#1e293b","bordercolor":"#334155",
               "steps":[{"range":[0,40],"color":"rgba(239,68,68,0.18)"},
                        {"range":[40,75],"color":"rgba(245,158,11,0.18)"},
                        {"range":[75,100],"color":"rgba(16,185,129,0.18)"}],
               "threshold":{"line":{"color":"#f8fafc","width":3},"thickness":0.8,"value":100}},
        domain={"x":[0.1,0.9],"y":[0.05,1]}))
    _fig5.update_layout(height=340,plot_bgcolor="#0f172a",paper_bgcolor="#0f172a",
        font=dict(color="#f8fafc"),margin=dict(l=40,r=40,t=10,b=10))
    st.plotly_chart(_fig5,use_container_width=True)
    _km1,_km2,_km3,_km4=st.columns(4)
    with _km1: st.metric("Avance Real",f"{_pr5}%")
    with _km2:
        st.markdown(f'''<div style="padding:8px 0;">
            <p style="color:#94a3b8;font-size:14px;margin:0 0 4px 0;">Faltante</p>
            <p style="color:{_falt_color};font-size:32px;font-weight:700;margin:0;">{_faltante}%</p>
            <p style="color:{_falt_color};font-size:12px;margin:4px 0 0 0;">{_falt_label}</p>
            </div>''',unsafe_allow_html=True)
    with _km3: st.metric("Días restantes",str(_dias_rest))
    with _km4: st.metric("Fecha límite",_ff.strftime("%d/%m/%Y"))
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
            
            columnas_base = ['ITEM', 'TAG', 'DESCRIPTION', 'AREA', 'PRO', 'TIPO INSTRUMENTOS', 'Categoria']
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
    
    tab1, tab2, tab3, tab4 = st.tabs(["📍 Por Área", "⚙️ Por Sistema", "🔧 Por Tipo", "🎯 Por Categoria"])
    
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
        if 'PRO' in df_filtrado.columns:
            analisis_sistema = df_filtrado.groupby('PRO').size().reset_index(name='Cantidad')
            
            fig_sistema = px.bar(
                analisis_sistema.sort_values('Cantidad', ascending=True),
                x='Cantidad',
                y='PRO',
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
        if 'Categoria' in df_filtrado.columns:
            analisis_prioridad = df_filtrado.groupby('Categoria').size().reset_index(name='Cantidad')
            
            col_p1, col_p2 = st.columns([2, 1])
            with col_p1:
                fig_prioridad = px.pie(
                    analisis_prioridad,
                    values='Cantidad',
                    names='Categoria',
                    title='Distribución por Categoria',
                    color='Categoria',
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
                
                excel_prioridad = crear_excel_descarga(analisis_prioridad, "Por Categoria")
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
        columnas_default = ['ITEM', 'TAG', 'TIPO INSTRUMENTOS', 'AREA', 'PRO', 
                           'DESCRIPTION', 'Categoria'] + actividades_existentes
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
