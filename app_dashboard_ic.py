"""
Dashboard de Seguimiento - Equipos I&C
======================================
Sistema completo de seguimiento y análisis de avance para proyectos de Instrumentación y Control.

Autor: Dashboard I&C
Fecha: Febrero 2026
Versión: 1.3.0 - Filtro Pendientes / Ejecutados / Todos en sección de Equipos

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

URL_DEFECTO = "https://drive.google.com/uc?export=download&id=1x_uQhW4EKXiEgbLzZpF_InphP2oIItlu"

COL_REQUIERE_SA = "Requiere Suministro de Aire"
COL_SA          = "Suiministro de Aire"

VALORES_NA = ["N/A", "NA", "n/a", "na", "N/a"]
VALORES_OK = ["OK", "SI", "Completado", "COMPLETADO", "ok", "X", "x", 1, True]
VALORES_SI = ["si", "sí", "s", "yes", "y", "1", "true"]

# ============================================================================
# INICIALIZAR SESSION STATE
# ============================================================================

if "filtros_inicializados" not in st.session_state:
    st.session_state.filtros_inicializados = False
    st.session_state.filtros = {}

# ============================================================================
# FUNCIONES AUXILIARES
# ============================================================================

def verificar_dependencias():
    dependencias = {"streamlit": "streamlit", "pandas": "pandas",
                    "plotly": "plotly", "openpyxl": "openpyxl"}
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
    data = {
        "ITEM": [1, 2, 3, 4, 5, 6, 7, 8],
        "TAG":  ["FT-001", "PT-002", "TT-003", "LT-004", "FV-005", "PT-006", "TT-007", "LT-008"],
        "TIPO INSTRUMENTOS": [
            "Transmisor de Flujo", "Transmisor de Presión", "Transmisor de Temperatura",
            "Transmisor de Nivel", "Válvula de Control", "Transmisor de Presión",
            "Transmisor de Temperatura", "Transmisor de Nivel"
        ],
        "AREA": ["Area 100", "Area 100", "Area 200", "Area 200", "Area 100", "Area 300", "Area 300", "Area 200"],
        "SISTEMA GENERAL": ["Vapor", "Vapor", "Agua", "Agua", "Vapor", "Combustible", "Combustible", "Agua"],
        "FECHA ENTREGA": [
            "2026-03-15", "2026-03-15", "2026-04-01", "2026-04-01",
            "2026-03-15", "2026-05-01", "2026-05-01", "2026-04-01"
        ],
        "SISTEMA BMS/SMC/DCS": ["DCS", "DCS", "PLC", "PLC", "DCS", "DCS", "PLC", "PLC"],
        "SISTEMA": ["Sistema A", "Sistema A", "Sistema B", "Sistema B",
                    "Sistema A", "Sistema C", "Sistema C", "Sistema B"],
        "SIGNAL ASSOCIATION": ["AI-001", "AI-002", "AI-003", "AI-004",
                                "AO-001", "AI-005", "AI-006", "AI-007"],
        "DESCRIPTION": [
            "Flujo de vapor principal", "Presión de vapor", "Temperatura agua",
            "Nivel tanque principal", "Control flujo vapor", "Presión combustible",
            "Temperatura combustible", "Nivel tanque secundario"
        ],
        "SIGNAL": ["4-20mA"] * 8,
        "I/O":   ["AI", "AI", "AI", "AI", "AO", "AI", "AI", "AI"],
        "Hito":      ["Hito 1", "Hito 1", "Hito 2", "Hito 2", "Hito 1", "Hito 3", "Hito 3", "Hito 2"],
        "Prioridad": ["Alta", "Media", "Alta", "Baja", "Alta", "Media", "Baja", "Alta"],
        "Pre Emsanblado":           ["OK", "OK", "Pendiente", "OK", "OK", "OK", "Pendiente", "OK"],
        "A Instalar":               ["OK"] * 8,
        "Instalación":              ["OK", "OK", "Pendiente", "OK", "OK", "OK", "Pendiente", "Pendiente"],
        "Canalización/Bandeja":     ["OK", "OK", "OK", "Pendiente", "OK", "OK", "Pendiente", "OK"],
        "Cableado":                 ["OK", "Pendiente", "Pendiente", "Pendiente", "OK", "Pendiente", "Pendiente", "Pendiente"],
        "Conexión Equipo":          ["OK", "Pendiente", "Pendiente", "Pendiente", "Pendiente", "Pendiente", "Pendiente", "Pendiente"],
        "Conexión DCS":             ["Pendiente"] * 8,
        "Marquillado Equipo":       ["OK"] * 8,
        "Marquillado Cable":        ["OK", "Pendiente", "Pendiente", "Pendiente", "OK", "Pendiente", "Pendiente", "Pendiente"],
        "Requiere Suministro de Aire": ["No", "No", "No", "No", "Si", "Si", "No", "Si"],
        "Suiministro de Aire":         ["",   "",   "",   "",  "OK", "",   "",   "OK"],
        "Pre-Comisionamiento":         ["Pendiente"] * 8
    }
    df = pd.DataFrame(data)
    try:
        with pd.ExcelWriter("Equipos_IC_Ejemplo.xlsx", engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Equipos I&C", index=False)
        return True, "Equipos_IC_Ejemplo.xlsx"
    except Exception as e:
        return False, str(e)


def crear_excel_descarga(df, nombre_hoja="Datos"):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=nombre_hoja, index=False)
    buffer.seek(0)
    return buffer


def crear_tabla_imagen(df, titulo="Tabla", max_filas=50):
    df_display = df.head(max_filas).copy()
    columnas = list(df_display.columns)
    valores = [df_display[col].tolist() for col in columnas]
    for i, val_list in enumerate(valores):
        valores[i] = [str(v)[:50] + "..." if len(str(v)) > 50 else str(v) for v in val_list]
    fig = go.Figure(data=[go.Table(
        header=dict(
            values=[f"<b>{col}</b>" for col in columnas],
            fill_color="#1f77b4",
            font=dict(color="white", size=11, family="Arial"),
            align="center", height=30
        ),
        cells=dict(
            values=valores,
            fill_color=[["#f0f0f0", "white"] * len(df_display)],
            font=dict(color="black", size=10, family="Arial"),
            align="left", height=25
        )
    )])
    altura = min(800 + (len(df_display) * 25), 4000)
    ancho  = max(1200, len(columnas) * 150)
    fig.update_layout(
        title=dict(
            text=f"<b>{titulo}</b><br><sub>{len(df)} equipos - Mostrando primeros {len(df_display)}</sub>",
            x=0.5, xanchor="center", font=dict(size=16, color="#1f77b4")
        ),
        height=altura, width=ancho,
        margin=dict(l=20, r=20, t=80, b=20)
    )
    return fig


def _formatear_fecha(valor):
    try:
        return pd.to_datetime(valor).strftime("%d/%m/%Y")
    except Exception:
        return str(valor)

# ============================================================================
# FUNCIONES PRINCIPALES
# ============================================================================

@st.cache_data(ttl=300)
def cargar_datos(url_excel):
    try:
        df = pd.read_excel(url_excel, sheet_name="Equipos I&C")
        if "ITEM" in df.columns:
            df = df[df["ITEM"].notna()].copy()
            df = df[df["ITEM"].astype(str).str.strip() != ""].copy()
        return df
    except Exception as e:
        st.error(f"Error al cargar el archivo: {str(e)}")
        return None


def _es_suministro(nombre):
    return "suministro" in nombre.lower() or "suiministro" in nombre.lower()


def _mask_requiere_sa(df):
    if COL_REQUIERE_SA in df.columns:
        return df[COL_REQUIERE_SA].astype(str).str.strip().str.lower().isin(VALORES_SI)
    if COL_SA in df.columns:
        return ~df[COL_SA].isin(VALORES_NA) & df[COL_SA].notna()
    return pd.Series([True] * len(df), index=df.index)


def _mask_pendiente_sa(df):
    if COL_SA not in df.columns:
        return pd.Series([False] * len(df), index=df.index)
    val = df[COL_SA]
    mask_nulo      = val.isna()
    mask_str_vacio = val.astype(str).str.strip() == ""
    mask_str_cero  = val.astype(str).str.strip() == "0"
    mask_num_cero  = pd.to_numeric(val, errors="coerce").fillna(1) == 0
    return mask_nulo | mask_str_vacio | mask_str_cero | mask_num_cero


def calcular_completados(df, actividad):
    if _es_suministro(actividad):
        mask_req  = _mask_requiere_sa(df)
        df_aplic  = df[mask_req].copy()
        total     = len(df_aplic)
        if total == 0:
            return 0, 0, 0.0, 0
        pendientes  = int(_mask_pendiente_sa(df_aplic).sum())
        completados = total - pendientes
        porcentaje  = (completados / total) * 100
        return completados, pendientes, porcentaje, total

    total = len(df)
    if total == 0:
        return 0, 0, 0.0, 0
    completados = df[actividad].notna().sum()
    try:
        v = df[actividad].isin(VALORES_OK).sum()
        if v > 0:
            completados = v
    except Exception:
        pass
    pendientes = total - completados
    porcentaje = (completados / total) * 100
    return completados, pendientes, porcentaje, total

# ============================================================================
# INTERFAZ PRINCIPAL
# ============================================================================

st.title("🏭 Dashboard de Seguimiento - Equipos I&C")
st.markdown("---")

# ============================================================================
# SIDEBAR
# ============================================================================

with st.sidebar:
    st.header("⚙️ Configuración")
    tab_url, tab_archivo, tab_ejemplo = st.tabs(["🌐 URL", "📁 Archivo", "🧪 Ejemplo"])

    with tab_url:
        url_excel = st.text_input(
            "URL del archivo Excel:", value=URL_DEFECTO,
            help="URL directa del archivo Excel en la nube"
        )
        usar_url = st.checkbox("Usar URL", value=True)
        if url_excel == URL_DEFECTO:
            st.success("✅ URL del proyecto cargada")

    with tab_archivo:
        archivo_subido = st.file_uploader(
            "Sube tu archivo Excel:", type=["xlsx", "xls"],
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
# PROCESAMIENTO Y VISUALIZACIÓN
# ============================================================================

if df is not None:

    if "ITEM" not in df.columns:
        st.warning("⚠️ La columna ITEM no existe en el archivo. Se usará el conteo total de filas.")

    actividades = [
        "A Instalar", "Instalación", "Canalización/Bandeja", "Cableado",
        "Conexión Equipo", "Conexión DCS", "Marquillado Equipo", "Marquillado Cable",
        "Suiministro de Aire", "Pre-Comisionamiento"
    ]
    actividades_existentes = [col for col in actividades if col in df.columns]

    # ========================================================================
    # FILTROS DINÁMICOS
    # ========================================================================

    with st.sidebar:
        st.header("🔍 Filtros")
        filtros_activos = {}

        if not st.session_state.filtros_inicializados:
            st.session_state.filtros = {}
            st.session_state.filtros_inicializados = True

        if "AREA" in df.columns:
            areas = ["Todas"] + sorted(df["AREA"].dropna().unique().tolist())
            default_areas = st.session_state.filtros.get("AREA", ["Todas"])
            area_sel = st.multiselect("Área:", areas, default=default_areas, key="filtro_area")
            st.session_state.filtros["AREA"] = area_sel
            if "Todas" not in area_sel:
                filtros_activos["AREA"] = area_sel

        if "SISTEMA GENERAL" in df.columns:
            sistemas_gen = ["Todos"] + sorted(df["SISTEMA GENERAL"].dropna().unique().tolist())
            default_sist_gen = st.session_state.filtros.get("SISTEMA GENERAL", ["Todos"])
            sistema_gen_sel = st.multiselect("Sistema General:", sistemas_gen,
                                             default=default_sist_gen, key="filtro_sist_gen")
            st.session_state.filtros["SISTEMA GENERAL"] = sistema_gen_sel
            if "Todos" not in sistema_gen_sel:
                filtros_activos["SISTEMA GENERAL"] = sistema_gen_sel

        if "SISTEMA BMS/SMC/DCS" in df.columns:
            sistemas_bms = ["Todos"] + sorted(df["SISTEMA BMS/SMC/DCS"].dropna().unique().tolist())
            default_sist_bms = st.session_state.filtros.get("SISTEMA BMS/SMC/DCS", ["Todos"])
            sistema_bms_sel = st.multiselect("Sistema BMS/SMC/DCS:", sistemas_bms,
                                             default=default_sist_bms, key="filtro_sist_bms")
            st.session_state.filtros["SISTEMA BMS/SMC/DCS"] = sistema_bms_sel
            if "Todos" not in sistema_bms_sel:
                filtros_activos["SISTEMA BMS/SMC/DCS"] = sistema_bms_sel

        if "TIPO INSTRUMENTOS" in df.columns:
            tipos = ["Todos"] + sorted(df["TIPO INSTRUMENTOS"].dropna().unique().tolist())
            default_tipos = st.session_state.filtros.get("TIPO INSTRUMENTOS", ["Todos"])
            tipo_sel = st.multiselect("Tipo Instrumento:", tipos, default=default_tipos, key="filtro_tipo")
            st.session_state.filtros["TIPO INSTRUMENTOS"] = tipo_sel
            if "Todos" not in tipo_sel:
                filtros_activos["TIPO INSTRUMENTOS"] = tipo_sel

        if "Prioridad" in df.columns:
            prioridades = ["Todas"] + sorted(df["Prioridad"].dropna().unique().tolist())
            default_prior = st.session_state.filtros.get("Prioridad", ["Todas"])
            prioridad_sel = st.multiselect("Prioridad:", prioridades, default=default_prior, key="filtro_prior")
            st.session_state.filtros["Prioridad"] = prioridad_sel
            if "Todas" not in prioridad_sel:
                filtros_activos["Prioridad"] = prioridad_sel

        if "Hito" in df.columns:
            hitos = ["Todos"] + sorted(df["Hito"].dropna().unique().tolist())
            default_hitos = st.session_state.filtros.get("Hito", ["Todos"])
            hito_sel = st.multiselect("Hito:", hitos, default=default_hitos, key="filtro_hito")
            st.session_state.filtros["Hito"] = hito_sel
            if "Todos" not in hito_sel:
                filtros_activos["Hito"] = hito_sel

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
    # MÉTRICAS DE AVANCE
    # ========================================================================

    st.header("📊 Métricas de Avance General")

    if COL_SA in actividades_existentes:
        total_global = len(df_filtrado)
        completados_sa, pendientes_sa, porcentaje_sa, total_aplicable_sa = calcular_completados(
            df_filtrado, COL_SA
        )
        no_requiere_sa = total_global - total_aplicable_sa

        if COL_REQUIERE_SA not in df_filtrado.columns:
            st.warning(
                f"⚠️ No se encontró la columna **'{COL_REQUIERE_SA}'** en el Excel. "
                "Se usará la lógica anterior (excluir N/A). "
                "Agrega esa columna con valores **Si / No** para activar la nueva lógica."
            )

        col_info_sa1, col_info_sa2, col_info_sa3, col_info_sa4 = st.columns(4)
        with col_info_sa1:
            msg1 = f"🔧 **Suministro de Aire**  \n\n📊 {total_aplicable_sa} equipos lo requieren"
            st.info(msg1)
        with col_info_sa2:
            msg2 = f"✅ **Completados**  \n\n{completados_sa} de {total_aplicable_sa} → {porcentaje_sa:.1f}%"
            st.success(msg2)
        with col_info_sa3:
            msg3 = f"⚠️ **Pendientes**  \n\n{pendientes_sa} equipos"
            st.warning(msg3)
        with col_info_sa4:
            lbl4 = "No Requiere" if COL_REQUIERE_SA in df_filtrado.columns else "No Aplica (N/A)"
            msg4 = f"⚪ **{lbl4}**  \n\n{no_requiere_sa} equipos"
            st.info(msg4)

        st.markdown("---")

    metricas_cols = st.columns(5)
    for idx, actividad in enumerate(actividades_existentes):
        with metricas_cols[idx % 5]:
            total = len(df_filtrado)
            if total > 0:
                completados, pendientes, porcentaje, total_aplicable = calcular_completados(
                    df_filtrado, actividad
                )
                if _es_suministro(actividad):
                    no_req   = total - total_aplicable
                    st.metric(label="Suministro de Aire",
                              value=f"{completados}/{total_aplicable}",
                              delta=f"{porcentaje:.1f}%")
                    st.caption(f"✅ {total_aplicable} requieren | ⚪ {no_req} no requieren")
                else:
                    etiqueta = actividad.replace("Suiministro", "Suministro")
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
            completados, pendientes, porcentaje, total_aplicable = calcular_completados(
                df_filtrado, actividad
            )
            avance_data.append({
                "Actividad":   actividad.replace("Suiministro", "Suministro"),
                "Completados": completados,
                "Pendientes":  pendientes,
                "Porcentaje":  porcentaje
            })
    df_avance = pd.DataFrame(avance_data)

    col_graph1, col_graph2 = st.columns(2)

    with col_graph1:
        fig_barras = go.Figure()
        fig_barras.add_trace(go.Bar(
            name="Completados", x=df_avance["Actividad"], y=df_avance["Completados"],
            marker_color="#10b981", text=df_avance["Completados"],
            textposition="inside", textfont=dict(size=12, color="white", family="Arial Black")
        ))
        fig_barras.add_trace(go.Bar(
            name="Pendientes", x=df_avance["Actividad"], y=df_avance["Pendientes"],
            marker_color="#ef4444", text=df_avance["Pendientes"],
            textposition="inside", textfont=dict(size=12, color="white", family="Arial Black")
        ))
        fig_barras.update_layout(
            barmode="stack", title="Estado de Actividades",
            xaxis_title="Actividad", yaxis_title="Cantidad",
            xaxis_tickangle=-45, height=400, showlegend=True
        )
        st.plotly_chart(fig_barras, use_container_width=True)

    with col_graph2:
        fig_pct = go.Figure()
        fig_pct.add_trace(go.Bar(
            x=df_avance["Actividad"], y=df_avance["Porcentaje"],
            marker=dict(color=df_avance["Porcentaje"], colorscale="RdYlGn", cmin=0, cmax=100),
            text=[f"{p:.1f}%" for p in df_avance["Porcentaje"]],
            textposition="outside", textfont=dict(size=11, color="black", family="Arial Black")
        ))
        fig_pct.update_layout(
            title="Porcentaje de Completitud por Actividad",
            xaxis_title="Actividad", yaxis_title="% Completado",
            xaxis_tickangle=-45, height=400,
            yaxis=dict(range=[0, max(df_avance["Porcentaje"].max() * 1.1, 10)])
        )
        st.plotly_chart(fig_pct, use_container_width=True)

    col_btn_img1, col_btn_img2, _ = st.columns([1, 1, 2])
    with col_btn_img1:
        try:
            st.download_button(
                label="📸 Descargar Gráfico Barras (PNG)",
                data=fig_barras.to_image(format="png", width=1200, height=600),
                file_name=f"grafico_barras_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                mime="image/png"
            )
        except Exception:
            st.info("ℹ️ PNG requiere configuración adicional")
    with col_btn_img2:
        try:
            st.download_button(
                label="📸 Descargar Gráfico % (PNG)",
                data=fig_pct.to_image(format="png", width=1200, height=600),
                file_name=f"grafico_porcentaje_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                mime="image/png"
            )
        except Exception:
            st.info("ℹ️ PNG requiere configuración adicional")

    st.markdown("---")

    # ========================================================================
    # EQUIPOS POR ACTIVIDAD — Filtro: Pendientes / Ejecutados / Todos
    # ========================================================================

    st.header("⚙️ Equipos por Actividad")

    # ── Controles en 2 columnas ──────────────────────────────────────────────
    col_act_sel, col_estado_sel = st.columns([3, 2])

    with col_act_sel:
        actividades_seleccionadas = st.multiselect(
            "Selecciona una o más actividades:",
            actividades_existentes,
            default=[actividades_existentes[0]] if actividades_existentes else []
        )

    with col_estado_sel:
        estado_filtro = st.radio(
            "Mostrar equipos:",
            options=["⚠️ Pendientes", "✅ Ejecutados", "📋 Todos"],
            index=0,
            horizontal=True,
            help=(
                "**⚠️ Pendientes:** equipos con al menos 1 actividad sin completar.  \n"
                "**✅ Ejecutados:** equipos con TODAS las actividades seleccionadas completadas.  \n"
                "**📋 Todos:** sin filtrar por estado."
            )
        )

    if actividades_seleccionadas:
        df_work = df_filtrado.copy()

        # ── Calcular máscaras Pendiente y Ejecutado para cada actividad ──────
        for actividad in actividades_seleccionadas:
            if _es_suministro(actividad):
                mask_req  = _mask_requiere_sa(df_work)
                mask_pend = _mask_pendiente_sa(df_work)
                df_work[f"{actividad}_Pendiente"] = mask_req & mask_pend
                df_work[f"{actividad}_Ejecutado"] = mask_req & ~mask_pend
            else:
                df_work[f"{actividad}_Pendiente"] = ~df_work[actividad].isin(VALORES_OK)
                df_work[f"{actividad}_Ejecutado"] =  df_work[actividad].isin(VALORES_OK)

        masks_pend = [df_work[f"{a}_Pendiente"] for a in actividades_seleccionadas]
        masks_ejec = [df_work[f"{a}_Ejecutado"] for a in actividades_seleccionadas]

        # ── Aplicar filtro de estado ─────────────────────────────────────────
        if "Pendientes" in estado_filtro:
            # Al menos UNA actividad pendiente
            df_resultado = df_work[pd.concat(masks_pend, axis=1).any(axis=1)].copy()
            icono_estado  = "⚠️"
            label_estado  = "con Pendientes"
            estado_nombre = "pendientes"
        elif "Ejecutados" in estado_filtro:
            # TODAS las actividades seleccionadas completadas
            df_resultado = df_work[pd.concat(masks_ejec, axis=1).all(axis=1)].copy()
            icono_estado  = "✅"
            label_estado  = "Ejecutados (todas completadas)"
            estado_nombre = "ejecutados"
        else:
            # Todos — sin filtro de estado
            df_resultado = df_work.copy()
            icono_estado  = "📋"
            label_estado  = "Total (sin filtro)"
            estado_nombre = "todos"

        # ── 4 métricas simultáneas (siempre visibles) ────────────────────────
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        with col_m1:
            pct_res = (len(df_resultado) / len(df_filtrado) * 100) if len(df_filtrado) > 0 else 0
            st.metric(
                f"{icono_estado} {label_estado}",
                len(df_resultado),
                delta=f"{pct_res:.1f}% del total"
            )
        with col_m2:
            st.metric("🗂️ Actividades Seleccionadas", len(actividades_seleccionadas))
        with col_m3:
            total_con_pend = int(pd.concat(masks_pend, axis=1).any(axis=1).sum())
            pct_pend = (total_con_pend / len(df_filtrado) * 100) if len(df_filtrado) > 0 else 0
            st.metric("⚠️ Con ≥1 Pendiente", total_con_pend, delta=f"{pct_pend:.1f}%")
        with col_m4:
            total_ejec_comp = int(pd.concat(masks_ejec, axis=1).all(axis=1).sum())
            pct_ejec = (total_ejec_comp / len(df_filtrado) * 100) if len(df_filtrado) > 0 else 0
            st.metric("✅ 100% Ejecutados", total_ejec_comp, delta=f"{pct_ejec:.1f}%")

        # ── Tabla de resultados ──────────────────────────────────────────────
        if len(df_resultado) > 0:
            n_act = len(actividades_seleccionadas)
            st.subheader(
                f"{icono_estado} {label_estado} — "
                f"{len(df_resultado)} equipo{'s' if len(df_resultado) != 1 else ''} "
                f"({n_act} actividad{'es' if n_act > 1 else ''} seleccionada{'s' if n_act > 1 else ''})"
            )

            # Columnas a mostrar
            columnas_base = ["ITEM", "TAG", "DESCRIPTION", "AREA",
                             "SISTEMA GENERAL", "TIPO INSTRUMENTOS", "Prioridad"]
            tiene_sa_sel  = any(_es_suministro(a) for a in actividades_seleccionadas)
            if tiene_sa_sel and COL_REQUIERE_SA in df_resultado.columns:
                columnas_base.append(COL_REQUIERE_SA)
            columnas_base    = [c for c in columnas_base if c in df_resultado.columns]
            columnas_mostrar = columnas_base + actividades_seleccionadas
            df_display       = df_resultado[columnas_mostrar].copy()

            # Colores en las celdas de actividades
            def colorear_celda(val, actividad):
                val_str = str(val).strip()
                if val_str in [str(v) for v in VALORES_OK]:
                    return "background-color: #d1fae5; color: #065f46;"   # verde
                elif val_str in ["", "nan", "0", "Pendiente", "pendiente"]:
                    return "background-color: #fee2e2; color: #991b1b;"   # rojo
                else:
                    return "background-color: #fef9c3; color: #78350f;"   # amarillo

            try:
                styled = df_display.style
                for act in actividades_seleccionadas:
                    if act in df_display.columns:
                        styled = styled.applymap(
                            lambda v, a=act: colorear_celda(v, a), subset=[act]
                        )
                st.dataframe(styled, use_container_width=True, height=420)
            except Exception:
                st.dataframe(df_display, use_container_width=True, height=420)

            st.info(
                f"💡 **Vista activa:** {estado_filtro}  |  "
                "🟢 Completado  🔴 Pendiente / Vacío  🟡 Otro valor"
            )

            # Botones de descarga
            col_dl1, col_dl2, col_dl3 = st.columns(3)
            act_nombre = "_".join([a.replace("/", "-").replace(" ", "_")
                                   for a in actividades_seleccionadas[:3]])
            if len(actividades_seleccionadas) > 3:
                act_nombre += f"_y_{len(actividades_seleccionadas)-3}_mas"

            with col_dl1:
                excel_buf = crear_excel_descarga(df_display, "Equipos")
                st.download_button(
                    label=f"📥 Descargar {icono_estado} {estado_nombre.capitalize()} (Excel)",
                    data=excel_buf,
                    file_name=f"{estado_nombre}_{act_nombre}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            with col_dl2:
                try:
                    titulo_tabla = (f"Equipos {estado_nombre.capitalize()} - "
                                    f"{', '.join(actividades_seleccionadas[:2])}")
                    if len(actividades_seleccionadas) > 2:
                        titulo_tabla += f" y {len(actividades_seleccionadas)-2} más"
                    fig_tabla = crear_tabla_imagen(df_display, titulo=titulo_tabla, max_filas=50)
                    st.download_button(
                        label="📸 Descargar Tabla (PNG)",
                        data=fig_tabla.to_image(
                            format="png",
                            width=fig_tabla.layout.width,
                            height=fig_tabla.layout.height
                        ),
                        file_name=f"tabla_{estado_nombre}_{act_nombre}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                        mime="image/png"
                    )
                except Exception:
                    st.info("ℹ️ PNG requiere configuración adicional")

            with col_dl3:
                if len(df_resultado) > 50:
                    st.warning(
                        f"⚠️ PNG muestra las primeras 50 filas de {len(df_resultado)}. "
                        "Descarga Excel para el listado completo."
                    )

            # Resumen expandible
            with st.expander("📊 Ver Resumen por Actividad"):
                resumen_data = []
                for actividad in actividades_seleccionadas:
                    n_pend = int(df_work[df_work[f"{actividad}_Pendiente"]].shape[0])
                    n_ejec = int(df_work[df_work[f"{actividad}_Ejecutado"]].shape[0])
                    total_act = len(df_work)
                    resumen_data.append({
                        "Actividad":     actividad.replace("Suiministro", "Suministro"),
                        "⚠️ Pendientes": n_pend,
                        "✅ Ejecutados":  n_ejec,
                        "📋 Total":       total_act,
                        "% Avance":      f"{(n_ejec / total_act * 100):.1f}%" if total_act > 0 else "0%"
                    })
                df_res_act = pd.DataFrame(resumen_data)
                st.dataframe(df_res_act, use_container_width=True)
                st.download_button(
                    label="📥 Descargar Resumen (Excel)",
                    data=crear_excel_descarga(df_res_act, "Resumen"),
                    file_name=f"resumen_actividades_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        else:
            if "Pendientes" in estado_filtro:
                st.success("🎉 ¡Todos los equipos tienen las actividades seleccionadas completadas!")
            elif "Ejecutados" in estado_filtro:
                st.warning("⚠️ Ningún equipo tiene TODAS las actividades seleccionadas completadas aún.")
            else:
                st.info("ℹ️ No hay equipos en el conjunto filtrado.")

    else:
        st.info("👆 Selecciona al menos una actividad para ver los equipos")

    st.markdown("---")

    # ========================================================================
    # RESUMEN POR SISTEMA GENERAL
    # ========================================================================

    st.header("🗂️ Resumen por Sistema General")

    if "SISTEMA GENERAL" in df_filtrado.columns:

        col_fecha_existe = "FECHA ENTREGA" in df_filtrado.columns

        if not col_fecha_existe:
            st.info(
                "ℹ️ La columna **'FECHA ENTREGA'** no existe en el archivo. "
                "Se mostrará solo la cantidad de equipos por sistema."
            )

        resumen_rows = []
        for sistema, grupo in df_filtrado.groupby("SISTEMA GENERAL", sort=True):
            row = {"Sistema General": sistema, "Cantidad de Equipos": len(grupo)}
            if col_fecha_existe:
                fechas_raw = grupo["FECHA ENTREGA"].dropna().unique()
                if len(fechas_raw) == 0:
                    row["Fecha de Entrega"] = "—"
                else:
                    fechas_fmt = []
                    for f in sorted(fechas_raw,
                                    key=lambda x: pd.to_datetime(x, errors="coerce") or x):
                        fechas_fmt.append(_formatear_fecha(f))
                    row["Fecha de Entrega"] = " / ".join(fechas_fmt)
            resumen_rows.append(row)

        df_res_sis = (
            pd.DataFrame(resumen_rows)
            .sort_values("Cantidad de Equipos", ascending=False)
            .reset_index(drop=True)
        )

        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        with col_m1:
            st.metric("🏭 Total Sistemas", df_res_sis.shape[0])
        with col_m2:
            st.metric("🔩 Total Equipos", int(df_res_sis["Cantidad de Equipos"].sum()))
        with col_m3:
            if col_fecha_existe:
                sin_fecha = int((df_res_sis["Fecha de Entrega"] == "—").sum())
                st.metric("📅 Sin Fecha Asignada", sin_fecha)
            else:
                st.metric("📅 Columna FECHA ENTREGA", "No encontrada")
        with col_m4:
            mayor = df_res_sis.iloc[0]
            st.metric("📌 Sistema Mayor", mayor["Sistema General"],
                      delta=f"{int(mayor['Cantidad de Equipos'])} equipos")

        st.markdown("")

        col_g, col_t = st.columns([2, 3])

        with col_g:
            altura_graf = max(300, df_res_sis.shape[0] * 55)
            fig_sis = px.bar(
                df_res_sis.sort_values("Cantidad de Equipos", ascending=True),
                x="Cantidad de Equipos", y="Sistema General",
                orientation="h", title="Equipos por Sistema General",
                color="Cantidad de Equipos", color_continuous_scale="Blues",
                text="Cantidad de Equipos"
            )
            fig_sis.update_traces(textposition="outside")
            fig_sis.update_layout(
                height=altura_graf, yaxis_title="",
                xaxis_title="Cantidad de Equipos",
                showlegend=False, coloraxis_showscale=False
            )
            st.plotly_chart(fig_sis, use_container_width=True)
            try:
                st.download_button(
                    label="📸 Descargar Gráfico (PNG)",
                    data=fig_sis.to_image(format="png", width=900, height=altura_graf),
                    file_name=f"sistemas_generales_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                    mime="image/png"
                )
            except Exception:
                st.info("ℹ️ PNG requiere configuración adicional")

        with col_t:
            st.subheader("📋 Tabla de Sistemas")
            st.dataframe(df_res_sis, use_container_width=True,
                         height=min(450, 80 + df_res_sis.shape[0] * 38))
            excel_sis = crear_excel_descarga(df_res_sis, "Sistemas Generales")
            st.download_button(
                label="📥 Descargar Tabla (Excel .xlsx)",
                data=excel_sis,
                file_name=f"resumen_sistemas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with st.expander("🔍 Ver detalle completo de equipos por sistema"):
            sistema_sel = st.selectbox(
                "Selecciona un sistema para ver sus equipos:",
                sorted(df_filtrado["SISTEMA GENERAL"].dropna().unique().tolist()),
                key="sel_sistema_detalle"
            )
            df_det = df_filtrado[df_filtrado["SISTEMA GENERAL"] == sistema_sel].copy()
            cols_det = ["ITEM", "TAG", "TIPO INSTRUMENTOS", "AREA", "DESCRIPTION"]
            if col_fecha_existe:
                cols_det.append("FECHA ENTREGA")
            cols_det += actividades_existentes
            cols_det  = [c for c in cols_det if c in df_det.columns]
            st.markdown(f"**{len(df_det)} equipos** en el sistema *{sistema_sel}*")
            st.dataframe(df_det[cols_det], use_container_width=True, height=350)
            st.download_button(
                label=f"📥 Descargar equipos de {sistema_sel} (Excel .xlsx)",
                data=crear_excel_descarga(df_det[cols_det], f"Sistema_{sistema_sel[:20]}"),
                file_name=f"sistema_{sistema_sel.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    else:
        st.warning("⚠️ La columna 'SISTEMA GENERAL' no existe en el archivo cargado.")

    st.markdown("---")

    # ========================================================================
    # ANÁLISIS MULTIDIMENSIONAL
    # ========================================================================

    st.header("🔍 Análisis Multidimensional")

    tab1, tab2, tab3, tab4 = st.tabs(["📍 Por Área", "⚙️ Por Sistema", "🔧 Por Tipo", "🎯 Por Prioridad"])

    with tab1:
        if "AREA" in df_filtrado.columns:
            analisis_area = df_filtrado.groupby("AREA").size().reset_index(name="Cantidad")
            col_a1, col_a2 = st.columns([2, 1])
            with col_a1:
                fig_area = px.pie(analisis_area, values="Cantidad", names="AREA",
                                  title="Distribución de Equipos por Área", hole=0.3)
                fig_area.update_traces(textposition="inside", textinfo="percent+label+value")
                st.plotly_chart(fig_area, use_container_width=True)
                try:
                    st.download_button(
                        label="📸 Descargar Gráfico (PNG)",
                        data=fig_area.to_image(format="png", width=1200, height=800),
                        file_name=f"analisis_area_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                        mime="image/png"
                    )
                except Exception:
                    st.info("ℹ️ PNG requiere configuración adicional")
            with col_a2:
                st.dataframe(analisis_area.sort_values("Cantidad", ascending=False),
                             use_container_width=True, height=400)
                st.download_button(
                    label="📥 Descargar Tabla (Excel)",
                    data=crear_excel_descarga(analisis_area, "Por Área"),
                    file_name=f"tabla_por_area_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    with tab2:
        if "SISTEMA GENERAL" in df_filtrado.columns:
            analisis_sistema = df_filtrado.groupby("SISTEMA GENERAL").size().reset_index(name="Cantidad")
            fig_sistema = px.bar(
                analisis_sistema.sort_values("Cantidad", ascending=True),
                x="Cantidad", y="SISTEMA GENERAL", title="Equipos por Sistema General",
                orientation="h", color="Cantidad", color_continuous_scale="Blues", text="Cantidad"
            )
            fig_sistema.update_traces(textposition="outside")
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
                except Exception:
                    st.info("ℹ️ PNG requiere configuración adicional")
            with col_s2:
                st.download_button(
                    label="📥 Descargar Tabla (Excel)",
                    data=crear_excel_descarga(analisis_sistema, "Por Sistema"),
                    file_name=f"tabla_sistema_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            st.dataframe(analisis_sistema.sort_values("Cantidad", ascending=False),
                         use_container_width=True)

    with tab3:
        if "TIPO INSTRUMENTOS" in df_filtrado.columns:
            analisis_tipo = df_filtrado.groupby("TIPO INSTRUMENTOS").size().reset_index(name="Cantidad")
            fig_tipo = px.bar(
                analisis_tipo.sort_values("Cantidad", ascending=False),
                x="TIPO INSTRUMENTOS", y="Cantidad", title="Equipos por Tipo de Instrumento",
                color="Cantidad", color_continuous_scale="Viridis", text="Cantidad"
            )
            fig_tipo.update_layout(xaxis_tickangle=-45)
            fig_tipo.update_traces(textposition="outside")
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
                except Exception:
                    st.info("ℹ️ PNG requiere configuración adicional")
            with col_t2:
                st.download_button(
                    label="📥 Descargar Tabla (Excel)",
                    data=crear_excel_descarga(analisis_tipo, "Por Tipo"),
                    file_name=f"tabla_tipo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            st.dataframe(analisis_tipo.sort_values("Cantidad", ascending=False),
                         use_container_width=True)

    with tab4:
        if "Prioridad" in df_filtrado.columns:
            analisis_prioridad = df_filtrado.groupby("Prioridad").size().reset_index(name="Cantidad")
            col_p1, col_p2 = st.columns([2, 1])
            with col_p1:
                fig_prioridad = px.pie(
                    analisis_prioridad, values="Cantidad", names="Prioridad",
                    title="Distribución por Prioridad", color="Prioridad",
                    color_discrete_map={"Alta": "#ef4444", "Media": "#f59e0b", "Baja": "#10b981"}
                )
                fig_prioridad.update_traces(textposition="inside", textinfo="percent+label+value")
                st.plotly_chart(fig_prioridad, use_container_width=True)
                try:
                    st.download_button(
                        label="📸 Descargar Gráfico (PNG)",
                        data=fig_prioridad.to_image(format="png", width=1200, height=800),
                        file_name=f"analisis_prioridad_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                        mime="image/png"
                    )
                except Exception:
                    st.info("ℹ️ PNG requiere configuración adicional")
            with col_p2:
                st.dataframe(analisis_prioridad.sort_values("Cantidad", ascending=False),
                             use_container_width=True, height=400)
                st.download_button(
                    label="📥 Descargar Tabla (Excel)",
                    data=crear_excel_descarga(analisis_prioridad, "Por Prioridad"),
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
    if buscar_tag and "TAG" in df_mostrar.columns:
        df_mostrar = df_mostrar[df_mostrar["TAG"].str.contains(buscar_tag, case=False, na=False)]

    if not mostrar_todas_columnas:
        columnas_default = ["ITEM", "TAG", "TIPO INSTRUMENTOS", "AREA", "SISTEMA GENERAL",
                            "DESCRIPTION", "Prioridad"] + actividades_existentes
        if COL_REQUIERE_SA in df_mostrar.columns and COL_SA in columnas_default:
            idx_sa = columnas_default.index(COL_SA)
            columnas_default.insert(idx_sa, COL_REQUIERE_SA)
        columnas_default = [col for col in columnas_default if col in df_mostrar.columns]
        df_mostrar = df_mostrar[columnas_default]

    st.dataframe(df_mostrar, use_container_width=True, height=500)

    col_dl1, _ = st.columns(2)
    with col_dl1:
        st.download_button(
            label="📥 Descargar Tabla Filtrada (Excel)",
            data=crear_excel_descarga(df_mostrar, "Equipos Filtrados"),
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
        st.success("✅ **Datos cargados correctamente**")
    with col_footer3:
        st.info("📈 **Versión:** 1.3.0")
