"""
Dashboard de Seguimiento - Equipos I&C
======================================
Sistema completo de seguimiento y análisis de avance para proyectos de Instrumentación y Control.

Autor: Dashboard I&C
Fecha: Febrero 2026
Versión: 2.0.0 - Nueva función: Reporte PDF por Sistema General

USO:
streamlit run app_dashboard_ic.py

DEPENDENCIAS:
pip install streamlit pandas openpyxl plotly xlrd kaleido reportlab
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
from io import BytesIO

# Imports para generación de PDF
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors as rlc
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table,
    TableStyle, Image as RLImage, PageBreak, HRFlowable
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
import plotly.io as pio

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
        'openpyxl': 'openpyxl',
        'reportlab': 'reportlab',
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
        'Hito S': [1, 1, 2, 2, 1, 3, 3, 2],
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
        'Pre-Comisionamiento': ['Pendiente'] * 8,
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
            values=[f'**{col}**' for col in columnas],
            fill_color='#1f77b4',
            font=dict(color='white', size=11, family='Arial'),
            align='center', height=30
        ),
        cells=dict(
            values=valores,
            fill_color=[['#f0f0f0', 'white'] * len(df_display)],
            font=dict(color='black', size=10, family='Arial'),
            align='left', height=25
        )
    )])
    altura = min(800 + (len(df_display) * 25), 4000)
    ancho = max(1200, len(columnas) * 150)
    fig.update_layout(
        title=dict(
            text=f'**{titulo}** — {len(df)} equipos totales - Mostrando primeros {len(df_display)}',
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
        df = pd.read_excel(url_excel, sheet_name="Equipos I&C",
                           keep_default_na=False, na_values=[""])
        if 'ITEM' in df.columns:
            df = df[df['ITEM'].notna()].copy()
            df = df[df['ITEM'].astype(str).str.strip() != ''].copy()
        if 'Hito S' in df.columns:
            def _fmt_hito(x):
                try:
                    n = float(x)
                    if n != n:
                        return x
                    return str(int(n)) if n == int(n) else str(round(n, 4)).rstrip('0').rstrip('.')
                except (ValueError, TypeError):
                    return x
            df['Hito S'] = df['Hito S'].apply(_fmt_hito)
            def _sort_key(v):
                try:
                    return float(v)
                except:
                    return float('inf')
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
    'Marquillado Equipo': 1.0,
    'Marquillado Cable': 1.0,
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
    'Marquillado Equipo': 2.5,
    'Marquillado Cable': 2.5,
    'Suiministro de Aire/Tubing': 0.0,
    'Pre-Comisionamiento': 5.0,
}


def calcular_completados(df, actividad):
    """
    Calcula cuántos equipos tienen una actividad completada.
    Para 'Suministro de Aire', solo cuenta equipos que NO tienen N/A.
    Returns: tuple: (completados, pendientes, porcentaje, total_aplicable)
    """
    if 'suministro' in actividad.lower() or 'suiministro' in actividad.lower():
        df_aplicable = df[~df[actividad].isin(['N/A', 'NA', 'n/a', 'na', 'N/a'])].copy()
        total = len(df_aplicable)
        if total == 0:
            return 0, 0, 0, 0
        completados = df_aplicable[actividad].isin(
            ['OK', 'SI', 'Completado', 'COMPLETADO', 'ok', 'X', 'x', 1, True]
        ).sum()
        pendientes = total - completados
        porcentaje = (completados / total) * 100 if total > 0 else 0
        return completados, pendientes, porcentaje, total
    total = len(df)
    if total == 0:
        return 0, 0, 0, 0
    completados = df[actividad].notna().sum()
    try:
        valores_completados = df[actividad].isin(
            ['OK', 'SI', 'Completado', 'COMPLETADO', 'ok', 'X', 'x', 1, True]
        ).sum()
        if valores_completados > 0:
            completados = valores_completados
    except:
        pass
    pendientes = total - completados
    porcentaje = (completados / total) * 100 if total > 0 else 0
    return completados, pendientes, porcentaje, total


# ============================================================================
# FUNCIONES PARA GENERACIÓN DE REPORTE PDF
# ============================================================================

def generar_graficas_pdf(df_s, actividades_vis, calcular_completados):
    """Genera las dos gráficas de barras como bytes PNG para el PDF."""
    nombres, completados_l, pendientes_l, porcentajes_l = [], [], [], []
    for act in actividades_vis:
        c, p, pct, total = calcular_completados(df_s, act)
        nombres.append(act.replace("Suiministro", "Suministro"))
        completados_l.append(c)
        pendientes_l.append(p)
        porcentajes_l.append(round(pct, 1))

    # Gráfica 1 – Estado de Actividades (barras apiladas)
    fig1 = go.Figure()
    fig1.add_trace(go.Bar(
        name="Completados", x=nombres, y=completados_l,
        marker_color="#22c55e",
        text=completados_l, textposition="inside",
        textfont=dict(color="white", size=12, family="Arial Black"),
    ))
    fig1.add_trace(go.Bar(
        name="Pendientes", x=nombres, y=pendientes_l,
        marker_color="#ef4444",
        text=pendientes_l, textposition="inside",
        textfont=dict(color="white", size=12, family="Arial Black"),
    ))
    fig1.update_layout(
        barmode="stack",
        title=dict(text="Estado de Actividades",
                   font=dict(size=13, color="#111827", family="Arial"), x=0),
        xaxis=dict(tickangle=-35, tickfont=dict(size=10)),
        yaxis=dict(title="Cantidad"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02,
                    xanchor="right", x=1, bgcolor="rgba(0,0,0,0)"),
        plot_bgcolor="white", paper_bgcolor="white",
        height=380, width=630,
        margin=dict(l=45, r=15, t=55, b=115),
    )

    # Gráfica 2 – Porcentaje de Completitud
    bar_colors = [
        "#22c55e" if p >= 70 else ("#84cc16" if p >= 50 else "#eab308")
        for p in porcentajes_l
    ]
    fig2 = go.Figure()
    fig2.add_trace(go.Bar(
        x=nombres, y=porcentajes_l,
        marker_color=bar_colors,
        text=[f"{p}%" for p in porcentajes_l],
        textposition="outside",
        textfont=dict(size=11, color="#111827", family="Arial Black"),
    ))
    fig2.update_layout(
        title=dict(text="Porcentaje de Completitud por Actividad",
                   font=dict(size=13, color="#111827", family="Arial"), x=0),
        xaxis=dict(tickangle=-35, tickfont=dict(size=10)),
        yaxis=dict(title="% Completado", range=[0, 115]),
        plot_bgcolor="white", paper_bgcolor="white",
        height=380, width=630,
        margin=dict(l=45, r=15, t=55, b=115),
        showlegend=False,
    )

    img1 = pio.to_image(fig1, format="png", scale=2)
    img2 = pio.to_image(fig2, format="png", scale=2)
    return img1, img2


def generar_pdf_reporte(df_filtrado, sistemas_pro, actividades_existentes,
                         calcular_completados, PESOS_CON_SA, PESOS_SIN_SA):
    """Genera el reporte PDF completo, una página por Sistema General."""

    # Paleta de colores
    C_BG    = rlc.HexColor("#0f172a")
    C_CARD  = rlc.HexColor("#1e293b")
    C_BORDER= rlc.HexColor("#334155")
    C_ORANGE= rlc.HexColor("#f59e0b")
    C_GREEN = rlc.HexColor("#22c55e")
    C_RED   = rlc.HexColor("#ef4444")
    C_BLUE  = rlc.HexColor("#3b82f6")
    C_WHITE = rlc.white
    C_GRAY  = rlc.HexColor("#6b7280")
    C_GRAY2 = rlc.HexColor("#9ca3af")
    C_DK    = rlc.HexColor("#1e3a5f")

    def ps(name, **kw):
        return ParagraphStyle(name, **kw)

    MARGIN    = 1.4 * cm
    buffer    = BytesIO()
    PAGE_SIZE = landscape(A4)
    CONTENT_W = PAGE_SIZE[0] - 2 * MARGIN  # ~27.6 cm

    doc = SimpleDocTemplate(
        buffer, pagesize=PAGE_SIZE,
        rightMargin=MARGIN, leftMargin=MARGIN,
        topMargin=MARGIN, bottomMargin=MARGIN,
    )

    fecha_gen       = datetime.now().strftime("%d/%m/%Y %H:%M")
    actividades_vis = [a for a in actividades_existentes if a != "A Instalar"]

    # Determinar sistemas a incluir en el reporte
    if not sistemas_pro or sistemas_pro == ["Todos"]:
        sistemas_iter = ["TODOS LOS SISTEMAS"]
        dfs_map = {"TODOS LOS SISTEMAS": df_filtrado}
    else:
        sistemas_iter = sistemas_pro
        dfs_map = {}
        if "PRO" in df_filtrado.columns:
            for s in sistemas_pro:
                dfs_map[s] = df_filtrado[df_filtrado["PRO"] == s].copy()
        else:
            dfs_map = {s: df_filtrado for s in sistemas_pro}

    story = []

    for idx_s, sistema in enumerate(sistemas_iter):
        if idx_s > 0:
            story.append(PageBreak())

        df_s = dfs_map[sistema]

        # ---------- Calcular métricas ----------
        _sa    = next((a for a in actividades_existentes
                       if "suiministro" in a.lower() or "suministro" in a.lower()), None)
        _ts    = calcular_completados(df_s, _sa)[3] if _sa else 0
        _sa_ok = _sa is not None and _ts > 0
        _pw    = PESOS_CON_SA if _sa_ok else PESOS_SIN_SA
        _modo  = "Con SA/Tubing" if _sa_ok else "Sin SA/Tubing"

        _spxp = _tc = _tp = 0
        metricas = []
        for act in actividades_existentes:
            c, p, pct, total = calcular_completados(df_s, act)
            _tc   += c
            _tp   += p
            _spxp += _pw.get(act, 0) * pct
            metricas.append(dict(act=act, c=c, p=p, pct=round(pct, 1),
                                  total=total, total_df=len(df_s)))

        pct_gen   = round(_spxp / 100, 1)
        n_equipos = len(df_s)
        n_acts    = len(actividades_vis)

        # =====================================================================
        # SECCIÓN 1 – HEADER OSCURO
        # =====================================================================
        # Sub-tabla izquierda: % + label
        left_tbl = Table(
            [[Paragraph(f"<b>{pct_gen}%</b>",
                        ps("pp", fontSize=26, textColor=C_ORANGE, fontName="Helvetica-Bold")),
              Paragraph(
                  f"AVANCE GENERAL DEL PROYECTO<br/>"
                  f'<font size="7" color="#6b7280">Pesos: {_modo}</font>',
                  ps("pa", fontSize=9, textColor=C_GRAY2, fontName="Helvetica-Bold", leading=13)
              )]],
            colWidths=[3.9*cm, 4.1*cm],
        )
        left_tbl.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), C_BG),
            ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
            ("LEFTPADDING",(0,0), (-1,-1), 10),
            ("TOPPADDING", (0,0), (-1,-1), 6),
            ("BOTTOMPADDING",(0,0),(-1,-1), 6),
        ]))

        # Sub-tabla derecha: 4 métricas
        right_w  = CONTENT_W - 8 * cm
        m_data = [
            [Paragraph(f"<b>{_tc}</b>",
                       ps("mv1", fontSize=20, textColor=C_GREEN,  fontName="Helvetica-Bold", alignment=TA_CENTER)),
             Paragraph(f"<b>{_tp}</b>",
                       ps("mv2", fontSize=20, textColor=C_RED,    fontName="Helvetica-Bold", alignment=TA_CENTER)),
             Paragraph(f"<b>{n_acts}</b>",
                       ps("mv3", fontSize=20, textColor=C_BLUE,   fontName="Helvetica-Bold", alignment=TA_CENTER)),
             Paragraph(f"<b>{n_equipos}</b>",
                       ps("mv4", fontSize=20, textColor=C_WHITE,  fontName="Helvetica-Bold", alignment=TA_CENTER))],
            [Paragraph("Completados", ps("ml1", fontSize=8, textColor=C_GREEN,  fontName="Helvetica", alignment=TA_CENTER)),
             Paragraph("Pendientes",  ps("ml2", fontSize=8, textColor=C_RED,    fontName="Helvetica", alignment=TA_CENTER)),
             Paragraph("Actividades", ps("ml3", fontSize=8, textColor=C_BLUE,   fontName="Helvetica", alignment=TA_CENTER)),
             Paragraph("Equipos",     ps("ml4", fontSize=8, textColor=C_GRAY,   fontName="Helvetica", alignment=TA_CENTER))],
        ]
        t_m = Table(m_data, colWidths=[right_w / 4] * 4)
        t_m.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,-1), C_BG),
            ("ALIGN",         (0,0), (-1,-1), "CENTER"),
            ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
            ("TOPPADDING",    (0,0), (-1,-1), 5),
            ("BOTTOMPADDING", (0,0), (-1,-1), 4),
        ]))

        t_hdr = Table([[left_tbl, t_m]], colWidths=[8*cm, right_w])
        t_hdr.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,-1), C_BG),
            ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
            ("LEFTPADDING",   (0,0), (0,-1),  0),
            ("RIGHTPADDING",  (-1,0),(-1,-1), 10),
            ("TOPPADDING",    (0,0), (-1,-1), 0),
            ("BOTTOMPADDING", (0,0), (-1,-1), 0),
            ("LINEAFTER",     (0,0), (0,-1),  0.5, C_BORDER),
        ]))

        # Etiqueta sistema + fecha (encima del header)
        t_top = Table(
            [[Paragraph(f"<b>Sistema General: {sistema}</b>",
                        ps("psl", fontSize=9, textColor=C_GRAY2, fontName="Helvetica-Bold")),
              Paragraph(fecha_gen,
                        ps("pfec", fontSize=8, textColor=C_GRAY, fontName="Helvetica", alignment=TA_RIGHT))]],
            colWidths=[CONTENT_W * 0.65, CONTENT_W * 0.35],
        )
        t_top.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,-1), C_BG),
            ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
            ("LEFTPADDING",   (0,0), (-1,-1), 12),
            ("RIGHTPADDING",  (0,0), (-1,-1), 12),
            ("TOPPADDING",    (0,0), (-1,-1), 5),
            ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ]))

        story.append(t_top)
        story.append(t_hdr)

        # Barra de progreso
        fw  = max(0.05, CONTENT_W * (pct_gen / 100))
        ew  = CONTENT_W - fw
        tb_f = Table([[""]], colWidths=[fw], rowHeights=[0.38*cm])
        tb_e = Table([[""]], colWidths=[ew], rowHeights=[0.38*cm])
        for t_, c_ in [(tb_f, C_ORANGE), (tb_e, C_BORDER)]:
            t_.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1), c_),
                                     ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),
                                     ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0)]))
        t_bar = Table([[tb_f, tb_e]], colWidths=[fw, ew], rowHeights=[0.38*cm])
        t_bar.setStyle(TableStyle([
            ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),
            ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0),
        ]))
        story.append(t_bar)
        story.append(Spacer(1, 0.3*cm))

        # =====================================================================
        # SECCIÓN 2 – CARDS DE ACTIVIDADES
        # =====================================================================
        story.append(Paragraph(
            "&#128203;  Estado por Actividad",
            ps("sh1", fontSize=11, textColor=C_DK, fontName="Helvetica-Bold",
               spaceBefore=4, spaceAfter=4),
        ))

        acts_show = [m for m in metricas if m["act"] != "A Instalar"]
        CPR       = 5
        card_w    = CONTENT_W / CPR

        for ri in range(0, len(acts_show), CPR):
            row = acts_show[ri: ri + CPR]
            while len(row) < CPR:
                row.append(None)
            r_t, r_c, r_p = [], [], []
            for m in row:
                if m is None:
                    r_t.append(""); r_c.append(""); r_p.append("")
                else:
                    aname = m["act"].replace("Suiministro", "Suministro")
                    pc    = "#22c55e" if m["pct"] >= 70 else ("#84cc16" if m["pct"] >= 50 else "#eab308")
                    is_sa = "suministro" in m["act"].lower() or "suiministro" in m["act"].lower()
                    na_c  = m["total_df"] - m["total"]
                    extra = (f'<br/><font size="7" color="#6b7280">'
                             f'&#10003; {m["total"]} aplican | &#9711; {na_c} N/A</font>'
                             if is_sa else "")
                    r_t.append(Paragraph(aname,
                        ps(f"at{ri}{aname}", fontSize=8, textColor=C_GRAY2, fontName="Helvetica")))
                    r_c.append(Paragraph(f'<b>{m["c"]}/{m["total"]}</b>',
                        ps(f"ac{ri}{aname}", fontSize=16, textColor=C_WHITE, fontName="Helvetica-Bold")))
                    r_p.append(Paragraph(
                        f'<font color="{pc}">&#9650; {m["pct"]}%</font>{extra}',
                        ps(f"ap{ri}{aname}", fontSize=8, textColor=rlc.HexColor(pc),
                           fontName="Helvetica-Bold", leading=11)))

            tc = Table([r_t, r_c, r_p], colWidths=[card_w] * CPR)
            sc = [
                ("BACKGROUND",    (0,0), (-1,-1), C_CARD),
                ("LEFTPADDING",   (0,0), (-1,-1), 10),
                ("RIGHTPADDING",  (0,0), (-1,-1), 6),
                ("TOPPADDING",    (0,0), (CPR-1, 0), 9),
                ("TOPPADDING",    (0,1), (-1, 2), 3),
                ("BOTTOMPADDING", (0,2), (-1, 2), 9),
                ("BOTTOMPADDING", (0,0), (-1, 1), 1),
                ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
            ]
            for ci in range(CPR - 1):
                sc.append(("LINEAFTER", (ci,0), (ci,-1), 0.5, C_BORDER))
            tc.setStyle(TableStyle(sc))
            story.append(tc)
            story.append(Spacer(1, 0.15*cm))

        story.append(Spacer(1, 0.35*cm))

        # =====================================================================
        # SECCIÓN 3 – GRÁFICAS
        # =====================================================================
        story.append(Paragraph(
            "&#128202;  Progreso por Actividad",
            ps("sh2", fontSize=11, textColor=C_DK, fontName="Helvetica-Bold",
               spaceBefore=4, spaceAfter=4),
        ))
        try:
            i1b, i2b = generar_graficas_pdf(df_s, actividades_vis, calcular_completados)
            cw2  = (CONTENT_W - 0.4*cm) / 2
            img1 = RLImage(BytesIO(i1b), width=cw2, height=7.8*cm)
            img2 = RLImage(BytesIO(i2b), width=cw2, height=7.8*cm)
            t_ch = Table([[img1, img2]], colWidths=[cw2, cw2])
            t_ch.setStyle(TableStyle([
                ("ALIGN",          (0,0), (-1,-1), "CENTER"),
                ("VALIGN",         (0,0), (-1,-1), "TOP"),
                ("LEFTPADDING",    (0,0), (-1,-1), 0),
                ("RIGHTPADDING",   (0,0), (-1,-1), 0),
            ]))
            story.append(t_ch)
        except Exception as e:
            story.append(Paragraph(f"Error generando graficas: {e}",
                                    ps("er", fontSize=9, textColor=rlc.red, fontName="Helvetica")))

        # Footer
        story.append(Spacer(1, 0.2*cm))
        story.append(HRFlowable(width="100%", thickness=0.4, color=rlc.HexColor("#e5e7eb")))
        story.append(Paragraph(
            f"Dashboard I&amp;C  |  Generado: {fecha_gen}  |  Sistema: {sistema}",
            ps("ft", fontSize=7, textColor=C_GRAY, fontName="Helvetica", alignment=TA_CENTER),
        ))

    doc.build(story)
    buffer.seek(0)
    return buffer


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
        'A Instalar', 'Instalación', 'Canalización/Bandeja', 'Cableado',
        'Conexión Equipo', 'Conexión DCS', 'Marquillado Equipo', 'Marquillado Cable',
        'Suiministro de Aire/Tubing', 'Pre-Comisionamiento'
    ]
    actividades_existentes = [col for col in actividades if col in df.columns]

    # ========================================================================
    # FILTROS DINÁMICOS CON SESSION STATE
    # ========================================================================
    with st.sidebar:
        st.header("🔍 Filtros")
        _COLS_F = [
            ("Hito",               "Hito",             "Todos"),
            ("PRO",                "Sistema General",  "Todos"),
            ("AREA",               "Area",             "Todas"),
            ("SISTEMA BMS/SMC/DCS","Sistema BMS/DCS",  "Todos"),
            ("TIPO INSTRUMENTOS",  "Tipo Instrumento", "Todos"),
            ("Hito S",             "Categoria",        "Todas"),
        ]
        _CA_F = [(c, l, t) for c, l, t in _COLS_F if c in df.columns]

        if 'Hito S' in df.columns:
            _vals_hito_ok = set(df['Hito S'].dropna().unique())
            _ss_h = st.session_state.get('dyn_Hito S', ['Todas'])
            _ss_h_clean = [v for v in _ss_h if v == 'Todas' or v in _vals_hito_ok] or ['Todas']
            if _ss_h_clean != _ss_h:
                st.session_state['dyn_Hito S'] = _ss_h_clean

        def _safe_sort(vals, col=None):
            if col == 'Hito S':
                def _sk(v):
                    try: return (0, float(v))
                    except: return (1, str(v))
                return sorted(vals, key=_sk)
            try:
                return sorted(vals)
            except TypeError:
                return sorted(vals, key=lambda x: (str(type(x).__name__), str(x)))

        def _opts_f(col_obj):
            df_t = df.copy()
            for col, lbl, tod in _CA_F:
                if col == col_obj:
                    continue
                v = st.session_state.get("dyn_" + col, [tod])
                if v and tod not in v:
                    df_t = df_t[df_t[col].isin(v)]
            vals = df_t[col_obj].dropna().unique().tolist()
            return _safe_sort(vals, col=col_obj)

        for _col_f, _lbl_f, _tod_f in _CA_F:
            _opts_v = _opts_f(_col_f)
            _cur_v  = st.session_state.get("dyn_" + _col_f, [_tod_f])
            _cur_v  = [x for x in _cur_v if x == _tod_f or x in _opts_v] or [_tod_f]
            st.multiselect(_lbl_f + ":", [_tod_f] + _opts_v, default=_cur_v, key="dyn_" + _col_f)

        if st.button("Resetear Filtros", type="secondary"):
            for _col_f, _, _ in _CA_F:
                st.session_state.pop("dyn_" + _col_f, None)
            st.rerun()

    df_filtrado = df.copy()
    for _cf, _tf in [("Hito","Todos"),("PRO","Todos"),("AREA","Todas"),
                      ("SISTEMA BMS/SMC/DCS","Todos"),("TIPO INSTRUMENTOS","Todos"),
                      ("Hito S","Todas")]:
        if _cf not in df.columns:
            continue
        _vf       = st.session_state.get("dyn_" + _cf, [_tf])
        _vals_col = set(df[_cf].dropna().unique())
        _vf       = [x for x in _vf if x == _tf or x in _vals_col] or [_tf]
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
    # MÉTRICAS DE SUMINISTRO DE AIRE
    # ========================================================================
    st.header("📊 Metricas de Avance General")
    if 'Suiministro de Aire/Tubing' in actividades_existentes:
        _tg2 = len(df_filtrado)
        _cs2, _ps2, _pp2, _ta2 = calcular_completados(df_filtrado, 'Suiministro de Aire/Tubing')
        _na2 = _tg2 - _ta2
        _bb1, _bb2, _bb3, _bb4 = st.columns(4)
        with _bb1:
            st.info(f"**🔧 Suministro de Aire/Tubing**\n\n{_ta2} equipos lo requieren")
        with _bb2:
            st.success(f"**✅ Completados**\n\n{_cs2} de {_ta2} — {_pp2:.1f}%")
        with _bb3:
            st.warning(f"**⚠️ Pendientes**\n\n{_ps2} equipos")
        with _bb4:
            st.info(f"**⚪ No Aplica**\n\n{_na2} equipos")
        st.markdown("---")

    # ========================================================================
    # AVANCE GENERAL + CARDS DE ACTIVIDADES
    # ========================================================================
    if actividades_existentes and len(df_filtrado) > 0:
        _sa3    = next((a for a in actividades_existentes
                        if "suiministro" in a.lower() or "suministro" in a.lower()), None)
        _ts3    = calcular_completados(df_filtrado, _sa3)[3] if _sa3 else 0
        _sa3_ok = _sa3 is not None and _ts3 > 0
        _pw3    = PESOS_CON_SA if _sa3_ok else PESOS_SIN_SA
        _modo3  = "Con SA/Tubing" if _sa3_ok else "Sin SA/Tubing"

        _spxp3 = 0.0; _tc3 = 0; _tp3 = 0; _det3 = []
        for _act3 in actividades_existentes:
            _c3, _p3, _pct3, _ = calcular_completados(df_filtrado, _act3)
            _tc3 += _c3; _tp3 += _p3
            _pe3  = _pw3.get(_act3, 0)
            _spxp3 += _pe3 * _pct3
            _det3.append({
                "Actividad":    _act3.replace("Suiministro", "Suministro"),
                "Peso":         _pe3,
                "Avance":       round(_pct3, 1),
                "Contribucion": round(_pe3 * _pct3 / 100, 2)
            })

        _pct_gen3 = round(_spxp3 / 100, 1)
        _cb3      = "#10b981" if _pct_gen3 >= 75 else ("#f59e0b" if _pct_gen3 >= 40 else "#ef4444")
        _faltante = round(100 - _pct_gen3, 1)
        _falt_label = "Faltante"
        _n_acts3    = len([a for a in actividades_existentes if a != "A Instalar"])

        # Tarjeta de avance general (HTML/CSS)
        _prog_pct = _pct_gen3 / 100
        st.markdown(f"""
        <div style="background:#0f172a;border-radius:12px;padding:20px 24px 16px 24px;margin-bottom:12px;">
          <div style="font-size:11px;color:#6b7280;font-weight:600;letter-spacing:1px;margin-bottom:6px;">
            AVANCE GENERAL DEL PROYECTO
          </div>
          <div style="display:flex;align-items:center;gap:32px;flex-wrap:wrap;">
            <div>
              <span style="font-size:42px;font-weight:800;color:{_cb3};">{_pct_gen3}%</span><br/>
              <span style="font-size:11px;color:#6b7280;">Pesos: {_modo3}</span>
            </div>
            <div style="flex:1;min-width:200px;">
              <div style="background:#334155;border-radius:6px;height:12px;overflow:hidden;margin-bottom:12px;">
                <div style="background:{_cb3};height:12px;width:{_pct_gen3}%;border-radius:6px;"></div>
              </div>
              <div style="display:flex;gap:40px;flex-wrap:wrap;">
                <div style="text-align:center;">
                  <div style="font-size:22px;font-weight:800;color:#22c55e;">{_tc3}</div>
                  <div style="font-size:11px;color:#22c55e;">Completados</div>
                </div>
                <div style="text-align:center;">
                  <div style="font-size:22px;font-weight:800;color:#ef4444;">{_tp3}</div>
                  <div style="font-size:11px;color:#ef4444;">Pendientes</div>
                </div>
                <div style="text-align:center;">
                  <div style="font-size:22px;font-weight:800;color:#3b82f6;">{_n_acts3}</div>
                  <div style="font-size:11px;color:#3b82f6;">Actividades</div>
                </div>
                <div style="text-align:center;">
                  <div style="font-size:22px;font-weight:800;color:#ffffff;">{len(df_filtrado)}</div>
                  <div style="font-size:11px;color:#6b7280;">Equipos</div>
                </div>
              </div>
            </div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        # Expander: detalle de pesos
        with st.expander("📊 Ver detalle de pesos por actividad"):
            _df_det3 = pd.DataFrame(_det3)
            st.dataframe(_df_det3, use_container_width=True, hide_index=True)

        # Cards individuales por actividad
        _acts_cards = [a for a in actividades_existentes if a != "A Instalar"]
        _row1       = _acts_cards[:5]
        _row2       = _acts_cards[5:]

        def _render_act_card(act, df_f, col):
            c, p, pct, total = calcular_completados(df_f, act)
            color = "#10b981" if pct >= 70 else ("#f59e0b" if pct >= 40 else "#ef4444")
            arrow = "↑"
            is_sa = "suministro" in act.lower() or "suiministro" in act.lower()
            label = act.replace("Suiministro", "Suministro")
            with col:
                st.markdown(f"""
                <div style="background:#1e293b;border-radius:8px;padding:12px 16px;">
                  <div style="font-size:11px;color:#9ca3af;margin-bottom:4px;">{label}</div>
                  <div style="font-size:26px;font-weight:700;color:#ffffff;">{c}/{total}</div>
                  <div style="font-size:12px;color:{color};font-weight:600;">{arrow} {pct:.1f}%</div>
                  {"<div style='font-size:10px;color:#6b7280;margin-top:4px;'>✅ " + str(total) + " aplican | ⚪ " + str(len(df_f)-total) + " N/A</div>" if is_sa else ""}
                </div>
                """, unsafe_allow_html=True)

        cols1 = st.columns(len(_row1))
        for _act, _col in zip(_row1, cols1):
            _render_act_card(_act, df_filtrado, _col)

        if _row2:
            st.markdown("<div style='margin-top:8px'></div>", unsafe_allow_html=True)
            cols2 = st.columns(len(_row2))
            for _act, _col in zip(_row2, cols2):
                _render_act_card(_act, df_filtrado, _col)

        st.markdown("---")

        # ====================================================================
        # BOTÓN REPORTE PDF
        # ====================================================================
        st.subheader("📄 Generar Reporte PDF")

        _sis_sel   = st.session_state.get("dyn_PRO", ["Todos"])
        _sis_label = ", ".join(_sis_sel) if "Todos" not in _sis_sel else "Todos los sistemas"

        col_pdf1, col_pdf2 = st.columns([3, 1])
        with col_pdf1:
            st.info(
                f"📋 El reporte incluirá una **página por Sistema General** seleccionado: "
                f"**{_sis_label}** · Cada página contiene métricas, cards de actividades y gráficas."
            )
        with col_pdf2:
            if st.button("📄 Generar Reporte PDF", type="primary", use_container_width=True):
                with st.spinner("Generando reporte PDF, por favor espere..."):
                    try:
                        if "Todos" in _sis_sel:
                            _sistemas_pdf = (
                                sorted(df_filtrado["PRO"].dropna().unique().tolist())
                                if "PRO" in df_filtrado.columns
                                else ["TODOS LOS SISTEMAS"]
                            )
                        else:
                            _sistemas_pdf = [s for s in _sis_sel if s != "Todos"]

                        _pdf_buf = generar_pdf_reporte(
                            df_filtrado, _sistemas_pdf, actividades_existentes,
                            calcular_completados, PESOS_CON_SA, PESOS_SIN_SA,
                        )
                        _fecha_fn = datetime.now().strftime("%Y%m%d_%H%M")
                        st.download_button(
                            label="⬇️ Descargar PDF",
                            data=_pdf_buf,
                            file_name=f"Reporte_IC_{_fecha_fn}.pdf",
                            mime="application/pdf",
                            use_container_width=True,
                        )
                        st.success("✅ Reporte generado exitosamente.")
                    except Exception as _e:
                        st.error(f"❌ Error al generar el PDF: {_e}")

        st.markdown("---")

        # ====================================================================
        # SECCIÓN: PROGRESO POR ACTIVIDAD (gráficas en el dashboard)
        # ====================================================================
        st.markdown("## 📊 Progreso por Actividad")

        _acts_graf = [a for a in actividades_existentes if a != "A Instalar"]
        _names_g, _comp_g, _pend_g, _pcts_g = [], [], [], []
        for _ag in _acts_graf:
            _cg, _pg, _pctg, _tg = calcular_completados(df_filtrado, _ag)
            _names_g.append(_ag.replace("Suiministro", "Suministro"))
            _comp_g.append(_cg)
            _pend_g.append(_pg)
            _pcts_g.append(round(_pctg, 1))

        col_g1, col_g2 = st.columns(2)

        with col_g1:
            fig_stack = go.Figure()
            fig_stack.add_trace(go.Bar(
                name="Completados", x=_names_g, y=_comp_g,
                marker_color="#22c55e",
                text=_comp_g, textposition="inside",
                textfont=dict(color="white", size=11, family="Arial Black"),
            ))
            fig_stack.add_trace(go.Bar(
                name="Pendientes", x=_names_g, y=_pend_g,
                marker_color="#ef4444",
                text=_pend_g, textposition="inside",
                textfont=dict(color="white", size=11, family="Arial Black"),
            ))
            fig_stack.update_layout(
                barmode="stack",
                title="Estado de Actividades",
                xaxis=dict(tickangle=-35),
                yaxis=dict(title="Cantidad"),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                plot_bgcolor="rgba(0,0,0,0)",
                paper_bgcolor="rgba(0,0,0,0)",
                font=dict(color="#e5e7eb"),
                height=420,
            )
            st.plotly_chart(fig_stack, use_container_width=True)

        with col_g2:
            _bar_colors_g = [
                "#22c55e" if p >= 70 else ("#84cc16" if p >= 50 else "#eab308")
                for p in _pcts_g
            ]
            fig_pct = go.Figure()
            fig_pct.add_trace(go.Bar(
                x=_names_g, y=_pcts_g,
                marker_color=_bar_colors_g,
                text=[f"{p}%" for p in _pcts_g],
                textposition="outside",
                textfont=dict(size=11, family="Arial Black"),
            ))
            fig_pct.update_layout(
                title="Porcentaje de Completitud por Actividad",
                xaxis=dict(tickangle=-35),
                yaxis=dict(title="% Completado", range=[0, 115]),
                plot_bgcolor="rgba(0,0,0,0)",
                paper_bgcolor="rgba(0,0,0,0)",
                font=dict(color="#e5e7eb"),
                height=420,
                showlegend=False,
            )
            st.plotly_chart(fig_pct, use_container_width=True)

        st.markdown("---")

        # ====================================================================
        # TABLA DE EQUIPOS (detalle)
        # ====================================================================
        st.markdown("## 📋 Detalle de Equipos")
        _cols_show = [c for c in ['ITEM','TAG','TIPO INSTRUMENTOS','AREA','PRO',
                                   'SISTEMA BMS/SMC/DCS','Hito','Hito S'] if c in df_filtrado.columns]
        _cols_show += [a for a in actividades_existentes if a != 'A Instalar']
        _df_show   = df_filtrado[[c for c in _cols_show if c in df_filtrado.columns]].copy()

        st.dataframe(_df_show, use_container_width=True, height=400)

        _excel_buf = crear_excel_descarga(df_filtrado, "Equipos Filtrados")
        st.download_button(
            label="⬇️ Descargar datos filtrados (Excel)",
            data=_excel_buf,
            file_name=f"Equipos_IC_filtrado_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
