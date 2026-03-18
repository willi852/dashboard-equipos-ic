"""
Dashboard de Seguimiento - Equipos I&C
======================================
Versión: 2.3.0 - Fix: barra avance general, sin badge en cards, PDF cards mejoradas

DEPENDENCIAS:
pip install streamlit pandas openpyxl plotly xlrd kaleido reportlab
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
from io import BytesIO

try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors as rlc
    from reportlab.lib.units import cm
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table,
        TableStyle, Image as RLImage, PageBreak, HRFlowable,
    )
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
    import plotly.io as pio
    REPORTLAB_OK = True
except ImportError:
    REPORTLAB_OK = False

# ============================================================================
# CONFIGURACIÓN
# ============================================================================
st.set_page_config(
    page_title="Dashboard Equipos I&C",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
html, body, [class*="css"] { font-family: 'Inter', 'Segoe UI', sans-serif; }
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: #1e293b; }
::-webkit-scrollbar-thumb { background: #475569; border-radius: 3px; }
</style>
""", unsafe_allow_html=True)

URL_DEFECTO = "https://drive.google.com/uc?export=download&id=1x_uQhW4EKXiEgbLzZpF_InphP2oIItlu"

if 'filtros_inicializados' not in st.session_state:
    st.session_state.filtros_inicializados = False
    st.session_state.filtros = {}

# ============================================================================
# FUNCIONES AUXILIARES
# ============================================================================

def generar_excel_ejemplo():
    data = {
        'ITEM': [1,2,3,4,5,6,7,8],
        'TAG': ['FT-001','PT-002','TT-003','LT-004','FV-005','PT-006','TT-007','LT-008'],
        'TIPO INSTRUMENTOS': ['Transmisor de Flujo','Transmisor de Presión','Transmisor de Temperatura',
            'Transmisor de Nivel','Válvula de Control','Transmisor de Presión',
            'Transmisor de Temperatura','Transmisor de Nivel'],
        'AREA': ['Area 100','Area 100','Area 200','Area 200','Area 100','Area 300','Area 300','Area 200'],
        'PRO': ['Vapor','Vapor','Agua','Agua','Vapor','Combustible','Combustible','Agua'],
        'SISTEMA BMS/SMC/DCS': ['DCS','DCS','PLC','PLC','DCS','DCS','PLC','PLC'],
        'SISTEMA': ['Sistema A','Sistema A','Sistema B','Sistema B','Sistema A','Sistema C','Sistema C','Sistema B'],
        'SIGNAL ASSOCIATION': ['AI-001','AI-002','AI-003','AI-004','AO-001','AI-005','AI-006','AI-007'],
        'DESCRIPTION': ['Flujo vapor','Presión vapor','Temp agua','Nivel tanque',
                         'Control flujo','Presión combustible','Temp combustible','Nivel secundario'],
        'SIGNAL': ['4-20mA']*8, 'I/O': ['AI','AI','AI','AI','AO','AI','AI','AI'],
        'Hito': ['Hito 1','Hito 1','Hito 2','Hito 2','Hito 1','Hito 3','Hito 3','Hito 2'],
        'Hito S': [1,1,2,2,1,3,3,2],
        'Pre Emsanblado': ['OK','OK','Pendiente','OK','OK','OK','Pendiente','OK'],
        'A Instalar': ['OK']*8,
        'Instalación': ['OK','OK','Pendiente','OK','OK','OK','Pendiente','Pendiente'],
        'Canalización/Bandeja': ['OK','OK','OK','Pendiente','OK','OK','Pendiente','OK'],
        'Cableado': ['OK','Pendiente','Pendiente','Pendiente','OK','Pendiente','Pendiente','Pendiente'],
        'Conexión Equipo': ['OK','Pendiente','Pendiente','Pendiente','Pendiente','Pendiente','Pendiente','Pendiente'],
        'Conexión DCS': ['Pendiente']*8,
        'Marquillado Equipo': ['OK']*8,
        'Marquillado Cable': ['OK','Pendiente','Pendiente','Pendiente','OK','Pendiente','Pendiente','Pendiente'],
        'Suiministro de Aire/Tubing': ['N/A','N/A','N/A','N/A','OK','N/A','N/A','N/A'],
        'Pre-Comisionamiento': ['Pendiente']*8,
    }
    df = pd.DataFrame(data)
    try:
        with pd.ExcelWriter('Equipos_IC_Ejemplo.xlsx', engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Equipos I&C', index=False)
        return True, "Equipos_IC_Ejemplo.xlsx"
    except Exception as e:
        return False, str(e)


def crear_excel_descarga(df, nombre_hoja="Datos"):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=nombre_hoja, index=False)
    buffer.seek(0)
    return buffer


# ============================================================================
# FUNCIONES PRINCIPALES
# ============================================================================

@st.cache_data(ttl=300)
def cargar_datos(url_excel):
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
                    if n != n: return x
                    return str(int(n)) if n == int(n) else str(round(n,4)).rstrip('0').rstrip('.')
                except (ValueError, TypeError):
                    return x
            df['Hito S'] = df['Hito S'].apply(_fmt_hito)
            def _sort_key(v):
                try: return float(v)
                except: return float('inf')
            df['_hito_s_sort'] = df['Hito S'].apply(_sort_key)
        return df
    except Exception as e:
        st.error(f"Error al cargar el archivo: {str(e)}")
        return None


PESOS_CON_SA = {
    'A Instalar': 0,
    'Instalacion': 20.0, 'Instalación': 20.0,
    'Canalizacion/Bandeja': 30.0, 'Canalización/Bandeja': 30.0,
    'Cableado': 20.0,
    'Conexion Equipo': 2.5,  'Conexión Equipo': 2.5,
    'Conexion DCS': 2.5,     'Conexión DCS': 2.5,
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
    'Conexion DCS': 10.0,    'Conexión DCS': 10.0,
    'Marquillado Equipo': 2.5,
    'Marquillado Cable': 2.5,
    'Suiministro de Aire/Tubing': 0.0,
    'Pre-Comisionamiento': 5.0,
}


def calcular_completados(df, actividad):
    if 'suministro' in actividad.lower() or 'suiministro' in actividad.lower():
        df_ap = df[~df[actividad].isin(['N/A','NA','n/a','na','N/a'])].copy()
        total = len(df_ap)
        if total == 0: return 0, 0, 0, 0
        completados = df_ap[actividad].isin(['OK','SI','Completado','COMPLETADO','ok','X','x',1,True]).sum()
        pendientes  = total - completados
        return completados, pendientes, (completados/total)*100, total
    total = len(df)
    if total == 0: return 0, 0, 0, 0
    completados = df[actividad].notna().sum()
    try:
        vc = df[actividad].isin(['OK','SI','Completado','COMPLETADO','ok','X','x',1,True]).sum()
        if vc > 0: completados = vc
    except: pass
    pendientes = total - completados
    return completados, pendientes, (completados/total)*100, total


# ============================================================================
# PDF FUNCTIONS — SIN MODIFICAR (excepto cards mejoradas)
# ============================================================================

def _pdf_status_color(pct):
    if pct >= 75:   return rlc.HexColor("#22c55e")
    elif pct >= 50: return rlc.HexColor("#f59e0b")
    elif pct >= 25: return rlc.HexColor("#f97316")
    else:           return rlc.HexColor("#ef4444")



# ============================================================================
# PDF: ProgressBar Flowable
# ============================================================================

class ProgressBar:
    """Flowable personalizado: barra de progreso dibujada directo en canvas."""
    def __init__(self, width, height, pct, color_fill, color_bg):
        self._width      = width
        self._height     = height
        self._pct        = min(100.0, max(0.0, float(pct)))
        self._color_fill = color_fill
        self._color_bg   = color_bg

    # Convertirlo en un Flowable real en el momento de usarlo
    def as_flowable(self):
        from reportlab.platypus import Flowable as _FL
        _w, _h, _p, _cf, _cb = (
            self._width, self._height, self._pct,
            self._color_fill, self._color_bg,
        )
        class _PB(_FL):
            def wrap(self_, aw, ah):
                return (_w, _h)
            def draw(self_):
                c = self_.canv
                c.saveState()
                c.setFillColor(_cb)
                c.rect(0, 0, _w, _h, fill=1, stroke=0)
                fw = _w * _p / 100.0
                if fw > 0.5:
                    c.setFillColor(_cf)
                    c.rect(0, 0, fw, _h, fill=1, stroke=0)
                c.restoreState()
        return _PB()


# ============================================================================
# PDF: GRÁFICAS
# ============================================================================

def generar_graficas_pdf(df_s, actividades_vis, calcular_completados):
    nombres, comp_l, pend_l, pcts_l = [], [], [], []
    for act in actividades_vis:
        c, p, pct, _ = calcular_completados(df_s, act)
        nombres.append(act.replace("Suiministro", "Suministro"))
        comp_l.append(c); pend_l.append(p); pcts_l.append(round(pct, 1))

    FONT = dict(family="Arial, sans-serif")

    fig1 = go.Figure()
    fig1.add_trace(go.Bar(
        name="Completados", x=nombres, y=comp_l,
        marker_color="#22c55e", marker_line_width=0,
        text=comp_l, textposition="inside",
        textfont=dict(color="white", size=11, family="Arial Black"),
    ))
    fig1.add_trace(go.Bar(
        name="Pendientes", x=nombres, y=pend_l,
        marker_color="#ef4444", marker_line_width=0,
        text=pend_l, textposition="inside",
        textfont=dict(color="white", size=11, family="Arial Black"),
    ))
    fig1.update_layout(
        barmode="stack",
        title=dict(text="Estado de Actividades",
                   font=dict(size=12, color="#111827", family="Arial Bold"), x=0),
        xaxis=dict(tickangle=-40, tickfont=dict(size=9, color="#374151"),
                   showgrid=False, zeroline=False),
        yaxis=dict(title="Cantidad", tickfont=dict(size=9),
                   gridcolor="#f3f4f6", zeroline=False),
        legend=dict(orientation="h", yanchor="bottom", y=1.0, xanchor="right", x=1,
                    bgcolor="rgba(0,0,0,0)", font=dict(size=10)),
        plot_bgcolor="white", paper_bgcolor="white",
        height=320, width=600, margin=dict(l=40, r=10, t=45, b=105), font=FONT,
    )

    bar_colors = ["#22c55e" if p >= 70 else ("#84cc16" if p >= 50 else "#eab308")
                  for p in pcts_l]
    fig2 = go.Figure()
    fig2.add_trace(go.Bar(
        x=nombres, y=pcts_l, marker_color=bar_colors, marker_line_width=0,
        text=[f"{p}%" for p in pcts_l], textposition="outside",
        textfont=dict(size=10, color="#111827", family="Arial Black"),
    ))
    fig2.update_layout(
        title=dict(text="Porcentaje de Completitud por Actividad",
                   font=dict(size=12, color="#111827", family="Arial Bold"), x=0),
        xaxis=dict(tickangle=-40, tickfont=dict(size=9, color="#374151"),
                   showgrid=False, zeroline=False),
        yaxis=dict(title="% Completado", range=[0, 118], tickfont=dict(size=9),
                   gridcolor="#f3f4f6", zeroline=False),
        plot_bgcolor="white", paper_bgcolor="white",
        height=320, width=600, margin=dict(l=40, r=10, t=45, b=105),
        showlegend=False, font=FONT,
    )

    img1 = pio.to_image(fig1, format="png", scale=2)
    img2 = pio.to_image(fig2, format="png", scale=2)
    return img1, img2


# ============================================================================
# PDF: REPORTE COMPLETO
# ============================================================================

def _pdf_status_color(pct):
    if pct >= 75:   return rlc.HexColor("#22c55e")
    elif pct >= 50: return rlc.HexColor("#f59e0b")
    elif pct >= 25: return rlc.HexColor("#f97316")
    else:           return rlc.HexColor("#ef4444")


def _hex(rl_color):
    """Convierte HexColor de ReportLab a string #rrggbb."""
    return "#{:02x}{:02x}{:02x}".format(
        int(rl_color.red * 255),
        int(rl_color.green * 255),
        int(rl_color.blue * 255),
    )


def generar_pdf_reporte(df_filtrado, sistemas_pro, actividades_existentes,
                         calcular_completados, PESOS_CON_SA, PESOS_SIN_SA,
                         filtros_info=None):
    """Genera reporte PDF: una página landscape A4 por sistema."""

    # ── Paleta ──────────────────────────────────────────────────────────────
    C_BG     = rlc.HexColor("#0f172a")
    C_CARD   = rlc.HexColor("#1e293b")
    C_SEP    = rlc.HexColor("#334155")
    C_BAR_BG = rlc.HexColor("#0d1f35")
    C_ORANGE = rlc.HexColor("#f59e0b")
    C_GREEN  = rlc.HexColor("#22c55e")
    C_RED    = rlc.HexColor("#ef4444")
    C_BLUE   = rlc.HexColor("#3b82f6")
    C_WHITE  = rlc.white
    C_GRAY   = rlc.HexColor("#6b7280")
    C_GRAY2  = rlc.HexColor("#9ca3af")
    C_DK     = rlc.HexColor("#1e3a5f")

    def ps(name, **kw):
        return ParagraphStyle(name, **kw)

    MARGIN    = 1.4 * cm
    PAGE_SIZE = landscape(A4)
    CW        = PAGE_SIZE[0] - 2 * MARGIN   # ~762 pt

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=PAGE_SIZE,
        rightMargin=MARGIN, leftMargin=MARGIN,
        topMargin=MARGIN, bottomMargin=MARGIN,
        title="Reporte Dashboard I&C",
    )

    fecha_gen       = datetime.now().strftime("%d/%m/%Y %H:%M")
    # Construir etiqueta de filtros activos para el título
    _fi = filtros_info or {}
    _hito_raw = _fi.get("Hito", "")
    # Formatear: quitar decimales, asegurar prefijo "Hito"
    def _fmt_hito_label(val):
        parts = []
        for v in val.split(","):
            v = v.strip()
            if not v: continue
            try:
                n = float(v)
                parts.append(f"Hito {int(n)}")
            except ValueError:
                # already a string like "Hito 1"
                parts.append(v)
        return ", ".join(parts)
    _hito_lbl = _fmt_hito_label(_hito_raw) if _hito_raw else ""
    filtros_label = f"  |  {_hito_lbl}" if _hito_lbl else ""
    actividades_vis = [a for a in actividades_existentes if a != "A Instalar"]

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

        # ── Calcular métricas ────────────────────────────────────────────────
        _sa    = next((a for a in actividades_existentes
                       if "suiministro" in a.lower() or "suministro" in a.lower()), None)
        _ts    = calcular_completados(df_s, _sa)[3] if _sa else 0
        _sa_ok = _sa is not None and _ts > 0
        _pw    = PESOS_CON_SA if _sa_ok else PESOS_SIN_SA
        _modo  = "Con SA/Tubing" if _sa_ok else "Sin SA/Tubing"

        _spxp = 0
        metricas = []
        for act in actividades_existentes:
            c, p, pct, total = calcular_completados(df_s, act)
            _spxp += _pw.get(act, 0) * pct
            metricas.append(dict(act=act, c=c, p=p, pct=round(pct, 1),
                                  total=total, total_df=len(df_s)))

        # Completados/Pendientes = equipos con Pre-Comisionamiento completado
        _pre_com_pdf = next(
            (a for a in actividades_existentes if 'pre-comisionamiento' in a.lower()),
            None
        )
        if _pre_com_pdf:
            _tc, _tp, _, _ = calcular_completados(df_s, _pre_com_pdf)
        else:
            _tc, _tp = 0, len(df_s)

        pct_gen   = round(_spxp / 100, 1)
        n_equipos = len(df_s)
        n_acts    = len(actividades_vis)

        # ════════════════════════════════════════════════════════════════════
        # BLOQUE 1 — HEADER
        # ════════════════════════════════════════════════════════════════════
        # ── Título centrado principal ─────────────────────────────────────────
        titulo_linea2 = f"Sistema General: {sistema}{filtros_label}"
        tit_block = Table(
            [[Paragraph(
                "<b>DASHBOARD DE SEGUIMIENTO — EQUIPOS I&amp;C</b>",
                ps("tit1", fontSize=13, textColor=C_WHITE, fontName="Helvetica-Bold",
                   alignment=TA_CENTER, spaceAfter=2),
              )],
             [Paragraph(
                titulo_linea2,
                ps("tit2", fontSize=10, textColor=C_ORANGE, fontName="Helvetica-Bold",
                   alignment=TA_CENTER),
              )],
            ],
            colWidths=[CW],
        )
        tit_block.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,-1), C_BG),
            ("ALIGN",         (0,0), (-1,-1), "CENTER"),
            ("TOPPADDING",    (0,0), (-1,-1), 10),
            ("BOTTOMPADDING", (0,0), (-1,0),  3),
            ("BOTTOMPADDING", (0,1), (-1,1),  10),
            ("LINEBELOW",     (0,1), (-1,1),  1.5, C_ORANGE),
        ]))

        row_top = Table(
            [[Paragraph(f"<b>Sistema General: {sistema}</b>",
                        ps("pt", fontSize=8, textColor=C_GRAY2, fontName="Helvetica-Bold")),
              Paragraph(fecha_gen,
                        ps("pf", fontSize=8, textColor=C_GRAY,
                           fontName="Helvetica", alignment=TA_RIGHT))]],
            colWidths=[CW * 0.6, CW * 0.4],
        )
        row_top.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,-1), C_BG),
            ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
            ("LEFTPADDING",   (0,0), (-1,-1), 12),
            ("RIGHTPADDING",  (0,0), (-1,-1), 12),
            ("TOPPADDING",    (0,0), (-1,-1), 7),
            ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ]))

        row_label = Table(
            [[Paragraph("AVANCE GENERAL DEL PROYECTO",
                        ps("pal", fontSize=8, textColor=C_GRAY,
                           fontName="Helvetica-Bold", letterSpacing=1.0))]],
            colWidths=[CW],
        )
        row_label.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,-1), C_BG),
            ("LEFTPADDING",   (0,0), (-1,-1), 12),
            ("TOPPADDING",    (0,0), (-1,-1), 2),
            ("BOTTOMPADDING", (0,0), (-1,-1), 2),
        ]))

        # métricas 4 columnas
        RW4 = (CW - 6.2 * cm) / 4
        m_inner = Table(
            [[Paragraph(f"<b>{_tc}</b>",    ps("v1", fontSize=18, textColor=C_GREEN,  fontName="Helvetica-Bold", alignment=TA_CENTER)),
              Paragraph(f"<b>{_tp}</b>",    ps("v2", fontSize=18, textColor=C_RED,    fontName="Helvetica-Bold", alignment=TA_CENTER)),
              Paragraph(f"<b>{n_acts}</b>", ps("v3", fontSize=18, textColor=C_BLUE,   fontName="Helvetica-Bold", alignment=TA_CENTER)),
              Paragraph(f"<b>{n_equipos}</b>", ps("v4", fontSize=18, textColor=C_WHITE, fontName="Helvetica-Bold", alignment=TA_CENTER))],
             [Paragraph("Equipos 100%", ps("l1", fontSize=8, textColor=C_GREEN, fontName="Helvetica", alignment=TA_CENTER)),
              Paragraph("Pre-Com. Pend.", ps("l2", fontSize=7.5, textColor=C_RED,   fontName="Helvetica", alignment=TA_CENTER)),
              Paragraph("Actividades", ps("l3", fontSize=8, textColor=C_BLUE,  fontName="Helvetica", alignment=TA_CENTER)),
              Paragraph("Equipos",     ps("l4", fontSize=8, textColor=C_GRAY,  fontName="Helvetica", alignment=TA_CENTER))],
            ],
            colWidths=[RW4] * 4,
        )
        m_inner.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,-1), C_BG),
            ("ALIGN",         (0,0), (-1,-1), "CENTER"),
            ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
            ("TOPPADDING",    (0,0), (-1,-1), 6),
            ("BOTTOMPADDING", (0,0), (-1,-1), 6),
            ("LINEAFTER",     (0,0), (2,-1),  0.5, C_SEP),
        ]))

        pct_left = Table(
            [[Paragraph(f"<b>{pct_gen}%</b>",
                        ps("pgr", fontSize=28, textColor=C_ORANGE, fontName="Helvetica-Bold")),
              Paragraph(f'<font size="7.5" color="#6b7280">Pesos:<br/>{_modo}</font>',
                        ps("pmd", fontSize=7.5, textColor=C_GRAY, fontName="Helvetica", leading=10))]],
            colWidths=[4.0 * cm, 2.2 * cm],
        )
        pct_left.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,-1), C_BG),
            ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
            ("LEFTPADDING",   (0,0), (-1,-1), 12),
            ("TOPPADDING",    (0,0), (-1,-1), 6),
            ("BOTTOMPADDING", (0,0), (-1,-1), 8),
        ]))

        row_main = Table([[pct_left, m_inner]], colWidths=[6.2 * cm, CW - 6.2 * cm])
        row_main.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,-1), C_BG),
            ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
            ("LINEAFTER",     (0,0), (0,-1),  0.8, C_SEP),
            ("LEFTPADDING",   (0,0), (-1,-1), 0),
            ("RIGHTPADDING",  (0,0), (-1,-1), 12),
            ("TOPPADDING",    (0,0), (-1,-1), 0),
            ("BOTTOMPADDING", (0,0), (-1,-1), 0),
        ]))

        story.append(tit_block)
        story.append(Spacer(1, 0.10 * cm))
        story.append(row_top)
        story.append(row_label)
        story.append(row_main)

        # ── Barra de progreso general ────────────────────────────────────────
        fw = max(2.0, CW * (pct_gen / 100))
        ew = max(2.0, CW - fw)
        tb_f = Table([[""]], colWidths=[fw],  rowHeights=[0.36 * cm])
        tb_e = Table([[""]], colWidths=[ew],  rowHeights=[0.36 * cm])
        for t_, c_ in [(tb_f, C_ORANGE), (tb_e, C_SEP)]:
            t_.setStyle(TableStyle([
                ("BACKGROUND",    (0,0), (-1,-1), c_),
                ("LEFTPADDING",   (0,0), (-1,-1), 0),
                ("RIGHTPADDING",  (0,0), (-1,-1), 0),
                ("TOPPADDING",    (0,0), (-1,-1), 0),
                ("BOTTOMPADDING", (0,0), (-1,-1), 0),
            ]))
        t_bar = Table([[tb_f, tb_e]], colWidths=[fw, ew], rowHeights=[0.36 * cm])
        t_bar.setStyle(TableStyle([
            ("LEFTPADDING",   (0,0), (-1,-1), 0),
            ("RIGHTPADDING",  (0,0), (-1,-1), 0),
            ("TOPPADDING",    (0,0), (-1,-1), 0),
            ("BOTTOMPADDING", (0,0), (-1,-1), 0),
        ]))
        story.append(t_bar)
        story.append(Spacer(1, 0.16 * cm))

        # ════════════════════════════════════════════════════════════════════
        # BLOQUE 2 — CARDS DE ACTIVIDADES (diseño limpio 3 filas)
        # ════════════════════════════════════════════════════════════════════
        story.append(Paragraph(
            "<b>&#9632;  Estado por Actividad</b>",
            ps("sh1", fontSize=9, textColor=C_DK, fontName="Helvetica-Bold",
               spaceBefore=1, spaceAfter=4, leftIndent=2),
        ))

        acts_show = [m for m in metricas if m["act"] != "A Instalar"]
        CPR    = 5
        CARD_W = CW / CPR
        PAD_H  = 10   # padding horizontal (pt)

        # 3 filas: nombre | conteo | porcentaje (+SA note si aplica)
        RH = [0.36 * cm, 0.65 * cm, 0.60 * cm]

        for ri in range(0, len(acts_show), CPR):
            row_acts = acts_show[ri: ri + CPR]
            while len(row_acts) < CPR:
                row_acts.append(None)

            r_name, r_count, r_pct = [], [], []

            for m in row_acts:
                if m is None:
                    r_name.append(""); r_count.append(""); r_pct.append("")
                    continue

                aname    = m["act"].replace("Suiministro", "Suministro")
                sc_color = _pdf_status_color(m["pct"])
                sc_hex   = _hex(sc_color)
                is_sa    = "suministro" in m["act"].lower() or "suiministro" in m["act"].lower()
                na_c     = m["total_df"] - m["total"]

                # Fila 0 — Nombre
                r_name.append(Paragraph(
                    aname,
                    ps(f"nm{ri}_{aname[:4]}", fontSize=7.5, textColor=C_GRAY2,
                       fontName="Helvetica", leading=9),
                ))

                # Fila 1 — Conteo grande  c / total
                r_count.append(Paragraph(
                    f'<b><font color="#f1f5f9" size="17">{m["c"]}</font></b>'
                    f'<font size="11" color="#475569">/{m["total"]}</font>',
                    ps(f"ct{ri}_{aname[:4]}", fontSize=17, textColor=C_WHITE,
                       fontName="Helvetica-Bold", leading=20),
                ))

                # Fila 2 — Porcentaje  ▲ XX.X%
                r_pct.append(Paragraph(
                    f'<font color="{sc_hex}" size="9.5"><b>&#9650; {m["pct"]}%</b></font>',
                    ps(f"pc{ri}_{aname[:4]}", fontSize=9.5, textColor=C_WHITE,
                       fontName="Helvetica", leading=12),
                ))

            tc = Table(
                [r_name, r_count, r_pct],
                colWidths=[CARD_W] * CPR,
                rowHeights=RH,
            )
            sty = [
                ("BACKGROUND",    (0,0), (-1,-1), C_CARD),
                ("LEFTPADDING",   (0,0), (-1,-1), PAD_H),
                ("RIGHTPADDING",  (0,0), (-1,-1), PAD_H),
                ("TOPPADDING",    (0,0), (-1,0),  6),
                ("BOTTOMPADDING", (0,0), (-1,0),  2),
                ("TOPPADDING",    (0,1), (-1,1),  3),
                ("BOTTOMPADDING", (0,1), (-1,1),  2),
                ("TOPPADDING",    (0,2), (-1,2),  2),
                ("BOTTOMPADDING", (0,2), (-1,2),  6),
                ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
                # separadores verticales
                *[("LINEAFTER", (ci,0), (ci,-1), 0.5, C_SEP) for ci in range(CPR-1)],
            ]
            # borde superior de color por tarjeta
            for ci, m in enumerate(row_acts):
                if m is not None:
                    sty.append(("LINEABOVE", (ci,0), (ci,0), 3, _pdf_status_color(m["pct"])))

            tc.setStyle(TableStyle(sty))
            story.append(tc)
            if ri + CPR < len(acts_show):
                story.append(Spacer(1, 0.10 * cm))

        story.append(Spacer(1, 0.18 * cm))

        # ════════════════════════════════════════════════════════════════════
        # BLOQUE 3 — GRÁFICAS
        # ════════════════════════════════════════════════════════════════════
        story.append(Paragraph(
            "<b>&#9632;  Progreso por Actividad</b>",
            ps("sh2", fontSize=9, textColor=C_DK, fontName="Helvetica-Bold",
               spaceBefore=1, spaceAfter=4, leftIndent=2),
        ))

        try:
            i1b, i2b = generar_graficas_pdf(df_s, actividades_vis, calcular_completados)
            cw2  = (CW - 0.3 * cm) / 2
            CH   = 6.4 * cm
            img1 = RLImage(BytesIO(i1b), width=cw2, height=CH)
            img2 = RLImage(BytesIO(i2b), width=cw2, height=CH)
            t_ch = Table([[img1, img2]], colWidths=[cw2, cw2], rowHeights=[CH])
            t_ch.setStyle(TableStyle([
                ("ALIGN",         (0,0), (-1,-1), "CENTER"),
                ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
                ("BACKGROUND",    (0,0), (-1,-1), rlc.HexColor("#f8fafc")),
                ("LEFTPADDING",   (0,0), (-1,-1), 2),
                ("RIGHTPADDING",  (0,0), (-1,-1), 2),
                ("TOPPADDING",    (0,0), (-1,-1), 2),
                ("BOTTOMPADDING", (0,0), (-1,-1), 2),
                ("LINEAFTER",     (0,0), (0,-1),  0.5, rlc.HexColor("#e2e8f0")),
                ("BOX",           (0,0), (-1,-1),  0.5, rlc.HexColor("#e2e8f0")),
            ]))
            story.append(t_ch)
        except Exception as e:
            story.append(Paragraph(f"Error generando graficas: {e}",
                ps("er", fontSize=9, textColor=rlc.red, fontName="Helvetica")))


        # ════════════════════════════════════════════════════════════════════
        # PÁGINA 2 — PENDIENTES POR ACTIVIDAD + EQUIPOS OPERATIVOS
        # ════════════════════════════════════════════════════════════════════
        story.append(PageBreak())

        # ── Mini-header página 2 ─────────────────────────────────────────────
        tit_block2 = Table(
            [[Paragraph(
                "<b>DASHBOARD DE SEGUIMIENTO — EQUIPOS I&amp;C</b>",
                ps("tit1b", fontSize=13, textColor=C_WHITE, fontName="Helvetica-Bold",
                   alignment=TA_CENTER, spaceAfter=2),
              )],
             [Paragraph(
                f"Sistema General: {sistema}{filtros_label}  —  Pendientes por Actividad",
                ps("tit2b", fontSize=10, textColor=C_ORANGE, fontName="Helvetica-Bold",
                   alignment=TA_CENTER),
              )],
             [Paragraph(
                fecha_gen,
                ps("tit3b", fontSize=8, textColor=C_GRAY, fontName="Helvetica",
                   alignment=TA_CENTER),
              )],
            ],
            colWidths=[CW],
        )
        tit_block2.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,-1), C_BG),
            ("ALIGN",         (0,0), (-1,-1), "CENTER"),
            ("TOPPADDING",    (0,0), (-1,-1), 10),
            ("BOTTOMPADDING", (0,0), (-1,1),  3),
            ("BOTTOMPADDING", (0,2), (-1,2),  10),
            ("LINEBELOW",     (0,1), (-1,1),  1.5, C_ORANGE),
        ]))
        story.append(tit_block2)
        story.append(Spacer(1, 0.25 * cm))

        # ── Tabla: pendientes por actividad ──────────────────────────────────
        story.append(Paragraph(
            "<b>&#9632;  Pendientes por Actividad</b>",
            ps("sh3", fontSize=9, textColor=C_DK, fontName="Helvetica-Bold",
               spaceBefore=2, spaceAfter=5, leftIndent=2),
        ))

        _NA_VALS = ['N/A', 'NA', 'n/a', 'na', 'N/a']
        _OK_VALS = ['OK', 'SI', 'Completado', 'COMPLETADO', 'ok', 'X', 'x', 1, True]

        CW_ACT  = 6.5 * cm
        CW_NUM  = (CW - CW_ACT) / 5

        def _ph2(txt, clr=C_WHITE):
            return Paragraph(f"<b>{txt}</b>",
                             ps(f"ph2{txt[:6]}", fontSize=8, textColor=clr,
                                fontName="Helvetica-Bold", alignment=TA_CENTER))

        def _pc2(txt, clr=C_WHITE, align=TA_CENTER):
            return Paragraph(str(txt),
                             ps(f"pc2{txt}", fontSize=8.5, textColor=clr,
                                fontName="Helvetica", alignment=align))

        tbl2_data = [[
            _ph2("Actividad",       C_GRAY2),
            _ph2("Total Equipos",   C_WHITE),
            _ph2("No Aplica (N/A)", C_GRAY),
            _ph2("Aplican",         C_BLUE),
            _ph2("&#10003; Completados", C_GREEN),
            _ph2("&#215; Pendientes",    C_RED),
        ]]

        tbl2_style = [
            # Header
            ("BACKGROUND",    (0,0), (-1,0),  rlc.HexColor("#0f172a")),
            ("LINEBELOW",     (0,0), (-1,0),   1.5, C_ORANGE),
            # Data rows alternadas
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [C_CARD, rlc.HexColor("#162032")]),
            # Borders
            ("LINEBELOW",     (0,1), (-1,-2),  0.3, C_SEP),
            ("LINEBELOW",     (0,-1), (-1,-1), 0.8, C_SEP),
            # Padding
            ("LEFTPADDING",   (0,0), (-1,-1),  8),
            ("RIGHTPADDING",  (0,0), (-1,-1),  8),
            ("TOPPADDING",    (0,0), (-1,-1),  6),
            ("BOTTOMPADDING", (0,0), (-1,-1),  6),
            ("VALIGN",        (0,0), (-1,-1),  "MIDDLE"),
            # Separadores verticales
            ("LINEAFTER",     (0,0), (0,-1),   0.5, C_SEP),
            ("LINEAFTER",     (1,0), (1,-1),   0.3, C_SEP),
            ("LINEAFTER",     (2,0), (2,-1),   0.3, C_SEP),
            ("LINEAFTER",     (3,0), (3,-1),   0.3, C_SEP),
            ("LINEAFTER",     (4,0), (4,-1),   0.3, C_SEP),
        ]

        _acts_for_table = [m for m in acts_show if 'pre-comisionamiento' not in m['act'].lower()]
        for ri_t, m in enumerate(_acts_for_table, start=1):
            aname    = m["act"].replace("Suiministro", "Suministro")
            total_df = m["total_df"]

            # Contar NA (no aplica)
            if m["act"] in df_s.columns:
                na_count  = int(df_s[m["act"]].isin(_NA_VALS).sum())
                aplica    = total_df - na_count
                comp      = int(df_s[m["act"]].isin(_OK_VALS).sum())
                pend      = aplica - comp
            else:
                na_count, aplica, comp, pend = 0, total_df, m["c"], m["p"]

            sc_col = _pdf_status_color(round(comp / aplica * 100, 1) if aplica > 0 else 0)

            # Color de pendientes: rojo si > 0, gris si 0
            pend_col = C_RED if pend > 0 else C_GREEN

            tbl2_data.append([
                Paragraph(aname,
                          ps(f"an2{ri_t}", fontSize=8.5, textColor=C_GRAY2,
                             fontName="Helvetica", alignment=TA_LEFT)),
                _pc2(total_df),
                _pc2(na_count if na_count > 0 else "—",
                     clr=C_GRAY if na_count == 0 else rlc.HexColor("#64748b")),
                _pc2(aplica,  clr=C_WHITE),
                Paragraph(f"<b>{comp}</b>",
                          ps(f"cp2{ri_t}", fontSize=8.5, textColor=C_GREEN,
                             fontName="Helvetica-Bold", alignment=TA_CENTER)),
                Paragraph(f"<b>{pend}</b>",
                          ps(f"pp2{ri_t}", fontSize=8.5, textColor=pend_col,
                             fontName="Helvetica-Bold", alignment=TA_CENTER)),
            ])
            # Borde izquierdo de color de estado
            tbl2_style.append(("LINEBEFORECOLOR" if False else "LINEBEFORE",
                                (0, ri_t), (0, ri_t), 3, sc_col))

        tbl2 = Table(tbl2_data, colWidths=[CW_ACT] + [CW_NUM] * 5)
        tbl2.setStyle(TableStyle(tbl2_style))
        story.append(tbl2)
        story.append(Spacer(1, 0.30 * cm))

        # ════════════════════════════════════════════════════════════════════
        # BLOQUE: PENDIENTES POR TIPO DE INSTRUMENTO (matriz actividad × tipo)
        # ════════════════════════════════════════════════════════════════════
        story.append(Paragraph(
            "<b>&#9632;  Pendientes por Tipo de Instrumento</b>",
            ps("sh5", fontSize=9, textColor=C_DK, fontName="Helvetica-Bold",
               spaceBefore=2, spaceAfter=5, leftIndent=2),
        ))

        _ABREV = {
            "Instalación":              "Instal.",
            "Instalacion":              "Instal.",
            "Canalización/Bandeja":     "Canal.",
            "Canalizacion/Bandeja":     "Canal.",
            "Cableado":                 "Cableado",
            "Conexión Equipo":          "Con.Eq.",
            "Conexion Equipo":          "Con.Eq.",
            "Conexión DCS":             "Con.DCS",
            "Conexion DCS":             "Con.DCS",
            "Marquillado Equipo":       "Marq.Eq.",
            "Marquillado Cable":        "Marq.Ca.",
            "Suiministro de Aire/Tubing": "Sum.Aire",
            "Suministro de Aire/Tubing":  "Sum.Aire",
            "Pre-Comisionamiento":      "Pre-Com.",
        }

        _OK_VALS2 = ['OK', 'SI', 'Completado', 'COMPLETADO', 'ok', 'X', 'x', 1, True]
        _NA_VALS2 = ['N/A', 'NA', 'n/a', 'na', 'N/a']

        # Calcular pendientes por tipo para cada actividad
        _tipo_col = "TIPO DE INSTRUMENTO" if "TIPO DE INSTRUMENTO" in df_s.columns else None
        _pend_matrix = {}   # {tipo: {act: count}}
        _acts_with_pend = []   # actividades que tienen al menos 1 pendiente

        if _tipo_col:
            for _mt in _acts_for_table:
                _act_t = _mt["act"]
                if _act_t not in df_s.columns:
                    continue
                # Filas que aplican (no NA) y no están completadas
                _mask_na   = df_s[_act_t].isin(_NA_VALS2)
                _mask_ok   = df_s[_act_t].isin(_OK_VALS2)
                _df_pend_t = df_s[~_mask_na & ~_mask_ok].copy()
                if len(_df_pend_t) == 0:
                    continue
                _acts_with_pend.append(_act_t)
                _tipo_counts = _df_pend_t[_tipo_col].fillna("Sin tipo").value_counts()
                for _tipo, _cnt in _tipo_counts.items():
                    if _tipo not in _pend_matrix:
                        _pend_matrix[_tipo] = {}
                    _pend_matrix[_tipo][_act_t] = int(_cnt)

        if _tipo_col and _pend_matrix and _acts_with_pend:
            # Ordenar tipos por total pendiente descendente
            _tipos_sorted = sorted(
                _pend_matrix.keys(),
                key=lambda t: sum(_pend_matrix[t].values()),
                reverse=True
            )

            # Anchos de columna
            _CW_TIPO = 5.8 * cm
            _n_act_cols = len(_acts_with_pend)
            _CW_ACT_COL = (CW - _CW_TIPO) / _n_act_cols

            # Fila de encabezado
            def _ph3(txt, clr=C_WHITE, fs=7.5):
                return Paragraph(f"<b>{txt}</b>",
                                 ps(f"ph3{txt[:5]}", fontSize=fs, textColor=clr,
                                    fontName="Helvetica-Bold", alignment=TA_CENTER,
                                    leading=9))

            _hdr_row = [_ph3("Tipo de Instrumento", C_GRAY2, fs=8)]
            for _at in _acts_with_pend:
                _lbl = _ABREV.get(_at, _at[:8])
                _total_pend = sum(
                    _pend_matrix.get(_t, {}).get(_at, 0) for _t in _tipos_sorted
                )
                _hdr_row.append(_ph3(f"{_lbl}\n({_total_pend})", C_RED, fs=7))

            _tbl3_data = [_hdr_row]
            _tbl3_style = [
                ("BACKGROUND",     (0,0), (-1,0),  rlc.HexColor("#0f172a")),
                ("LINEBELOW",      (0,0), (-1,0),   1.5, C_ORANGE),
                ("ROWBACKGROUNDS", (0,1), (-1,-1), [C_CARD, rlc.HexColor("#162032")]),
                ("LINEBELOW",      (0,1), (-1,-2),  0.3, C_SEP),
                ("LINEBELOW",      (0,-1), (-1,-1), 0.8, C_SEP),
                ("LEFTPADDING",    (0,0), (-1,-1),  5),
                ("RIGHTPADDING",   (0,0), (-1,-1),  5),
                ("TOPPADDING",     (0,0), (-1,-1),  5),
                ("BOTTOMPADDING",  (0,0), (-1,-1),  5),
                ("VALIGN",         (0,0), (-1,-1),  "MIDDLE"),
                ("LINEAFTER",      (0,0), (0,-1),   0.5, C_SEP),
            ]
            # Separadores entre columnas de actividades
            for _ci in range(1, _n_act_cols):
                _tbl3_style.append(("LINEAFTER", (_ci,0), (_ci,-1), 0.3, C_SEP))

            _total_row_vals = {_at: 0 for _at in _acts_with_pend}

            for _ri3, _tipo in enumerate(_tipos_sorted, start=1):
                _row3 = [Paragraph(
                    str(_tipo),
                    ps(f"ti3{_ri3}", fontSize=8, textColor=C_GRAY2,
                       fontName="Helvetica", alignment=TA_LEFT, leading=9)
                )]
                for _at in _acts_with_pend:
                    _cnt = _pend_matrix.get(_tipo, {}).get(_at, 0)
                    _total_row_vals[_at] += _cnt
                    if _cnt > 0:
                        _cell_par = Paragraph(
                            f"<b>{_cnt}</b>",
                            ps(f"ct3{_ri3}{_at[:3]}", fontSize=8.5, textColor=C_RED,
                               fontName="Helvetica-Bold", alignment=TA_CENTER))
                        _tbl3_style.append(
                            ("BACKGROUND", (_acts_with_pend.index(_at)+1, _ri3),
                             (_acts_with_pend.index(_at)+1, _ri3),
                             rlc.HexColor("#2d1515")))
                    else:
                        _cell_par = Paragraph(
                            "—",
                            ps(f"ct3z{_ri3}{_at[:3]}", fontSize=8, textColor=C_GRAY,
                               fontName="Helvetica", alignment=TA_CENTER))
                    _row3.append(_cell_par)
                _tbl3_data.append(_row3)

            # Fila de totales
            _tot_row = [Paragraph("<b>TOTAL</b>",
                                  ps("tot3", fontSize=8, textColor=C_WHITE,
                                     fontName="Helvetica-Bold", alignment=TA_LEFT))]
            for _at in _acts_with_pend:
                _tot_row.append(Paragraph(
                    f"<b>{_total_row_vals[_at]}</b>",
                    ps(f"tot3{_at[:3]}", fontSize=8.5, textColor=C_ORANGE,
                       fontName="Helvetica-Bold", alignment=TA_CENTER)))
            _tbl3_data.append(_tot_row)
            _tbl3_style.append(("BACKGROUND",    (0,-1), (-1,-1), rlc.HexColor("#0f172a")))
            _tbl3_style.append(("LINEABOVE",     (0,-1), (-1,-1), 1.0, C_ORANGE))

            _tbl3 = Table(
                _tbl3_data,
                colWidths=[_CW_TIPO] + [_CW_ACT_COL] * _n_act_cols,
            )
            _tbl3.setStyle(TableStyle(_tbl3_style))
            story.append(_tbl3)
        else:
            story.append(Paragraph(
                "Columna TIPO DE INSTRUMENTO no disponible o sin pendientes.",
                ps("noti3", fontSize=8, textColor=C_GRAY, fontName="Helvetica"),
            ))

        story.append(Spacer(1, 0.35 * cm))

        # ── Equipos completamente operativos ─────────────────────────────────
        story.append(Paragraph(
            "<b>&#9632;  Equipos Completamente Operativos</b>",
            ps("sh4", fontSize=9, textColor=C_DK, fontName="Helvetica-Bold",
               spaceBefore=2, spaceAfter=5, leftIndent=2),
        ))

        _pre_com2 = next(
            (a for a in actividades_existentes if 'pre-comisionamiento' in a.lower()), None
        )
        if _pre_com2 and _pre_com2 in df_s.columns:
            _op_comp = int(df_s[_pre_com2].isin(_OK_VALS).sum())
            _op_total = len(df_s)
            _op_pend  = _op_total - _op_comp
            _op_pct   = round(_op_comp / _op_total * 100, 1) if _op_total > 0 else 0
            _op_col   = _pdf_status_color(_op_pct)
            _op_hex   = _hex(_op_col)

            # Barra de progreso operativos
            _op_fw = max(2.0, CW * (_op_pct / 100))
            _op_ew = max(2.0, CW - _op_fw)

            op_bar_f = Table([[""]], colWidths=[_op_fw], rowHeights=[0.30 * cm])
            op_bar_e = Table([[""]], colWidths=[_op_ew], rowHeights=[0.30 * cm])
            for _t, _c in [(op_bar_f, _op_col), (op_bar_e, C_SEP)]:
                _t.setStyle(TableStyle([
                    ("BACKGROUND",    (0,0), (-1,-1), _c),
                    ("LEFTPADDING",   (0,0), (-1,-1), 0),
                    ("RIGHTPADDING",  (0,0), (-1,-1), 0),
                    ("TOPPADDING",    (0,0), (-1,-1), 0),
                    ("BOTTOMPADDING", (0,0), (-1,-1), 0),
                ]))
            op_bar = Table([[op_bar_f, op_bar_e]], colWidths=[_op_fw, _op_ew],
                            rowHeights=[0.30 * cm])
            op_bar.setStyle(TableStyle([
                ("LEFTPADDING",   (0,0), (-1,-1), 0),
                ("RIGHTPADDING",  (0,0), (-1,-1), 0),
                ("TOPPADDING",    (0,0), (-1,-1), 0),
                ("BOTTOMPADDING", (0,0), (-1,-1), 0),
            ]))

            # Card resumen operativos — 3 columnas, 2 filas (valores + etiquetas)
            _op_cw = CW / 3
            op_inner = Table(
                [
                    [Paragraph(f"<b>{_op_comp}</b>",
                               ps("opv1", fontSize=28, textColor=_op_col,
                                  fontName="Helvetica-Bold", alignment=TA_CENTER,
                                  leading=32)),
                     Paragraph(f"<b>{_op_pct}%</b>",
                               ps("opp1", fontSize=28, textColor=_op_col,
                                  fontName="Helvetica-Bold", alignment=TA_CENTER,
                                  leading=32)),
                     Paragraph(f"<b>{_op_pend}</b>",
                               ps("opv2", fontSize=28, textColor=C_RED,
                                  fontName="Helvetica-Bold", alignment=TA_CENTER,
                                  leading=32)),
                    ],
                    [Paragraph("Equipos Operativos",
                               ps("opl1", fontSize=9, textColor=_op_col,
                                  fontName="Helvetica", alignment=TA_CENTER,
                                  leading=11)),
                     Paragraph(f"de {_op_total} equipos totales",
                               ps("opl2", fontSize=9, textColor=C_GRAY2,
                                  fontName="Helvetica", alignment=TA_CENTER,
                                  leading=11)),
                     Paragraph("Pendientes Pre-Com.",
                               ps("opl3", fontSize=9, textColor=C_RED,
                                  fontName="Helvetica", alignment=TA_CENTER,
                                  leading=11)),
                    ],
                ],
                colWidths=[_op_cw, _op_cw, _op_cw],
            )
            op_inner.setStyle(TableStyle([
                ("BACKGROUND",    (0,0), (-1,-1), C_BG),
                ("VALIGN",        (0,0), (2,0),  "BOTTOM"),
                ("VALIGN",        (0,1), (2,1),  "TOP"),
                ("ALIGN",         (0,0), (-1,-1), "CENTER"),
                ("LEFTPADDING",   (0,0), (-1,-1), 8),
                ("RIGHTPADDING",  (0,0), (-1,-1), 8),
                ("TOPPADDING",    (0,0), (2,0),  12),
                ("BOTTOMPADDING", (0,0), (2,0),  4),
                ("TOPPADDING",    (0,1), (2,1),  4),
                ("BOTTOMPADDING", (0,1), (2,1),  12),
                ("LINEAFTER",     (0,0), (0,-1), 0.8, C_SEP),
                ("LINEAFTER",     (1,0), (1,-1), 0.8, C_SEP),
            ]))
            story.append(op_inner)
            story.append(Spacer(1, 0.15 * cm))
            story.append(op_bar)
        else:
            story.append(Paragraph(
                "Columna Pre-Comisionamiento no encontrada.",
                ps("nopc", fontSize=8, textColor=C_GRAY, fontName="Helvetica"),
            ))

        # ── Footer página 2 ───────────────────────────────────────────────────
        story.append(Spacer(1, 0.20 * cm))
        story.append(HRFlowable(width="100%", thickness=0.4,
                                 color=rlc.HexColor("#334155"), spaceAfter=3))
        story.append(Paragraph(
            f"Dashboard I&amp;C  &#183;  Generado: {fecha_gen}"
            f"  &#183;  Sistema: {sistema}  &#183;  Pág. 2 de 2",
            ps("ft2", fontSize=7, textColor=C_GRAY,
               fontName="Helvetica", alignment=TA_CENTER),
        ))

    doc.build(story)
    buffer.seek(0)
    return buffer


# ============================================================================
# UI HELPERS
# ============================================================================

def _color_por_pct(pct):
    if pct >= 75:   return "#22c55e", "rgba(34,197,94,0.08)"
    elif pct >= 50: return "#f59e0b", "rgba(245,158,11,0.08)"
    elif pct >= 25: return "#f97316", "rgba(249,115,22,0.08)"
    else:           return "#ef4444", "rgba(239,68,68,0.08)"


def _render_act_card_v2(act, df_f, col, calcular_completados):
    """Card sin badge de estado — borde superior + barra de progreso + conteos."""
    c, p, pct, total = calcular_completados(df_f, act)
    label = act.replace("Suiministro", "Suministro")
    color, _ = _color_por_pct(pct)
    is_sa     = "suministro" in act.lower() or "suiministro" in act.lower()
    na_count  = len(df_f) - total

    sa_html = ""
    if is_sa:
        sa_html = (
            f"<div style='display:flex;gap:6px;margin-top:6px;flex-wrap:wrap;'>"
            f"<span style='background:rgba(34,197,94,0.15);color:#22c55e;"
            f"font-size:10px;padding:2px 7px;border-radius:4px;font-weight:600;'>"
            f"&#10003; {total} aplican</span>"
            f"<span style='background:rgba(100,116,139,0.15);color:#94a3b8;"
            f"font-size:10px;padding:2px 7px;border-radius:4px;'>"
            f"&#9675; {na_count} N/A</span>"
            f"</div>"
        )

    with col:
        st.markdown(
            f"<div style='"
            f"background:#1e293b;"
            f"border-radius:12px;"
            f"padding:16px 18px 14px 18px;"
            f"border:1px solid #334155;"
            f"border-top:3px solid {color};'>"
            # Nombre actividad
            f"<div style='font-size:11px;color:#94a3b8;font-weight:600;"
            f"letter-spacing:0.4px;margin-bottom:10px;text-transform:uppercase;'>"
            f"{label}</div>"
            # Número grande
            f"<div style='font-size:30px;font-weight:800;color:#f1f5f9;"
            f"line-height:1;margin-bottom:10px;'>"
            f"{c}<span style='font-size:16px;color:#64748b;font-weight:500;'>/{total}</span></div>"
            # Barra de progreso
            f"<div style='background:#0f172a;border-radius:4px;height:6px;"
            f"overflow:hidden;margin-bottom:8px;'>"
            f"<div style='background:{color};height:6px;width:{pct:.1f}%;border-radius:4px;'></div>"
            f"</div>"
            # Porcentaje + completados / pendientes
            f"<div style='display:flex;justify-content:space-between;align-items:center;'>"
            f"<span style='font-size:18px;font-weight:800;color:{color};'>{pct:.1f}%</span>"
            f"<div style='display:flex;gap:10px;'>"
            f"<span style='font-size:11px;color:#22c55e;font-weight:600;'>&#10003; {c}</span>"
            f"<span style='font-size:11px;color:#ef4444;font-weight:600;'>&#215; {p}</span>"
            f"</div></div>"
            f"{sa_html}"
            f"</div>",
            unsafe_allow_html=True,
        )


def _render_metric_sa(ta, cs, ps_count, pp, na, col1, col2, col3, col4):
    cards = [
        (col1, "&#128295;", "Suministro de Aire/Tubing", str(ta),
         "equipos lo requieren", "#3b82f6", "#1d4ed8"),
        (col2, "&#9989;", "Completados", str(cs),
         f"de {ta} &nbsp;&middot;&nbsp; {pp:.1f}%", "#22c55e", "#15803d"),
        (col3, "&#9888;&#65039;", "Pendientes", str(ps_count),
         "equipos por completar", "#f59e0b", "#a16207"),
        (col4, "&#9898;", "No Aplica", str(na),
         "equipos excluidos", "#64748b", "#334155"),
    ]
    for col, icon, title, value, subtitle, color, border in cards:
        with col:
            st.markdown(
                f"<div style='background:rgba(30,41,59,0.6);"
                f"border:1px solid {border}55;"
                f"border-left:4px solid {color};"
                f"border-radius:10px;padding:16px 18px;'>"
                f"<div style='display:flex;align-items:center;gap:8px;margin-bottom:8px;'>"
                f"<span style='font-size:16px;'>{icon}</span>"
                f"<span style='font-size:12px;color:#94a3b8;font-weight:600;"
                f"letter-spacing:0.3px;'>{title}</span></div>"
                f"<div style='font-size:28px;font-weight:800;color:{color};"
                f"line-height:1;margin-bottom:4px;'>{value}</div>"
                f"<div style='font-size:12px;color:#64748b;'>{subtitle}</div>"
                f"</div>",
                unsafe_allow_html=True,
            )


# ============================================================================
# INTERFAZ PRINCIPAL
# ============================================================================

st.title("&#127981; Dashboard de Seguimiento - Equipos I&C")
st.markdown("---")

with st.sidebar:
    st.header("&#9881;&#65039; Configuración")
    tab_url, tab_archivo, tab_ejemplo = st.tabs(["&#127760; URL", "&#128193; Archivo", "&#129514; Ejemplo"])
    with tab_url:
        url_excel = st.text_input("URL del archivo Excel:", value=URL_DEFECTO)
        usar_url  = st.checkbox("Usar URL", value=True)
        if url_excel == URL_DEFECTO:
            st.success("&#10003; URL del proyecto cargada")
    with tab_archivo:
        archivo_subido = st.file_uploader("Sube tu archivo Excel:", type=['xlsx','xls'])
    with tab_ejemplo:
        st.info("&#128161; Genera un archivo Excel de ejemplo para probar la aplicación")
        if st.button("&#128202; Generar Archivo de Ejemplo", type="secondary"):
            exito, mensaje = generar_excel_ejemplo()
            if exito:
                st.success(f"&#10003; Archivo generado: {mensaje}")
            else:
                st.error(f"&#10060; Error: {mensaje}")
    if st.button("&#128260; Cargar/Actualizar Datos", type="primary"):
        st.cache_data.clear()
        st.success("&#10003; Datos actualizados")
    st.markdown("---")

df = None
if archivo_subido is not None:
    df = cargar_datos(archivo_subido)
    fuente_datos = f"&#128193; Archivo: {archivo_subido.name}"
elif usar_url and url_excel:
    df = cargar_datos(url_excel)
    fuente_datos = "&#127760; Google Drive" if "drive.google.com" in url_excel else "&#127760; URL en la nube"
else:
    st.info("&#128072; Por favor, selecciona una fuente de datos en el panel lateral")

if df is not None:
    if 'ITEM' not in df.columns:
        st.warning("&#9888; La columna ITEM no existe en el archivo.")

    actividades = [
        'A Instalar','Instalación','Canalización/Bandeja','Cableado',
        'Conexión Equipo','Conexión DCS','Marquillado Equipo','Marquillado Cable',
        'Suiministro de Aire/Tubing','Pre-Comisionamiento'
    ]
    actividades_existentes = [col for col in actividades if col in df.columns]

    with st.sidebar:
        st.header("&#128269; Filtros")
        _COLS_F = [
            ("Hito","Hito","Todos"),("PRO","Sistema General","Todos"),
            ("AREA","Area","Todas"),("SISTEMA BMS/SMC/DCS","Sistema BMS/DCS","Todos"),
            ("TIPO INSTRUMENTOS","Tipo Instrumento","Todos"),("Hito S","Categoria","Todas"),
        ]
        _CA_F = [(c,l,t) for c,l,t in _COLS_F if c in df.columns]
        if 'Hito S' in df.columns:
            _vals_hito_ok = set(df['Hito S'].dropna().unique())
            _ss_h = st.session_state.get('dyn_Hito S', ['Todas'])
            _ss_h_clean = [v for v in _ss_h if v=='Todas' or v in _vals_hito_ok] or ['Todas']
            if _ss_h_clean != _ss_h:
                st.session_state['dyn_Hito S'] = _ss_h_clean

        def _safe_sort(vals, col=None):
            if col == 'Hito S':
                def _sk(v):
                    try: return (0, float(v))
                    except: return (1, str(v))
                return sorted(vals, key=_sk)
            try: return sorted(vals)
            except TypeError: return sorted(vals, key=lambda x: (str(type(x).__name__), str(x)))

        def _opts_f(col_obj):
            df_t = df.copy()
            for col, lbl, tod in _CA_F:
                if col == col_obj: continue
                v = st.session_state.get("dyn_"+col, [tod])
                if v and tod not in v:
                    df_t = df_t[df_t[col].isin(v)]
            return _safe_sort(df_t[col_obj].dropna().unique().tolist(), col=col_obj)

        for _col_f, _lbl_f, _tod_f in _CA_F:
            _opts_v = _opts_f(_col_f)
            _cur_v  = st.session_state.get("dyn_"+_col_f, [_tod_f])
            _cur_v  = [x for x in _cur_v if x==_tod_f or x in _opts_v] or [_tod_f]
            st.multiselect(_lbl_f+":", [_tod_f]+_opts_v, default=_cur_v, key="dyn_"+_col_f)

        if st.button("Resetear Filtros", type="secondary"):
            for _col_f, _, _ in _CA_F:
                st.session_state.pop("dyn_"+_col_f, None)
            st.rerun()

    df_filtrado = df.copy()
    for _cf, _tf in [("Hito","Todos"),("PRO","Todos"),("AREA","Todas"),
                      ("SISTEMA BMS/SMC/DCS","Todos"),("TIPO INSTRUMENTOS","Todos"),
                      ("Hito S","Todas")]:
        if _cf not in df.columns: continue
        _vf       = st.session_state.get("dyn_"+_cf, [_tf])
        _vals_col = set(df[_cf].dropna().unique())
        _vf       = [x for x in _vf if x==_tf or x in _vals_col] or [_tf]
        if _vf and _tf not in _vf:
            df_filtrado = df_filtrado[df_filtrado[_cf].isin(_vf)]

    filtros_activos = {}

    # INFO GENERAL
    col_i1, col_i2, col_i3, col_i4 = st.columns(4)
    with col_i1: st.info(f"&#128203; **Total Equipos:** {len(df)}")
    with col_i2: st.info(f"&#128269; **Equipos Filtrados:** {len(df_filtrado)}")
    with col_i3: st.success(f"&#10003; **Filtros Activos:** {len(filtros_activos)}")
    with col_i4: st.info(f"&#128202; **Fuente:** {fuente_datos}")
    st.markdown("---")

    # SUMINISTRO AIRE
    st.header("&#128202; Métricas de Avance General")
    if 'Suiministro de Aire/Tubing' in actividades_existentes:
        _tg2 = len(df_filtrado)
        _cs2, _ps2, _pp2, _ta2 = calcular_completados(df_filtrado, 'Suiministro de Aire/Tubing')
        _na2 = _tg2 - _ta2
        _bb1, _bb2, _bb3, _bb4 = st.columns(4)
        _render_metric_sa(_ta2, _cs2, _ps2, _pp2, _na2, _bb1, _bb2, _bb3, _bb4)
        st.markdown("<div style='margin-top:8px'></div>", unsafe_allow_html=True)

    if actividades_existentes and len(df_filtrado) > 0:
        _sa3    = next((a for a in actividades_existentes
                        if "suiministro" in a.lower() or "suministro" in a.lower()), None)
        _ts3    = calcular_completados(df_filtrado, _sa3)[3] if _sa3 else 0
        _sa3_ok = _sa3 is not None and _ts3 > 0
        _pw3    = PESOS_CON_SA if _sa3_ok else PESOS_SIN_SA
        _modo3  = "Con SA/Tubing" if _sa3_ok else "Sin SA/Tubing"

        _spxp3 = 0.0; _det3 = []
        for _act3 in actividades_existentes:
            _c3, _p3, _pct3, _ = calcular_completados(df_filtrado, _act3)
            _pe3 = _pw3.get(_act3, 0)
            _spxp3 += _pe3 * _pct3
            _det3.append({"Actividad": _act3.replace("Suiministro","Suministro"),
                           "Peso": _pe3, "Avance": round(_pct3,1),
                           "Contribucion": round(_pe3*_pct3/100, 2)})

        # Completados/Pendientes = equipos con Pre-Comisionamiento completado
        # Representa cuántos equipos están al 100% del flujo completo
        _pre_com_col = next(
            (a for a in actividades_existentes if 'pre-comisionamiento' in a.lower()),
            None
        )
        if _pre_com_col:
            _tc3, _tp3, _, _ = calcular_completados(df_filtrado, _pre_com_col)
        else:
            _tc3 = 0
            _tp3 = len(df_filtrado)

        _pct_gen3 = round(_spxp3/100, 1)
        _cb3      = "#22c55e" if _pct_gen3>=75 else ("#f59e0b" if _pct_gen3>=40 else "#ef4444")
        _n_acts3  = len([a for a in actividades_existentes if a != "A Instalar"])
        _pct_str  = f"{_pct_gen3:.1f}" if '.' in str(_pct_gen3) else str(_pct_gen3)

        # AVANCE GENERAL — HTML sin grid, sin comentarios, sin position:absolute
        st.markdown(
            f"<div style='"
            f"background:linear-gradient(135deg,#0f172a 0%,#1e293b 100%);"
            f"border-radius:14px;padding:22px 28px 20px 28px;"
            f"margin:8px 0 12px 0;border:1px solid #334155;'>"
            # label
            f"<div style='font-size:10px;color:#475569;font-weight:700;"
            f"letter-spacing:1.5px;margin-bottom:10px;'>AVANCE GENERAL DEL PROYECTO</div>"
            # flex container
            f"<div style='display:flex;align-items:center;gap:32px;flex-wrap:wrap;'>"
            # % principal
            f"<div>"
            f"<div style='font-size:52px;font-weight:900;color:{_cb3};"
            f"line-height:1;letter-spacing:-2px;'>{_pct_gen3}%</div>"
            f"<div style='font-size:11px;color:#475569;margin-top:4px;'>"
            f"Pesos: <b style='color:#64748b;'>{_modo3}</b></div>"
            f"</div>"
            # barra + 4 métricas
            f"<div style='flex:1;min-width:260px;'>"
            # barra
            f"<div style='background:#0f172a;border-radius:6px;height:8px;"
            f"overflow:hidden;margin-bottom:18px;border:1px solid #334155;'>"
            f"<div style='background:{_cb3};height:100%;width:{_pct_gen3}%;border-radius:6px;'>"
            f"</div></div>"
            # 4 métricas con flex
            f"<div style='display:flex;justify-content:space-around;flex-wrap:wrap;'>"
            f"<div style='text-align:center;padding:10px 20px;border-right:1px solid #334155;'>"
            f"<div style='font-size:26px;font-weight:800;color:#22c55e;line-height:1;'>{_tc3}</div>"
            f"<div style='font-size:11px;color:#22c55e;margin-top:3px;font-weight:500;'>Equipos 100%</div>"
            f"</div>"
            f"<div style='text-align:center;padding:10px 20px;border-right:1px solid #334155;'>"
            f"<div style='font-size:26px;font-weight:800;color:#ef4444;line-height:1;'>{_tp3}</div>"
            f"<div style='font-size:11px;color:#ef4444;margin-top:3px;font-weight:500;'>Pre-Com. Pend.</div>"
            f"</div>"
            f"<div style='text-align:center;padding:10px 20px;border-right:1px solid #334155;'>"
            f"<div style='font-size:26px;font-weight:800;color:#3b82f6;line-height:1;'>{_n_acts3}</div>"
            f"<div style='font-size:11px;color:#3b82f6;margin-top:3px;font-weight:500;'>Actividades</div>"
            f"</div>"
            f"<div style='text-align:center;padding:10px 20px;'>"
            f"<div style='font-size:26px;font-weight:800;color:#f1f5f9;line-height:1;'>{len(df_filtrado)}</div>"
            f"<div style='font-size:11px;color:#64748b;margin-top:3px;font-weight:500;'>Equipos</div>"
            f"</div>"
            f"</div></div></div></div>",
            unsafe_allow_html=True,
        )

        with st.expander("&#128202; Ver detalle de pesos por actividad"):
            st.dataframe(pd.DataFrame(_det3), use_container_width=True, hide_index=True)

        # CARDS ACTIVIDADES
        st.markdown(
            "<div style='font-size:13px;color:#94a3b8;font-weight:600;"
            "letter-spacing:0.5px;margin:16px 0 10px 2px;'>&#9642; ESTADO POR ACTIVIDAD</div>",
            unsafe_allow_html=True,
        )

        _acts_cards = [a for a in actividades_existentes if a != "A Instalar"]
        _row1 = _acts_cards[:5]
        _row2 = _acts_cards[5:]

        cols1 = st.columns(len(_row1))
        for _act, _col in zip(_row1, cols1):
            _render_act_card_v2(_act, df_filtrado, _col, calcular_completados)

        if _row2:
            st.markdown("<div style='margin-top:10px'></div>", unsafe_allow_html=True)
            n2 = len(_row2)
            if n2 < 5:
                _pad = (5 - n2) // 2
                _cols_all = st.columns(5)
                for i, _act in enumerate(_row2):
                    _render_act_card_v2(_act, df_filtrado, _cols_all[_pad + i], calcular_completados)
            else:
                cols2 = st.columns(n2)
                for _act, _col in zip(_row2, cols2):
                    _render_act_card_v2(_act, df_filtrado, _col, calcular_completados)

        st.markdown("<div style='margin-top:16px'></div>", unsafe_allow_html=True)
        st.markdown("---")

        # BOTÓN PDF
        st.subheader("&#128196; Generar Reporte PDF")
        _sis_sel   = st.session_state.get("dyn_PRO", ["Todos"])
        _sis_label = ", ".join(_sis_sel) if "Todos" not in _sis_sel else "Todos los sistemas"
        col_pdf1, col_pdf2 = st.columns([3, 1])
        with col_pdf1:
            st.info(f"&#128203; El reporte incluirá **una página por Sistema General**: **{_sis_label}**.")
        with col_pdf2:
            if st.button("&#128196; Generar Reporte PDF", type="primary", use_container_width=True):
                if not REPORTLAB_OK:
                    st.error("&#10060; Falta instalar reportlab: pip install reportlab")
                else:
                    with st.spinner("Generando reporte PDF..."):
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
                                filtros_info={
                                    "Hito": ", ".join([str(v) for v in st.session_state.get("dyn_Hito", ["Todos"]) if str(v) != "Todos"]),
                                },
                            )
                            _fecha_fn = datetime.now().strftime("%Y%m%d_%H%M")
                            st.download_button(
                                label="&#11015;&#65039; Descargar PDF",
                                data=_pdf_buf,
                                file_name=f"Reporte_IC_{_fecha_fn}.pdf",
                                mime="application/pdf",
                                use_container_width=True,
                            )
                            st.success("&#10003; Reporte generado exitosamente.")
                        except Exception as _e:
                            st.error(f"&#10060; Error al generar el PDF: {_e}")

        st.markdown("---")

        # GRÁFICAS
        st.markdown("## &#128202; Progreso por Actividad")
        _acts_g = [a for a in actividades_existentes if a != "A Instalar"]
        _nm_g, _cp_g, _pd_g, _pt_g = [], [], [], []
        for _ag in _acts_g:
            _cg, _pg, _pctg, _ = calcular_completados(df_filtrado, _ag)
            _nm_g.append(_ag.replace("Suiministro","Suministro"))
            _cp_g.append(_cg); _pd_g.append(_pg); _pt_g.append(round(_pctg,1))

        col_g1, col_g2 = st.columns(2)
        with col_g1:
            fig_s = go.Figure()
            fig_s.add_trace(go.Bar(name="Completados", x=_nm_g, y=_cp_g,
                marker_color="#22c55e", marker_line_width=0,
                text=_cp_g, textposition="inside",
                textfont=dict(color="white",size=11,family="Arial Black")))
            fig_s.add_trace(go.Bar(name="Pendientes", x=_nm_g, y=_pd_g,
                marker_color="#ef4444", marker_line_width=0,
                text=_pd_g, textposition="inside",
                textfont=dict(color="white",size=11,family="Arial Black")))
            fig_s.update_layout(
                barmode="stack", title="Estado de Actividades",
                xaxis=dict(tickangle=-35,showgrid=False),
                yaxis=dict(title="Cantidad",gridcolor="#1e293b"),
                legend=dict(orientation="h",yanchor="bottom",y=1.02,
                            xanchor="right",x=1,bgcolor="rgba(0,0,0,0)"),
                plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                font=dict(color="#e2e8f0"), height=420)
            st.plotly_chart(fig_s, use_container_width=True)

        with col_g2:
            _bc = ["#22c55e" if p>=70 else ("#84cc16" if p>=50 else "#eab308") for p in _pt_g]
            fig_p = go.Figure()
            fig_p.add_trace(go.Bar(x=_nm_g, y=_pt_g, marker_color=_bc, marker_line_width=0,
                text=[f"{p}%" for p in _pt_g], textposition="outside",
                textfont=dict(size=11,family="Arial Black")))
            fig_p.update_layout(
                title="Porcentaje de Completitud por Actividad",
                xaxis=dict(tickangle=-35,showgrid=False),
                yaxis=dict(title="% Completado",range=[0,115],gridcolor="#1e293b"),
                plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                font=dict(color="#e2e8f0"), height=420, showlegend=False)
            st.plotly_chart(fig_p, use_container_width=True)

        st.markdown("---")

        # TABLA DETALLE
        st.markdown("## &#128203; Detalle de Equipos")
        _cols_s = [c for c in ['ITEM','TAG','TIPO INSTRUMENTOS','AREA','PRO',
                                 'SISTEMA BMS/SMC/DCS','Hito','Hito S'] if c in df_filtrado.columns]
        _cols_s += [a for a in actividades_existentes if a != 'A Instalar']
        st.dataframe(df_filtrado[[c for c in _cols_s if c in df_filtrado.columns]],
                     use_container_width=True, height=400)

        _excel_buf = crear_excel_descarga(df_filtrado, "Equipos Filtrados")
        st.download_button(
            label="&#11015;&#65039; Descargar datos filtrados (Excel)",
            data=_excel_buf,
            file_name=f"Equipos_IC_filtrado_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
