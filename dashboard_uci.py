# -*- coding: utf-8 -*-
"""
Dashboard UCI Adultos - Panel de Indicadores
Hospital Infantil Universitario de San Jose
"""

import dash
from dash import dcc, html, Input, Output, State
import dash_bootstrap_components as dbc
import plotly.graph_objects as go
import numpy as np
import pandas as pd
import json
import colorsys
from io import BytesIO
import logging
from typing import Any
from pathlib import Path
from time import time

# Imports del paquete modular (paquete dashboard_uci/, no este script)
from dashboard_uci import (
    BASE_DIR,
    DataLoader,
    DataProcessor,
    chart_factory,
    create_export_data,
    MESES,
    MESES_COMPLETOS,
    GRAPH_CONFIG,
    EPS_LIST
)
# Tipo para config de Plotly (dcc.Graph acepta dict pero el type checker espera Config)
_GRAPH_CFG: Any = GRAPH_CONFIG

# Donut IAAS: staticPlot puede afectar el layout de la leyenda en algunos navegadores
_GRAPH_CFG_IAAS: Any = {**GRAPH_CONFIG, "staticPlot": False}
# Radar comparación: hover por vértice / relleno (requiere staticPlot False)
_GRAPH_CFG_RADAR: Any = {**GRAPH_CONFIG, "staticPlot": False}


# 14 tonos fijos (familias distintas: un azul, un naranja, etc.) → orden = ranking mayor→menor
_COLORS_IAAS_FIXED: tuple[str, ...] = (
    "#E63946",  # rojo
    "#F77F00",  # naranja
    "#FCBF49",  # amarillo
    "#4ADE80",  # verde
    "#14B8A6",  # teal
    "#BE185D",  # magenta oscuro
    "#2563EB",  # azul
    "#4F46E5",  # índigo
    "#65A30D",  # verde oliva intenso
    "#EC4899",  # rosa
    "#F43F5E",  # coral
    "#581C87",  # púrpura oscuro
    "#A16207",  # ocre / marrón dorado
    "#92400E",  # marrón rojizo
)


def _palette_iaas_n(n: int) -> list[str]:
    """
    Paleta IAAS: hasta 14 colores fijos bien diferenciados; si hay más categorías,
    completa con HSV (caso raro).
    """
    if n <= 0:
        return []
    fixed = list(_COLORS_IAAS_FIXED[: min(n, len(_COLORS_IAAS_FIXED))])
    if n <= len(_COLORS_IAAS_FIXED):
        return fixed
    for i in range(len(_COLORS_IAAS_FIXED), n):
        h = (i * 0.618033988749895) % 1.0
        r, g, b = colorsys.hsv_to_rgb(h, 0.78, 0.92)
        fixed.append(f"#{int(r * 255):02x}{int(g * 255):02x}{int(b * 255):02x}")
    return fixed


# Configuracion de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def _morbilidad_total_anual(df_morb: pd.DataFrame) -> float:
    """
    Total anual mostrado: coincide con la suma de los subtotales mensuales
    (filas Total) de la hoja MORBILIDAD del Excel cuando existen.
    Si el detalle por diagnóstico no cuadra con esos pies, el Excel manda.
    """
    if df_morb.empty or "Total" not in df_morb.columns:
        return 0.0
    attrs = getattr(df_morb, "attrs", None) or {}
    if "total_anual" in attrs:
        v = attrs["total_anual"]
        if v is not None:
            return float(v)
    return float(pd.to_numeric(df_morb["Total"], errors="coerce").fillna(0).sum())


def _morbilidad_total_para_meses(df_morb: pd.DataFrame, meses_filtrados: list[str]) -> float:
    """
    Total de casos de morbilidad para los meses seleccionados: suma de los pies
    mensuales (attrs totales_mensuales) alineada con la hoja MORBILIDAD.
    Prioridad sobre total_anual cuando hay filtro de meses.
    """
    if df_morb.empty or "Total" not in df_morb.columns:
        return 0.0
    attrs = getattr(df_morb, "attrs", None) or {}
    monthly = attrs.get("totales_mensuales")
    if isinstance(monthly, dict) and meses_filtrados:
        s = int(sum(int(monthly.get(m, 0) or 0) for m in meses_filtrados))
        if s > 0:
            return float(s)
    return _morbilidad_total_anual(df_morb)


def _enteros_prop_por_suma(values: np.ndarray, target: int) -> np.ndarray:
    """Reparte `target` en enteros proporcionales a `values` (método mayores restos)."""
    values = np.asarray(values, dtype=float)
    s = float(values.sum())
    n = len(values)
    if n == 0 or target <= 0:
        return np.zeros(n, dtype=int)
    if s <= 0:
        return np.zeros(n, dtype=int)
    scaled = values * (target / s)
    floor = np.floor(scaled).astype(int)
    rem = int(target - floor.sum())
    frac = scaled - floor
    order = np.argsort(-frac)
    out = floor.copy()
    for j in range(rem):
        out[order[j % n]] += 1
    return out


def _morbilidad_escalada_al_total_oficial(df_morb: pd.DataFrame, total_oficial: float) -> pd.DataFrame:
    """
    Escala Hombre/Mujer/Total para que la suma de Total coincida con total_oficial
    (pies mensuales), repartiendo la diferencia de forma proporcional al detalle.
    """
    if df_morb.empty or "Total" not in df_morb.columns:
        return df_morb.copy()
    t = pd.to_numeric(df_morb["Total"], errors="coerce").fillna(0)
    detail = float(t.sum())
    target = int(round(total_oficial))
    if detail <= 0 or target <= 0 or abs(detail - target) < 1e-6:
        return df_morb.copy()
    new_totals = _enteros_prop_por_suma(t.values, target)
    out = df_morb.copy()
    for i, idx in enumerate(df_morb.index):
        Ti = int(new_totals[i])
        hi = float(pd.to_numeric(df_morb.loc[idx, "Hombre"], errors="coerce") or 0)
        mi = float(pd.to_numeric(df_morb.loc[idx, "Mujer"], errors="coerce") or 0)
        sm = hi + mi
        out.loc[idx, "Total"] = Ti
        if Ti <= 0:
            out.loc[idx, "Hombre"] = 0
            out.loc[idx, "Mujer"] = 0
        elif sm <= 0:
            out.loc[idx, "Hombre"] = Ti // 2
            out.loc[idx, "Mujer"] = Ti - int(out.loc[idx, "Hombre"])
        else:
            hi_n = int(round(Ti * hi / sm))
            hi_n = max(0, min(hi_n, Ti))
            out.loc[idx, "Hombre"] = hi_n
            out.loc[idx, "Mujer"] = Ti - hi_n
    return out


def _debug_log(run_id: str, hypothesis_id: str, location: str, message: str, data: dict[str, Any]) -> None:
    """Escribe logs NDJSON para depuracion en modo debug."""
    payload = {
        "sessionId": "57990a",
        "runId": run_id,
        "hypothesisId": hypothesis_id,
        "location": location,
        "message": message,
        "data": data,
        "timestamp": int(time() * 1000),
    }
    log_path = BASE_DIR / "debug-57990a.log"
    with log_path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(payload, ensure_ascii=False) + "\n")


def _build_compare_rows(meses_filtrados: list[str]) -> list[dict[str, Any]]:
    """Construye filas de comparación anual coherentes con meses seleccionados."""
    rows: list[dict[str, Any]] = []
    available_years = sorted(PROCESSOR_BY_YEAR.keys())
    for y in available_years:
        _ensure_year_data_fresh(y)
        proc_y = PROCESSOR_BY_YEAR[y]
        data_y = DATA_BY_YEAR.get(y, DATA)
        k_resumen = proc_y.get_kpis_resumen(meses_filtrados)
        k_gestion = proc_y.get_kpis_gestion(meses_filtrados)

        row_td = proc_y.safe_get_row('TOTAL DIAS ESTANCIA', meses_filtrados, silent=True)
        total_dias_meses = float(pd.to_numeric(row_td, errors='coerce').fillna(0).sum()) if isinstance(row_td, pd.Series) else 0.0
        total_dias_y = total_dias_meses if total_dias_meses > 0 else float(k_gestion.get("total_dias_estancia", 0) or 0)

        total_fallecidos_y = int(k_resumen.get("total_fallecidos", 0) or 0)
        total_egresos_y = float(k_resumen.get("total_egresos", 0) or 0)
        mort_uci_y = (total_fallecidos_y / total_egresos_y * 100.0) if total_egresos_y > 0 else 0.0

        df_eventos_det_y = data_y.get('eventos_adversos_detalle')
        if isinstance(df_eventos_det_y, pd.DataFrame) and not df_eventos_det_y.empty and 'Mes' in df_eventos_det_y.columns:
            total_eventos_y = int(df_eventos_det_y[df_eventos_det_y['Mes'].isin(meses_filtrados)].shape[0])
        else:
            df_eventos_y = data_y.get('eventos_adversos')
            total_eventos_y = int(pd.to_numeric(df_eventos_y.get('Cantidad', pd.Series(dtype=float)), errors='coerce').fillna(0).sum()) if isinstance(df_eventos_y, pd.DataFrame) else 0

        df_germ_det = data_y.get('germenes_detalle')
        if isinstance(df_germ_det, pd.DataFrame) and not df_germ_det.empty and 'Mes' in df_germ_det.columns:
            total_germenes_y = int(df_germ_det[df_germ_det['Mes'].isin(meses_filtrados)].shape[0])
        else:
            df_germenes_y = data_y.get('germenes')
            total_germenes_y = int(pd.to_numeric(df_germenes_y.get('Cantidad', pd.Series(dtype=float)), errors='coerce').fillna(0).sum()) if isinstance(df_germenes_y, pd.DataFrame) else 0

        df_morb_y = data_y.get('morbilidad')
        total_diag_y = 0
        if isinstance(df_morb_y, pd.DataFrame) and not df_morb_y.empty:
            attrs = getattr(df_morb_y, "attrs", None) or {}
            monthly = attrs.get("totales_mensuales", {})
            if isinstance(monthly, dict):
                total_diag_y = int(sum(int(monthly.get(m, 0) or 0) for m in meses_filtrados))
            if total_diag_y <= 0 and 'Total' in df_morb_y.columns:
                total_diag_y = int(round(_morbilidad_total_anual(df_morb_y)))

        rows.append({
            "year": y,
            "ingresos": int(k_resumen.get("total_ingresos", 0) or 0),
            "ocupacion": float(k_resumen.get("promedio_ocupacional", 0) or 0),
            "mortalidad": float(k_resumen.get("mortalidad", 0) or 0),
            "estancia": float(k_resumen.get("promedio_estancia", 0) or 0),
            "giro_cama": float(k_resumen.get("giro_cama", 0) or 0),
            "total_dias": float(total_dias_y),
            "fallecidos": total_fallecidos_y,
            "mort_ucI": float(mort_uci_y),
            "total_eventos": total_eventos_y,
            "total_germenes": total_germenes_y,
            "total_diag": total_diag_y,
        })
    return rows

# ==============================================================================
# INICIALIZACION DE LA APLICACION
# ==============================================================================

app = dash.Dash(
    __name__,
    external_stylesheets=[dbc.themes.BOOTSTRAP],
    suppress_callback_exceptions=True
)
server = app.server

# Cargar datos al inicio (multianio)
logger.info("Cargando datos...")
YEAR_FILES = {
    "2025": "Estadistica 2025.xlsx",
    "2026": "Estadistica 2026.xlsx",
}


def _empty_data() -> dict[str, pd.DataFrame]:
    return {
        'estadistica': pd.DataFrame(),
        'morbilidad': pd.DataFrame(),
        'germenes': pd.DataFrame(),
        'germenes_detalle': pd.DataFrame(),
        'fallecidos': pd.DataFrame(),
        'eventos_adversos': pd.DataFrame(),
        'eventos_adversos_detalle': pd.DataFrame(),
        'gram_sot': pd.DataFrame()
    }


DATA_BY_YEAR: dict[str, dict[str, pd.DataFrame]] = {}
PROCESSOR_BY_YEAR: dict[str, DataProcessor] = {}
_YEAR_FILE_MTIME: dict[str, float] = {}

for year, filename in YEAR_FILES.items():
    excel_path = BASE_DIR / "data" / filename
    if not excel_path.exists():
        logger.warning("Archivo no encontrado para %s: %s", year, excel_path)
        continue

    try:
        year_data = DataLoader(Path(excel_path)).load_all()
        DATA_BY_YEAR[year] = year_data
        PROCESSOR_BY_YEAR[year] = DataProcessor(year_data)
        try:
            _YEAR_FILE_MTIME[year] = excel_path.stat().st_mtime
        except OSError:
            pass
        logger.info("Datos cargados para %s desde %s", year, filename)
        # region agent log
        _debug_log(
            "pre-fix",
            "H2",
            "dashboard_uci.py:data_load_loop",
            "Year loaded",
            {
                "year": year,
                "file": filename,
                "estadisticaRows": int(year_data.get("estadistica", pd.DataFrame()).shape[0]),
                "morbilidadRows": int(year_data.get("morbilidad", pd.DataFrame()).shape[0]),
            },
        )
        # endregion
    except Exception as e:
        logger.error("Error cargando datos de %s: %s", year, e)
        # region agent log
        _debug_log(
            "pre-fix",
            "H2",
            "dashboard_uci.py:data_load_loop",
            "Year failed to load",
            {"year": year, "file": filename, "error": str(e)},
        )
        # endregion

if not PROCESSOR_BY_YEAR:
    DATA_BY_YEAR["2026"] = _empty_data()
    PROCESSOR_BY_YEAR["2026"] = DataProcessor(DATA_BY_YEAR["2026"])


def _ensure_year_data_fresh(year_key: str) -> None:
    """Recarga datos desde el Excel si el archivo en disco cambió (sin reiniciar el servidor)."""
    if year_key not in YEAR_FILES:
        return
    excel_path = BASE_DIR / "data" / YEAR_FILES[year_key]
    if not excel_path.exists():
        return
    try:
        mtime = excel_path.stat().st_mtime
    except OSError:
        return
    if _YEAR_FILE_MTIME.get(year_key) == mtime:
        return
    try:
        year_data = DataLoader(excel_path).load_all()
        DATA_BY_YEAR[year_key] = year_data
        PROCESSOR_BY_YEAR[year_key] = DataProcessor(year_data)
        _YEAR_FILE_MTIME[year_key] = mtime
        logger.info("Datos recargados para %s tras cambio en el Excel", year_key)
    except Exception as e:
        logger.error("Error recargando datos de %s: %s", year_key, e)


AVAILABLE_YEARS = sorted(PROCESSOR_BY_YEAR.keys())
DEFAULT_YEAR = "2025" if "2025" in PROCESSOR_BY_YEAR else AVAILABLE_YEARS[0]

# Opción especial para "Todos los meses" en el filtro
VALOR_TODOS_MESES = '__TODOS__'

# Variables por compatibilidad con otras pestañas (default 2026)
DATA = DATA_BY_YEAR[DEFAULT_YEAR]
PROCESSOR = PROCESSOR_BY_YEAR[DEFAULT_YEAR]

# region agent log
_debug_log(
    "pre-fix",
    "H2",
    "dashboard_uci.py:year_registry",
    "Year registry initialized",
    {"availableYears": AVAILABLE_YEARS, "defaultYear": DEFAULT_YEAR},
)
# endregion

# ==============================================================================
# LAYOUT DE LA APLICACION
# ==============================================================================

app.layout = html.Div([
    dcc.Store(id='selected-year-store', data=DEFAULT_YEAR),
    dcc.Store(id='pending-year-store', data=DEFAULT_YEAR),
    dcc.Store(id='selected-months-store', data=MESES),
    # Header
    html.Div([
        html.Div([
            html.Div([
                html.Img(src='/assets/Logo.png', style={
                    'height': '70px',
                    'marginRight': '20px'
                }, className="logo-img"),
                html.Div([
                    html.H1("UCI Adultos", style={
                        'color': '#ffffff', 'fontWeight': '800', 'margin': '0',
                        'fontSize': '32px', 'lineHeight': '1.1', 'letterSpacing': '-0.5px'
                    }),
                    html.P("Panel de Indicadores 2025-2026", style={
                        'color': '#cbd5e1', 'fontSize': '14px', 'margin': '4px 0 0 0',
                        'fontWeight': '500', 'letterSpacing': '0.5px'
                    }, className="header-subtitle")
                ])
            ], style={'display': 'flex', 'alignItems': 'center', 'justifyContent': 'center', 'flex': '1'}),
            html.Div([
                dbc.Button("Exportar Excel", id="btn-export-excel", color="success", size="sm", className="me-2"),
                dbc.Button("Exportar CSV", id="btn-export-csv", color="primary", size="sm", className="me-2"),
                dcc.Download(id="download-excel"),
                dcc.Download(id="download-csv"),
                dcc.Download(id="download-compare-excel"),
                dcc.Download(id="download-compare-csv"),
            ], style={'display': 'flex', 'gap': '8px', 'position': 'absolute', 'right': '20px', 'flexWrap': 'wrap', 'justifyContent': 'flex-end'})
        ], style={'display': 'flex', 'justifyContent': 'center', 'alignItems': 'center', 'position': 'relative', 'padding': '20px 0'})
    ], className="header-main", style={
        'padding': '16px 24px'
    }),
    # Tabs
    dbc.Tabs(id="tabs", active_tab='tab-resumen', children=[
        dbc.Tab(label='Resumen General', tab_id='tab-resumen', label_style={'cursor': 'pointer'}),
        dbc.Tab(label='Gesti\u00f3n y Calidad', tab_id='tab-gestion', label_style={'cursor': 'pointer'}),
        dbc.Tab(label='Especialidades', tab_id='tab-especialidades', label_style={'cursor': 'pointer'}),
        dbc.Tab(label='Mortalidad', tab_id='tab-mortalidad', label_style={'cursor': 'pointer'}),
    ]),
    html.Div(id='tab-content', className="tab-content"),
    # Footer
    html.Div([
        html.Div([
            html.Div([
                html.Div([
                    html.P("Hospital Infantil Universitario de San Jos\u00e9", style={
                        'color': '#f1f5f9', 'fontWeight': '600', 'margin': '0', 'fontSize': '14px',
                        'letterSpacing': '0.3px'
                    }),
                    html.P("Unidad de Cuidados Intensivos Adultos", style={
                        'color': '#94a3b8', 'margin': '4px 0 0 0', 'fontSize': '12px'
                    })
                ]),
            ], style={'display': 'flex', 'alignItems': 'center', 'gap': '12px'}),
            html.P("Panel de Indicadores 2025-2026", style={
                'color': '#64748b', 'margin': '0', 'fontSize': '12px'
            })
        ], style={'display': 'flex', 'justifyContent': 'space-between', 'alignItems': 'center', 'flexWrap': 'wrap', 'gap': '16px'})
    ], className="dashboard-footer")
], style={
    'background': 'transparent',
    'minHeight': '100vh',
    'fontFamily': "'Roboto', -apple-system, BlinkMacSystemFont, sans-serif"
}, id='app-container')


# ==============================================================================
# CALLBACKS
# ==============================================================================

@app.callback(
    Output('tab-content', 'children'),
    [
        Input('tabs', 'active_tab'),
        Input('selected-year-store', 'data'),
        Input('pending-year-store', 'data'),
        Input('selected-months-store', 'data'),
    ]
)
def render_content(tab, selected_year, pending_year, selected_months):
    """
    Renderiza el contenido de cada pestaña.
    - En Resumen, no se reconstruye el layout cuando cambian stores aplicados.
    - En otras pestañas, se usa el periodo global aplicado (año + meses).
    """
    ctx = dash.callback_context
    triggered_id = ctx.triggered[0]['prop_id'].split('.')[0] if ctx.triggered else None

    # Si estamos en Resumen y el cambio viene solo del a\u00f1o aplicado, mantenemos el layout actual
    if tab == 'tab-resumen' and triggered_id == 'selected-year-store':
        return dash.no_update

    # Cambio solo en el desplegable de a\u00f1o (sin Aplicar): no re-renderizar Resumen para no pisar el valor del dropdown
    if tab == 'tab-resumen' and triggered_id == 'pending-year-store':
        return dash.no_update

    if tab == 'tab-resumen' and triggered_id == 'selected-months-store':
        return dash.no_update

    selected_resolved = str(selected_year) if selected_year is not None else DEFAULT_YEAR
    selected_resolved = selected_resolved if selected_resolved in PROCESSOR_BY_YEAR else DEFAULT_YEAR

    if tab == 'tab-resumen':
        return render_resumen(selected_resolved)

    _ = pending_year  # se conserva para sincronizar el dropdown, sin aplicar fuera de Resumen.
    meses_filtrados = (
        [m for m in selected_months if m in MESES]
        if isinstance(selected_months, list) and selected_months
        else list(MESES)
    )
    year_key = selected_resolved

    if tab == 'tab-gestion':
        return render_gestion(year_key, meses_filtrados)
    elif tab == 'tab-especialidades':
        return render_especialidades(year_key, meses_filtrados)
    elif tab == 'tab-mortalidad':
        return render_mortalidad(year_key, meses_filtrados)
    return html.Div([html.H5("Selecciona una pesta\u00f1a", style={'color': '#1e293b'})])


def render_resumen(year_key: str):
    filtros = html.Div([
        html.Div([
            html.Span("Per\u00edodo de an\u00e1lisis", style={
                'color': '#f1f5f9', 'fontWeight': '600', 'fontSize': '14px'
            })
        ], style={'marginBottom': '12px'}, className="filtros-titulo"),
        html.Div([
            # Año: etiqueta + dropdown
            html.Div([
                html.Label("Año", className="filtro-label", style={
                    'color': '#94a3b8', 'fontSize': '12px', 'fontWeight': '600', 'marginBottom': '6px', 'display': 'block'
                }),
                dcc.Dropdown(
                    id='filtro-anio',
                    options=[{'label': anio, 'value': anio} for anio in AVAILABLE_YEARS],
                    value=year_key,
                    clearable=False,
                    persistence=True,
                    persistence_type='session',
                    className='uci-dd-dark',
                    style={
                        'fontSize': '13px',
                        'width': '100%',
                        'color': '#f1f5f9',
                        'backgroundColor': '#1e293b',
                    },
                ),
            ], className="filtro-grupo"),
            # Meses: etiqueta + dropdown
            html.Div([
                html.Label("Meses", className="filtro-label", style={
                    'color': '#94a3b8', 'fontSize': '12px', 'fontWeight': '600', 'marginBottom': '6px', 'display': 'block'
                }),
                dcc.Dropdown(
                    id='filtro-mes-inicio',
                    options=[{'label': MESES_COMPLETOS[mes], 'value': mes} for mes in MESES],
                    value=MESES,
                    multi=True,
                    placeholder="Seleccionar meses",
                    className='uci-dd-dark',
                    style={
                        'fontSize': '13px',
                        'width': '100%',
                        'color': '#f1f5f9',
                        'backgroundColor': '#1e293b',
                    },
                ),
            ], className="filtro-grupo"),
            # Botones
            html.Div([
                dbc.Button("Aplicar", id='btn-aplicar-filtros', color="success", size="sm",
                           title="Actualizar con año y meses seleccionados", className="filtro-btn"),
                dbc.Button("Todos", id='btn-todos-meses', color="secondary", size="sm",
                           title="Seleccionar todos los meses", className="filtro-btn"),
                dcc.Checklist(
                    id='comparar-anios',
                    options=[{'label': 'Comparar 2025 vs 2026', 'value': 'compare'}],
                    value=[],
                    inputStyle={'marginRight': '6px'},
                    labelStyle={'color': '#e5e7eb', 'fontSize': '12px', 'fontWeight': '500'}
                ),
                dbc.Button(
                    "Exportar comparación Excel",
                    id="btn-export-compare-excel",
                    color="success",
                    size="sm",
                    className="filtro-btn me-2",
                    title="Descarga la tabla 2025 vs 2026 en Excel según los meses guardados en Resumen (tras Aplicar)",
                ),
                dbc.Button(
                    "Exportar comparación CSV",
                    id="btn-export-compare-csv",
                    color="primary",
                    size="sm",
                    className="filtro-btn",
                    title="Descarga la tabla 2025 vs 2026 según los meses guardados en Resumen (tras Aplicar)",
                ),
            ], className="filtros-acciones")
        ], style={'display': 'grid', 'gridTemplateColumns': '100px minmax(0, 1fr)', 'alignItems': 'start', 'columnGap': '16px', 'rowGap': '16px'}, className="filtros-inner")
    ], style={
        'background': 'rgba(30, 41, 59, 0.8)',
        'padding': '14px 20px',
        'borderRadius': '8px',
        'marginBottom': '16px',
        'border': '1px solid rgba(148, 163, 184, 0.15)'
    }, className="filtros-section")
    return html.Div([filtros, html.Div(id='resumen-content')])


def render_gestion(year_key: str, meses_filtrados: list[str] | None = None):
    try:
        _ensure_year_data_fresh(year_key)
        processor = PROCESSOR_BY_YEAR.get(year_key, PROCESSOR)
        data_year = DATA_BY_YEAR.get(year_key, DATA)
        meses = [m for m in (meses_filtrados or MESES) if m in MESES]
        if not meses:
            meses = list(MESES)
        # region agent log
        _debug_log(
            "pre-fix",
            "H5",
            "dashboard_uci.py:render_gestion:entry",
            "Gestion tab resolved year",
            {
                "globalDefaultYear": DEFAULT_YEAR,
                "resolvedYear": year_key,
                "mesesFiltrados": meses,
            },
        )
        # endregion
        kpis = processor.get_kpis_gestion(meses)
        if len(meses) == 1:
            rango_gestion = f"{MESES_COMPLETOS.get(meses[0], meses[0])} {year_key}"
        elif len(meses) == len(MESES):
            rango_gestion = processor.etiqueta_rango_periodo(year_key)
        else:
            rango_gestion = f"{', '.join([MESES_COMPLETOS.get(m, m)[:3] for m in meses])} {year_key}"
        kpi_items_g = [
            (float(kpis['total_dias_estancia'] or 0), chart_factory.create_kpi_card("Total D\u00edas Estancia", f"{int(kpis['total_dias_estancia'])}", rango_gestion, "#10b981")),
            (float(kpis['capacidad_uci'] or 0), chart_factory.create_kpi_card("Capacidad UCI", f"{int(kpis['capacidad_uci'])}", "Camas", "#8b5cf6")),
            (float(kpis['promedio_dias_vm'] or 0), chart_factory.create_kpi_card("Promedio D\u00edas VM", f"{kpis['promedio_dias_vm']:.2f}", "Ventilaci\u00f3n Mec\u00e1nica", "#06b6d4")),
            (float(kpis['promedio_estancia'] or 0), chart_factory.create_kpi_card("Estancia Promedio", f"{kpis['promedio_estancia']:.2f}", "D\u00edas", "#6366f1")),
            (float(kpis['giro_cama'] or 0), chart_factory.create_kpi_card("Giro cama", f"{kpis['giro_cama']:.2f}", "Promedio", "#3b82f6")),
        ]
        kpi_cards_g = [c for _, c in sorted(kpi_items_g, key=lambda x: x[0], reverse=True)]

        datos_estancia = processor.safe_get_row('PROMEDIO ESTANCIA', meses)
        fig_estancia = chart_factory.create_line_chart(datos_estancia, 'Estancia Promedio (D\u00edas)', color='#6366f1', decimals=1)
        datos_giro = processor.safe_get_row('GIRO CAMA', meses)
        fig_giro = chart_factory.create_bar_chart(datos_giro, 'Giro cama', color='#3b82f6', decimals=2)

        # Grafico de VM - datos mensuales
        datos_vm_mensual = kpis.get('promedio_dias_vm_mensual', [])
        if not datos_vm_mensual or len(datos_vm_mensual) != len(meses):
            datos_vm_mensual = [round(kpis['promedio_dias_vm'], 2) for _ in meses]
        vm_text_positions = ['top center'] * len(meses)
        if vm_text_positions:
            vm_text_positions[0] = 'top right'
            vm_text_positions[-1] = 'top left'
        fig_vm = go.Figure()
        fig_vm.add_trace(go.Scatter(
            x=meses, y=datos_vm_mensual,
            mode='lines+markers+text',
            line=dict(color='#06b6d4', width=3.2),
            marker=dict(size=8, color='#06b6d4'),
            cliponaxis=False,
            fill='tozeroy',
            fillcolor='rgba(6, 182, 212, 0.18)',
            text=[f'{v:.1f}' if v > 0 else '' for v in datos_vm_mensual],
            textposition=vm_text_positions,
            textfont=dict(color='#f1f5f9', size=12),
            hovertemplate='<b>%{x}</b><br>Promedio VM: %{y:.1f}<extra></extra>'
        ))
        fig_vm.update_layout(
            title={'text': 'Promedio de D\u00edas de VM (Tendencia)', 'font': {'size': 14, 'color': '#f1f5f9'}, 'x': 0.5},
            transition={'duration': 0},
            plot_bgcolor='#1e293b', paper_bgcolor='#1e293b',
            height=470, margin=dict(l=50, r=30, t=50, b=60),
            hovermode='x unified',
            xaxis={
                'tickfont': {'color': '#cbd5e1', 'size': 12},
                'gridcolor': 'rgba(148, 163, 184, 0.1)',
                'range': [-0.35, len(meses) - 0.65]
            },
            yaxis={'tickfont': {'color': '#94a3b8', 'size': 12}, 'gridcolor': 'rgba(148, 163, 184, 0.1)'}
        )

        # Variantes de nombre de fila en ESTADISTICA (Excel 2025/2026 pueden diferir en espacios / DE / tildes)
        _its_casos_filas = [
            'N\u00daMERO DE BACTERIEMIAS ASOC.  A  C. CENTRAL',
            'N\u00daMERO DE BACTERIEMIAS ASOC. A C. CENTRAL',
            'NUMERO DE BACTERIEMIAS ASOC.  A  C. CENTRAL',
            'NUMERO DE BACTERIEMIAS ASOC. A C. CENTRAL',
            'BACTERIEMIAS ASOCIADAS A CATETER CENTRAL',
            'BACTERIEMIAS ASOC',
        ]
        _its_dias_filas = [
            'DIAS CAT\u00c9TER CENTRAL',
            'D\u00cdAS CAT\u00c9TER CENTRAL',
            'DIAS CATETER CENTRAL',
            'DIAS CATER CENTRAL',
        ]
        _nav_casos_filas = [
            'NUMERO CASOS NAV',
            'N\u00daMERO CASOS NAV',
            'NUMERO DE CASOS NAV',
            'N\u00daMERO DE CASOS NAV',
            'CASOS NAV',
        ]
        _nav_dias_filas = [
            'DIAS VENTILADOR',
            'D\u00cdAS VENTILADOR',
            'DIAS DE VENTILADOR',
        ]
        fig_bacteriemia = chart_factory.create_line_chart(
            processor.serie_tasa_por_mil_dias(_its_casos_filas, _its_dias_filas, meses),
            'Tasa de ITS AC (por 1.000 d\u00edas cat\u00e9ter central)', color='#ef4444', decimals=1
        )
        fig_nav = chart_factory.create_line_chart(
            processor.serie_tasa_por_mil_dias(_nav_casos_filas, _nav_dias_filas, meses),
            'Tasa NAV (por 1.000 d\u00edas VM)', color='#ef4444', decimals=1
        )

        # Germenes - Donut
        if not data_year['germenes'].empty:
            df_germ_det = data_year.get('germenes_detalle')
            if isinstance(df_germ_det, pd.DataFrame) and not df_germ_det.empty and 'Mes' in df_germ_det.columns:
                df_germ_det = df_germ_det[df_germ_det['Mes'].isin(meses)]
                if not df_germ_det.empty and 'Germen' in df_germ_det.columns:
                    g_counts = df_germ_det['Germen'].value_counts()
                    germenes_labels = g_counts.index.astype(str).tolist()
                    germenes_values = g_counts.values.tolist()
                else:
                    germenes_labels = []
                    germenes_values = []
            else:
                germenes_labels = data_year['germenes']['Germen'].tolist()
                germenes_values = data_year['germenes']['Cantidad'].tolist()
            datos_germenes = sorted(zip(germenes_values, germenes_labels), reverse=True)
            germenes_values_sorted = [d[0] for d in datos_germenes]
            germenes_labels_sorted = [d[1] for d in datos_germenes]
            total_germenes = sum(germenes_values)
            n_g = len(germenes_labels_sorted)
            colors_germenes = _palette_iaas_n(n_g)
            fig_germenes = go.Figure(data=[go.Pie(
                labels=germenes_labels_sorted,
                values=germenes_values_sorted,
                hole=0.55,
                marker=dict(colors=colors_germenes, line=dict(color='#0f172a', width=2)),
                textinfo='percent',
                textposition='inside',
                textfont=dict(size=12, color='white', family='Roboto'),
                hovertemplate='<b>%{label}</b><br>Cantidad: %{value}<br>Porcentaje: %{percent}<extra></extra>',
                direction='clockwise',
                sort=False,
                rotation=90,
                pull=[0.03 if i == 0 else 0 for i in range(n_g)]
            )])
            fig_germenes.add_annotation(
                text=f"<b style='font-size:22px'>{int(total_germenes)}</b><br><span style='color:#94a3b8'>Aislados</span>",
                x=0.5, y=0.5,
                font=dict(size=14, color='#f1f5f9', family='Roboto'),
                showarrow=False
            )
            fig_germenes.update_layout(
                title={'text': 'G\u00e9rmenes Aislados (IAAS)', 'font': {'size': 13, 'color': '#f1f5f9', 'family': 'Roboto'}, 'x': 0.5},
                transition={'duration': 0},
                plot_bgcolor='#1e293b', paper_bgcolor='#1e293b',
                autosize=True, height=420,
                margin=dict(l=20, r=230, t=50, b=70),
                legend=dict(
                    orientation='v',
                    xref='paper',
                    yref='paper',
                    x=1.01,
                    xanchor='left',
                    y=1,
                    yanchor='top',
                    bgcolor='rgba(15, 23, 42, 0.92)',
                    font=dict(color='#e2e8f0', size=10, family='Roboto'),
                    bordercolor='rgba(148, 163, 184, 0.35)',
                    borderwidth=1,
                    itemsizing='constant',
                    traceorder='normal',
                ),
                showlegend=True
            )
        else:
            fig_germenes = go.Figure()
            fig_germenes.update_layout(
                title={'text': 'Gérmenes Aislados (IAAS)', 'font': {'size': 16, 'color': '#f1f5f9', 'family': 'Roboto'}, 'x': 0.5},
                plot_bgcolor='#1e293b',
                paper_bgcolor='#1e293b',
                height=520,
                margin=dict(l=30, r=30, t=60, b=30),
                xaxis={'visible': False},
                yaxis={'visible': False},
                showlegend=False
            )
            fig_germenes.add_annotation(
                text="<b>Sin datos de gérmenes</b><br><span style='color:#94a3b8'>Carga IAAS para visualizar este gráfico</span>",
                x=0.5, y=0.5, xref='paper', yref='paper',
                showarrow=False,
                font=dict(size=14, color='#e2e8f0', family='Roboto'),
                align='center'
            )

        # Eventos adversos
        df_eventos_det = data_year.get('eventos_adversos_detalle', pd.DataFrame())
        if isinstance(df_eventos_det, pd.DataFrame) and not df_eventos_det.empty and {'Mes', 'Evento'}.issubset(df_eventos_det.columns):
            df_eventos_base = df_eventos_det[df_eventos_det['Mes'].isin(meses)]
            eventos_counts = (
                df_eventos_base['Evento'].value_counts().reset_index(name='Cantidad').rename(columns={'index': 'Evento'})
                if not df_eventos_base.empty
                else pd.DataFrame(columns=['Evento', 'Cantidad'])
            )
        else:
            df_eventos_base = df_eventos_det
            eventos_counts = data_year['eventos_adversos']
        if not eventos_counts.empty:
            df_eventos = eventos_counts.sort_values('Cantidad', ascending=False)
            if len(df_eventos) > 12:
                df_eventos_top = df_eventos.head(12).copy()
                otros_eventos = int(df_eventos.iloc[12:]['Cantidad'].sum())
                if otros_eventos > 0:
                    df_eventos_top = pd.concat(
                        [df_eventos_top, pd.DataFrame([{'Evento': 'Otros', 'Cantidad': otros_eventos}])],
                        ignore_index=True
                    )
                # Orden Pareto: por frecuencia (si "Otros" suma más, debe ir a la izquierda)
                df_eventos = df_eventos_top.sort_values('Cantidad', ascending=False).reset_index(drop=True)
            else:
                df_eventos = df_eventos.reset_index(drop=True)
            total_eventos = max(float(df_eventos['Cantidad'].sum()), 1.0)
            df_eventos['Acumulado %'] = (df_eventos['Cantidad'].cumsum() / total_eventos) * 100
            n_bar = len(df_eventos)
            bar_colors = _palette_iaas_n(n_bar)
            bar_values = pd.to_numeric(df_eventos['Cantidad'], errors='coerce').fillna(0.0)
            # Resaltar especialmente los "1" para que se lean mejor en barras pequeñas.
            bar_labels_main = [f"{int(v)}" if float(v) > 1 else '' for v in bar_values]
            acum_values = pd.to_numeric(df_eventos['Acumulado %'], errors='coerce').fillna(0.0)
            acum_labels = [f"{float(v):.1f}%" for v in acum_values]
            fig_eventos = go.Figure()
            fig_eventos.add_trace(go.Bar(
                x=df_eventos['Evento'],
                y=bar_values,
                name='N.º de eventos',
                marker_color=bar_colors,
                text=bar_labels_main,
                textposition='inside',
                insidetextanchor='middle',
                textfont=dict(color='#f8fafc', size=12, family='Roboto'),
                hovertemplate='<b>%{x}</b><br>Eventos: %{y}<extra></extra>'
            ))
            # Mostrar solo los "1" centrados y en negrita, sin trazas duplicadas.
            for evento, v in zip(df_eventos['Evento'], bar_values):
                if float(v) == 1.0:
                    fig_eventos.add_annotation(
                        x=evento,
                        y=0.5,
                        xref='x',
                        yref='y',
                        text='1',
                        showarrow=False,
                        font=dict(color='#ffffff', size=13, family='Arial Black'),
                        align='center',
                    )
            fig_eventos.add_trace(go.Scatter(
                x=df_eventos['Evento'],
                y=acum_values,
                name='Acumulado %',
                mode='lines+markers',
                yaxis='y2',
                line=dict(color='#fbbf24', width=2.5),
                marker=dict(size=6, color='#fbbf24'),
                hovertemplate='<b>%{x}</b><br>Acumulado: %{y:.1f}%<extra></extra>'
            ))
            # Etiquetas del acumulado con fondo para que no se mezclen con las barras.
            for i, (evento, v, lbl) in enumerate(zip(df_eventos['Evento'], acum_values, acum_labels)):
                # Siempre anclar al punto de la línea acumulada (y2).
                ann_y = float(v)
                ann_yref = 'y2'
                if i in (0, 1):
                    yshift = 18
                else:
                    yshift = -14 if (i % 2 == 1 or float(v) >= 97.0) else 14
                xshift = 0
                fig_eventos.add_annotation(
                    x=evento,
                    y=ann_y,
                    xref='x',
                    yref=ann_yref,
                    text=lbl,
                    showarrow=False,
                    yshift=yshift,
                    xshift=xshift,
                    font=dict(color='#fef9c3', size=11, family='Roboto'),
                    align='center',
                    bgcolor='rgba(15, 23, 42, 0.92)',
                    bordercolor='rgba(251, 191, 36, 0.55)',
                    borderwidth=1,
                    borderpad=2,
                )
            fig_eventos.update_layout(
                title={
                    'text': 'Eventos adversos (Pareto)',
                    'subtitle': {
                        'text': (
                            'Barras: veces que ocurrió cada tipo de evento (año). '
                            'Línea amarilla: % acumulado — muestra qué parte del total '
                            'explican los eventos más frecuentes (regla 80/20).'
                        ),
                        'font': {'size': 11, 'color': '#94a3b8', 'family': 'Roboto'},
                    },
                    'font': {'size': 13, 'color': '#f1f5f9', 'family': 'Roboto'},
                    'x': 0.5,
                    'xanchor': 'center',
                },
                transition={'duration': 0},
                plot_bgcolor='#1e293b', paper_bgcolor='#1e293b',
                autosize=True, height=600, margin=dict(l=48, r=36, t=72, b=180),
                hovermode='x unified',
                xaxis={
                    'tickfont': {'color': '#e2e8f0', 'size': 10},
                    'tickangle': -42,
                    'ticklabelstandoff': 10,
                    'automargin': True,
                    'gridcolor': 'rgba(148, 163, 184, 0.1)'
                },
                yaxis={
                    'title': {'text': 'N.º de eventos', 'font': {'color': '#94a3b8'}},
                    'tickfont': {'color': '#94a3b8'},
                    'gridcolor': 'rgba(148, 163, 184, 0.1)',
                },
                yaxis2={
                    'overlaying': 'y',
                    'side': 'right',
                    'range': [0, 118],
                    'tickfont': {'color': '#fbbf24'},
                    'title': {
                        'text': '% acumulado (de izq. a der.)',
                        'font': {'color': '#fbbf24'},
                    },
                    'showgrid': False
                },
                legend=dict(
                    orientation="h", y=-0.50, x=0.5, xanchor="center",
                    bgcolor='rgba(30, 41, 59, 0.8)', font={'color': '#f1f5f9', 'size': 12},
                    bordercolor='rgba(148, 163, 184, 0.3)', borderwidth=1
                ),
                bargap=0.3
            )
        else:
            fig_eventos = go.Figure()
            fig_eventos.update_layout(
                title={'text': 'Eventos Adversos por Categoría', 'font': {'size': 14, 'color': '#f1f5f9'}, 'x': 0.5},
                plot_bgcolor='#1e293b',
                paper_bgcolor='#1e293b',
                height=640,
                margin=dict(l=55, r=40, t=55, b=120),
                xaxis={'visible': False},
                yaxis={'visible': False},
                showlegend=False
            )
            fig_eventos.add_annotation(
                text="<b>Sin datos de eventos adversos</b>",
                x=0.5, y=0.5, xref='paper', yref='paper',
                showarrow=False,
                font=dict(size=14, color='#e2e8f0', family='Roboto')
            )

        # Distribución de pacientes con eventos adversos (1, 2, 3+ eventos)
        df_det = df_eventos_base if isinstance(df_eventos_base, pd.DataFrame) else data_year.get('eventos_adversos_detalle', pd.DataFrame())
        if not df_det.empty and 'Nombre' in df_det.columns and 'Evento' in df_det.columns:
            eventos_por_pac = df_det.groupby('Nombre').size()
            uno = int((eventos_por_pac == 1).sum())
            dos = int((eventos_por_pac == 2).sum())
            tres_mas = int((eventos_por_pac >= 3).sum())
            dist_labels = ['1 evento', '2 eventos', '3+ eventos']
            dist_vals = [uno, dos, tres_mas]
            if sum(dist_vals) > 0:
                fig_dist_pac = go.Figure(data=[go.Bar(
                    x=dist_labels, y=dist_vals,
                    marker_color=['#6366f1', '#8b5cf6', '#a855f7'],
                    text=dist_vals, textposition='auto',
                    textfont=dict(color='#f1f5f9', size=12),
                    hovertemplate='<b>%{x}</b><br>Pacientes: %{y}<extra></extra>'
                )])
                fig_dist_pac.update_layout(
                    title={'text': 'Distribución de Pacientes con Eventos Adversos', 'font': {'size': 14, 'color': '#f1f5f9'}, 'x': 0.5},
                    plot_bgcolor='#1e293b', paper_bgcolor='#1e293b',
                    height=320, margin=dict(l=50, r=30, t=50, b=50),
                    xaxis={'tickfont': {'color': '#cbd5e1'}, 'gridcolor': 'rgba(148, 163, 184, 0.1)'},
                    yaxis={'tickfont': {'color': '#94a3b8'}, 'gridcolor': 'rgba(148, 163, 184, 0.1)'},
                    showlegend=False
                )
            else:
                fig_dist_pac = None
        else:
            fig_dist_pac = None

        # Distribución por especialidad de pacientes con eventos adversos (no mostrar en 2025).
        # Coherente con 1/2/3+ eventos: contar pacientes únicos.
        if year_key != '2025' and not df_det.empty and {'Nombre', 'Especialidad'}.issubset(df_det.columns):
            esp_counts = df_det.groupby('Nombre')['Especialidad'].first().value_counts()
            esp_labels = [str(x).strip() or 'Sin especificar' for x in esp_counts.index.tolist()]
            esp_values = esp_counts.tolist()
            if esp_values:
                fig_esp_ea = chart_factory.create_donut_chart(
                    esp_labels, [float(v) for v in esp_values],
                    'Distribución por Especialidad (Pacientes Únicos con Eventos Adversos)',
                    auto_sort=True
                )
                fig_esp_ea.update_layout(height=380)
            else:
                fig_esp_ea = None
        else:
            fig_esp_ea = None

        return html.Div([
            html.H5("Indicadores de Gesti\u00f3n", style={'color': '#f1f5f9', 'fontWeight': '600'}, className="mb-3"),
            dbc.Row([
                *[dbc.Col(k, xs=12, md=True) for k in kpi_cards_g]
            ], className="mb-4 g-3"),
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_estancia, config=_GRAPH_CFG, style={'height': '350px'}), xs=12, className="plotly-chart plotly-chart--sm")
            ], className="mb-4 g-3"),
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_giro, config=_GRAPH_CFG, style={'height': '350px'}), xs=12, className="plotly-chart plotly-chart--sm")
            ], className="mb-3 g-3"),
            dbc.Row([
                dbc.Col(xs=12, md=1),
                dbc.Col(dcc.Graph(figure=fig_vm, config=_GRAPH_CFG, style={'height': '470px'}), xs=12, md=10, className="plotly-chart plotly-chart--lg"),
                dbc.Col(xs=12, md=1)
            ], className="mb-4", justify="center"),
            html.H5("Indicadores de Calidad e Infecciones", style={'color': '#f1f5f9', 'fontWeight': '600', 'marginTop': '16px', 'marginBottom': '16px', 'textAlign': 'center'}),
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_bacteriemia, config=_GRAPH_CFG, style={'height': '350px'}), xs=12, className="plotly-chart plotly-chart--sm")
            ], className="mb-4 g-3"),
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_nav, config=_GRAPH_CFG, style={'height': '350px'}), xs=12, className="plotly-chart plotly-chart--sm")
            ], className="mb-3 g-3"),
            html.H5("G\u00e9rmenes Aislados y Eventos Adversos", style={'color': '#f1f5f9', 'fontWeight': '600', 'marginTop': '16px', 'marginBottom': '16px', 'textAlign': 'center'}),
            dbc.Row([
                dbc.Col(xs=12, md=2),
                dbc.Col(dcc.Graph(figure=fig_germenes, config=_GRAPH_CFG_IAAS, style={'height': '520px'}), xs=12, md=8, className="plotly-chart plotly-chart--xl"),
                dbc.Col(xs=12, md=2)
            ], className="mb-4", justify="center"),
            dbc.Row([dbc.Col(dcc.Graph(figure=fig_eventos, config=_GRAPH_CFG, style={'height': '640px'}), xs=12, className="plotly-chart plotly-chart--xxl")], className="mb-4 g-3"),
        ] + ([dbc.Row(
            [dbc.Col(dcc.Graph(figure=fig_dist_pac, config=_GRAPH_CFG, style={'height': '320px'}), xs=12, md=6, className="plotly-chart plotly-chart--sm"),
             dbc.Col(dcc.Graph(figure=fig_esp_ea, config=_GRAPH_CFG, style={'height': '380px'}), xs=12, md=6, className="plotly-chart plotly-chart--md")]
            if fig_dist_pac and fig_esp_ea
            else [dbc.Col(dcc.Graph(figure=(fig_dist_pac or fig_esp_ea), config=_GRAPH_CFG, style={'height': '380px'}), xs=12, className="plotly-chart plotly-chart--md")],
            className="mb-5 g-3"
        )] if (fig_dist_pac or fig_esp_ea) else []))
    except Exception as e:
        logger.error(f"Error en render_gestion: {e}")
        return html.Div([html.H3(f"Error: {str(e)}", style={'color': '#e74c3c'})])


def render_especialidades(year_key: str, meses_filtrados: list[str] | None = None):
    try:
        _ensure_year_data_fresh(year_key)
        processor = PROCESSOR_BY_YEAR.get(year_key, PROCESSOR)
        meses = [m for m in (meses_filtrados or MESES) if m in MESES]
        if not meses:
            meses = list(MESES)
        total_medicina = float(pd.to_numeric(processor.safe_get_row('MEDICINA INTERNA', meses, silent=True), errors='coerce').fillna(0).sum())
        total_cirugia = float(pd.to_numeric(processor.safe_get_row('CIRUGIA GENERAL', meses, silent=True), errors='coerce').fillna(0).sum())
        total_neuro = float(pd.to_numeric(processor.safe_get_row('NEUROCIRUGIA', meses, silent=True), errors='coerce').fillna(0).sum())
        total_neurologia = float(pd.to_numeric(processor.safe_get_row('NEUROLOGIA', meses, silent=True), errors='coerce').fillna(0).sum())
        total = max(total_medicina + total_cirugia + total_neuro + total_neurologia, 1)

        kpi_items_esp = [
            (float(total_medicina), chart_factory.create_kpi_card("Medicina Interna", f"{int(total_medicina)}", f"{processor.formatear_porcentaje(total_medicina/total*100, False)}%", "#3498db")),
            (float(total_cirugia), chart_factory.create_kpi_card("Cirug\u00eda General", f"{int(total_cirugia)}", f"{processor.formatear_porcentaje(total_cirugia/total*100, False)}%", "#e74c3c")),
            (float(total_neuro), chart_factory.create_kpi_card("Neurocirug\u00eda", f"{int(total_neuro)}", f"{processor.formatear_porcentaje(total_neuro/total*100, False)}%", "#9b59b6")),
            (float(total_neurologia), chart_factory.create_kpi_card("Neurolog\u00eda", f"{int(total_neurologia)}", f"{processor.formatear_porcentaje(total_neurologia/total*100, False)}%", "#f39c12")),
        ]
        kpi_cards_esp = [c for _, c in sorted(kpi_items_esp, key=lambda x: x[0], reverse=True)]

        series_medicina = pd.to_numeric(processor.safe_get_row('MEDICINA INTERNA', meses), errors='coerce').fillna(0)
        series_cirugia = pd.to_numeric(processor.safe_get_row('CIRUGIA GENERAL', meses), errors='coerce').fillna(0)
        series_neuro = pd.to_numeric(processor.safe_get_row('NEUROCIRUGIA', meses), errors='coerce').fillna(0)
        series_neurologia = pd.to_numeric(processor.safe_get_row('NEUROLOGIA', meses), errors='coerce').fillna(0)
        df_especialidades = pd.DataFrame({
            'Mes': meses,
            'Medicina Interna': series_medicina.values,
            'Cirugía General': series_cirugia.values,
            'Neurocirugía': series_neuro.values,
            'Neurología': series_neurologia.values,
        })
        df_especialidades['Total'] = (
            df_especialidades['Medicina Interna']
            + df_especialidades['Cirugía General']
            + df_especialidades['Neurocirugía']
            + df_especialidades['Neurología']
        ).replace(0, 1)

        def _esp_label_vol(v: Any) -> str:
            try:
                x = float(v)
            except (TypeError, ValueError):
                return ''
            if x <= 0:
                return ''
            return str(int(round(x)))

        m_vol = df_especialidades['Medicina Interna'].astype(float)
        c_vol = df_especialidades['Cirugía General'].astype(float)
        n_vol = df_especialidades['Neurocirugía'].astype(float)
        l_vol = df_especialidades['Neurología'].astype(float)
        # Centro vertical de cada franja apilada (para etiquetar el número por especialidad/mes)
        y_mid_mi = m_vol / 2
        y_mid_cg = m_vol + c_vol / 2
        y_mid_nc = m_vol + c_vol + n_vol / 2
        y_mid_nl = m_vol + c_vol + n_vol + l_vol / 2

        fig_especialidades_area = go.Figure()
        fig_especialidades_area.add_trace(go.Scatter(
            x=df_especialidades['Mes'], y=df_especialidades['Medicina Interna'],
            name='Medicina Interna', mode='lines', stackgroup='one', line=dict(width=0.8, color='#3498db')
        ))
        fig_especialidades_area.add_trace(go.Scatter(
            x=df_especialidades['Mes'], y=df_especialidades['Cirugía General'],
            name='Cirugía General', mode='lines', stackgroup='one', line=dict(width=0.8, color='#e74c3c')
        ))
        fig_especialidades_area.add_trace(go.Scatter(
            x=df_especialidades['Mes'], y=df_especialidades['Neurocirugía'],
            name='Neurocirugía', mode='lines', stackgroup='one', line=dict(width=0.8, color='#9b59b6')
        ))
        fig_especialidades_area.add_trace(go.Scatter(
            x=df_especialidades['Mes'], y=df_especialidades['Neurología'],
            name='Neurología', mode='lines', stackgroup='one', line=dict(width=0.8, color='#f39c12')
        ))
        # Mismo color claro en todas las capas: mejor contraste que el color de la serie sobre el fondo oscuro
        _txt_font_area = dict(
            size=12,
            color='#e2e8f0',
            family='Arial',
        )
        for y_mid, series_vol in (
            (y_mid_mi, m_vol),
            (y_mid_cg, c_vol),
            (y_mid_nc, n_vol),
            (y_mid_nl, l_vol),
        ):
            fig_especialidades_area.add_trace(go.Scatter(
                x=df_especialidades['Mes'],
                y=y_mid,
                mode='text',
                text=[_esp_label_vol(v) for v in series_vol],
                textposition='middle center',
                textfont=_txt_font_area,
                showlegend=False,
                hoverinfo='skip',
                cliponaxis=False,
            ))
        fig_especialidades_area.update_layout(
            title={'text': 'Volumen Mensual por Especialidad (Área Apilada)', 'font': {'size': 13, 'color': '#f1f5f9'}, 'x': 0.5},
            transition={'duration': 0},
            plot_bgcolor='#1e293b', paper_bgcolor='#1e293b',
            autosize=True, height=360, margin=dict(l=48, r=24, t=50, b=55),
            hovermode='x unified',
            xaxis={'tickfont': {'color': '#cbd5e1', 'size': 11}, 'gridcolor': 'rgba(148, 163, 184, 0.1)'},
            yaxis={'title': 'Pacientes', 'tickfont': {'color': '#94a3b8'}, 'gridcolor': 'rgba(148, 163, 184, 0.1)'},
            legend=dict(
                orientation="h", y=-0.18, x=0.5, xanchor="center",
                bgcolor='rgba(30, 41, 59, 0.8)', font={'color': '#f1f5f9', 'size': 11},
                bordercolor='rgba(148, 163, 184, 0.3)', borderwidth=1
            )
        )

        fig_especialidades_mix = go.Figure()
        # Tipografía normal (sin negrilla): evita aspecto de “resaltador” de Arial Black
        _txt_mix = dict(color='#e2e8f0', size=11, family='Arial')
        _bar_text_kw = dict(
            insidetextanchor='middle',
            textangle=0,
            textfont=_txt_mix,
            constraintext='none',
        )
        pct_mi = (df_especialidades['Medicina Interna'] / df_especialidades['Total']) * 100
        pct_cg = (df_especialidades['Cirugía General'] / df_especialidades['Total']) * 100
        pct_nc = (df_especialidades['Neurocirugía'] / df_especialidades['Total']) * 100
        # Franja superior (Neurología): si el % es bajo, el texto dentro se rota o recorta → fuera de la barra
        pct_nl = (df_especialidades['Neurología'] / df_especialidades['Total']) * 100
        nl_vals = df_especialidades['Neurología'].astype(float)
        text_pos_nl: list[str] = []
        for p, v in zip(pct_nl, nl_vals):
            if float(v) <= 0:
                text_pos_nl.append('none')
            elif float(p) < 14.0:
                text_pos_nl.append('outside')
            else:
                text_pos_nl.append('inside')
        fig_especialidades_mix.add_trace(go.Bar(
            x=meses, y=pct_mi,
            name='Medicina Interna', marker_color='#3498db',
            text=[f"{float(v):.1f}%" if float(v) > 0 else '' for v in pct_mi],
            textposition='inside',
            **_bar_text_kw,
        ))
        fig_especialidades_mix.add_trace(go.Bar(
            x=meses, y=pct_cg,
            name='Cirugía General', marker_color='#e74c3c',
            text=[f"{float(v):.1f}%" if float(v) > 0 else '' for v in pct_cg],
            textposition='inside',
            **_bar_text_kw,
        ))
        fig_especialidades_mix.add_trace(go.Bar(
            x=meses, y=pct_nc,
            name='Neurocirugía', marker_color='#9b59b6',
            text=[f"{float(v):.1f}%" if float(v) > 0 else '' for v in pct_nc],
            textposition='inside',
            **_bar_text_kw,
        ))
        fig_especialidades_mix.add_trace(go.Bar(
            x=meses, y=pct_nl,
            name='Neurología', marker_color='#f39c12',
            text=[f"{float(v):.1f}%" if float(v) > 0 else '' for v in pct_nl],
            textposition=text_pos_nl,
            cliponaxis=False,
            **_bar_text_kw,
            insidetextfont=_txt_mix,
            outsidetextfont=_txt_mix,
        ))
        fig_especialidades_mix.update_layout(
            title={'text': 'Composición % por Mes (100% Apilada)', 'font': {'size': 13, 'color': '#f1f5f9'}, 'x': 0.5},
            transition={'duration': 0},
            barmode='stack',
            plot_bgcolor='#1e293b', paper_bgcolor='#1e293b',
            autosize=True, height=360, margin=dict(l=48, r=24, t=72, b=55),
            hovermode='x unified',
            xaxis={'tickfont': {'color': '#cbd5e1', 'size': 11}, 'gridcolor': 'rgba(148, 163, 184, 0.1)'},
            yaxis={
                'title': '% del total mensual',
                'tickfont': {'color': '#94a3b8'},
                'ticksuffix': '%',
                # Rango >100% para que las etiquetas "outside" del trazo superior no se recorten (p. ej. el 4 en Ene)
                'range': [0, 118],
                'tickmode': 'linear',
                'tick0': 0,
                'dtick': 20,
                'gridcolor': 'rgba(148, 163, 184, 0.1)',
            },
            legend=dict(
                orientation="h", y=-0.18, x=0.5, xanchor="center",
                bgcolor='rgba(30, 41, 59, 0.8)', font={'color': '#f1f5f9', 'size': 11},
                bordercolor='rgba(148, 163, 184, 0.3)', borderwidth=1
            )
        )

        return html.Div([
            html.H5("Principales Especialidades", style={'color': '#f1f5f9', 'fontWeight': '600', 'marginBottom': '12px'}),
            dbc.Row([dbc.Col(k, xs=12, md=3) for k in kpi_cards_esp], className="mb-3 g-3"),
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_especialidades_area, config=_GRAPH_CFG, style={'height': '380px'}), xs=12, className="plotly-chart plotly-chart--md")
            ], className="mb-3 g-3"),
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_especialidades_mix, config=_GRAPH_CFG, style={'height': '380px'}), xs=12, className="plotly-chart plotly-chart--md")
            ], className="g-3")
        ])
    except Exception as e:
        logger.error(f"Error en render_especialidades: {e}")
        return html.Div([html.H3(f"Error: {str(e)}", style={'color': '#e74c3c'})])


def render_mortalidad(year_key: str, meses_filtrados: list[str] | None = None):
    try:
        _ensure_year_data_fresh(year_key)
        processor = PROCESSOR_BY_YEAR.get(year_key, PROCESSOR)
        data_year = DATA_BY_YEAR.get(year_key, DATA)
        meses = [m for m in (meses_filtrados or MESES) if m in MESES]
        if not meses:
            meses = list(MESES)
        kpis_mort = processor.get_kpis_mortalidad(meses)
        if len(meses) == 1:
            rango_mort = f"{MESES_COMPLETOS.get(meses[0], meses[0])} {year_key}"
        elif len(meses) == len(MESES):
            rango_mort = processor.etiqueta_rango_periodo(year_key)
        else:
            rango_mort = f"{', '.join([MESES_COMPLETOS.get(m, m)[:3] for m in meses])} {year_key}"
        kpi_items_mort = [
            (float(kpis_mort.get('total_fallecidos') or 0), chart_factory.create_kpi_card("Total Fallecidos", f"{int(kpis_mort['total_fallecidos'])}", rango_mort, "#e74c3c")),
            (float(kpis_mort.get('mortalidad_general') or 0), chart_factory.create_kpi_card("% Mortalidad General", processor.formatear_porcentaje(kpis_mort['mortalidad_general']), rango_mort, "#c0392b")),
            (float(kpis_mort.get('fallecidos_covid') or 0), chart_factory.create_kpi_card("Fallecidos COVID-19", f"{int(kpis_mort['fallecidos_covid'])}", rango_mort, "#34495e")),
            (float(kpis_mort.get('mortalidad_trauma') or 0), chart_factory.create_kpi_card("% Mortalidad Trauma", processor.formatear_porcentaje(kpis_mort['mortalidad_trauma']), "Promedio", "#f39c12")),
        ]
        kpi_cards_mort = [c for _, c in sorted(kpi_items_mort, key=lambda x: x[0], reverse=True)]
        kpi_estancia_mort = [
            (float(kpis_mort.get('promedio_dias_estancia') or 0), chart_factory.create_kpi_card("Promedio D\u00edas Estancia", f"{kpis_mort['promedio_dias_estancia']:.1f}", "Fallecidos", "#9b59b6")),
            (float(kpis_mort.get('mediana_dias_estancia') or 0), chart_factory.create_kpi_card("Mediana D\u00edas de Estancia", f"{kpis_mort['mediana_dias_estancia']:.1f}", "Fallecidos", "#3498db")),
        ]
        kpi_cards_estancia_mort = [c for _, c in sorted(kpi_estancia_mort, key=lambda x: x[0], reverse=True)]

        fig_mort = chart_factory.create_line_chart(
            processor.serie_mortalidad_general_mensual(meses),
            '% Mortalidad General', color='#e74c3c', decimals=1
        )
        fig_trauma = chart_factory.create_grouped_bar_chart(
            ['Pacientes con diagn\u00f3stico de Trauma', 'Pacientes fallecidos por Trauma'],
            ['Pacientes Trauma', 'Fallecidos Trauma'], 'Trauma: Pacientes vs Fallecidos',
            processor=processor, meses=meses, decimals=0
        )
        df_fallecidos_plot = data_year['fallecidos']
        if isinstance(df_fallecidos_plot, pd.DataFrame) and not df_fallecidos_plot.empty and 'Mes' in df_fallecidos_plot.columns:
            meses_full = [MESES_COMPLETOS.get(m, m) for m in meses]
            df_fallecidos_plot = df_fallecidos_plot[df_fallecidos_plot['Mes'].isin(meses_full)]
        fig_piramide = chart_factory.create_pyramid_chart(df_fallecidos_plot)

        # Distribución por EPS (periodo filtrado): misma suma que Total Fallecidos
        total_ref_fallecidos = int(kpis_mort.get('total_fallecidos') or 0)
        eps_labels, eps_values = processor.get_eps_data(meses)
        if total_ref_fallecidos > 0:
            eps_labels, eps_values = DataProcessor.reconcile_fallecidos_distribution(
                eps_labels,
                list(eps_values) if eps_values else [],
                total_ref_fallecidos,
                residual_label="Sin clasificar EPS",
            )
        if eps_values:
            eps_df = pd.DataFrame({'EPS': eps_labels, 'Cantidad': eps_values})
            eps_df = eps_df.sort_values('Cantidad', ascending=False)

            def _short_eps_label(name: str) -> str:
                normalized = str(name).strip().upper()
                label_map = {
                    'ADMINISTRADORA DE RECURSOS SGSS': 'ADRES',
                    'FIDUSUARIA LA PREVISORA': 'FIDUPREVISORA',
                    'NUEVA  EPS': 'NUEVA EPS',
                    'CAPITAL SALUD': 'CAPITAL SALUD',
                }
                shortened = label_map.get(normalized, normalized)
                return shortened if len(shortened) <= 18 else f"{shortened[:17]}…"

            eps_df['EPS_DISPLAY'] = eps_df['EPS'].apply(_short_eps_label)
            total_eps = int(sum(eps_values))
            eps_color_map = {
                'FAMISANAR': '#2563eb',
                'COMPENSAR': '#ef4444',
                'SALUD TOTAL': '#10b981',
                'NUEVA EPS': '#f59e0b',
                'NUEVA  EPS': '#f59e0b',
                'FIDUSUARIA LA PREVISORA': '#8b5cf6',
                'CAPITAL SALUD': '#06b6d4',
                'CAPITAL SALUD ': '#06b6d4',
                'SANITAS': '#ec4899',
                'SURA': '#84cc16',
                'ALIANSALUD': '#f97316',
                'EMCOSALUD': '#14b8a6',
                'EMSSANAR': '#eab308',
                'SOAT': '#64748b',
                'PARTICULAR': '#22c55e',
                'ADRES': '#7c3aed',
                'ADMINISTRADORA DE RECURSOS SGSS': '#7c3aed',
                'OTRA': '#94a3b8',
                'OTRAS EPS': '#475569',
                'SIN CLASIFICAR EPS': '#64748b',
            }
            fallback_colors = ['#0ea5e9', '#f43f5e', '#a855f7', '#f97316', '#14b8a6', '#65a30d']
            resolved_eps_colors = []
            for i, original_name in enumerate(eps_df['EPS'].tolist()):
                key = str(original_name).strip().upper()
                resolved_eps_colors.append(eps_color_map.get(key, fallback_colors[i % len(fallback_colors)]))
            eps_df['Color'] = resolved_eps_colors
            eps_df['Porcentaje'] = (eps_df['Cantidad'] / max(total_eps, 1)) * 100
            fig_eps_mort = go.Figure(data=[go.Bar(
                y=eps_df['EPS_DISPLAY'],
                x=eps_df['Cantidad'],
                orientation='h',
                marker=dict(color=eps_df['Color'], line=dict(color='#0f172a', width=1.5)),
                customdata=eps_df[['EPS', 'Porcentaje']],
                text=[f"{int(v)} ({p:.1f}%)" for v, p in zip(eps_df['Cantidad'], eps_df['Porcentaje'])],
                textposition='auto',
                textfont=dict(size=10, color='white', family='Roboto'),
                hovertemplate='<b>%{customdata[0]}</b><br>Cantidad: %{x:.0f}<br>Porcentaje: %{customdata[1]:.1f}%<extra></extra>'
            )])
            fig_eps_mort.update_layout(
                title={'text': f'Distribución de fallecidos por EPS (año {year_key})', 'font': {'size': 13, 'color': '#f1f5f9'}, 'x': 0.5},
                transition={'duration': 0},
                plot_bgcolor='#1e293b', paper_bgcolor='#1e293b',
                font={'color': '#f1f5f9'},
                autosize=True, height=360, margin=dict(l=120, r=28, t=50, b=32),
                showlegend=False,
                xaxis={'tickfont': {'color': '#94a3b8'}, 'gridcolor': 'rgba(148, 163, 184, 0.1)'},
                yaxis={
                    'tickfont': {'color': '#cbd5e1', 'size': 10},
                    'automargin': True,
                    'ticklabelstandoff': 8,
                    'categoryorder': 'total ascending'
                }
            )
        else:
            fig_eps_mort = go.Figure()
            fig_eps_mort.update_layout(
                title={'text': 'Distribución de fallecidos por EPS', 'font': {'size': 14, 'color': '#f1f5f9'}, 'x': 0.5},
                plot_bgcolor='#1e293b',
                paper_bgcolor='#1e293b',
                height=380,
                margin=dict(l=20, r=20, t=60, b=20),
                xaxis={'visible': False},
                yaxis={'visible': False},
                showlegend=False
            )
            fig_eps_mort.add_annotation(
                text="<b>Sin datos de EPS</b>",
                x=0.5, y=0.5, xref='paper', yref='paper',
                showarrow=False,
                font=dict(size=14, color='#e2e8f0', family='Roboto')
            )

        # Distribución por especialidad (periodo filtrado): misma suma que Total Fallecidos
        total_medicina = float(pd.to_numeric(processor.safe_get_row('MEDICINA INTERNA', meses, silent=True), errors='coerce').fillna(0).sum())
        total_cirugia = float(pd.to_numeric(processor.safe_get_row('CIRUGIA GENERAL', meses, silent=True), errors='coerce').fillna(0).sum())
        total_neuro = float(pd.to_numeric(processor.safe_get_row('NEUROCIRUGIA', meses, silent=True), errors='coerce').fillna(0).sum())
        total_neurologia = float(pd.to_numeric(processor.safe_get_row('NEUROLOGIA', meses, silent=True), errors='coerce').fillna(0).sum())
        espec_labels = ['Medicina Interna', 'Cirugía General', 'Neurocirugía', 'Neurología']
        espec_values = [
            float(total_medicina or 0),
            float(total_cirugia or 0),
            float(total_neuro or 0),
            float(total_neurologia or 0),
        ]
        if total_ref_fallecidos > 0:
            espec_labels, espec_values = DataProcessor.reconcile_fallecidos_distribution(
                espec_labels,
                espec_values,
                total_ref_fallecidos,
                residual_label="Otras especialidades",
            )
        if total_ref_fallecidos > 0 and espec_values and any(v > 0 for v in espec_values):
            espec_colors = ['#3498db', '#e74c3c', '#9b59b6', '#f39c12', '#64748b']
            fig_espec_mort = chart_factory.create_donut_chart(
                labels=espec_labels,
                values=espec_values,
                title=f'Distribución de fallecidos por Especialidad (año {year_key})',
                colors=espec_colors[: len(espec_labels)],
                auto_sort=True
            )
        else:
            fig_espec_mort = go.Figure()
            fig_espec_mort.update_layout(
                title={'text': 'Distribución de fallecidos por Especialidad', 'font': {'size': 14, 'color': '#f1f5f9'}, 'x': 0.5},
                plot_bgcolor='#1e293b',
                paper_bgcolor='#1e293b',
                height=380,
                margin=dict(l=20, r=20, t=60, b=20),
                xaxis={'visible': False},
                yaxis={'visible': False},
                showlegend=False
            )
            fig_espec_mort.add_annotation(
                text="<b>Sin datos por especialidad</b>",
                x=0.5, y=0.5, xref='paper', yref='paper',
                showarrow=False,
                font=dict(size=14, color='#e2e8f0', family='Roboto')
            )

        return html.Div([
            dbc.Row([dbc.Col(k, xs=12, md=3) for k in kpi_cards_mort], className="mb-3 g-3"),
            html.H6("D\u00edas de Estancia - Fallecidos", style={'color': '#f1f5f9', 'fontWeight': '600', 'marginBottom': '10px', 'marginTop': '8px'}),
            dbc.Row([dbc.Col(k, xs=12, md=6) for k in kpi_cards_estancia_mort], className="mb-3 g-3"),
            dbc.Row([
                dbc.Col(xs=12, md=2),
                dbc.Col(dcc.Graph(figure=fig_mort, config=_GRAPH_CFG, style={'height': '350px'}), xs=12, md=8, className="plotly-chart plotly-chart--sm"),
                dbc.Col(xs=12, md=2)
            ], className="mb-3", justify="center"),
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_trauma, config=_GRAPH_CFG, style={'height': '380px'}), xs=12, className="plotly-chart plotly-chart--md")
            ], className="mb-4 g-3"),
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_piramide, config=_GRAPH_CFG, style={'height': '400px'}), xs=12, className="plotly-chart plotly-chart--lg")
            ], className="g-3"),
            html.H5("Distribución de fallecidos por EPS y Especialidad", style={'color': '#f1f5f9', 'fontWeight': '600', 'marginTop': '18px', 'marginBottom': '12px', 'textAlign': 'center'}),
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_eps_mort, config=_GRAPH_CFG, style={'height': '380px'}), xs=12, md=7, className="plotly-chart plotly-chart--md"),
                dbc.Col(dcc.Graph(figure=fig_espec_mort, config=_GRAPH_CFG, style={'height': '380px'}), xs=12, md=5, className="plotly-chart plotly-chart--md"),
            ], className="mb-4 g-3")
        ])
    except Exception as e:
        logger.error(f"Error en render_mortalidad: {e}")
        return html.Div([html.H3(f"Error: {str(e)}", style={'color': '#e74c3c'})])


# Callback para Resumen - Solo se aplica al pulsar Aplicar o Todos
@app.callback(
    [Output('filtro-mes-inicio', 'value'),
     Output('resumen-content', 'children'),
     Output('selected-months-store', 'data'),
     Output('selected-year-store', 'data')],
    [Input('btn-todos-meses', 'n_clicks'),
     Input('btn-aplicar-filtros', 'n_clicks'),
     Input('tabs', 'active_tab')],
    [State('filtro-anio', 'value'),
     State('filtro-mes-inicio', 'value'),
     State('comparar-anios', 'value')]
)
def actualizar_resumen(_btn_todos, _btn_aplicar, active_tab, anio_seleccionado, meses_seleccionados, comparar_anios):
    # Solo procesar cuando estamos en Resumen (evita actualizar componente inexistente)
    if active_tab != 'tab-resumen':
        return dash.no_update, dash.no_update, dash.no_update, dash.no_update

    ctx = dash.callback_context
    triggered_id = ctx.triggered[0]['prop_id'].split('.')[0] if ctx.triggered else None
    # region agent log
    _debug_log(
        "pre-fix",
        "H1",
        "dashboard_uci.py:actualizar_resumen:entry",
        "Resumen callback triggered",
        {
            "triggeredId": triggered_id,
            "anioSeleccionado": anio_seleccionado,
            "mesesSeleccionados": meses_seleccionados,
        },
    )
    # endregion
    meses_output_value = dash.no_update
    if triggered_id == 'btn-todos-meses':
        # Selecciona explícitamente todos los meses en el dropdown
        meses_seleccionados = MESES.copy()
        meses_output_value = MESES.copy()
    try:
        year_key = str(anio_seleccionado) if anio_seleccionado is not None else DEFAULT_YEAR
        if year_key not in PROCESSOR_BY_YEAR:
            year_key = DEFAULT_YEAR
        _ensure_year_data_fresh(year_key)

        processor = PROCESSOR_BY_YEAR[year_key]
        data_actual = DATA_BY_YEAR[year_key]
        # region agent log
        _debug_log(
            "pre-fix",
            "H3",
            "dashboard_uci.py:actualizar_resumen:year_resolution",
            "Year resolved for callback",
            {
                "resolvedYear": year_key,
                "fallbackUsed": str(anio_seleccionado) != year_key if anio_seleccionado is not None else True,
                "morbilidadRows": int(data_actual.get("morbilidad", pd.DataFrame()).shape[0]),
            },
        )
        # endregion

        if meses_seleccionados and isinstance(meses_seleccionados, list) and len(meses_seleccionados) > 0:
            meses_filtrados = [m for m in meses_seleccionados if m in MESES]
            if not meses_filtrados:
                meses_filtrados = MESES
        else:
            meses_filtrados = MESES

        if len(meses_filtrados) == 1:
            rango_texto = MESES_COMPLETOS.get(meses_filtrados[0], meses_filtrados[0])
        elif len(meses_filtrados) == len(MESES):
            rango_texto = processor.etiqueta_rango_periodo(year_key)
        else:
            rango_texto = ", ".join([MESES_COMPLETOS.get(m, m)[:3] for m in meses_filtrados])

        rango_texto = f"{rango_texto} {year_key}"
        year_store_output = dash.no_update
        if triggered_id in ('btn-aplicar-filtros', 'btn-todos-meses'):
            year_store_output = year_key

        kpis = processor.get_kpis_resumen(meses_filtrados)
        # region agent log
        _debug_log(
            "pre-fix",
            "H4",
            "dashboard_uci.py:actualizar_resumen:kpis",
            "Resumen KPI snapshot",
            {
                "year": year_key,
                "mesesFiltrados": meses_filtrados,
                "totalIngresos": float(kpis.get("total_ingresos", 0) or 0),
                "mortalidad": float(kpis.get("mortalidad", 0) or 0),
            },
        )
        # endregion
        kpi_items_r = [
            (float(kpis.get('total_ingresos') or 0), chart_factory.create_kpi_card("Total Ingresos", f"{int(kpis['total_ingresos'])}", rango_texto, "#6366f1")),
            (float(kpis.get('promedio_ocupacional') or 0), chart_factory.create_kpi_card("% Ocupaci\u00f3n Promedio", processor.formatear_porcentaje(kpis['promedio_ocupacional']), rango_texto, "#3b82f6")),
            (float(kpis.get('mortalidad') or 0), chart_factory.create_kpi_card("% Mortalidad General", processor.formatear_porcentaje(kpis['mortalidad']), f"Total {rango_texto}", "#ef4444")),
            (float(kpis.get('promedio_estancia') or 0), chart_factory.create_kpi_card("Estancia Promedio", f"{kpis['promedio_estancia']:.2f}", f"D\u00edas - {rango_texto}", "#10b981")),
            (float(kpis.get('giro_cama') or 0), chart_factory.create_kpi_card("Giro cama", f"{kpis['giro_cama']:.2f}", f"Promedio {rango_texto}", "#8b5cf6")),
        ]
        kpi_cards_r = [c for _, c in sorted(kpi_items_r, key=lambda x: x[0], reverse=True)]

        # EPS (filtrado por meses seleccionados)
        eps_data = []
        for eps in EPS_LIST:
            eps_row = processor.safe_get_row(eps, meses_filtrados, silent=True)
            if isinstance(eps_row, pd.Series):
                eps_total = float(pd.to_numeric(eps_row, errors='coerce').fillna(0).sum())
            else:
                eps_total = 0.0
            if eps_total > 0:
                eps_data.append((eps.strip(), eps_total))
        eps_data.sort(key=lambda item: item[1], reverse=True)
        eps_labels = [item[0] for item in eps_data]
        eps_values = [item[1] for item in eps_data]
        if eps_values:
            eps_df = pd.DataFrame({'EPS': eps_labels, 'Cantidad': eps_values})
            eps_df = eps_df.sort_values('Cantidad', ascending=False)
            def _short_eps_label(name: str) -> str:
                normalized = str(name).strip().upper()
                label_map = {
                    'ADMINISTRADORA DE RECURSOS SGSS': 'ADRES',
                    'FIDUSUARIA LA PREVISORA': 'FIDUPREVISORA',
                    'NUEVA  EPS': 'NUEVA EPS',
                    'CAPITAL SALUD': 'CAPITAL SALUD',
                }
                shortened = label_map.get(normalized, normalized)
                return shortened if len(shortened) <= 18 else f"{shortened[:17]}…"

            eps_df['EPS_DISPLAY'] = eps_df['EPS'].apply(_short_eps_label)
            total_eps = int(sum(eps_values))
            eps_color_map = {
                'FAMISANAR': '#2563eb',
                'COMPENSAR': '#ef4444',
                'SALUD TOTAL': '#10b981',
                'NUEVA EPS': '#f59e0b',
                'NUEVA  EPS': '#f59e0b',
                'FIDUSUARIA LA PREVISORA': '#8b5cf6',
                'CAPITAL SALUD': '#06b6d4',
                'CAPITAL SALUD ': '#06b6d4',
                'SANITAS': '#ec4899',
                'SURA': '#84cc16',
                'ALIANSALUD': '#f97316',
                'EMCOSALUD': '#14b8a6',
                'EMSSANAR': '#eab308',
                'SOAT': '#64748b',
                'PARTICULAR': '#22c55e',
                'ADRES': '#7c3aed',
                'ADMINISTRADORA DE RECURSOS SGSS': '#7c3aed',
                'OTRA': '#94a3b8',
                'OTRAS EPS': '#475569',
            }
            fallback_colors = ['#0ea5e9', '#f43f5e', '#a855f7', '#f97316', '#14b8a6', '#65a30d']
            resolved_eps_colors = []
            for i, original_name in enumerate(eps_df['EPS'].tolist()):
                key = str(original_name).strip().upper()
                resolved_eps_colors.append(eps_color_map.get(key, fallback_colors[i % len(fallback_colors)]))
            eps_df['Color'] = resolved_eps_colors
            eps_df['Porcentaje'] = (eps_df['Cantidad'] / max(total_eps, 1)) * 100
            fig_eps = go.Figure(data=[go.Bar(
                y=eps_df['EPS_DISPLAY'],
                x=eps_df['Cantidad'],
                orientation='h',
                marker=dict(color=eps_df['Color'], line=dict(color='#0f172a', width=1.5)),
                customdata=eps_df[['EPS', 'Porcentaje']],
                text=[f"{int(v)} ({p:.1f}%)" for v, p in zip(eps_df['Cantidad'], eps_df['Porcentaje'])],
                textposition='auto',
                textfont=dict(size=10, color='white', family='Roboto'),
                hovertemplate='<b>%{customdata[0]}</b><br>Cantidad: %{x:.0f}<br>Porcentaje: %{customdata[1]:.1f}%<extra></extra>'
            )])
            fig_eps.update_layout(
                title={'text': f'Distribuci\u00f3n por EPS - Barras (Total: {total_eps:,})', 'font': {'size': 13, 'color': '#f1f5f9'}, 'x': 0.5},
                transition={'duration': 0},
                plot_bgcolor='#1e293b', paper_bgcolor='#1e293b',
                font={'color': '#f1f5f9'},
                autosize=True, height=480, margin=dict(l=120, r=28, t=50, b=32),
                showlegend=False,
                xaxis={'tickfont': {'color': '#94a3b8'}, 'gridcolor': 'rgba(148, 163, 184, 0.1)'},
                yaxis={
                    'tickfont': {'color': '#cbd5e1', 'size': 10},
                    'automargin': True,
                    'ticklabelstandoff': 8,
                    'categoryorder': 'total ascending'
                }
            )
        else:
            fig_eps = go.Figure()
            fig_eps.update_layout(
                title={'text': 'Distribución por EPS', 'font': {'size': 14, 'color': '#f1f5f9'}, 'x': 0.5},
                plot_bgcolor='#1e293b',
                paper_bgcolor='#1e293b',
                height=520,
                margin=dict(l=20, r=20, t=60, b=20),
                xaxis={'visible': False},
                yaxis={'visible': False},
                showlegend=False
            )
            fig_eps.add_annotation(
                text="<b>Sin datos de EPS</b>",
                x=0.5, y=0.5, xref='paper', yref='paper',
                showarrow=False,
                font=dict(size=14, color='#e2e8f0', family='Roboto')
            )

        # Bloque opcional de comparación entre años (tabla simple)
        compare_enabled = isinstance(comparar_anios, list) and 'compare' in comparar_anios
        compare_block = []
        if compare_enabled:
            available_years = sorted(PROCESSOR_BY_YEAR.keys())
            if len(available_years) >= 2:
                compare_rows = _build_compare_rows(meses_filtrados)
                period_label = (
                    MESES_COMPLETOS.get(meses_filtrados[0], meses_filtrados[0])
                    if len(meses_filtrados) == 1
                    else (processor.etiqueta_rango_periodo(year_key) if len(meses_filtrados) == len(MESES)
                          else ", ".join([MESES_COMPLETOS.get(m, m)[:3] for m in meses_filtrados]))
                )

                header_style = {
                    'backgroundColor': 'rgba(15,23,42,0.98)',
                    'color': '#e5e7eb',
                    'fontSize': '11px',
                    'fontWeight': '700',
                    'borderBottom': '1px solid rgba(148,163,184,0.5)',
                    'borderRight': '1px solid rgba(55,65,81,0.7)',
                    'textAlign': 'center',
                    'padding': '6px 10px',
                    'whiteSpace': 'normal',
                    'textTransform': 'uppercase',
                    'letterSpacing': '0.04em'
                }

                body_style = {
                    'backgroundColor': 'rgba(15,23,42,0.98)',
                    'borderTop': 'none',
                    'borderBottom': '1px solid rgba(55,65,81,0.7)',
                    'borderRight': '1px solid rgba(55,65,81,0.7)',
                    'fontSize': '12px',
                    'color': '#e5e7eb',
                    'textAlign': 'center',
                    'padding': '5px 10px',
                    'whiteSpace': 'nowrap'
                }

                compare_table = dbc.Table(
                    [
                        html.Thead(html.Tr([
                            html.Th("Año", style=header_style),
                            html.Th("Total ingresos", style=header_style),
                            html.Th("% Ocupación", style=header_style),
                            html.Th("% Mortalidad", style=header_style),
                            html.Th("Estancia prom. (días)", style=header_style),
                            html.Th("Giro cama prom.", style=header_style),
                            html.Th("Total días estancia", style=header_style),
                            html.Th("Total fallecidos", style=header_style),
                            html.Th("Eventos adversos", style=header_style),
                            html.Th("Gérmenes aislados", style=dict(header_style, minWidth='110px')),
                            html.Th("Casos morbilidad", style=dict(header_style, minWidth='115px')),
                        ])),
                        html.Tbody([
                            html.Tr([
                                html.Td(row["year"], style=body_style),
                                html.Td(f"{row['ingresos']:,}".replace(",", "."), style=body_style),
                                html.Td(f"{row['ocupacion']:.1f}%", style=body_style),
                                html.Td(f"{row['mortalidad']:.1f}%", style=body_style),
                                html.Td(f"{row['estancia']:.2f}", style=body_style),
                                html.Td(f"{row['giro_cama']:.2f}", style=body_style),
                                html.Td(f"{row['total_dias']:.0f}", style=body_style),
                                html.Td(f"{row['fallecidos']}", style=body_style),
                                html.Td(f"{row['total_eventos']}", style=body_style),
                                html.Td(f"{row['total_germenes']}", style=body_style),
                                html.Td(f"{row['total_diag']}", style=body_style),
                            ]) for row in compare_rows
                        ])
                    ],
                    bordered=True,
                    hover=True,
                    size="sm",
                    style={
                        'marginTop': '4px',
                        'marginBottom': '0',
                        'backgroundColor': 'transparent',
                        'width': '100%',
                        'borderCollapse': 'collapse',
                        'tableLayout': 'fixed'
                    }
                )

                def _compare_mobile_metric_rows(row: dict[str, Any]) -> list[tuple[str, str]]:
                    """Pares (etiqueta, valor) para vista móvil en tarjetas."""
                    return [
                        ("Total ingresos", f"{row['ingresos']:,}".replace(",", ".")),
                        ("% Ocupación", f"{row['ocupacion']:.1f}%"),
                        ("% Mortalidad", f"{row['mortalidad']:.1f}%"),
                        ("Estancia prom. (días)", f"{row['estancia']:.2f}"),
                        ("Giro cama prom.", f"{row['giro_cama']:.2f}"),
                        ("Total días estancia", f"{row['total_dias']:.0f}"),
                        ("Total fallecidos", str(row["fallecidos"])),
                        ("Eventos adversos", str(row["total_eventos"])),
                        ("Gérmenes aislados", str(row["total_germenes"])),
                        ("Casos morbilidad", str(row["total_diag"])),
                    ]

                compare_table_mobile = html.Div(
                    [
                        dbc.Card(
                            dbc.CardBody([
                                html.H6(
                                    f"Año {row['year']}",
                                    style={
                                        'color': '#f8fafc', 'fontWeight': '700', 'fontSize': '16px',
                                        'marginBottom': '10px', 'paddingBottom': '10px',
                                        'borderBottom': '1px solid rgba(148, 163, 184, 0.35)',
                                    },
                                ),
                                *[
                                    html.Div(
                                        [
                                            html.Span(
                                                lbl,
                                                style={
                                                    'color': '#94a3b8', 'fontSize': '12px', 'lineHeight': '1.35',
                                                    'paddingRight': '10px', 'flex': '1', 'minWidth': '0',
                                                },
                                            ),
                                            html.Span(
                                                val,
                                                style={
                                                    'color': '#e5e7eb', 'fontSize': '13px', 'fontWeight': '600',
                                                    'textAlign': 'right', 'flexShrink': '0',
                                                },
                                            ),
                                        ],
                                        style={
                                            'display': 'flex', 'justifyContent': 'space-between',
                                            'alignItems': 'flex-start', 'gap': '8px',
                                            'padding': '9px 0',
                                            'borderBottom': '1px solid rgba(71, 85, 105, 0.45)',
                                        },
                                    )
                                    for lbl, val in _compare_mobile_metric_rows(row)
                                ],
                            ]),
                            className="compare-mobile-year-card",
                            style={
                                'background': 'rgba(15, 23, 42, 0.96)',
                                'border': '1px solid rgba(148, 163, 184, 0.28)',
                                'borderRadius': '10px',
                            },
                        )
                        for row in compare_rows
                    ],
                    className="d-md-none compare-table-mobile-stack",
                    style={'marginTop': '10px'},
                )

                compare_df = pd.DataFrame(compare_rows)
                year_labels = compare_df["year"].tolist()

                # Paleta fija: un color distinto por métrica en todo el bloque comparativo (sin repetir)
                _compare_color_keys = [
                    "total_dias", "total_diag",
                    "ingresos", "fallecidos", "total_eventos", "total_germenes",
                    "ocupacion", "mortalidad", "estancia", "giro_cama",
                ]
                # Paleta fija de alto contraste para que cada métrica se diferencie claramente.
                _color_cmp = {
                    "total_dias": "#2563eb",      # azul fuerte
                    "total_diag": "#f59e0b",      # ámbar
                    "ingresos": "#10b981",        # verde
                    "fallecidos": "#ef4444",      # rojo
                    "total_eventos": "#8b5cf6",   # violeta
                    "total_germenes": "#06b6d4",  # cian
                    "ocupacion": "#3b82f6",       # azul medio
                    "mortalidad": "#f97316",      # naranja intenso
                    "estancia": "#84cc16",        # lima
                    "giro_cama": "#ec4899",       # fucsia
                }

                def _range_y_vol(keys: list[str]) -> list[float]:
                    m = float(compare_df[keys].max().max()) if keys else 0.0
                    if m <= 0:
                        return [0.0, 1.0]
                    return [0.0, m * 1.22]

                def _range_y_rates(keys: list[str]) -> list[float]:
                    m = float(compare_df[keys].max().max()) if keys else 0.0
                    if m <= 0:
                        return [0.0, 1.0]
                    return [0.0, m * 1.22]

                def _fig_compare_bars_vol(title: str, specs: list[tuple[str, str]], height: int) -> go.Figure:
                    fig = go.Figure()
                    keys = [k for k, _ in specs]
                    for key, label in specs:
                        fig.add_trace(go.Bar(
                            x=year_labels,
                            y=compare_df[key],
                            name=label,
                            marker_color=_color_cmp[key],
                            text=[f"{int(v):,}".replace(",", ".") for v in compare_df[key]],
                            textposition='outside',
                            textfont=dict(color='#e2e8f0', size=10),
                            cliponaxis=False,
                            hovertemplate=f"<b>{label}</b><br>Año: %{{x}}<br>Valor: %{{y:.0f}}<extra></extra>"
                        ))
                    fig.update_layout(
                        title={'text': title, 'font': {'size': 13, 'color': '#f1f5f9'}, 'x': 0.5},
                        barmode='group',
                        plot_bgcolor='#1e293b', paper_bgcolor='#1e293b',
                        height=height, margin=dict(l=45, r=20, t=48, b=45),
                        xaxis={'tickfont': {'color': '#cbd5e1'}, 'gridcolor': 'rgba(148, 163, 184, 0.1)'},
                        yaxis={
                            'tickfont': {'color': '#94a3b8'},
                            'gridcolor': 'rgba(148, 163, 184, 0.1)',
                            'rangemode': 'tozero'
                        },
                        legend=dict(
                            orientation="h", y=-0.22, x=0.5, xanchor="center",
                            bgcolor='rgba(30, 41, 59, 0.8)', font={'color': '#f1f5f9', 'size': 11},
                            bordercolor='rgba(148, 163, 184, 0.3)', borderwidth=1
                        ),
                    )
                    fig.update_yaxes(range=_range_y_vol(keys))
                    return fig

                def _fig_compare_radar(compare_df: pd.DataFrame, year_labels: list) -> go.Figure:
                    """Gráfico de araña/radar: cada eje es una métrica; radio = % del máximo entre años."""
                    radar_specs = [
                        ("total_dias", "Días estancia"),
                        ("total_diag", "Morbilidad"),
                        ("ingresos", "Ingresos"),
                        ("fallecidos", "Fallecidos"),
                        ("total_eventos", "Eventos adv."),
                        ("total_germenes", "Gérmenes"),
                        ("ocupacion", "% Ocupación"),
                        ("mortalidad", "% Mortalidad"),
                        ("estancia", "Estancia (d)"),
                        ("giro_cama", "Giro cama"),
                    ]
                    theta_labels = [lbl for _, lbl in radar_specs]
                    theta = theta_labels + [theta_labels[0]]
                    fig_r = go.Figure()
                    # Colores muy contrastantes sobre fondo oscuro (naranja vs verde lima; líneas gruesas)
                    year_line_colors = ("#fb923c", "#a3e635")
                    year_fill_colors = ("rgba(251, 146, 60, 0.35)", "rgba(163, 230, 53, 0.35)")
                    for yi, year in enumerate(year_labels):
                        sub = compare_df[compare_df["year"].astype(str) == str(year)]
                        if sub.empty:
                            continue
                        row = sub.iloc[0]
                        r_norm: list[float] = []
                        # Por vértice: etiqueta, valor real, máximo entre años en esa métrica, % relativo
                        meta_rows: list[tuple[str, float, float, float]] = []
                        for key, lbl in radar_specs:
                            v = float(row[key])
                            colmax = float(compare_df[key].max())
                            rel_pct = (v / colmax) * 100.0 if colmax > 0 else 0.0
                            meta_rows.append((lbl, v, colmax, rel_pct))
                            r_norm.append(rel_pct)
                        r_norm_closed = r_norm + [r_norm[0]]
                        meta_closed = meta_rows + [meta_rows[0]]
                        customdata = [[a, b, c, d] for (a, b, c, d) in meta_closed]
                        yc = year_line_colors[yi % len(year_line_colors)]
                        fig_r.add_trace(go.Scatterpolar(
                            r=r_norm_closed,
                            theta=theta,
                            name=str(year),
                            mode="lines+markers",
                            line=dict(color=yc, width=3.5),
                            marker=dict(
                                size=6,
                                color=yc,
                                opacity=0.95,
                                line=dict(color="#f8fafc", width=0.8),
                            ),
                            fillcolor=year_fill_colors[yi % len(year_fill_colors)],
                            fill="toself",
                            opacity=1,
                            hoveron="points+fills",
                            customdata=customdata,
                            hovertemplate=(
                                "<b style='font-size:14px'>%{customdata[0]}</b><br>"
                                "<span style='color:#94a3b8'>Año %{fullData.name}</span><br><br>"
                                "Valor: <b>%{customdata[1]:,.2f}</b><br>"
                                "Máximo entre años (esta métrica): %{customdata[2]:,.2f}<br>"
                                "Relativo al máximo: <b>%{customdata[3]:.1f}%</b><extra></extra>"
                            ),
                        ))
                    fig_r.update_layout(
                        title={
                            "text": "Comparación en gráfico de araña (radar)",
                            "x": 0.5,
                            "font": {"size": 15, "color": "#f1f5f9"},
                        },
                        polar=dict(
                            # Dominio amplio: gráfico más grande; márgenes del papel reservan espacio a etiquetas
                            domain=dict(x=[0.06, 0.94], y=[0.04, 0.88]),
                            bgcolor="#1e293b",
                            radialaxis=dict(
                                visible=True,
                                range=[0, 100],
                                dtick=25,
                                tick0=0,
                                # Etiquetas 0–100: color cálido y tamaño mayor (evita confundirse con la rejilla)
                                tickfont=dict(color="#fef08a", size=13, family="Roboto"),
                                gridcolor="rgba(71, 85, 105, 0.75)",
                                gridwidth=1,
                                linecolor="rgba(148, 163, 184, 0.55)",
                                linewidth=1,
                            ),
                            angularaxis=dict(
                                tickfont=dict(color="#e2e8f0", size=10, family="Roboto"),
                                direction="clockwise",
                                gridcolor="rgba(71, 85, 105, 0.55)",
                                ticklen=10,
                                ticks="outside",
                                showline=True,
                                linecolor="rgba(148, 163, 184, 0.35)",
                            ),
                        ),
                        paper_bgcolor="#1e293b",
                        height=780,
                        margin=dict(l=72, r=72, t=64, b=88),
                        legend=dict(
                            orientation="h",
                            y=-0.11,
                            yanchor="top",
                            x=0.5,
                            xanchor="center",
                            font=dict(color="#f1f5f9", size=12),
                            bgcolor="rgba(30, 41, 59, 0.85)",
                            bordercolor="rgba(148, 163, 184, 0.3)",
                            borderwidth=1,
                        ),
                        hovermode="closest",
                        hoverlabel=dict(
                            bgcolor="rgba(15, 23, 42, 0.96)",
                            bordercolor="rgba(148, 163, 184, 0.55)",
                            font_size=13,
                            font_family="Roboto",
                            font_color="#f1f5f9",
                            align="left",
                        ),
                        font=dict(color="#f1f5f9"),
                    )
                    return fig_r

                # Volúmenes: dos escalas (altos vs resto) para que no aplaste las barras pequeñas
                vol_specs_large = [
                    ("total_dias", "Total días estancia"),
                    ("total_diag", "Casos morbilidad (anual)"),
                ]
                vol_specs_small = [
                    ("ingresos", "Total ingresos"),
                    ("fallecidos", "Total fallecidos"),
                    ("total_eventos", "Eventos adversos"),
                    ("total_germenes", "Gérmenes aislados (anual)"),
                ]
                fig_compare_vol_hi = _fig_compare_bars_vol(
                    "Comparación visual - Volúmenes altos (días estancia / morbilidad)",
                    vol_specs_large,
                    340,
                )
                fig_compare_vol_lo = _fig_compare_bars_vol(
                    "Comparación visual - Volúmenes restantes",
                    vol_specs_small,
                    360,
                )

                fig_compare_rates = go.Figure()
                rate_specs = [
                    ("ocupacion", "% Ocupación"),
                    ("mortalidad", "% Mortalidad"),
                    ("estancia", "Estancia prom. (días)"),
                    ("giro_cama", "Giro cama prom."),
                ]
                for key, label in rate_specs:
                    fig_compare_rates.add_trace(go.Bar(
                        x=year_labels,
                        y=compare_df[key],
                        name=label,
                        marker_color=_color_cmp[key],
                        text=[f"{v:.1f}" for v in compare_df[key]],
                        textposition='outside',
                        textfont=dict(color='#e2e8f0', size=10),
                        cliponaxis=False,
                        hovertemplate=f"<b>{label}</b><br>Año: %{{x}}<br>Valor: %{{y:.1f}}<extra></extra>"
                    ))
                fig_compare_rates.update_layout(
                    title={'text': 'Comparación visual - Tasas y promedios', 'font': {'size': 13, 'color': '#f1f5f9'}, 'x': 0.5},
                    barmode='group',
                    plot_bgcolor='#1e293b', paper_bgcolor='#1e293b',
                    height=370, margin=dict(l=45, r=20, t=48, b=45),
                    xaxis={'tickfont': {'color': '#cbd5e1'}, 'gridcolor': 'rgba(148, 163, 184, 0.1)'},
                    yaxis={
                        'tickfont': {'color': '#94a3b8'},
                        'gridcolor': 'rgba(148, 163, 184, 0.1)',
                        'rangemode': 'tozero'
                    },
                    legend=dict(
                        orientation="h", y=-0.25, x=0.5, xanchor="center",
                        bgcolor='rgba(30, 41, 59, 0.8)', font={'color': '#f1f5f9', 'size': 11},
                        bordercolor='rgba(148, 163, 184, 0.3)', borderwidth=1
                    ),
                )
                fig_compare_rates.update_yaxes(range=_range_y_rates([k for k, _ in rate_specs]))

                fig_compare_radar = _fig_compare_radar(compare_df, year_labels)

                compare_block = [
                    html.H6("Comparación por año (meses seleccionados)", style={
                        'color': '#f8fafc', 'fontWeight': '700', 'marginTop': '6px', 'fontSize': '17px', 'letterSpacing': '0.2px'
                    }),
                    html.P(
                        f"Periodo aplicado: {period_label}",
                        style={'color': '#e2e8f0', 'fontSize': '13px', 'margin': '4px 0 0 0', 'fontWeight': '500'}
                    ),
                    html.P(
                        "Nota: Gérmenes y morbilidad ahora también siguen el filtro de meses cuando la hoja trae separadores mensuales.",
                        style={'color': '#cbd5e1', 'fontSize': '12px', 'margin': '6px 0 0 0'}
                    ),
                    html.P(
                        "Exportación: usa los botones de comparación junto al selector \"Comparar 2025 vs 2026\".",
                        style={'color': '#cbd5e1', 'fontSize': '12px', 'margin': '6px 0 0 0'}
                    ),
                    html.Div([
                        html.Div(
                            compare_table,
                            style={
                                'overflowX': 'auto',
                                'WebkitOverflowScrolling': 'touch',
                                'marginTop': '8px',
                                'borderRadius': '8px',
                            },
                            className="d-none d-md-block table-responsive-mobile",
                        ),
                        compare_table_mobile,
                    ]),
                    html.Div([
                        dbc.Card(
                            dbc.CardBody([
                                html.Div(row["year"], style={'fontWeight': '700', 'color': '#f1f5f9', 'marginBottom': '8px'}),
                                html.Div(f"Ingresos: {row['ingresos']:,}".replace(",", "."), style={'color': '#cbd5e1', 'fontSize': '12px'}),
                                html.Div(f"Fallecidos: {row['fallecidos']}", style={'color': '#cbd5e1', 'fontSize': '12px'}),
                                html.Div(f"Mortalidad: {row['mortalidad']:.1f}%", style={'color': '#cbd5e1', 'fontSize': '12px'}),
                                html.Div(f"Eventos: {row['total_eventos']}", style={'color': '#cbd5e1', 'fontSize': '12px'}),
                            ]),
                            style={'background': 'rgba(15,23,42,0.8)', 'border': '1px solid rgba(148,163,184,0.25)'}
                        ) for row in compare_rows
                    ], style={'display': 'grid', 'gridTemplateColumns': 'repeat(auto-fit, minmax(180px, 1fr))', 'gap': '10px', 'marginTop': '10px'})
                    ,
                    html.Div([
                        html.Span("Leyenda única:", style={'color': '#e5e7eb', 'fontSize': '12px', 'marginRight': '10px'}),
                        *[
                            html.Span([
                                html.Span(style={'display': 'inline-block', 'width': '10px', 'height': '10px', 'background': _color_cmp[k], 'marginRight': '6px', 'borderRadius': '2px'}),
                                lbl
                            ], style={'color': '#cbd5e1', 'fontSize': '11px', 'marginRight': '12px'})
                            for k, lbl in [
                                ('total_dias', 'Días estancia'),
                                ('total_diag', 'Morbilidad'),
                                ('ingresos', 'Ingresos'),
                                ('fallecidos', 'Fallecidos'),
                                ('total_eventos', 'Eventos'),
                                ('total_germenes', 'Gérmenes'),
                                ('ocupacion', '% Ocupación'),
                                ('mortalidad', '% Mortalidad'),
                                ('estancia', 'Estancia prom.'),
                                ('giro_cama', 'Giro cama'),
                            ]
                        ],
                    ], style={'marginTop': '12px', 'marginBottom': '2px', 'display': 'flex', 'flexWrap': 'wrap', 'alignItems': 'center'}),
                    dbc.Row([
                        dbc.Col(
                            dcc.Graph(figure=fig_compare_radar, config=_GRAPH_CFG_RADAR, style={'height': '780px'}),
                            xs=12,
                            className="px-1 py-2 graph-radar-compare plotly-chart",
                        )
                    ], className="mb-2 g-3"),
                    dbc.Row([
                        dbc.Col(dcc.Graph(figure=fig_compare_vol_hi, config=_GRAPH_CFG, style={'height': '340px'}), xs=12, className="plotly-chart plotly-chart--cmp")
                    ], className="mt-3 mb-2 g-3"),
                    dbc.Row([
                        dbc.Col(dcc.Graph(figure=fig_compare_vol_lo, config=_GRAPH_CFG, style={'height': '360px'}), xs=12, className="plotly-chart plotly-chart--cmp")
                    ], className="mb-2 g-3"),
                    dbc.Row([
                        dbc.Col(dcc.Graph(figure=fig_compare_rates, config=_GRAPH_CFG, style={'height': '370px'}), xs=12, className="plotly-chart plotly-chart--cmp")
                    ], className="mb-1 g-3"),
                ]
                # Modo comparación exclusivo: si está activo, oculta el resto del resumen.
                return meses_output_value, html.Div([
                    html.Div(
                        compare_block,
                        style={'padding': '4px 0 8px', 'marginBottom': '14px'}
                    )
                ]), meses_filtrados, year_store_output

        # Ocupacional
        datos_ocup = processor.safe_get_row('% OCUPACIONAL', meses_filtrados)
        fig_ocup = go.Figure()
        fig_ocup.add_trace(go.Scatter(
            x=meses_filtrados, y=datos_ocup.values,
            mode='lines+markers',
            line=dict(color='#10b981', width=2),
            marker=dict(size=8, color='#10b981'),
            hovertemplate='<b>%{x}</b><br>Ocupación: %{y:.1f}%<extra></extra>'
        ))
        fig_ocup.update_layout(
            title={'text': '% Ocupacional Mensual', 'font': {'size': 14, 'color': '#f1f5f9'}, 'x': 0.5},
            transition={'duration': 0},
            plot_bgcolor='#1e293b', paper_bgcolor='#1e293b',
            height=350, margin=dict(l=50, r=30, t=50, b=50),
            hovermode='x unified',
            xaxis={'tickfont': {'color': '#cbd5e1', 'size': 11}, 'gridcolor': 'rgba(148, 163, 184, 0.1)'},
            yaxis={'tickfont': {'color': '#94a3b8'}, 'gridcolor': 'rgba(148, 163, 184, 0.1)'}
        )

        # Ingresos vs Egresos
        datos_ing = processor.safe_get_row('NUMERO DE INGRESOS DEL MES', meses_filtrados)
        datos_eg = processor.safe_get_row('N\u00daMERO DE EGRESOS DEL MES', meses_filtrados)
        fig_ie = go.Figure()
        fig_ie.add_trace(go.Bar(
            name='Ingresos', x=meses_filtrados, y=datos_ing.values, marker_color='#6366f1',
            opacity=0.88,
            hovertemplate='<b>%{x}</b><br>Ingresos: %{y:.0f}<extra></extra>'
        ))
        fig_ie.add_trace(go.Scatter(
            name='Egresos', x=meses_filtrados, y=datos_eg.values, mode='lines+markers+text',
            line=dict(color='#22d3ee', width=3), marker=dict(size=8, color='#22d3ee'),
            text=[f'{v:.0f}' if pd.notna(v) else '' for v in datos_eg.values], textposition='top center',
            textfont=dict(color='#f1f5f9', size=10),
            hovertemplate='<b>%{x}</b><br>Egresos: %{y:.0f}<extra></extra>'
        ))
        fig_ie.update_layout(
            title={'text': 'Ingresos (Barras) vs Egresos (Línea)', 'font': {'size': 14, 'color': '#f1f5f9'}, 'x': 0.5},
            transition={'duration': 0},
            plot_bgcolor='#1e293b', paper_bgcolor='#1e293b',
            height=380, margin=dict(l=50, r=30, t=50, b=55),
            hovermode='x unified',
            legend=dict(orientation="h", y=-0.12, x=0.5, xanchor="center", bgcolor='rgba(30, 41, 59, 0.75)', font={'color': '#f1f5f9', 'size': 11}, bordercolor='rgba(148, 163, 184, 0.25)', borderwidth=1),
            xaxis={'tickfont': {'color': '#cbd5e1', 'size': 11}, 'automargin': True, 'gridcolor': 'rgba(148, 163, 184, 0.1)'},
            yaxis={'tickfont': {'color': '#94a3b8'}, 'gridcolor': 'rgba(148, 163, 184, 0.1)'}
        )

        # Morbilidad (total oficial = pies mensuales en Excel para meses_filtrados; escalar detalle si no cuadra)
        if not data_actual['morbilidad'].empty:
            df_morb = data_actual['morbilidad'].sort_values('Total', ascending=False)
            total_morb = _morbilidad_total_para_meses(df_morb, meses_filtrados)
            detail_morb = float(pd.to_numeric(df_morb['Total'], errors='coerce').fillna(0).sum())
            if abs(detail_morb - total_morb) > 1e-6:
                df_morb = _morbilidad_escalada_al_total_oficial(df_morb, total_morb)
                df_morb = df_morb.sort_values('Total', ascending=False)
            df_morb_top = df_morb.head(8).copy()
            # No usar el índice 'OTROS': choca con el grupo diagnóstico homónimo del Excel y
            # sobrescribe su valor; la suma del treemap quedaba por debajo del total oficial.
            if len(df_morb) > 8:
                tail = df_morb.iloc[8:]
                otros_total = float(pd.to_numeric(tail['Total'], errors='coerce').fillna(0).sum())
                if otros_total > 0:
                    h_rest = int(pd.to_numeric(tail['Hombre'], errors='coerce').fillna(0).sum())
                    m_rest = int(pd.to_numeric(tail['Mujer'], errors='coerce').fillna(0).sum())
                    df_morb_top = pd.concat(
                        [
                            df_morb_top,
                            pd.DataFrame(
                                [{'Hombre': h_rest, 'Mujer': m_rest, 'Total': otros_total}],
                                index=['Resto (demás grupos)'],
                            ),
                        ]
                    )
            colors_morb = ['#6366f1', '#3b82f6', '#10b981', '#8b5cf6', '#06b6d4', '#f43f5e', '#14b8a6', '#f59e0b', '#84cc16', '#ec4899', '#0ea5e9', '#fb923c']
            fig_morb = go.Figure(data=[go.Treemap(
                labels=[str(idx) for idx in df_morb_top.index],
                parents=[''] * len(df_morb_top),
                values=[float(v) for v in df_morb_top['Total'].values],
                marker=dict(
                    colors=[colors_morb[i % len(colors_morb)] for i in range(len(df_morb_top))],
                    line=dict(color='#0f172a', width=2)
                ),
                texttemplate='<b>%{label}</b><br>%{value:.0f}',
                textfont=dict(size=11, color='white', family='Roboto'),
                hovertemplate='<b>%{label}</b><br>Total: %{value:.0f}<br>Participación: %{percentRoot:.1%}<extra></extra>'
            )])
            total_treemap = int(pd.to_numeric(df_morb_top['Total'], errors='coerce').fillna(0).sum())
            fig_morb.update_layout(
                title={'text': f'Morbilidad por Grupo Diagnóstico (Top, Total: {total_treemap})', 'font': {'size': 15, 'color': '#f1f5f9', 'family': 'Roboto'}, 'x': 0.5},
                transition={'duration': 0},
                plot_bgcolor='#1e293b', paper_bgcolor='#1e293b',
                height=380, margin=dict(l=20, r=20, t=50, b=20),
                showlegend=False,
                uniformtext=dict(minsize=10, mode='hide')
            )

            # Morbilidad por Sexo - Barras normales
            total_hombres = df_morb['Hombre'].sum()
            total_mujeres = df_morb['Mujer'].sum()
            total_sexo = total_hombres + total_mujeres
            total_sexo_safe = max(float(total_sexo), 1.0)
            pct_h = (float(total_hombres) / total_sexo_safe) * 100
            pct_m = (float(total_mujeres) / total_sexo_safe) * 100
            fig_morb_gen = go.Figure()
            fig_morb_gen.add_trace(go.Bar(
                x=['Hombres'],
                y=[total_hombres],
                name='Hombres',
                width=0.25,
                marker_color='#3b82f6',
                text=[f'{int(total_hombres)} ({pct_h:.1f}%)'],
                textposition='inside',
                textfont=dict(color='#ffffff', size=12),
                customdata=[[pct_h]],
                hovertemplate='Hombres: %{y:.0f} (%{customdata[0]:.1f}%)<extra></extra>'
            ))
            fig_morb_gen.add_trace(go.Bar(
                x=['Mujeres'],
                y=[total_mujeres],
                name='Mujeres',
                width=0.25,
                marker_color='#ec4899',
                text=[f'{int(total_mujeres)} ({pct_m:.1f}%)'],
                textposition='inside',
                textfont=dict(color='#ffffff', size=12),
                customdata=[[pct_m]],
                hovertemplate='Mujeres: %{y:.0f} (%{customdata[0]:.1f}%)<extra></extra>'
            ))
            fig_morb_gen.update_layout(
                title={'text': f'Morbilidad por Sexo (Total: {int(total_sexo)})', 'font': {'size': 15, 'color': '#f1f5f9', 'family': 'Roboto'}, 'x': 0.5},
                transition={'duration': 0},
                barmode='group',
                bargap=0.7,
                plot_bgcolor='#1e293b', paper_bgcolor='#1e293b',
                height=380, margin=dict(l=20, r=20, t=50, b=20),
                xaxis={'tickfont': {'color': '#cbd5e1'}, 'gridcolor': 'rgba(148, 163, 184, 0.1)'},
                yaxis={'tickfont': {'color': '#94a3b8'}, 'gridcolor': 'rgba(148, 163, 184, 0.1)'},
                legend=dict(
                    orientation="h", y=-0.2, x=0.5, xanchor="center",
                    bgcolor='rgba(30, 41, 59, 0.8)', font={'color': '#f1f5f9', 'size': 11},
                    bordercolor='rgba(148, 163, 184, 0.3)', borderwidth=1
                )
            )
        else:
            fig_morb = go.Figure()
            fig_morb_gen = go.Figure()
            fig_morb.update_layout(
                title={'text': 'Ranking de Morbilidad', 'font': {'size': 15, 'color': '#f1f5f9', 'family': 'Roboto'}, 'x': 0.5},
                plot_bgcolor='#1e293b',
                paper_bgcolor='#1e293b',
                height=380,
                margin=dict(l=50, r=30, t=50, b=40),
                xaxis={'visible': False},
                yaxis={'visible': False},
                showlegend=False
            )
            fig_morb.add_annotation(
                text="<b>Sin datos de morbilidad</b>",
                x=0.5, y=0.5, xref='paper', yref='paper',
                showarrow=False,
                font=dict(size=14, color='#e2e8f0', family='Roboto')
            )
            fig_morb_gen.update_layout(
                title={'text': 'Morbilidad por Sexo', 'font': {'size': 15, 'color': '#f1f5f9', 'family': 'Roboto'}, 'x': 0.5},
                plot_bgcolor='#1e293b',
                paper_bgcolor='#1e293b',
                height=380,
                margin=dict(l=20, r=20, t=50, b=20),
                xaxis={'visible': False},
                yaxis={'visible': False},
                showlegend=False
            )
            fig_morb_gen.add_annotation(
                text="<b>Sin datos de sexo por morbilidad</b>",
                x=0.5, y=0.5, xref='paper', yref='paper',
                showarrow=False,
                font=dict(size=14, color='#e2e8f0', family='Roboto')
            )

        return meses_output_value, html.Div([
            html.Div(
                compare_block,
                style={
                    'background': 'rgba(15, 23, 42, 0.9)',
                    'border': '1px solid rgba(148, 163, 184, 0.35)',
                    'borderRadius': '10px',
                    'padding': '10px 6px 12px',
                    'marginBottom': '16px',
                    'boxShadow': '0 10px 25px rgba(15,23,42,0.7)',
                    'overflow': 'hidden'
                }
            ) if compare_block else html.Div(),
            dbc.Row([dbc.Col(k, xs=6, sm=6, md=True) for k in kpi_cards_r], className="mb-4 g-3"),
            dbc.Row([dbc.Col(dcc.Graph(figure=fig_ocup, config=_GRAPH_CFG, style={'height': '350px'}), xs=12, className="plotly-chart plotly-chart--sm")], className="mb-4"),
            dbc.Row([dbc.Col(dcc.Graph(figure=fig_ie, config=_GRAPH_CFG, style={'height': '380px'}), xs=12, className="plotly-chart plotly-chart--md")], className="mb-4"),
            dbc.Row([dbc.Col(dcc.Graph(figure=fig_eps, config=_GRAPH_CFG, style={'height': '520px'}), xs=12, className="plotly-chart plotly-chart--xl")], className="mb-4 g-3"),
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_morb, config=_GRAPH_CFG, style={'height': '380px'}), xs=12, className="plotly-chart plotly-chart--md")
            ], className="mb-4 g-3"),
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_morb_gen, config=_GRAPH_CFG, style={'height': '380px'}), xs=12, className="plotly-chart plotly-chart--md")
            ], className="mb-2 g-3")
        ]), meses_filtrados, year_store_output
    except Exception as e:
        logger.error(f"Error en actualizar_resumen: {e}")
        return dash.no_update, html.Div([html.H3(f"Error: {str(e)}", style={'color': '#dc2626'})]), dash.no_update, dash.no_update


@app.callback(
    Output('pending-year-store', 'data'),
    [Input('filtro-anio', 'value'), Input('tabs', 'active_tab')],
    [State('selected-year-store', 'data')],
    prevent_initial_call=True
)
def sync_pending_year(anio_seleccionado, active_tab, selected_year):
    """Dropdown: guarda selección (filtro-anio) sin aplicar. Al entrar a Resumen, muestra año aplicado."""
    ctx = dash.callback_context
    triggered_id = ctx.triggered[0]['prop_id'].split('.')[0] if ctx.triggered else None
    if triggered_id == 'tabs' and active_tab == 'tab-resumen':
        year_key = str(selected_year) if selected_year is not None else DEFAULT_YEAR
        return year_key if year_key in PROCESSOR_BY_YEAR else DEFAULT_YEAR
    if triggered_id == 'filtro-anio':
        year_key = str(anio_seleccionado) if anio_seleccionado is not None else DEFAULT_YEAR
        return year_key if year_key in PROCESSOR_BY_YEAR else DEFAULT_YEAR
    return dash.no_update


# Callbacks de exportacion
@app.callback(
    Output("download-excel", "data"),
    Input("btn-export-excel", "n_clicks"),
    State('selected-year-store', 'data'),
    State('selected-months-store', 'data'),
    prevent_initial_call=True
)
def export_excel(_n_clicks, selected_year, selected_months):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    year_candidate = str(selected_year) if selected_year is not None else DEFAULT_YEAR
    year_key = year_candidate if year_candidate in DATA_BY_YEAR else DEFAULT_YEAR
    _ensure_year_data_fresh(year_key)
    meses_export = selected_months if isinstance(selected_months, list) and selected_months else MESES
    datos = create_export_data(DATA_BY_YEAR[year_key], meses_export)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in datos.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
            ws = writer.sheets[sheet_name[:31]]

            # Estilo de encabezados
            header_fill = PatternFill(fill_type="solid", fgColor="1E3A5F")
            header_font = Font(color="FFFFFF", bold=True)
            header_alignment = Alignment(horizontal="center", vertical="center")
            thin_border = Border(
                left=Side(style="thin", color="D1D5DB"),
                right=Side(style="thin", color="D1D5DB"),
                top=Side(style="thin", color="D1D5DB"),
                bottom=Side(style="thin", color="D1D5DB"),
            )

            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = header_alignment
                cell.border = thin_border

            # Formato de cuerpo + bordes
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                for cell in row:
                    cell.border = thin_border
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = "#,##0.00"

            # Congelar cabecera y activar autofiltro
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            # Ajuste automático de anchura por contenido
            for col in ws.columns:
                col_letter = col[0].column_letter
                max_length = 0
                for cell in col:
                    value_len = len(str(cell.value)) if cell.value is not None else 0
                    if value_len > max_length:
                        max_length = value_len
                ws.column_dimensions[col_letter].width = min(max(max_length + 2, 12), 45)
    return dcc.send_bytes(output.getvalue(), "dashboard_uci_export.xlsx")  # type: ignore[attr-defined]


@app.callback(
    Output("download-csv", "data"),
    Input("btn-export-csv", "n_clicks"),
    State('selected-year-store', 'data'),
    State('selected-months-store', 'data'),
    prevent_initial_call=True
)
def export_csv(_n_clicks, selected_year, selected_months):
    year_candidate = str(selected_year) if selected_year is not None else DEFAULT_YEAR
    year_key = year_candidate if year_candidate in DATA_BY_YEAR else DEFAULT_YEAR
    _ensure_year_data_fresh(year_key)
    meses_export = selected_months if isinstance(selected_months, list) and selected_months else MESES
    datos = create_export_data(DATA_BY_YEAR[year_key], meses_export)
    if 'Resumen General' in datos:
        return dcc.send_data_frame(datos['Resumen General'].to_csv, "dashboard_uci_resumen.csv")  # type: ignore[attr-defined]
    return dash.no_update


@app.callback(
    Output("download-compare-excel", "data"),
    Input("btn-export-compare-excel", "n_clicks"),
    State('selected-months-store', 'data'),
    prevent_initial_call=True
)
def export_compare_excel(n_clicks, selected_months):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    if not n_clicks:
        raise dash.exceptions.PreventUpdate
    ctx = dash.callback_context
    if not ctx.triggered or ctx.triggered[0]["prop_id"].split(".")[0] != "btn-export-compare-excel":
        raise dash.exceptions.PreventUpdate
    meses_export = selected_months if isinstance(selected_months, list) and selected_months else MESES
    meses_filtrados = [m for m in meses_export if m in MESES]
    if not meses_filtrados:
        meses_filtrados = MESES
    rows = _build_compare_rows(meses_filtrados)
    if not rows:
        return dash.no_update
    df = pd.DataFrame(rows)

    column_map = {
        "year": "Año",
        "ingresos": "Total ingresos",
        "ocupacion": "% Ocupación",
        "mortalidad": "% Mortalidad",
        "estancia": "Estancia prom. (días)",
        "giro_cama": "Giro cama prom.",
        "total_dias": "Total días estancia",
        "fallecidos": "Total fallecidos",
        "mort_ucI": "% Mortalidad UCI",
        "total_eventos": "Eventos adversos",
        "total_germenes": "Gérmenes aislados",
        "total_diag": "Casos morbilidad",
    }
    ordered_columns = [c for c in column_map if c in df.columns]
    df_export = df[ordered_columns].rename(columns=column_map)

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export.to_excel(writer, sheet_name="Comparación 2025-2026", index=False)
        ws = writer.sheets["Comparación 2025-2026"]

        header_fill = PatternFill(fill_type="solid", fgColor="1E3A5F")
        header_font = Font(color="FFFFFF", bold=True)
        header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell_alignment = Alignment(horizontal="center", vertical="center")
        thin_border = Border(
            left=Side(style="thin", color="D1D5DB"),
            right=Side(style="thin", color="D1D5DB"),
            top=Side(style="thin", color="D1D5DB"),
            bottom=Side(style="thin", color="D1D5DB"),
        )

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
            cell.border = thin_border

        # Valores de % vienen en escala 0–100 (no 0–1): usar sufijo "%" sin escalar de Excel.
        percent_columns = {"% Ocupación", "% Mortalidad", "% Mortalidad UCI"}
        two_decimal_columns = {"Estancia prom. (días)", "Giro cama prom."}
        int_columns = {
            "Total ingresos", "Total días estancia", "Total fallecidos",
            "Eventos adversos", "Gérmenes aislados", "Casos morbilidad",
        }

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.border = thin_border
                cell.alignment = cell_alignment
                header_name = ws.cell(row=1, column=cell.column).value
                if header_name == "Año" and isinstance(cell.value, (int, float)):
                    cell.number_format = "0"
                elif header_name in percent_columns and isinstance(cell.value, (int, float)):
                    cell.number_format = '0.0"%"'
                elif header_name in two_decimal_columns and isinstance(cell.value, (int, float)):
                    cell.number_format = "0.00"
                elif header_name in int_columns and isinstance(cell.value, (int, float)):
                    cell.number_format = "#,##0"

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for col in ws.columns:
            col_letter = col[0].column_letter
            max_length = 0
            for cell in col:
                value_len = len(str(cell.value)) if cell.value is not None else 0
                if value_len > max_length:
                    max_length = value_len
            ws.column_dimensions[col_letter].width = min(max(max_length + 2, 12), 30)

    return dcc.send_bytes(output.getvalue(), "comparacion_2025_2026.xlsx")  # type: ignore[attr-defined]


@app.callback(
    Output("download-compare-csv", "data"),
    Input("btn-export-compare-csv", "n_clicks"),
    State('selected-months-store', 'data'),
    prevent_initial_call=True
)
def export_compare_csv(n_clicks, selected_months):
    # Evita descargas espurias al montar botones dinámicos (el botón vive fijo en el header).
    if not n_clicks:
        raise dash.exceptions.PreventUpdate
    ctx = dash.callback_context
    if not ctx.triggered or ctx.triggered[0]["prop_id"].split(".")[0] != "btn-export-compare-csv":
        raise dash.exceptions.PreventUpdate
    meses_export = selected_months if isinstance(selected_months, list) and selected_months else MESES
    meses_filtrados = [m for m in meses_export if m in MESES]
    if not meses_filtrados:
        meses_filtrados = MESES
    rows = _build_compare_rows(meses_filtrados)
    if not rows:
        return dash.no_update
    df = pd.DataFrame(rows)
    return dcc.send_data_frame(df.to_csv, "comparacion_2025_2026.csv", index=False)  # type: ignore[attr-defined]


# ==============================================================================
# PUNTO DE ENTRADA
# ==============================================================================

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8050)
