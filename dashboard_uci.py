# -*- coding: utf-8 -*-
"""
Dashboard UCI Adultos - Panel de Indicadores
Hospital Infantil Universitario de San Jose
"""

import dash
from dash import dcc, html, Input, Output, State
import dash_bootstrap_components as dbc
import plotly.graph_objects as go
import pandas as pd
import json
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

# Configuracion de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


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
        'fallecidos': pd.DataFrame(),
        'eventos_adversos': pd.DataFrame(),
        'gram_sot': pd.DataFrame()
    }


DATA_BY_YEAR: dict[str, dict[str, pd.DataFrame]] = {}
PROCESSOR_BY_YEAR: dict[str, DataProcessor] = {}

for year, filename in YEAR_FILES.items():
    excel_path = BASE_DIR / "data" / filename
    if not excel_path.exists():
        logger.warning("Archivo no encontrado para %s: %s", year, excel_path)
        continue

    try:
        year_data = DataLoader(Path(excel_path)).load_all()
        DATA_BY_YEAR[year] = year_data
        PROCESSOR_BY_YEAR[year] = DataProcessor(year_data)
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

AVAILABLE_YEARS = sorted(PROCESSOR_BY_YEAR.keys())
DEFAULT_YEAR = "2025" if "2025" in PROCESSOR_BY_YEAR else AVAILABLE_YEARS[0]

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
                dbc.Button("Exportar CSV", id="btn-export-csv", color="primary", size="sm"),
                dcc.Download(id="download-excel"),
                dcc.Download(id="download-csv"),
            ], style={'display': 'flex', 'gap': '8px', 'position': 'absolute', 'right': '20px'})
        ], style={'display': 'flex', 'justifyContent': 'center', 'alignItems': 'center', 'position': 'relative', 'padding': '20px 0'})
    ], className="header-main", style={
        'background': 'linear-gradient(135deg, #0f172a 0%, #1e3a5f 100%)',
        'borderBottom': '1px solid rgba(148, 163, 184, 0.2)',
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
    'background': 'linear-gradient(135deg, #0f172a 0%, #1e293b 100%)',
    'minHeight': '100vh',
    'fontFamily': "'Roboto', -apple-system, BlinkMacSystemFont, sans-serif"
}, id='app-container')


# ==============================================================================
# CALLBACKS
# ==============================================================================

@app.callback(
    Output('tab-content', 'children'),
    [Input('tabs', 'active_tab'), Input('selected-year-store', 'data')]
)
def render_content(tab, selected_year):
    year_candidate = str(selected_year) if selected_year is not None else DEFAULT_YEAR
    year_key = year_candidate if year_candidate in PROCESSOR_BY_YEAR else DEFAULT_YEAR
    if tab == 'tab-resumen':
        return render_resumen(year_key)
    elif tab == 'tab-gestion':
        return render_gestion(year_key)
    elif tab == 'tab-especialidades':
        return render_especialidades(year_key)
    elif tab == 'tab-mortalidad':
        return render_mortalidad(year_key)
    return html.Div([html.H5("Selecciona una pesta\u00f1a", style={'color': '#1e293b'})])


def render_resumen(year_key: str):
    filtros = html.Div([
        html.Div([
            html.Span("Per\u00edodo de an\u00e1lisis", style={
                'color': '#f1f5f9', 'fontWeight': '600', 'fontSize': '14px'
            })
        ], style={'marginBottom': '12px'}),
        html.Div([
            html.Div([
                html.Div("Año", style={
                    'color': '#94a3b8', 'fontSize': '12px', 'fontWeight': '600', 'lineHeight': '1'
                }),
                html.Div("Meses", style={
                    'color': '#94a3b8', 'fontSize': '12px', 'fontWeight': '600', 'lineHeight': '1'
                }),
                html.Div("", style={'height': '12px'})
            ], style={
                'display': 'grid',
                'gridTemplateColumns': '120px minmax(420px, 1fr) auto',
                'alignItems': 'end',
                'columnGap': '12px',
                'marginBottom': '6px'
            }),
            html.Div([
                dcc.Dropdown(
                    id='filtro-anio',
                    options=[{'label': anio, 'value': anio} for anio in AVAILABLE_YEARS],
                    value=year_key,
                    clearable=False,
                    persistence=True,
                    persistence_type='session',
                    style={'fontSize': '13px'}
                ),
                dcc.Dropdown(
                    id='filtro-mes-inicio',
                    options=[{'label': MESES_COMPLETOS[mes], 'value': mes} for mes in MESES],
                    value=MESES,
                    multi=True,
                    placeholder="Seleccionar meses...",
                    style={'fontSize': '13px'}
                ),
                html.Div([
                    dbc.Button("Aplicar", id='btn-aplicar-filtros', color="success", size="sm"),
                    dbc.Button("Todos", id='btn-todos-meses', color="secondary", size="sm",
                              style={'marginLeft': '8px'})
                ], style={'display': 'flex', 'alignItems': 'center'})
            ], style={
                'display': 'grid',
                'gridTemplateColumns': '120px minmax(420px, 1fr) auto',
                'alignItems': 'center',
                'columnGap': '12px'
            })
        ], style={'display': 'flex', 'flexDirection': 'column', 'gap': '0'})
    ], style={
        'background': 'rgba(30, 41, 59, 0.8)',
        'padding': '14px 20px',
        'borderRadius': '8px',
        'marginBottom': '16px',
        'border': '1px solid rgba(148, 163, 184, 0.15)'
    }, className="filtros-section")
    return html.Div([filtros, html.Div(id='resumen-content')])


def render_gestion(year_key: str):
    try:
        processor = PROCESSOR_BY_YEAR.get(year_key, PROCESSOR)
        data_year = DATA_BY_YEAR.get(year_key, DATA)
        # region agent log
        _debug_log(
            "pre-fix",
            "H5",
            "dashboard_uci.py:render_gestion:entry",
            "Gestion tab resolved year",
            {"globalDefaultYear": DEFAULT_YEAR, "resolvedYear": year_key},
        )
        # endregion
        kpis = processor.get_kpis_gestion()
        kpi1 = chart_factory.create_kpi_card("Estancia Promedio", f"{kpis['promedio_estancia']:.2f}", "D\u00edas", "#6366f1")
        kpi2 = chart_factory.create_kpi_card("Giro de Cama", f"{kpis['giro_cama']:.2f}", "Promedio", "#3b82f6")
        kpi3 = chart_factory.create_kpi_card("Total D\u00edas Estancia", f"{int(kpis['total_dias_estancia'])}", "Ene - Dic", "#10b981")
        kpi4 = chart_factory.create_kpi_card("Capacidad UCI", f"{int(kpis['capacidad_uci'])}", "Camas", "#8b5cf6")
        kpi5 = chart_factory.create_kpi_card("Promedio D\u00edas VM", f"{kpis['promedio_dias_vm']:.2f}", "Ventilaci\u00f3n Mec\u00e1nica", "#06b6d4")

        datos_estancia = processor.safe_get_row('PROMEDIO ESTANCIA', MESES)
        fig_estancia = chart_factory.create_line_chart(datos_estancia, 'Estancia Promedio (D\u00edas)', color='#6366f1', decimals=1)
        datos_giro = processor.safe_get_row('GIRO CAMA', MESES)
        fig_giro = chart_factory.create_bar_chart(datos_giro, 'Giro de Cama', color='#3b82f6', decimals=2)

        # Grafico de VM - datos mensuales
        datos_vm_mensual = kpis.get('promedio_dias_vm_mensual', [])
        if not datos_vm_mensual or len(datos_vm_mensual) != len(MESES):
            datos_vm_mensual = [round(kpis['promedio_dias_vm'], 2) for _ in MESES]
        fig_vm = go.Figure()
        fig_vm.add_trace(go.Scatter(
            x=MESES, y=datos_vm_mensual,
            mode='lines+markers',
            line=dict(color='#06b6d4', width=3),
            marker=dict(size=7, color='#06b6d4'),
            fill='tozeroy',
            fillcolor='rgba(6, 182, 212, 0.18)',
            hovertemplate='<b>%{x}</b><br>Promedio VM: %{y:.1f}<extra></extra>'
        ))
        fig_vm.update_layout(
            title={'text': 'Promedio de D\u00edas de VM (Tendencia)', 'font': {'size': 14, 'color': '#f1f5f9'}, 'x': 0.5},
            transition={'duration': 0},
            plot_bgcolor='#1e293b', paper_bgcolor='#1e293b',
            height=350, margin=dict(l=50, r=30, t=50, b=50),
            hovermode='x unified',
            xaxis={'tickfont': {'color': '#cbd5e1', 'size': 11}, 'gridcolor': 'rgba(148, 163, 184, 0.1)'},
            yaxis={'tickfont': {'color': '#94a3b8'}, 'gridcolor': 'rgba(148, 163, 184, 0.1)'}
        )

        fig_bacteriemia = chart_factory.create_line_chart(
            processor.safe_get_row('TASA DE BACTERIEMIA ASOC.  A C. CENTRAL', MESES),
            'Tasa de ITS AC', color='#ef4444', decimals=1
        )
        fig_nav = chart_factory.create_line_chart(
            processor.safe_get_row('TASA NEUMON\u00cdA ASOCIADA A VENTILADOR', MESES),
            'Tasa NAV', color='#ef4444', decimals=1
        )

        # Germenes - Donut
        if not data_year['germenes'].empty:
            germenes_labels = data_year['germenes']['Germen'].tolist()
            germenes_values = data_year['germenes']['Cantidad'].tolist()
            datos_germenes = sorted(zip(germenes_values, germenes_labels), reverse=True)
            germenes_values_sorted = [d[0] for d in datos_germenes]
            germenes_labels_sorted = [d[1] for d in datos_germenes]
            total_germenes = sum(germenes_values)
            colors_germenes = ['#6366f1', '#3b82f6', '#10b981', '#8b5cf6', '#06b6d4', '#f43f5e', '#14b8a6', '#f59e0b', '#84cc16', '#ec4899', '#0ea5e9', '#fb923c']
            fig_germenes = go.Figure(data=[go.Pie(
                labels=germenes_labels_sorted,
                values=germenes_values_sorted,
                hole=0.55,
                marker=dict(colors=colors_germenes[:len(germenes_labels_sorted)], line=dict(color='#0f172a', width=2)),
                textinfo='percent',
                textposition='inside',
                textfont=dict(size=12, color='white', family='Roboto'),
                hovertemplate='<b>%{label}</b><br>Cantidad: %{value}<br>Porcentaje: %{percent}<extra></extra>',
                direction='clockwise',
                sort=False,
                pull=[0.03 if i == 0 else 0 for i in range(len(germenes_labels_sorted))]
            )])
            fig_germenes.add_annotation(
                text=f"<b style='font-size:22px'>{int(total_germenes)}</b><br><span style='color:#94a3b8'>Aislados</span>",
                x=0.5, y=0.5,
                font=dict(size=14, color='#f1f5f9', family='Roboto'),
                showarrow=False
            )
            fig_germenes.update_layout(
                title={'text': 'G\u00e9rmenes Aislados (IAAS)', 'font': {'size': 16, 'color': '#f1f5f9', 'family': 'Roboto'}, 'x': 0.5},
                transition={'duration': 0},
                plot_bgcolor='#1e293b', paper_bgcolor='#1e293b',
                height=520, margin=dict(l=30, r=300, t=60, b=30),
                legend=dict(
                    orientation='v', yanchor='middle', y=0.5, xanchor='left', x=1.08,
                    bgcolor='rgba(15, 23, 42, 0.9)',
                    font=dict(color='#e2e8f0', size=9, family='Roboto'),
                    bordercolor='rgba(148, 163, 184, 0.3)', borderwidth=1,
                    itemsizing='constant', tracegroupgap=3
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
        if not data_year['eventos_adversos'].empty:
            df_eventos = data_year['eventos_adversos'].sort_values('Cantidad', ascending=False)
            if len(df_eventos) > 12:
                df_eventos_top = df_eventos.head(12).copy()
                otros_eventos = int(df_eventos.iloc[12:]['Cantidad'].sum())
                if otros_eventos > 0:
                    df_eventos_top = pd.concat(
                        [df_eventos_top, pd.DataFrame([{'Evento': 'Otros', 'Cantidad': otros_eventos}])],
                        ignore_index=True
                    )
                df_eventos = df_eventos_top
            total_eventos = max(float(df_eventos['Cantidad'].sum()), 1.0)
            df_eventos['Acumulado %'] = (df_eventos['Cantidad'].cumsum() / total_eventos) * 100
            colors_eventos = [
                '#ef4444', '#6366f1', '#10b981', '#f59e0b', '#8b5cf6', '#06b6d4', '#ec4899',
                '#f97316', '#14b8a6', '#a855f7', '#0ea5e9', '#84cc16', '#e11d48', '#22c55e',
                '#3b82f6', '#eab308', '#d946ef', '#0891b2', '#f43f5e', '#4ade80', '#818cf8',
                '#fb923c', '#2dd4bf', '#c084fc', '#38bdf8', '#a3e635', '#fb7185', '#34d399'
            ]
            fig_eventos = go.Figure()
            fig_eventos.add_trace(go.Bar(
                x=df_eventos['Evento'],
                y=df_eventos['Cantidad'],
                name='Cantidad',
                marker_color=[colors_eventos[i % len(colors_eventos)] for i in range(len(df_eventos))],
                text=df_eventos['Cantidad'],
                textposition='auto',
                textfont=dict(color='#f1f5f9', size=11),
                hovertemplate='<b>%{x}</b><br>Eventos: %{y}<extra></extra>'
            ))
            fig_eventos.add_trace(go.Scatter(
                x=df_eventos['Evento'],
                y=df_eventos['Acumulado %'],
                name='Acumulado %',
                mode='lines+markers',
                yaxis='y2',
                line=dict(color='#fbbf24', width=2.5),
                marker=dict(size=6, color='#fbbf24'),
                hovertemplate='<b>%{x}</b><br>Acumulado: %{y:.1f}%<extra></extra>'
            ))
            fig_eventos.update_layout(
                title={'text': 'Eventos Adversos (Pareto)', 'font': {'size': 14, 'color': '#f1f5f9'}, 'x': 0.5},
                transition={'duration': 0},
                plot_bgcolor='#1e293b', paper_bgcolor='#1e293b',
                height=550, margin=dict(l=50, r=30, t=50, b=220),
                hovermode='x unified',
                xaxis={'tickfont': {'color': '#e2e8f0', 'size': 9}, 'tickangle': -35, 'automargin': True, 'gridcolor': 'rgba(148, 163, 184, 0.1)'},
                yaxis={'title': 'Cantidad', 'tickfont': {'color': '#94a3b8'}, 'gridcolor': 'rgba(148, 163, 184, 0.1)'},
                yaxis2={
                    'overlaying': 'y',
                    'side': 'right',
                    'range': [0, 110],
                    'tickfont': {'color': '#fbbf24'},
                    'title': {'text': 'Acumulado %', 'font': {'color': '#fbbf24'}},
                    'showgrid': False
                },
                legend=dict(
                    orientation="h", y=-0.18, x=0.5, xanchor="center",
                    bgcolor='rgba(30, 41, 59, 0.8)', font={'color': '#f1f5f9', 'size': 11},
                    bordercolor='rgba(148, 163, 184, 0.3)', borderwidth=1
                )
            )
        else:
            fig_eventos = go.Figure()
            fig_eventos.update_layout(
                title={'text': 'Eventos Adversos por Categoría', 'font': {'size': 14, 'color': '#f1f5f9'}, 'x': 0.5},
                plot_bgcolor='#1e293b',
                paper_bgcolor='#1e293b',
                height=550,
                margin=dict(l=50, r=30, t=50, b=80),
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

        return html.Div([
            html.H5("Indicadores de Gesti\u00f3n", style={'color': '#f1f5f9', 'fontWeight': '600'}, className="mb-3"),
            dbc.Row([
                dbc.Col(kpi1, xs=12, md=True), dbc.Col(kpi2, xs=12, md=True),
                dbc.Col(kpi3, xs=12, md=True), dbc.Col(kpi4, xs=12, md=True),
                dbc.Col(kpi5, xs=12, md=True)
            ], className="mb-4 g-3"),
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_estancia, config=_GRAPH_CFG), xs=12)
            ], className="mb-4 g-3"),
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_giro, config=_GRAPH_CFG), xs=12)
            ], className="mb-3 g-3"),
            dbc.Row([
                dbc.Col(xs=12, md=3),
                dbc.Col(dcc.Graph(figure=fig_vm, config=_GRAPH_CFG), xs=12, md=6),
                dbc.Col(xs=12, md=3)
            ], className="mb-4", justify="center"),
            html.H5("Indicadores de Calidad e Infecciones", style={'color': '#f1f5f9', 'fontWeight': '600', 'marginTop': '16px', 'marginBottom': '16px', 'textAlign': 'center'}),
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_bacteriemia, config=_GRAPH_CFG), xs=12)
            ], className="mb-4 g-3"),
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_nav, config=_GRAPH_CFG), xs=12)
            ], className="mb-3 g-3"),
            html.H5("G\u00e9rmenes Aislados y Eventos Adversos", style={'color': '#f1f5f9', 'fontWeight': '600', 'marginTop': '16px', 'marginBottom': '16px', 'textAlign': 'center'}),
            dbc.Row([
                dbc.Col(xs=12, md=2),
                dbc.Col(dcc.Graph(figure=fig_germenes, config=_GRAPH_CFG), xs=12, md=8),
                dbc.Col(xs=12, md=2)
            ], className="mb-4", justify="center"),
            dbc.Row([dbc.Col(dcc.Graph(figure=fig_eventos, config=_GRAPH_CFG), xs=12)])
        ])
    except Exception as e:
        logger.error(f"Error en render_gestion: {e}")
        return html.Div([html.H3(f"Error: {str(e)}", style={'color': '#e74c3c'})])


def render_especialidades(year_key: str):
    try:
        processor = PROCESSOR_BY_YEAR.get(year_key, PROCESSOR)
        suma_col = processor.get_suma_column()
        total_medicina = processor.safe_get_value('MEDICINA INTERNA', suma_col)
        total_cirugia = processor.safe_get_value('CIRUGIA GENERAL', suma_col)
        total_neuro = processor.safe_get_value('NEUROCIRUGIA', suma_col)
        total_neurologia = processor.safe_get_value('NEUROLOGIA', suma_col)
        total = max(total_medicina + total_cirugia + total_neuro + total_neurologia, 1)

        kpi1 = chart_factory.create_kpi_card("Medicina Interna", f"{int(total_medicina)}", f"{processor.formatear_porcentaje(total_medicina/total*100, False)}%", "#3498db")
        kpi2 = chart_factory.create_kpi_card("Cirug\u00eda General", f"{int(total_cirugia)}", f"{processor.formatear_porcentaje(total_cirugia/total*100, False)}%", "#27ae60")
        kpi3 = chart_factory.create_kpi_card("Neurocirug\u00eda", f"{int(total_neuro)}", f"{processor.formatear_porcentaje(total_neuro/total*100, False)}%", "#16a085")
        kpi4 = chart_factory.create_kpi_card("Neurolog\u00eda", f"{int(total_neurologia)}", f"{processor.formatear_porcentaje(total_neurologia/total*100, False)}%", "#9b59b6")

        series_medicina = pd.to_numeric(processor.safe_get_row('MEDICINA INTERNA', MESES), errors='coerce').fillna(0)
        series_cirugia = pd.to_numeric(processor.safe_get_row('CIRUGIA GENERAL', MESES), errors='coerce').fillna(0)
        series_neuro = pd.to_numeric(processor.safe_get_row('NEUROCIRUGIA', MESES), errors='coerce').fillna(0)
        series_neurologia = pd.to_numeric(processor.safe_get_row('NEUROLOGIA', MESES), errors='coerce').fillna(0)
        df_especialidades = pd.DataFrame({
            'Mes': MESES,
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

        fig_especialidades_area = go.Figure()
        fig_especialidades_area.add_trace(go.Scatter(
            x=df_especialidades['Mes'], y=df_especialidades['Medicina Interna'],
            name='Medicina Interna', mode='lines', stackgroup='one', line=dict(width=0.8, color='#3498db')
        ))
        fig_especialidades_area.add_trace(go.Scatter(
            x=df_especialidades['Mes'], y=df_especialidades['Cirugía General'],
            name='Cirugía General', mode='lines', stackgroup='one', line=dict(width=0.8, color='#27ae60')
        ))
        fig_especialidades_area.add_trace(go.Scatter(
            x=df_especialidades['Mes'], y=df_especialidades['Neurocirugía'],
            name='Neurocirugía', mode='lines', stackgroup='one', line=dict(width=0.8, color='#16a085')
        ))
        fig_especialidades_area.add_trace(go.Scatter(
            x=df_especialidades['Mes'], y=df_especialidades['Neurología'],
            name='Neurología', mode='lines', stackgroup='one', line=dict(width=0.8, color='#9b59b6')
        ))
        fig_especialidades_area.update_layout(
            title={'text': 'Volumen Mensual por Especialidad (Área Apilada)', 'font': {'size': 14, 'color': '#f1f5f9'}, 'x': 0.5},
            transition={'duration': 0},
            plot_bgcolor='#1e293b', paper_bgcolor='#1e293b',
            height=380, margin=dict(l=50, r=30, t=50, b=60),
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
        fig_especialidades_mix.add_trace(go.Bar(
            x=MESES, y=(df_especialidades['Medicina Interna'] / df_especialidades['Total']) * 100,
            name='Medicina Interna', marker_color='#3498db'
        ))
        fig_especialidades_mix.add_trace(go.Bar(
            x=MESES, y=(df_especialidades['Cirugía General'] / df_especialidades['Total']) * 100,
            name='Cirugía General', marker_color='#27ae60'
        ))
        fig_especialidades_mix.add_trace(go.Bar(
            x=MESES, y=(df_especialidades['Neurocirugía'] / df_especialidades['Total']) * 100,
            name='Neurocirugía', marker_color='#16a085'
        ))
        fig_especialidades_mix.add_trace(go.Bar(
            x=MESES, y=(df_especialidades['Neurología'] / df_especialidades['Total']) * 100,
            name='Neurología', marker_color='#9b59b6'
        ))
        fig_especialidades_mix.update_layout(
            title={'text': 'Composición % por Mes (100% Apilada)', 'font': {'size': 14, 'color': '#f1f5f9'}, 'x': 0.5},
            transition={'duration': 0},
            barmode='stack',
            plot_bgcolor='#1e293b', paper_bgcolor='#1e293b',
            height=380, margin=dict(l=50, r=30, t=50, b=60),
            hovermode='x unified',
            xaxis={'tickfont': {'color': '#cbd5e1', 'size': 11}, 'gridcolor': 'rgba(148, 163, 184, 0.1)'},
            yaxis={
                'title': '% del total mensual',
                'tickfont': {'color': '#94a3b8'},
                'ticksuffix': '%',
                'range': [0, 100],
                'gridcolor': 'rgba(148, 163, 184, 0.1)'
            },
            legend=dict(
                orientation="h", y=-0.18, x=0.5, xanchor="center",
                bgcolor='rgba(30, 41, 59, 0.8)', font={'color': '#f1f5f9', 'size': 11},
                bordercolor='rgba(148, 163, 184, 0.3)', borderwidth=1
            )
        )

        return html.Div([
            html.H5("Principales Especialidades", style={'color': '#f1f5f9', 'fontWeight': '600', 'marginBottom': '12px'}),
            dbc.Row([dbc.Col(kpi1, xs=12, md=3), dbc.Col(kpi2, xs=12, md=3), dbc.Col(kpi3, xs=12, md=3), dbc.Col(kpi4, xs=12, md=3)], className="mb-3 g-3"),
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_especialidades_area, config=_GRAPH_CFG), xs=12)
            ], className="mb-3 g-3"),
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_especialidades_mix, config=_GRAPH_CFG), xs=12)
            ], className="g-3")
        ])
    except Exception as e:
        logger.error(f"Error en render_especialidades: {e}")
        return html.Div([html.H3(f"Error: {str(e)}", style={'color': '#e74c3c'})])


def render_mortalidad(year_key: str):
    try:
        processor = PROCESSOR_BY_YEAR.get(year_key, PROCESSOR)
        data_year = DATA_BY_YEAR.get(year_key, DATA)
        kpis_mort = processor.get_kpis_mortalidad()
        kpi1 = chart_factory.create_kpi_card("Total Fallecidos", f"{int(kpis_mort['total_fallecidos'])}", "Ene - Dic", "#e74c3c")
        kpi2 = chart_factory.create_kpi_card("% Mortalidad General", processor.formatear_porcentaje(kpis_mort['mortalidad_general']), "Ene - Dic", "#c0392b")
        kpi3 = chart_factory.create_kpi_card("Fallecidos COVID-19", f"{int(kpis_mort['fallecidos_covid'])}", "Ene - Dic", "#34495e")
        kpi4 = chart_factory.create_kpi_card("% Mortalidad Trauma", processor.formatear_porcentaje(kpis_mort['mortalidad_trauma']), "Promedio", "#f39c12")
        kpi_prom_dias = chart_factory.create_kpi_card("Promedio D\u00edas Estancia", f"{kpis_mort['promedio_dias_estancia']:.1f}", "Fallecidos", "#9b59b6")
        kpi_mediana = chart_factory.create_kpi_card("Mediana D\u00edas de Estancia", f"{kpis_mort['mediana_dias_estancia']:.1f}", "Fallecidos", "#3498db")

        fig_mort = chart_factory.create_line_chart(
            processor.safe_get_row('% DE MORTALIDAD GENERAL', MESES),
            '% Mortalidad General', color='#e74c3c', decimals=1
        )
        fig_trauma = chart_factory.create_grouped_bar_chart(
            ['Pacientes con diagn\u00f3stico de Trauma', 'Pacientes fallecidos por Trauma'],
            ['Pacientes Trauma', 'Fallecidos Trauma'], 'Trauma: Pacientes vs Fallecidos',
            processor=processor, decimals=0
        )
        fig_piramide = chart_factory.create_pyramid_chart(data_year['fallecidos'])

        return html.Div([
            dbc.Row([dbc.Col(kpi1, xs=12, md=3), dbc.Col(kpi2, xs=12, md=3), dbc.Col(kpi3, xs=12, md=3), dbc.Col(kpi4, xs=12, md=3)], className="mb-3 g-3"),
            html.H6("D\u00edas de Estancia - Fallecidos", style={'color': '#f1f5f9', 'fontWeight': '600', 'marginBottom': '10px', 'marginTop': '8px'}),
            dbc.Row([dbc.Col(kpi_prom_dias, xs=12, md=6), dbc.Col(kpi_mediana, xs=12, md=6)], className="mb-3 g-3"),
            dbc.Row([
                dbc.Col(xs=12, md=2),
                dbc.Col(dcc.Graph(figure=fig_mort, config=_GRAPH_CFG), xs=12, md=8),
                dbc.Col(xs=12, md=2)
            ], className="mb-3", justify="center"),
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_trauma, config=_GRAPH_CFG), xs=12)
            ], className="mb-4 g-3"),
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_piramide, config=_GRAPH_CFG), xs=12)
            ], className="g-3")
        ])
    except Exception as e:
        logger.error(f"Error en render_mortalidad: {e}")
        return html.Div([html.H3(f"Error: {str(e)}", style={'color': '#e74c3c'})])


# Callback para Resumen
@app.callback(
    [Output('filtro-mes-inicio', 'value'),
     Output('resumen-content', 'children'),
     Output('selected-months-store', 'data')],
    [Input('btn-todos-meses', 'n_clicks'),
     Input('btn-aplicar-filtros', 'n_clicks'),
     Input('filtro-anio', 'value'),
     Input('filtro-mes-inicio', 'value')]
)
def actualizar_resumen(_btn_todos, _btn_aplicar, anio_seleccionado, meses_seleccionados):
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
    if triggered_id == 'btn-todos-meses':
        meses_seleccionados = MESES
    try:
        year_key = str(anio_seleccionado) if anio_seleccionado is not None else DEFAULT_YEAR
        if year_key not in PROCESSOR_BY_YEAR:
            year_key = DEFAULT_YEAR

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
        else:
            meses_filtrados = MESES

        if len(meses_filtrados) == 1:
            rango_texto = MESES_COMPLETOS.get(meses_filtrados[0], meses_filtrados[0])
        elif len(meses_filtrados) == len(MESES):
            rango_texto = "Ene - Dic"
        else:
            rango_texto = ", ".join([MESES_COMPLETOS.get(m, m)[:3] for m in meses_filtrados])

        rango_texto = f"{rango_texto} {year_key}"

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
        kpi_ingresos = chart_factory.create_kpi_card("Total Ingresos", f"{int(kpis['total_ingresos'])}", rango_texto, "#6366f1")
        kpi_ocup = chart_factory.create_kpi_card("% Ocupaci\u00f3n Promedio", processor.formatear_porcentaje(kpis['promedio_ocupacional']), rango_texto, "#3b82f6")
        kpi_mort = chart_factory.create_kpi_card("% Mortalidad General", processor.formatear_porcentaje(kpis['mortalidad']), f"Total {rango_texto}", "#ef4444")
        kpi_estancia = chart_factory.create_kpi_card("Estancia Promedio", f"{kpis['promedio_estancia']:.2f}", f"D\u00edas - {rango_texto}", "#10b981")
        kpi_giro = chart_factory.create_kpi_card("Giro de Cama", f"{kpis['giro_cama']:.2f}", f"Promedio {rango_texto}", "#8b5cf6")

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
            # Agrupar categorías pequeñas para mejorar legibilidad en donut
            if len(eps_df) > 9:
                otros_eps = int(eps_df.iloc[9:]['Cantidad'].sum())
                eps_df = eps_df.head(9).copy()
                if otros_eps > 0:
                    eps_df = pd.concat(
                        [eps_df, pd.DataFrame([{'EPS': 'OTRAS EPS', 'Cantidad': otros_eps}])],
                        ignore_index=True
                    )
            # Orden ascendente para que el donut inicie con el menor porcentaje.
            eps_df = eps_df.sort_values('Cantidad', ascending=True).reset_index(drop=True)

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
            fig_eps = go.Figure(data=[go.Pie(
                labels=eps_df['EPS_DISPLAY'],
                values=eps_df['Cantidad'],
                customdata=eps_df[['EPS']],
                hole=0.58,
                sort=False,
                marker=dict(
                    colors=resolved_eps_colors,
                    line=dict(color='#0f172a', width=2)
                ),
                textinfo='percent',
                textposition='inside',
                textfont=dict(size=11, color='white', family='Roboto'),
                hovertemplate='<b>%{customdata[0]}</b><br>Cantidad: %{value}<br>Porcentaje: %{percent}<extra></extra>'
            )])
            fig_eps.add_annotation(
                text=f"<b>Total</b><br>{total_eps:,}".replace(',', '.'),
                x=0.5, y=0.5,
                font=dict(size=14, color='#f1f5f9', family='Roboto'),
                showarrow=False
            )
            fig_eps.update_traces(
                texttemplate='%{percent}'
            )
            fig_eps.update_layout(
                title={'text': f'Distribuci\u00f3n por EPS - Donut (Total: {total_eps:,})', 'font': {'size': 14, 'color': '#f1f5f9'}, 'x': 0.5},
                transition={'duration': 0},
                plot_bgcolor='#1e293b', paper_bgcolor='#1e293b',
                font={'color': '#f1f5f9'},
                height=520, margin=dict(l=20, r=210, t=60, b=20),
                uniformtext=dict(minsize=10, mode='hide'),
                showlegend=True,
                legend=dict(
                    orientation='v',
                    yanchor='middle',
                    y=0.5,
                    xanchor='left',
                    x=1.02,
                    bgcolor='rgba(30, 41, 59, 0.82)',
                    bordercolor='rgba(148, 163, 184, 0.3)',
                    borderwidth=1,
                    font=dict(color='#f1f5f9', size=10)
                )
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

        # Morbilidad
        if not data_actual['morbilidad'].empty:
            df_morb = data_actual['morbilidad'].sort_values('Total', ascending=False)
            total_morb = float(df_morb['Total'].sum())
            df_morb_top = df_morb.head(8).copy()
            if len(df_morb) > 8:
                otros_total = float(df_morb.iloc[8:]['Total'].sum())
                if otros_total > 0:
                    df_morb_top.loc['OTROS'] = {'Hombre': 0, 'Mujer': 0, 'Total': otros_total}
            colors_morb = ['#6366f1', '#3b82f6', '#10b981', '#8b5cf6', '#06b6d4', '#f43f5e', '#14b8a6', '#f59e0b', '#84cc16', '#ec4899', '#0ea5e9', '#fb923c']
            fig_morb = go.Figure()
            for i, (diagnostico, fila) in enumerate(df_morb_top.iterrows()):
                valor = float(fila['Total'])
                fig_morb.add_trace(go.Bar(
                    y=[diagnostico],
                    x=[valor],
                    name=str(diagnostico),
                    orientation='h',
                    marker_color=colors_morb[i % len(colors_morb)],
                    text=[f"{int(valor)}"],
                    textposition='auto',
                    textfont=dict(color='#f1f5f9', size=11, family='Roboto'),
                    hovertemplate='<b>%{y}</b><br>Total: %{x:.0f}<extra></extra>',
                    showlegend=True
                ))
            fig_morb.update_layout(
                title={'text': f'Ranking de Morbilidad (Top diagnósticos, Total: {int(total_morb)})', 'font': {'size': 15, 'color': '#f1f5f9', 'family': 'Roboto'}, 'x': 0.5},
                transition={'duration': 0},
                plot_bgcolor='#1e293b', paper_bgcolor='#1e293b',
                height=380, margin=dict(l=130, r=220, t=50, b=40),
                xaxis={'tickfont': {'color': '#e2e8f0', 'size': 10}, 'gridcolor': 'rgba(148, 163, 184, 0.1)'},
                yaxis={'tickfont': {'color': '#94a3b8'}, 'gridcolor': 'rgba(148, 163, 184, 0.1)', 'categoryorder': 'total ascending', 'automargin': True},
                legend=dict(
                    title={'text': 'Diagnóstico', 'font': {'color': '#f1f5f9', 'size': 11}},
                    orientation='v',
                    yanchor='middle', y=0.5,
                    xanchor='left', x=1.02,
                    bgcolor='rgba(30, 41, 59, 0.82)',
                    font={'color': '#f1f5f9', 'size': 10},
                    bordercolor='rgba(148, 163, 184, 0.28)',
                    borderwidth=1
                ),
                bargap=0.25
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
                marker_color='#3b82f6',
                text=[f'{int(total_hombres)} ({pct_h:.1f}%)'],
                textposition='inside',
                customdata=[[pct_h]],
                hovertemplate='Hombres: %{y:.0f} (%{customdata[0]:.1f}%)<extra></extra>'
            ))
            fig_morb_gen.add_trace(go.Bar(
                x=['Mujeres'],
                y=[total_mujeres],
                name='Mujeres',
                marker_color='#ec4899',
                text=[f'{int(total_mujeres)} ({pct_m:.1f}%)'],
                textposition='inside',
                customdata=[[pct_m]],
                hovertemplate='Mujeres: %{y:.0f} (%{customdata[0]:.1f}%)<extra></extra>'
            ))
            fig_morb_gen.update_layout(
                title={'text': f'Distribuci\u00f3n por Sexo (Total: {int(total_sexo)})', 'font': {'size': 15, 'color': '#f1f5f9', 'family': 'Roboto'}, 'x': 0.5},
                transition={'duration': 0},
                barmode='group',
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
                title={'text': 'Distribución por Sexo', 'font': {'size': 15, 'color': '#f1f5f9', 'family': 'Roboto'}, 'x': 0.5},
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

        return meses_filtrados, html.Div([
            dbc.Row([
                dbc.Col(kpi_ingresos, xs=6, sm=6, md=True),
                dbc.Col(kpi_ocup, xs=6, sm=6, md=True),
                dbc.Col(kpi_mort, xs=6, sm=6, md=True),
                dbc.Col(kpi_estancia, xs=6, sm=6, md=True),
                dbc.Col(kpi_giro, xs=6, sm=6, md=True)
            ], className="mb-4 g-3"),
            dbc.Row([dbc.Col(dcc.Graph(figure=fig_ocup, config=_GRAPH_CFG), xs=12)], className="mb-4"),
            dbc.Row([dbc.Col(dcc.Graph(figure=fig_ie, config=_GRAPH_CFG), xs=12)], className="mb-4"),
            dbc.Row([dbc.Col(dcc.Graph(figure=fig_eps, config=_GRAPH_CFG), xs=12)], className="mb-4 g-3"),
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_morb, config=_GRAPH_CFG), xs=12)
            ], className="mb-4 g-3"),
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_morb_gen, config=_GRAPH_CFG), xs=12)
            ], className="mb-3 g-3")
        ]), meses_filtrados
    except Exception as e:
        logger.error(f"Error en actualizar_resumen: {e}")
        return dash.no_update, html.Div([html.H3(f"Error: {str(e)}", style={'color': '#dc2626'})]), dash.no_update


@app.callback(
    Output('selected-year-store', 'data'),
    Input('filtro-anio', 'value'),
    prevent_initial_call=True
)
def sync_selected_year(anio_seleccionado):
    year_key = str(anio_seleccionado) if anio_seleccionado is not None else DEFAULT_YEAR
    return year_key if year_key in PROCESSOR_BY_YEAR else DEFAULT_YEAR


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
    meses_export = selected_months if isinstance(selected_months, list) and selected_months else MESES
    datos = create_export_data(DATA_BY_YEAR[year_key], meses_export)
    if 'Resumen General' in datos:
        return dcc.send_data_frame(datos['Resumen General'].to_csv, "dashboard_uci_resumen.csv")  # type: ignore[attr-defined]
    return dash.no_update


# ==============================================================================
# PUNTO DE ENTRADA
# ==============================================================================

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8050)
