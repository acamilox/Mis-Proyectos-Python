# -*- coding: utf-8 -*-
"""
Dashboard UCI Adultos - Panel de Indicadores
Hospital Infantil Universitario de San José

Aplicación Dash modular con visualizaciones Plotly.
Versión refactorizada con arquitectura modular.

Author: Hospital Infantil Universitario de San José
"""

import dash
from dash import dcc, html, Input, Output, State
import dash_bootstrap_components as dbc
import plotly.graph_objects as go
import pandas as pd
from io import BytesIO
import logging

# Imports del paquete modular
from dashboard_uci import (
    load_data,
    DataProcessor,
    ChartFactory,
    chart_factory,
    MESES,
    MESES_COMPLETOS,
    GRAPH_CONFIG
)

# Configuración de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# ============================================
# INICIALIZACIÓN DE LA APLICACIÓN
# ============================================

app = dash.Dash(
    __name__, 
    external_stylesheets=[dbc.themes.BOOTSTRAP],
    suppress_callback_exceptions=True
)
server = app.server

# Cargar datos al inicio
logger.info("Cargando datos...")
try:
    DATA = load_data()
    PROCESSOR = DataProcessor(DATA)
    logger.info("Datos cargados exitosamente")
except Exception as e:
    logger.error(f"Error cargando datos: {e}")
    DATA = {
        'estadistica': pd.DataFrame(),
        'morbilidad': pd.DataFrame(),
        'germenes': pd.DataFrame(),
        'fallecidos': pd.DataFrame(),
        'eventos_adversos': pd.DataFrame(),
        'gram_sot': pd.DataFrame()
    }
    PROCESSOR = DataProcessor(DATA)

# ============================================
# LAYOUT DE LA APLICACIÓN
# ============================================

app.layout = html.Div([
    # Header
    html.Div([
        html.Div([
            html.Div([
                html.Img(src='/assets/Logo.png', style={
                    'height': '48px',
                    'marginRight': '16px'
                }, className="logo-img"),
                html.Div([
                    html.H1("UCI Adultos", style={
                        'color': '#f8fafc', 'fontWeight': '700', 'margin': '0',
                        'fontSize': '20px', 'lineHeight': '1.2'
                    }),
                    html.P("Panel de Indicadores 2026", style={
                        'color': '#94a3b8', 'fontSize': '13px', 'margin': '0',
                        'fontWeight': '400'
                    }, className="header-subtitle")
                ])
            ], style={'display': 'flex', 'alignItems': 'center'}),
            html.Div([
                dbc.Button("Exportar Excel", id="btn-export-excel", color="success", size="sm", className="me-2"),
                dbc.Button("Exportar CSV", id="btn-export-csv", color="primary", size="sm"),
                dcc.Download(id="download-excel"),
                dcc.Download(id="download-csv"),
            ], style={'display': 'flex', 'gap': '8px', 'flexWrap': 'wrap'})
        ], style={'display': 'flex', 'justifyContent': 'space-between', 'alignItems': 'center', 'flexWrap': 'wrap', 'gap': '16px'})
    ], className="header-main"),
    
    # Tabs
    dbc.Tabs(id="tabs", active_tab='tab-resumen', children=[
        dbc.Tab(label='Resumen General', tab_id='tab-resumen'),
        dbc.Tab(label='Gestión y Calidad', tab_id='tab-gestion'),
        dbc.Tab(label='Especialidades', tab_id='tab-especialidades'),
        dbc.Tab(label='Mortalidad', tab_id='tab-mortalidad'),
    ]),
    
    # Content
    html.Div(id='tab-content', className="tab-content"),
    
    # Footer
    html.Div([
        html.Div([
            html.Div([
                html.P("Hospital Infantil Universitario de San José", style={
                    'color': '#f1f5f9', 'fontWeight': '600', 'margin': '0', 'fontSize': '13px'
                }),
                html.P("Unidad de Cuidados Intensivos Adultos", style={
                    'color': '#94a3b8', 'margin': '0', 'fontSize': '12px'
                })
            ]),
            html.P("2026 Panel de Indicadores", style={
                'color': '#64748b', 'margin': '0', 'fontSize': '12px'
            })
        ], style={'display': 'flex', 'justifyContent': 'space-between', 'alignItems': 'center', 'flexWrap': 'wrap', 'gap': '12px'})
    ], className="dashboard-footer")
], style={
    'background': 'linear-gradient(135deg, #0f172a 0%, #1e293b 100%)',
    'minHeight': '100vh',
    'fontFamily': "'Inter', -apple-system, BlinkMacSystemFont, sans-serif"
}, id='app-container')


# ============================================
# CALLBACKS
# ============================================

@app.callback(Output('tab-content', 'children'), Input('tabs', 'active_tab'))
def render_content(tab):
    """Renderiza el contenido según la pestaña seleccionada"""
    
    if tab == 'tab-resumen':
        return render_resumen()
    elif tab == 'tab-gestion':
        return render_gestion()
    elif tab == 'tab-especialidades':
        return render_especialidades()
    elif tab == 'tab-mortalidad':
        return render_mortalidad()
    
    return html.Div([html.H5("Selecciona una pestaña", style={'color': '#1e293b'})])


def render_resumen():
    """Renderiza el módulo de Resumen General"""
    filtros = html.Div([
        html.Div([
            html.Span("Período de análisis", style={
                'color': '#f1f5f9', 'fontWeight': '600', 'fontSize': '14px'
            })
        ], style={'marginBottom': '12px'}),
        html.Div([
            html.Div([
                dcc.Dropdown(
                    id='filtro-mes-inicio',
                    options=[{'label': MESES_COMPLETOS[mes], 'value': mes} for mes in MESES],
                    value=MESES,
                    multi=True,
                    placeholder="Seleccionar meses...",
                    style={'fontSize': '13px', 'minWidth': '350px'}
                )
            ]),
            html.Div([
                dbc.Button("Aplicar", id='btn-aplicar-filtros', color="success", size="sm"),
                dbc.Button("Todos", id='btn-todos-meses', color="secondary", size="sm",
                          style={'marginLeft': '8px'})
            ], style={'marginLeft': '12px', 'display': 'flex', 'alignItems': 'center'})
        ], style={'display': 'flex', 'alignItems': 'center', 'flexWrap': 'wrap', 'gap': '8px'})
    ], style={
        'background': 'rgba(30, 41, 59, 0.8)',
        'padding': '16px 20px',
        'borderRadius': '8px',
        'marginBottom': '20px',
        'border': '1px solid rgba(148, 163, 184, 0.15)'
    }, className="filtros-section")
    
    return html.Div([filtros, html.Div(id='resumen-content')])


def render_gestion():
    """Renderiza el módulo de Gestión y Calidad"""
    try:
        kpis = PROCESSOR.get_kpis_gestion()
        
        kpi1 = chart_factory.create_kpi_card("Estancia Promedio", f"{kpis['promedio_estancia']:.2f}", "Días", "#2c3e50")
        kpi2 = chart_factory.create_kpi_card("Giro de Cama", f"{kpis['giro_cama']:.2f}", "Promedio", "#27ae60")
        kpi3 = chart_factory.create_kpi_card("Total Días Estancia", f"{int(kpis['total_dias_estancia'])}", "Ene - Dic", "#16a085")
        kpi4 = chart_factory.create_kpi_card("Capacidad UCI", f"{int(kpis['capacidad_uci'])}", "Camas", "#f39c12")
        kpi5 = chart_factory.create_kpi_card("Promedio Días VM", f"{kpis['promedio_dias_vm']:.2f}", "Ventilación Mecánica", "#9b59b6")
        
        # Gráficos
        datos_estancia = PROCESSOR.safe_get_row('PROMEDIO ESTANCIA', MESES)
        fig_estancia = chart_factory.create_bar_chart(datos_estancia, 'Estancia Promedio (Días)', color='#16a085', decimals=1)
        
        datos_giro = PROCESSOR.safe_get_row('GIRO CAMA', MESES)
        fig_giro = chart_factory.create_bar_chart(datos_giro, 'Giro de Cama', color='#27ae60', decimals=2)
        
        # Gráfico de VM
        fig_vm = go.Figure()
        fig_vm.add_trace(go.Bar(
            x=MESES, y=[round(kpis['promedio_dias_vm'], 2) for _ in MESES],
            marker_color='#9b59b6',
            text=[f'{round(kpis["promedio_dias_vm"], 1)}' for _ in MESES],
            textposition='inside',
            textfont=dict(color='white', size=11)
        ))
        fig_vm.update_layout(
            title={'text': 'Promedio de Días de VM', 'font': {'size': 14, 'color': '#f1f5f9'}, 'x': 0.5},
            plot_bgcolor='#1e293b', paper_bgcolor='#1e293b',
            height=350, bargap=0.3, margin=dict(l=50, r=30, t=50, b=50),
            xaxis={'tickfont': {'color': '#cbd5e1', 'size': 11}, 'gridcolor': 'rgba(148, 163, 184, 0.1)'},
            yaxis={'tickfont': {'color': '#94a3b8'}, 'gridcolor': 'rgba(148, 163, 184, 0.1)'}
        )
        
        # Infecciones
        fig_bacteriemia = chart_factory.create_line_chart(
            PROCESSOR.safe_get_row('TASA DE BACTERIEMIA ASOC.  A C. CENTRAL', MESES),
            'Tasa de ITS AC', color='#e74c3c', decimals=1
        )
        fig_nav = chart_factory.create_bar_chart(
            PROCESSOR.safe_get_row('TASA NEUMONÍA ASOCIADA A VENTILADOR', MESES),
            'Tasa NAV', color='#e74c3c', decimals=1
        )
        
        # Gérmenes - Donut profesional con diseño mejorado
        if not DATA['germenes'].empty:
            germenes_labels = DATA['germenes']['Germen'].tolist()
            germenes_values = DATA['germenes']['Cantidad'].tolist()
            
            # Ordenar de mayor a menor
            datos_germenes = sorted(zip(germenes_values, germenes_labels), reverse=True)
            germenes_values_sorted = [d[0] for d in datos_germenes]
            germenes_labels_sorted = [d[1] for d in datos_germenes]
            total_germenes = sum(germenes_values)
            
            # Paleta de colores distintiva
            colors_germenes = ['#ef4444', '#8b5cf6', '#10b981', '#f59e0b', '#06b6d4', '#ec4899', '#6366f1', '#14b8a6', '#f97316', '#a855f7', '#0ea5e9', '#84cc16', '#22c55e', '#e11d48', '#0891b2', '#d946ef', '#facc15', '#2563eb']
            
            fig_germenes = go.Figure(data=[go.Pie(
                labels=germenes_labels_sorted,
                values=germenes_values_sorted,
                hole=0.55,
                marker=dict(
                    colors=colors_germenes[:len(germenes_labels_sorted)], 
                    line=dict(color='#0f172a', width=2)
                ),
                textinfo='percent',
                textposition='inside',
                textfont=dict(size=9, color='white', family='Inter'),
                hovertemplate='<b>%{label}</b><br>Cantidad: %{value}<br>Porcentaje: %{percent}<extra></extra>',
                direction='clockwise',
                sort=False,
                pull=[0.03 if i == 0 else 0 for i in range(len(germenes_labels_sorted))]
            )])
            
            fig_germenes.add_annotation(
                text=f"<b style='font-size:22px'>{int(total_germenes)}</b><br><span style='color:#94a3b8'>Aislados</span>",
                x=0.5, y=0.5,
                font=dict(size=14, color='#f1f5f9', family='Inter'),
                showarrow=False
            )
            
            fig_germenes.update_layout(
                title={'text': 'Gérmenes Aislados (IAAS)', 'font': {'size': 16, 'color': '#f1f5f9', 'family': 'Inter'}, 'x': 0.5},
                plot_bgcolor='#1e293b', paper_bgcolor='#1e293b',
                height=520, margin=dict(l=30, r=280, t=60, b=30),
                legend=dict(
                    orientation='v', yanchor='middle', y=0.5, xanchor='left', x=0.95,
                    bgcolor='rgba(15, 23, 42, 0.9)', 
                    font=dict(color='#e2e8f0', size=9, family='Inter'),
                    bordercolor='rgba(148, 163, 184, 0.3)', borderwidth=1,
                    itemsizing='constant',
                    tracegroupgap=3
                ),
                showlegend=True
            )
        else:
            fig_germenes = go.Figure()
        
        # Eventos adversos - Barras verticales con etiquetas visibles
        if not DATA['eventos_adversos'].empty:
            df_eventos = DATA['eventos_adversos'].sort_values('Cantidad', ascending=False)
            colors_eventos = ['#ef4444', '#6366f1', '#10b981', '#f59e0b', '#8b5cf6', '#06b6d4', '#ec4899']
            fig_eventos = go.Figure(data=[go.Bar(
                x=df_eventos['Evento'],
                y=df_eventos['Cantidad'],
                marker_color=[colors_eventos[i % len(colors_eventos)] for i in range(len(df_eventos))],
                text=df_eventos['Cantidad'],
                textposition='outside',
                textfont=dict(color='#f1f5f9', size=11)
            )])
            fig_eventos.update_layout(
                title={'text': 'Eventos Adversos por Categoría', 'font': {'size': 14, 'color': '#f1f5f9'}, 'x': 0.5},
                plot_bgcolor='#1e293b', paper_bgcolor='#1e293b',
                height=550, margin=dict(l=50, r=30, t=50, b=280),
                xaxis={'tickfont': {'color': '#e2e8f0', 'size': 8}, 'tickangle': 45, 'gridcolor': 'rgba(148, 163, 184, 0.1)'},
                yaxis={'tickfont': {'color': '#94a3b8'}, 'gridcolor': 'rgba(148, 163, 184, 0.1)'}
            )
        else:
            fig_eventos = go.Figure()
        
        return html.Div([
            html.H5("Indicadores de Gestión", style={'color': '#f1f5f9', 'fontWeight': '600'}, className="mb-4"),
            dbc.Row([
                dbc.Col(kpi1, xs=12, md=True), 
                dbc.Col(kpi2, xs=12, md=True), 
                dbc.Col(kpi3, xs=12, md=True), 
                dbc.Col(kpi4, xs=12, md=True), 
                dbc.Col(kpi5, xs=12, md=True)
            ], className="mb-5 g-4"),
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_estancia, config=GRAPH_CONFIG), xs=12, md=6), 
                dbc.Col(dcc.Graph(figure=fig_giro, config=GRAPH_CONFIG), xs=12, md=6)
            ], className="mb-4 g-4"),
            dbc.Row([
                dbc.Col(xs=12, md=3),
                dbc.Col(dcc.Graph(figure=fig_vm, config=GRAPH_CONFIG), xs=12, md=6),
                dbc.Col(xs=12, md=3)
            ], className="mb-5", justify="center"),
            html.H5("Indicadores de Calidad e Infecciones", style={'color': '#f1f5f9', 'fontWeight': '600', 'marginTop': '30px', 'marginBottom': '30px', 'textAlign': 'center'}),
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_bacteriemia, config=GRAPH_CONFIG), xs=12, md=6), 
                dbc.Col(dcc.Graph(figure=fig_nav, config=GRAPH_CONFIG), xs=12, md=6)
            ], className="mb-4 g-4"),
            html.H5("Gérmenes Aislados y Eventos Adversos", style={'color': '#f1f5f9', 'fontWeight': '600', 'marginTop': '30px', 'marginBottom': '30px', 'textAlign': 'center'}),
            dbc.Row([
                dbc.Col(xs=12, md=2),
                dbc.Col(dcc.Graph(figure=fig_germenes, config=GRAPH_CONFIG), xs=12, md=8),
                dbc.Col(xs=12, md=2)
            ], className="mb-5", justify="center"),
            dbc.Row([dbc.Col(dcc.Graph(figure=fig_eventos, config=GRAPH_CONFIG), xs=12)])
        ])
    except Exception as e:
        logger.error(f"Error en render_gestion: {e}")
        return html.Div([html.H3(f"Error: {str(e)}", style={'color': '#e74c3c'})])


def render_especialidades():
    """Renderiza el módulo de Especialidades"""
    try:
        suma_col = PROCESSOR.get_suma_column()
        
        total_medicina = PROCESSOR.safe_get_value('MEDICINA INTERNA', suma_col)
        total_cirugia = PROCESSOR.safe_get_value('CIRUGIA GENERAL', suma_col)
        total_neuro = PROCESSOR.safe_get_value('NEUROCIRUGIA', suma_col)
        total_neurologia = PROCESSOR.safe_get_value('NEUROLOGIA', suma_col)
        
        total = max(total_medicina + total_cirugia + total_neuro + total_neurologia, 1)
        
        kpi1 = chart_factory.create_kpi_card("Medicina Interna", f"{int(total_medicina)}", f"{PROCESSOR.formatear_porcentaje(total_medicina/total*100, False)}%", "#3498db")
        kpi2 = chart_factory.create_kpi_card("Cirugía General", f"{int(total_cirugia)}", f"{PROCESSOR.formatear_porcentaje(total_cirugia/total*100, False)}%", "#27ae60")
        kpi3 = chart_factory.create_kpi_card("Neurocirugía", f"{int(total_neuro)}", f"{PROCESSOR.formatear_porcentaje(total_neuro/total*100, False)}%", "#16a085")
        kpi4 = chart_factory.create_kpi_card("Neurología", f"{int(total_neurologia)}", f"{PROCESSOR.formatear_porcentaje(total_neurologia/total*100, False)}%", "#9b59b6")
        
        fig_medicina = chart_factory.create_bar_chart(PROCESSOR.safe_get_row('MEDICINA INTERNA', MESES), 'Medicina Interna', color='#3498db', decimals=0)
        fig_cirugia = chart_factory.create_bar_chart(PROCESSOR.safe_get_row('CIRUGIA GENERAL', MESES), 'Cirugía General', color='#27ae60', decimals=0)
        fig_neuro = chart_factory.create_bar_chart(PROCESSOR.safe_get_row('NEUROCIRUGIA', MESES), 'Neurocirugía', color='#16a085', decimals=0)
        fig_neurologia = chart_factory.create_bar_chart(PROCESSOR.safe_get_row('NEUROLOGIA', MESES), 'Neurología', color='#9b59b6', decimals=0)
        
        return html.Div([
            html.H5("Principales Especialidades", style={'color': '#f1f5f9', 'fontWeight': '600', 'marginBottom': '20px'}),
            dbc.Row([dbc.Col(kpi1, xs=12, md=3), dbc.Col(kpi2, xs=12, md=3), dbc.Col(kpi3, xs=12, md=3), dbc.Col(kpi4, xs=12, md=3)], className="mb-4"),
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_medicina, config=GRAPH_CONFIG), xs=12, md=6), 
                dbc.Col(dcc.Graph(figure=fig_cirugia, config=GRAPH_CONFIG), xs=12, md=6)
            ], className="mb-4"),
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_neuro, config=GRAPH_CONFIG), xs=12, md=6), 
                dbc.Col(dcc.Graph(figure=fig_neurologia, config=GRAPH_CONFIG), xs=12, md=6)
            ])
        ])
    except Exception as e:
        logger.error(f"Error en render_especialidades: {e}")
        return html.Div([html.H3(f"Error: {str(e)}", style={'color': '#e74c3c'})])


def render_mortalidad():
    """Renderiza el módulo de Mortalidad"""
    try:
        kpis_mort = PROCESSOR.get_kpis_mortalidad()
        
        kpi1 = chart_factory.create_kpi_card("Total Fallecidos", f"{int(kpis_mort['total_fallecidos'])}", "Ene - Dic", "#e74c3c")
        kpi2 = chart_factory.create_kpi_card("% Mortalidad General", PROCESSOR.formatear_porcentaje(kpis_mort['mortalidad_general']), "Ene - Dic", "#c0392b")
        kpi3 = chart_factory.create_kpi_card("Fallecidos COVID-19", f"{int(kpis_mort['fallecidos_covid'])}", "Ene - Dic", "#34495e")
        kpi4 = chart_factory.create_kpi_card("% Mortalidad Trauma", PROCESSOR.formatear_porcentaje(kpis_mort['mortalidad_trauma']), "Promedio", "#f39c12")
        
        kpi_prom_dias = chart_factory.create_kpi_card("Promedio Días Estancia", f"{kpis_mort['promedio_dias_estancia']:.1f}", "Fallecidos", "#9b59b6")
        kpi_mediana = chart_factory.create_kpi_card("Mediana Días de Estancia", f"{kpis_mort['mediana_dias_estancia']:.1f}", "Fallecidos", "#3498db")
        
        fig_mort = chart_factory.create_line_chart(
            PROCESSOR.safe_get_row('% DE MORTALIDAD GENERAL', MESES),
            '% Mortalidad General', color='#e74c3c', decimals=1
        )
        
        fig_trauma = chart_factory.create_grouped_bar_chart(
            ['Pacientes con diagnóstico de Trauma', 'Pacientes fallecidos por Trauma'],
            ['Pacientes Trauma', 'Fallecidos Trauma'], 'Trauma: Pacientes vs Fallecidos',
            processor=PROCESSOR, decimals=0
        )
        
        fig_piramide = chart_factory.create_pyramid_chart(DATA['fallecidos'])
        
        return html.Div([
            dbc.Row([dbc.Col(kpi1, xs=12, md=3), dbc.Col(kpi2, xs=12, md=3), dbc.Col(kpi3, xs=12, md=3), dbc.Col(kpi4, xs=12, md=3)], className="mb-4"),
            html.H6("Días de Estancia - Fallecidos", style={'color': '#f1f5f9', 'fontWeight': '600', 'marginBottom': '15px', 'marginTop': '10px'}),
            dbc.Row([dbc.Col(kpi_prom_dias, xs=12, md=6), dbc.Col(kpi_mediana, xs=12, md=6)], className="mb-4"),
            dbc.Row([
                dbc.Col(xs=12, md=2),
                dbc.Col(dcc.Graph(figure=fig_mort, config=GRAPH_CONFIG), xs=12, md=8),
                dbc.Col(xs=12, md=2)
            ], className="mb-4", justify="center"),
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_trauma, config=GRAPH_CONFIG), xs=12, md=6), 
                dbc.Col(dcc.Graph(figure=fig_piramide, config=GRAPH_CONFIG), xs=12, md=6)
            ], className="mt-3")
        ])
    except Exception as e:
        logger.error(f"Error en render_mortalidad: {e}")
        return html.Div([html.H3(f"Error: {str(e)}", style={'color': '#e74c3c'})])


# Callback para Resumen
@app.callback(
    [Output('filtro-mes-inicio', 'value'),
     Output('resumen-content', 'children')],
    [Input('btn-todos-meses', 'n_clicks'),
     Input('btn-aplicar-filtros', 'n_clicks')],
    [State('filtro-mes-inicio', 'value')]
)
def actualizar_resumen(btn_todos, btn_aplicar, meses_seleccionados):
    ctx = dash.callback_context
    triggered_id = ctx.triggered[0]['prop_id'].split('.')[0] if ctx.triggered else None
    
    if triggered_id == 'btn-todos-meses':
        meses_seleccionados = MESES
    
    try:
        if meses_seleccionados and isinstance(meses_seleccionados, list) and len(meses_seleccionados) > 0:
            meses_filtrados = [m for m in meses_seleccionados if m in MESES]
        else:
            meses_filtrados = MESES
        
        # Texto del rango
        if len(meses_filtrados) == 1:
            rango_texto = MESES_COMPLETOS.get(meses_filtrados[0], meses_filtrados[0])
        elif len(meses_filtrados) == len(MESES):
            rango_texto = "Ene - Dic"
        else:
            rango_texto = ", ".join([MESES_COMPLETOS.get(m, m)[:3] for m in meses_filtrados])
        
        kpis = PROCESSOR.get_kpis_resumen(meses_filtrados)
        
        kpi_ingresos = chart_factory.create_kpi_card("Total Ingresos", f"{int(kpis['total_ingresos'])}", rango_texto, "#3498db")
        kpi_ocup = chart_factory.create_kpi_card("% Ocupación Promedio", PROCESSOR.formatear_porcentaje(kpis['promedio_ocupacional']), rango_texto, "#2c3e50")
        kpi_mort = chart_factory.create_kpi_card("% Mortalidad General", PROCESSOR.formatear_porcentaje(kpis['mortalidad']), f"Total {rango_texto}", "#e74c3c")
        kpi_estancia = chart_factory.create_kpi_card("Estancia Promedio", f"{kpis['promedio_estancia']:.2f}", f"Días - {rango_texto}", "#27ae60")
        kpi_giro = chart_factory.create_kpi_card("Giro de Cama", f"{kpis['giro_cama']:.2f}", f"Promedio {rango_texto}", "#f39c12")
        
        # EPS - Barras verticales con etiquetas visibles
        eps_labels, eps_values = PROCESSOR.get_eps_data()
        if eps_values:
            eps_df = pd.DataFrame({'EPS': eps_labels, 'Cantidad': eps_values})
            eps_df = eps_df.sort_values('Cantidad', ascending=False)
            colors_eps = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#06b6d4', '#ec4899']
            fig_eps = go.Figure(data=[go.Bar(
                x=eps_df['EPS'],
                y=eps_df['Cantidad'],
                marker_color=[colors_eps[i % len(colors_eps)] for i in range(len(eps_df))],
                text=eps_df['Cantidad'].astype(int),
                textposition='outside',
                textfont=dict(color='#f1f5f9', size=11)
            )])
            fig_eps.update_layout(
                title={'text': f'Distribución por EPS (Total: {int(sum(eps_values)):,})', 'font': {'size': 14, 'color': '#f1f5f9'}, 'x': 0.5},
                plot_bgcolor='#1e293b', paper_bgcolor='#1e293b',
                height=520, margin=dict(l=50, r=30, t=50, b=200),
                xaxis={'tickfont': {'color': '#e2e8f0', 'size': 9}, 'tickangle': 45, 'gridcolor': 'rgba(148, 163, 184, 0.1)'},
                yaxis={'tickfont': {'color': '#94a3b8'}, 'gridcolor': 'rgba(148, 163, 184, 0.1)'}
            )
        else:
            fig_eps = go.Figure()
        
        # Ocupacional
        datos_ocup = PROCESSOR.safe_get_row('% OCUPACIONAL', meses_filtrados)
        fig_ocup = go.Figure()
        fig_ocup.add_trace(go.Scatter(
            x=meses_filtrados, y=datos_ocup.values,
            mode='lines+markers+text',
            line=dict(color='#10b981', width=2),
            marker=dict(size=8, color='#10b981'),
            text=[f'{v:.1f}' if pd.notna(v) else '' for v in datos_ocup.values],
            textposition='top center',
            textfont=dict(color='#f1f5f9', size=10)
        ))
        fig_ocup.update_layout(
            title={'text': '% Ocupacional Mensual', 'font': {'size': 14, 'color': '#f1f5f9'}, 'x': 0.5},
            plot_bgcolor='#1e293b', paper_bgcolor='#1e293b',
            height=350, margin=dict(l=50, r=30, t=50, b=50),
            xaxis={'tickfont': {'color': '#cbd5e1', 'size': 11}, 'gridcolor': 'rgba(148, 163, 184, 0.1)'},
            yaxis={'tickfont': {'color': '#94a3b8'}, 'gridcolor': 'rgba(148, 163, 184, 0.1)'}
        )
        
        # Ingresos vs Egresos
        datos_ing = PROCESSOR.safe_get_row('NUMERO DE INGRESOS DEL MES', meses_filtrados)
        datos_eg = PROCESSOR.safe_get_row('NÚMERO DE EGRESOS DEL MES', meses_filtrados)
        fig_ie = go.Figure()
        fig_ie.add_trace(go.Bar(name='Ingresos', x=meses_filtrados, y=datos_ing.values, marker_color='#3b82f6', text=datos_ing.values, texttemplate='%{y:.0f}', textposition='inside', textfont=dict(color='white', size=9)))
        fig_ie.add_trace(go.Bar(name='Egresos', x=meses_filtrados, y=datos_eg.values, marker_color='#10b981', text=datos_eg.values, texttemplate='%{y:.0f}', textposition='inside', textfont=dict(color='white', size=9)))
        fig_ie.update_layout(
            title={'text': 'Ingresos vs. Egresos', 'font': {'size': 14, 'color': '#f1f5f9'}, 'x': 0.5},
            barmode='group', plot_bgcolor='#1e293b', paper_bgcolor='#1e293b',
            height=380, margin=dict(l=50, r=30, t=50, b=70),
            legend=dict(orientation="h", y=-0.15, x=0.5, xanchor="center", bgcolor='rgba(30, 41, 59, 0.8)', font={'color': '#f1f5f9', 'size': 12}, bordercolor='rgba(148, 163, 184, 0.3)', borderwidth=1),
            xaxis={'tickfont': {'color': '#cbd5e1', 'size': 11}, 'gridcolor': 'rgba(148, 163, 184, 0.1)'},
            yaxis={'tickfont': {'color': '#94a3b8'}, 'gridcolor': 'rgba(148, 163, 184, 0.1)'}
        )
        
        # Morbilidad - Barras verticales
        if not DATA['morbilidad'].empty:
            df_morb = DATA['morbilidad'].sort_values('Total', ascending=False)
            total_morb = df_morb['Total'].sum()
            
            # Paleta de colores vibrante
            colors_morb = ['#ef4444', '#8b5cf6', '#06b6d4', '#f59e0b', '#ec4899', '#10b981', '#6366f1', '#14b8a6', '#f97316', '#a855f7', '#0ea5e9', '#84cc16']
            
            fig_morb = go.Figure(data=[go.Bar(
                x=df_morb.index,
                y=df_morb['Total'],
                marker_color=colors_morb[:len(df_morb)],
                text=df_morb['Total'],
                textposition='outside',
                textfont=dict(color='#f1f5f9', size=11, family='Inter')
            )])
            
            fig_morb.update_layout(
                title={'text': f'Morbilidad por Diagnóstico (Total: {int(total_morb)})', 'font': {'size': 15, 'color': '#f1f5f9', 'family': 'Inter'}, 'x': 0.5},
                plot_bgcolor='#1e293b', paper_bgcolor='#1e293b',
                height=380, margin=dict(l=50, r=30, t=50, b=100),
                xaxis={'tickfont': {'color': '#e2e8f0', 'size': 9}, 'tickangle': -45, 'gridcolor': 'rgba(148, 163, 184, 0.1)'},
                yaxis={'tickfont': {'color': '#94a3b8'}, 'gridcolor': 'rgba(148, 163, 184, 0.1)'},
                bargap=0.25
            )
            
            # Morbilidad por Sexo - Gráfico de Torta (Pie)
            total_hombres = df_morb['Hombre'].sum()
            total_mujeres = df_morb['Mujer'].sum()
            total_sexo = total_hombres + total_mujeres
            pct_h = round(total_hombres / total_sexo * 100, 1)
            pct_m = round(total_mujeres / total_sexo * 100, 1)
            
            fig_morb_gen = go.Figure(data=[go.Pie(
                labels=['Hombres', 'Mujeres'],
                values=[total_hombres, total_mujeres],
                hole=0,  # Sin hueco = torta
                marker=dict(
                    colors=['#3b82f6', '#ec4899'], 
                    line=dict(color='#0f172a', width=2)
                ),
                textinfo='label+percent+value',
                textposition='inside',
                textfont=dict(size=14, color='white', family='Inter'),
                hovertemplate='<b>%{label}</b><br>Cantidad: %{value}<br>Porcentaje: %{percent}<extra></extra>',
                direction='clockwise',
                sort=False,
                pull=[0.02, 0.02]
            )])
            
            fig_morb_gen.update_layout(
                title={'text': f'Distribución por Sexo (Total: {int(total_sexo)})', 'font': {'size': 15, 'color': '#f1f5f9', 'family': 'Inter'}, 'x': 0.5},
                plot_bgcolor='#1e293b', paper_bgcolor='#1e293b',
                height=380, margin=dict(l=20, r=20, t=50, b=20),
                showlegend=False
            )
        else:
            fig_morb = go.Figure()
            fig_morb_gen = go.Figure()
        
        return meses_filtrados, html.Div([
            # KPIs en fila
            dbc.Row([
                dbc.Col(kpi_ingresos, xs=6, sm=6, md=True), 
                dbc.Col(kpi_ocup, xs=6, sm=6, md=True),
                dbc.Col(kpi_mort, xs=6, sm=6, md=True), 
                dbc.Col(kpi_estancia, xs=6, sm=6, md=True),
                dbc.Col(kpi_giro, xs=6, sm=6, md=True)
            ], className="mb-4 g-3"),
            
            # Gráficos de línea y barras agrupadas
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_ocup, config=GRAPH_CONFIG), xs=12, lg=6), 
                dbc.Col(dcc.Graph(figure=fig_ie, config=GRAPH_CONFIG), xs=12, lg=6)
            ], className="mb-4 g-3"),
            
            # EPS
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_eps, config=GRAPH_CONFIG), xs=12)
            ], className="mb-4 g-3"),
            
            # Morbilidad y Sexo lado a lado
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_morb, config=GRAPH_CONFIG), xs=12, lg=8),
                dbc.Col(dcc.Graph(figure=fig_morb_gen, config=GRAPH_CONFIG), xs=12, lg=4)
            ], className="mb-3 g-3")
        ])
    except Exception as e:
        logger.error(f"Error en actualizar_resumen: {e}")
        return dash.no_update, html.Div([html.H3(f"Error: {str(e)}", style={'color': '#dc2626'})])


# Callbacks de exportación
@app.callback(
    Output("download-excel", "data"),
    Input("btn-export-excel", "n_clicks"),
    prevent_initial_call=True
)
def export_excel(n_clicks):
    """Exporta datos a Excel"""
    from dashboard_uci import create_export_data
    datos = create_export_data(DATA)
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in datos.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    
    return dcc.send_bytes(output.getvalue(), "dashboard_uci_export.xlsx")


@app.callback(
    Output("download-csv", "data"),
    Input("btn-export-csv", "n_clicks"),
    prevent_initial_call=True
)
def export_csv(n_clicks):
    """Exporta datos a CSV (Resumen)"""
    from dashboard_uci import create_export_data
    datos = create_export_data(DATA)
    
    if 'Resumen General' in datos:
        return dcc.send_data_frame(datos['Resumen General'].to_csv, "dashboard_uci_resumen.csv")
    
    return dash.no_update


# ============================================
# PUNTO DE ENTRADA
# ============================================

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8050)
