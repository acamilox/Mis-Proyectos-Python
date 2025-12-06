# -*- coding: utf-8 -*-
import dash
from dash import dcc, html, Input, Output
import dash_bootstrap_components as dbc
import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
from openpyxl import load_workbook
import os

# Obtener directorio del script
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, 'datos_iaas.xlsx')

# Inicializar df globalmente
df = pd.DataFrame()

# --- 1. CARGA DE DATOS ---
try:
    # Leer Excel: la fila 2 (índice 2) contiene los headers
    df_raw = pd.read_excel(EXCEL_PATH, sheet_name='ESTADISTICA', header=2)
    
    # La primera columna contiene los nombres de las filas
    df = df_raw.set_index('DATOS')
    
    # Limpiar datos
    df = df.replace(['*', '#DIV/0!'], None)
    
    # Convertir columnas numéricas
    for col in df.columns:
        if col not in ['PROMEDIO', None]:
            try:
                df[col] = pd.to_numeric(df[col], errors='coerce')
            except:
                pass
    
    print(f"✅ Datos cargados: {len(df)} filas, {len(df.columns)} columnas")
    print(f"\n📋 Columnas disponibles: {list(df.columns)}")
    print(f"\n📋 Primeras 10 filas:")
    for i, idx in enumerate(df.index[:10]):
        print(f"  {i}: '{idx}'")
except Exception as e:
    print(f"❌ Error cargando datos: {e}")
    import traceback
    traceback.print_exc()

# --- CARGA DE DATOS DE MORBILIDAD ---
df_morbilidad = pd.DataFrame()
try:
    df_morb_raw = pd.read_excel(EXCEL_PATH, sheet_name='MORBILIDAD', header=None)
    
    # Procesar datos de morbilidad - extraer totales por diagnóstico
    diagnosticos = ['ACV', 'CARDIOVASCULAR', 'RESPIRATORIO', 'SEPSIS', 'POP', 'DM', 
                    'GINECO-OBSTE', 'TRAUMA', 'TOXICOLOGIA', 'OTROS']
    
    morbilidad_data = {diag: {'Hombre': 0, 'Mujer': 0, 'Total': 0} for diag in diagnosticos}
    
    for i, row in df_morb_raw.iterrows():
        diag = str(row[0]).strip().upper() if pd.notna(row[0]) else ''
        # Normalizar nombres de diagnósticos
        if 'CARVIO' in diag or 'CARDIO' in diag:
            diag = 'CARDIOVASCULAR'
        
        if diag in [d.upper() for d in diagnosticos]:
            diag_key = next((d for d in diagnosticos if d.upper() == diag), None)
            if diag_key:
                hombre = pd.to_numeric(row[1], errors='coerce')
                mujer = pd.to_numeric(row[2], errors='coerce')
                total = pd.to_numeric(row[3], errors='coerce')
                if pd.notna(hombre): morbilidad_data[diag_key]['Hombre'] += int(hombre)
                if pd.notna(mujer): morbilidad_data[diag_key]['Mujer'] += int(mujer)
                if pd.notna(total): morbilidad_data[diag_key]['Total'] += int(total)
    
    df_morbilidad = pd.DataFrame(morbilidad_data).T
    df_morbilidad = df_morbilidad.sort_values('Total', ascending=False)
    print(f"✅ Datos de morbilidad cargados: {len(df_morbilidad)} diagnósticos")
except Exception as e:
    print(f"⚠️ Error cargando morbilidad: {e}")

# --- CARGA DE DATOS DE GÉRMENES AISLADOS (IAAS) ---
df_germenes = pd.DataFrame()
try:
    df_iaas = pd.read_excel(EXCEL_PATH, sheet_name='IAAS', header=None)
    # Columna F (índice 5) contiene los gérmenes
    germenes_raw = df_iaas[5].dropna()
    germenes_raw = germenes_raw[germenes_raw.astype(str).str.strip() != '']
    
    # Normalizar nombres de gérmenes (simplificar nombres largos)
    def normalizar_germen(g):
        g = str(g).strip()
        if 'K. pneumoniae' in g or 'K.pneumoniae' in g:
            if 'carbapenemasa' in g.lower():
                return 'K. pneumoniae KPC'
            return 'K. pneumoniae'
        if 'Acinetobacter' in g:
            return 'A. baumannii'
        if 'E. coli' in g:
            return 'E. coli'
        if 'Pseudomona' in g:
            return 'P. aeruginosa'
        if 'S. hominis' in g:
            return 'S. hominis'
        if 'S. aureus' in g:
            return 'S. aureus'
        if 'Candida' in g:
            return 'Candida spp.'
        return g
    
    germenes_norm = germenes_raw.apply(normalizar_germen)
    df_germenes = germenes_norm.value_counts().reset_index()
    df_germenes.columns = ['Germen', 'Cantidad']
    print(f"✅ Datos de gérmenes cargados: {len(df_germenes)} tipos")
except Exception as e:
    print(f"⚠️ Error cargando gérmenes: {e}")

# --- CARGA DE OTROS EVENTOS ADVERSOS (celdas azules en columna E) ---
df_eventos_adversos = pd.DataFrame()
try:
    wb = load_workbook(EXCEL_PATH)
    ws = wb['EVENTOS ADVERSOS']
    
    eventos_azules = []
    for i in range(1, 100):
        cell = ws.cell(i, 5)  # Columna E
        if cell.value:
            fill_color = cell.fill.fgColor.rgb if cell.fill.fgColor else None
            # Detectar color azul FF00B0F0
            if fill_color and fill_color == 'FF00B0F0':
                evento = str(cell.value).strip()
                eventos_azules.append(evento)
    
    # Normalizar y agrupar eventos similares
    def normalizar_evento(e):
        e_lower = e.lower()
        if 'desplazamiento' in e_lower or 'autoretiro' in e_lower:
            return 'Desplazamiento/Autoretiro dispositivo'
        if 'error' in e_lower or 'medicamento' in e_lower:
            return 'Error medicamentos/procedimiento'
        if 'lecash' in e_lower or 'marsi' in e_lower or 'lesión' in e_lower:
            return 'Lesiones (LECASH/MARSI)'
        if 'caida' in e_lower:
            return 'Caídas'
        return 'Otros'
    
    eventos_norm = [normalizar_evento(e) for e in eventos_azules]
    df_eventos_adversos = pd.Series(eventos_norm).value_counts().reset_index()
    df_eventos_adversos.columns = ['Evento', 'Cantidad']
    print(f"✅ Datos de eventos adversos cargados: {len(eventos_azules)} eventos")
except Exception as e:
    print(f"⚠️ Error cargando eventos adversos: {e}")

# Meses disponibles (abreviados como en Excel)
meses_disponibles = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Agos', 'Sep', 'Oct']

# Mapeo de meses abreviados a nombres completos para mostrar en UI
meses_completos = {
    'Ene': 'Enero', 'Feb': 'Febrero', 'Mar': 'Marzo', 'Abr': 'Abril',
    'May': 'Mayo', 'Jun': 'Junio', 'Jul': 'Julio', 'Agos': 'Agosto',
    'Sep': 'Septiembre', 'Oct': 'Octubre', 'Nov': 'Noviembre', 'Dic': 'Diciembre'
}

COLOR_PALETTE = {
    'primary': '#2c3e50',      # Azul oscuro profesional
    'secondary': '#27ae60',    # Verde esmeralda
    'danger': '#e74c3c',       # Rojo coral
    'warning': '#f39c12',      # Naranja suave
    'info': '#34495e',         # Gris azulado
    'success': '#27ae60',      # Verde esmeralda
    'accent': '#16a085',       # Verde azulado
    'gradient': ['#2c3e50', '#34495e', '#16a085', '#27ae60']
}

def safe_get_row(row_name, columns, silent=False):
    try:
        # Buscar el nombre de fila de forma flexible (ignorando mayúsculas/minúsculas y espacios)
        matching_rows = [idx for idx in df.index if row_name.lower() in str(idx).lower()]
        if matching_rows:
            actual_row = matching_rows[0]
            if isinstance(columns, list):
                return df.loc[actual_row, columns]
            else:
                return df.loc[actual_row, columns]
        else:
            if not silent:
                print(f"⚠️ No se encontró la fila: '{row_name}'")
            if isinstance(columns, list):
                return pd.Series([0] * len(columns), index=columns)
            return 0
    except Exception as e:
        if not silent:
            print(f"⚠️ Error al obtener '{row_name}': {e}")
        if isinstance(columns, list):
            return pd.Series([0] * len(columns), index=columns)
        return 0

def create_bar_chart(row_name, title, color=None, silent=False, decimals=0):
    data = safe_get_row(row_name, meses_disponibles, silent=silent)
    fig = px.bar(x=meses_disponibles, y=data, title=title)
    
    # Formato del hover según decimales
    if decimals == 0:
        hover_format = '<b>%{x}</b><br>Valor: %{y:.0f}<extra></extra>'
    else:
        hover_format = f'<b>%{{x}}</b><br>Valor: %{{y:.{decimals}f}}<extra></extra>'
    
    fig.update_traces(
        marker_color=color or COLOR_PALETTE['primary'],
        marker_line_color='rgba(255,255,255,0.1)',
        marker_line_width=1,
        hovertemplate=hover_format
    )
    fig.update_layout(
        showlegend=False,
        plot_bgcolor='#1a202c',
        paper_bgcolor='#2d3748',
        title={'font': {'size': 15, 'color': '#e2e8f0', 'family': 'Arial, sans-serif'}, 'x': 0.5, 'xanchor': 'center'},
        xaxis={'gridcolor': '#4a5568', 'showline': True, 'linecolor': '#4a5568', 'color': '#a0aec0'},
        yaxis={'gridcolor': '#4a5568', 'showline': True, 'linecolor': '#4a5568', 'color': '#a0aec0', 'tickformat': 'd'},
        margin=dict(l=50, r=30, t=50, b=40),
        hovermode='closest',
        hoverlabel=dict(
            bgcolor='white',
            font_size=14,
            font_family="Arial",
            font_color='#1a202c',
            bordercolor='#4a5568'
        ),
        transition={'duration': 500}
    )
    return fig

def create_line_chart(row_name, title, color=None, silent=False, decimals=2):
    data = safe_get_row(row_name, meses_disponibles, silent=silent)
    fig = px.line(x=meses_disponibles, y=data, title=title, markers=True)
    
    # Formato del hover según decimales
    if decimals == 0:
        hover_format = '<b>%{x}</b><br>Valor: %{y:.0f}<extra></extra>'
    else:
        hover_format = f'<b>%{{x}}</b><br>Valor: %{{y:.{decimals}f}}<extra></extra>'
    
    fig.update_traces(
        line_color=color or '#27ae60',
        line_width=3,
        marker=dict(size=10, line=dict(width=2, color='#2d3748')),
        hovertemplate=hover_format
    )
    
    # Configurar yaxis - para rangos pequeños usar dtick=1 para evitar repetición de valores
    yaxis_config = {'gridcolor': '#4a5568', 'showline': True, 'linecolor': '#4a5568', 'color': '#a0aec0'}
    if isinstance(data, pd.Series):
        max_val = data.max() if not data.empty else 0
        if pd.notna(max_val) and max_val <= 5:
            yaxis_config['dtick'] = 1  # Forzar ticks cada 1 unidad para rangos pequeños
    
    fig.update_layout(
        showlegend=False,
        plot_bgcolor='#1a202c',
        paper_bgcolor='#2d3748',
        title={'font': {'size': 15, 'color': '#e2e8f0', 'family': 'Arial, sans-serif'}, 'x': 0.5, 'xanchor': 'center'},
        xaxis={'gridcolor': '#4a5568', 'showline': True, 'linecolor': '#4a5568', 'color': '#a0aec0'},
        yaxis=yaxis_config,
        margin=dict(l=50, r=30, t=50, b=40),
        hovermode='closest',
        hoverlabel=dict(
            bgcolor='white',
            font_size=14,
            font_family="Arial",
            font_color='#1a202c',
            bordercolor='#4a5568'
        ),
        transition={'duration': 500}
    )
    return fig

def create_grouped_bar_chart(row_names, labels, title, decimals=0):
    """Crear gráfico de barras agrupadas"""
    fig = go.Figure()
    colors = ['#2c3e50', '#27ae60', '#16a085', '#f39c12']
    
    for i, (row_name, label) in enumerate(zip(row_names, labels)):
        data = safe_get_row(row_name, meses_disponibles)
        
        # Formato del hover según decimales
        if decimals == 0:
            hover_template = f'<b>%{{x}}</b><br>{label}: %{{y:.0f}}<extra></extra>'
        else:
            hover_template = f'<b>%{{x}}</b><br>{label}: %{{y:.{decimals}f}}<extra></extra>'
        
        fig.add_trace(go.Bar(
            name=label,
            x=meses_disponibles,
            y=data,
            marker_color=colors[i % len(colors)],
            marker_line_color='rgba(0,0,0,0.1)',
            marker_line_width=1,
            hovertemplate=hover_template
        ))
    
    fig.update_layout(
        title={
            'text': title, 
            'font': {'size': 15, 'color': '#e2e8f0', 'family': 'Arial, sans-serif'}, 
            'x': 0.5, 
            'xanchor': 'center',
            'y': 0.98,
            'yanchor': 'top'
        },
        barmode='group',
        plot_bgcolor='#1a202c',
        paper_bgcolor='#2d3748',
        xaxis={'gridcolor': '#4a5568', 'showline': True, 'linecolor': '#4a5568', 'color': '#e2e8f0', 'tickfont': {'size': 12, 'color': '#e2e8f0', 'weight': 'bold'}},
        yaxis={'gridcolor': '#4a5568', 'showline': True, 'linecolor': '#4a5568', 'color': '#e2e8f0', 'tickfont': {'size': 12, 'color': '#e2e8f0', 'weight': 'bold'}, 'tickformat': 'd'},
        margin=dict(l=50, r=30, t=90, b=40),
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.15,
            xanchor="center",
            x=0.5,
            bgcolor='rgba(26,32,44,0.95)',
            bordercolor='#4a5568',
            borderwidth=1,
            font={'color': '#ffffff', 'size': 13, 'weight': 'bold'}
        ),
        hovermode='closest',
        hoverlabel=dict(
            bgcolor='rgba(255, 255, 255, 0.95)',
            font_size=13,
            font_family="Arial",
            bordercolor='#4a5568'
        ),
        transition={'duration': 500}
    )
    return fig

def create_kpi_card(title, value, subtitle="", color="#3498db", icon="📊"):
    return dbc.Card(
        dbc.CardBody([
            html.Div([
                html.Span(icon, style={
                    'fontSize': '32px', 
                    'marginRight': '12px',
                    'filter': 'drop-shadow(0 0 8px rgba(255,255,255,0.3))',
                    'flexShrink': '0'
                }),
                html.Div([
                    html.H6(title, style={
                        'color': '#e2e8f0', 
                        'fontSize': '12px', 
                        'marginBottom': '8px', 
                        'fontWeight': '600',
                        'textTransform': 'uppercase',
                        'letterSpacing': '0.5px',
                        'textShadow': '0 2px 4px rgba(0,0,0,0.5)',
                        'lineHeight': '1.2',
                        'whiteSpace': 'nowrap',
                        'overflow': 'hidden',
                        'textOverflow': 'ellipsis'
                    }),
                    html.H2(value, style={
                        'color': '#ffffff', 
                        'fontWeight': '900', 
                        'marginBottom': '5px', 
                        'fontSize': '28px',
                        'textShadow': f'0 0 30px {color}, 0 0 60px {color}, 0 0 90px {color}60, 0 4px 8px rgba(0,0,0,0.8)',
                        'letterSpacing': '-1px',
                        'lineHeight': '1'
                    }),
                    html.P(subtitle, style={
                        'color': '#cbd5e0', 
                        'fontSize': '11px', 
                        'margin': '0',
                        'fontWeight': '500',
                        'textShadow': '0 1px 2px rgba(0,0,0,0.5)',
                        'whiteSpace': 'nowrap'
                    })
                ], style={'flex': '1', 'minWidth': '0', 'overflow': 'hidden'})
            ], style={'display': 'flex', 'alignItems': 'center'})
        ], style={'padding': '12px'}),
        className="kpi-card",
        style={
            'border': 'none',
            'borderLeft': f'4px solid {color}',
            'borderRadius': '12px',
            'boxShadow': f'0 8px 20px rgba(0,0,0,0.4), 0 0 40px {color}20, inset 0 1px 0 rgba(255,255,255,0.1)',
            'background': 'linear-gradient(145deg, #1a202c 0%, #2d3748 100%)',
            'transition': 'all 0.3s ease',
            'cursor': 'pointer',
            'height': '100%',
            'minHeight': '100px'
        }
    )

app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP], suppress_callback_exceptions=True)
server = app.server  # Para Render/Gunicorn

# CSS personalizado
app.index_string = '''
<!DOCTYPE html>
<html>
    <head>
        {%metas%}
        <title>{%title%}</title>
        {%favicon%}
        {%css%}
        <style>
            @keyframes fadeIn {
                from { opacity: 0; transform: translateY(20px); }
                to { opacity: 1; transform: translateY(0); }
            }
            
            .kpi-card:hover {
                transform: translateY(-8px) scale(1.02);
                box-shadow: 0 12px 24px rgba(0,0,0,0.4) !important;
            }
            
            .fade-in {
                animation: fadeIn 0.6s ease-out;
            }
            
            .nav-tabs .nav-link {
                color: #a0aec0;
                border: none;
                border-bottom: 3px solid transparent;
                transition: all 0.3s ease;
                background: transparent;
            }
            
            .nav-tabs .nav-link:hover {
                border-bottom: 3px solid #16a085;
                color: #16a085;
                text-shadow: 0 0 10px rgba(22,160,133,0.5);
            }
            
            .nav-tabs .nav-link.active {
                color: #16a085;
                border-bottom: 3px solid #16a085;
                background: transparent;
                font-weight: 600;
                text-shadow: 0 0 15px rgba(22,160,133,0.6);
            }
            
            /* Estilos para dropdowns oscuros - Dash dcc.Dropdown */
            .Select-control, .Select--single > .Select-control {
                background-color: #2d3748 !important;
                border-color: #4a5568 !important;
            }
            
            .Select-value-label, .Select-placeholder, .Select-value {
                color: #ffffff !important;
            }
            
            .Select-input > input {
                color: #ffffff !important;
            }
            
            .Select-menu-outer, .Select-menu {
                background-color: #2d3748 !important;
                border-color: #4a5568 !important;
            }
            
            .Select-option, .VirtualizedSelectOption {
                background-color: #2d3748 !important;
                color: #ffffff !important;
            }
            
            .Select-option:hover, .Select-option.is-focused {
                background-color: #16a085 !important;
                color: #ffffff !important;
            }
            
            .Select-option.is-selected {
                background-color: #16a085 !important;
                color: #ffffff !important;
            }
            
            .Select-arrow-zone .Select-arrow {
                border-color: #ffffff transparent transparent !important;
            }
            
            .Select.is-open > .Select-control .Select-arrow {
                border-color: transparent transparent #ffffff !important;
            }
            
            .Select-clear-zone {
                color: #ffffff !important;
            }
            
            /* Forzar color blanco en el texto seleccionado */
            .Select--single > .Select-control .Select-value .Select-value-label {
                color: #ffffff !important;
            }
            
            .Select-placeholder {
                color: #a0aec0 !important;
            }
            
            /* Para el input de búsqueda */
            .Select-input input {
                color: #ffffff !important;
                background: transparent !important;
            }
            
            /* Estilos adicionales para Dash Dropdown (versión moderna) */
            .dash-dropdown .Select-control {
                background-color: #2d3748 !important;
                border: 1px solid #4a5568 !important;
                border-radius: 8px !important;
            }
            
            .dash-dropdown .Select-value-label {
                color: #ffffff !important;
            }
            
            .dash-dropdown .Select-placeholder {
                color: #a0aec0 !important;
            }
            
            .dash-dropdown .Select-input input {
                color: #ffffff !important;
            }
            
            .dash-dropdown .Select-menu-outer {
                background-color: #2d3748 !important;
                border: 1px solid #4a5568 !important;
                border-radius: 0 0 8px 8px !important;
            }
            
            .dash-dropdown .VirtualizedSelectOption {
                background-color: #2d3748 !important;
                color: #ffffff !important;
            }
            
            .dash-dropdown .VirtualizedSelectFocusedOption {
                background-color: #16a085 !important;
                color: #ffffff !important;
            }
            
            /* Estilos para el contenedor del dropdown */
            #filtro-mes .Select-control,
            #filtro-anio .Select-control {
                background-color: #2d3748 !important;
                border: 1px solid #4a5568 !important;
                border-radius: 8px !important;
                min-height: 38px !important;
            }
            
            #filtro-mes .Select-value-label,
            #filtro-anio .Select-value-label {
                color: #ffffff !important;
                font-weight: 500 !important;
            }
            
            #filtro-mes .Select-arrow,
            #filtro-anio .Select-arrow {
                border-color: #ffffff transparent transparent !important;
            }
            
            /* Estilo para el RangeSlider */
            .rc-slider-track {
                background-color: #16a085 !important;
            }
            
            .rc-slider-handle {
                border-color: #16a085 !important;
                background-color: #16a085 !important;
            }
            
            .rc-slider-rail {
                background-color: #4a5568 !important;
            }
        </style>
    </head>
    <body>
        {%app_entry%}
        <footer>
            {%config%}
            {%scripts%}
            {%renderer%}
        </footer>
    </body>
</html>
'''

app.layout = html.Div([
    dbc.Container([
        # Header con animación
        html.Div([
            html.Div([
                html.H1("🏥 Dashboard UCI Adultos 2025", style={
                    'color': '#ffffff',
                    'fontWeight': '800',
                    'marginBottom': '10px',
                    'fontSize': '36px',
                    'letterSpacing': '1px',
                    'textShadow': '0 0 40px rgba(22,160,133,0.8), 0 0 80px rgba(22,160,133,0.4), 0 4px 8px rgba(0,0,0,0.5)',
                    'textTransform': 'uppercase'
                }),
                html.P("Sistema de Monitoreo y Análisis de Indicadores", style={
                    'color': '#e2e8f0',
                    'fontSize': '15px',
                    'marginBottom': '20px',
                    'fontWeight': '500',
                    'letterSpacing': '0.5px',
                    'textShadow': '0 2px 4px rgba(0,0,0,0.5)'
                })
            ]),
            html.Div([
                dbc.Button("📊 Exportar a Excel", color="success", size="sm", className="me-2", 
                          style={'borderRadius': '8px', 'fontWeight': '500'}),
                dbc.Button("📄 Exportar a CSV", color="primary", size="sm",
                          style={'borderRadius': '8px', 'fontWeight': '500'}),
            ], style={'marginBottom': '20px'})
        ], className="fade-in"),
        
        # Tabs con estilo mejorado
        dbc.Tabs(id="tabs", active_tab='tab-resumen', children=[
            dbc.Tab(label='📊 Resumen General', tab_id='tab-resumen'),
            dbc.Tab(label='📈 Gestión y Calidad', tab_id='tab-gestion'),
            dbc.Tab(label='💀 Mortalidad', tab_id='tab-mortalidad'),
            dbc.Tab(label='🏥 Especialidades y FPS', tab_id='tab-especialidades'),
        ], style={'marginBottom': '25px', 'borderBottom': '2px solid #e8eaed'}),
        
        html.Div(id='tab-content', className="p-3")
    
    ], fluid=True, style={
        'background': 'linear-gradient(145deg, #0f2027 0%, #203a43 50%, #2c5364 100%)',
        'padding': '35px',
        'minHeight': '100vh',
        'borderRadius': '20px',
        'boxShadow': '0 20px 60px rgba(0,0,0,0.4), inset 0 1px 0 rgba(255,255,255,0.1)'
    })
], style={
    'background': 'linear-gradient(135deg, #0f2027 0%, #203a43 25%, #2c5364 50%, #0f2027 75%, #203a43 100%)',
    'minHeight': '100vh',
    'padding': '25px',
    'fontFamily': 'Arial, sans-serif'
})

@app.callback(Output('tab-content', 'children'), Input('tabs', 'active_tab'))
def render_content(tab):
    if tab == 'tab-resumen':
        try:
            if len(df.columns) == 0:
                return html.Div([html.H3("⚠️ No hay datos")])
            
            # Filtros de mes y año con botón Aplicar
            filtros = html.Div([
                dbc.Row([
                    dbc.Col([
                        html.Label("📅 Mes Inicio:", style={'color': '#e2e8f0', 'fontWeight': '600', 'marginBottom': '8px'}),
                        dcc.Dropdown(
                            id='filtro-mes-inicio',
                            options=[{'label': meses_completos[mes], 'value': mes} for mes in meses_disponibles],
                            value='Ene',
                            style={'backgroundColor': '#2d3748', 'color': '#1a202c'},
                            className='dropdown-dark',
                            clearable=False
                        )
                    ], xs=12, md=3),
                    dbc.Col([
                        html.Label("📅 Mes Fin:", style={'color': '#e2e8f0', 'fontWeight': '600', 'marginBottom': '8px'}),
                        dcc.Dropdown(
                            id='filtro-mes-fin',
                            options=[{'label': meses_completos[mes], 'value': mes} for mes in meses_disponibles],
                            value='Oct',
                            style={'backgroundColor': '#2d3748', 'color': '#1a202c'},
                            className='dropdown-dark',
                            clearable=False
                        )
                    ], xs=12, md=3),
                    dbc.Col([
                        html.Label("📆 Año:", style={'color': '#e2e8f0', 'fontWeight': '600', 'marginBottom': '8px'}),
                        dcc.Dropdown(
                            id='filtro-anio',
                            options=[
                                {'label': '2025', 'value': '2025'},
                                {'label': '2024', 'value': '2024'},
                                {'label': '2023', 'value': '2023'},
                            ],
                            value='2025',
                            style={'backgroundColor': '#2d3748', 'color': '#1a202c'},
                            className='dropdown-dark',
                            clearable=False
                        )
                    ], xs=12, md=3),
                    dbc.Col([
                        html.Label(" ", style={'color': 'transparent', 'marginBottom': '8px', 'display': 'block'}),
                        dbc.Button(
                            "🔍 Aplicar Filtros",
                            id='btn-aplicar-filtros',
                            color="success",
                            className="w-100",
                            style={
                                'borderRadius': '8px',
                                'fontWeight': '600',
                                'padding': '9px 20px',
                                'background': 'linear-gradient(145deg, #16a085 0%, #27ae60 100%)',
                                'border': 'none',
                                'boxShadow': '0 4px 15px rgba(22,160,133,0.4)',
                                'transition': 'all 0.3s ease',
                                'height': '38px',
                                'marginTop': '0px'
                            }
                        )
                    ], xs=12, md=3, style={'display': 'flex', 'flexDirection': 'column', 'justifyContent': 'flex-end'}),
                ], className="mb-2"),
            ], style={
                'background': 'linear-gradient(145deg, #1a202c 0%, #2d3748 100%)',
                'padding': '20px',
                'borderRadius': '15px',
                'marginBottom': '25px',
                'boxShadow': '0 4px 15px rgba(0,0,0,0.3)'
            })
            
            # Contenedor para KPIs y gráficos (se actualizará con el callback)
            return html.Div([
                filtros,
                html.Div(id='resumen-content')
            ])
        except Exception as e:
            return html.Div([html.H3(f"Error: {str(e)}")])



    elif tab == 'tab-mortalidad':
        try:
            suma_col = [c for c in df.columns if str(c).strip() in ['∑', 'SUMA']]
            suma_col = suma_col[0] if suma_col else '∑'
            
            # KPIs de mortalidad
            total_fallecidos = safe_get_row('FALLECIDOS', suma_col)
            fallecidos_covid = safe_get_row('FALLECIDOS COVID-19', suma_col)
            total_egresos = safe_get_row('NÚMERO DE EGRESOS DEL MES', suma_col)
            pacientes_trauma = safe_get_row('Pacientes con diagnóstico de Trauma', suma_col)
            fallecidos_trauma = safe_get_row('Pacientes fallecidos por Trauma', suma_col)
            
            if pd.isna(total_fallecidos): total_fallecidos = 0
            if pd.isna(fallecidos_covid): fallecidos_covid = 0
            if pd.isna(total_egresos) or total_egresos == 0: total_egresos = 1
            if pd.isna(pacientes_trauma): pacientes_trauma = 0
            if pd.isna(fallecidos_trauma): fallecidos_trauma = 0
            
            mortalidad_general = (total_fallecidos / total_egresos * 100)
            # Mortalidad trauma: suma de % mortalidad mensual / 10 meses con datos
            datos_mortalidad_trauma = safe_get_row('% mortalidad en trauma', meses_disponibles, silent=True)
            mortalidad_trauma = datos_mortalidad_trauma.sum() / 10 if isinstance(datos_mortalidad_trauma, pd.Series) else 0
            
            kpi1 = create_kpi_card("Total Fallecidos", f"{int(total_fallecidos)}", "Ene - Oct", "#e74c3c", "💀")
            kpi2 = create_kpi_card("% Mortalidad General", f"{mortalidad_general:.2f}%", "Ene - Oct", "#c0392b", "⚰️")
            kpi3 = create_kpi_card("Fallecidos COVID-19", f"{int(fallecidos_covid)}", "Ene - Oct", "#34495e", "🦠")
            kpi4 = create_kpi_card("% Mortalidad Trauma", f"{mortalidad_trauma:.2f}%", "Promedio 10 meses", "#f39c12", "🚑")
            
            fig_fallecidos = create_bar_chart('FALLECIDOS', 'Número de Fallecidos', color='#e74c3c', decimals=0)
            fig_mortalidad = create_line_chart('% DE MORATLIDAD GENERAL', '% Mortalidad General', color='#e74c3c', decimals=2)
            fig_trauma = create_grouped_bar_chart(
                ['Pacientes con diagnóstico de Trauma', 'Pacientes fallecidos por Trauma'],
                ['Pacientes Trauma', 'Fallecidos Trauma'],
                'Trauma: Pacientes vs Fallecidos',
                decimals=0
            )
            
            return html.Div([
                dbc.Row([
                    dbc.Col(kpi1, xs=12, md=3),
                    dbc.Col(kpi2, xs=12, md=3),
                    dbc.Col(kpi3, xs=12, md=3),
                    dbc.Col(kpi4, xs=12, md=3),
                ], className="mb-4"),
                dbc.Row([
                    dbc.Col(dcc.Graph(figure=fig_fallecidos, config={'displayModeBar': False}), xs=12, md=6),
                    dbc.Col(dcc.Graph(figure=fig_mortalidad, config={'displayModeBar': False}), xs=12, md=6),
                ]),
                dbc.Row([
                    dbc.Col(dcc.Graph(figure=fig_trauma, config={'displayModeBar': False}), xs=12, md=6),
                ], className="mt-3")
            ])
        except Exception as e:
            return html.Div([html.H3(f"Error: {str(e)}")])

    elif tab == 'tab-especialidades':
        try:
            suma_col = [c for c in df.columns if str(c).strip() in ['∑', 'SUMA']]
            suma_col = suma_col[0] if suma_col else '∑'
            
            # KPIs de Especialidades
            total_medicina_interna = safe_get_row('MEDICINA INTERNA', suma_col, silent=True)
            total_cirugia = safe_get_row('CIRUGIA GENERAL', suma_col, silent=True)
            total_neurocirugia = safe_get_row('NEUROCIRUGIA', suma_col, silent=True)
            total_ortopedia = safe_get_row('ORTOPEDIA', suma_col, silent=True)
            
            if pd.isna(total_medicina_interna): total_medicina_interna = 0
            if pd.isna(total_cirugia): total_cirugia = 0
            if pd.isna(total_neurocirugia): total_neurocirugia = 0
            if pd.isna(total_ortopedia): total_ortopedia = 0
            
            kpi1 = create_kpi_card("Medicina Interna", f"{int(total_medicina_interna)}", "Total pacientes", "#2c3e50", "⚕️")
            kpi2 = create_kpi_card("Cirugía General", f"{int(total_cirugia)}", "Total pacientes", "#27ae60", "🔪")
            kpi3 = create_kpi_card("Neurocirugía", f"{int(total_neurocirugia)}", "Total pacientes", "#16a085", "🧠")
            kpi4 = create_kpi_card("Ortopedia", f"{int(total_ortopedia)}", "Total pacientes", "#9b59b6", "🦴")
            
            # Gráficos de especialidades principales
            fig_medicina = create_bar_chart('MEDICINA INTERNA', 'Medicina Interna', color='#2c3e50', silent=True, decimals=0)
            fig_cirugia = create_bar_chart('CIRUGIA GENERAL', 'Cirugía General', color='#27ae60', silent=True, decimals=0)
            fig_neurocirugia = create_bar_chart('NEUROCIRUGIA', 'Neurocirugía', color='#16a085', silent=True, decimals=0)
            fig_ortopedia = create_bar_chart('ORTOPEDIA', 'Ortopedia', color='#9b59b6', silent=True, decimals=0)
            
            return html.Div([
                html.H5("Principales Especialidades", style={'color': '#e2e8f0', 'fontWeight': '600'}, className="mb-3"),
                dbc.Row([
                    dbc.Col(kpi1, xs=12, md=3),
                    dbc.Col(kpi2, xs=12, md=3),
                    dbc.Col(kpi3, xs=12, md=3),
                    dbc.Col(kpi4, xs=12, md=3),
                ], className="mb-4"),
                dbc.Row([
                    dbc.Col(dcc.Graph(figure=fig_medicina, config={'displayModeBar': False}), xs=12, md=3),
                    dbc.Col(dcc.Graph(figure=fig_cirugia, config={'displayModeBar': False}), xs=12, md=3),
                    dbc.Col(dcc.Graph(figure=fig_neurocirugia, config={'displayModeBar': False}), xs=12, md=3),
                    dbc.Col(dcc.Graph(figure=fig_ortopedia, config={'displayModeBar': False}), xs=12, md=3),
                ])
            ])
        except Exception as e:
            return html.Div([html.H3(f"Error: {str(e)}")])

    elif tab == 'tab-gestion':
        try:
            suma_col = [c for c in df.columns if str(c).strip() in ['∑', 'SUMA']]
            suma_col = suma_col[0] if suma_col else '∑'
            
            # KPIs de gestión - calcular promedio desde los meses
            datos_estancia = safe_get_row('PROMEDIO ESTANCIA', meses_disponibles)
            promedio_estancia = datos_estancia[datos_estancia > 0].mean() if isinstance(datos_estancia, pd.Series) and len(datos_estancia[datos_estancia > 0]) > 0 else 0
            
            datos_giro = safe_get_row('GIRO CAMA', meses_disponibles)
            giro_cama = datos_giro[datos_giro > 0].mean() if isinstance(datos_giro, pd.Series) and len(datos_giro[datos_giro > 0]) > 0 else 0
            
            total_dias_estancia = safe_get_row('TOTAL DIAS ESTANCIA', suma_col)
            capacidad = safe_get_row('CAPACIDAD INSTALADA: UCI', 'Ene')
            
            if pd.isna(promedio_estancia): promedio_estancia = 0
            if pd.isna(giro_cama): giro_cama = 0
            if pd.isna(total_dias_estancia): total_dias_estancia = 0
            if pd.isna(capacidad): capacidad = 0
            
            kpi1 = create_kpi_card("Estancia Promedio", f"{promedio_estancia:.2f}", "Días", "#2c3e50", "⏱️")
            kpi2 = create_kpi_card("Giro de Cama", f"{giro_cama:.2f}", "Promedio", "#27ae60", "🔄")
            kpi3 = create_kpi_card("Total Días Estancia", f"{int(total_dias_estancia)}", "Ene - Oct", "#16a085", "📅")
            kpi4 = create_kpi_card("Capacidad UCI", f"{int(capacidad)}", "Camas", "#f39c12", "🛏️")
            
            fig_estancia = create_line_chart('PROMEDIO ESTANCIA', 'Estancia Promedio (Días)', color='#16a085', decimals=2)
            fig_giro = create_line_chart('GIRO CAMA', 'Giro de Cama', color='#27ae60', decimals=2)
            # Gráficos de infecciones
            fig_bacteriemia = create_line_chart('TASA DE BACTERIEMIA ASOC. A  C. CENTRAL', 'Tasa de Bacteriemia', color='#e74c3c', silent=True, decimals=2)
            fig_nav = create_line_chart('TASA NEUMONÍA ASOCIADA A VENTILADOR', 'Tasa NAV', color='#e74c3c', silent=True, decimals=2)
            fig_istu = create_line_chart('TASA INFECCIÓN ASOCIADA A C. URINARIO', 'Tasa de ISTU AC', color='#f39c12', silent=True, decimals=2)
            # Gráficos de seguridad del paciente
            fig_lesiones = create_line_chart('TASA LESIONES POR PRESION', 'Tasa Lesiones por Presión', color='#9b59b6', silent=True, decimals=2)
            fig_autoextubaciones = create_line_chart('% AUTOEXTUBACIONES', '% Autoextubaciones', color='#3498db', silent=True, decimals=2)
            fig_caidas = create_line_chart('% CAÍDAS', '% Caídas', color='#e67e22', silent=True, decimals=2)
            
            # Gráfico de Gérmenes Aislados
            fig_germenes = go.Figure()
            if not df_germenes.empty:
                colors_germenes = ['#e74c3c', '#3498db', '#2ecc71', '#f39c12', '#9b59b6', 
                                   '#1abc9c', '#e67e22', '#34495e', '#16a085', '#c0392b']
                fig_germenes = go.Figure(data=[
                    go.Bar(
                        x=df_germenes['Germen'],
                        y=df_germenes['Cantidad'],
                        marker_color=colors_germenes[:len(df_germenes)],
                        hovertemplate='<b>%{x}</b><br>Casos: %{y}<extra></extra>'
                    )
                ])
                fig_germenes.update_layout(
                    title={'text': 'Gérmenes Aislados (IAAS)', 'font': {'size': 15, 'color': '#e2e8f0'}, 'x': 0.5, 'xanchor': 'center'},
                    plot_bgcolor='#1a202c',
                    paper_bgcolor='#2d3748',
                    xaxis={'gridcolor': '#4a5568', 'color': '#e2e8f0', 'tickangle': -45},
                    yaxis={'gridcolor': '#4a5568', 'color': '#e2e8f0'},
                    margin=dict(l=50, r=30, t=50, b=100),
                    hoverlabel=dict(bgcolor='white', font_size=14, font_color='#1a202c')
                )
            
            # Gráfico de Otros Eventos Adversos
            colors_eventos = ['#e74c3c', '#3498db', '#2ecc71', '#f39c12', '#9b59b6']
            if not df_eventos_adversos.empty:
                fig_eventos = go.Figure(data=[
                    go.Bar(
                        x=df_eventos_adversos['Evento'],
                        y=df_eventos_adversos['Cantidad'],
                        marker_color=colors_eventos[:len(df_eventos_adversos)],
                        hovertemplate='<b>%{x}</b><br>Casos: %{y}<extra></extra>'
                    )
                ])
            else:
                fig_eventos = go.Figure(data=[go.Bar(x=['Sin datos'], y=[0], marker_color='#4a5568')])
            
            fig_eventos.update_layout(
                title={'text': 'Otros Eventos Adversos', 'font': {'size': 15, 'color': '#e2e8f0'}, 'x': 0.5, 'xanchor': 'center'},
                plot_bgcolor='#1a202c',
                paper_bgcolor='#2d3748',
                xaxis={'gridcolor': '#4a5568', 'color': '#e2e8f0', 'tickangle': -25},
                yaxis={'gridcolor': '#4a5568', 'color': '#e2e8f0'},
                margin=dict(l=50, r=30, t=50, b=100),
                hoverlabel=dict(bgcolor='white', font_size=14, font_color='#1a202c')
            )
            
            return html.Div([
                html.H5("Indicadores de Gestión", style={'color': '#ffffff', 'fontWeight': '600'}, className="mb-3"),
                dbc.Row([
                    dbc.Col(kpi1, xs=12, md=3),
                    dbc.Col(kpi2, xs=12, md=3),
                    dbc.Col(kpi3, xs=12, md=3),
                    dbc.Col(kpi4, xs=12, md=3),
                ], className="mb-4"),
                dbc.Row([
                    dbc.Col(dcc.Graph(figure=fig_estancia, config={'displayModeBar': False}), xs=12, md=6),
                    dbc.Col(dcc.Graph(figure=fig_giro, config={'displayModeBar': False}), xs=12, md=6),
                ]),
                html.H5("Indicadores de Calidad e Infecciones", style={'color': '#ffffff', 'fontWeight': '600', 'marginTop': '50px', 'marginBottom': '30px', 'textAlign': 'center'}),
                dbc.Row([
                    dbc.Col(dcc.Graph(figure=fig_bacteriemia, config={'displayModeBar': False}), xs=12, md=6),
                    dbc.Col(dcc.Graph(figure=fig_nav, config={'displayModeBar': False}), xs=12, md=6),
                ]),
                dbc.Row([
                    dbc.Col(dcc.Graph(figure=fig_istu, config={'displayModeBar': False}), xs=12, md=6),
                    dbc.Col(dcc.Graph(figure=fig_lesiones, config={'displayModeBar': False}), xs=12, md=6),
                ], className="mt-3"),
                dbc.Row([
                    dbc.Col(dcc.Graph(figure=fig_autoextubaciones, config={'displayModeBar': False}), xs=12, md=6),
                    dbc.Col(dcc.Graph(figure=fig_caidas, config={'displayModeBar': False}), xs=12, md=6),
                ], className="mt-3"),
                html.H5("Gérmenes Aislados y Eventos Adversos", style={'color': '#ffffff', 'fontWeight': '600', 'marginTop': '50px', 'marginBottom': '30px', 'textAlign': 'center'}),
                dbc.Row([
                    dbc.Col(dcc.Graph(figure=fig_germenes, config={'displayModeBar': False}), xs=12, md=6),
                    dbc.Col(dcc.Graph(figure=fig_eventos, config={'displayModeBar': False}), xs=12, md=6),
                ], className="mt-3")
            ])
        except Exception as e:
            return html.Div([html.H3(f"Error: {str(e)}")])
    
    return html.Div([html.H5("Selecciona una pestaña")])


# Callback para aplicar filtros en Resumen General
@app.callback(
    Output('resumen-content', 'children'),
    [Input('btn-aplicar-filtros', 'n_clicks')],
    [dash.dependencies.State('filtro-mes-inicio', 'value'),
     dash.dependencies.State('filtro-mes-fin', 'value'),
     dash.dependencies.State('filtro-anio', 'value')],
    prevent_initial_call=False
)
def actualizar_resumen(n_clicks, mes_inicio, mes_fin, anio):
    try:
        # Determinar rango de meses
        if mes_inicio and mes_fin:
            idx_inicio = meses_disponibles.index(mes_inicio) if mes_inicio in meses_disponibles else 0
            idx_fin = meses_disponibles.index(mes_fin) if mes_fin in meses_disponibles else len(meses_disponibles) - 1
            
            # Asegurar que inicio <= fin
            if idx_inicio > idx_fin:
                idx_inicio, idx_fin = idx_fin, idx_inicio
            
            meses_filtrados = meses_disponibles[idx_inicio:idx_fin + 1]
        else:
            meses_filtrados = meses_disponibles
        
        # Texto del rango para subtítulos
        rango_texto = f"{meses_filtrados[0]} - {meses_filtrados[-1]}" if len(meses_filtrados) > 1 else meses_filtrados[0]
        
        # Obtener columna de suma
        suma_col = [c for c in df.columns if str(c).strip() in ['∑', 'SUMA']]
        suma_col = suma_col[0] if suma_col else '∑'
        
        # Calcular KPIs con meses filtrados
        datos_ocupacional = safe_get_row('% OCUPACIONAL', meses_filtrados)
        promedio_ocupacional = datos_ocupacional[datos_ocupacional > 0].mean() if len(datos_ocupacional[datos_ocupacional > 0]) > 0 else 0
        
        # Sumar ingresos, fallecidos y egresos de los meses filtrados
        datos_ingresos_kpi = safe_get_row('NUMERO DE INGRESOS DEL MES', meses_filtrados)
        total_ingresos = datos_ingresos_kpi.sum() if isinstance(datos_ingresos_kpi, pd.Series) else datos_ingresos_kpi
        
        datos_fallecidos = safe_get_row('FALLECIDOS', meses_filtrados)
        total_fallecidos = datos_fallecidos.sum() if isinstance(datos_fallecidos, pd.Series) else datos_fallecidos
        
        datos_egresos = safe_get_row('NÚMERO DE EGRESOS DEL MES', meses_filtrados)
        total_egresos = datos_egresos.sum() if isinstance(datos_egresos, pd.Series) else datos_egresos
        
        # Calcular promedio de estancia
        datos_estancia = safe_get_row('PROMEDIO ESTANCIA', meses_filtrados)
        promedio_estancia = datos_estancia[datos_estancia > 0].mean() if len(datos_estancia[datos_estancia > 0]) > 0 else 0
        
        # Calcular giro de cama
        datos_giro = safe_get_row('GIRO CAMA', meses_filtrados)
        giro_cama = datos_giro[datos_giro > 0].mean() if len(datos_giro[datos_giro > 0]) > 0 else 0
        
        # Validar valores
        if pd.isna(promedio_ocupacional): promedio_ocupacional = 0
        if pd.isna(total_ingresos): total_ingresos = 0
        if pd.isna(total_fallecidos): total_fallecidos = 0
        if pd.isna(total_egresos) or total_egresos == 0: total_egresos = 1
        if pd.isna(promedio_estancia): promedio_estancia = 0
        if pd.isna(giro_cama): giro_cama = 0
        
        mortalidad = (total_fallecidos / total_egresos * 100)
        
        # Crear KPIs - Primera fila
        kpi_ingresos = create_kpi_card("Total Ingresos", f"{int(total_ingresos)}", rango_texto, "#3498db", "📥")
        kpi1 = create_kpi_card("% Ocupación Promedio", f"{promedio_ocupacional:.2f}%", rango_texto, "#2c3e50", "🏥")
        kpi2 = create_kpi_card("% Mortalidad General", f"{mortalidad:.2f}%", f"Total {rango_texto}", "#e74c3c", "💔")
        kpi3 = create_kpi_card("Estancia Promedio", f"{promedio_estancia:.2f}", "Días", "#27ae60", "⏱️")
        kpi4 = create_kpi_card("Giro de Cama", f"{giro_cama:.2f}", f"Promedio {rango_texto}", "#f39c12", "🔄")
        
        # EPS principales: Famisanar, Compensar, Salud Total
        total_famisanar = safe_get_row('FAMISANAR', suma_col, silent=True)
        total_compensar = safe_get_row('COMPENSAR', suma_col, silent=True)
        total_salud_total = safe_get_row('SALUD TOTAL', suma_col, silent=True)
        
        if pd.isna(total_famisanar): total_famisanar = 0
        if pd.isna(total_compensar): total_compensar = 0
        if pd.isna(total_salud_total): total_salud_total = 0
        
        # Calcular porcentajes
        total_ingresos_pct = total_ingresos if total_ingresos > 0 else 1
        pct_famisanar = (total_famisanar / total_ingresos_pct * 100)
        pct_compensar = (total_compensar / total_ingresos_pct * 100)
        pct_salud_total = (total_salud_total / total_ingresos_pct * 100)
        
        kpi_eps1 = create_kpi_card("Famisanar", f"{int(total_famisanar)}", f"{pct_famisanar:.1f}% del total", "#27ae60", "🏢")
        kpi_eps2 = create_kpi_card("Compensar", f"{int(total_compensar)}", f"{pct_compensar:.1f}% del total", "#f39c12", "🏢")
        kpi_eps3 = create_kpi_card("Salud Total", f"{int(total_salud_total)}", f"{pct_salud_total:.1f}% del total", "#e74c3c", "🏢")
        
        # Crear gráficos con meses filtrados
        datos_ocup = safe_get_row('% OCUPACIONAL', meses_filtrados)
        fig_ocupacional = px.line(x=meses_filtrados, y=datos_ocup, title='% Ocupacional Mensual', markers=True)
        fig_ocupacional.update_traces(
            line_color='#16a085',
            line_width=3,
            marker=dict(size=10, line=dict(width=2, color='#2d3748')),
            hovertemplate='<b>%{x}</b><br>Valor: %{y:.2f}<extra></extra>'
        )
        fig_ocupacional.update_layout(
            showlegend=False,
            plot_bgcolor='#1a202c',
            paper_bgcolor='#2d3748',
            title={'font': {'size': 15, 'color': '#e2e8f0'}, 'x': 0.5, 'xanchor': 'center'},
            xaxis={'gridcolor': '#4a5568', 'showline': True, 'linecolor': '#4a5568', 'color': '#a0aec0'},
            yaxis={'gridcolor': '#4a5568', 'showline': True, 'linecolor': '#4a5568', 'color': '#a0aec0'},
            margin=dict(l=50, r=30, t=50, b=40),
            hovermode='closest',
            hoverlabel=dict(bgcolor='white', font_size=14, font_color='#1a202c')
        )
        
        # Gráfico de ingresos vs egresos
        datos_ingresos = safe_get_row('NUMERO DE INGRESOS DEL MES', meses_filtrados)
        datos_egresos_chart = safe_get_row('NÚMERO DE EGRESOS DEL MES', meses_filtrados)
        
        fig_ingresos_egresos = go.Figure()
        fig_ingresos_egresos.add_trace(go.Bar(
            name='Ingresos',
            x=meses_filtrados,
            y=datos_ingresos,
            marker_color='#2c3e50',
            hovertemplate='<b>%{x}</b><br>Ingresos: %{y:.0f}<extra></extra>'
        ))
        fig_ingresos_egresos.add_trace(go.Bar(
            name='Egresos',
            x=meses_filtrados,
            y=datos_egresos_chart,
            marker_color='#27ae60',
            hovertemplate='<b>%{x}</b><br>Egresos: %{y:.0f}<extra></extra>'
        ))
        fig_ingresos_egresos.update_layout(
            title={'text': 'Ingresos vs. Egresos Mensuales', 'font': {'size': 15, 'color': '#e2e8f0'}, 'x': 0.5, 'xanchor': 'center'},
            barmode='group',
            plot_bgcolor='#1a202c',
            paper_bgcolor='#2d3748',
            xaxis={'gridcolor': '#4a5568', 'color': '#e2e8f0'},
            yaxis={'gridcolor': '#4a5568', 'color': '#e2e8f0', 'tickformat': 'd'},
            margin=dict(l=50, r=30, t=90, b=40),
            legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5, bgcolor='rgba(26,32,44,0.95)', font={'color': '#ffffff'}),
            hovermode='closest',
            hoverlabel=dict(bgcolor='white', font_size=14, font_color='#1a202c')
        )
        
        # Gráficos de Morbilidad
        fig_morbilidad_total = go.Figure()
        if not df_morbilidad.empty:
            fig_morbilidad_total = go.Figure(data=[
                go.Bar(
                    x=df_morbilidad.index,
                    y=df_morbilidad['Total'],
                    marker_color=['#e74c3c', '#3498db', '#2ecc71', '#f39c12', '#9b59b6', 
                                  '#1abc9c', '#e67e22', '#34495e', '#16a085', '#c0392b'],
                    hovertemplate='<b>%{x}</b><br>Total: %{y}<extra></extra>'
                )
            ])
            fig_morbilidad_total.update_layout(
                title={'text': 'Morbilidad por Diagnóstico (Total Acumulado)', 'font': {'size': 15, 'color': '#e2e8f0'}, 'x': 0.5, 'xanchor': 'center'},
                plot_bgcolor='#1a202c',
                paper_bgcolor='#2d3748',
                xaxis={'gridcolor': '#4a5568', 'color': '#e2e8f0', 'tickangle': -45},
                yaxis={'gridcolor': '#4a5568', 'color': '#e2e8f0'},
                margin=dict(l=50, r=30, t=50, b=80),
                hoverlabel=dict(bgcolor='white', font_size=14, font_color='#1a202c')
            )
        
        fig_morbilidad_genero = go.Figure()
        if not df_morbilidad.empty:
            fig_morbilidad_genero = go.Figure(data=[
                go.Bar(name='Hombres', x=df_morbilidad.index, y=df_morbilidad['Hombre'], 
                       marker_color='#3498db', hovertemplate='<b>%{x}</b><br>Hombres: %{y}<extra></extra>'),
                go.Bar(name='Mujeres', x=df_morbilidad.index, y=df_morbilidad['Mujer'], 
                       marker_color='#e91e63', hovertemplate='<b>%{x}</b><br>Mujeres: %{y}<extra></extra>')
            ])
            fig_morbilidad_genero.update_layout(
                title={'text': 'Morbilidad por Género', 'font': {'size': 15, 'color': '#e2e8f0'}, 'x': 0.5, 'xanchor': 'center'},
                barmode='group',
                plot_bgcolor='#1a202c',
                paper_bgcolor='#2d3748',
                xaxis={'gridcolor': '#4a5568', 'color': '#e2e8f0', 'tickangle': -45},
                yaxis={'gridcolor': '#4a5568', 'color': '#e2e8f0'},
                margin=dict(l=50, r=30, t=50, b=100),
                legend=dict(orientation="h", yanchor="top", y=-0.38, xanchor="center", x=0.5, 
                           bgcolor='rgba(26,32,44,0.95)', font={'color': '#ffffff'}),
                hoverlabel=dict(bgcolor='white', font_size=14, font_color='#1a202c')
            )
        
        return html.Div([
            dbc.Row([
                dbc.Col(kpi_ingresos, xs=12, sm=6, md=4, lg=True, className="mb-3"),
                dbc.Col(kpi1, xs=12, sm=6, md=4, lg=True, className="mb-3"),
                dbc.Col(kpi2, xs=12, sm=6, md=4, lg=True, className="mb-3"),
                dbc.Col(kpi3, xs=12, sm=6, md=4, lg=True, className="mb-3"),
                dbc.Col(kpi4, xs=12, sm=6, md=4, lg=True, className="mb-3"),
            ], className="mb-3 g-3"),
            html.H6("Principales EPS", style={'color': '#e2e8f0', 'fontWeight': '600', 'marginBottom': '15px', 'marginTop': '10px'}),
            dbc.Row([
                dbc.Col(kpi_eps1, xs=12, sm=6, md=4, className="mb-3"),
                dbc.Col(kpi_eps2, xs=12, sm=6, md=4, className="mb-3"),
                dbc.Col(kpi_eps3, xs=12, sm=6, md=4, className="mb-3"),
            ], className="mb-4 g-3"),
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_ocupacional, config={'displayModeBar': False}), xs=12, md=6),
                dbc.Col(dcc.Graph(figure=fig_ingresos_egresos, config={'displayModeBar': False}), xs=12, md=6),
            ]),
            html.H6("Morbilidad", style={'color': '#ffffff', 'fontWeight': '600', 'marginTop': '40px', 'marginBottom': '20px', 'textAlign': 'center'}),
            dbc.Row([
                dbc.Col(dcc.Graph(figure=fig_morbilidad_total, config={'displayModeBar': False}), xs=12, md=6),
                dbc.Col(dcc.Graph(figure=fig_morbilidad_genero, config={'displayModeBar': False}), xs=12, md=6),
            ])
        ])
    except Exception as e:
        return html.Div([html.H3(f"Error: {str(e)}")])


if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=8050)


