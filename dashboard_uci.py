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

# --- CARGA DE DATOS DE FALLECIDOS (MORTALIDAD) - Edad, Sexo y Días de Estancia ---
df_fallecidos = pd.DataFrame()
try:
    import unicodedata
    df_mort = pd.read_excel(EXCEL_PATH, sheet_name='MORTALIDAD', header=None, usecols='A:F')
    
    meses_separadores = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 
                         'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE']
    
    meses_map = {'ENERO': 'Ene', 'FEBRERO': 'Feb', 'MARZO': 'Mar', 'ABRIL': 'Abr', 
                 'MAYO': 'May', 'JUNIO': 'Jun', 'JULIO': 'Jul', 'AGOSTO': 'Agos',
                 'SEPTIEMBRE': 'Sep', 'OCTUBRE': 'Oct', 'NOVIEMBRE': 'Nov', 'DICIEMBRE': 'Dic'}
    
    nombres_femeninos = ['maria', 'ana', 'luz', 'rosa', 'carmen', 'martha', 'marta', 'gloria', 
                         'patricia', 'sandra', 'claudia', 'diana', 'adriana', 'carolina', 'andrea',
                         'monica', 'laura', 'paula', 'angela', 'beatriz', 'cecilia', 'clara',
                         'cristina', 'elena', 'esperanza', 'eva', 'flor', 'graciela', 'hilda',
                         'irma', 'isabel', 'julia', 'liliana', 'lucia', 'luisa', 'margarita',
                         'marina', 'mercedes', 'nancy', 'nelly', 'nidia', 'olga', 'oliva',
                         'pilar', 'rocio', 'ruth', 'silvia', 'sonia', 'stella', 'teresa',
                         'victoria', 'virginia', 'yolanda', 'blanca', 'barbara', 'nesly',
                         'claribel', 'evelia', 'zulema', 'angelica', 'milena', 'ines']
    
    def normalizar_texto(s):
        return ''.join(c for c in unicodedata.normalize('NFD', s.lower()) if unicodedata.category(c) != 'Mn')
    
    def es_nombre_femenino(nombre_completo):
        nombre_norm = normalizar_texto(nombre_completo)
        primer_nombre = nombre_norm.split()[0] if nombre_norm.split() else ''
        for nf in nombres_femeninos:
            if nf == primer_nombre or primer_nombre.startswith(nf):
                return True
        return False
    
    def extraer_dias(valor):
        if pd.isna(valor):
            return None
        if isinstance(valor, (int, float)):
            return int(valor)
        # Manejar casos como "42 horas", "36 horas"
        val_str = str(valor).lower().strip()
        if 'hora' in val_str:
            try:
                horas = int(''.join(filter(str.isdigit, val_str)))
                return max(1, horas // 24)  # Convertir a días, mínimo 1
            except:
                return 1
        try:
            return int(float(val_str))
        except:
            return None
    
    fallecidos_data = []
    mes_actual = None
    for i, row in df_mort.iterrows():
        nombre = str(row[0]).strip() if pd.notna(row[0]) else ''
        
        # Detectar separador de mes
        nombre_upper = nombre.upper().replace(' ', '')
        if nombre_upper in [m.upper() for m in meses_separadores]:
            mes_actual = meses_map.get(nombre_upper, None)
            continue
        
        edad = row[1]
        dias_estancia = extraer_dias(row[2])
        
        if not nombre or nombre.upper() == 'NAN':
            continue
        if not pd.notna(edad) or not isinstance(edad, (int, float)):
            continue
        
        sexo = 'Mujer' if es_nombre_femenino(nombre) else 'Hombre'
        fallecidos_data.append({
            'Nombre': nombre, 
            'Edad': int(edad), 
            'Sexo': sexo, 
            'Dias_Estancia': dias_estancia,
            'Mes': mes_actual
        })
    
    df_fallecidos = pd.DataFrame(fallecidos_data)
    print(f"✅ Datos de fallecidos cargados: {len(df_fallecidos)} pacientes")
except Exception as e:
    print(f"⚠️ Error cargando fallecidos: {e}")

# --- CARGA DE OTROS EVENTOS ADVERSOS (celdas azules en columna E) ---
df_eventos_adversos = pd.DataFrame()
try:
    wb = load_workbook(EXCEL_PATH)
    ws = wb['EVENTOS ADVERSOS']
    
    eventos_azules = []
    for i in range(1, 150):
        cell = ws.cell(i, 5)
        if cell.value:
            fill_color = cell.fill.fgColor.rgb if cell.fill.fgColor else None
            if fill_color and fill_color == 'FF00B0F0':
                evento = str(cell.value).strip()
                eventos_azules.append(evento)
    
    # Crear DataFrame con cada evento individual (sin agrupar)
    df_eventos_adversos = pd.DataFrame({'Evento': eventos_azules, 'Cantidad': [1] * len(eventos_azules)})
    print(f"✅ Datos de eventos adversos cargados: {len(eventos_azules)} eventos")
except Exception as e:
    print(f"⚠️ Error cargando eventos adversos: {e}")

# Meses disponibles
meses_disponibles = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Agos', 'Sep', 'Oct']

meses_completos = {
    'Ene': 'Enero', 'Feb': 'Febrero', 'Mar': 'Marzo', 'Abr': 'Abril',
    'May': 'Mayo', 'Jun': 'Junio', 'Jul': 'Julio', 'Agos': 'Agosto',
    'Sep': 'Septiembre', 'Oct': 'Octubre', 'Nov': 'Noviembre', 'Dic': 'Diciembre'
}

# Paleta de colores - Tema Oscuro/Negro
COLOR_PALETTE = {
    'primary': '#3498db',
    'secondary': '#27ae60',
    'accent': '#16a085',
    'success': '#27ae60',
    'danger': '#e74c3c',
    'warning': '#f39c12',
    'info': '#34495e',
    'light': '#e2e8f0',
    'dark': '#0a0a0a',
    'muted': '#a0aec0',
    'bg_primary': '#000000',
    'bg_secondary': '#111111',
    'bg_card': '#1a1a1a',
    'text_primary': '#ffffff',
    'text_secondary': '#e2e8f0',
    'border': '#333333',
}

def capitalizar_texto(texto):
    if not texto or not isinstance(texto, str):
        return texto
    return texto[0].upper() + texto[1:].lower() if len(texto) > 1 else texto.upper()

def safe_get_row(row_name, columns, silent=False):
    try:
        matching_rows = [idx for idx in df.index if row_name.lower() in str(idx).lower()]
        if matching_rows:
            actual_row = matching_rows[0]
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
        title={'font': {'size': 16, 'color': '#ffffff'}, 'x': 0.5, 'xanchor': 'center'},
        xaxis={'gridcolor': '#4a5568', 'showline': True, 'linecolor': '#4a5568', 'color': '#e2e8f0'},
        yaxis={'gridcolor': '#4a5568', 'showline': True, 'linecolor': '#4a5568', 'color': '#e2e8f0', 'tickformat': 'd'},
        margin=dict(l=50, r=30, t=60, b=40),
        hoverlabel=dict(bgcolor='white', font_size=14, font_color='#000000', bordercolor='#333333')
    )
    return fig

def create_line_chart(row_name, title, color=None, silent=False, decimals=2):
    data = safe_get_row(row_name, meses_disponibles, silent=silent)
    fig = px.line(x=meses_disponibles, y=data, title=title, markers=True)
    
    if decimals == 0:
        hover_format = '<b>%{x}</b><br>Valor: %{y:.0f}<extra></extra>'
    else:
        hover_format = f'<b>%{{x}}</b><br>Valor: %{{y:.{decimals}f}}<extra></extra>'
    
    fig.update_traces(
        line_color=color or COLOR_PALETTE['secondary'],
        line_width=3,
        marker=dict(size=8, line=dict(width=2, color='#1a1a1a')),
        hovertemplate=hover_format
    )
    
    yaxis_config = {'gridcolor': '#4a5568', 'showline': True, 'linecolor': '#4a5568', 'color': '#e2e8f0'}
    if isinstance(data, pd.Series):
        max_val = data.max() if not data.empty else 0
        if pd.notna(max_val) and max_val <= 5:
            yaxis_config['dtick'] = 1
    
    fig.update_layout(
        showlegend=False,
        plot_bgcolor='#1a202c',
        paper_bgcolor='#2d3748',
        title={'font': {'size': 16, 'color': '#ffffff'}, 'x': 0.5, 'xanchor': 'center'},
        xaxis={'gridcolor': '#4a5568', 'showline': True, 'linecolor': '#4a5568', 'color': '#e2e8f0'},
        yaxis=yaxis_config,
        margin=dict(l=50, r=30, t=60, b=40),
        hoverlabel=dict(bgcolor='white', font_size=14, font_color='#000000', bordercolor='#333333')
    )
    return fig

def create_grouped_bar_chart(row_names, labels, title, decimals=0):
    fig = go.Figure()
    colors = [COLOR_PALETTE['primary'], COLOR_PALETTE['secondary'], COLOR_PALETTE['accent'], COLOR_PALETTE['success']]
    
    for i, (row_name, label) in enumerate(zip(row_names, labels)):
        data = safe_get_row(row_name, meses_disponibles)
        
        if decimals == 0:
            hover_template = f'<b>%{{x}}</b><br>{label}: %{{y:.0f}}<extra></extra>'
        else:
            hover_template = f'<b>%{{x}}</b><br>{label}: %{{y:.{decimals}f}}<extra></extra>'
        
        fig.add_trace(go.Bar(
            name=label, x=meses_disponibles, y=data,
            marker_color=colors[i % len(colors)],
            hovertemplate=hover_template
        ))
    
    fig.update_layout(
        title={'text': title, 'font': {'size': 16, 'color': '#ffffff'}, 'x': 0.5, 'xanchor': 'center'},
        barmode='group',
        plot_bgcolor='#1a202c',
        paper_bgcolor='#2d3748',
        xaxis={'gridcolor': '#4a5568', 'color': '#e2e8f0'},
        yaxis={'gridcolor': '#4a5568', 'color': '#e2e8f0', 'tickformat': 'd'},
        margin=dict(l=50, r=30, t=90, b=40),
        legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5, 
                   bgcolor='rgba(0,0,0,0.8)', font={'color': '#ffffff'}),
        hoverlabel=dict(bgcolor='white', font_size=13, font_color='#000000', bordercolor='#333333')
    )
    return fig

def create_kpi_card(title, value, subtitle="", color=None, icon="📊"):
    return dbc.Card(
        dbc.CardBody([
            html.Div([
                html.Span(icon, style={'fontSize': '28px', 'marginRight': '12px', 'flexShrink': '0'}),
                html.Div([
                    html.H6(capitalizar_texto(title), style={
                        'color': '#a0aec0', 'fontSize': '13px', 'marginBottom': '8px', 
                        'fontWeight': '600', 'textTransform': 'uppercase', 'letterSpacing': '0.5px'
                    }),
                    html.H2(value, style={
                        'color': '#ffffff', 'fontWeight': '700', 
                        'marginBottom': '5px', 'fontSize': '28px', 'lineHeight': '1',
                        'textShadow': f'0 0 20px {color or COLOR_PALETTE["primary"]}'
                    }),
                    html.P(capitalizar_texto(subtitle), style={
                        'color': '#a0aec0', 'fontSize': '12px', 'margin': '0', 'fontWeight': '500'
                    })
                ], style={'flex': '1', 'minWidth': '0', 'overflow': 'hidden'})
            ], style={'display': 'flex', 'alignItems': 'center'})
        ], style={'padding': '16px'}),
        className="kpi-card h-100",
        style={
            'border': 'none', 'borderLeft': f'4px solid {color or COLOR_PALETTE["primary"]}',
            'borderRadius': '12px', 'boxShadow': '0 8px 32px rgba(0,0,0,0.3)',
            'background': 'linear-gradient(145deg, #132238 0%, #1a365d 100%)', 'transition': 'all 0.3s ease', 'height': '100%'
        }
    )

app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP], suppress_callback_exceptions=True)
server = app.server


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
            @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
            
            body {
                font-family: 'Inter', sans-serif !important;
                background: linear-gradient(135deg, #0c1929 0%, #1a365d 25%, #1e3a5f 50%, #0f2744 75%, #0a1628 100%);
                color: #e2e8f0;
                min-height: 100vh;
            }
            
            @keyframes fadeIn {
                from { opacity: 0; transform: translateY(20px); }
                to { opacity: 1; transform: translateY(0); }
            }
            
            .kpi-card {
                transition: all 0.3s ease;
                border-radius: 12px !important;
                box-shadow: 0 8px 32px rgba(0,0,0,0.4), 0 0 40px rgba(26,54,93,0.2) !important;
                backdrop-filter: blur(10px);
            }
            
            .kpi-card:hover {
                transform: translateY(-5px);
                box-shadow: 0 12px 40px rgba(0,0,0,0.5), 0 0 60px rgba(52,152,219,0.15) !important;
            }
            
            .fade-in { animation: fadeIn 0.6s ease-out; }
            
            .nav-tabs {
                border-bottom: none !important;
            }
            
            .nav-tabs .nav-link {
                color: #a0aec0;
                border: none;
                border-bottom: 3px solid transparent;
                transition: all 0.3s ease;
                background: transparent;
                font-weight: 500;
                padding: 12px 16px;
                margin-right: 8px;
                border-radius: 8px 8px 0 0;
            }
            
            .nav-tabs .nav-link:hover {
                border-bottom: 3px solid #3498db;
                color: #3498db;
                background-color: rgba(52, 152, 219, 0.1);
            }
            
            .nav-tabs .nav-link.active {
                color: #3498db;
                border-bottom: 3px solid #3498db;
                background: rgba(52, 152, 219, 0.15);
                font-weight: 600;
            }
            
            .tab-content {
                background: linear-gradient(145deg, #0f2744 0%, #132238 100%);
                border-radius: 0 12px 12px 12px;
                padding: 25px;
                box-shadow: 0 8px 32px rgba(0,0,0,0.3);
            }
            
            .Select-control { background-color: rgba(19,34,56,0.9) !important; border-color: #1e3a5f !important; border-radius: 8px !important; }
            .Select-value-label, .Select-placeholder, .Select-value { color: #ffffff !important; }
            .Select-menu-outer { background-color: #132238 !important; border-color: #1e3a5f !important; }
            .Select-option { background-color: #132238 !important; color: #ffffff !important; }
            .Select-option:hover, .Select-option.is-focused { background-color: #3498db !important; color: #ffffff !important; }
            .Select-option.is-selected { background-color: #3498db !important; color: #ffffff !important; }
            
            .dashboard-footer {
                background-color: transparent;
                padding: 0 0 20px 0;
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
        # Header
        html.Div([
            html.Div([
                html.Div([
                    html.Img(src='/assets/Logo.png', style={'height': '90px', 'marginRight': '25px', 'filter': 'drop-shadow(0 4px 8px rgba(0,0,0,0.4))'}),
                    html.Div([
                        html.H1("DASHBOARD UCI ADULTOS 2025", style={
                            'color': '#ffffff', 'fontWeight': '700', 'marginBottom': '5px',
                            'fontSize': '28px', 'letterSpacing': '0.5px', 'margin': '0',
                            'textShadow': '0 0 30px rgba(52,152,219,0.5)'
                        }),
                        html.P("Sistema de Monitoreo y Análisis de Indicadores", style={
                            'color': '#a0aec0', 'fontSize': '14px', 'marginBottom': '0', 'fontWeight': '400'
                        })
                    ])
                ], style={'display': 'flex', 'alignItems': 'center', 'marginBottom': '15px'})
            ]),
            html.Div([
                dbc.Button("📤 Exportar a Excel", color="success", size="sm", className="me-2", style={'borderRadius': '6px', 'fontWeight': '500'}),
                dbc.Button("📄 Exportar a CSV", color="primary", size="sm", style={'borderRadius': '6px', 'fontWeight': '500'}),
            ], style={'marginBottom': '20px'})
        ], className="fade-in"),
        
        # Tabs
        dbc.Tabs(id="tabs", active_tab='tab-resumen', children=[
            dbc.Tab(label='✅ Resumen general', tab_id='tab-resumen'),
            dbc.Tab(label='📦 Gestión y calidad', tab_id='tab-gestion'),
            dbc.Tab(label='🏥 Especialidades y FPS', tab_id='tab-especialidades'),
            dbc.Tab(label='💀 Mortalidad', tab_id='tab-mortalidad'),
        ], style={'marginBottom': '0', 'borderBottom': 'none', 'backgroundColor': 'rgba(19,34,56,0.8)', 'borderRadius': '12px 12px 0 0', 'padding': '0 10px'}),
        
        html.Div(id='tab-content', className="tab-content"),
        
        # Footer
        html.Div([
            html.Hr(style={'borderColor': '#333333', 'margin': '30px 0 20px 0'}),
            html.Div([
                html.Div([
                    html.P("Hospital Infantil Universitario de San José", style={'color': '#e2e8f0', 'fontWeight': '600', 'margin': '0', 'fontSize': '14px'}),
                    html.P("Unidad de Cuidados Intensivos - Adultos", style={'color': '#a0aec0', 'margin': '0', 'fontSize': '12px'})
                ]),
                html.Div([
                    html.P("© 2025 - Dashboard de Indicadores UCI", style={'color': '#718096', 'margin': '0', 'fontSize': '12px'})
                ], style={'textAlign': 'right'})
            ], style={'display': 'flex', 'justifyContent': 'space-between', 'alignItems': 'center', 'flexWrap': 'wrap', 'gap': '15px'})
        ], className="dashboard-footer")
        
    ], fluid=True, style={'maxWidth': '1400px', 'padding': '30px 20px'})
], style={'background': 'linear-gradient(135deg, #0c1929 0%, #1a365d 25%, #1e3a5f 50%, #0f2744 75%, #0a1628 100%)', 'minHeight': '100vh', 'fontFamily': "'Inter', sans-serif"})

@app.callback(Output('tab-content', 'children'), Input('tabs', 'active_tab'))
def render_tab_content(tab):
    if tab == 'tab-resumen':
        try:
            if len(df.columns) == 0:
                return html.Div([html.H3("⚠️ No hay datos", style={'color': '#ffffff'})])
            
            filtros = html.Div([
                dbc.Row([
                    dbc.Col([
                        html.Label("📅 Meses (selecciona hasta 3):", style={'color': '#e2e8f0', 'fontWeight': '600', 'marginBottom': '8px'}),
                        dcc.Dropdown(
                            id='filtro-mes-inicio',
                            options=[{'label': meses_completos[mes], 'value': mes} for mes in meses_disponibles],
                            value=['Agos', 'Sep', 'Oct'],
                            multi=True,
                            placeholder="Selecciona meses..."
                        )
                    ], xs=12, md=6),
                    dbc.Col([
                        html.Label(" ", style={'color': 'transparent', 'marginBottom': '8px', 'display': 'block'}),
                        dbc.Button("🔍 Aplicar Filtros", id='btn-aplicar-filtros', color="success", className="w-100",
                                  style={'borderRadius': '8px', 'fontWeight': '600', 'height': '38px'})
                    ], xs=12, md=3),
                    dbc.Col([
                        html.Label(" ", style={'color': 'transparent', 'marginBottom': '8px', 'display': 'block'}),
                        dbc.Button("📊 Todos los meses", id='btn-todos-meses', color="secondary", className="w-100",
                                  style={'borderRadius': '8px', 'fontWeight': '500', 'height': '38px'})
                    ], xs=12, md=3),
                ], className="mb-2"),
            ], style={'background': 'linear-gradient(145deg, #132238 0%, #1a365d 100%)', 'padding': '20px', 'borderRadius': '15px', 'marginBottom': '25px'})
            
            return html.Div([filtros, html.Div(id='resumen-content')])
        except Exception as e:
            return html.Div([html.H3(f"Error: {str(e)}", style={'color': '#e74c3c'})])

    elif tab == 'tab-mortalidad':
        try:
            suma_col = [c for c in df.columns if str(c).strip() in ['∑', 'SUMA']]
            suma_col = suma_col[0] if suma_col else '∑'
            
            total_fallecidos = safe_get_row('FALLECIDOS', suma_col)
            fallecidos_covid = safe_get_row('FALLECIDOS COVID-19', suma_col)
            total_egresos = safe_get_row('NÚMERO DE EGRESOS DEL MES', suma_col)
            
            if pd.isna(total_fallecidos): total_fallecidos = 0
            if pd.isna(fallecidos_covid): fallecidos_covid = 0
            if pd.isna(total_egresos) or total_egresos == 0: total_egresos = 1
            
            mortalidad_general = (total_fallecidos / total_egresos * 100)
            datos_mortalidad_trauma = safe_get_row('% mortalidad en trauma', meses_disponibles, silent=True)
            mortalidad_trauma = datos_mortalidad_trauma.sum() / 10 if isinstance(datos_mortalidad_trauma, pd.Series) else 0
            
            # Calcular estadísticas de días de estancia de fallecidos
            dias_estancia_validos = df_fallecidos['Dias_Estancia'].dropna() if not df_fallecidos.empty else pd.Series([0])
            if len(dias_estancia_validos) == 0:
                dias_estancia_validos = pd.Series([0])
            
            promedio_dias = dias_estancia_validos.mean()
            minimo_dias = dias_estancia_validos.min()
            maximo_dias = dias_estancia_validos.max()
            mediana_dias = dias_estancia_validos.median()
            
            # KPIs principales
            kpi1 = create_kpi_card("Total Fallecidos", f"{int(total_fallecidos)}", "Ene - Oct", "#e74c3c", "💀")
            kpi2 = create_kpi_card("% Mortalidad General", f"{mortalidad_general:.2f}%", "Ene - Oct", "#c0392b", "⚰️")
            kpi3 = create_kpi_card("Fallecidos COVID-19", f"{int(fallecidos_covid)}", "Ene - Oct", "#34495e", "🦠")
            kpi4 = create_kpi_card("% Mortalidad Trauma", f"{mortalidad_trauma:.2f}%", "Promedio", "#f39c12", "🚑")
            
            # KPIs de días de estancia fallecidos
            kpi_prom_dias = create_kpi_card("Promedio Días Estancia", f"{promedio_dias:.1f}", "Fallecidos", "#9b59b6", "📊")
            kpi_min_dias = create_kpi_card("Mínimo Días", f"{int(minimo_dias)}", "Fallecidos", "#27ae60", "⬇️")
            kpi_max_dias = create_kpi_card("Máximo Días", f"{int(maximo_dias)}", "Fallecidos", "#e67e22", "⬆️")
            kpi_mediana_dias = create_kpi_card("Mediana Días", f"{mediana_dias:.1f}", "Fallecidos", "#3498db", "📈")
            
            fig_mortalidad = create_line_chart('% DE MORATLIDAD GENERAL', '% Mortalidad General', color='#e74c3c', decimals=2)
            fig_trauma = create_grouped_bar_chart(
                ['Pacientes con diagnóstico de Trauma', 'Pacientes fallecidos por Trauma'],
                ['Pacientes Trauma', 'Fallecidos Trauma'], 'Trauma: Pacientes vs Fallecidos', decimals=0)

            # Gráfico Pirámide Poblacional de Edad y Sexo de Fallecidos
            fig_edad_sexo = go.Figure()
            if not df_fallecidos.empty:
                bins = [0, 40, 50, 60, 70, 80, 90, 100]
                labels = ['<40', '40-49', '50-59', '60-69', '70-79', '80-89', '90+']
                df_fallecidos['Rango_Edad'] = pd.cut(df_fallecidos['Edad'], bins=bins, labels=labels, right=False)
                edad_sexo = df_fallecidos.groupby(['Rango_Edad', 'Sexo']).size().unstack(fill_value=0)
                if 'Hombre' not in edad_sexo.columns: edad_sexo['Hombre'] = 0
                if 'Mujer' not in edad_sexo.columns: edad_sexo['Mujer'] = 0
                
                # Calcular totales y porcentajes
                total_fallecidos_chart = len(df_fallecidos)
                hombres_vals = edad_sexo['Hombre'].values
                mujeres_vals = edad_sexo['Mujer'].values
                hombres_pct = [(v / total_fallecidos_chart * 100) if total_fallecidos_chart > 0 else 0 for v in hombres_vals]
                mujeres_pct = [(v / total_fallecidos_chart * 100) if total_fallecidos_chart > 0 else 0 for v in mujeres_vals]
                
                # Pirámide: hombres a la izquierda (valores negativos), mujeres a la derecha
                fig_edad_sexo = go.Figure()
                fig_edad_sexo.add_trace(go.Bar(
                    name='Hombres', y=edad_sexo.index.astype(str), x=[-v for v in hombres_vals],
                    orientation='h', marker_color='#3498db',
                    text=[f'{int(v)} ({p:.1f}%)' for v, p in zip(hombres_vals, hombres_pct)],
                    textposition='inside', textfont=dict(color='white', size=11),
                    hovertemplate='<b>%{y}</b><br>Hombres: %{customdata[0]} (%{customdata[1]:.1f}%)<extra></extra>',
                    customdata=list(zip(hombres_vals, hombres_pct))
                ))
                fig_edad_sexo.add_trace(go.Bar(
                    name='Mujeres', y=edad_sexo.index.astype(str), x=mujeres_vals,
                    orientation='h', marker_color='#e91e63',
                    text=[f'{int(v)} ({p:.1f}%)' for v, p in zip(mujeres_vals, mujeres_pct)],
                    textposition='inside', textfont=dict(color='white', size=11),
                    hovertemplate='<b>%{y}</b><br>Mujeres: %{customdata[0]} (%{customdata[1]:.1f}%)<extra></extra>',
                    customdata=list(zip(mujeres_vals, mujeres_pct))
                ))
                
                max_val = max(max(hombres_vals), max(mujeres_vals)) if len(hombres_vals) > 0 else 5
                fig_edad_sexo.update_layout(
                    title={'text': 'Pirámide de Edad y Sexo - Fallecidos', 'font': {'size': 16, 'color': '#ffffff'}, 'x': 0.5},
                    barmode='relative', plot_bgcolor='#1a202c', paper_bgcolor='#2d3748',
                    xaxis={'title': '← Hombres | Mujeres →', 'gridcolor': '#4a5568', 'color': '#e2e8f0',
                           'range': [-(max_val + 2), max_val + 2], 'tickvals': list(range(-int(max_val)-1, int(max_val)+2)),
                           'ticktext': [str(abs(x)) for x in range(-int(max_val)-1, int(max_val)+2)]},
                    yaxis={'title': 'Rango de Edad', 'gridcolor': '#4a5568', 'color': '#e2e8f0'},
                    legend=dict(orientation="h", y=-0.15, x=0.5, xanchor="center", bgcolor='rgba(0,0,0,0.8)', font={'color': '#ffffff'}),
                    hoverlabel=dict(bgcolor='white', font_color='#000000'),
                    margin=dict(l=50, r=50, t=60, b=60)
                )
            
            return html.Div([
                dbc.Row([dbc.Col(kpi1, xs=12, md=3), dbc.Col(kpi2, xs=12, md=3), dbc.Col(kpi3, xs=12, md=3), dbc.Col(kpi4, xs=12, md=3)], className="mb-4"),
                html.H6("Días de Estancia - Fallecidos", style={'color': '#e2e8f0', 'fontWeight': '600', 'marginBottom': '15px', 'marginTop': '10px'}),
                dbc.Row([dbc.Col(kpi_prom_dias, xs=12, md=3), dbc.Col(kpi_min_dias, xs=12, md=3), dbc.Col(kpi_max_dias, xs=12, md=3), dbc.Col(kpi_mediana_dias, xs=12, md=3)], className="mb-4"),
                dbc.Row([
                    dbc.Col(xs=12, md=2),
                    dbc.Col(dcc.Graph(figure=fig_mortalidad), xs=12, md=8),
                    dbc.Col(xs=12, md=2)
                ], className="mb-4", justify="center"),
                dbc.Row([dbc.Col(dcc.Graph(figure=fig_trauma), xs=12, md=6), dbc.Col(dcc.Graph(figure=fig_edad_sexo), xs=12, md=6)], className="mt-3")
            ])
        except Exception as e:
            return html.Div([html.H3(f"Error: {str(e)}", style={'color': '#e74c3c'})])

    elif tab == 'tab-especialidades':
        try:
            suma_col = [c for c in df.columns if str(c).strip() in ['∑', 'SUMA']]
            suma_col = suma_col[0] if suma_col else '∑'
            
            total_medicina = safe_get_row('MEDICINA INTERNA', suma_col, silent=True)
            total_cirugia = safe_get_row('CIRUGIA GENERAL', suma_col, silent=True)
            total_neuro = safe_get_row('NEUROCIRUGIA', suma_col, silent=True)
            total_orto = safe_get_row('ORTOPEDIA', suma_col, silent=True)
            
            if pd.isna(total_medicina): total_medicina = 0
            if pd.isna(total_cirugia): total_cirugia = 0
            if pd.isna(total_neuro): total_neuro = 0
            if pd.isna(total_orto): total_orto = 0
            
            # Calcular total de especialidades y porcentajes
            total_especialidades = total_medicina + total_cirugia + total_neuro + total_orto
            if total_especialidades == 0: total_especialidades = 1
            
            pct_medicina = (total_medicina / total_especialidades * 100)
            pct_cirugia = (total_cirugia / total_especialidades * 100)
            pct_neuro = (total_neuro / total_especialidades * 100)
            pct_orto = (total_orto / total_especialidades * 100)
            
            kpi1 = create_kpi_card("Medicina Interna", f"{int(total_medicina)}", f"{pct_medicina:.1f}% del total", "#3498db", "⚕️")
            kpi2 = create_kpi_card("Cirugía General", f"{int(total_cirugia)}", f"{pct_cirugia:.1f}% del total", "#27ae60", "🔪")
            kpi3 = create_kpi_card("Neurocirugía", f"{int(total_neuro)}", f"{pct_neuro:.1f}% del total", "#16a085", "🧠")
            kpi4 = create_kpi_card("Ortopedia", f"{int(total_orto)}", f"{pct_orto:.1f}% del total", "#9b59b6", "🦴")
            
            fig_medicina = create_bar_chart('MEDICINA INTERNA', 'Medicina Interna', color='#3498db', silent=True, decimals=0)
            fig_cirugia = create_bar_chart('CIRUGIA GENERAL', 'Cirugía General', color='#27ae60', silent=True, decimals=0)
            fig_neuro = create_bar_chart('NEUROCIRUGIA', 'Neurocirugía', color='#16a085', silent=True, decimals=0)
            fig_orto = create_bar_chart('ORTOPEDIA', 'Ortopedia', color='#9b59b6', silent=True, decimals=0)
            
            return html.Div([
                html.H5("Principales Especialidades", style={'color': '#ffffff', 'fontWeight': '600', 'marginBottom': '20px'}),
                dbc.Row([dbc.Col(kpi1, xs=12, md=3), dbc.Col(kpi2, xs=12, md=3), dbc.Col(kpi3, xs=12, md=3), dbc.Col(kpi4, xs=12, md=3)], className="mb-4"),
                dbc.Row([dbc.Col(dcc.Graph(figure=fig_medicina), xs=12, md=6), dbc.Col(dcc.Graph(figure=fig_cirugia), xs=12, md=6)], className="mb-4"),
                dbc.Row([dbc.Col(dcc.Graph(figure=fig_neuro), xs=12, md=6), dbc.Col(dcc.Graph(figure=fig_orto), xs=12, md=6)])
            ])
        except Exception as e:
            return html.Div([html.H3(f"Error: {str(e)}", style={'color': '#e74c3c'})])

    elif tab == 'tab-gestion':
        try:
            suma_col = [c for c in df.columns if str(c).strip() in ['∑', 'SUMA']]
            suma_col = suma_col[0] if suma_col else '∑'
            
            datos_estancia = safe_get_row('PROMEDIO ESTANCIA', meses_disponibles)
            promedio_estancia = datos_estancia[datos_estancia > 0].mean() if isinstance(datos_estancia, pd.Series) and len(datos_estancia[datos_estancia > 0]) > 0 else 0
            datos_giro = safe_get_row('GIRO CAMA', meses_disponibles)
            giro_cama = datos_giro[datos_giro > 0].mean() if isinstance(datos_giro, pd.Series) and len(datos_giro[datos_giro > 0]) > 0 else 0
            total_dias = safe_get_row('TOTAL DIAS ESTANCIA', suma_col)
            capacidad = safe_get_row('CAPACIDAD INSTALADA: UCI', 'Ene')
            
            if pd.isna(promedio_estancia): promedio_estancia = 0
            if pd.isna(giro_cama): giro_cama = 0
            if pd.isna(total_dias): total_dias = 0
            if pd.isna(capacidad): capacidad = 0
            
            kpi1 = create_kpi_card("Estancia Promedio", f"{promedio_estancia:.2f}", "Días", "#2c3e50", "⏱️")
            kpi2 = create_kpi_card("Giro de Cama", f"{giro_cama:.2f}", "Promedio", "#27ae60", "🔄")
            kpi3 = create_kpi_card("Total Días Estancia", f"{int(total_dias)}", "Ene - Oct", "#16a085", "📅")
            kpi4 = create_kpi_card("Capacidad UCI", f"{int(capacidad)}", "Camas", "#f39c12", "🛏️")
            
            fig_estancia = create_line_chart('PROMEDIO ESTANCIA', 'Estancia Promedio (Días)', color='#16a085', decimals=2)
            fig_giro = create_line_chart('GIRO CAMA', 'Giro de Cama', color='#27ae60', decimals=2)
            
            # Gráfico Promedio de Días de VM (días ventilador / pacientes ventilados)
            datos_dias_vm = safe_get_row('DIAS VENTILADOR', meses_disponibles, silent=True)
            datos_pac_vm = safe_get_row('PACIENTES VENTILADOS', meses_disponibles, silent=True)
            
            # Calcular promedio días VM por mes
            promedio_vm = []
            for i in range(len(meses_disponibles)):
                dias = datos_dias_vm.iloc[i] if isinstance(datos_dias_vm, pd.Series) else 0
                pac = datos_pac_vm.iloc[i] if isinstance(datos_pac_vm, pd.Series) else 0
                if pd.notna(dias) and pd.notna(pac) and pac > 0:
                    promedio_vm.append(round(dias / pac, 2))
                else:
                    promedio_vm.append(0)
            
            fig_promedio_vm = go.Figure()
            fig_promedio_vm.add_trace(go.Scatter(
                x=meses_disponibles, y=promedio_vm, mode='lines+markers',
                line=dict(color='#9b59b6', width=3),
                marker=dict(size=10, line=dict(width=2, color='#1a1a1a')),
                hovertemplate='<b>%{x}</b><br>Promedio días VM: %{y:.2f}<extra></extra>'
            ))
            fig_promedio_vm.update_layout(
                title={'text': 'Promedio de Días de VM', 'font': {'size': 16, 'color': '#ffffff'}, 'x': 0.5},
                plot_bgcolor='#1a202c', paper_bgcolor='#2d3748',
                xaxis={'gridcolor': '#4a5568', 'showline': True, 'linecolor': '#4a5568', 'color': '#e2e8f0'},
                yaxis={'gridcolor': '#4a5568', 'showline': True, 'linecolor': '#4a5568', 'color': '#e2e8f0'},
                margin=dict(l=50, r=30, t=60, b=40),
                hoverlabel=dict(bgcolor='white', font_size=14, font_color='#000000')
            )
            
            fig_bacteriemia = create_line_chart('TASA DE BACTERIEMIA ASOC. A  C. CENTRAL', 'Tasa de ITS AC', color='#e74c3c', silent=True, decimals=2)
            fig_nav = create_line_chart('TASA NEUMONÍA ASOCIADA A VENTILADOR', 'Tasa NAV', color='#e74c3c', silent=True, decimals=2)
            fig_istu = create_line_chart('TASA INFECCIÓN ASOCIADA A C. URINARIO', 'Tasa de ISTU AC', color='#f39c12', silent=True, decimals=2)
            fig_lesiones = create_line_chart('TASA LESIONES POR PRESION', 'Tasa Lesiones por Presión', color='#9b59b6', silent=True, decimals=2)

            # Gráfico de Gérmenes Aislados (centrado)
            fig_germenes = go.Figure()
            if not df_germenes.empty:
                colors_germenes = ['#e74c3c', '#3498db', '#2ecc71', '#f39c12', '#9b59b6', '#1abc9c', '#e67e22', '#34495e', '#16a085', '#c0392b', '#27ae60', '#8e44ad', '#2980b9', '#d35400', '#7f8c8d']
                fig_germenes = go.Figure(data=[go.Bar(x=df_germenes['Germen'], y=df_germenes['Cantidad'],
                    marker_color=colors_germenes[:len(df_germenes)], hovertemplate='<b>%{x}</b><br>Casos: %{y}<extra></extra>')])
                fig_germenes.update_layout(title={'text': 'Gérmenes Aislados (IAAS)', 'font': {'size': 16, 'color': '#ffffff'}, 'x': 0.5},
                    plot_bgcolor='#1a202c', paper_bgcolor='#2d3748', xaxis={'gridcolor': '#4a5568', 'color': '#e2e8f0', 'tickangle': -45},
                    yaxis={'gridcolor': '#4a5568', 'color': '#e2e8f0', 'dtick': 1}, margin=dict(l=50, r=50, t=50, b=120),
                    height=450, hoverlabel=dict(bgcolor='white', font_color='#000000'))
            
            # Gráfico de Eventos Adversos - Barras horizontales expandido
            fig_eventos = go.Figure()
            if not df_eventos_adversos.empty:
                colors_eventos = ['#e74c3c', '#3498db', '#2ecc71', '#f39c12', '#9b59b6', '#1abc9c', '#e67e22', '#c0392b', '#27ae60', '#8e44ad', '#16a085', '#d35400', '#7f8c8d', '#2980b9', '#e91e63']
                
                # Invertir orden para que el primero quede arriba
                df_eventos_plot = df_eventos_adversos.iloc[::-1].reset_index(drop=True)
                
                fig_eventos = go.Figure(data=[go.Bar(
                    y=df_eventos_plot['Evento'],
                    x=df_eventos_plot['Cantidad'],
                    orientation='h',
                    marker_color=colors_eventos[:len(df_eventos_plot)][::-1],
                    text=df_eventos_plot['Cantidad'],
                    textposition='outside',
                    textfont=dict(color='#ffffff', size=12),
                    hovertemplate='<b>%{y}</b><br>Casos: %{x}<extra></extra>'
                )])
                
            fig_eventos.update_layout(
                title={'text': 'Otros Eventos Adversos', 'font': {'size': 16, 'color': '#ffffff'}, 'x': 0.5},
                plot_bgcolor='#1a202c', paper_bgcolor='#2d3748',
                xaxis={'gridcolor': '#4a5568', 'color': '#e2e8f0', 'dtick': 1, 'range': [0, 1.5]},
                yaxis={'gridcolor': 'rgba(74,85,104,0.3)', 'color': '#e2e8f0', 'tickfont': {'size': 11}},
                margin=dict(l=380, r=50, t=50, b=30),
                height=550,
                hoverlabel=dict(bgcolor='white', font_color='#000000', font_size=12),
                bargap=0.3
            )

            return html.Div([
                html.H5("Indicadores de Gestión", style={'color': '#ffffff', 'fontWeight': '600'}, className="mb-3"),
                dbc.Row([dbc.Col(kpi1, xs=12, md=3), dbc.Col(kpi2, xs=12, md=3), dbc.Col(kpi3, xs=12, md=3), dbc.Col(kpi4, xs=12, md=3)], className="mb-4"),
                dbc.Row([dbc.Col(dcc.Graph(figure=fig_estancia), xs=12, md=4), dbc.Col(dcc.Graph(figure=fig_giro), xs=12, md=4), dbc.Col(dcc.Graph(figure=fig_promedio_vm), xs=12, md=4)]),
                html.H5("Indicadores de Calidad e Infecciones", style={'color': '#ffffff', 'fontWeight': '600', 'marginTop': '40px', 'marginBottom': '20px', 'textAlign': 'center'}),
                dbc.Row([dbc.Col(dcc.Graph(figure=fig_bacteriemia), xs=12, md=6), dbc.Col(dcc.Graph(figure=fig_nav), xs=12, md=6)]),
                dbc.Row([dbc.Col(dcc.Graph(figure=fig_istu), xs=12, md=6), dbc.Col(dcc.Graph(figure=fig_lesiones), xs=12, md=6)], className="mt-3"),
                html.H5("Gérmenes Aislados y Eventos Adversos", style={'color': '#ffffff', 'fontWeight': '600', 'marginTop': '40px', 'marginBottom': '20px', 'textAlign': 'center'}),
                dbc.Row([
                    dbc.Col(xs=12, md=2),
                    dbc.Col(dcc.Graph(figure=fig_germenes), xs=12, md=8),
                    dbc.Col(xs=12, md=2)
                ], className="mb-4", justify="center"),
                dbc.Row([
                    dbc.Col(dcc.Graph(figure=fig_eventos), xs=12)
                ])
            ])
        except Exception as e:
            return html.Div([html.H3(f"Error: {str(e)}", style={'color': '#e74c3c'})])
    
    return html.Div([html.H5("Selecciona una pestaña", style={'color': '#ffffff'})])


# Callback para seleccionar todos los meses y actualizar contenido
@app.callback(
    [Output('filtro-mes-inicio', 'value'),
     Output('resumen-content', 'children')],
    [Input('btn-todos-meses', 'n_clicks'),
     Input('btn-aplicar-filtros', 'n_clicks')],
    [dash.dependencies.State('filtro-mes-inicio', 'value')],
    prevent_initial_call=False
)
def actualizar_resumen(btn_todos, btn_aplicar, meses_seleccionados):
    ctx = dash.callback_context
    triggered_id = ctx.triggered[0]['prop_id'].split('.')[0] if ctx.triggered else None
    
    # Si se presionó "Todos los meses", usar todos los meses
    if triggered_id == 'btn-todos-meses':
        meses_seleccionados = meses_disponibles
    
    try:
        # Manejar selección múltiple de meses
        if meses_seleccionados and isinstance(meses_seleccionados, list) and len(meses_seleccionados) > 0:
            meses_filtrados = [m for m in meses_seleccionados if m in meses_disponibles]
        else:
            meses_filtrados = meses_disponibles
        
        # Crear texto del rango
        if len(meses_filtrados) == 1:
            rango_texto = meses_completos.get(meses_filtrados[0], meses_filtrados[0])
        elif len(meses_filtrados) == len(meses_disponibles):
            rango_texto = "Ene - Oct"
        else:
            rango_texto = ", ".join([meses_completos.get(m, m)[:3] for m in meses_filtrados])
        
        suma_col = [c for c in df.columns if str(c).strip() in ['∑', 'SUMA']]
        suma_col = suma_col[0] if suma_col else '∑'
        
        datos_ocupacional = safe_get_row('% OCUPACIONAL', meses_filtrados)
        promedio_ocupacional = datos_ocupacional[datos_ocupacional > 0].mean() if len(datos_ocupacional[datos_ocupacional > 0]) > 0 else 0
        datos_ingresos_kpi = safe_get_row('NUMERO DE INGRESOS DEL MES', meses_filtrados)
        total_ingresos = datos_ingresos_kpi.sum() if isinstance(datos_ingresos_kpi, pd.Series) else datos_ingresos_kpi
        datos_fallecidos = safe_get_row('FALLECIDOS', meses_filtrados)
        total_fallecidos = datos_fallecidos.sum() if isinstance(datos_fallecidos, pd.Series) else datos_fallecidos
        datos_egresos = safe_get_row('NÚMERO DE EGRESOS DEL MES', meses_filtrados)
        total_egresos = datos_egresos.sum() if isinstance(datos_egresos, pd.Series) else datos_egresos
        datos_estancia = safe_get_row('PROMEDIO ESTANCIA', meses_filtrados)
        promedio_estancia = datos_estancia[datos_estancia > 0].mean() if len(datos_estancia[datos_estancia > 0]) > 0 else 0
        datos_giro = safe_get_row('GIRO CAMA', meses_filtrados)
        giro_cama = datos_giro[datos_giro > 0].mean() if len(datos_giro[datos_giro > 0]) > 0 else 0
        
        if pd.isna(promedio_ocupacional): promedio_ocupacional = 0
        if pd.isna(total_ingresos): total_ingresos = 0
        if pd.isna(total_fallecidos): total_fallecidos = 0
        if pd.isna(total_egresos) or total_egresos == 0: total_egresos = 1
        if pd.isna(promedio_estancia): promedio_estancia = 0
        if pd.isna(giro_cama): giro_cama = 0
        mortalidad = (total_fallecidos / total_egresos * 100)

        kpi_ingresos = create_kpi_card("Total Ingresos", f"{int(total_ingresos)}", rango_texto, "#3498db", "📥")
        kpi1 = create_kpi_card("% Ocupación Promedio", f"{promedio_ocupacional:.2f}%", rango_texto, "#2c3e50", "🏥")
        kpi2 = create_kpi_card("% Mortalidad General", f"{mortalidad:.2f}%", f"Total {rango_texto}", "#e74c3c", "💔")
        kpi3 = create_kpi_card("Estancia Promedio", f"{promedio_estancia:.2f}", f"Días - {rango_texto}", "#27ae60", "⏱️")
        kpi4 = create_kpi_card("Giro de Cama", f"{giro_cama:.2f}", f"Promedio {rango_texto}", "#f39c12", "🔄")
        
        # EPS
        total_famisanar = safe_get_row('FAMISANAR', suma_col, silent=True)
        total_compensar = safe_get_row('COMPENSAR', suma_col, silent=True)
        total_salud_total = safe_get_row('SALUD TOTAL', suma_col, silent=True)
        if pd.isna(total_famisanar): total_famisanar = 0
        if pd.isna(total_compensar): total_compensar = 0
        if pd.isna(total_salud_total): total_salud_total = 0
        total_ingresos_anual = safe_get_row('NUMERO DE INGRESOS DEL MES', suma_col, silent=True)
        if pd.isna(total_ingresos_anual) or total_ingresos_anual == 0: total_ingresos_anual = 1
        pct_famisanar = (total_famisanar / total_ingresos_anual * 100)
        pct_compensar = (total_compensar / total_ingresos_anual * 100)
        pct_salud_total = (total_salud_total / total_ingresos_anual * 100)
        
        kpi_eps1 = create_kpi_card("Famisanar", f"{int(total_famisanar)}", f"{pct_famisanar:.1f}% del total", "#27ae60", "🏢")
        kpi_eps2 = create_kpi_card("Compensar", f"{int(total_compensar)}", f"{pct_compensar:.1f}% del total", "#f39c12", "🏢")
        kpi_eps3 = create_kpi_card("Salud Total", f"{int(total_salud_total)}", f"{pct_salud_total:.1f}% del total", "#e74c3c", "🏢")

        # Gráficos
        datos_ocup = safe_get_row('% OCUPACIONAL', meses_filtrados)
        fig_ocupacional = px.line(x=meses_filtrados, y=datos_ocup, title='% Ocupacional Mensual', markers=True)
        fig_ocupacional.update_traces(line_color='#16a085', line_width=3, marker=dict(size=10))
        fig_ocupacional.update_layout(showlegend=False, plot_bgcolor='#1a202c', paper_bgcolor='#2d3748',
            title={'font': {'size': 16, 'color': '#ffffff'}, 'x': 0.5}, xaxis={'gridcolor': '#4a5568', 'color': '#e2e8f0'},
            yaxis={'gridcolor': '#4a5568', 'color': '#e2e8f0'}, hoverlabel=dict(bgcolor='white', font_color='#000000'))
        
        datos_ingresos = safe_get_row('NUMERO DE INGRESOS DEL MES', meses_filtrados)
        datos_egresos_chart = safe_get_row('NÚMERO DE EGRESOS DEL MES', meses_filtrados)
        fig_ingresos_egresos = go.Figure()
        fig_ingresos_egresos.add_trace(go.Bar(name='Ingresos', x=meses_filtrados, y=datos_ingresos, marker_color='#3498db'))
        fig_ingresos_egresos.add_trace(go.Bar(name='Egresos', x=meses_filtrados, y=datos_egresos_chart, marker_color='#27ae60'))
        fig_ingresos_egresos.update_layout(title={'text': 'Ingresos vs. Egresos', 'font': {'size': 16, 'color': '#ffffff'}, 'x': 0.5},
            barmode='group', plot_bgcolor='#1a202c', paper_bgcolor='#2d3748', xaxis={'gridcolor': '#4a5568', 'color': '#e2e8f0'},
            yaxis={'gridcolor': '#4a5568', 'color': '#e2e8f0'}, legend=dict(orientation="h", y=-0.15, x=0.5, xanchor="center", 
            bgcolor='rgba(26,32,44,0.95)', font={'color': '#ffffff'}), hoverlabel=dict(bgcolor='white', font_color='#000000'))

        # Gráficos de Morbilidad
        fig_morbilidad_total = go.Figure()
        fig_morbilidad_genero = go.Figure()
        if not df_morbilidad.empty:
            colors_morb = ['#e74c3c', '#3498db', '#2ecc71', '#f39c12', '#9b59b6', '#1abc9c', '#e67e22', '#34495e', '#16a085', '#c0392b']
            fig_morbilidad_total = go.Figure(data=[go.Bar(x=df_morbilidad.index, y=df_morbilidad['Total'], marker_color=colors_morb)])
            fig_morbilidad_total.update_layout(title={'text': 'Morbilidad por Diagnóstico', 'font': {'size': 16, 'color': '#ffffff'}, 'x': 0.5},
                plot_bgcolor='#1a202c', paper_bgcolor='#2d3748', xaxis={'gridcolor': '#4a5568', 'color': '#e2e8f0', 'tickangle': -45},
                yaxis={'gridcolor': '#4a5568', 'color': '#e2e8f0'}, hoverlabel=dict(bgcolor='white', font_color='#000000'))
            
            fig_morbilidad_genero = go.Figure(data=[
                go.Bar(name='Hombres', x=df_morbilidad.index, y=df_morbilidad['Hombre'], marker_color='#3498db'),
                go.Bar(name='Mujeres', x=df_morbilidad.index, y=df_morbilidad['Mujer'], marker_color='#e91e63')])
            fig_morbilidad_genero.update_layout(title={'text': 'Morbilidad por Sexo', 'font': {'size': 16, 'color': '#ffffff'}, 'x': 0.5},
                barmode='group', plot_bgcolor='#1a202c', paper_bgcolor='#2d3748', xaxis={'gridcolor': '#4a5568', 'color': '#e2e8f0', 'tickangle': -45},
                yaxis={'gridcolor': '#4a5568', 'color': '#e2e8f0'}, legend=dict(orientation="h", y=-0.35, x=0.5, xanchor="center",
                bgcolor='rgba(26,32,44,0.95)', font={'color': '#ffffff'}), hoverlabel=dict(bgcolor='white', font_color='#000000'))

        return meses_filtrados, html.Div([
            dbc.Row([dbc.Col(kpi_ingresos, xs=12, sm=6, md=True, className="mb-3"), dbc.Col(kpi1, xs=12, sm=6, md=True, className="mb-3"),
                     dbc.Col(kpi2, xs=12, sm=6, md=True, className="mb-3"), dbc.Col(kpi3, xs=12, sm=6, md=True, className="mb-3"),
                     dbc.Col(kpi4, xs=12, sm=6, md=True, className="mb-3")], className="mb-3 g-3"),
            html.H6("Principales EPS", style={'color': '#e2e8f0', 'fontWeight': '600', 'marginBottom': '15px', 'marginTop': '10px'}),
            dbc.Row([dbc.Col(kpi_eps1, xs=12, sm=6, md=4, className="mb-3"), dbc.Col(kpi_eps2, xs=12, sm=6, md=4, className="mb-3"),
                     dbc.Col(kpi_eps3, xs=12, sm=6, md=4, className="mb-3")], className="mb-4 g-3"),
            dbc.Row([dbc.Col(dcc.Graph(figure=fig_ocupacional), xs=12, md=6), dbc.Col(dcc.Graph(figure=fig_ingresos_egresos), xs=12, md=6)]),
            html.H6("Morbilidad", style={'color': '#ffffff', 'fontWeight': '600', 'marginTop': '30px', 'marginBottom': '20px', 'textAlign': 'center'}),
            dbc.Row([dbc.Col(dcc.Graph(figure=fig_morbilidad_total), xs=12, md=6), dbc.Col(dcc.Graph(figure=fig_morbilidad_genero), xs=12, md=6)])
        ])
    except Exception as e:
        return dash.no_update, html.Div([html.H3(f"Error: {str(e)}", style={'color': '#e74c3c'})])


if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=8050)
