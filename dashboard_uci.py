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

# --- CARGA DE DATOS DE FALLECIDOS (MORTALIDAD) - Edad y Sexo ---
df_fallecidos = pd.DataFrame()
try:
    import unicodedata
    df_mort = pd.read_excel(EXCEL_PATH, sheet_name='MORTALIDAD', header=None, usecols='A:F')
    
    meses_separadores = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 
                         'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE']
    
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
    
    fallecidos_data = []
    for i, row in df_mort.iterrows():
        nombre = str(row[0]).strip() if pd.notna(row[0]) else ''
        edad = row[1]
        
        if nombre.upper().replace(' ', '') in [m.upper() for m in meses_separadores]:
            continue
        if not nombre or nombre.upper() == 'NAN':
            continue
        if not pd.notna(edad) or not isinstance(edad, (int, float)):
            continue
        
        sexo = 'Mujer' if es_nombre_femenino(nombre) else 'Hombre'
        fallecidos_data.append({'Nombre': nombre, 'Edad': int(edad), 'Sexo': sexo})
    
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
    for i in range(1, 100):
        cell = ws.cell(i, 5)
        if cell.value:
            fill_color = cell.fill.fgColor.rgb if cell.fill.fgColor else None
            if fill_color and fill_color == 'FF00B0F0':
                evento = str(cell.value).strip()
                eventos_azules.append(evento)
    
    def normalizar_evento(e):
        e_lower = e.lower()
        if 'desplazamiento' in e_lower:
            return 'Desplazamiento'
        if 'autoretiro' in e_lower:
            return 'Autoretiro'
        if 'error' in e_lower:
            return 'Error procedimiento'
        if 'medicamento' in e_lower:
            return 'Error medicamentos'
        if 'lecash' in e_lower:
            return 'LECASH'
        if 'marsi' in e_lower:
            return 'MARSI'
        if 'lesión' in e_lower or 'lesion' in e_lower:
            return 'Lesiones'
        if 'caida' in e_lower or 'caída' in e_lower:
            return 'Caídas'
        return e[:20] if len(e) > 20 else e
    
    eventos_norm = [normalizar_evento(e) for e in eventos_azules]
    df_eventos_adversos = pd.Series(eventos_norm).value_counts().reset_index()
    df_eventos_adversos.columns = ['Evento', 'Cantidad']
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

# Paleta de colores profesional
COLOR_PALETTE = {
    'primary': '#1a365d',
    'secondary': '#2c5282',
    'accent': '#3182ce',
    'success': '#38a169',
    'danger': '#e53e3e',
    'warning': '#dd6b20',
    'info': '#4299e1',
    'light': '#f7fafc',
    'dark': '#1a202c',
    'muted': '#718096',
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
        plot_bgcolor='#f8fafc',
        paper_bgcolor='#ffffff',
        title={'font': {'size': 16, 'color': COLOR_PALETTE['dark']}, 'x': 0.5, 'xanchor': 'center'},
        xaxis={'gridcolor': '#e2e8f0', 'showline': True, 'linecolor': '#cbd5e0', 'color': COLOR_PALETTE['dark']},
        yaxis={'gridcolor': '#e2e8f0', 'showline': True, 'linecolor': '#cbd5e0', 'color': COLOR_PALETTE['dark'], 'tickformat': 'd'},
        margin=dict(l=50, r=30, t=60, b=40),
        hoverlabel=dict(bgcolor='white', font_size=14, font_color=COLOR_PALETTE['dark'], bordercolor='#cbd5e0')
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
        marker=dict(size=8, line=dict(width=2, color='#ffffff')),
        hovertemplate=hover_format
    )
    
    yaxis_config = {'gridcolor': '#e2e8f0', 'showline': True, 'linecolor': '#cbd5e0', 'color': COLOR_PALETTE['dark']}
    if isinstance(data, pd.Series):
        max_val = data.max() if not data.empty else 0
        if pd.notna(max_val) and max_val <= 5:
            yaxis_config['dtick'] = 1
    
    fig.update_layout(
        showlegend=False,
        plot_bgcolor='#f8fafc',
        paper_bgcolor='#ffffff',
        title={'font': {'size': 16, 'color': COLOR_PALETTE['dark']}, 'x': 0.5, 'xanchor': 'center'},
        xaxis={'gridcolor': '#e2e8f0', 'showline': True, 'linecolor': '#cbd5e0', 'color': COLOR_PALETTE['dark']},
        yaxis=yaxis_config,
        margin=dict(l=50, r=30, t=60, b=40),
        hoverlabel=dict(bgcolor='white', font_size=14, font_color=COLOR_PALETTE['dark'], bordercolor='#cbd5e0')
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
        title={'text': title, 'font': {'size': 16, 'color': COLOR_PALETTE['dark']}, 'x': 0.5, 'xanchor': 'center'},
        barmode='group',
        plot_bgcolor='#f8fafc',
        paper_bgcolor='#ffffff',
        xaxis={'gridcolor': '#e2e8f0', 'color': COLOR_PALETTE['dark']},
        yaxis={'gridcolor': '#e2e8f0', 'color': COLOR_PALETTE['dark'], 'tickformat': 'd'},
        margin=dict(l=50, r=30, t=90, b=40),
        legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5, 
                   bgcolor='rgba(255,255,255,0.95)', font={'color': COLOR_PALETTE['dark']}),
        hoverlabel=dict(bgcolor='white', font_size=13, bordercolor='#cbd5e0')
    )
    return fig

def create_kpi_card(title, value, subtitle="", color=None, icon="📊"):
    return dbc.Card(
        dbc.CardBody([
            html.Div([
                html.Span(icon, style={'fontSize': '28px', 'marginRight': '12px', 'flexShrink': '0'}),
                html.Div([
                    html.H6(capitalizar_texto(title), style={
                        'color': COLOR_PALETTE['muted'], 'fontSize': '13px', 'marginBottom': '8px', 
                        'fontWeight': '600', 'textTransform': 'uppercase', 'letterSpacing': '0.5px'
                    }),
                    html.H2(value, style={
                        'color': color or COLOR_PALETTE['primary'], 'fontWeight': '700', 
                        'marginBottom': '5px', 'fontSize': '28px', 'lineHeight': '1'
                    }),
                    html.P(capitalizar_texto(subtitle), style={
                        'color': COLOR_PALETTE['muted'], 'fontSize': '12px', 'margin': '0', 'fontWeight': '500'
                    })
                ], style={'flex': '1', 'minWidth': '0', 'overflow': 'hidden'})
            ], style={'display': 'flex', 'alignItems': 'center'})
        ], style={'padding': '16px'}),
        className="kpi-card h-100",
        style={
            'border': 'none', 'borderLeft': f'4px solid {color or COLOR_PALETTE["primary"]}',
            'borderRadius': '8px', 'boxShadow': '0 4px 12px rgba(0,0,0,0.08)',
            'background': '#ffffff', 'transition': 'all 0.3s ease', 'height': '100%'
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
                background-color: #f8fafc;
                color: #1a365d;
            }
            
            @keyframes fadeIn {
                from { opacity: 0; transform: translateY(20px); }
                to { opacity: 1; transform: translateY(0); }
            }
            
            .kpi-card {
                transition: all 0.3s ease;
                border-radius: 8px !important;
                box-shadow: 0 4px 12px rgba(0,0,0,0.08) !important;
            }
            
            .kpi-card:hover {
                transform: translateY(-5px);
                box-shadow: 0 8px 16px rgba(0,0,0,0.12) !important;
            }
            
            .fade-in { animation: fadeIn 0.6s ease-out; }
            
            .nav-tabs .nav-link {
                color: #718096;
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
                border-bottom: 3px solid #3182ce;
                color: #3182ce;
                background-color: rgba(49, 130, 206, 0.05);
            }
            
            .nav-tabs .nav-link.active {
                color: #2c5282;
                border-bottom: 3px solid #2c5282;
                background: rgba(44, 82, 130, 0.08);
                font-weight: 600;
            }
            
            .tab-content {
                background-color: white;
                border-radius: 0 8px 8px 8px;
                padding: 20px;
                box-shadow: 0 4px 12px rgba(0,0,0,0.05);
            }
            
            .Select-control { background-color: #ffffff !important; border-color: #cbd5e0 !important; border-radius: 6px !important; }
            .Select-value-label, .Select-placeholder, .Select-value { color: #1a365d !important; }
            .Select-menu-outer { background-color: #ffffff !important; border-color: #cbd5e0 !important; }
            .Select-option { background-color: #ffffff !important; color: #1a365d !important; }
            .Select-option:hover, .Select-option.is-focused { background-color: #ebf8ff !important; color: #2c5282 !important; }
            .Select-option.is-selected { background-color: #3182ce !important; color: #ffffff !important; }
            
            .dashboard-footer {
                background-color: #ffffff;
                border-top: 1px solid #e2e8f0;
                padding: 20px 0;
                margin-top: 40px;
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
                    html.Img(src='/assets/Logo.png', style={'height': '60px', 'marginRight': '20px', 'filter': 'drop-shadow(0 2px 4px rgba(0,0,0,0.1))'}),
                    html.Div([
                        html.H1("DASHBOARD UCI ADULTOS 2025", style={
                            'color': COLOR_PALETTE['primary'], 'fontWeight': '700', 'marginBottom': '5px',
                            'fontSize': '28px', 'letterSpacing': '0.5px', 'margin': '0'
                        }),
                        html.P("Sistema de Monitoreo y Análisis de Indicadores", style={
                            'color': COLOR_PALETTE['muted'], 'fontSize': '14px', 'marginBottom': '0', 'fontWeight': '400'
                        })
                    ])
                ], style={'display': 'flex', 'alignItems': 'center', 'marginBottom': '15px'})
            ]),
            html.Div([
                dbc.Button("📊 Exportar a Excel", color="primary", size="sm", className="me-2", style={'borderRadius': '6px', 'fontWeight': '500'}),
                dbc.Button("📄 Exportar a CSV", color="secondary", size="sm", style={'borderRadius': '6px', 'fontWeight': '500'}),
            ], style={'marginBottom': '20px'})
        ], className="fade-in"),
        
        # Tabs
        dbc.Tabs(id="tabs", active_tab='tab-resumen', children=[
            dbc.Tab(label='📊 Resumen general', tab_id='tab-resumen'),
            dbc.Tab(label='📈 Gestión y calidad', tab_id='tab-gestion'),
            dbc.Tab(label='💀 Mortalidad', tab_id='tab-mortalidad'),
            dbc.Tab(label='🏥 Especialidades y FPS', tab_id='tab-especialidades'),
        ], style={'marginBottom': '0', 'borderBottom': 'none', 'backgroundColor': '#ffffff', 'borderRadius': '8px 8px 0 0', 'padding': '0 10px'}),
        
        html.Div(id='tab-content', className="tab-content"),
        
        # Footer
        html.Div([
            html.Div([
                html.Div([
                    html.P("Hospital Infantil Universitario de San José", style={'color': COLOR_PALETTE['dark'], 'fontWeight': '600', 'margin': '0', 'fontSize': '14px'}),
                    html.P("Unidad de Cuidados Intensivos - Adultos", style={'color': COLOR_PALETTE['muted'], 'margin': '0', 'fontSize': '12px'})
                ]),
                html.Div([
                    html.P("© 2025 - Dashboard de Indicadores UCI", style={'color': COLOR_PALETTE['muted'], 'margin': '0', 'fontSize': '12px'})
                ], style={'textAlign': 'right'})
            ], style={'display': 'flex', 'justifyContent': 'space-between', 'alignItems': 'center', 'flexWrap': 'wrap', 'gap': '15px'})
        ], className="dashboard-footer")
        
    ], fluid=True, style={'maxWidth': '1400px', 'padding': '30px 20px', 'backgroundColor': '#f8fafc'})
], style={'backgroundColor': '#f8fafc', 'minHeight': '100vh', 'fontFamily': "'Inter', sans-serif"})
