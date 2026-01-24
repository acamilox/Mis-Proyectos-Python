# -*- coding: utf-8 -*-
import dash
from dash import dcc, html, Input, Output, State
import dash_bootstrap_components as dbc
import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import os
from io import BytesIO
from datetime import datetime

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
    
    # Procesar datos de morbilidad - extraer totales por diagnóstico (todas las categorías)
    diagnosticos = ['ACV', 'CARDIOVASCULAR', 'RESPIRATORIO', 'SEPSIS', 'POP', 'DM', 
                    'GINECO-OBSTE', 'GINECO-ONCOLOGIA', 'GINECO', 'TRAUMA', 'TOXICOLOGIA', 'OTROS']
    
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
    # Filtrar solo los que tienen datos > 0 y ordenar de mayor a menor
    df_morbilidad = df_morbilidad[df_morbilidad['Total'] > 0]
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

# --- CARGA DE TODOS LOS EVENTOS ADVERSOS (columna E) ---
df_eventos_adversos = pd.DataFrame()
try:
    wb = load_workbook(EXCEL_PATH)
    ws = wb['EVENTOS ADVERSOS']
    
    meses_sep = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 
                 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE']
    
    def normalizar_evento(e):
        """Normaliza el nombre del evento para agrupar variantes similares"""
        e_lower = e.lower().strip()
        
        # Autoextubación
        if 'autoextubación' in e_lower or 'autoextubacion' in e_lower:
            return 'Autoextubación'
        
        # Autoretiro de dispositivos (especificar cuál)
        if 'autoretiro' in e_lower:
            if 'cvc' in e_lower:
                return 'Autoretiro CVC'
            if 'picc' in e_lower:
                return 'Autoretiro PICC'
            if 'ventriculostomia' in e_lower:
                return 'Autoretiro Ventriculostomía'
            return 'Autoretiro de dispositivo'
        
        # Desplazamiento de dispositivos (especificar cuál)
        if 'desplazamiento' in e_lower:
            if 'gastrostomia' in e_lower or 'sonda' in e_lower:
                return 'Desplazamiento sonda/gastrostomía'
            if 'cvc' in e_lower or 'venoso central' in e_lower:
                return 'Desplazamiento CVC'
            if 'hemodialisis' in e_lower or 'hemodiálisis' in e_lower:
                return 'Desplazamiento catéter hemodiálisis'
            if 'arterial' in e_lower:
                return 'Desplazamiento línea arterial'
            if 'torax' in e_lower or 'tórax' in e_lower:
                return 'Desplazamiento tubo tórax'
            if 'nefrostomia' in e_lower:
                return 'Desplazamiento nefrostomía'
            if 'biliar' in e_lower:
                return 'Desplazamiento derivación biliar'
            if 'parecentesis' in e_lower or 'paracentesis' in e_lower:
                return 'Desplazamiento catéter paracentesis'
            return 'Desplazamiento de dispositivo'
        
        # LPP (Lesiones por presión)
        if 'lpp' in e_lower or ('lesion' in e_lower and 'presion' in e_lower):
            if 'dispositivo' in e_lower or 'medico' in e_lower or 'collar' in e_lower:
                return 'LPP por dispositivo médico'
            if 'sacra' in e_lower or 'sacro' in e_lower:
                return 'LPP región sacra'
            if 'talon' in e_lower or 'talón' in e_lower:
                return 'LPP talón'
            if 'occipital' in e_lower:
                return 'LPP occipital'
            if 'gluteo' in e_lower or 'glúteo' in e_lower:
                return 'LPP glúteo'
            return 'LPP'
        
        # LECASH
        if 'lecash' in e_lower:
            return 'LECASH'
        
        # MARSI
        if 'marsi' in e_lower:
            return 'Lesión MARSI'
        
        # Errores de medicación
        if 'error' in e_lower or 'medicamento' in e_lower or 'medicación' in e_lower:
            if 'contraste' in e_lower:
                return 'Error medicamento (contraste)'
            if 'retiro' in e_lower:
                return 'Error en retiro CVC'
            if 'retraso' in e_lower:
                return 'Retraso administración medicamento'
            return 'Error de medicación'
        
        # Problema con medicamento
        if 'problema' in e_lower and 'medicamento' in e_lower:
            return 'Problema con medicamento'
        
        # Retraso medicamento
        if 'retraso' in e_lower:
            return 'Retraso administración medicamento'
        
        # Caídas
        if 'caida' in e_lower or 'caída' in e_lower:
            return 'Caída'
        
        # Flebitis
        if 'flebitis' in e_lower:
            return 'Flebitis'
        
        # Neumotórax
        if 'neumotórax' in e_lower or 'neumotorax' in e_lower:
            return 'Neumotórax'
        
        # Si no coincide con nada, devolver el evento original capitalizado
        return e.strip().capitalize()
    
    todos_eventos = []
    for i in range(1, 200):
        cell = ws.cell(i, 5)
        if cell.value:
            evento = str(cell.value).strip()
            if evento.upper() not in meses_sep and evento and evento.upper() != 'NAN':
                todos_eventos.append(normalizar_evento(evento))
    
    # Agrupar eventos y contar
    df_eventos_adversos = pd.Series(todos_eventos).value_counts().reset_index()
    df_eventos_adversos.columns = ['Evento', 'Cantidad']
    print(f"✅ Datos de eventos adversos cargados: {len(todos_eventos)} eventos, {len(df_eventos_adversos)} categorías")
except Exception as e:
    print(f"⚠️ Error cargando eventos adversos: {e}")

# --- CARGA DE DATOS GRAM - SOT (Trasplantes/Donaciones) ---
df_gram_sot = pd.DataFrame()
try:
    df_gram_raw = pd.read_excel(EXCEL_PATH, sheet_name='GRAM - SOT', header=None)
    meses_gram = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Agos', 'Sep', 'Oct', 'Nov', 'Dic']
    meses_map_gram = {
        'enero': 'Ene', 'febrero': 'Feb', 'marzo': 'Mar', 'abril': 'Abr',
        'mayo': 'May', 'junio': 'Jun', 'julio': 'Jul', 'agosto': 'Agos',
        'septiembre': 'Sep', 'octubre': 'Oct', 'noviembre': 'Nov', 'diciembre': 'Dic'
    }
    gram_data = {}
    for i, row in df_gram_raw.iterrows():
        mes_raw = str(row[0]).strip().lower() if pd.notna(row[0]) else ''
        mes_key = meses_map_gram.get(mes_raw, None)
        if mes_key:
            valor = row[1] if pd.notna(row[1]) else 0
            gram_data[mes_key] = int(valor) if pd.notna(valor) else 0
    
    df_gram_sot = pd.DataFrame([gram_data])
    df_gram_sot = df_gram_sot.reindex(columns=meses_gram, fill_value=0)
    print(f"✅ Datos GRAM-SOT cargados: {df_gram_sot.values.sum()} total")
except Exception as e:
    print(f"⚠️ Error cargando GRAM-SOT: {e}")

# Meses disponibles
meses_disponibles = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Agos', 'Sep', 'Oct', 'Nov', 'Dic']

meses_completos = {
    'Ene': 'Enero', 'Feb': 'Febrero', 'Mar': 'Marzo', 'Abr': 'Abril',
    'May': 'Mayo', 'Jun': 'Junio', 'Jul': 'Julio', 'Agos': 'Agosto',
    'Sep': 'Septiembre', 'Oct': 'Octubre', 'Nov': 'Noviembre', 'Dic': 'Diciembre'
}

# Paleta de colores - Tema Premium Moderno
COLOR_PALETTE = {
    'primary': '#6366f1',      # Indigo vibrante
    'secondary': '#10b981',    # Esmeralda
    'accent': '#06b6d4',       # Cyan
    'success': '#22c55e',      # Verde brillante
    'danger': '#ef4444',       # Rojo coral
    'warning': '#f59e0b',      # Ámbar
    'info': '#8b5cf6',         # Violeta
    'light': '#f1f5f9',
    'dark': '#0f172a',
    'muted': '#94a3b8',
    'bg_primary': '#0f172a',   # Slate 900
    'bg_secondary': '#1e293b', # Slate 800
    'bg_card': '#1e293b',
    'text_primary': '#f8fafc',
    'text_secondary': '#cbd5e1',
    'border': '#334155',
    'gradient_start': '#0f172a',
    'gradient_end': '#1e1b4b',
}

def capitalizar_texto(texto):
    if not texto or not isinstance(texto, str):
        return texto
    return texto[0].upper() + texto[1:].lower() if len(texto) > 1 else texto.upper()

def formato_porcentaje(valor, con_simbolo=True):
    """Formatea porcentaje: sin decimales si es entero, con 1 decimal si no."""
    simbolo = "%" if con_simbolo else ""
    if valor == int(valor):
        return f"{int(valor)}{simbolo}"
    return f"{valor:.1f}{simbolo}"

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
    
    bar_color = color or COLOR_PALETTE['primary']
    if decimals == 0:
        hover_format = '<b>%{x}</b><br>Valor: %{y:.0f}<extra></extra>'
        text_template = '%{y:.0f}'
    else:
        hover_format = f'<b>%{{x}}</b><br>Valor: %{{y:.{decimals}f}}<extra></extra>'
        text_template = f'%{{y:.{decimals}f}}'
    
    fig.update_traces(
        marker_color=bar_color,
        marker_line_color='rgba(255,255,255,0.2)',
        marker_line_width=1,
        hovertemplate=hover_format,
        text=data,
        texttemplate=text_template,
        textposition='inside',
        insidetextanchor='middle',
        textfont=dict(color='white', size=12, family='Inter', weight=700)
    )
    fig.update_layout(
        showlegend=False,
        plot_bgcolor='rgba(15,23,42,0.6)',
        paper_bgcolor='rgba(30,41,59,0.8)',
        title={'font': {'size': 14, 'color': '#f1f5f9', 'family': 'Inter'}, 'x': 0.5, 'xanchor': 'center'},
        xaxis={'gridcolor': 'rgba(148,163,184,0.15)', 'showline': False, 'color': '#94a3b8', 'tickfont': {'size': 11}, 'title': ''},
        yaxis={'gridcolor': 'rgba(148,163,184,0.15)', 'showline': False, 'color': '#94a3b8', 'tickformat': 'd', 'tickfont': {'size': 11}, 'title': ''},
        margin=dict(l=50, r=30, t=50, b=40),
        height=380,
        hoverlabel=dict(bgcolor='#1e293b', font_size=13, font_color='#f1f5f9', bordercolor=bar_color),
        bargap=0.3
    )
    return fig

def create_line_chart(row_name, title, color=None, silent=False, decimals=1):
    data = safe_get_row(row_name, meses_disponibles, silent=silent)
    line_color = color or COLOR_PALETTE['secondary']
    
    # Formatear texto según decimales
    if decimals == 0:
        text_values = [f'{int(v)}' if pd.notna(v) else '' for v in data]
    else:
        text_values = [f'{v:.{decimals}f}' if pd.notna(v) else '' for v in data]
    
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=meses_disponibles, 
        y=data,
        mode='lines+markers+text',
        line=dict(color=line_color, width=2),
        marker=dict(size=10, color=line_color, line=dict(width=2, color='#0f172a')),
        text=text_values,
        textposition='top center',
        textfont=dict(color='#f1f5f9', size=11),
        hovertemplate=f'<b>%{{x}}</b><br>Valor: %{{y:.{decimals}f}}<extra></extra>'
    ))
    
    yaxis_config = {'gridcolor': 'rgba(148,163,184,0.15)', 'showline': False, 'color': '#94a3b8', 'tickfont': {'size': 11}}
    if isinstance(data, pd.Series):
        max_val = data.max() if not data.empty else 0
        if pd.notna(max_val) and max_val <= 5:
            yaxis_config['dtick'] = 1
    
    fig.update_layout(
        title={'text': title, 'font': {'size': 14, 'color': '#f1f5f9', 'family': 'Inter'}, 'x': 0.5, 'xanchor': 'center'},
        showlegend=False,
        plot_bgcolor='rgba(15,23,42,0.6)',
        paper_bgcolor='rgba(30,41,59,0.8)',
        xaxis={'gridcolor': 'rgba(148,163,184,0.15)', 'showline': False, 'color': '#94a3b8', 'tickfont': {'size': 11}, 'title': ''},
        yaxis=yaxis_config,
        margin=dict(l=50, r=30, t=50, b=40),
        height=380,
        hoverlabel=dict(bgcolor='#1e293b', font_size=13, font_color='#f1f5f9', bordercolor=line_color)
    )
    return fig

def create_grouped_bar_chart(row_names, labels, title, decimals=0):
    fig = go.Figure()
    colors = ['#6366f1', '#10b981', '#06b6d4', '#f59e0b']
    
    for i, (row_name, label) in enumerate(zip(row_names, labels)):
        data = safe_get_row(row_name, meses_disponibles)
        
        if decimals == 0:
            hover_template = f'<b>%{{x}}</b><br>{label}: %{{y:.0f}}<extra></extra>'
            text_template = '%{y:.0f}'
        else:
            hover_template = f'<b>%{{x}}</b><br>{label}: %{{y:.{decimals}f}}<extra></extra>'
            text_template = f'%{{y:.{decimals}f}}'
        
        fig.add_trace(go.Bar(
            name=label, x=meses_disponibles, y=data,
            marker_color=colors[i % len(colors)],
            marker_line_color='rgba(255,255,255,0.2)',
            marker_line_width=1,
            hovertemplate=hover_template,
            text=data,
            texttemplate=text_template,
            textposition='inside',
            insidetextanchor='middle',
            textfont=dict(color='white', size=11, family='Inter')
        ))
    
    fig.update_layout(
        title={'text': title, 'font': {'size': 14, 'color': '#f1f5f9', 'family': 'Inter'}, 'x': 0.5, 'xanchor': 'center'},
        barmode='group',
        plot_bgcolor='rgba(15,23,42,0.6)',
        paper_bgcolor='rgba(30,41,59,0.8)',
        xaxis={'gridcolor': 'rgba(148,163,184,0.15)', 'color': '#94a3b8', 'tickfont': {'size': 11}, 'title': ''},
        yaxis={'gridcolor': 'rgba(148,163,184,0.15)', 'color': '#94a3b8', 'tickformat': 'd', 'tickfont': {'size': 11}, 'title': ''},
        margin=dict(l=50, r=30, t=50, b=70),
        height=400,
        legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5, 
                   bgcolor='rgba(15,23,42,0.9)', font={'color': '#f1f5f9', 'size': 11}, bordercolor='#334155', borderwidth=1),
        hoverlabel=dict(bgcolor='#1e293b', font_size=13, font_color='#f1f5f9', bordercolor='#6366f1'),
        bargap=0.25
    )
    return fig

def create_donut_chart(labels, values, title, colors=None, auto_sort=True):
    """Crea un gráfico de dona ordenado de mayor a menor"""
    # Paleta de colores
    if colors is None:
        colors = [
            '#6366f1', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', 
            '#06b6d4', '#ec4899', '#84cc16', '#f97316', '#14b8a6',
            '#a855f7', '#eab308', '#3b82f6', '#22c55e', '#e11d48',
            '#0ea5e9', '#d946ef', '#facc15', '#64748b', '#fb923c',
        ]
    
    if auto_sort:
        # Ordenar de mayor a menor
        datos = sorted(zip(values, labels, colors[:len(labels)]), reverse=True)
        values_final = [d[0] for d in datos]
        labels_final = [d[1] for d in datos]
        colors_final = [d[2] for d in datos]
    else:
        values_final = list(values)
        labels_final = list(labels)
        colors_final = list(colors[:len(labels)])
    
    total = sum(values_final)
    
    # Calcular rotación para centrar el mayor en las 12 (arriba)
    if total > 0:
        primer_pct = values_final[0] / total
        primer_angulo = primer_pct * 360
        rotation_ajustado = -primer_angulo / 2
    else:
        rotation_ajustado = 0
    
    fig = go.Figure(data=[go.Pie(
        labels=labels_final,
        values=values_final,
        hole=0.5,
        marker=dict(colors=colors_final, line=dict(color='#0f172a', width=2)),
        textinfo='percent',
        textposition='inside',
        textfont=dict(size=11, color='white'),
        hovertemplate='<b>%{label}</b><br>Cantidad: %{value}<br>Porcentaje: %{percent}<extra></extra>',
        direction='clockwise',
        sort=False,
        rotation=rotation_ajustado
    )])
    
    # Total en el centro
    fig.add_annotation(
        text=f"<b>Total</b><br>{int(total):,}".replace(',', '.'),
        x=0.5, y=0.5,
        font=dict(size=16, color='#f1f5f9'),
        showarrow=False,
        xref='paper', yref='paper'
    )
    
    fig.update_layout(
        title={'text': title, 'font': {'size': 14, 'color': '#f1f5f9', 'family': 'Inter'}, 'x': 0.5, 'xanchor': 'center'},
        plot_bgcolor='rgba(15,23,42,0.6)',
        paper_bgcolor='rgba(30,41,59,0.8)',
        showlegend=True,
        legend=dict(
            orientation='v',
            yanchor='middle',
            y=0.5,
            xanchor='left',
            x=1.02,
            bgcolor='rgba(15,23,42,0.8)',
            bordercolor='#334155',
            borderwidth=1,
            font=dict(color='#f1f5f9', size=10),
            itemsizing='constant'
        ),
        margin=dict(l=20, r=150, t=50, b=30),
        height=420,
        hoverlabel=dict(bgcolor='#1e293b', font_size=13, font_color='#f1f5f9')
    )
    return fig

def create_kpi_card(title, value, subtitle="", color=None, icon="📊", trend=None, trend_value=None, inverse=False):
    """
    Crea una tarjeta KPI con indicador de tendencia opcional.
    
    Args:
        trend: 'up', 'down', 'stable' o None
        trend_value: texto como "+5.2%" o "-3.1%"
        inverse: True si 'down' es positivo (ej: mortalidad, infecciones)
    """
    accent_color = color or COLOR_PALETTE["primary"]
    if accent_color in ['#34495e', '#2c3e50', '#1e293b']:
        accent_color = '#64748b'
    
    # Determinar color y símbolo de tendencia
    trend_element = None
    if trend and trend_value:
        if trend == 'up':
            trend_color = '#ef4444' if inverse else '#22c55e'  # Rojo si inverse, verde si no
            trend_icon = '▲'
        elif trend == 'down':
            trend_color = '#22c55e' if inverse else '#ef4444'  # Verde si inverse, rojo si no
            trend_icon = '▼'
        else:
            trend_color = '#94a3b8'
            trend_icon = '●'
        
        trend_element = html.Div([
            html.Div('Variación', style={
                'fontSize': '10px',
                'color': '#94a3b8',
                'fontWeight': '500',
                'marginBottom': '2px',
                'textTransform': 'uppercase',
                'letterSpacing': '0.5px'
            }),
            html.Div(trend_value, style={
                'fontSize': '16px', 
                'fontWeight': '700',
                'color': trend_color,
                'marginBottom': '4px'
            }),
            html.Div([
                html.Span(trend_icon, style={'marginRight': '4px', 'fontSize': '10px'}),
                html.Span('Respecto a meses anteriores', style={'fontSize': '9px'})
            ], style={
                'color': '#64748b',
                'display': 'flex',
                'alignItems': 'center',
                'justifyContent': 'center',
            })
        ], style={
            'textAlign': 'center',
            'marginTop': '8px',
            'padding': '8px 12px',
            'borderRadius': '12px',
            'background': f'{trend_color}10',
            'border': f'1px solid {trend_color}25'
        })
    
    card_content = [
        html.Div([
            html.Span(icon, style={
                'fontSize': '28px', 
                'filter': 'drop-shadow(0 0 6px rgba(255,255,255,0.2))'
            })
        ], style={
            'width': '52px', 'height': '52px', 'borderRadius': '14px',
            'background': f'linear-gradient(135deg, {accent_color}30 0%, {accent_color}15 100%)',
            'display': 'flex', 'alignItems': 'center', 'justifyContent': 'center',
            'marginBottom': '14px', 'border': f'1px solid {accent_color}40'
        }),
        html.H6(capitalizar_texto(title), style={
            'color': '#94a3b8', 'fontSize': '11px', 'marginBottom': '8px', 
            'fontWeight': '600', 'textTransform': 'uppercase', 'letterSpacing': '0.5px',
            'minHeight': '26px', 'display': 'flex', 'alignItems': 'center', 'justifyContent': 'center'
        }),
        html.H2(value, style={
            'color': '#f1f5f9', 'fontWeight': '700', 
            'marginBottom': '6px', 'fontSize': '28px', 'lineHeight': '1'
        }),
        html.P(capitalizar_texto(subtitle), style={
            'color': '#64748b', 'fontSize': '12px', 'margin': '0', 'fontWeight': '500',
            'minHeight': '18px'
        })
    ]
    
    if trend_element:
        card_content.append(trend_element)
    else:
        # Placeholder para mantener altura consistente cuando no hay tendencia
        card_content.append(html.Div(style={'height': '32px'}))
    
    return dbc.Card(
        dbc.CardBody([
            html.Div(card_content, style={'textAlign': 'center'})
        ], style={'padding': '20px 16px'}),
        className="kpi-card h-100",
        style={
            'border': f'1px solid {accent_color}30',
            'borderRadius': '16px', 
            'boxShadow': f'0 4px 20px rgba(0,0,0,0.3)',
            'background': 'linear-gradient(145deg, #1e293b 0%, #0f172a 100%)',
            'transition': 'all 0.3s ease', 
            'height': '100%',
            'minHeight': '220px',
            'position': 'relative',
            'overflow': 'hidden'
        }
    )


def calcular_tendencia(datos, inverse=False):
    """Calcula la tendencia comparando los últimos meses con datos."""
    if not isinstance(datos, pd.Series):
        return None, None
    
    # Filtrar valores válidos
    valores = datos.dropna()
    valores = valores[valores > 0]
    
    if len(valores) < 2:
        return None, None
    
    # Comparar último valor con el promedio de los anteriores
    ultimo = valores.iloc[-1]
    promedio_anterior = valores.iloc[:-1].mean()
    
    if promedio_anterior == 0:
        return 'stable', '0%'
    
    cambio_pct = ((ultimo - promedio_anterior) / promedio_anterior) * 100
    
    if abs(cambio_pct) < 2:
        trend = 'stable'
    elif cambio_pct > 0:
        trend = 'up'
    else:
        trend = 'down'
    
    trend_value = f"{'+' if cambio_pct > 0 else ''}{cambio_pct:.1f}%"
    return trend, trend_value


app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP], suppress_callback_exceptions=True)
server = app.server

# Configuración para gráficos estáticos (sin zoom, pan, etc.)
GRAPH_CONFIG = {
    'staticPlot': True,
    'displayModeBar': False
}


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
            @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
            
            * {
                box-sizing: border-box;
            }
            
            body {
                font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
                background: linear-gradient(135deg, #0f172a 0%, #1e1b4b 50%, #0f172a 100%);
                color: #f1f5f9;
                min-height: 100vh;
            }
            
            @keyframes fadeInUp {
                from { opacity: 0; transform: translateY(20px); }
                to { opacity: 1; transform: translateY(0); }
            }
            
            .kpi-card {
                transition: all 0.3s ease;
                border-radius: 16px !important;
                backdrop-filter: blur(10px);
            }
            
            .kpi-card:hover {
                transform: translateY(-4px);
                box-shadow: 0 12px 30px rgba(0,0,0,0.4) !important;
            }
            
            .fade-in { animation: fadeInUp 0.6s ease-out; }
            
            .nav-tabs {
                border-bottom: none !important;
                gap: 6px;
            }
            
            .nav-tabs .nav-link {
                color: #94a3b8;
                border: none;
                border-radius: 10px;
                transition: all 0.3s ease;
                background: rgba(30, 41, 59, 0.6);
                font-weight: 500;
                padding: 12px 18px;
                font-size: 14px;
                border: 1px solid transparent;
            }
            
            .nav-tabs .nav-link:hover {
                color: #f1f5f9;
                background: rgba(99, 102, 241, 0.15);
                border-color: rgba(99, 102, 241, 0.3);
            }
            
            .nav-tabs .nav-link.active {
                color: #f1f5f9;
                background: linear-gradient(135deg, rgba(99, 102, 241, 0.25) 0%, rgba(139, 92, 246, 0.15) 100%);
                font-weight: 600;
                border-color: rgba(99, 102, 241, 0.4);
            }
            
            .tab-content {
                background: transparent;
                border-radius: 16px;
                padding: 24px;
                padding-top: 8px;
                border: none;
                margin-top: 0;
            }
            
            /* Dropdown styles */
            .Select-control { 
                background: rgba(30, 41, 59, 0.9) !important; 
                border: 1px solid rgba(99, 102, 241, 0.3) !important; 
                border-radius: 10px !important;
                min-height: 42px !important;
            }
            .Select-control:hover {
                border-color: rgba(99, 102, 241, 0.5) !important;
            }
            .Select-value-label, .Select-placeholder, .Select-value { color: #f1f5f9 !important; }
            .Select-menu-outer { 
                background: #1e293b !important; 
                border: 1px solid rgba(99, 102, 241, 0.3) !important;
                border-radius: 10px !important;
                margin-top: 4px !important;
            }
            .Select-option { 
                background: #1e293b !important; 
                color: #f1f5f9 !important;
                padding: 10px 14px !important;
            }
            .Select-option:hover, .Select-option.is-focused { 
                background: rgba(99, 102, 241, 0.25) !important; 
            }
            .Select-option.is-selected { 
                background: rgba(99, 102, 241, 0.4) !important; 
            }
            
            /* Button styles */
            .btn {
                border-radius: 10px !important;
                font-weight: 600 !important;
                font-size: 14px !important;
                transition: all 0.3s ease !important;
                border: none !important;
                display: flex !important;
                align-items: center !important;
                justify-content: center !important;
            }
            
            .btn-success {
                background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
            }
            
            .btn-success:hover {
                transform: translateY(-2px);
                box-shadow: 0 4px 12px rgba(16, 185, 129, 0.3) !important;
            }
            
            .btn-primary {
                background: linear-gradient(135deg, #6366f1 0%, #4f46e5 100%) !important;
            }
            
            .btn-primary:hover {
                transform: translateY(-2px);
                box-shadow: 0 4px 12px rgba(99, 102, 241, 0.3) !important;
            }
            
            .btn-secondary {
                background: rgba(51, 65, 85, 0.8) !important;
                border: 1px solid rgba(148, 163, 184, 0.2) !important;
            }
            
            .btn-secondary:hover {
                background: rgba(71, 85, 105, 0.9) !important;
                transform: translateY(-2px);
            }
            
            .dashboard-footer {
                background: transparent;
                padding: 0 0 20px 0;
            }
            
            @keyframes pulse {
                0%, 100% { opacity: 1; }
                50% { opacity: 0.7; }
            }
            
            .pulse-animation {
                animation: pulse 2s ease-in-out infinite;
            }
            
            /* Scrollbar */
            ::-webkit-scrollbar {
                width: 6px;
                height: 6px;
            }
            ::-webkit-scrollbar-track {
                background: #0f172a;
            }
            ::-webkit-scrollbar-thumb {
                background: #334155;
                border-radius: 3px;
            }
            ::-webkit-scrollbar-thumb:hover {
                background: #475569;
            }
            
            /* Graph containers */
            .js-plotly-plot {
                border-radius: 12px;
                overflow: hidden;
            }
            
            /* Section titles */
            h5 {
                color: #f1f5f9 !important;
                -webkit-text-fill-color: #f1f5f9 !important;
            }
            
            h6 {
                color: #cbd5e1 !important;
                -webkit-text-fill-color: #cbd5e1 !important;
            }
            
            /* ========== RESPONSIVE MOBILE ========== */
            @media (max-width: 768px) {
                /* Header */
                h1 {
                    font-size: 16px !important;
                }
                
                /* KPI Cards */
                .kpi-card {
                    min-height: 180px !important;
                    margin-bottom: 12px;
                }
                
                .kpi-card h2 {
                    font-size: 22px !important;
                }
                
                .kpi-card h6 {
                    font-size: 10px !important;
                }
                
                /* Tabs */
                .nav-tabs .nav-link {
                    padding: 8px 12px !important;
                    font-size: 12px !important;
                }
                
                .nav-tabs {
                    flex-wrap: nowrap;
                    overflow-x: auto;
                    -webkit-overflow-scrolling: touch;
                }
                
                /* Tab content */
                .tab-content {
                    padding: 12px 8px !important;
                }
                
                /* Graphs */
                .js-plotly-plot {
                    margin-bottom: 16px;
                }
                
                /* Footer */
                .dashboard-footer {
                    padding-left: 12px !important;
                    padding-right: 12px !important;
                }
                
                .dashboard-footer p {
                    font-size: 11px !important;
                }
                
                .dashboard-footer > div > div {
                    flex-direction: column !important;
                    text-align: center !important;
                }
                
                .dashboard-footer > div > div > div {
                    text-align: center !important;
                }
                
                /* Filtros */
                .filtros-section {
                    padding: 12px !important;
                }
                
                /* Buttons */
                .btn {
                    padding: 6px 12px !important;
                    font-size: 12px !important;
                }
                
                /* Container padding */
                .container-fluid {
                    padding-left: 8px !important;
                    padding-right: 8px !important;
                }
            }
            
            @media (max-width: 480px) {
                h1 {
                    font-size: 14px !important;
                }
                
                .kpi-card h2 {
                    font-size: 20px !important;
                }
                
                .nav-tabs .nav-link {
                    padding: 6px 10px !important;
                    font-size: 11px !important;
                }
                
                .logo-img {
                    height: 40px !important;
                }
                
                .header-subtitle {
                    display: none;
                }
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
                    html.Img(src='/assets/Logo.png', style={
                        'height': '70px', 
                        'filter': 'drop-shadow(0 4px 8px rgba(0,0,0,0.3))',
                        'maxHeight': '50px'
                    }, className="logo-img"),
                    html.Div([
                        html.H1("DASHBOARD UCI ADULTOS 2025", style={
                            'color': '#f8fafc', 'fontWeight': '700', 'marginBottom': '4px',
                            'fontSize': '24px', 'letterSpacing': '-0.3px', 'margin': '0'
                        }),
                        html.P("Sistema de Monitoreo y Análisis de Indicadores", style={
                            'color': '#64748b', 'fontSize': '13px', 'marginBottom': '0', 'fontWeight': '500'
                        }, className="header-subtitle")
                    ], style={'marginLeft': '20px'})
                ], style={'display': 'flex', 'alignItems': 'center', 'marginBottom': '20px', 'flexWrap': 'wrap', 'gap': '10px'})
            ]),
            html.Div([
                dbc.Button("📤 Exportar Excel", id="btn-export-excel", color="success", size="sm", className="me-2"),
                dbc.Button("📄 Exportar CSV", id="btn-export-csv", color="primary", size="sm"),
                dcc.Download(id="download-excel"),
                dcc.Download(id="download-csv"),
            ], style={'marginBottom': '24px', 'display': 'flex', 'gap': '12px', 'flexWrap': 'wrap'})
        ], className="fade-in", style={'padding': '0 24px'}),
        
        # Tabs
        dbc.Tabs(id="tabs", active_tab='tab-resumen', children=[
            dbc.Tab(label='📊 Resumen General', tab_id='tab-resumen'),
            dbc.Tab(label='⚙️ Gestión y Calidad', tab_id='tab-gestion'),
            dbc.Tab(label='🏥 Especialidades', tab_id='tab-especialidades'),
            dbc.Tab(label='📋 Mortalidad', tab_id='tab-mortalidad'),
        ], style={'marginBottom': '0', 'borderBottom': 'none', 'padding': '0 24px'}),
        
        html.Div(id='tab-content', className="tab-content"),
        
        # Footer
        html.Div([
            html.Hr(style={
                'borderColor': '#334155',
                'borderWidth': '1px',
                'margin': '0 0 20px 0',
                'opacity': '0.5'
            }),
            html.Div([
                html.Div([
                    html.P("Hospital Infantil Universitario de San José", style={
                        'color': '#f1f5f9', 'fontWeight': '600', 'margin': '0', 'fontSize': '14px'
                    }),
                    html.P("Unidad de Cuidados Intensivos - Adultos", style={
                        'color': '#64748b', 'margin': '0', 'fontSize': '12px'
                    })
                ]),
                html.Div([
                    html.P("© 2025 - Dashboard de Indicadores UCI", style={
                        'color': '#475569', 'margin': '0', 'fontSize': '12px'
                    })
                ], style={'textAlign': 'right'})
            ], style={'display': 'flex', 'justifyContent': 'space-between', 'alignItems': 'center', 'flexWrap': 'wrap', 'gap': '16px'})
        ], className="dashboard-footer", style={'marginTop': '40px', 'paddingBottom': '20px', 'paddingLeft': '24px', 'paddingRight': '24px'})
        
    ], fluid=True, style={'maxWidth': '1440px', 'padding': '32px 0'}, id='main-container')
], style={
    'background': 'linear-gradient(135deg, #0f172a 0%, #1e1b4b 50%, #0f172a 100%)',
    'minHeight': '100vh', 
    'fontFamily': "'Inter', -apple-system, BlinkMacSystemFont, sans-serif"
}, id='app-container')

@app.callback(Output('tab-content', 'children'), Input('tabs', 'active_tab'))
def render_tab_content(tab):
    if tab == 'tab-resumen':
        try:
            if len(df.columns) == 0:
                return html.Div([html.H3("⚠️ No hay datos", style={'color': '#ffffff'})])
            
            filtros = html.Div([
                html.Div([
                    html.Div([
                        html.Span("📅", style={'fontSize': '16px', 'marginRight': '8px'}),
                        html.Span("Filtrar por Meses", style={
                            'color': '#f1f5f9', 'fontWeight': '600', 'fontSize': '14px'
                        })
                    ], style={'display': 'flex', 'alignItems': 'center', 'marginBottom': '12px'}),
                    html.Div([
                        html.Div([
                            dcc.Dropdown(
                                id='filtro-mes-inicio',
                                options=[{'label': meses_completos[mes], 'value': mes} for mes in meses_disponibles],
                                value=meses_disponibles,
                                multi=True,
                                placeholder="Selecciona meses...",
                                style={'fontSize': '13px', 'minWidth': '400px'}
                            )
                        ]),
                        html.Div([
                            dbc.Button("✓ Aplicar", id='btn-aplicar-filtros', color="success", size="sm",
                                      style={'fontSize': '12px', 'padding': '8px 16px', 'fontWeight': '600'}),
                            dbc.Button("Todos", id='btn-todos-meses', color="secondary", size="sm",
                                      style={'fontSize': '12px', 'padding': '8px 16px', 'fontWeight': '600', 'marginLeft': '8px'})
                        ], style={'marginLeft': '12px', 'display': 'flex', 'alignItems': 'center'})
                    ], style={'display': 'flex', 'alignItems': 'center', 'flexWrap': 'wrap', 'gap': '8px'}),
                ])
            ], style={
                'background': 'linear-gradient(135deg, rgba(30,41,59,0.95) 0%, rgba(15,23,42,0.95) 100%)',
                'padding': '16px 20px', 'borderRadius': '12px', 'marginBottom': '20px',
                'border': '1px solid rgba(99,102,241,0.25)',
                'boxShadow': '0 2px 10px rgba(0,0,0,0.15)'
            }, className="filtros-section")
            
            return html.Div([filtros, html.Div(id='resumen-content')])
        except Exception as e:
            return html.Div([html.H3(f"Error: {str(e)}", style={'color': '#e74c3c'})])

    elif tab == 'tab-mortalidad':
        try:
            suma_col = [c for c in df.columns if str(c).strip() in ['∑', 'SUMA']]
            suma_col = suma_col[0] if suma_col else '∑'
            
            # Usar el valor de la hoja ESTADISTICA (dato oficial)
            total_fallecidos = safe_get_row('FALLECIDOS', suma_col)
            if pd.isna(total_fallecidos): total_fallecidos = len(df_fallecidos) if not df_fallecidos.empty else 0
            fallecidos_covid = safe_get_row('FALLECIDOS COVID-19', suma_col)
            total_egresos = safe_get_row('NÚMERO DE EGRESOS DEL MES', suma_col)
            
            if pd.isna(fallecidos_covid): fallecidos_covid = 0
            if pd.isna(total_egresos) or total_egresos == 0: total_egresos = 1
            
            # Usar el promedio de mortalidad del Excel
            mortalidad_general = safe_get_row('% DE MORTALIDAD GENERAL', 'PROMEDIO', silent=True)
            if pd.isna(mortalidad_general): mortalidad_general = (total_fallecidos / total_egresos * 100)
            
            # Usar el promedio de mortalidad trauma del Excel
            mortalidad_trauma = safe_get_row('% mortalidad en trauma', 'PROMEDIO', silent=True)
            if pd.isna(mortalidad_trauma): mortalidad_trauma = 0
            
            # Calcular estadísticas de días de estancia de fallecidos
            dias_estancia_validos = df_fallecidos['Dias_Estancia'].dropna() if not df_fallecidos.empty else pd.Series([0])
            if len(dias_estancia_validos) == 0:
                dias_estancia_validos = pd.Series([0])
            
            promedio_dias = dias_estancia_validos.mean()
            minimo_dias = dias_estancia_validos.min()
            maximo_dias = dias_estancia_validos.max()
            mediana_dias = dias_estancia_validos.median()
            
            # KPIs principales
            kpi1 = create_kpi_card("Total Fallecidos", f"{int(total_fallecidos)}", "Ene - Dic", "#e74c3c", "🕊️")
            kpi2 = create_kpi_card("% Mortalidad General", formato_porcentaje(mortalidad_general), "Ene - Dic", "#c0392b", "⚰️")
            kpi3 = create_kpi_card("Fallecidos COVID-19", f"{int(fallecidos_covid)}", "Ene - Dic", "#34495e", "🦠")
            kpi4 = create_kpi_card("% Mortalidad Trauma", formato_porcentaje(mortalidad_trauma), "Promedio", "#f39c12", "🚑")
            
            # KPIs de días de estancia fallecidos
            kpi_prom_dias = create_kpi_card("Promedio Días Estancia", f"{promedio_dias:.1f}", "Fallecidos", "#9b59b6", "📊")
            kpi_mediana_dias = create_kpi_card("Mediana Días de Estancia", f"{mediana_dias:.1f}", "Fallecidos", "#3498db", "📈")
            
            fig_mortalidad = create_line_chart('% DE MORTALIDAD GENERAL', '% Mortalidad General', color='#e74c3c', decimals=1)
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
                    title={'text': 'Pirámide de Edad y Sexo - Fallecidos', 'font': {'size': 14, 'color': '#f1f5f9'}, 'x': 0.5},
                    barmode='relative', 
                    plot_bgcolor='rgba(15,23,42,0.6)',
                    paper_bgcolor='rgba(30,41,59,0.8)',
                    xaxis={'title': '', 'gridcolor': 'rgba(148,163,184,0.15)', 'color': '#94a3b8',
                           'range': [-(max_val + 2), max_val + 2], 'showticklabels': False, 'zeroline': False},
                    yaxis={'title': '', 'gridcolor': 'rgba(148,163,184,0.15)', 'color': '#94a3b8'},
                    legend=dict(orientation="h", y=-0.15, x=0.5, xanchor="center", bgcolor='rgba(15,23,42,0.9)', 
                               font={'color': '#f1f5f9', 'size': 11}, bordercolor='#334155', borderwidth=1),
                    hoverlabel=dict(bgcolor='#1e293b', font_color='#f1f5f9'),
                    margin=dict(l=60, r=30, t=50, b=60),
                    height=400
                )
            
            return html.Div([
                dbc.Row([dbc.Col(kpi1, xs=12, md=3), dbc.Col(kpi2, xs=12, md=3), dbc.Col(kpi3, xs=12, md=3), dbc.Col(kpi4, xs=12, md=3)], className="mb-4"),
                html.H6("Días de Estancia - Fallecidos", style={'color': '#e2e8f0', 'fontWeight': '600', 'marginBottom': '15px', 'marginTop': '10px'}),
                dbc.Row([dbc.Col(kpi_prom_dias, xs=12, md=6), dbc.Col(kpi_mediana_dias, xs=12, md=6)], className="mb-4"),
                dbc.Row([
                    dbc.Col(xs=12, md=2),
                    dbc.Col(dcc.Graph(figure=fig_mortalidad, config=GRAPH_CONFIG), xs=12, md=8),
                    dbc.Col(xs=12, md=2)
                ], className="mb-4", justify="center"),
                dbc.Row([dbc.Col(dcc.Graph(figure=fig_trauma, config=GRAPH_CONFIG), xs=12, md=6), dbc.Col(dcc.Graph(figure=fig_edad_sexo, config=GRAPH_CONFIG), xs=12, md=6)], className="mt-3")
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
            total_neurologia = safe_get_row('NEUROLOGIA', suma_col, silent=True)
            
            if pd.isna(total_medicina): total_medicina = 0
            if pd.isna(total_cirugia): total_cirugia = 0
            if pd.isna(total_neuro): total_neuro = 0
            if pd.isna(total_neurologia): total_neurologia = 0
            
            # Calcular total de especialidades y porcentajes
            total_especialidades = total_medicina + total_cirugia + total_neuro + total_neurologia
            if total_especialidades == 0: total_especialidades = 1
            
            pct_medicina = (total_medicina / total_especialidades * 100)
            pct_cirugia = (total_cirugia / total_especialidades * 100)
            pct_neuro = (total_neuro / total_especialidades * 100)
            pct_neurologia = (total_neurologia / total_especialidades * 100)
            
            kpi1 = create_kpi_card("Medicina Interna", f"{int(total_medicina)}", f"{formato_porcentaje(pct_medicina, False)}% del total", "#3498db", "🩺")
            kpi2 = create_kpi_card("Cirugía General", f"{int(total_cirugia)}", f"{formato_porcentaje(pct_cirugia, False)}% del total", "#27ae60", "💉")
            kpi3 = create_kpi_card("Neurocirugía", f"{int(total_neuro)}", f"{formato_porcentaje(pct_neuro, False)}% del total", "#16a085", "🧠")
            kpi4 = create_kpi_card("Neurología", f"{int(total_neurologia)}", f"{formato_porcentaje(pct_neurologia, False)}% del total", "#9b59b6", "🧬")
            
            fig_medicina = create_bar_chart('MEDICINA INTERNA', 'Medicina Interna', color='#3498db', silent=True, decimals=0)
            fig_cirugia = create_bar_chart('CIRUGIA GENERAL', 'Cirugía General', color='#27ae60', silent=True, decimals=0)
            fig_neuro = create_bar_chart('NEUROCIRUGIA', 'Neurocirugía', color='#16a085', silent=True, decimals=0)
            fig_neurologia = create_bar_chart('NEUROLOGIA', 'Neurología', color='#9b59b6', silent=True, decimals=0)
            
            return html.Div([
                html.H5("Principales Especialidades", style={'color': '#ffffff', 'fontWeight': '600', 'marginBottom': '20px'}),
                dbc.Row([dbc.Col(kpi1, xs=12, md=3), dbc.Col(kpi2, xs=12, md=3), dbc.Col(kpi3, xs=12, md=3), dbc.Col(kpi4, xs=12, md=3)], className="mb-4"),
                dbc.Row([dbc.Col(dcc.Graph(figure=fig_medicina, config=GRAPH_CONFIG), xs=12, md=6), dbc.Col(dcc.Graph(figure=fig_cirugia, config=GRAPH_CONFIG), xs=12, md=6)], className="mb-4"),
                dbc.Row([dbc.Col(dcc.Graph(figure=fig_neuro, config=GRAPH_CONFIG), xs=12, md=6), dbc.Col(dcc.Graph(figure=fig_neurologia, config=GRAPH_CONFIG), xs=12, md=6)])
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
            
            # Calcular promedio general de días VM
            promedio_vm_validos = [v for v in promedio_vm if v > 0]
            promedio_vm_general = sum(promedio_vm_validos) / len(promedio_vm_validos) if promedio_vm_validos else 0
            
            kpi1 = create_kpi_card("Estancia Promedio", f"{promedio_estancia:.2f}", "Días", "#2c3e50", "⏱️")
            kpi2 = create_kpi_card("Giro de Cama", f"{giro_cama:.2f}", "Promedio", "#27ae60", "🔄")
            kpi3 = create_kpi_card("Total Días Estancia", f"{int(total_dias)}", "Ene - Dic", "#16a085", "📅")
            kpi4 = create_kpi_card("Capacidad UCI", f"{int(capacidad)}", "Camas", "#f39c12", "🛏️")
            kpi5 = create_kpi_card("Promedio Días VM", f"{promedio_vm_general:.2f}", "Ventilación Mecánica", "#9b59b6", "🫁")
            
            fig_estancia = create_bar_chart('PROMEDIO ESTANCIA', 'Estancia Promedio (Días)', color='#16a085', decimals=1)
            fig_giro = create_bar_chart('GIRO CAMA', 'Giro de Cama', color='#27ae60', decimals=2)
            
            fig_promedio_vm = go.Figure()
            text_vm = [f'{v:.1f}' if pd.notna(v) else '' for v in promedio_vm]
            fig_promedio_vm.add_trace(go.Bar(
                x=meses_disponibles, y=promedio_vm,
                marker_color='#9b59b6',
                marker_line_color='rgba(255,255,255,0.2)',
                marker_line_width=1,
                text=text_vm,
                texttemplate='%{text}',
                textposition='inside',
                insidetextanchor='middle',
                textfont=dict(color='white', size=12, family='Inter'),
                hovertemplate='<b>%{x}</b><br>Promedio días VM: %{y:.1f}<extra></extra>'
            ))
            fig_promedio_vm.update_layout(
                title={'text': 'Promedio de Días de VM', 'font': {'size': 14, 'color': '#f1f5f9'}, 'x': 0.5},
                plot_bgcolor='rgba(15,23,42,0.6)', paper_bgcolor='rgba(30,41,59,0.8)',
                xaxis={'gridcolor': 'rgba(148,163,184,0.15)', 'showline': False, 'color': '#94a3b8', 'title': ''},
                yaxis={'gridcolor': 'rgba(148,163,184,0.15)', 'showline': False, 'color': '#94a3b8', 'title': ''},
                margin=dict(l=50, r=30, t=50, b=40),
                height=380,
                bargap=0.3,
                hoverlabel=dict(bgcolor='#1e293b', font_size=13, font_color='#f1f5f9')
            )
            
            fig_bacteriemia = create_line_chart('TASA DE BACTERIEMIA ASOC. A  C. CENTRAL', 'Tasa de ITS AC', color='#e74c3c', silent=True, decimals=1)
            fig_nav = create_bar_chart('TASA NEUMONÍA ASOCIADA A VENTILADOR', 'Tasa NAV', color='#e74c3c', silent=True, decimals=1)
            fig_istu = create_line_chart('TASA INFECCIÓN ASOCIADA A C. URINARIO', 'Tasa de ISTU AC', color='#f39c12', silent=True, decimals=1)
            fig_lesiones = create_bar_chart('TASA LESIONES POR PRESION', 'Tasa Lesiones por Presión', color='#9b59b6', silent=True, decimals=1)

            # Gráfico de Gérmenes Aislados (donut)
            fig_germenes = go.Figure()
            if not df_germenes.empty:
                # Donut para gérmenes
                fig_germenes = create_donut_chart(
                    labels=df_germenes['Germen'].tolist(),
                    values=df_germenes['Cantidad'].tolist(),
                    title='Gérmenes Aislados (IAAS)',
                    auto_sort=True
                )
            
            # Gráfico de Eventos Adversos - Barras horizontales
            fig_eventos = go.Figure()
            if not df_eventos_adversos.empty:
                # Paleta amplia de 25 colores únicos
                colors_eventos = [
                    '#ef4444', '#6366f1', '#10b981', '#f59e0b', '#8b5cf6', 
                    '#06b6d4', '#f97316', '#ec4899', '#22c55e', '#a855f7',
                    '#14b8a6', '#eab308', '#3b82f6', '#84cc16', '#e11d48',
                    '#0ea5e9', '#d946ef', '#facc15', '#fb7185', '#4ade80',
                    '#38bdf8', '#c084fc', '#fbbf24', '#2dd4bf', '#f472b6'
                ]
                
                # Invertir orden para que el mayor quede arriba
                df_eventos_plot = df_eventos_adversos.iloc[::-1].reset_index(drop=True)
                n_eventos = len(df_eventos_plot)
                
                # Asignar colores únicos a cada barra
                bar_colors = [colors_eventos[i % len(colors_eventos)] for i in range(n_eventos)][::-1]
                
                fig_eventos = go.Figure(data=[go.Bar(
                    y=df_eventos_plot['Evento'],
                    x=df_eventos_plot['Cantidad'],
                    orientation='h',
                    marker_color=bar_colors,
                    text=df_eventos_plot['Cantidad'],
                    textposition='outside',
                    textfont=dict(color='#f1f5f9', size=12),
                    hovertemplate='<b>%{y}</b><br>Casos: %{x}<extra></extra>'
                )])
            
            max_val = df_eventos_adversos['Cantidad'].max() if not df_eventos_adversos.empty else 10
            fig_eventos.update_layout(
                title={'text': 'Eventos Adversos por Categoría', 'font': {'size': 14, 'color': '#f1f5f9'}, 'x': 0.5},
                plot_bgcolor='rgba(15,23,42,0.6)', paper_bgcolor='rgba(30,41,59,0.8)',
                xaxis={'gridcolor': 'rgba(148,163,184,0.15)', 'color': '#94a3b8', 'dtick': 2, 'range': [0, max_val + 3], 'title': ''},
                yaxis={'gridcolor': 'rgba(148,163,184,0.1)', 'color': '#f1f5f9', 'tickfont': {'size': 10}, 'title': '', 'ticksuffix': '  '},
                margin=dict(l=280, r=50, t=50, b=30),
                height=500,
                hoverlabel=dict(bgcolor='#1e293b', font_color='#f1f5f9', font_size=12),
                bargap=0.2
            )

            # Gráfico GRAM-SOT (Donaciones/Trasplantes)
            fig_gram_sot = go.Figure()
            total_gram_sot = 0
            if not df_gram_sot.empty:
                gram_values = df_gram_sot.iloc[0].values.tolist()
                total_gram_sot = sum([v for v in gram_values if pd.notna(v)])
                fig_gram_sot = go.Figure(data=[go.Bar(
                    x=meses_disponibles, y=gram_values,
                    marker_color='#8b5cf6',
                    hovertemplate='<b>%{x}</b><br>Casos: %{y}<extra></extra>'
                )])
                fig_gram_sot.update_layout(
                    title={'text': f'GRAM - SOT (Total: {int(total_gram_sot)})', 'font': {'size': 14, 'color': '#f1f5f9'}, 'x': 0.5},
                    plot_bgcolor='rgba(15,23,42,0.6)', paper_bgcolor='rgba(30,41,59,0.8)',
                    xaxis={'gridcolor': 'rgba(148,163,184,0.15)', 'color': '#94a3b8', 'title': ''},
                    yaxis={'gridcolor': 'rgba(148,163,184,0.15)', 'color': '#94a3b8', 'dtick': 1, 'title': ''},
                    margin=dict(l=50, r=30, t=50, b=40),
                    height=380,
                    hoverlabel=dict(bgcolor='#1e293b', font_color='#f1f5f9')
                )

            return html.Div([
                html.H5("Indicadores de Gestión", style={'color': '#ffffff', 'fontWeight': '600'}, className="mb-4"),
                dbc.Row([dbc.Col(kpi1, xs=12, md=True), dbc.Col(kpi2, xs=12, md=True), dbc.Col(kpi3, xs=12, md=True), dbc.Col(kpi4, xs=12, md=True), dbc.Col(kpi5, xs=12, md=True)], className="mb-5 g-4"),
                dbc.Row([dbc.Col(dcc.Graph(figure=fig_estancia, config=GRAPH_CONFIG), xs=12, md=6), dbc.Col(dcc.Graph(figure=fig_giro, config=GRAPH_CONFIG), xs=12, md=6)], className="mb-4 g-4"),
                dbc.Row([
                    dbc.Col(xs=12, md=3),
                    dbc.Col(dcc.Graph(figure=fig_promedio_vm, config=GRAPH_CONFIG), xs=12, md=6),
                    dbc.Col(xs=12, md=3)
                ], className="mb-5", justify="center"),
                html.H5("Indicadores de Calidad e Infecciones", style={'color': '#ffffff', 'fontWeight': '600', 'marginTop': '30px', 'marginBottom': '30px', 'textAlign': 'center'}),
                dbc.Row([dbc.Col(dcc.Graph(figure=fig_bacteriemia, config=GRAPH_CONFIG), xs=12, md=6), dbc.Col(dcc.Graph(figure=fig_nav, config=GRAPH_CONFIG), xs=12, md=6)], className="mb-4 g-4"),
                dbc.Row([dbc.Col(dcc.Graph(figure=fig_istu, config=GRAPH_CONFIG), xs=12, md=6), dbc.Col(dcc.Graph(figure=fig_lesiones, config=GRAPH_CONFIG), xs=12, md=6)], className="mb-5 g-4"),
                html.H5("Gérmenes Aislados y Eventos Adversos", style={'color': '#ffffff', 'fontWeight': '600', 'marginTop': '30px', 'marginBottom': '30px', 'textAlign': 'center'}),
                dbc.Row([
                    dbc.Col(xs=12, md=2),
                    dbc.Col(dcc.Graph(figure=fig_germenes, config=GRAPH_CONFIG), xs=12, md=8),
                    dbc.Col(xs=12, md=2)
                ], className="mb-5", justify="center"),
                dbc.Row([
                    dbc.Col(dcc.Graph(figure=fig_eventos, config=GRAPH_CONFIG), xs=12)
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
            rango_texto = "Ene - Dic"
        else:
            rango_texto = ", ".join([meses_completos.get(m, m)[:3] for m in meses_filtrados])
        
        suma_col = [c for c in df.columns if str(c).strip() in ['∑', 'SUMA']]
        suma_col = suma_col[0] if suma_col else '∑'
        
        datos_ocupacional = safe_get_row('% OCUPACIONAL', meses_filtrados)
        promedio_ocupacional = datos_ocupacional[datos_ocupacional > 0].mean() if len(datos_ocupacional[datos_ocupacional > 0]) > 0 else 0
        datos_ingresos_kpi = safe_get_row('NUMERO DE INGRESOS DEL MES', meses_filtrados)
        total_ingresos = datos_ingresos_kpi.sum() if isinstance(datos_ingresos_kpi, pd.Series) else datos_ingresos_kpi
        # Usar datos de fallecidos: si son todos los meses, usar ESTADISTICA (valor oficial)
        if len(meses_filtrados) == len(meses_disponibles):
            total_fallecidos = safe_get_row('FALLECIDOS', suma_col)
            if pd.isna(total_fallecidos): total_fallecidos = 0
        elif not df_fallecidos.empty:
            total_fallecidos = len(df_fallecidos[df_fallecidos['Mes'].isin(meses_filtrados)])
        else:
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
        
        # Usar el promedio de mortalidad del Excel
        mortalidad = safe_get_row('% DE MORTALIDAD GENERAL', 'PROMEDIO', silent=True)
        if pd.isna(mortalidad): mortalidad = (total_fallecidos / total_egresos * 100)
        
        # KPIs
        kpi_ingresos = create_kpi_card("Total Ingresos", f"{int(total_ingresos)}", rango_texto, "#3498db", "📥")
        kpi1 = create_kpi_card("% Ocupación Promedio", formato_porcentaje(promedio_ocupacional), rango_texto, "#2c3e50", "🏥")
        kpi2 = create_kpi_card("% Mortalidad General", formato_porcentaje(mortalidad), f"Total {rango_texto}", "#e74c3c", "💔")
        kpi3 = create_kpi_card("Estancia Promedio", f"{promedio_estancia:.2f}", f"Días - {rango_texto}", "#27ae60", "⏱️")
        kpi4 = create_kpi_card("Giro de Cama", f"{giro_cama:.2f}", f"Promedio {rango_texto}", "#f39c12", "🔄")
        
        # EPS - Todas las EPS individuales
        total_famisanar = safe_get_row('FAMISANAR', suma_col, silent=True)
        total_compensar = safe_get_row('COMPENSAR', suma_col, silent=True)
        total_salud_total = safe_get_row('SALUD TOTAL', suma_col, silent=True)
        total_nueva_eps = safe_get_row('NUEVA  EPS', suma_col, silent=True)
        total_capital_salud = safe_get_row('CAPITAL SALUD ', suma_col, silent=True)
        total_sanitas = safe_get_row('SANITAS', suma_col, silent=True)
        total_emcosalud = safe_get_row('EMCOSALUD', suma_col, silent=True)
        total_sura = safe_get_row('SURA', suma_col, silent=True)
        total_aliansalud = safe_get_row('ALIANSALUD', suma_col, silent=True)
        total_emssanar = safe_get_row('EMSSANAR', suma_col, silent=True)
        
        # Limpiar valores nulos
        if pd.isna(total_famisanar): total_famisanar = 0
        if pd.isna(total_compensar): total_compensar = 0
        if pd.isna(total_salud_total): total_salud_total = 0
        if pd.isna(total_nueva_eps): total_nueva_eps = 0
        if pd.isna(total_capital_salud): total_capital_salud = 0
        if pd.isna(total_sanitas): total_sanitas = 0
        if pd.isna(total_emcosalud): total_emcosalud = 0
        if pd.isna(total_sura): total_sura = 0
        if pd.isna(total_aliansalud): total_aliansalud = 0
        if pd.isna(total_emssanar): total_emssanar = 0
        
        # Crear lista de EPS con sus valores (solo las que tienen datos > 0)
        eps_data = [
            ('Famisanar', total_famisanar),
            ('Compensar', total_compensar),
            ('Salud Total', total_salud_total),
            ('Nueva EPS', total_nueva_eps),
            ('Capital Salud', total_capital_salud),
            ('Sanitas', total_sanitas),
            ('Emcosalud', total_emcosalud),
            ('Sura', total_sura),
            ('Aliansalud', total_aliansalud),
            ('Emssanar', total_emssanar),
        ]
        
        # Filtrar EPS con valores > 0 y ordenar de mayor a menor
        eps_data = [(nombre, valor) for nombre, valor in eps_data if valor > 0]
        eps_data.sort(key=lambda x: x[1], reverse=True)
        
        eps_labels = [e[0] for e in eps_data]
        eps_values = [e[1] for e in eps_data]
        
        # Donut de distribución EPS - Ordenado de mayor a menor
        fig_eps_donut = create_donut_chart(
            labels=eps_labels,
            values=eps_values,
            title='Distribución por EPS',
            auto_sort=False  # Ya está ordenado
        )

        # Gráficos
        datos_ocup = safe_get_row('% OCUPACIONAL', meses_filtrados)
        fig_ocupacional = go.Figure()
        fig_ocupacional.add_trace(go.Scatter(
            x=meses_filtrados, 
            y=datos_ocup,
            mode='lines+markers+text',
            line=dict(color='#10b981', width=2),
            marker=dict(size=10, color='#10b981', line=dict(width=2, color='#0f172a')),
            text=[f'{v:.1f}' if pd.notna(v) else '' for v in datos_ocup],
            textposition='top center',
            textfont=dict(color='#f1f5f9', size=11),
            hovertemplate='<b>%{x}</b><br>Ocupación: %{y:.1f}%<extra></extra>'
        ))
        fig_ocupacional.update_layout(
            title={'text': '% Ocupacional Mensual', 'font': {'size': 14, 'color': '#f1f5f9'}, 'x': 0.5},
            showlegend=False, 
            plot_bgcolor='rgba(15,23,42,0.6)', 
            paper_bgcolor='rgba(30,41,59,0.8)',
            xaxis={'gridcolor': 'rgba(148,163,184,0.15)', 'color': '#94a3b8', 'title': ''},
            yaxis={'gridcolor': 'rgba(148,163,184,0.15)', 'color': '#94a3b8', 'title': ''}, 
            margin=dict(l=50, r=30, t=50, b=40),
            height=380,
            hoverlabel=dict(bgcolor='#1e293b', font_color='#f1f5f9')
        )
        
        datos_ingresos = safe_get_row('NUMERO DE INGRESOS DEL MES', meses_filtrados)
        datos_egresos_chart = safe_get_row('NÚMERO DE EGRESOS DEL MES', meses_filtrados)
        fig_ingresos_egresos = go.Figure()
        fig_ingresos_egresos.add_trace(go.Bar(name='Ingresos', x=meses_filtrados, y=datos_ingresos, marker_color='#6366f1',
            text=datos_ingresos, texttemplate='%{y:.0f}', textposition='inside', textfont=dict(color='white', size=9)))
        fig_ingresos_egresos.add_trace(go.Bar(name='Egresos', x=meses_filtrados, y=datos_egresos_chart, marker_color='#10b981',
            text=datos_egresos_chart, texttemplate='%{y:.0f}', textposition='inside', textfont=dict(color='white', size=9)))
        fig_ingresos_egresos.update_layout(title={'text': 'Ingresos vs. Egresos', 'font': {'size': 14, 'color': '#f1f5f9'}, 'x': 0.5},
            barmode='group', plot_bgcolor='rgba(15,23,42,0.6)', paper_bgcolor='rgba(30,41,59,0.8)', 
            xaxis={'gridcolor': 'rgba(148,163,184,0.15)', 'color': '#94a3b8', 'title': ''},
            yaxis={'gridcolor': 'rgba(148,163,184,0.15)', 'color': '#94a3b8', 'title': ''}, 
            margin=dict(l=50, r=30, t=50, b=70),
            height=400,
            legend=dict(orientation="h", y=-0.18, x=0.5, xanchor="center", 
            bgcolor='rgba(15,23,42,0.9)', font={'color': '#f1f5f9', 'size': 11}, bordercolor='#334155', borderwidth=1), 
            hoverlabel=dict(bgcolor='#1e293b', font_color='#f1f5f9'))

        # Gráficos de Morbilidad
        fig_morbilidad_total = go.Figure()
        fig_morbilidad_genero = go.Figure()
        if not df_morbilidad.empty:
            # Donut chart para morbilidad por diagnóstico - auto_sort=True ordena de mayor a menor
            fig_morbilidad_total = create_donut_chart(
                labels=df_morbilidad.index.tolist(),
                values=df_morbilidad['Total'].tolist(),
                title='Morbilidad por Diagnóstico',
                auto_sort=True
            )
            
            fig_morbilidad_genero = go.Figure(data=[
                go.Bar(name='Hombres', x=df_morbilidad.index, y=df_morbilidad['Hombre'], marker_color='#3b82f6',
                    text=df_morbilidad['Hombre'], texttemplate='%{y:.0f}', textposition='inside', textfont=dict(color='white', size=9)),
                go.Bar(name='Mujeres', x=df_morbilidad.index, y=df_morbilidad['Mujer'], marker_color='#ec4899',
                    text=df_morbilidad['Mujer'], texttemplate='%{y:.0f}', textposition='inside', textfont=dict(color='white', size=9))])
            fig_morbilidad_genero.update_layout(title={'text': 'Morbilidad por Sexo', 'font': {'size': 14, 'color': '#f1f5f9'}, 'x': 0.5},
                barmode='group', plot_bgcolor='rgba(15,23,42,0.6)', paper_bgcolor='rgba(30,41,59,0.8)', 
                xaxis={'gridcolor': 'rgba(148,163,184,0.15)', 'color': '#94a3b8', 'tickangle': -45, 'title': ''},
                yaxis={'gridcolor': 'rgba(148,163,184,0.15)', 'color': '#94a3b8', 'title': ''}, 
                margin=dict(l=50, r=30, t=50, b=130),
                height=420,
                legend=dict(orientation="h", y=-0.45, x=0.5, xanchor="center",
                bgcolor='rgba(15,23,42,0.9)', font={'color': '#f1f5f9', 'size': 11}, bordercolor='#334155', borderwidth=1), 
                hoverlabel=dict(bgcolor='#1e293b', font_color='#f1f5f9'))

        return meses_filtrados, html.Div([
            dbc.Row([dbc.Col(kpi_ingresos, xs=12, sm=6, md=True, className="mb-3"), dbc.Col(kpi1, xs=12, sm=6, md=True, className="mb-3"),
                     dbc.Col(kpi2, xs=12, sm=6, md=True, className="mb-3"), dbc.Col(kpi3, xs=12, sm=6, md=True, className="mb-3"),
                     dbc.Col(kpi4, xs=12, sm=6, md=True, className="mb-3")], className="mb-3 g-3"),
            dbc.Row([dbc.Col(dcc.Graph(figure=fig_ocupacional, config=GRAPH_CONFIG), xs=12, md=6), dbc.Col(dcc.Graph(figure=fig_ingresos_egresos, config=GRAPH_CONFIG), xs=12, md=6)]),
            dbc.Row([dbc.Col(dcc.Graph(figure=fig_eps_donut, config=GRAPH_CONFIG), xs=12, md=6), dbc.Col(dcc.Graph(figure=fig_morbilidad_total, config=GRAPH_CONFIG), xs=12, md=6)], className="mt-3"),
            html.H6("Morbilidad por Sexo", style={'color': '#ffffff', 'fontWeight': '600', 'marginTop': '30px', 'marginBottom': '20px', 'textAlign': 'center'}),
            dbc.Row([
                dbc.Col(xs=12, md=2),
                dbc.Col(dcc.Graph(figure=fig_morbilidad_genero, config=GRAPH_CONFIG), xs=12, md=8),
                dbc.Col(xs=12, md=2)
            ])
        ])
    except Exception as e:
        return dash.no_update, html.Div([html.H3(f"Error: {str(e)}", style={'color': '#e74c3c'})])





# --- FUNCIÓN PARA GENERAR DATOS DE EXPORTACIÓN ---
def generar_datos_exportacion():
    """Genera un diccionario con todos los datos organizados para exportar"""
    datos = {}
    suma_col = [c for c in df.columns if str(c).strip() in ['∑', 'SUMA']]
    suma_col = suma_col[0] if suma_col else '∑'
    
    # 1. Resumen General
    resumen = []
    indicadores_resumen = [
        ('% OCUPACIONAL', '% Ocupación'),
        ('NUMERO DE INGRESOS DEL MES', 'Ingresos'),
        ('NÚMERO DE EGRESOS DEL MES', 'Egresos'),
        ('FALLECIDOS', 'Fallecidos'),
        ('PROMEDIO ESTANCIA', 'Estancia Promedio (días)'),
        ('GIRO CAMA', 'Giro de Cama'),
    ]
    for row_name, label in indicadores_resumen:
        fila = {'Indicador': label}
        data = safe_get_row(row_name, meses_disponibles, silent=True)
        for i, mes in enumerate(meses_disponibles):
            fila[meses_completos.get(mes, mes)] = data.iloc[i] if isinstance(data, pd.Series) else 0
        total = safe_get_row(row_name, suma_col, silent=True)
        fila['Total'] = total if not pd.isna(total) else 0
        resumen.append(fila)
    datos['Resumen General'] = pd.DataFrame(resumen)
    
    # 2. EPS
    eps_data = []
    eps_list = ['FAMISANAR', 'COMPENSAR', 'SALUD TOTAL', 'NUEVA  EPS', 'CAPITAL SALUD ', 'SANITAS', 'SURA', 'ALIANSALUD']
    for eps in eps_list:
        fila = {'EPS': eps.strip()}
        data = safe_get_row(eps, meses_disponibles, silent=True)
        for i, mes in enumerate(meses_disponibles):
            fila[meses_completos.get(mes, mes)] = data.iloc[i] if isinstance(data, pd.Series) else 0
        total = safe_get_row(eps, suma_col, silent=True)
        fila['Total'] = total if not pd.isna(total) else 0
        eps_data.append(fila)
    datos['EPS'] = pd.DataFrame(eps_data)
    
    # 3. Especialidades
    esp_data = []
    especialidades = ['MEDICINA INTERNA', 'CIRUGIA GENERAL', 'NEUROCIRUGIA', 'NEUROLOGIA', 'ORTOPEDIA', 'UROLOGIA', 'GINECOLOGIA']
    for esp in especialidades:
        fila = {'Especialidad': esp.title()}
        data = safe_get_row(esp, meses_disponibles, silent=True)
        for i, mes in enumerate(meses_disponibles):
            fila[meses_completos.get(mes, mes)] = data.iloc[i] if isinstance(data, pd.Series) else 0
        total = safe_get_row(esp, suma_col, silent=True)
        fila['Total'] = total if not pd.isna(total) else 0
        esp_data.append(fila)
    datos['Especialidades'] = pd.DataFrame(esp_data)
    
    # 4. Indicadores de Calidad
    calidad_data = []
    indicadores_calidad = [
        ('TASA DE BACTERIEMIA ASOC. A  C. CENTRAL', 'Tasa ITS AC'),
        ('TASA NEUMONÍA ASOCIADA A VENTILADOR', 'Tasa NAV'),
        ('TASA INFECCIÓN ASOCIADA A C. URINARIO', 'Tasa ISTU AC'),
        ('TASA LESIONES POR PRESION', 'Tasa Lesiones por Presión'),
    ]
    for row_name, label in indicadores_calidad:
        fila = {'Indicador': label}
        data = safe_get_row(row_name, meses_disponibles, silent=True)
        for i, mes in enumerate(meses_disponibles):
            fila[meses_completos.get(mes, mes)] = data.iloc[i] if isinstance(data, pd.Series) else 0
        calidad_data.append(fila)
    datos['Indicadores Calidad'] = pd.DataFrame(calidad_data)
    
    # 5. Mortalidad
    mort_data = []
    indicadores_mort = [
        ('FALLECIDOS', 'Total Fallecidos'),
        ('FALLECIDOS COVID-19', 'Fallecidos COVID-19'),
        ('% DE MORTALIDAD GENERAL', '% Mortalidad General'),
        ('Pacientes con diagnóstico de Trauma', 'Pacientes Trauma'),
        ('Pacientes fallecidos por Trauma', 'Fallecidos Trauma'),
        ('% mortalidad en trauma', '% Mortalidad Trauma'),
        ('Pacientes Obstétricas', 'Pacientes Obstétricas'),
        ('Pacientes fallecidas Ginecobstetricia', 'Fallecidas Gineco-Obstetricia'),
        ('%. Mortalidad Obstétrica', '% Mortalidad Obstétrica'),
    ]
    for row_name, label in indicadores_mort:
        fila = {'Indicador': label}
        data = safe_get_row(row_name, meses_disponibles, silent=True)
        for i, mes in enumerate(meses_disponibles):
            fila[meses_completos.get(mes, mes)] = data.iloc[i] if isinstance(data, pd.Series) else 0
        mort_data.append(fila)
    datos['Mortalidad'] = pd.DataFrame(mort_data)
    
    # 6. Morbilidad
    if not df_morbilidad.empty:
        datos['Morbilidad'] = df_morbilidad.reset_index().rename(columns={'index': 'Diagnóstico'})
    
    # 7. Gérmenes Aislados
    if not df_germenes.empty:
        datos['Gérmenes Aislados'] = df_germenes.copy()
    
    # 8. Eventos Adversos
    if not df_eventos_adversos.empty:
        datos['Eventos Adversos'] = df_eventos_adversos.copy()
    
    # 9. GRAM-SOT
    if not df_gram_sot.empty:
        gram_export = {'Indicador': 'GRAM - SOT'}
        for mes in meses_disponibles:
            gram_export[meses_completos.get(mes, mes)] = df_gram_sot[mes].iloc[0] if mes in df_gram_sot.columns else 0
        gram_export['Total'] = df_gram_sot.values.sum()
        datos['GRAM-SOT'] = pd.DataFrame([gram_export])
    
    # 10. Ventilación Mecánica
    vm_data = []
    indicadores_vm = [
        ('PACIENTES VENTILADOS', 'Pacientes Ventilados'),
        ('DIAS VENTILADOR', 'Días Ventilador'),
        ('TASA NEUMONÍA ASOCIADA A VENTILADOR', 'Tasa NAV'),
    ]
    for row_name, label in indicadores_vm:
        fila = {'Indicador': label}
        data = safe_get_row(row_name, meses_disponibles, silent=True)
        for i, mes in enumerate(meses_disponibles):
            fila[meses_completos.get(mes, mes)] = data.iloc[i] if isinstance(data, pd.Series) else 0
        total = safe_get_row(row_name, suma_col, silent=True)
        fila['Total'] = total if not pd.isna(total) else 0
        vm_data.append(fila)
    datos['Ventilacion Mecanica'] = pd.DataFrame(vm_data)
    
    # 11. Fallecidos detalle
    if not df_fallecidos.empty:
        df_fall_export = df_fallecidos[['Nombre', 'Edad', 'Sexo', 'Dias_Estancia', 'Mes']].copy()
        df_fall_export.columns = ['Nombre', 'Edad', 'Sexo', 'Días Estancia', 'Mes']
        datos['Detalle Fallecidos'] = df_fall_export
    
    return datos


# --- CALLBACK EXPORTAR A EXCEL ---
@app.callback(
    Output("download-excel", "data"),
    Input("btn-export-excel", "n_clicks"),
    prevent_initial_call=True
)
def exportar_excel(n_clicks):
    if not n_clicks:
        return dash.no_update
    
    datos = generar_datos_exportacion()
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for nombre_hoja, df_hoja in datos.items():
            df_hoja.to_excel(writer, sheet_name=nombre_hoja[:31], index=False)
            
            # Aplicar estilos
            ws = writer.sheets[nombre_hoja[:31]]
            
            # Estilos
            header_fill = PatternFill(start_color="1a365d", end_color="1a365d", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF", size=11)
            cell_border = Border(
                left=Side(style='thin', color='CCCCCC'),
                right=Side(style='thin', color='CCCCCC'),
                top=Side(style='thin', color='CCCCCC'),
                bottom=Side(style='thin', color='CCCCCC')
            )
            
            # Aplicar estilos a encabezados
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = cell_border
            
            # Aplicar estilos a datos y ajustar ancho de columnas
            for col_idx, col in enumerate(ws.columns, 1):
                max_length = 0
                for row_idx, cell in enumerate(col, 1):
                    if row_idx > 1:  # Datos
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                        cell.border = cell_border
                        # Formato numérico
                        if isinstance(cell.value, float):
                            cell.number_format = '#,##0.00'
                        elif isinstance(cell.value, int):
                            cell.number_format = '#,##0'
                    # Calcular ancho
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                # Ajustar ancho de columna
                ws.column_dimensions[col[0].column_letter].width = min(max_length + 3, 40)
            
            # Primera columna alineada a la izquierda
            for row in ws.iter_rows(min_row=2, max_col=1):
                for cell in row:
                    cell.alignment = Alignment(horizontal='left', vertical='center')
    
    output.seek(0)
    fecha = datetime.now().strftime("%Y%m%d_%H%M")
    return dcc.send_bytes(output.getvalue(), f"Dashboard_UCI_{fecha}.xlsx")


# --- CALLBACK EXPORTAR A CSV ---
@app.callback(
    Output("download-csv", "data"),
    Input("btn-export-csv", "n_clicks"),
    prevent_initial_call=True
)
def exportar_csv(n_clicks):
    if not n_clicks:
        return dash.no_update
    
    datos = generar_datos_exportacion()
    output = BytesIO()
    
    # Combinar todos los DataFrames en un solo CSV con separadores
    lineas = []
    lineas.append("DASHBOARD UCI ADULTOS 2025 - REPORTE DE DATOS")
    lineas.append(f"Fecha de generación: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    lineas.append("")
    
    for nombre_seccion, df_seccion in datos.items():
        lineas.append("=" * 80)
        lineas.append(nombre_seccion.upper())
        lineas.append("=" * 80)
        lineas.append(df_seccion.to_csv(index=False, sep=';', decimal=','))
        lineas.append("")
    
    contenido = "\n".join(lineas)
    
    fecha = datetime.now().strftime("%Y%m%d_%H%M")
    return dict(content=contenido, filename=f"Dashboard_UCI_{fecha}.csv")


if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=8050)
