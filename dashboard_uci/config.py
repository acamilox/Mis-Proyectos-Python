# -*- coding: utf-8 -*-
"""
Configuración del Dashboard UCI Adultos
Hospital Infantil Universitario de San José
"""

from pathlib import Path
import os

# Directorios
BASE_DIR = Path(__file__).parent.parent
EXCEL_PATH = BASE_DIR / "data" / "datos_iaas.xlsx"
ASSETS_DIR = BASE_DIR / "assets"

# Meses
MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", 
         "Jul", "Agos", "Sep", "Oct", "Nov", "Dic"]

MESES_COMPLETOS = {
    'Ene': 'Enero', 'Feb': 'Febrero', 'Mar': 'Marzo', 'Abr': 'Abril',
    'May': 'Mayo', 'Jun': 'Junio', 'Jul': 'Julio', 'Agos': 'Agosto',
    'Sep': 'Septiembre', 'Oct': 'Octubre', 'Nov': 'Noviembre', 'Dic': 'Diciembre'
}

# Paleta de colores - Estilo JetBrains Darcula / New UI
COLORS = {
    'primary': '#3574f0',      # Azul selección JetBrains
    'secondary': '#cc7832',    # Naranja keywords
    'accent': '#629755',       # Verde comentarios
    'success': '#6a8759',      # Verde strings
    'danger': '#ff453a',       # Rojo error
    'warning': '#f59e0b',      # Ámbar
    'info': '#9876aa',         # Violeta variables
    'light': '#bcbec4',        # Texto estándar
    'dark': '#1e1f22',         # Fondo New UI
    'muted': '#70727a',        # Texto secundario
    'bg_primary': '#1e1f22',   # Fondo principal
    'bg_secondary': '#2b2d30', # Fondo secundario
    'bg_card': '#2b2d30',
    'text_primary': '#bcbec4',
    'text_secondary': '#70727a',
    'border': '#323232',
    'gradient_start': '#1e1f22',
    'gradient_end': '#2b2d30',
}

# Paleta extendida para gráficas (Estilo Syntax Highlighting)
CHART_COLORS = [
    '#cc7832', # Orange (Keyword)
    '#6a8759', # Green (String)
    '#9876aa', # Purple (Attribute)
    '#6897bb', # Blue (Number)
    '#ffc66d', # Yellow (Function)
    '#629755', # Dark Green (Comment)
    '#3574f0', # Bright Blue (Link)
    '#e06c75', # Soft Red
    '#56b6c2', # Cyan
    '#abb2bf', # Grey
    '#c678dd', # Magenta
    '#d19a66', # Light Orange
]

# Diagnósticos de morbilidad
DIAGNOSTICOS_MORBILIDAD = [
    'ACV', 'CARDIOVASCULAR', 'RESPIRATORIO', 'SEPSIS', 'POP', 'DM', 
    'GINECO-OBSTE', 'GINECO-ONCOLOGIA', 'GINECO', 'TRAUMA', 'TOXICOLOGIA', 'OTROS'
]

# Nombres femeninos para detección de sexo
NOMBRES_FEMENINOS = [
    'maria', 'ana', 'luz', 'rosa', 'carmen', 'martha', 'marta', 'gloria', 
    'patricia', 'sandra', 'claudia', 'diana', 'adriana', 'carolina', 'andrea',
    'monica', 'laura', 'paula', 'angela', 'beatriz', 'cecilia', 'clara',
    'cristina', 'elena', 'esperanza', 'eva', 'flor', 'graciela', 'hilda',
    'irma', 'isabel', 'julia', 'liliana', 'lucia', 'luisa', 'margarita',
    'marina', 'mercedes', 'nancy', 'nelly', 'nidia', 'olga', 'oliva',
    'pilar', 'rocio', 'ruth', 'silvia', 'sonia', 'stella', 'teresa',
    'victoria', 'virginia', 'yolanda', 'blanca', 'barbara', 'nesly',
    'claribel', 'evelia', 'zulema', 'angelica', 'milena', 'ines'
]

# Meses para sheets de Excel
MESES_SHEET_ENERO = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 
                     'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE']

MESES_MAP_GRAM = {
    'enero': 'Ene', 'febrero': 'Feb', 'marzo': 'Mar', 'abril': 'Abr',
    'mayo': 'May', 'junio': 'Jun', 'julio': 'Jul', 'agosto': 'Agos',
    'septiembre': 'Sep', 'octubre': 'Oct', 'noviembre': 'Nov', 'diciembre': 'Dic'
}

# Configuración de gráficos
GRAPH_CONFIG = {
    'staticPlot': True,
    'displayModeBar': False
}

# EPS soportadas
EPS_LIST = [
    'FAMISANAR', 'COMPENSAR', 'SALUD TOTAL', 'NUEVA  EPS', 'CAPITAL SALUD ',
    'SANITAS', 'EMCOSALUD', 'SURA', 'ALIANSALUD', 'EMSSANAR',
    'FIDUSUARIA LA PREVISORA', 'SOAT', 'PARTICULAR', 
    'ADMINISTRADORA DE RECURSOS SGSS', 'OTRA'
]
