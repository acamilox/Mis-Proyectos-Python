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

# Paleta de colores - Tema Moderno y Elegante (Sin Amarillo)
COLORS = {
    'primary': '#3776ab',      # Azul Python
    'secondary': '#6366f1',    # Indigo vibrante (Reemplazo del amarillo)
    'accent': '#06b6d4',       # Cyan
    'success': '#10b981',      # Esmeralda
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
    'gradient_end': '#1e3a5f',
},

# Paleta extendida para gráficas (Colores bonitos y variados)
CHART_COLORS = [
    '#6366f1', # Indigo
    '#10b981', # Esmeralda
    '#3b82f6', # Azul brillante
    '#8b5cf6', # Violeta
    '#06b6d4', # Cyan
    '#f43f5e', # Rosa fuerte
    '#14b8a6', # Teal
    '#f59e0b', # Ámbar/Naranja
    '#84cc16', # Lima
    '#ec4899', # Rosa
    '#0ea5e9', # Sky
    '#fb923c', # Naranja
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
