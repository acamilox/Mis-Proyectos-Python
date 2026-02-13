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

# Paleta de colores - Tema Python Islands Dark
COLORS = {
    'primary': '#3776ab',      # Azul Python (Python Blue)
    'secondary': '#ffd43b',    # Amarillo Python (Python Yellow)
    'accent': '#4b8bbe',       # Azul claro Python
    'success': '#10b981',      # Esmeralda (mantenido para estados positivos)
    'danger': '#ef4444',       # Rojo coral
    'warning': '#f59e0b',      # Ámbar
    'info': '#64748b',         # Slate
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
    'gradient_end': '#1e3a5f', # Azul profundo (Islands Dark feel)
}

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
