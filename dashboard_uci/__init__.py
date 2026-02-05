# -*- coding: utf-8 -*-
"""
Dashboard UCI Adultos
Hospital Infantil Universitario de San José

Paquete modular para el dashboard de indicadores de la UCI.
"""

__version__ = "2.0.0"
__author__ = "Hospital Infantil Universitario de San José"

from .config import (
    BASE_DIR,
    EXCEL_PATH,
    ASSETS_DIR,
    MESES,
    MESES_COMPLETOS,
    COLORS,
    DIAGNOSTICOS_MORBILIDAD,
    NOMBRES_FEMENINOS,
    GRAPH_CONFIG,
    EPS_LIST
)

from .loaders import DataLoader, load_data
from .processors import DataProcessor, create_export_data
from .chart_factory import ChartFactory, chart_factory

__all__ = [
    # Config
    'BASE_DIR',
    'EXCEL_PATH',
    'ASSETS_DIR', 
    'MESES',
    'MESES_COMPLETOS',
    'COLORS',
    'GRAPH_CONFIG',
    # Loaders
    'DataLoader',
    'load_data',
    # Processors
    'DataProcessor',
    'create_export_data',
    # Chart Factory
    'ChartFactory',
    'chart_factory',
]
