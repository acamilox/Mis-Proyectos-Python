# -*- coding: utf-8 -*-
"""
Procesadores de datos para el Dashboard UCI
Transformaciones y cálculos sobre los datos cargados
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Optional, Tuple, Any
from datetime import datetime
import logging

logger = logging.getLogger(__name__)


class DataProcessor:
    """Clase para procesar y transformar datos del Dashboard"""
    
    def __init__(self, data: Dict[str, pd.DataFrame]):
        from .config import MESES
        self.data = data
        self.df = data.get('estadistica', pd.DataFrame())
        self.df_morbilidad = data.get('morbilidad', pd.DataFrame())
        self.df_germenes = data.get('germenes', pd.DataFrame())
        self.df_fallecidos = data.get('fallecidos', pd.DataFrame())
        self.df_eventos = data.get('eventos_adversos', pd.DataFrame())
        self.df_gram = data.get('gram_sot', pd.DataFrame())
        self.meses = MESES  # Lista de meses disponibles
    
    def get_suma_column(self) -> str:
        """Obtiene el nombre de la columna de suma"""
        suma_cols = [c for c in self.df.columns if str(c).strip() in ['∑', 'SUMA']]
        return suma_cols[0] if suma_cols else '∑'
    
    def safe_get_row(self, row_name: str, columns: List[str], silent: bool = False) -> pd.Series:
        """
        Obtiene una fila del DataFrame buscando por nombre parcial
        """
        try:
            matching_rows = [idx for idx in self.df.index 
                           if row_name.lower() in str(idx).lower()]
            if matching_rows:
                actual_row = matching_rows[0]
                return self.df.loc[actual_row, columns]
            else:
                if not silent:
                    logger.warning(f"No se encontró la fila: '{row_name}'")
                return pd.Series([0] * len(columns), index=columns)
        except Exception as e:
            if not silent:
                logger.error(f"Error al obtener '{row_name}': {e}")
            return pd.Series([0] * len(columns), index=columns)
    
    def safe_get_value(self, row_name: str, column: str = None) -> Any:
        """Obtiene un valor específico del DataFrame"""
        try:
            matching_rows = [idx for idx in self.df.index 
                           if row_name.lower() in str(idx).lower()]
            if matching_rows:
                actual_row = matching_rows[0]
                if column:
                    return self.df.loc[actual_row, column]
                else:
                    return self.df.loc[actual_row]
            return 0
        except Exception as e:
            logger.error(f"Error obteniendo valor '{row_name}': {e}")
            return 0
    
    def calcular_tendencia(self, datos: pd.Series, inverse: bool = False) -> Tuple[str, str]:
        """
        Calcula la tendencia comparando los últimos meses con datos
        Returns: (trend: 'up'|'down'|'stable', trend_value: str)
        """
        if not isinstance(datos, pd.Series):
            return 'stable', '0%'
        
        valores = datos.dropna()
        valores = valores[valores > 0]
        
        if len(valores) < 2:
            return 'stable', '0%'
        
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
    
    def calcular_promedio_mensual(self, row_name: str, meses: List[str] = None) -> float:
        """Calcula el promedio de una fila para meses seleccionados"""
        if meses is None:
            meses = self.meses
        
        datos = self.safe_get_row(row_name, meses)
        datos_validos = datos[datos > 0]
        
        if len(datos_validos) > 0:
            return datos_validos.mean()
        return 0.0
    
    def calcular_total(self, row_name: str, columna: str = None) -> float:
        """Calcula el total de una fila"""
        if columna:
            return self.safe_get_value(row_name, columna)
        return self.safe_get_value(row_name)
    
    def formatear_porcentaje(self, valor: float, con_simbolo: bool = True) -> str:
        """Formatea porcentaje: sin decimales si es entero, con 1 decimal si no"""
        simbolo = "%" if con_simbolo else ""
        if pd.isna(valor):
            return f"0{simbolo}"
        if valor == int(valor):
            return f"{int(valor)}{simbolo}"
        return f"{valor:.1f}{simbolo}"
    
    def capitalizar_texto(self, texto: str) -> str:
        """Capitaliza la primera letra del texto"""
        if not texto or not isinstance(texto, str):
            return texto
        return texto[0].upper() + texto[1:].lower() if len(texto) > 1 else texto.upper()
    
    def get_kpis_resumen(self, meses: List[str] = None) -> Dict[str, Any]:
        """Calcula KPIs para el módulo de Resumen"""
        if meses is None:
            meses = self.meses
        
        suma_col = self.get_suma_column()
        
        # Datos principales
        datos_ocupacional = self.safe_get_row('% OCUPACIONAL', meses)
        promedio_ocupacional = datos_ocupacional[datos_ocupacional > 0].mean() if len(datos_ocupacional[datos_ocupacional > 0]) > 0 else 0
        
        datos_ingresos = self.safe_get_row('NUMERO DE INGRESOS DEL MES', meses)
        total_ingresos = datos_ingresos.sum() if isinstance(datos_ingresos, pd.Series) else datos_ingresos
        
        datos_egresos = self.safe_get_row('NÚMERO DE EGRESOS DEL MES', meses)
        total_egresos = datos_egresos.sum() if isinstance(datos_egresos, pd.Series) else datos_egresos
        
        datos_fallecidos = self.safe_get_row('FALLECIDOS', meses)
        total_fallecidos = datos_fallecidos.sum() if isinstance(datos_fallecidos, pd.Series) else datos_fallecidos
        
        datos_estancia = self.safe_get_row('PROMEDIO ESTANCIA', meses)
        promedio_estancia = datos_estancia[datos_estancia > 0].mean() if len(datos_estancia[datos_estancia > 0]) > 0 else 0
        
        datos_giro = self.safe_get_row('GIRO CAMA', meses)
        giro_cama = datos_giro[datos_giro > 0].mean() if len(datos_giro[datos_giro > 0]) > 0 else 0
        
        # Mortalidad
        mortalidad = self.safe_get_value('% DE MORTALIDAD GENERAL', 'PROMEDIO')
        if pd.isna(mortalidad):
            mortalidad = (total_fallecidos / total_egresos * 100) if total_egresos > 0 else 0
        
        return {
            'promedio_ocupacional': promedio_ocupacional,
            'total_ingresos': total_ingresos,
            'total_egresos': total_egresos,
            'total_fallecidos': total_fallecidos,
            'promedio_estancia': promedio_estancia,
            'giro_cama': giro_cama,
            'mortalidad': mortalidad
        }
    
    def get_kpis_mortalidad(self) -> Dict[str, Any]:
        """Calcula KPIs para el módulo de Mortalidad"""
        suma_col = self.get_suma_column()
        
        total_fallecidos = self.safe_get_value('FALLECIDOS', suma_col)
        if pd.isna(total_fallecidos):
            total_fallecidos = len(self.df_fallecidos) if not self.df_fallecidos.empty else 0
        
        fallecidos_covid = self.safe_get_value('FALLECIDOS COVID-19', suma_col)
        total_egresos = self.safe_get_value('NÚMERO DE EGRESOS DEL MES', suma_col)
        
        if pd.isna(fallecidos_covid):
            fallecidos_covid = 0
        if pd.isna(total_egresos) or total_egresos == 0:
            total_egresos = 1
        
        mortalidad_general = self.safe_get_value('% DE MORTALIDAD GENERAL', 'PROMEDIO')
        if pd.isna(mortalidad_general):
            mortalidad_general = (total_fallecidos / total_egresos * 100)
        
        mortalidad_trauma = self.safe_get_value('% mortalidad en trauma', 'PROMEDIO')
        if pd.isna(mortalidad_trauma):
            mortalidad_trauma = 0
        
        # Estadísticas de días de estancia
        dias_estancia_validos = self.df_fallecidos['Dias_Estancia'].dropna() if not self.df_fallecidos.empty else pd.Series([0])
        if len(dias_estancia_validos) == 0:
            dias_estancia_validos = pd.Series([0])
        
        promedio_dias = dias_estancia_validos.mean()
        minimo_dias = dias_estancia_validos.min()
        maximo_dias = dias_estancia_validos.max()
        mediana_dias = dias_estancia_validos.median()
        
        return {
            'total_fallecidos': total_fallecidos,
            'fallecidos_covid': fallecidos_covid,
            'total_egresos': total_egresos,
            'mortalidad_general': mortalidad_general,
            'mortalidad_trauma': mortalidad_trauma,
            'promedio_dias_estancia': promedio_dias,
            'minimo_dias_estancia': minimo_dias,
            'maximo_dias_estancia': maximo_dias,
            'mediana_dias_estancia': mediana_dias
        }
    
    def get_kpis_gestion(self) -> Dict[str, Any]:
        """Calcula KPIs para el módulo de Gestión"""
        suma_col = self.get_suma_column()
        meses = self.meses
        
        datos_estancia = self.safe_get_row('PROMEDIO ESTANCIA', meses)
        promedio_estancia = datos_estancia[datos_estancia > 0].mean() if len(datos_estancia[datos_estancia > 0]) > 0 else 0
        
        datos_giro = self.safe_get_row('GIRO CAMA', meses)
        giro_cama = datos_giro[datos_giro > 0].mean() if len(datos_giro[datos_giro > 0]) > 0 else 0
        
        total_dias = self.safe_get_value('TOTAL DIAS ESTANCIA', suma_col)
        capacidad = self.safe_get_value('CAPACIDAD INSTALADA: UCI', 'Ene')
        
        # Promedio de días de VM
        datos_dias_vm = self.safe_get_row('DIAS VENTILADOR', meses, silent=True)
        datos_pac_vm = self.safe_get_row('PACIENTES VENTILADOS', meses, silent=True)
        
        promedio_vm = []
        for i in range(len(meses)):
            dias = datos_dias_vm.iloc[i] if isinstance(datos_dias_vm, pd.Series) else 0
            pac = datos_pac_vm.iloc[i] if isinstance(datos_pac_vm, pd.Series) else 0
            if pd.notna(dias) and pd.notna(pac) and pac > 0:
                promedio_vm.append(round(dias / pac, 2))
            else:
                promedio_vm.append(0)
        
        promedio_vm_validos = [v for v in promedio_vm if v > 0]
        promedio_vm_general = sum(promedio_vm_validos) / len(promedio_vm_validos) if promedio_vm_validos else 0
        
        return {
            'promedio_estancia': promedio_estancia,
            'giro_cama': giro_cama,
            'total_dias_estancia': total_dias,
            'capacidad_uci': capacidad,
            'promedio_dias_vm': promedio_vm_general,
            'promedio_dias_vm_mensual': promedio_vm
        }
    
    def get_eps_data(self) -> Tuple[List[str], List[float]]:
        """Obtiene datos de EPS ordenados por cantidad"""
        from .config import EPS_LIST
        
        suma_col = self.get_suma_column()
        
        eps_data = []
        for eps in EPS_LIST:
            total = self.safe_get_value(eps, suma_col)
            if pd.notna(total) and total > 0:
                eps_data.append((eps.strip(), float(total)))
        
        eps_data.sort(key=lambda x: x[1], reverse=True)
        
        labels = [e[0] for e in eps_data]
        values = [e[1] for e in eps_data]
        
        return labels, values


def create_export_data(data: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    """Genera datos organizados para exportación"""
    from .config import MESES, MESES_COMPLETOS
    from .loaders import DataLoader
    
    loader = DataLoader()
    processor = DataProcessor(data)
    
    datos = {}
    suma_col = processor.get_suma_column()
    
    # 1. Resumen General
    resumen = []
    indicadores = [
        ('% OCUPACIONAL', '% Ocupación'),
        ('NUMERO DE INGRESOS DEL MES', 'Ingresos'),
        ('NÚMERO DE EGRESOS DEL MES', 'Egresos'),
        ('FALLECIDOS', 'Fallecidos'),
        ('PROMEDIO ESTANCIA', 'Estancia Promedio (días)'),
        ('GIRO CAMA', 'Giro de Cama'),
    ]
    
    for row_name, label in indicadores:
        fila = {'Indicador': label}
        row_data = processor.safe_get_row(row_name, MESES)
        for i, mes in enumerate(MESES):
            fila[MESES_COMPLETOS.get(mes, mes)] = row_data.iloc[i] if isinstance(row_data, pd.Series) else 0
        total = processor.safe_get_value(row_name, suma_col)
        fila['Total'] = total if not pd.isna(total) else 0
        resumen.append(fila)
    
    datos['Resumen General'] = pd.DataFrame(resumen)
    
    # 2. Morbilidad
    if not data['morbilidad'].empty:
        datos['Morbilidad'] = data['morbilidad'].reset_index().rename(columns={'index': 'Diagnóstico'})
    
    # 3. Gérmenes
    if not data['germenes'].empty:
        datos['Germenes'] = data['germenes']
    
    # 4. Fallecidos
    if not data['fallecidos'].empty:
        datos['Fallecidos'] = data['fallecidos']
    
    # 5. Eventos Adversos
    if not data['eventos_adversos'].empty:
        datos['Eventos Adversos'] = data['eventos_adversos']
    
    return datos
