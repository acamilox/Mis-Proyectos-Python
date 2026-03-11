# -*- coding: utf-8 -*-
"""
Carga de datos desde Excel para el Dashboard UCI
"""

import pandas as pd
from openpyxl import load_workbook
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import unicodedata
import logging

from .config import (
    EXCEL_PATH, MESES, MESES_COMPLETOS, DIAGNOSTICOS_MORBILIDAD,
    NOMBRES_FEMENINOS, MESES_SHEET_ENERO, MESES_MAP_GRAM
)

logger = logging.getLogger(__name__)


class DataLoader:
    """Clase para cargar y validar datos desde Excel"""
    
    def __init__(self, excel_path: Path = EXCEL_PATH):
        self.excel_path = excel_path
        self.df = pd.DataFrame()
        self.df_morbilidad = pd.DataFrame()
        self.df_germenes = pd.DataFrame()
        self.df_fallecidos = pd.DataFrame()
        self.df_eventos_adversos = pd.DataFrame()
        self.df_gram_sot = pd.DataFrame()
    
    def load_all(self) -> Dict[str, pd.DataFrame]:
        """Carga todos los datos desde Excel"""
        self._load_estadistica()
        self._load_morbilidad()
        self._load_germenes()
        self._load_fallecidos()
        self._load_eventos_adversos()
        self._load_gram_sot()
        
        return {
            'estadistica': self.df,
            'morbilidad': self.df_morbilidad,
            'germenes': self.df_germenes,
            'fallecidos': self.df_fallecidos,
            'eventos_adversos': self.df_eventos_adversos,
            'gram_sot': self.df_gram_sot
        }
    
    def _load_estadistica(self) -> None:
        """Carga datos principales de la hoja ESTADISTICA"""
        try:
            self.df = pd.read_excel(self.excel_path, sheet_name='ESTADISTICA', header=2)
            self.df = self.df.set_index('DATOS')
            self.df = self.df.replace(['*', '#DIV/0!'], None)
            
            for col in self.df.columns:
                if col not in ['PROMEDIO', None]:
                    try:
                        self.df[col] = pd.to_numeric(self.df[col], errors='coerce')
                    except (ValueError, TypeError):
                        pass
            
            logger.info(f"Datos ESTADISTICA cargados: {len(self.df)} filas, {len(self.df.columns)} columnas")
        except Exception as e:
            logger.error(f"Error cargando ESTADISTICA: {e}")
            self.df = pd.DataFrame()
    
    def _load_morbilidad(self) -> None:
        """Carga datos de morbilidad por diagnóstico"""
        try:
            df_morb_raw = pd.read_excel(self.excel_path, sheet_name='MORBILIDAD', header=None)
            
            morbilidad_data = {diag: {'Hombre': 0, 'Mujer': 0, 'Total': 0} for diag in DIAGNOSTICOS_MORBILIDAD}
            
            for _, row in df_morb_raw.iterrows():
                diag = str(row[0]).strip().upper() if pd.notna(row[0]) else ''
                diag = self._normalize_diagnostico(diag)
                
                if diag in morbilidad_data:
                    hombre = pd.to_numeric(row[1], errors='coerce')
                    mujer = pd.to_numeric(row[2], errors='coerce')
                    total = pd.to_numeric(row[3], errors='coerce')
                    
                    if pd.notna(hombre):
                        morbilidad_data[diag]['Hombre'] += int(hombre)
                    if pd.notna(mujer):
                        morbilidad_data[diag]['Mujer'] += int(mujer)
                    if pd.notna(total):
                        morbilidad_data[diag]['Total'] += int(total)
            
            self.df_morbilidad = pd.DataFrame(morbilidad_data).T
            self.df_morbilidad = self.df_morbilidad[self.df_morbilidad['Total'] > 0]
            self.df_morbilidad = self.df_morbilidad.sort_values('Total', ascending=False)
            
            logger.info(f"Datos MORBILIDAD cargados: {len(self.df_morbilidad)} diagnósticos")
        except Exception as e:
            logger.error(f"Error cargando MORBILIDAD: {e}")
            self.df_morbilidad = pd.DataFrame()
    
    def _normalize_diagnostico(self, diag: str) -> str:
        """Normaliza nombres de diagnósticos"""
        if not diag or diag in [m.upper().replace(' ', '') for m in MESES_SHEET_ENERO]:
            return ''
        if 'CARVIO' in diag or 'CARDIO' in diag:
            return 'CARDIOVASCULAR'
        if diag in ('HOMBRE', 'MUJER', 'TOTAL', ''):
            return ''
        return diag
    
    def _load_germenes(self) -> None:
        """Carga datos de gérmenes aislados (IAAS)"""
        try:
            df_iaas = pd.read_excel(self.excel_path, sheet_name='IAAS', header=None)

            # Algunos archivos (p.ej. 2026 parcial) no incluyen la columna esperada.
            # En ese caso devolvemos dataset vacio sin romper el flujo del dashboard.
            if 5 not in df_iaas.columns:
                logger.warning(
                    "Hoja IAAS sin columna de germenes (columna 5). Se carga vacio."
                )
                self.df_germenes = pd.DataFrame(columns=['Germen', 'Cantidad'])
                return

            germenes_raw = df_iaas[5].dropna()
            germenes_raw = germenes_raw[germenes_raw.astype(str).str.strip() != '']
            
            germenes_norm = germenes_raw.apply(self._normalize_germen)
            self.df_germenes = germenes_norm.value_counts().reset_index()
            self.df_germenes.columns = ['Germen', 'Cantidad']
            
            logger.info(f"Datos IAAS cargados: {len(self.df_germenes)} tipos de gérmenes")
        except Exception as e:
            logger.error(f"Error cargando IAAS: {e}")
            self.df_germenes = pd.DataFrame()
    
    def _normalize_germen(self, g: str) -> str:
        """Normaliza nombres de gérmenes"""
        g = str(g).strip()
        
        germenes_map = {
            'K. pneumoniae': ['K. pneumoniae', 'K.pneumoniae'],
            'K. pneumoniae KPC': ['carbapenemasa'],
            'A. baumannii': ['Acinetobacter'],
            'E. coli': ['E. coli'],
            'P. aeruginosa': ['Pseudomona'],
            'S. hominis': ['S. hominis'],
            'S. aureus': ['S. aureus'],
            'Candida spp.': ['Candida']
        }
        
        for normalized, variants in germenes_map.items():
            if any(v in g for v in variants):
                if normalized == 'K. pneumoniae' and 'carbapenemasa' in g.lower():
                    return 'K. pneumoniae KPC'
                return normalized
        
        return g
    
    def _load_fallecidos(self) -> None:
        """Carga datos de mortalidad - pacientes fallecidos"""
        try:
            df_mort = pd.read_excel(self.excel_path, sheet_name='MORTALIDAD', 
                                   header=None, usecols='A:F')
            
            fallecidos_data = []
            mes_actual = None
            
            for _, row in df_mort.iterrows():
                nombre = str(row[0]).strip() if pd.notna(row[0]) else ''
                
                # Detectar separador de mes
                nombre_upper = nombre.upper().replace(' ', '')
                if nombre_upper in [m.upper() for m in MESES_SHEET_ENERO]:
                    mes_actual = MESES_COMPLETOS.get(nombre_upper[:3], None)
                    if mes_actual is None:
                        # Buscar en MESES_COMPLETOS
                        for k, v in MESES_COMPLETOS.items():
                            if nombre_upper.startswith(k.upper()):
                                mes_actual = v
                                break
                    continue
                
                edad = row[1]
                dias_estancia = self._extraer_dias(row[2])
                
                if not nombre or nombre.upper() == 'NAN':
                    continue
                if not pd.notna(edad) or not isinstance(edad, (int, float)):
                    continue
                
                sexo = 'Mujer' if self._es_nombre_femenino(nombre) else 'Hombre'
                fallecidos_data.append({
                    'Nombre': nombre,
                    'Edad': int(edad),
                    'Sexo': sexo,
                    'Dias_Estancia': dias_estancia,
                    'Mes': mes_actual
                })
            
            self.df_fallecidos = pd.DataFrame(fallecidos_data)
            logger.info(f"Datos MORTALIDAD cargados: {len(self.df_fallecidos)} pacientes")
        except Exception as e:
            logger.error(f"Error cargando MORTALIDAD: {e}")
            self.df_fallecidos = pd.DataFrame()
    
    def _es_nombre_femenino(self, nombre_completo: str) -> bool:
        """Determina si el nombre es femenino"""
        nombre_norm = self._normalizar_texto(nombre_completo)
        primer_nombre = nombre_norm.split()[0] if nombre_norm.split() else ''
        
        for nf in NOMBRES_FEMENINOS:
            if nf == primer_nombre or primer_nombre.startswith(nf):
                return True
        return False
    
    def _normalizar_texto(self, s: str) -> str:
        """Elimina acentos y normaliza texto"""
        return ''.join(c for c in unicodedata.normalize('NFD', s.lower()) 
                      if unicodedata.category(c) != 'Mn')
    
    def _extraer_dias(self, valor) -> Optional[int]:
        """Extrae días de estancia de diferentes formatos"""
        if pd.isna(valor):
            return None
        if isinstance(valor, (int, float)):
            return int(valor)
        
        val_str = str(valor).lower().strip()
        if 'hora' in val_str:
            try:
                horas = int(''.join(filter(str.isdigit, val_str)))
                return max(1, horas // 24)
            except:
                return 1
        
        try:
            return int(float(val_str))
        except:
            return None
    
    def _load_eventos_adversos(self) -> None:
        """Carga datos de eventos adversos"""
        try:
            wb = load_workbook(self.excel_path)
            ws = wb['EVENTOS ADVERSOS']
            
            # Encabezado de columna E en el Excel (Fila 1) - no es un evento real
            excluir_encabezado = 'tipo de evento o incidente adverso'
            todos_eventos = []
            for i in range(1, 200):
                cell = ws.cell(i, 5)
                if cell.value:
                    evento = str(cell.value).strip()
                    evento_upper = evento.upper()
                    evento_norm = evento.strip().lower()
                    if (evento_upper not in MESES_SHEET_ENERO
                            and evento
                            and evento_upper != 'NAN'
                            and evento_norm != excluir_encabezado):
                        todos_eventos.append(self._normalizar_evento(evento))
            
            self.df_eventos_adversos = pd.Series(todos_eventos).value_counts().reset_index()
            self.df_eventos_adversos.columns = ['Evento', 'Cantidad']
            
            logger.info(f"Datos EVENTOS ADVERSOS cargados: {len(todos_eventos)} eventos")
        except Exception as e:
            logger.error(f"Error cargando EVENTOS ADVERSOS: {e}")
            self.df_eventos_adversos = pd.DataFrame()
    
    def _normalizar_evento(self, e: str) -> str:
        """Normaliza nombres de eventos adversos"""
        e_lower = e.lower().strip()
        
        # Mapeo de eventos (más específicos primero)
        eventos_map = {
            'Ruptura accidental balón traqueostomía': ['ruptura', 'balón', 'traqueostom'],
            'Desplazamiento tubo tórax': ['desplazamiento', 'tubo'],
            'LPP talón y sacra': ['lpp', 'talon', 'sacra'],
            'LPP occipital y LECASH': ['lpp', 'occipital', 'lecash'],
            'LPP glúteo y LECASH': ['lpp', 'gluteo', 'lecash'],
            'LPP occipital categoría II': ['lpp', 'occipital', 'categor'],
            'Autoextubación': ['autoextubación', 'autoextubacion'],
            'Autoretiro CVC': ['autoretiro', 'cvc'],
            'Autoretiro PICC': ['autoretiro', 'picc'],
            'Autoretiro Ventriculostomía': ['autoretiro', 'ventriculostomia'],
            'Desplazamiento sonda/gastrostomía': ['desplazamiento', 'gastrostomia', 'sonda'],
            'Desplazamiento CVC': ['desplazamiento', 'cvc', 'venoso central'],
            'Desplazamiento catéter hemodiálisis': ['desplazamiento', 'hemodialisis', 'hemodiálisis'],
            'Desplazamiento línea arterial': ['desplazamiento', 'arterial'],
            'Desplazamiento nefrostomía': ['desplazamiento', 'nefrostomia'],
            'Desplazamiento derivación biliar': ['desplazamiento', 'biliar'],
            'Desplazamiento catéter paracentesis': ['desplazamiento', 'parecentesis', 'paracentesis'],
            'LPP por dispositivo médico': ['lpp', 'lesion', 'presion', 'dispositivo', 'medico', 'collar'],
            'LPP región sacra': ['lpp', 'lesion', 'presion', 'sacra', 'sacro'],
            'LPP talón': ['lpp', 'talon', 'talón'],
            'LPP occipital': ['lpp', 'occipital'],
            'LPP glúteo': ['lpp', 'gluteo', 'glúteo'],
            'LECASH': ['lecash'],
            'Lesión MARSI': ['marsi'],
            'Error prescripción': ['error', 'prescripci'],
            'Error medicamento (contraste)': ['error', 'medicamento', 'contraste'],
            'Error en retiro CVC': ['error', 'retiro', 'cvc'],
            'Retraso administración medicamento': ['retraso', 'medicamento', 'administración'],
            'Error de medicación': ['error', 'medicamento', 'medicación'],
            'Problema con medicamento': ['problema', 'medicamento'],
            'Caída': ['caida', 'caída'],
            'Flebitis': ['flebitis'],
            'Neumotórax': ['neumotórax', 'neumotorax'],
        }
        
        for normalized, variants in eventos_map.items():
            if all(v in e_lower for v in variants):
                return normalized
        
        return e.strip().capitalize()
    
    def _load_gram_sot(self) -> None:
        """Carga datos de GRAM - SOT (Donaciones/Trasplantes)"""
        try:
            df_gram_raw = pd.read_excel(self.excel_path, sheet_name='GRAM - SOT', header=None)
            
            gram_data = {}
            for _, row in df_gram_raw.iterrows():
                mes_raw = str(row[0]).strip().lower() if pd.notna(row[0]) else ''
                mes_key = MESES_MAP_GRAM.get(mes_raw, None)
                if mes_key:
                    valor = row[1] if pd.notna(row[1]) else 0
                    gram_data[mes_key] = int(valor) if pd.notna(valor) else 0
            
            self.df_gram_sot = pd.DataFrame([gram_data])
            self.df_gram_sot = self.df_gram_sot.reindex(columns=MESES, fill_value=0)
            
            logger.info(f"Datos GRAM-SOT cargados: {self.df_gram_sot.values.sum()} total")
        except ValueError:
            logger.warning("Hoja GRAM - SOT no disponible en el archivo. Se carga vacio.")
            self.df_gram_sot = pd.DataFrame()
        except Exception as e:
            logger.error(f"Error cargando GRAM - SOT: {e}")
            self.df_gram_sot = pd.DataFrame()


def load_data() -> Dict[str, pd.DataFrame]:
    """Función de conveniencia para cargar todos los datos"""
    loader = DataLoader()
    return loader.load_all()
