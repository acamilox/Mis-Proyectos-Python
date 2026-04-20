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


def _find_sheet_name(excel_path: Path, candidates: Tuple[str, ...]) -> Optional[str]:
    """
    Devuelve el nombre exacto de la hoja en el libro si coincide con algún candidato
    (ignora mayúsculas y espacios al inicio/fin).
    """
    try:
        xf = pd.ExcelFile(excel_path)
        by_lower = {str(n).strip().lower(): n for n in xf.sheet_names}
        for c in candidates:
            k = str(c).strip().lower()
            if k in by_lower:
                return str(by_lower[k])
    except Exception as exc:
        logger.warning("No se pudieron listar las hojas de %s: %s", excel_path, exc)
    return None


def _sum_morbilidad_pies_mensuales(df_morb_raw: pd.DataFrame) -> Optional[int]:
    """
    Suma los subtotales mensuales de la hoja MORBILIDAD (filas con col2 'Total' y
    sin diagnóstico en col0). Coincide con la suma que se obtiene en el Excel al
    totalizar por mes. Si el detalle por filas no cuadra con ese pie (error de
    digitación en una celda), el total oficial del Excel es esta suma.
    """
    total = 0
    n = 0
    for _, row in df_morb_raw.iterrows():
        c0 = row[0]
        c2 = row[2]
        c3 = row[3]
        col0_empty = pd.isna(c0) or not str(c0).strip() or str(c0).strip().upper() == "NAN"
        if not col0_empty:
            continue
        if pd.isna(c2) or str(c2).strip().lower() != "total":
            continue
        v = pd.to_numeric(c3, errors="coerce")
        if pd.notna(v):
            total += int(v)
            n += 1
    if n < 1:
        return None
    return total


def _morbilidad_pies_por_mes(df_morb_raw: pd.DataFrame) -> Dict[str, int]:
    """Extrae subtotales mensuales (fila Total) de la hoja MORBILIDAD."""
    meses_upper = {m.upper(): k for k, m in MESES_COMPLETOS.items()}
    mes_actual: Optional[str] = None
    out: Dict[str, int] = {k: 0 for k in MESES}
    for _, row in df_morb_raw.iterrows():
        c0 = row[0]
        c2 = row[2]
        c3 = row[3]
        if pd.notna(c0):
            s0 = str(c0).strip().upper()
            if s0 in meses_upper:
                mes_actual = meses_upper[s0]
                continue
        if mes_actual is None:
            continue
        if pd.isna(c2) or str(c2).strip().lower() != "total":
            continue
        v = pd.to_numeric(c3, errors="coerce")
        if pd.notna(v):
            out[mes_actual] += int(v)
    return out


class DataLoader:
    """Clase para cargar y validar datos desde Excel"""
    
    def __init__(self, excel_path: Path = EXCEL_PATH):
        self.excel_path = excel_path
        self.df = pd.DataFrame()
        self.df_morbilidad = pd.DataFrame()
        self.df_germenes = pd.DataFrame()
        self.df_germenes_detalle = pd.DataFrame()
        self.df_fallecidos = pd.DataFrame()
        self.df_eventos_adversos = pd.DataFrame()
        self.df_eventos_detalle = pd.DataFrame()
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
            'germenes_detalle': self.df_germenes_detalle,
            'fallecidos': self.df_fallecidos,
            'eventos_adversos': self.df_eventos_adversos,
            'eventos_adversos_detalle': self.df_eventos_detalle,
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

            detail_sum = int(self.df_morbilidad["Total"].sum())
            footer_sum = _sum_morbilidad_pies_mensuales(df_morb_raw)
            self.df_morbilidad.attrs["totales_mensuales"] = _morbilidad_pies_por_mes(df_morb_raw)
            if footer_sum is not None:
                self.df_morbilidad.attrs["total_anual"] = footer_sum
                if footer_sum != detail_sum:
                    logger.warning(
                        "MORBILIDAD: suma detalle (%s) != suma pies mensuales (%s); "
                        "se usa total_anual del Excel (pies).",
                        detail_sum,
                        footer_sum,
                    )
            else:
                self.df_morbilidad.attrs["total_anual"] = detail_sum
            
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
            sh_iaas = _find_sheet_name(self.excel_path, ("IAAS", "I.A.A.S."))
            if not sh_iaas:
                logger.warning(
                    "Hoja IAAS no encontrada en %s (hojas disponibles: usar nombre 'IAAS'). Se carga vacío.",
                    self.excel_path.name,
                )
                self.df_germenes = pd.DataFrame()
                self.df_germenes_detalle = pd.DataFrame()
                return
            df_iaas = pd.read_excel(self.excel_path, sheet_name=sh_iaas, header=None)
            if df_iaas.empty:
                self.df_germenes = pd.DataFrame(columns=['Germen', 'Cantidad'])
                self.df_germenes_detalle = pd.DataFrame(columns=['Mes', 'Germen'])
                logger.warning("Hoja IAAS vacía. Se carga dataset vacío.")
                return

            columnas_num = sorted([c for c in df_iaas.columns if isinstance(c, int)])
            # Formato histórico: columna >=5 contiene germen.
            germ_cols = [c for c in columnas_num if c >= 5]
            # Formato nuevo 2026: suele venir con 5 columnas (0..4) y col4 es tipo de IAAS.
            fallback_col = 4 if 4 in columnas_num else (columnas_num[-1] if columnas_num else None)
            iaas_mode = 'germenes' if germ_cols else 'tipos_iaas'

            def _resolve_mes(value) -> Optional[str]:
                if pd.isna(value):
                    return None
                token = str(value).strip()
                if not token:
                    return None
                token_norm = self._normalizar_texto(token)
                if token_norm in MESES_MAP_GRAM:
                    return MESES_MAP_GRAM[token_norm]
                token_up = token.upper()
                if token_up in MESES_SHEET_ENERO:
                    return MESES_MAP_GRAM.get(token_norm)
                for raw_mes, mes_short in MESES_MAP_GRAM.items():
                    if token_norm.startswith(raw_mes[:3]):
                        return mes_short
                return None

            def _is_data_row_with_age(row) -> bool:
                c0 = row[0] if 0 in df_iaas.columns else None
                if pd.isna(c0) or not str(c0).strip():
                    return False
                c1 = row[1] if 1 in df_iaas.columns else None
                return pd.notna(pd.to_numeric(c1, errors='coerce'))

            filas_detalle: List[Dict[str, str]] = []
            mes_actual: Optional[str] = None
            header_tokens = {
                'germen', 'germenes', 'microorganismo', 'microorganismos',
                'tipo de germen', 'tipo microorganismo'
            }

            for _, row in df_iaas.iterrows():
                mes_detectado = _resolve_mes(row[0] if 0 in df_iaas.columns else None)
                if mes_detectado:
                    mes_actual = mes_detectado
                    continue

                valor_fuente: Optional[str] = None
                if germ_cols:
                    # Tomar columna más a la derecha con valor (suele ser germen más específico).
                    for c in sorted(germ_cols, reverse=True):
                        v = row[c] if c in df_iaas.columns else None
                        if pd.notna(v) and str(v).strip():
                            valor_fuente = str(v).strip()
                            break
                elif fallback_col is not None and _is_data_row_with_age(row):
                    v = row[fallback_col] if fallback_col in df_iaas.columns else None
                    if pd.notna(v) and str(v).strip():
                        valor_fuente = str(v).strip()

                if not valor_fuente:
                    continue

                fuente_norm = self._normalizar_texto(valor_fuente)
                if not fuente_norm or fuente_norm in ('nan', 'none'):
                    continue
                if fuente_norm in header_tokens:
                    continue

                germen_norm = self._normalize_germen(valor_fuente)
                if germen_norm and str(germen_norm).upper() != "NAN":
                    filas_detalle.append({"Mes": mes_actual or "N/A", "Germen": str(germen_norm).strip()})

            # Fallback final: si no salió nada en modo nuevo, intentar lectura plana de la columna fallback.
            if not filas_detalle and not germ_cols and fallback_col is not None:
                mes_actual = None
                for _, row in df_iaas.iterrows():
                    mes_detectado = _resolve_mes(row[0] if 0 in df_iaas.columns else None)
                    if mes_detectado:
                        mes_actual = mes_detectado
                        continue
                    v = row[fallback_col] if fallback_col in df_iaas.columns else None
                    if pd.isna(v) or not str(v).strip():
                        continue
                    raw = str(v).strip()
                    raw_norm = self._normalizar_texto(raw)
                    if raw_norm in ('nan', 'none') or raw_norm in header_tokens:
                        continue
                    filas_detalle.append({"Mes": mes_actual or "N/A", "Germen": raw})

            self.df_germenes_detalle = (
                pd.DataFrame(filas_detalle)
                if filas_detalle else
                pd.DataFrame(columns=['Mes', 'Germen'])
            )
            if not self.df_germenes_detalle.empty:
                germenes_norm = self.df_germenes_detalle['Germen'].astype(str).str.strip()
                germenes_norm = germenes_norm[germenes_norm != ""]
                self.df_germenes = germenes_norm.value_counts().reset_index()
                self.df_germenes.columns = ['Germen', 'Cantidad']
            else:
                self.df_germenes = pd.DataFrame(columns=['Germen', 'Cantidad'])

            self.df_germenes.attrs['iaas_mode'] = iaas_mode
            self.df_germenes_detalle.attrs['iaas_mode'] = iaas_mode
            if iaas_mode == 'tipos_iaas':
                logger.info("Datos IAAS cargados (modo tipos): %s categorías", len(self.df_germenes))
            else:
                logger.info("Datos IAAS cargados: %s tipos de gérmenes", len(self.df_germenes))
        except Exception as e:
            logger.error(f"Error cargando IAAS: {e}")
            self.df_germenes = pd.DataFrame()
            self.df_germenes_detalle = pd.DataFrame()
    
    def _normalize_germen(self, g: str) -> str:
        """Normaliza nombres de gérmenes"""
        g = str(g).strip()
        g_lower = g.lower()
        
        germenes_map = {
            'K. pneumoniae': ['k. pneumoniae', 'k.pneumoniae', 'klebsiella pneumoniae'],
            'K. pneumoniae KPC': ['carbapenemasa', 'kpc'],
            'A. baumannii': ['acinetobacter', 'a. baumannii'],
            'E. coli': ['e. coli', 'escherichia coli'],
            'P. aeruginosa': ['pseudomona', 'p. aeruginosa'],
            'S. hominis': ['s. hominis', 'staphylococcus hominis'],
            'S. aureus': ['s. aureus', 'staphylococcus aureus'],
            'Candida spp.': ['candida'],
            'C. difficile': ['clostridium difficile', 'c. difficile'],
        }
        
        for normalized, variants in germenes_map.items():
            if any(v in g_lower for v in variants):
                if normalized == 'K. pneumoniae' and ('carbapenemasa' in g_lower or 'kpc' in g_lower):
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
        """Carga datos de eventos adversos (agregado y detalle con paciente/especialidad)"""
        try:
            sh_ea = _find_sheet_name(
                self.excel_path,
                ("EVENTOS ADVERSOS", "EVENTO ADVERSO", "EVENTOS ADVERSO"),
            )
            if not sh_ea:
                logger.warning(
                    "Hoja EVENTOS ADVERSOS no encontrada en %s. Se carga vacío.",
                    self.excel_path.name,
                )
                self.df_eventos_adversos = pd.DataFrame()
                self.df_eventos_detalle = pd.DataFrame()
                return
            wb = load_workbook(self.excel_path, data_only=True)
            ws = wb[sh_ea]
            excluir_encabezado = 'tipo de evento o incidente adverso'
            todos_eventos = []
            filas_detalle = []
            mes_actual = None
            for i in range(1, 1200):
                col1 = ws.cell(i, 1).value
                col5 = ws.cell(i, 5).value
                col6 = ws.cell(i, 6).value
                col7 = ws.cell(i, 7).value
                # Detectar separador de mes (ENERO, FEBRERO, etc.)
                if col1:
                    nombre_raw = str(col1).strip().lower()
                    for k, v in MESES_MAP_GRAM.items():
                        if nombre_raw == k or nombre_raw.startswith(k[:3]):
                            mes_actual = v
                            break
                    if mes_actual and nombre_raw.upper() in [m.upper() for m in MESES_SHEET_ENERO]:
                        continue
                # Fila de dato: nombre (col1), edad parseable (col2) y evento (col5).
                # Excel a veces guarda la edad como texto ("68"); antes solo aceptábamos int/float.
                edad_val = ws.cell(i, 2).value
                if isinstance(edad_val, (int, float)) and not isinstance(edad_val, bool):
                    edad_ok = pd.notna(edad_val)
                else:
                    edad_ok = pd.notna(pd.to_numeric(edad_val, errors='coerce'))
                is_data_row = (col1 and str(col1).strip()
                    and edad_ok
                    and col5 and str(col5).strip().lower() != excluir_encabezado)
                if is_data_row and col5:
                    evento_norm = self._normalizar_evento(str(col5).strip())
                    evento_upper = str(col5).strip().upper()
                    if evento_upper != 'NAN':
                        todos_eventos.append(evento_norm)
                        especialidad = str(col7).strip() if col7 and pd.notna(col7) else 'Sin especificar'
                        eps = str(col6).strip() if col6 and pd.notna(col6) else 'Sin especificar'
                        nombre = str(col1).strip()
                        filas_detalle.append({
                            'Nombre': nombre,
                            'Evento': evento_norm,
                            'Especialidad': especialidad,
                            'EPS': eps,
                            'Mes': mes_actual or 'N/A'
                        })
            self.df_eventos_adversos = pd.Series(todos_eventos).value_counts().reset_index()
            self.df_eventos_adversos.columns = ['Evento', 'Cantidad']
            self.df_eventos_detalle = pd.DataFrame(filas_detalle) if filas_detalle else pd.DataFrame(
                columns=['Nombre', 'Evento', 'Especialidad', 'EPS', 'Mes'])
            logger.info(f"Datos EVENTOS ADVERSOS cargados: {len(todos_eventos)} eventos, {len(filas_detalle)} filas detalle")
        except Exception as e:
            logger.error(f"Error cargando EVENTOS ADVERSOS: {e}")
            self.df_eventos_adversos = pd.DataFrame()
            self.df_eventos_detalle = pd.DataFrame()
    
    def _normalizar_evento(self, e: str) -> str:
        """Normaliza nombres de eventos adversos"""
        e_lower = e.lower().strip()

        # Variantes ortográficas: cualquier coincidencia basta (no AND como en frases compuestas)
        eventos_sinonimos: Tuple[Tuple[str, Tuple[str, ...]], ...] = (
            ('Autoextubación', ('autoextubación', 'autoextubacion', 'auto extub', 'auto-extub', 'autoextub')),
            ('Neumotórax', ('neumotórax', 'neumotorax')),
            ('Caída', ('caida', 'caída')),
        )
        for normalized, opciones in eventos_sinonimos:
            if any(op in e_lower for op in opciones):
                return normalized

        # Mapeo de eventos compuestos: todas las subcadenas deben aparecer (AND)
        eventos_map = {
            'Ruptura accidental balón traqueostomía': ['ruptura', 'balón', 'traqueostom'],
            'Desplazamiento tubo tórax': ['desplazamiento', 'tubo'],
            'LPP talón y sacra': ['lpp', 'talon', 'sacra'],
            'LPP occipital y LECASH': ['lpp', 'occipital', 'lecash'],
            'LPP glúteo y LECASH': ['lpp', 'gluteo', 'lecash'],
            'LPP occipital categoría II': ['lpp', 'occipital', 'categor'],
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
            'Flebitis': ['flebitis'],
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
