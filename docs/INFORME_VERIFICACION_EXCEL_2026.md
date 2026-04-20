# Informe de verificación: Excel 2026 vs Dashboard

**Archivo:** `data/Estadistica 2026.xlsx`  
**Fecha de verificación:** 10 de abril de 2026 (revisión tras actualización IAAS)

---

## Resumen

Se comparó el Excel actual con la lógica de carga del dashboard (`dashboard_uci/loaders.py`, `dashboard_uci.py`, `processors.py`) usando la validación automática.

**Resultado general:** carga correcta y consistente en los indicadores principales.  
El módulo IAAS ya no está vacío: ahora se detecta el formato nuevo por tipo de infección y se refleja en la dona.

**Comprobación automática:** `python scripts/verificar_excel_2026.py` (requiere `pandas` y `openpyxl`).

---

## Datos validados

| Indicador | Resultado actual | Estado |
| :-------- | :--------------- | :----- |
| Total ingresos | 187 | OK |
| Total egresos | 190 | OK |
| Total fallecidos (Resumen) | 27 | OK |
| Fallecidos (hoja MORTALIDAD) | 27 | OK |
| % Ocupación promedio | 82.9% | OK |
| Estancia promedio | 8.40 días | OK |
| % Mortalidad | 14.2% | OK |
| Morbilidad total (hoja MORBILIDAD) | 374 | OK |
| Eventos adversos | 17 | OK |
| IAAS total en dashboard | 7 | OK |

---

## IAAS: estado actual

La hoja IAAS ya trae registros en formato nuevo (detalle por paciente/tipo de infección).

- Filas IAAS detectadas en formato nuevo: **9**
- Categorías IAAS únicas detectadas: **5**
- Total que muestra el dashboard tras normalización: **7**

**Estado:** OK.  
La visualización de IAAS en `Gestión y Calidad` está operativa y ya no aparece en cero por falta de formato.

---

## Diferencias frente al informe anterior

1. **Fallecidos:** antes había desalineación entre RESUMEN y MORTALIDAD; ahora ambos muestran **27**.
2. **IAAS:** antes se reportaba hoja casi vacía; ahora hay datos cargados y visibles en la dona.
3. **Etiquetas en UI:** se ajustaron textos a "IAAS registradas" cuando aplica el formato nuevo.

---

## Observaciones vigentes

- La hoja **GRAM - SOT** no está en el archivo actual. El sistema sigue cargando vacío en ese módulo (comportamiento esperado si no aplica para 2026).

---

## Recomendaciones

1. Mantener la estructura IAAS actual (separadores por mes + filas de detalle) para evitar pérdidas en la carga.
2. Ejecutar `python scripts/verificar_excel_2026.py` cada vez que se reemplace el Excel.
3. Incluir hoja **GRAM - SOT** solo si se van a reportar esos datos en 2026.
