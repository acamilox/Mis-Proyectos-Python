# Informe de verificación: Excel 2026 vs Dashboard

**Archivo:** `data/Estadistica 2026.xlsx`  
**Fecha de verificación:** 28 de marzo de 2026 (revisión tras actualización de datos)

---

## Resumen

Se compararon los datos del Excel con la estructura y lógica que usa el dashboard (`dashboard_uci/loaders.py`, `processors.py`). El archivo **carga sin errores**; persisten **inconsistencias** entre hojas y **hojas incompletas** respecto al archivo 2025 de referencia.

**Comprobación automática:** `python scripts/verificar_excel_2026.py` (requiere dependencias del proyecto: `pandas`, `openpyxl`).

---

## ✅ Datos coherentes

| Indicador              | Excel (ESTADISTICA)   | Cálculo Dashboard              | Estado |
| :--------------------- | :-------------------- | :----------------------------- | :----- |
| Total ingresos         | 142 (∑)               | Suma mensual                   | OK     |
| Total egresos          | 139 (∑)               | Suma mensual                   | OK     |
| % Ocupación promedio   | ~88.7% (Ene+Feb)      | Promedio de valores > 0        | OK     |
| Promedio estancia      | ~8.1 días             | Promedio mensual               | OK     |
| Giro cama              | ~3.3                  | Promedio mensual               | OK     |
| Morbilidad             | 284 total             | Suma por diagnóstico           | OK     |
| Eventos adversos       | 11 filas              | Carga desde hoja EVENTOS ADV.  | OK     |

---

## ❌ Inconsistencias detectadas

### 1. FALLECIDOS: ESTADISTICA vs MORTALIDAD

| Fuente                                 | Valor                   |
| :------------------------------------- | :---------------------- |
| **ESTADISTICA** (fila FALLECIDOS)      | **8**                   |
| **MORTALIDAD** (pacientes listados)    | **20** (8 Ene + 12 Feb) |

**Impacto:**

- La pestaña **Resumen General** usa ESTADISTICA → muestra **8 fallecidos**
- La pestaña **Mortalidad** usa la hoja MORTALIDAD → muestra **20 fallecidos**

**Causa probable:** La fila FALLECIDOS en ESTADISTICA solo está actualizada con Enero (8). Febrero (12 fallecidos) no está incorporado en esa fila.

**Recomendación:** Actualizar en ESTADISTICA la fila FALLECIDOS para incluir Febrero = 12, de modo que la columna ∑ sea 20 y coincida con MORTALIDAD.

---

### 2. IAAS: hoja casi vacía

La hoja **IAAS** tiene:

- 1 sola fila con el texto "ENERO"
- Sin columna de gérmenes (el loader espera columna F / índice 5)

**Impacto:** El dashboard no mostrará gérmenes aislados; el gráfico de IAAS quedará vacío.

**Recomendación:** Completar la hoja IAAS con la estructura esperada (nombre del germen en columna F) o confirmar si para 2026 aún no hay datos de IAAS.

---

### 3. Hoja GRAM - SOT ausente

El archivo 2026 no incluye la hoja **"GRAM - SOT"** (donaciones/trasplantes).

**Impacto:** El loader carga un dataset vacío; no habrá gráficos relacionados.

**Recomendación:** Añadir la hoja si existen datos de GRAM-SOT para 2026, o mantener el comportamiento actual (vacío) si no aplica.

---

## Estructura del Excel 2026

### Hojas presentes

- ESTADISTICA  
- EVENTOS ADVERSOS  
- MORTALIDAD  
- IAAS  
- MORBILIDAD  

### Meses con datos en ESTADISTICA

- **Enero** y **Febrero** tienen valores
- **Mar–Dic** aparecen como 0 o vacíos (aún no se cargan datos)

### Columnas en ESTADISTICA

Ene, Feb, Mar, Abr, May, Jun, Jul, Agos, Sep, Oct, Nov, Dic, ∑, PROMEDIO (y a veces una columna extra `Unnamed: 15` al final; el código no la usa).

### Comparación con `Estadistica 2025.xlsx`

| Aspecto | 2025 | 2026 (actual) |
| :------ | :--- | :-------------- |
| Hoja **GRAM - SOT** | Presente | **Ausente** |
| Hoja **IAAS** | ~55 filas × 6 columnas (gérmenes en col. F) | **1 celda** (“ENERO”), sin datos de pacientes/gérmenes |

---

## Acciones recomendadas

1. **Crítico:** Alinear FALLECIDOS entre ESTADISTICA y MORTALIDAD (actualizar fila FALLECIDOS con el total real de 20).
2. **Opcional:** Completar la hoja IAAS si hay registros de gérmenes para 2026.
3. **Opcional:** Incluir la hoja GRAM - SOT si se van a reportar donaciones/trasplantes.
