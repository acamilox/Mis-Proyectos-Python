# Informe de verificación: Excel 2025 vs Dashboard

**Archivo:** `data/Estadistica 2025.xlsx`

---

## Resumen

El Excel 2025 tiene estructura completa. Hay una **ligera diferencia** en fallecidos que el dashboard ya resuelve usando la hoja MORTALIDAD como fuente.

---

## Datos verificados

| Indicador        | ESTADISTICA      | MORTALIDAD / Otras | Estado      |
| :--------------- | :--------------- | :----------------- | :---------- |
| Total ingresos   | 1065             | -                  | OK          |
| Total egresos    | 1044             | -                  | OK          |
| FALLECIDOS       | 92 (∑)           | 91 pacientes       | Ajustado*   |
| IAAS (gérmenes)  | -                | 29 registros       | OK          |
| GRAM-SOT         | -                | Hoja presente      | OK          |
| Morbilidad       | -                | Hoja presente      | OK          |
| Eventos adversos | -                | Hoja presente      | OK          |

\* El dashboard usa la hoja **MORTALIDAD** (91) como fuente prioritaria, igual que para 2026.

---

## Fallecidos: 92 vs 91

- **ESTADISTICA** (fila FALLECIDOS, columna ∑): 92  
- **MORTALIDAD** (pacientes listados): 91  

Hay 1 paciente de diferencia. Con la lógica actual, el dashboard muestra **91** (MORTALIDAD) para mantener coherencia con la pestaña Mortalidad.

---

## Conclusión

El Excel 2025 está en buen estado. No se requiere ningún cambio adicional en el código: la prioridad de MORTALIDAD sobre ESTADISTICA aplica igual para 2025 y 2026.
