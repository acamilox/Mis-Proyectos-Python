# Dashboard UCI Adultos 2026

Sistema de monitoreo y análisis de indicadores para la Unidad de Cuidados Intensivos de Adultos del Hospital Infantil Universitario de San José.

![Python](https://img.shields.io/badge/Python-3.11-blue)
![Dash](https://img.shields.io/badge/Dash-2.14+-green)
![Plotly](https://img.shields.io/badge/Plotly-5.18+-orange)

## Descripción

Dashboard interactivo desarrollado en Python con Dash y Plotly que permite visualizar y analizar los principales indicadores de gestión, calidad y mortalidad de la UCI. Los datos se cargan desde un archivo Excel (`datos_iaas.xlsx`) con múltiples hojas de información.

## Características

- **Visualización interactiva** de indicadores con gráficos dinámicos
- **Filtros por mes** para análisis temporal
- **4 módulos principales**:
  - Resumen General
  - Gestión y Calidad
  - Especialidades y EPS
  - Mortalidad
- **Exportación de datos** a Excel y CSV
- **Diseño responsive** con tema oscuro profesional

## Módulos

### Resumen General
- KPIs: Ingresos, Ocupación, Mortalidad, Estancia Promedio, Giro de Cama
- Principales EPS (Famisanar, Compensar, Salud Total)
- Gráficos de ocupación mensual e ingresos vs egresos
- Morbilidad por diagnóstico y sexo

### Gestión y Calidad
- Indicadores de estancia y giro de cama
- Promedio de días de ventilación mecánica
- Tasas de infecciones (ITS AC, NAV, ISTU AC, LPP)
- Gérmenes aislados (IAAS)
- Eventos adversos

### Especialidades y EPS
- Medicina Interna, Cirugía General, Neurocirugía, Ortopedia
- Distribución porcentual por especialidad
- Tendencias mensuales

### Mortalidad
- Total de fallecidos y porcentaje de mortalidad
- Estadísticas de días de estancia (promedio, mínimo, máximo, mediana)
- Mortalidad por trauma
- Pirámide poblacional por edad y sexo

## Requisitos

```
dash>=2.14.0
dash-bootstrap-components>=1.5.0
plotly>=5.18.0
pandas>=2.0.0
openpyxl>=3.1.0
gunicorn>=21.0.0
```

## Instalación

1. Clonar el repositorio:
```bash
git clone https://github.com/acamilox/Dashboard-UCI-Adultos-2025.git dashboard-uci
cd dashboard-uci
```

2. Crear entorno virtual (opcional pero recomendado):
```bash
python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate     # Windows
```

3. Instalar dependencias:
```bash
pip install -r requirements.txt
```

4. Asegurarse de tener el archivo `datos_iaas.xlsx` en el directorio raíz.

## Uso

### Desarrollo local
```bash
python dashboard_uci.py
```
El dashboard estará disponible en `http://localhost:8050`

### Producción
```bash
gunicorn dashboard_uci:server --bind 0.0.0.0:8050
```

## Estructura del Proyecto

```
dashboard-uci-adultos-2025/
├── assets/
│   └── Logo.png
├── dashboard_uci.py      # Aplicación principal
├── datos_iaas.xlsx       # Datos fuente
├── requirements.txt      # Dependencias
├── render.yaml           # Configuración de despliegue
└── README.md
```

## Estructura del Excel

El archivo `datos_iaas.xlsx` debe contener las siguientes hojas:

| Hoja | Descripción |
|------|-------------|
| ESTADISTICA | Indicadores mensuales principales |
| MORBILIDAD | Diagnósticos por sexo |
| IAAS | Infecciones y gérmenes aislados |
| MORTALIDAD | Detalle de pacientes fallecidos |
| EVENTOS ADVERSOS | Registro de eventos |

## Despliegue

El proyecto está configurado para desplegarse en Render. El archivo `render.yaml` contiene la configuración necesaria.

## Tecnologías

- **Backend**: Python, Dash
- **Visualización**: Plotly, Dash Bootstrap Components
- **Datos**: Pandas, OpenPyXL
- **Servidor**: Gunicorn

## Autor

Hospital Infantil Universitario de San José - UCI Adultos

## Licencia

Uso interno - Hospital Infantil Universitario de San José © 2026

