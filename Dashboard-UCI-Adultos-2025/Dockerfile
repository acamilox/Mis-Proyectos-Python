# Dashboard UCI Adultos - Contenedor Docker
FROM python:3.14.2-slim

WORKDIR /app

# Dependencias del sistema (openpyxl puede necesitar libs para Excel)
RUN apt-get update && apt-get install -y --no-install-recommends \
    && rm -rf /var/lib/apt/lists/*

# Copiar requisitos e instalar dependencias
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copiar código y datos
COPY app.py .
COPY dashboard_uci.py .
COPY dashboard_uci/ ./dashboard_uci/
COPY assets/ ./assets/
COPY data/ ./data/

# Puerto del dashboard
EXPOSE 8050

# Ejecutar con gunicorn (producción)
CMD ["gunicorn", "--bind", "0.0.0.0:8050", "--workers", "1", "--threads", "2", "app:server"]
