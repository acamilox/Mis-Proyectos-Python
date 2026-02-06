# Manual de Despliegue en AWS Lightsail
## Dashboard UCI Adultos 2026

---

## Información del Servidor

| Dato | Valor |
|------|-------|
| **IP Estática** | 52.86.146.83 |
| **Dominio** | http://dashboard-uci.com |
| **Puerto directo** | http://52.86.146.83:8050 |
| **Región** | Virginia (us-east-1) |
| **Plan** | $5 USD/mes |

---

## Flujo de Trabajo Diario

### Subir cambios al servidor

1. **Local**: Hacer commit y push
```bash
git add .
git commit -m "descripción del cambio"
git push
```

2. **AWS**: Conectar por SSH y actualizar
```bash
cd ~/dashboard-uci && git pull && cp dashboard_uci.py app.py && sudo systemctl restart dashboard-uci
```

O simplemente:
```bash
~/actualizar.sh
```

> **Nota**: En el servidor el archivo se llama `app.py` (renombrado para evitar conflicto con la carpeta `dashboard_uci/`)

---

## Comandos Útiles

### Estado del servicio
```bash
sudo systemctl status dashboard-uci
```

### Reiniciar servicio
```bash
sudo systemctl restart dashboard-uci
```

### Ver logs en tiempo real
```bash
sudo journalctl -u dashboard-uci -f
```

### Ver últimos 30 logs
```bash
sudo journalctl -u dashboard-uci -n 30
```

---

## Configuración Inicial (Solo primera vez)

### 1. Crear instancia en Lightsail

1. Ve a [https://lightsail.aws.amazon.com/](https://lightsail.aws.amazon.com/)
2. Click en **"Create instance"**
3. Configurar:
   - Platform: **Linux/Unix**
   - Blueprint: **Ubuntu 24.04 LTS**
   - Plan: **$5 USD/month** (1 GB RAM, 2 vCPUs)
   - Nombre: `dashboard-uci`

### 2. Configurar firewall

En la pestaña **Networking** → **IPv4 Firewall**:
- Agregar regla TCP puerto **8050**
- Agregar regla TCP puerto **80** (para nginx)

### 3. Asignar IP estática

En **Networking** → **Public IP** → **Create static IP**

### 4. Instalar dependencias

```bash
sudo apt update && sudo apt install -y python3-pip python3-venv git nginx
```

### 5. Clonar repositorio

```bash
git clone https://github.com/acamilox/Dashboard-UCI-Adultos-2025.git dashboard-uci
cd dashboard-uci
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

### 6. Crear servicio systemd

```bash
sudo nano /etc/systemd/system/dashboard-uci.service
```

Contenido:
```ini
[Unit]
Description=Dashboard UCI Adultos
After=network.target

[Service]
User=ubuntu
WorkingDirectory=/home/ubuntu/dashboard-uci
Environment="PATH=/home/ubuntu/dashboard-uci/venv/bin"
ExecStart=/home/ubuntu/dashboard-uci/venv/bin/gunicorn app:server -b 0.0.0.0:8050 --workers 2 --timeout 120
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
```

Activar:
```bash
sudo systemctl daemon-reload
sudo systemctl enable dashboard-uci
sudo systemctl start dashboard-uci
```

### 7. Configurar nginx (proxy para dominio)

```bash
sudo nano /etc/nginx/sites-available/dashboard-uci
```

Contenido:
```nginx
server {
    listen 80;
    server_name dashboard-uci.com;

    location / {
        proxy_pass http://127.0.0.1:8050;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
    }
}
```

Activar:
```bash
sudo ln -s /etc/nginx/sites-available/dashboard-uci /etc/nginx/sites-enabled/
sudo nginx -t
sudo systemctl restart nginx
```

### 8. Crear script de actualización

```bash
nano ~/actualizar.sh
```

Contenido:
```bash
#!/bin/bash
cd ~/dashboard-uci
git pull
cp dashboard_uci.py app.py
sudo systemctl restart dashboard-uci
echo "Dashboard actualizado!"
```

```bash
chmod +x ~/actualizar.sh
```

---

## Solución de Problemas

### El dashboard no carga
```bash
sudo systemctl status dashboard-uci
sudo journalctl -u dashboard-uci -n 50
```

### Error 502 Bad Gateway
El servicio gunicorn no está corriendo:
```bash
sudo systemctl restart dashboard-uci
```

### Servidor se cae por memoria
El plan de $5 tiene solo 1GB RAM. Si se cae frecuentemente, considera upgrade a $10/mes.

### Los cambios no se reflejan
1. Verificar que el push llegó: `git log --oneline -3`
2. Copiar archivo: `cp dashboard_uci.py app.py`
3. Reiniciar: `sudo systemctl restart dashboard-uci`
4. Limpiar caché del navegador: Ctrl+F5
