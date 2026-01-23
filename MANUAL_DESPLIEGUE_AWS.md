# Manual de Despliegue en AWS Lightsail
## Dashboard UCI Adultos 2025

---

## Índice
1. [Crear cuenta y acceder a AWS](#1-crear-cuenta-y-acceder-a-aws)
2. [Crear instancia en Lightsail](#2-crear-instancia-en-lightsail)
3. [Configurar el firewall](#3-configurar-el-firewall)
4. [Conectarse por SSH](#4-conectarse-por-ssh)
5. [Instalar dependencias del sistema](#5-instalar-dependencias-del-sistema)
6. [Clonar el repositorio](#6-clonar-el-repositorio)
7. [Configurar el entorno Python](#7-configurar-el-entorno-python)
8. [Probar la aplicación](#8-probar-la-aplicación)
9. [Crear servicio systemd](#9-crear-servicio-systemd)
10. [Crear script de actualización](#10-crear-script-de-actualización)
11. [Comandos útiles](#11-comandos-útiles)

---

## 1. Crear cuenta y acceder a AWS

1. Ve a [https://lightsail.aws.amazon.com/](https://lightsail.aws.amazon.com/)
2. Si tienes cuenta AWS, selecciona **"Root user"** e ingresa con tu correo y contraseña
3. Si no tienes cuenta, click en **"Create a new AWS account"**

---

## 2. Crear instancia en Lightsail

1. Click en **"Create instance"**

2. **Instance location**: 
   - Deja Virginia (us-east-1) o selecciona la región más cercana

3. **Pick your instance image**:
   - Platform: **Linux/Unix**
   - Blueprint: Click en **"Operating System (OS) only"**
   - Selecciona: **Ubuntu 24.04 LTS**

4. **Choose your instance plan**:
   - Selecciona **$5 USD/month** (1 GB RAM, 2 vCPUs, 40 GB SSD)
   - Incluye 90 días gratis

5. **Identify your instance**:
   - Nombre: `dashboard-uci` (o el nombre que prefieras)

6. Click en **"Create instance"**

7. Espera a que el estado cambie de "Starting" a **"Running"**

---

## 3. Configurar el firewall

1. Click en el nombre de tu instancia (ej: `dashboard-uci`)
2. Ve a la pestaña **"Networking"**
3. En la sección **"IPv4 Firewall"**, click en **"+ Add rule"**
4. Configura:
   - Application: **Custom**
   - Protocol: **TCP**
   - Port: **8050**
5. Click en **"Create"**

---

## 4. Conectarse por SSH

1. En tu instancia, ve a la pestaña **"Connect"**
2. Click en **"Connect using SSH"**
3. Se abrirá una terminal en el navegador

---

## 5. Instalar dependencias del sistema

Ejecuta este comando:

```bash
sudo apt update && sudo apt install -y python3-pip python3-venv git
```

---

## 6. Clonar el repositorio

### 6.1 Crear Personal Access Token en GitHub

1. Ve a GitHub → Tu perfil → **Settings**
2. Baja hasta **Developer settings** (al final del menú izquierdo)
3. Click en **Personal access tokens** → **Tokens (classic)**
4. Click en **"Generate new token (classic)"**
5. Configura:
   - Note: `AWS Dashboard` (o cualquier nombre)
   - Expiration: Selecciona la duración que prefieras
   - Scopes: Marca **repo** (acceso a repositorios privados)
6. Click en **"Generate token"**
7. **IMPORTANTE**: Copia el token, solo lo verás una vez

### 6.2 Clonar el repositorio

```bash
git clone https://github.com/TU_USUARIO/TU_REPOSITORIO.git dashboard-uci
```

Cuando te pida:
- **Username**: Tu usuario de GitHub
- **Password**: Pega el token que copiaste (no se verá mientras escribes)

---

## 7. Configurar el entorno Python

Ejecuta estos comandos uno por uno:

```bash
cd dashboard-uci
```

```bash
python3 -m venv venv
```

```bash
source venv/bin/activate
```

```bash
pip install -r requirements.txt
```

---

## 8. Probar la aplicación

```bash
python dashboard_uci.py
```

Abre en tu navegador: `http://TU_IP_PUBLICA:8050`

Si funciona, presiona **Ctrl+C** para detenerlo.

---

## 9. Crear servicio systemd

Esto hace que la aplicación corra automáticamente y se reinicie si el servidor se reinicia.

### 9.1 Crear el archivo de servicio

```bash
sudo nano /etc/systemd/system/dashboard-uci.service
```

Pega este contenido:

```
[Unit]
Description=Dashboard UCI Adultos
After=network.target

[Service]
User=ubuntu
WorkingDirectory=/home/ubuntu/dashboard-uci
Environment="PATH=/home/ubuntu/dashboard-uci/venv/bin"
ExecStart=/home/ubuntu/dashboard-uci/venv/bin/gunicorn dashboard_uci:server -b 0.0.0.0:8050 --workers 2 --timeout 120
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
```

Guarda con **Ctrl+O**, **Enter**, y sal con **Ctrl+X**

### 9.2 Activar el servicio

```bash
sudo systemctl daemon-reload
```

```bash
sudo systemctl enable dashboard-uci
```

```bash
sudo systemctl start dashboard-uci
```

### 9.3 Verificar que esté corriendo

```bash
sudo systemctl status dashboard-uci
```

Debe decir **"active (running)"** en verde.

---

## 10. Crear script de actualización

Este script facilita actualizar la aplicación cuando hagas cambios.

```bash
nano ~/actualizar.sh
```

Pega este contenido:

```bash
#!/bin/bash
cd ~/dashboard-uci
git pull
sudo systemctl restart dashboard-uci
echo "Dashboard actualizado!"
```

Guarda con **Ctrl+O**, **Enter**, y sal con **Ctrl+X**

Hazlo ejecutable:

```bash
chmod +x ~/actualizar.sh
```

### Guardar credenciales de Git (opcional)

Para no tener que escribir usuario y token cada vez:

```bash
git config --global credential.helper store
```

La próxima vez que ejecutes el script y pongas tus credenciales, se guardarán.

---

## 11. Comandos útiles

### Ver estado del servicio
```bash
sudo systemctl status dashboard-uci
```

### Reiniciar el servicio
```bash
sudo systemctl restart dashboard-uci
```

### Detener el servicio
```bash
sudo systemctl stop dashboard-uci
```

### Ver logs del servicio
```bash
sudo journalctl -u dashboard-uci -f
```

### Actualizar la aplicación (después de hacer push a GitHub)
```bash
~/actualizar.sh
```

---

## Resumen del flujo de trabajo

1. Haces cambios en tu código local
2. Haces commit y push a GitHub
3. Te conectas por SSH a AWS (pestaña Connect → Connect using SSH)
4. Ejecutas `~/actualizar.sh`
5. Refrescas el navegador (F5)

---

## Datos de tu instancia

- **IP Pública**: 18.207.148.186
- **URL del Dashboard**: http://18.207.148.186:8050
- **Región**: Virginia (us-east-1)
- **Plan**: $5 USD/mes (90 días gratis)

---

## Notas importantes

- El dashboard sigue corriendo aunque cierres la terminal SSH
- Si el servidor se reinicia, el dashboard se inicia automáticamente
- La IP pública puede cambiar si detienes y reinicias la instancia (puedes asignar una IP estática en Lightsail si lo necesitas)
