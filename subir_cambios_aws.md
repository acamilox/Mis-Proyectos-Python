## Pasos para guardar cambios local y desplegarlos en AWS
1. Hacer git add .
2. Hacer Commit y Push a GitHub
3. Abrir terminal SSH de AWS
4. Ejecutar cd ~/dashboard-uci && git pull && cp dashboard_uci.py app.py && sudo systemctl restart dashboard-uci

## Puerto (IP Estática)
http://52.86.146.83:8050

## Dominio
http://dashboard-uci.com