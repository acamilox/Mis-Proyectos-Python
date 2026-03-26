#!/usr/bin/env bash
set -euo pipefail

REPO_DIR="/home/devuser/dashboard-uci"
SERVICE_NAME="dashboard-uci"

cd "$REPO_DIR"

# Obtener el estado remoto sin forzar un pull en cada tick.
git fetch origin --quiet

LOCAL_REV="$(git rev-parse HEAD)"
REMOTE_REV="$(git rev-parse origin/main)"

if [[ "$LOCAL_REV" == "$REMOTE_REV" ]]; then
  exit 0
fi

# Traer cambios e intentar rebase local para mantener historia limpia.
git pull --rebase --quiet

# Reiniciar para que cambios (incl. assets CSS) queden reflejados.
systemctl restart "$SERVICE_NAME"

