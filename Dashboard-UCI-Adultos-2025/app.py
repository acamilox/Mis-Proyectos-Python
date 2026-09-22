# -*- coding: utf-8 -*-
"""
Entry point para gunicorn.
Evita conflicto de nombres: Python importa el paquete dashboard_uci/ en lugar del script.
Este módulo carga el script principal y re-exporta app y server.
"""
import importlib.util
import sys
from pathlib import Path

# Cargar dashboard_uci.py (el script, no el paquete)
script_path = Path(__file__).resolve().parent / "dashboard_uci.py"
spec = importlib.util.spec_from_file_location("_main_script", script_path)
_mod = importlib.util.module_from_spec(spec)
sys.modules["_main_script"] = _mod
spec.loader.exec_module(_mod)

app = _mod.app
server = _mod.server
