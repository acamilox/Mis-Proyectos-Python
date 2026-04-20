# -*- coding: utf-8 -*-
"""
Script de verificación: compara datos del Excel Estadistica 2026.xlsx
con lo que muestra el Dashboard UCI.
"""
import sys
import io
# Evitar problemas de encoding en Windows
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import pandas as pd
from openpyxl import load_workbook

EXCEL_2026 = ROOT / "data" / "Estadistica 2026.xlsx"
MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Agos", "Sep", "Oct", "Nov", "Dic"]
MESES_SHEET = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
               "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]


def load_estadistica(path):
    df = pd.read_excel(path, sheet_name="ESTADISTICA", header=2)
    df = df.set_index("DATOS")
    df = df.replace(["*", "#DIV/0!"], None)
    for col in df.columns:
        if col not in ["PROMEDIO", None]:
            try:
                df[col] = pd.to_numeric(df[col], errors="coerce")
            except (ValueError, TypeError):
                pass
    return df


def get_suma_col(df):
    for c in df.columns:
        if str(c).strip() in ["∑", "SUMA"]:
            return c
    return "∑"


def main():
    if not EXCEL_2026.exists():
        print(f"ERROR: No existe {EXCEL_2026}")
        return 1

    print("=" * 70)
    print("VERIFICACIÓN EXCEL 2026 vs DASHBOARD")
    print("=" * 70)

    # 1. Cargar ESTADISTICA
    df = load_estadistica(EXCEL_2026)
    suma_col = get_suma_col(df)

    print("\n--- HOJA ESTADISTICA ---")
    print(f"  Filas: {len(df)}, Num columnas: {len(df.columns)}")

    # Filas clave y sus valores
    filas = ["NUMERO DE INGRESOS DEL MES", "NÚMERO DE EGRESOS DEL MES",
             "FALLECIDOS", "PROMEDIO ESTANCIA", "GIRO CAMA", "% OCUPACIONAL",
             "% DE MORTALIDAD GENERAL", "TOTAL DIAS ESTANCIA", "CAPACIDAD INSTALADA: UCI"]
    for nombre in filas:
        idx_matches = [i for i in df.index if str(i).strip().upper() == nombre.upper() or nombre.upper() in str(i).upper()]
        if idx_matches:
            row = df.loc[idx_matches[0]]
            if suma_col in df.columns:
                val = row.get(suma_col)
            else:
                val = pd.to_numeric(row, errors="coerce").fillna(0).sum()
            print(f"  {nombre}: {val}")
        else:
            print(f"  {nombre}: [no encontrada]")

    # 2. Leer MORTALIDAD
    print("\n--- HOJA MORTALIDAD ---")
    try:
        df_mort = pd.read_excel(EXCEL_2026, sheet_name="MORTALIDAD", header=None, usecols="A:F")
        rows_data = 0
        for _, r in df_mort.iterrows():
            if pd.notna(r[0]) and str(r[0]).strip() and pd.notna(r[1]) and isinstance(r[1], (int, float)):
                if str(r[0]).upper() not in MESES_SHEET:
                    rows_data += 1
        print(f"  Filas con nombre+edad: {rows_data}")
    except Exception as e:
        print(f"  Error: {e}")

    # 3. Leer MORBILIDAD
    print("\n--- HOJA MORBILIDAD ---")
    try:
        df_morb = pd.read_excel(EXCEL_2026, sheet_name="MORBILIDAD", header=None)
        total_morb = pd.to_numeric(df_morb[3], errors="coerce").fillna(0).sum()
        print(f"  Total columna 3 (Total): {int(total_morb)}")
    except Exception as e:
        print(f"  Error: {e}")

    # 4. Leer IAAS
    print("\n--- HOJA IAAS ---")
    try:
        df_iaas = pd.read_excel(EXCEL_2026, sheet_name="IAAS", header=None)
        if 5 in df_iaas.columns:
            n_germenes = df_iaas[5].dropna().astype(str).str.strip()
            n_germenes = n_germenes[(n_germenes != "") & (n_germenes.str.lower() != "nan")]
            print(f"  Filas con gérmenes (formato clásico): {len(n_germenes)}")
        elif 4 in df_iaas.columns:
            filas_iaas = 0
            categorias = set()
            for _, row in df_iaas.iterrows():
                c0 = row[0] if 0 in df_iaas.columns else None
                c1 = row[1] if 1 in df_iaas.columns else None
                c4 = row[4]
                if pd.isna(c0) or not str(c0).strip():
                    continue
                # Saltar separadores de mes
                if str(c0).strip().upper() in MESES_SHEET:
                    continue
                if pd.isna(pd.to_numeric(c1, errors='coerce')):
                    continue
                if pd.isna(c4) or not str(c4).strip():
                    continue
                val = str(c4).strip()
                if val.lower() == "nan":
                    continue
                filas_iaas += 1
                categorias.add(val)
            print(f"  Filas IAAS (formato nuevo): {filas_iaas}")
            print(f"  Categorías IAAS únicas: {len(categorias)}")
        else:
            print("  IAAS: formato no reconocido (sin col 4 ni col 5)")
    except Exception as e:
        print(f"  Error: {e}")

    # 5. Leer EVENTOS ADVERSOS
    print("\n--- HOJA EVENTOS ADVERSOS ---")
    try:
        wb = load_workbook(EXCEL_2026, read_only=True, data_only=True)
        ws = wb["EVENTOS ADVERSOS"]
        count = 0
        for i in range(1, 300):
            c1 = ws.cell(i, 1).value
            c2 = ws.cell(i, 2).value
            c5 = ws.cell(i, 5).value
            if c1 and str(c1).strip() and pd.notna(c2) and c5 and "tipo de evento" not in str(c5).lower():
                if str(c1).upper() not in MESES_SHEET:
                    count += 1
        wb.close()
        print(f"  Filas con evento: {count}")
    except Exception as e:
        print(f"  Error: {e}")

    # 6. Usar DataLoader si está disponible
    print("\n--- DATOS SEGÚN DASHBOARD (DataLoader) ---")
    try:
        from dashboard_uci.loaders import DataLoader
        from dashboard_uci.processors import DataProcessor
        loader = DataLoader(EXCEL_2026)
        data = loader.load_all()
        proc = DataProcessor(data)
        kpis = proc.get_kpis_resumen(MESES)
        print(f"  Total ingresos: {kpis['total_ingresos']}")
        print(f"  Total egresos: {kpis['total_egresos']}")
        print(f"  Total fallecidos: {kpis['total_fallecidos']}")
        print(f"  % Ocupación promedio: {kpis['promedio_ocupacional']:.1f}%")
        print(f"  Estancia promedio: {kpis['promedio_estancia']:.2f}")
        print(f"  % Mortalidad: {kpis['mortalidad']:.1f}%")
        print(f"  Fallecidos (hoja MORTALIDAD): {len(data['fallecidos'])}")
        print(f"  Eventos adversos: {data['eventos_adversos']['Cantidad'].sum() if not data['eventos_adversos'].empty else 0}")
        print(f"  IAAS total: {data['germenes']['Cantidad'].sum() if not data['germenes'].empty else 0}")
    except ImportError as e:
        print("  (Omitido: módulo dashboard_uci no disponible)")

    print("\n" + "=" * 70)
    return 0


if __name__ == "__main__":
    sys.exit(main())
