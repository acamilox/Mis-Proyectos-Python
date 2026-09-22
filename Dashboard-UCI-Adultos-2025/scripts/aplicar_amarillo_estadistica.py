# -*- coding: utf-8 -*-
"""Aplica en ESTADISTICA solo los valores de celdas marcadas en amarillo (theme 4)."""
from __future__ import annotations

import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
EXCEL = ROOT / "data" / "Estadistica 2026.xlsx"

# Cambios de datos en amarillo (fila, col, valor) — Abr 2026 y demás según Excel
YELLOW_DATA_UPDATES = [
    (88, 5, 0),   # NÚMERO DE BACTERIEMIAS — Abr
    (93, 5, 0),   # NÚMERO CASOS INFECCIÓN ASOC. A C. URINARIO — Abr
    (95, 5, 1),   # INFECCION DEL TORRENTE SANGUINEO — Abr
]


def _col_letter(col: int) -> str:
    s = ""
    while col:
        col, r = divmod(col - 1, 26)
        s = chr(65 + r) + s
    return s


def _is_yellow(cell) -> bool:
    fill = cell.fill
    if not fill or fill.fill_type != "solid":
        return False
    fg = fill.fgColor
    if fg is None:
        return False
    if getattr(fg, "theme", None) == 4:
        return True
    rgb = str(getattr(fg, "rgb", "") or "").upper()
    return any(x in rgb for x in ("FFFF00", "FFEB9C", "FFF2CC", "FFC000", "FFFFFF99"))


def discover_yellow_updates(source: Path) -> list[tuple[int, int, object]]:
    """Lee celdas amarillas en ESTADISTICA (solo columnas de meses, col >= 2)."""
    wb = openpyxl.load_workbook(source, data_only=True)
    wb_fmt = openpyxl.load_workbook(source, data_only=False)
    ws, ws_fmt = wb["ESTADISTICA"], wb_fmt["ESTADISTICA"]
    updates: list[tuple[int, int, object]] = []
    for r in range(4, ws.max_row + 1):
        for c in range(2, 15):
            if not _is_yellow(ws_fmt.cell(r, c)):
                continue
            val = ws.cell(r, c).value
            if val is None:
                continue
            updates.append((r, c, val))
    wb.close()
    wb_fmt.close()
    return updates


def patch_sheet_xml(sheet_xml: str, updates: list[tuple[int, int, object]]) -> str:
    for r, c, val in updates:
        ref = f"{_col_letter(c)}{r}"
        new_cell = f'<c r="{ref}"><v>{val}</v></c>'
        block = re.search(rf'<c r="{re.escape(ref)}"[^>]*>.*?</c>', sheet_xml, re.DOTALL)
        empty = re.search(rf'<c r="{re.escape(ref)}"[^>]*/>', sheet_xml)
        if block:
            sheet_xml = sheet_xml[: block.start()] + new_cell + sheet_xml[block.end() :]
        elif empty:
            sheet_xml = sheet_xml[: empty.start()] + new_cell + sheet_xml[empty.end() :]
        else:
            print(f"  AVISO: celda {ref} no encontrada en XML", file=sys.stderr)
    return sheet_xml


def apply_yellow_to_base(base: Path, updates: list[tuple[int, int, object]], out: Path) -> None:
    tmpdir = Path(tempfile.mkdtemp())
    try:
        with zipfile.ZipFile(base, "r") as zin:
            zin.extractall(tmpdir)
        # ESTADISTICA suele ser sheet1 en este libro
        sheet_path = tmpdir / "xl" / "worksheets" / "sheet1.xml"
        xml = sheet_path.read_text(encoding="utf-8")
        sheet_path.write_text(patch_sheet_xml(xml, updates), encoding="utf-8")
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for f in tmpdir.rglob("*"):
                if f.is_file():
                    zout.write(f, f.relative_to(tmpdir).as_posix())
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def main() -> int:
    if not EXCEL.exists():
        print(f"No existe {EXCEL}", file=sys.stderr)
        return 1

    # Fuente de marcas amarillas: archivo actual del usuario + celdas Abr obligatorias
    discovered = discover_yellow_updates(EXCEL)
    by_rc = {(r, c): v for r, c, v in YELLOW_DATA_UPDATES}
    for r, c, v in discovered:
        by_rc[(r, c)] = v
    updates = sorted(by_rc.items(), key=lambda x: (x[0][0], x[0][1]))
    updates = [(r, c, v) for (r, c), v in updates]
    print(f"Celdas a aplicar: {len(updates)} (amarillo detectado: {len(discovered)})")

    # Base = último commit (conserva fórmulas y caché de valores)
    base_bytes = subprocess.check_output(["git", "show", "HEAD:data/Estadistica 2026.xlsx"], cwd=ROOT)
    tmp = Path(tempfile.mkdtemp())
    base_path = tmp / "base.xlsx"
    base_path.write_bytes(base_bytes)

    out_path = tmp / "out.xlsx"
    apply_yellow_to_base(base_path, updates, out_path)
    shutil.copy2(out_path, EXCEL)
    shutil.rmtree(tmp, ignore_errors=True)

    for r, c, v in updates:
        print(f"  {_col_letter(c)}{r} = {v}")
    print(f"Guardado: {EXCEL}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
