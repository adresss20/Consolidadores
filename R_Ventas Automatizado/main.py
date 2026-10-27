# -*- coding: utf-8 -*-
"""
main.py - Punto de entrada del proceso de actualizacion de ventas.

FLUJO DE USO (Fase 1 - descarga manual):
  1. Descarga del ERP los 3 reportes (ver "Proceso manual.txt", pasos 1-4):
        - Ventas Diarias  -> REPORTE_VENTAS_DIARIAS_*.xlsx
        - Utilidad        -> REPORTE_UTILIDAD_*.xlsx
        - Consulta Ventas -> ConsultaVentas*.xlsx
     usando el rango de fechas [ultima registrada + 1 .. ayer].
  2. Coloca los 3 archivos en la carpeta  input/
  3. Ejecuta:   python main.py
  4. El consolidado se actualiza en sitio (con backup previo automatico).

No hace falta borrar los archivos de input/: se toma siempre el mas reciente
por patron de nombre, y el anexado es idempotente (no duplica ventas ya
registradas).
"""

from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

import consolidador as C

# --------------------------------------------------------------------------- #
#  Configuracion (ajusta rutas aqui si cambias la estructura)
# --------------------------------------------------------------------------- #
BASE = Path(__file__).resolve().parent
INPUT_DIR = BASE / "input"
BACKUPS_DIR = BASE / "backups"
CONSOLIDADO = BASE / "Reporte consolidado de ventas.xlsx"

# Patrones para localizar cada reporte dentro de input/
PATRONES = {
    "ventas": "REPORTE_VENTAS_DIARIAS*.xlsx",
    "utilidad": "REPORTE_UTILIDAD*.xlsx",
    "consulta": "ConsultaVentas*.xlsx",
}


def _mas_reciente(patron: str) -> Path | None:
    """Devuelve el archivo mas reciente (por fecha de modificacion) que
    coincide con el patron dentro de input/, ignorando temporales de Excel."""
    candidatos = [
        p for p in INPUT_DIR.glob(patron)
        if not p.name.startswith("~$")
    ]
    if not candidatos:
        return None
    return max(candidatos, key=lambda p: p.stat().st_mtime)


def localizar_reportes() -> dict[str, Path]:
    faltan = []
    encontrados: dict[str, Path] = {}
    for clave, patron in PATRONES.items():
        p = _mas_reciente(patron)
        if p is None:
            faltan.append(f"{clave} (patron '{patron}')")
        else:
            encontrados[clave] = p
    if faltan:
        print("ERROR: faltan reportes en la carpeta input/:")
        for f in faltan:
            print("   -", f)
        print(f"\nColoca los archivos descargados en: {INPUT_DIR}")
        sys.exit(1)
    return encontrados


def main() -> None:
    if not CONSOLIDADO.exists():
        print(f"ERROR: no se encontro el consolidado:\n   {CONSOLIDADO}")
        sys.exit(1)

    reportes = localizar_reportes()
    print("Reportes detectados en input/:")
    for clave, p in reportes.items():
        print(f"   {clave:9s}: {p.name}")
    print()

    ventas = C.leer_ventas_diarias(reportes["ventas"])
    consulta = C.cargar_consulta_ventas(reportes["consulta"])
    utilidad = C.cargar_utilidad(reportes["utilidad"])

    res = C.actualizar_consolidado(
        consolidado_path=CONSOLIDADO,
        ventas=ventas,
        consulta=consulta,
        utilidad=utilidad,
        fecha_corte=date.today(),   # toma ventas hasta ayer (excluye hoy)
        backups_dir=BACKUPS_DIR,
    )

    print(res.resumen())

    if res.filas_nuevas == 0:
        print("\nEl consolidado ya estaba al dia. No se anexaron filas.")
    else:
        print(f"\nListo. Se anexaron {res.filas_nuevas} lineas al consolidado.")


if __name__ == "__main__":
    main()
