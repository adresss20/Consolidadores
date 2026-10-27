"""
main.py
-------
Punto de entrada del Unificador de Catálogos de Farmacia.

Uso:
    python main.py                        # usa el .xlsx más reciente en catalogos/
    python main.py ruta/al/catalogo.xlsx  # especifica el archivo manualmente

El script:
  1. Lee config.yaml
  2. Carga cada hoja de proveedor del Excel
  3. Consolida todos los SKUs únicos (por EAN)
  4. Calcula rankings de precio (más barato, segundo más barato, # proveedores)
  5. Exporta resultado formateado a output/consolidado_YYYYMMDD_HHMMSS.xlsx
"""

import sys
import logging
from pathlib import Path
from datetime import datetime

from src.config_loader import load_config
from src.loader import load_all_suppliers
from src.merger import merge_suppliers
from src.analyzer import analyze_prices
from src.exporter import export_to_excel

# ── Directorios del proyecto ───────────────────────────────────────────────────
BASE_DIR     = Path(__file__).parent
CATALOGOS_DIR = BASE_DIR / "catalogos"
OUTPUT_DIR   = BASE_DIR / "output"
LOGS_DIR     = BASE_DIR / "logs"
CONFIG_PATH  = BASE_DIR / "config.yaml"


def setup_logging() -> logging.Logger:
    LOGS_DIR.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file  = LOGS_DIR / f"consolidado_{timestamp}.log"

    fmt = "%(asctime)s [%(levelname)-8s] %(name)s: %(message)s"
    logging.basicConfig(
        level=logging.INFO,
        format=fmt,
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    return logging.getLogger("main")


def resolve_input_path(argv: list) -> Path:
    """Determina el archivo de catálogos a usar."""
    if len(argv) > 1:
        path = Path(argv[1])
        if not path.exists():
            raise FileNotFoundError(f"Archivo no encontrado: {path}")
        return path

    # Buscar el .xlsx más reciente en catalogos/
    CATALOGOS_DIR.mkdir(exist_ok=True)
    xlsx_files = sorted(
        CATALOGOS_DIR.glob("*.xlsx"),
        key=lambda f: f.stat().st_mtime,
        reverse=True,
    )
    if not xlsx_files:
        raise FileNotFoundError(
            f"No hay archivos .xlsx en '{CATALOGOS_DIR}'.\n"
            f"Coloca el libro de catálogos ahí o pasa la ruta como argumento:\n"
            f"    python main.py ruta/al/catalogo.xlsx"
        )
    return xlsx_files[0]


def main() -> int:
    """Orquesta todo el flujo. Retorna 0 si exitoso, 1 si error fatal."""
    logger = setup_logging()
    sep = "=" * 60

    logger.info(sep)
    logger.info("  UNIFICADOR DE CATÁLOGOS — FARMACIA")
    logger.info(sep)

    try:
        # ── Paso 1: Resolver archivo de entrada ────────────────────────────────
        input_path = resolve_input_path(sys.argv)
        logger.info(f"Archivo de catálogos: {input_path.resolve()}")

        # ── Paso 2: Cargar configuración ───────────────────────────────────────
        logger.info("\n[1/5] Cargando configuración...")
        config = load_config(str(CONFIG_PATH))

        # ── Paso 3: Cargar hojas de proveedores ────────────────────────────────
        logger.info("\n[2/5] Cargando catálogos de proveedores...")
        supplier_dfs = load_all_suppliers(input_path, config)

        if not supplier_dfs:
            logger.error(
                "No se pudo cargar ningún proveedor.\n"
                "Revisa que los nombres de hojas en config.yaml coincidan con el Excel."
            )
            return 1

        failed = set(config["suppliers"].keys()) - set(supplier_dfs.keys())
        if failed:
            logger.warning(
                f"Proveedores que NO se pudieron cargar (ver errores arriba): {failed}"
            )

        # ── Paso 4: Consolidar ────────────────────────────────────────────────
        logger.info("\n[3/5] Consolidando catálogos...")
        consolidated = merge_suppliers(supplier_dfs, config)

        # ── Paso 5: Analizar precios ──────────────────────────────────────────
        logger.info("\n[4/5] Analizando precios...")
        consolidated = analyze_prices(consolidated, list(supplier_dfs.keys()))

        # ── Paso 6: Exportar ──────────────────────────────────────────────────
        OUTPUT_DIR.mkdir(exist_ok=True)
        timestamp   = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = OUTPUT_DIR / f"consolidado_{timestamp}.xlsx"

        logger.info(f"\n[5/5] Exportando → {output_path.resolve()}")
        export_to_excel(consolidated, output_path, list(supplier_dfs.keys()))

        logger.info(sep)
        logger.info(f"  ✅  LISTO: {output_path.name}")
        logger.info(f"      {len(consolidated):,} SKUs únicos consolidados")
        logger.info(sep)
        return 0

    except FileNotFoundError as exc:
        logger.error(f"\n⛔  Archivo no encontrado:\n   {exc}")
        return 1
    except KeyError as exc:
        logger.error(
            f"\n⛔  Columna no encontrada en el Excel.\n"
            f"   Detalle: {exc}\n"
            f"   Verifica que los nombres de columnas en config.yaml sean correctos."
        )
        return 1
    except Exception as exc:
        logger.error(f"\n⛔  Error inesperado: {exc}", exc_info=True)
        return 1


if __name__ == "__main__":
    sys.exit(main())
