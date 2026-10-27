"""
Punto de entrada único de la automatización.

Uso:
    python main.py

Lee los archivos de scraping más recientes de cada subcarpeta en
`entradas/<farmacia>/`, cruza los catálogos por EAN (o por nombre cuando
no hay EAN), y genera `salida/Catalogo_Consolidado_<fecha_hora>.xlsx`.

Ninguna farmacia con problemas tumba el proceso completo: se excluye esa
farmacia, se deja constancia en el log y en la hoja "Resumen del proceso",
y se sigue con las demás. Solo se aborta si NINGUNA farmacia pudo cargarse.
"""
from __future__ import annotations

import logging
import sys
from datetime import datetime

from config import ENTRADAS_DIR, LOGS_DIR, SALIDA_DIR, FARMACIAS
from consolidar import ORDEN_FARMACIAS, construir_catalogo_consolidado
from io_utils import cargar_farmacia
from matching import emparejar_sin_ean
from reporte import generar_reporte


def _configurar_logging() -> logging.Logger:
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    archivo_log = LOGS_DIR / f"run_{timestamp}.log"

    logger = logging.getLogger("catalogo")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%Y-%m-%d %H:%M:%S")

    fh = logging.FileHandler(archivo_log, encoding="utf-8")
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(fmt)
    logger.addHandler(sh)

    return logger


def main() -> int:
    logger = _configurar_logging()
    logger.info("Iniciando generación del catálogo consolidado...")

    resultados_carga = {}
    por_farmacia = {}

    for farmacia_key in ORDEN_FARMACIAS:
        nombre_visible = FARMACIAS[farmacia_key]["nombre_visible"]
        resultado = cargar_farmacia(farmacia_key, ENTRADAS_DIR)
        resultados_carga[farmacia_key] = resultado

        for w in resultado.warnings:
            logger.warning("[%s] %s", nombre_visible, w)

        if resultado.errores:
            for e in resultado.errores:
                logger.error("[%s] %s", nombre_visible, e)
            logger.error("[%s] Se excluye esta farmacia de la corrida.", nombre_visible)
            continue

        logger.info(
            "[%s] %d productos cargados (EAN válidos: %d, sin EAN: %d, duplicados: %d, precios inválidos: %d)",
            nombre_visible, len(resultado.productos), resultado.ean_validos,
            resultado.ean_invalidos, resultado.duplicados_ean, resultado.precios_invalidos,
        )
        if resultado.errores_scraping_fuente:
            logger.warning(
                "[%s] El propio scraping reportó %d errores (ver hoja 'Errores' del archivo original).",
                nombre_visible, resultado.errores_scraping_fuente,
            )
        por_farmacia[farmacia_key] = resultado.productos

    if not por_farmacia:
        logger.error(
            "No se pudo cargar NINGUNA farmacia. Revisa que los archivos estén en "
            "'entradas/<farmacia>/' y que tengan el formato esperado. Abortando."
        )
        return 1

    logger.info("Buscando coincidencias por nombre para productos sin EAN...")
    resultados_matching, sin_identificar = emparejar_sin_ean(por_farmacia)
    n_asumidos = sum(1 for e in resultados_matching.values())
    logger.info(
        "Emparejamiento por nombre: %d productos asumidos, %d quedaron sin identificar.",
        n_asumidos, len(sin_identificar),
    )

    logger.info("Consolidando catálogo...")
    filas = construir_catalogo_consolidado(por_farmacia, resultados_matching)
    logger.info("Catálogo consolidado: %d productos únicos (por EAN).", len(filas))

    ruta_salida = generar_reporte(filas, sin_identificar, resultados_carga, SALIDA_DIR)
    logger.info("Listo. Archivo generado: %s", ruta_salida)

    return 0


if __name__ == "__main__":
    sys.exit(main())
