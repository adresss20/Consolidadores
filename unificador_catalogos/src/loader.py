"""
loader.py
---------
Lee cada hoja de proveedor del libro Excel y la convierte a un DataFrame
estandarizado con columnas canónicas.

DataFrame de salida por proveedor:
  ean, precio, descripcion, laboratorio, familia, generico, sustancia, presentacion
"""

import pandas as pd
import logging
from pathlib import Path

from src.normalizer import (
    normalize_ean,
    normalize_price,
    normalize_text,
    normalize_generico,
)

logger = logging.getLogger(__name__)


def load_all_suppliers(excel_path: Path, config: dict) -> dict:
    """
    Carga todos los proveedores definidos en config.
    Retorna dict { nombre_proveedor: DataFrame }.
    Los proveedores que fallen se omiten con error en log pero NO detienen el proceso.
    """
    suppliers = {}
    for name, cfg in config["suppliers"].items():
        logger.info(f"  ► Cargando {name!r} (hoja: {cfg['sheet']!r})...")
        try:
            df = _load_supplier_sheet(excel_path, name, cfg)
            suppliers[name] = df
            logger.info(
                f"    ✓ {len(df)} productos con EAN y precio válidos"
            )
        except Exception as exc:
            logger.error(
                f"    ✗ Error al cargar {name!r}: {exc}",
                exc_info=True,
            )
    return suppliers


def _load_supplier_sheet(excel_path: Path, name: str, cfg: dict) -> pd.DataFrame:
    """Lee una hoja del Excel y normaliza columnas al esquema canónico."""

    # --- Lectura raw (todo como string para evitar conversiones automáticas) ---
    try:
        raw = pd.read_excel(
            excel_path,
            sheet_name=cfg["sheet"],
            dtype=str,
            engine="openpyxl",
        )
    except Exception as exc:
        raise RuntimeError(
            f"No se pudo leer la hoja '{cfg['sheet']}' del archivo. "
            f"Verifica que la hoja exista y no esté protegida. Detalle: {exc}"
        ) from exc

    if raw.empty:
        raise ValueError(f"La hoja '{cfg['sheet']}' está vacía.")

    # Strip espacios en nombres de columnas (FarmaMayoreo tiene 'PRECIO ')
    raw.columns = [str(c).strip() for c in raw.columns]

    # En caso de columnas duplicadas, pandas las renombra como 'col', 'col.1', etc.
    # El config siempre apunta al nombre sin sufijo → tomamos la primera ocurrencia.

    # --- Columna EAN / SKU (obligatoria) ---
    sku_col = cfg["sku_col"].strip()
    if sku_col not in raw.columns:
        raise KeyError(
            f"Columna EAN '{sku_col}' no encontrada en la hoja '{cfg['sheet']}'. "
            f"Columnas disponibles: {list(raw.columns)}"
        )

    # --- Columna Precio (obligatoria) ---
    price_col = cfg["price_col"].strip()
    if price_col not in raw.columns:
        raise KeyError(
            f"Columna precio '{price_col}' no encontrada en la hoja '{cfg['sheet']}'. "
            f"Columnas disponibles: {list(raw.columns)}"
        )

    # --- Construir DataFrame canónico ---
    canonical = pd.DataFrame()
    canonical["ean"] = raw[sku_col].apply(normalize_ean)
    canonical["precio"] = raw[price_col].apply(normalize_price)

    # Campos informativos opcionales
    optional_map = {
        "descripcion": ("desc_col", normalize_text),
        "laboratorio": ("lab_col", normalize_text),
        "familia": ("familia_col", normalize_text),
        "generico": ("generico_col", normalize_generico),
        "sustancia": ("sustancia_col", normalize_text),
        "presentacion": ("presentacion_col", normalize_text),
    }
    for canonical_name, (config_key, normalizer_fn) in optional_map.items():
        col = cfg.get(config_key)
        if col and col.strip() in raw.columns:
            canonical[canonical_name] = raw[col.strip()].apply(normalizer_fn)
        else:
            canonical[canonical_name] = pd.NA

    # --- Filtrar filas sin EAN ---
    before = len(canonical)
    canonical = canonical[canonical["ean"].notna()].copy()
    dropped_ean = before - len(canonical)
    if dropped_ean:
        logger.warning(f"    {name}: {dropped_ean} filas sin EAN descartadas")

    # --- Filtrar filas sin precio ---
    before = len(canonical)
    canonical = canonical[canonical["precio"].notna()].copy()
    dropped_price = before - len(canonical)
    if dropped_price:
        logger.info(
            f"    {name}: {dropped_price} filas sin precio (no aparecerán en precios)"
        )

    if canonical.empty:
        raise ValueError(
            f"El proveedor '{name}' no tiene filas válidas con EAN y precio."
        )

    # --- Deduplicar EANs (mismo EAN con precios distintos → quedamos con el más bajo) ---
    dups_mask = canonical.duplicated(subset=["ean"], keep=False)
    n_dups = dups_mask.sum()
    if n_dups:
        logger.warning(
            f"    {name}: {n_dups} filas con EAN duplicado — "
            f"se conserva el precio más bajo por EAN"
        )
        canonical = (
            canonical
            .sort_values("precio")
            .drop_duplicates(subset=["ean"], keep="first")
        )

    return canonical.reset_index(drop=True)
