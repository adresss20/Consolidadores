"""
analyzer.py
-----------
Calcula métricas de precio por SKU:
  - NUM_PROVEEDORES    → cuántos proveedores tienen precio para ese SKU
  - PROVEEDOR_MAS_BARATO  → nombre del proveedor con el precio más bajo
  - PRECIO_MAS_BARATO     → precio más bajo
  - SEGUNDO_MAS_BARATO    → nombre del segundo proveedor más económico
  - PRECIO_SEGUNDO        → su precio
  - AHORRO_VS_SEGUNDO     → diferencia de precio entre 1ro y 2do (cuánto se ahorra eligiendo el más barato)
"""

import pandas as pd
import logging

logger = logging.getLogger(__name__)


def analyze_prices(df: pd.DataFrame, supplier_names: list) -> pd.DataFrame:
    """
    Agrega columnas de análisis de precio al DataFrame consolidado.

    Parámetros:
        df:             DataFrame con columna precio_<nombre> por proveedor
        supplier_names: lista de nombres de proveedores en el orden del config

    Retorna:
        DataFrame enriquecido con columnas de ranking.
    """
    price_cols = {name: f"precio_{name}" for name in supplier_names}

    # Construir matriz de precios para vectorizar el cálculo
    price_matrix = df[[col for col in price_cols.values() if col in df.columns]].copy()
    price_matrix.columns = [name for name, col in price_cols.items() if col in df.columns]
    valid_suppliers = list(price_matrix.columns)

    # Número de proveedores con precio
    df["NUM_PROVEEDORES"] = price_matrix.notna().sum(axis=1).astype("Int64")

    # Calcular ranking fila por fila (eficiente para N pequeño de proveedores)
    rankings = price_matrix.apply(_rank_prices, axis=1, result_type="expand")
    rankings.columns = [
        "PROVEEDOR_MAS_BARATO",
        "PRECIO_MAS_BARATO",
        "SEGUNDO_MAS_BARATO",
        "PRECIO_SEGUNDO",
        "AHORRO_VS_SEGUNDO",
    ]

    df = pd.concat([df, rankings], axis=1)

    # Estadísticas de cobertura para el log
    n_total = len(df)
    n_multi = (df["NUM_PROVEEDORES"] > 1).sum()
    n_single = (df["NUM_PROVEEDORES"] == 1).sum()
    n_none = (df["NUM_PROVEEDORES"] == 0).sum()

    logger.info(
        f"Análisis de precios:\n"
        f"  Total SKUs:               {n_total:,}\n"
        f"  Con 2+ proveedores:       {n_multi:,}\n"
        f"  Con 1 proveedor:          {n_single:,}\n"
        f"  Sin precio en ninguno:    {n_none:,}"
    )

    return df


def _rank_prices(row: pd.Series):
    """
    Para una fila de precios {proveedor: precio}, retorna los 2 más baratos.
    Diseñado para usarse con .apply(axis=1).
    """
    # Filtrar solo precios válidos (no NaN)
    available = {
        supplier: price
        for supplier, price in row.items()
        if pd.notna(price) and price > 0
    }

    if not available:
        return pd.NA, pd.NA, pd.NA, pd.NA, pd.NA

    # Ordenar por precio ascendente
    sorted_prices = sorted(available.items(), key=lambda x: x[1])

    prov_1 = sorted_prices[0][0]
    precio_1 = round(sorted_prices[0][1], 2)

    if len(sorted_prices) > 1:
        prov_2 = sorted_prices[1][0]
        precio_2 = round(sorted_prices[1][1], 2)
        ahorro = round(precio_2 - precio_1, 2)
    else:
        prov_2 = pd.NA
        precio_2 = pd.NA
        ahorro = pd.NA

    return prov_1, precio_1, prov_2, precio_2, ahorro
