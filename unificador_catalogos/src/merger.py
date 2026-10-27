"""
merger.py
---------
Une todos los DataFrames de proveedores en un único consolidado:
  - Un row por EAN único (unión de todos los catálogos)
  - Una columna de precio por proveedor
  - Campos informativos llenados por prioridad (definida en config)
"""

import pandas as pd
import logging

logger = logging.getLogger(__name__)

INFO_FIELDS = [
    "descripcion",
    "laboratorio",
    "familia",
    "generico",
    "sustancia",
    "presentacion",
]


def merge_suppliers(supplier_dfs: dict, config: dict) -> pd.DataFrame:
    """
    Parámetros:
        supplier_dfs: dict { nombre_proveedor: DataFrame_canónico }
        config:       configuración completa (para leer info_priority)

    Retorna:
        DataFrame con columnas:
          ean, descripcion, laboratorio, familia, generico, sustancia, presentacion,
          precio_<Proveedor1>, precio_<Proveedor2>, ...
    """
    priority = config.get("info_priority", list(supplier_dfs.keys()))

    # 1. Recolectar todos los EANs únicos
    all_eans: set = set()
    for name, df in supplier_dfs.items():
        all_eans.update(df["ean"].dropna().unique())

    logger.info(f"EANs únicos en todos los catálogos: {len(all_eans):,}")

    master = pd.DataFrame({"ean": sorted(all_eans)})

    # 2. Inicializar columnas de info en blanco
    for field in INFO_FIELDS:
        master[field] = pd.NA

    # 3. Llenar info por prioridad (de MENOR a MAYOR para que la más alta gane)
    #    El truco: iteramos en orden inverso de prioridad; el último (mayor prioridad)
    #    sobreescribe a los anteriores.
    fill_order = list(reversed(priority))  # menor primero, mayor al final

    for name in fill_order:
        if name not in supplier_dfs:
            continue
        df = supplier_dfs[name]
        ean_index = df.set_index("ean")

        for field in INFO_FIELDS:
            if field not in ean_index.columns:
                continue

            # Serie con valores del proveedor, indexada por EAN
            supplier_values = ean_index[field]

            # Mapeo EAN → valor
            mapping = (
                supplier_values
                .dropna()
                .to_dict()
            )
            if not mapping:
                continue

            # Aplicar al master: el proveedor de mayor prioridad sobreescribe
            mapped = master["ean"].map(mapping)          # NaN donde no tiene datos
            # Sobreescribir solo donde este proveedor SÍ tiene dato
            has_value = mapped.notna()
            master.loc[has_value, field] = mapped[has_value]

    # 4. Agregar columna de precio por proveedor
    for name, df in supplier_dfs.items():
        price_map = df.set_index("ean")["precio"].to_dict()
        master[f"precio_{name}"] = master["ean"].map(price_map)
        n_with_price = master[f"precio_{name}"].notna().sum()
        logger.info(f"  {name}: {n_with_price:,} SKUs con precio en el consolidado")

    logger.info(f"Consolidado inicial: {len(master):,} rows × {len(master.columns)} columnas")
    return master
