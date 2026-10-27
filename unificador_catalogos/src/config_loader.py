"""
config_loader.py
----------------
Lee y valida el archivo config.yaml.
"""

import yaml
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

REQUIRED_SUPPLIER_KEYS = {"sheet", "sku_col", "price_col"}


def load_config(config_path: str) -> dict:
    """Lee config.yaml y valida que tenga la estructura esperada."""
    path = Path(config_path)
    if not path.exists():
        raise FileNotFoundError(
            f"No se encontró el archivo de configuración: {config_path}\n"
            f"Asegúrate de que 'config.yaml' esté en la misma carpeta que main.py."
        )

    with open(path, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)

    if not isinstance(config, dict) or "suppliers" not in config:
        raise ValueError(
            "config.yaml inválido: debe contener una clave 'suppliers' con los proveedores."
        )

    # Validar cada proveedor
    for name, cfg in config["suppliers"].items():
        missing = REQUIRED_SUPPLIER_KEYS - set(cfg.keys())
        if missing:
            raise ValueError(
                f"Configuración incompleta para '{name}': faltan las claves {missing}"
            )

    suppliers = list(config["suppliers"].keys())
    logger.info(f"Configuración cargada — {len(suppliers)} proveedores: {suppliers}")
    return config
