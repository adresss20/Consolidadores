"""
normalizer.py
-------------
Funciones de limpieza de datos:
  - normalize_ean   → convierte EAN a string limpio
  - normalize_price → convierte precio a float, descarta inválidos
  - normalize_text  → limpia campos de texto
"""

import pandas as pd
import logging

logger = logging.getLogger(__name__)


def normalize_ean(value) -> str:
    """
    Convierte un valor a string EAN limpio.
    - Elimina espacios
    - Elimina sufijo '.0' de int-floats (e.g. '7501234567890.0' → '7501234567890')
    - Retorna pd.NA si está vacío o es un marcador de celda vacía (' ')
    """
    if pd.isna(value):
        return pd.NA
    s = str(value).strip()
    if s in ("", " ", "nan", "None"):
        return pd.NA
    # Quitar sufijo .0 producido al leer enteros como float
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s


def normalize_price(value) -> float:
    """
    Convierte un valor a float.
    - Elimina comas de miles (e.g. '1,234.56' → 1234.56)
    - Retorna pd.NA si el valor es vacío, nulo o cero/negativo
    """
    if pd.isna(value):
        return pd.NA
    s = str(value).strip()
    if s in ("", " ", "nan", "None"):
        return pd.NA
    try:
        f = float(s.replace(",", ""))
        return f if f > 0 else pd.NA
    except (ValueError, TypeError):
        return pd.NA


def normalize_text(value) -> str:
    """Limpia un campo de texto: strip, None→pd.NA."""
    if pd.isna(value):
        return pd.NA
    s = str(value).strip()
    return s if s not in ("", "nan", "None") else pd.NA


def normalize_generico(value) -> str:
    """Normaliza el campo genérico a 'Sí' / 'No' / pd.NA."""
    if pd.isna(value):
        return pd.NA
    s = str(value).strip()
    if s in ("1", "1.0", "SI", "SÍ", "S", "TRUE", "True"):
        return "Sí"
    if s in ("0", "0.0", "NO", "N", "FALSE", "False"):
        return "No"
    return pd.NA
