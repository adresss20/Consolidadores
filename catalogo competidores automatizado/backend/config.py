"""
Configuración central del proyecto: rutas, mapeo de columnas por farmacia
y umbrales del matching por nombre.

Si el sitio de una farmacia cambia el nombre de una columna, o si se agrega
una farmacia nueva, este es el único archivo que debería necesitar edición.
"""
from pathlib import Path

# ---------------------------------------------------------------------------
# Rutas del proyecto
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent.parent
ENTRADAS_DIR = BASE_DIR / "entradas"
SALIDA_DIR = BASE_DIR / "salida"
LOGS_DIR = BASE_DIR / "logs"

# ---------------------------------------------------------------------------
# Mapeo de columnas por farmacia
#
# hoja:          nombre de la hoja dentro del .xlsx que contiene el catálogo
# header_row:    índice (0-based) de la fila que contiene los encabezados
# ean:           columna con el código de barras / EAN
# nombre:        columna con el nombre del producto
# marca:         columna con marca/laboratorio/fabricante (para el matching por nombre)
# precio_lista:  columna con el precio regular / de lista
# precio_final:  columna con el precio de venta / con descuento (el que paga el cliente)
# sku_origen:    columna con el identificador interno de la farmacia (solo informativo)
# ---------------------------------------------------------------------------
FARMACIAS = {
    "benavides": {
        "nombre_visible": "Benavides",
        "hoja": "Catálogo",
        "header_row": 0,
        "ean": "EAN / Barcode",
        "nombre": "Nombre Completo",
        "marca": "Marca",
        "precio_lista": "Precio ($)",
        "precio_final": "Precio Promo ($)",
        "sku_origen": "SKU",
    },
    "similares": {
        "nombre_visible": "Similares",
        "hoja": "Catalogo",
        "header_row": 0,
        "ean": "EAN / Barcode",
        "nombre": "Nombre Producto",
        # La columna "Marca" de Similares es casi siempre "PICK UP" (no es una
        # marca real por producto), así que no sirve para el matching por nombre.
        "marca": None,
        "precio_lista": "Precio Lista (MXN)",
        "precio_final": "Precio Venta (MXN)",
        "sku_origen": "SKU ID",
    },
    "vitau": {
        "nombre_visible": "Vitau",
        "hoja": "Catálogo",
        "header_row": 0,
        "ean": "EAN",
        "nombre": "Nombre",
        "marca": "Marca",
        "precio_lista": "Precio Lista",
        "precio_final": "Precio Venta",
        "sku_origen": "ID",
    },
    "fahorro": {
        "nombre_visible": "Fahorro",
        "hoja": "Catálogo",
        "header_row": 0,
        "ean": "EAN/SKU",
        "nombre": "Nombre",
        "marca": None,  # Fahorro no trae columna de marca dedicada
        "precio_lista": "Precio Regular MXN",
        "precio_final": "Precio Final MXN",
        "sku_origen": "EAN/SKU",
    },
    "guadalajara": {
        "nombre_visible": "Guadalajara",
        "hoja": "Catálogo",
        "header_row": 0,
        "ean": "EAN",
        "nombre": "Nombre",
        "marca": "Marca",
        "precio_lista": "Precio Lista",
        "precio_final": "Precio Venta",
        "sku_origen": "PID",
    },
    "san_pablo": {
        "nombre_visible": "San Pablo",
        "hoja": "Catálogo",
        "header_row": 2,  # las 2 primeras filas son título/espacio, el encabezado real es la fila 3
        "ean": "UPC_EAN",
        "nombre": "Nombre",
        "marca": "Marca",
        "precio_lista": "Precio_Lista",
        "precio_final": "Precio_Venta",
        "sku_origen": "Item_Code",
        "hoja_errores": "Errores",  # hoja opcional con log de errores del propio scraping
    },
}

# ---------------------------------------------------------------------------
# Validación de EAN: longitudes numéricas plausibles (UPC-A=12, EAN-13, GTIN-14, EAN-8)
# ---------------------------------------------------------------------------
LONGITUDES_EAN_VALIDAS = {8, 12, 13, 14}

# ---------------------------------------------------------------------------
# Umbrales del matching por nombre (para filas sin EAN válido)
# ---------------------------------------------------------------------------
UMBRAL_MATCH_ALTO = 0.90   # >= esto y coincide marca+gramaje -> se asume el match
UMBRAL_MATCH_AMBIGUO = 0.75  # entre ambiguo y alto -> se reporta como ambiguo en "SKUs sin identificar"

# Duplicados de EAN dentro del mismo archivo de origen
CRITERIO_DUPLICADO = "precio_final_minimo"  # se conserva la fila con menor precio_final
