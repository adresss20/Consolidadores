# Unificador de Catálogos — Farmacia

Consolida automáticamente los catálogos de proveedores en un solo Excel formateado.

---

## Instalación

```bash
pip install -r requirements.txt
```

---

## Cómo correr

**Opción A — archivo automático** (usa el `.xlsx` más reciente en `catalogos/`):

```bash
python main.py
```

**Opción B — especificar archivo**:

```bash
python main.py ruta/al/catalogo.xlsx
```

El resultado queda en `output/consolidado_YYYYMMDD_HHMMSS.xlsx`.  
Los logs de ejecución quedan en `logs/`.

---

## Estructura de carpetas

```
unificador_catalogos/
├── main.py              ← ejecutar esto
├── config.yaml          ← mappeo de columnas por proveedor
├── requirements.txt
├── catalogos/           ← poner aquí el Excel con las hojas de proveedores
├── output/              ← aquí aparece el consolidado generado
├── logs/                ← logs de cada ejecución
└── src/
    ├── config_loader.py
    ├── loader.py
    ├── normalizer.py
    ├── merger.py
    ├── analyzer.py
    └── exporter.py
```

---

## Formato del Excel de entrada

El libro debe tener **una hoja por proveedor**, con los nombres exactos definidos en `config.yaml` (ej. `Adsa`, `FarmaMayoreo`, `Levic`, etc.).

Los nombres de columnas de cada hoja deben coincidir con los que están en `config.yaml`. Si un proveedor cambia sus columnas, solo actualizar `config.yaml`.

---

## Agregar un proveedor nuevo

Añadir un bloque en `config.yaml`:

```yaml
suppliers:
  NuevoProveedor:
    sheet: NombreDeLaHoja
    sku_col: COLUMNA_EAN_O_CODIGO
    price_col: COLUMNA_PRECIO_FINAL
    desc_col: COLUMNA_DESCRIPCION      # null si no existe
    lab_col: COLUMNA_LABORATORIO       # null si no existe
    familia_col: COLUMNA_FAMILIA       # null si no existe
    generico_col: COLUMNA_GENERICO     # null si no existe
    sustancia_col: COLUMNA_SUSTANCIA   # null si no existe
    presentacion_col: null             # null si no existe
```

---

## Columnas del consolidado generado

| Columna | Descripción |
|---|---|
| SKU / EAN | Código de barras (llave de match entre proveedores) |
| DESCRIPCIÓN | Nombre del producto |
| LABORATORIO | Fabricante / laboratorio |
| FAMILIA | Categoría del producto |
| GENÉRICO | Sí / No |
| SUSTANCIA ACTIVA | Principio activo |
| PRESENTACIÓN | Formato (tabletas, ml, sobres, etc.) |
| ADSA / FARMA MAYOREO / ... | Precio de cada proveedor |
| PROV. MÁS BARATO | Nombre del proveedor con menor precio |
| PRECIO MÁS BARATO | Precio más bajo disponible |
| 2DO MÁS BARATO | Segundo proveedor más económico |
| PRECIO 2DO | Su precio |
| AHORRO vs 2DO | Diferencia de precio entre 1ro y 2do |
| # PROVEEDORES | Cuántos proveedores ofrecen ese SKU |

> Celdas en **verde** = precio más barato del SKU.  
> Celdas en **amarillo** = segundo precio más barato.
