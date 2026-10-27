# Actualización automática del Reporte de Ventas — Fase 1

Automatiza los **pasos 5 a 11** del `Proceso manual.txt`. Tú sigues descargando
los 3 reportes del ERP (pasos 1–4); el script hace todo el trabajo de Excel y
**anexa** las ventas nuevas al consolidado existente (no crea uno nuevo).

## Requisito

```bash
pip install openpyxl
```

## Cómo usarlo

1. Entra al ERP y descarga los 3 reportes con el rango
   **[última fecha registrada + 1 … ayer]** (pasos 1–4 del proceso manual):
   - Ventas Diarias  → archivo `REPORTE_VENTAS_DIARIAS_*.xlsx`
   - Utilidad        → archivo `REPORTE_UTILIDAD_*.xlsx`
   - Consulta Ventas → archivo `ConsultaVentas*.xlsx`
2. Copia los 3 archivos a la carpeta **`input/`**.
3. Ejecuta:

   ```bash
   python main.py
   ```

4. El consolidado `Reporte consolidado de ventas.xlsx` se actualiza en sitio.
   Antes de escribir, se guarda una copia en `backups/` con fecha y hora.

> No hace falta borrar los archivos viejos de `input/`: el script toma siempre
> el **más reciente** por patrón de nombre. Y si corres el proceso dos veces,
> **no duplica** ventas (es idempotente): solo agrega fechas posteriores a la
> última registrada y hasta ayer.

## Estructura

```
R_Ventas Automatizado/
├── main.py                          # ejecuta esto
├── consolidador.py                  # motor (lógica de negocio)
├── Reporte consolidado de ventas.xlsx   # se actualiza en sitio
├── Proceso manual.txt               # referencia del proceso original
├── input/                           # aquí sueltas los 3 descargados
├── backups/                         # copia de seguridad automática por corrida
└── output/                          # (reservado)
```

## Qué automatiza y qué mejora sobre el proceso manual

Reproduce **idénticamente** el resultado del proceso manual (verificado celda por
celda), y además:

- **Calcula el rango de fechas** a partir de la última fecha del consolidado
  (evita el error humano de elegir mal el rango).
- **Detecta y descarta** la fila "Totales del período" de Ventas Diarias.
- **Normaliza el canal** (`UBER EATS  ` → `UBER EATS`) y clasifica comisiones
  por contenido, no por texto exacto.
- **Unifica tipos** de ticket y EAN para que los cruces (BUSCARX) no fallen.
- **Backup automático** antes de cada escritura.
- **Idempotente**: correr de más no duplica datos.

## Reglas de negocio implementadas

| Columna consolidado | Origen | Regla |
|---|---|---|
| A CLIENTE (canal) | ConsultaVentas | BUSCARX(ticket → CLIENTE), normalizado |
| B FECHA VENTA | Ventas Diarias `Fecha` | parte fecha |
| C HORA | Ventas Diarias `Fecha` | parte hora |
| D MES | fórmula | `=TEXT(B,"mmmm")` |
| E SEMANA | fórmula | `=ISOWEEKNUM(B)` |
| F TICKET | Ventas Diarias `Ticket` | directo |
| G CLAVE ARTICULO | Ventas Diarias `EAN` | directo |
| H ARTICULO | Ventas Diarias `Artículo` | directo |
| I CANTIDAD | Ventas Diarias `Cantidad` | directo |
| J TOTAL VENTA | Ventas Diarias `Total` | directo |
| K COSTO | Utilidad `Costos` | BUSCARX(EAN → Costos), **unitario c/impuestos** |
| L COSTO TOTAL | fórmula | `=K*I` |
| M COMISIONES | según canal | `=J*14%` si UBER/RAPPI, si no `0` |
| N IVA | según canal | `=M*16%` si UBER/RAPPI, si no `0` |
| O UTILIDAD | fórmula | `=J-L-M-N` |

## Nota sobre la Fase 2 (scraper)

La descarga automática desde el ERP (`fullfarma.mx/AppGTC`) se dejó fuera de la
Fase 1 a propósito: aporta poco tiempo (bajar 3 archivos son ~2–3 min) y es la
parte más frágil (login, sesión, formularios que pueden cambiar). El motor de
esta fase funciona de forma robusta sin depender de eso. Cuando quieras, se
puede añadir un `scraper.py` (Playwright) como módulo opcional que solo se
encargue de dejar los 3 archivos en `input/`.
