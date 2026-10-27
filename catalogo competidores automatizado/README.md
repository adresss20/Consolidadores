# Catálogo consolidado de competidores

Automatización que combina los catálogos de scraping de 6 farmacias
(Benavides, Similares, Vitau, Fahorro, Guadalajara, San Pablo) en un solo
Excel, cruzando productos por código de barras (EAN) y señalando cuál
farmacia tiene el precio más bajo en cada producto.

## Estructura de carpetas

```
entradas/
  benavides/     <- coloca aquí el .xlsx más reciente de Benavides
  similares/     <- coloca aquí el .xlsx más reciente de Similares
  vitau/         <- coloca aquí el .xlsx más reciente de Vitau
  fahorro/       <- coloca aquí el .xlsx más reciente de Fahorro
  guadalajara/   <- coloca aquí el .xlsx más reciente de Guadalajara
  san_pablo/     <- coloca aquí el .xlsx más reciente de San Pablo
salida/          <- aquí aparece cada catálogo consolidado generado
backend/         <- todo el código y la configuración del proyecto
logs/            <- un archivo .log por cada corrida
```

## Cómo actualizar el catálogo consolidado

1. Cada vez que tengas un scraping nuevo de una farmacia, sustituye el
   archivo `.xlsx` dentro de su carpeta en `entradas/<farmacia>/` (borra el
   viejo o déjalo ahí, el script siempre toma el más reciente por fecha de
   modificación; si deja más de uno te avisa en el log cuál usó).
2. Desde la carpeta del proyecto, corre:

   ```
   pip install -r requirements.txt   # solo la primera vez
   python backend/main.py
   ```

3. El archivo `salida/Catalogo_Consolidado_<fecha>_<hora>.xlsx` es el
   resultado. Contiene 3 hojas:
   - **Catálogo Consolidado**: un renglón por producto (identificado por
     EAN), con el nombre y precio de cada farmacia (vacío si esa farmacia
     no lo tiene), y cuál es la más barata.
   - **SKUs sin identificar**: productos sin EAN que no se pudieron cruzar
     con confianza contra los demás catálogos (con el motivo y los
     candidatos considerados, para revisión manual).
   - **Resumen del proceso**: qué archivo se usó por farmacia, cuántos
     productos se cargaron, cuántos EAN inválidos/duplicados/precios
     inválidos hubo, y cualquier farmacia excluida por error.

No es necesario borrar el archivo consolidado anterior: cada corrida crea
uno nuevo con fecha y hora en el nombre, así queda un histórico en
`salida/`.

## Si una farmacia cambia su estructura de columnas

Todo el mapeo de columnas por farmacia vive en `backend/config.py`
(diccionario `FARMACIAS`). Si el sitio de origen cambia el nombre de una
columna o el nombre de la hoja, se actualiza ahí — el resto del código no
necesita tocarse.

## Cómo se decide qué producto es "el mismo" entre farmacias

- Si el producto tiene EAN válido en el archivo de origen, se usa
  directamente como llave de cruce.
- Si no tiene EAN, se busca en los demás catálogos un producto de la
  misma marca y el mismo gramaje/presentación (ej. "500 mg", "30
  tabletas") con un nombre muy similar. Si la coincidencia es de alta
  confianza, se asume ese EAN (queda marcado en la columna "Coincidencias
  asumidas por nombre"). Si la coincidencia es dudosa o no existe, el
  producto no se fuerza al consolidado: aparece en la hoja "SKUs sin
  identificar" para revisión manual.

## Manejo de errores

- Farmacia sin archivo en su carpeta, o con más de un `.xlsx`: se avisa en
  el log y en "Resumen del proceso"; el proceso sigue con las demás.
- Columnas esperadas que no se encuentran (p. ej. cambió el nombre de una
  columna en el sitio de origen): esa farmacia se excluye de la corrida,
  el resto continúa.
- EAN corrupto, duplicado dentro del mismo archivo, o precio inválido
  (vacío, texto, cero o negativo): se detectan, se cuentan y se registran;
  nunca se inventan datos.
- El proceso solo se detiene por completo si NINGUNA farmacia pudo
  cargarse.
