"""
Construye las filas del catálogo consolidado a partir de:
  - los productos con EAN real (agrupados directamente por EAN)
  - los productos sin EAN que el matching por nombre pudo emparejar
    (agrupados bajo el EAN real prestado, o bajo un EAN sintético
    "ASUMIDO-N" cuando el match fue entre dos productos sin EAN)

También calcula, por cada fila del consolidado, cuál es la farmacia más
barata según el precio final (el que paga el cliente).
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

from config import FARMACIAS
from io_utils import Producto
from matching import Emparejado, SinIdentificar

ORDEN_FARMACIAS = list(FARMACIAS.keys())


@dataclass
class FilaConsolidada:
    ean: str
    por_farmacia: dict = field(default_factory=dict)  # farmacia -> Producto
    metodos_asumidos: list = field(default_factory=list)  # ["Vitau: Nombre (92%)"]

    def farmacia_mas_barata(self) -> tuple[str, Optional[float]]:
        mejor_precio = None
        mejores_farmacias = []
        for farmacia, producto in self.por_farmacia.items():
            precio = producto.precio_final
            if precio is None:
                continue
            if mejor_precio is None or precio < mejor_precio:
                mejor_precio = precio
                mejores_farmacias = [farmacia]
            elif precio == mejor_precio:
                mejores_farmacias.append(farmacia)
        if mejor_precio is None:
            return "", None
        nombres = [FARMACIAS[f]["nombre_visible"] for f in mejores_farmacias]
        return " / ".join(nombres), mejor_precio


def construir_catalogo_consolidado(
    por_farmacia: dict[str, list[Producto]],
    resultados_matching: dict[tuple, Emparejado],
) -> list[FilaConsolidada]:
    grupos: dict[str, FilaConsolidada] = {}

    for farmacia, productos in por_farmacia.items():
        for idx, producto in enumerate(productos):
            if producto.ean is not None:
                ean_final = producto.ean
                metodo = "EAN"
            else:
                emparejado = resultados_matching.get((farmacia, idx))
                if emparejado is None:
                    continue  # quedó en "SKUs sin identificar", no entra al consolidado
                ean_final = emparejado.ean_final
                metodo = emparejado.metodo

            fila = grupos.get(ean_final)
            if fila is None:
                fila = FilaConsolidada(ean=ean_final)
                grupos[ean_final] = fila

            fila.por_farmacia[farmacia] = producto
            if metodo != "EAN":
                fila.metodos_asumidos.append(f"{FARMACIAS[farmacia]['nombre_visible']}: {metodo}")

    return list(grupos.values())
