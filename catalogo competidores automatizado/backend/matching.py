"""
Matching de productos sin EAN válido contra el resto de los catálogos.

Estrategia (en dos pasadas, para no depender del orden de las farmacias):

1) Se intenta emparejar cada producto sin EAN contra productos de OTRAS
   farmacias que SÍ tienen EAN real, exigiendo:
     - marca compatible (igual o muy similar, si ambas la traen)
     - gramaje/presentación compatible (si ambos lo extraen del nombre)
     - similitud de nombre >= UMBRAL_MATCH_ALTO
   Si hay match, se le "presta" el EAN real del candidato.

2) Lo que sigue sin EAN se intenta emparejar entre sí (dos productos sin
   EAN, de farmacias distintas, que claramente son el mismo producto). Si
   hay match, se les asigna un identificador sintético compartido
   ("ASUMIDO-N").

Lo que no logra un match de alta confianza en ninguna pasada queda
reportado en "SKUs sin identificar" con el motivo y los mejores candidatos
considerados, en vez de forzarse dentro del consolidado.

Para mantener el espacio de búsqueda manejable (catálogos de miles de
filas), se usa "bloqueo": solo se comparan productos que comparten al
menos un token de gramaje o el primer token significativo del nombre.
"""
from __future__ import annotations

import difflib
import re
from dataclasses import dataclass
from typing import Optional

from config import UMBRAL_MATCH_ALTO, UMBRAL_MATCH_AMBIGUO
from io_utils import Producto, normalizar_marca, normalizar_nombre

GRAMAJE_RE = re.compile(
    r"(\d+(?:\.\d+)?)\s*"
    r"(mg|mcg|g|kg|ml|l|ui|iu|tabletas?|tabs?|capsulas?|comprimidos?|unidades?|ud|pzas?|piezas?|sobres?|amp(?:olletas?)?|frasco)"
)

STOPWORDS_PREFIJO = {"de", "la", "el", "los", "las", "con", "para", "en"}


@dataclass
class Emparejado:
    ean_final: str
    metodo: str  # "EAN" o "Nombre (NN%)"


@dataclass
class SinIdentificar:
    producto: Producto
    motivo: str
    candidatos: str  # descripción legible de los mejores candidatos considerados


def extraer_gramaje(nombre_normalizado: str) -> frozenset:
    return frozenset(GRAMAJE_RE.findall(nombre_normalizado))


def primer_token_significativo(nombre_normalizado: str) -> str:
    for token in nombre_normalizado.split():
        if len(token) >= 3 and token not in STOPWORDS_PREFIJO:
            return token
    return nombre_normalizado[:3]


def _marca_compatible(marca_a: str, marca_b: str) -> bool:
    if not marca_a or not marca_b:
        return True  # sin info de marca en alguno de los dos lados: no se puede descartar
    if marca_a == marca_b:
        return True
    return difflib.SequenceMatcher(None, marca_a, marca_b).ratio() >= 0.85


def _gramaje_compatible(g_a: frozenset, g_b: frozenset) -> bool:
    if not g_a and not g_b:
        return True  # ninguno trae gramaje detectable: no se puede descartar por esto
    if not g_a or not g_b:
        return False  # uno lo trae y el otro no: mejor ser conservador
    return bool(g_a & g_b)


@dataclass
class _Item:
    farmacia: str
    idx: int
    producto: Producto
    nombre_norm: str
    marca_norm: str
    gramaje: frozenset


def _construir_items(por_farmacia: dict[str, list[Producto]], con_ean: bool) -> list[_Item]:
    items = []
    for farmacia, productos in por_farmacia.items():
        for idx, p in enumerate(productos):
            tiene_ean = p.ean is not None
            if tiene_ean != con_ean:
                continue
            nombre_norm = normalizar_nombre(p.nombre)
            items.append(
                _Item(
                    farmacia=farmacia,
                    idx=idx,
                    producto=p,
                    nombre_norm=nombre_norm,
                    marca_norm=normalizar_marca(p.marca),
                    gramaje=extraer_gramaje(nombre_norm),
                )
            )
    return items


def _construir_indices(items: list[_Item]) -> tuple[dict, dict]:
    por_gramaje: dict[tuple, list[_Item]] = {}
    por_prefijo: dict[str, list[_Item]] = {}
    for item in items:
        for token in item.gramaje:
            por_gramaje.setdefault(token, []).append(item)
        prefijo = primer_token_significativo(item.nombre_norm)
        por_prefijo.setdefault(prefijo, []).append(item)
    return por_gramaje, por_prefijo


def _candidatos_para(item: _Item, por_gramaje: dict, por_prefijo: dict) -> list[_Item]:
    vistos = {}
    for token in item.gramaje:
        for cand in por_gramaje.get(token, []):
            if cand.farmacia != item.farmacia:
                vistos[(cand.farmacia, cand.idx)] = cand
    prefijo = primer_token_significativo(item.nombre_norm)
    for cand in por_prefijo.get(prefijo, []):
        if cand.farmacia != item.farmacia:
            vistos[(cand.farmacia, cand.idx)] = cand
    return list(vistos.values())


def _mejor_candidato(item: _Item, candidatos: list[_Item]) -> tuple[Optional[_Item], float, list[tuple[_Item, float]]]:
    puntuados = []
    for cand in candidatos:
        if not _marca_compatible(item.marca_norm, cand.marca_norm):
            continue
        if not _gramaje_compatible(item.gramaje, cand.gramaje):
            continue
        score = difflib.SequenceMatcher(None, item.nombre_norm, cand.nombre_norm).ratio()
        puntuados.append((cand, score))
    puntuados.sort(key=lambda par: par[1], reverse=True)
    if not puntuados:
        return None, 0.0, []
    mejor, score = puntuados[0]
    return mejor, score, puntuados[:3]


def emparejar_sin_ean(por_farmacia: dict[str, list[Producto]]) -> tuple[dict, list[SinIdentificar]]:
    """
    Devuelve:
      - resultados: dict {(farmacia, idx_producto): Emparejado} solo para los
        productos sin EAN que lograron un match de alta confianza.
      - sin_identificar: lista de SinIdentificar para los que no lo lograron.
    """
    con_ean_items = _construir_items(por_farmacia, con_ean=True)
    sin_ean_items = _construir_items(por_farmacia, con_ean=False)

    resultados: dict[tuple, Emparejado] = {}
    sin_identificar: list[SinIdentificar] = []
    pendientes: list[_Item] = []

    # --- Pasada 1: contra productos con EAN real -------------------------
    por_gramaje_ean, por_prefijo_ean = _construir_indices(con_ean_items)
    for item in sin_ean_items:
        candidatos = _candidatos_para(item, por_gramaje_ean, por_prefijo_ean)
        mejor, score, top = _mejor_candidato(item, candidatos)
        if mejor is not None and score >= UMBRAL_MATCH_ALTO:
            resultados[(item.farmacia, item.idx)] = Emparejado(
                ean_final=mejor.producto.ean,
                metodo=f"Nombre ({int(round(score * 100))}%)",
            )
        else:
            pendientes.append(item)
            item._mejor_pasada1 = (mejor, score, top)  # type: ignore[attr-defined]

    # --- Pasada 2: pendientes entre sí (ambos sin EAN) --------------------
    por_gramaje_pend, por_prefijo_pend = _construir_indices(pendientes)

    resueltos_pasada2 = set()
    contador_sinteticos = 0
    for item in pendientes:
        clave = (item.farmacia, item.idx)
        if clave in resueltos_pasada2:
            continue
        candidatos = [
            c for c in _candidatos_para(item, por_gramaje_pend, por_prefijo_pend)
            if (c.farmacia, c.idx) not in resueltos_pasada2
        ]
        mejor, score, top = _mejor_candidato(item, candidatos)
        if mejor is not None and score >= UMBRAL_MATCH_ALTO:
            contador_sinteticos += 1
            ean_sintetico = f"ASUMIDO-{contador_sinteticos}"
            pct = int(round(score * 100))
            resultados[clave] = Emparejado(ean_final=ean_sintetico, metodo=f"Nombre ({pct}%)")
            resultados[(mejor.farmacia, mejor.idx)] = Emparejado(ean_final=ean_sintetico, metodo=f"Nombre ({pct}%)")
            resueltos_pasada2.add(clave)
            resueltos_pasada2.add((mejor.farmacia, mejor.idx))

    # --- Lo que sigue sin resolver: a "SKUs sin identificar" --------------
    for item in pendientes:
        clave = (item.farmacia, item.idx)
        if clave in resueltos_pasada2:
            continue
        mejor1, score1, top1 = getattr(item, "_mejor_pasada1", (None, 0.0, []))
        if score1 >= UMBRAL_MATCH_AMBIGUO:
            motivo = "Coincidencia ambigua (no alcanza el umbral de confianza)"
            descripciones = [
                f"{c.farmacia}: '{c.producto.nombre}' ({int(round(s * 100))}%)" for c, s in top1
            ]
        else:
            motivo = "Sin coincidencia en los demás catálogos"
            descripciones = []
        sin_identificar.append(
            SinIdentificar(
                producto=item.producto,
                motivo=motivo,
                candidatos="; ".join(descripciones),
            )
        )

    return resultados, sin_identificar
