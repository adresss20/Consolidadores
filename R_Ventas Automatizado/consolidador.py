# -*- coding: utf-8 -*-
"""
consolidador.py
---------------
Motor de actualizacion del "Reporte consolidado de ventas".

Reemplaza los pasos manuales 5 a 11 del archivo "Proceso manual.txt":
toma los 3 reportes descargados del ERP (Ventas Diarias, Utilidad y
Consulta de Ventas) y ANEXA las ventas nuevas al final del consolidado
existente. No crea un archivo nuevo: actualiza el que ya se mantiene.

Reglas de negocio (confirmadas contra los datos reales):
  A CLIENTE (canal)  <- BUSCARX(ticket -> ConsultaVentas.CLIENTE)   (normalizado)
  B FECHA VENTA      <- parte fecha de Ventas Diarias "Fecha"
  C HORA             <- parte hora  de Ventas Diarias "Fecha"
  D MES              =  TEXT(B,"mmmm")                (formula)
  E SEMANA           =  ISOWEEKNUM(B)                 (formula)
  F TICKET           <- Ventas Diarias "Ticket"
  G CLAVE ARTICULO   <- Ventas Diarias "EAN"
  H ARTICULO         <- Ventas Diarias "Articulo"
  I CANTIDAD         <- Ventas Diarias "Cantidad"
  J TOTAL VENTA      <- Ventas Diarias "Total"
  K COSTO            <- BUSCARX(EAN -> Utilidad.Costos)   (unitario c/impuestos)
  L COSTO TOTAL      =  K*I                            (formula)
  M COMISIONES       =  J*14%  si canal UBER/RAPPI, si no 0
  N IVA              =  M*16%  si canal UBER/RAPPI, si no 0
  O UTILIDAD         =  J-L-M-N                        (formula)
"""

from __future__ import annotations

import shutil
from dataclasses import dataclass, field
from datetime import datetime, date
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter


# --------------------------------------------------------------------------- #
#  Normalizadores de claves de cruce
# --------------------------------------------------------------------------- #
def norm_ticket(v) -> str | None:
    """El ticket viene como texto en Ventas Diarias ('1746') y como numero
    (1746.0) en ConsultaVentas. Se unifica a un string entero."""
    if v is None or v == "":
        return None
    try:
        return str(int(float(v)))
    except (ValueError, TypeError):
        return str(v).strip()


def norm_ean(v) -> str | None:
    """EAN/clave de articulo: se unifica a string sin '.0' colgante."""
    if v is None or v == "":
        return None
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def parse_fecha(raw) -> tuple[datetime, str] | None:
    """'01/08/2026, 17:58:17' -> (datetime(2026,8,1), '17:58:17').
    Acepta separador coma o espacio. Formato de fecha dd/mm/yyyy."""
    if raw is None:
        return None
    if isinstance(raw, datetime):
        return datetime(raw.year, raw.month, raw.day), raw.strftime("%H:%M:%S")
    s = str(raw).strip()
    if "," in s:
        d_part, t_part = s.split(",", 1)
    elif " " in s:
        d_part, t_part = s.split(" ", 1)
    else:
        d_part, t_part = s, ""
    d_part, t_part = d_part.strip(), t_part.strip()
    dt = datetime.strptime(d_part, "%d/%m/%Y")
    return dt, t_part


def es_canal_comision(cliente: str | None) -> bool:
    """UBER EATS y RAPPI cobran comision (14%) + IVA sobre comision (16%).
    PUBLICO GENERAL no. Se compara sin espacios sobrantes."""
    c = (cliente or "").upper()
    return "UBER" in c or "RAPPI" in c


def norm_cliente(cliente: str | None) -> str:
    """Quita espacios sobrantes del canal ('UBER EATS  ' -> 'UBER EATS')."""
    return (cliente or "").strip()


# --------------------------------------------------------------------------- #
#  Resultado de la corrida (para reportar al usuario)
# --------------------------------------------------------------------------- #
@dataclass
class Resultado:
    ultima_fecha_previa: date | None = None
    fecha_corte: date | None = None            # ayer
    filas_leidas: int = 0                       # lineas en Ventas Diarias (sin total)
    filas_nuevas: int = 0                       # lineas efectivamente anexadas
    primera_fila_nueva: int | None = None
    ultima_fila_nueva: int | None = None
    backup: Path | None = None
    avisos: list[str] = field(default_factory=list)

    def resumen(self) -> str:
        L = []
        L.append("=" * 60)
        L.append(" RESULTADO DE LA ACTUALIZACION")
        L.append("=" * 60)
        L.append(f"  Ultima fecha previa en consolidado : {self.ultima_fecha_previa}")
        L.append(f"  Fecha de corte (ayer)              : {self.fecha_corte}")
        L.append(f"  Lineas leidas de Ventas Diarias    : {self.filas_leidas}")
        L.append(f"  Lineas nuevas anexadas             : {self.filas_nuevas}")
        if self.filas_nuevas:
            L.append(f"  Rango de filas escritas            : {self.primera_fila_nueva}..{self.ultima_fila_nueva}")
        if self.backup:
            L.append(f"  Backup creado                      : {self.backup.name}")
        if self.avisos:
            L.append("  AVISOS:")
            for a in self.avisos:
                L.append(f"    - {a}")
        L.append("=" * 60)
        return "\n".join(L)


# --------------------------------------------------------------------------- #
#  Lectura de reportes de origen
# --------------------------------------------------------------------------- #
def _ws(path: Path):
    wb = openpyxl.load_workbook(path, data_only=True)
    return wb.worksheets[0]


def cargar_consulta_ventas(path: Path) -> dict[str, str]:
    """ticket_normalizado -> CLIENTE (canal). Columnas: A TICKET, D CLIENTE."""
    ws = _ws(path)
    mapa: dict[str, str] = {}
    for r in range(2, ws.max_row + 1):
        tk = norm_ticket(ws.cell(r, 1).value)
        cli = ws.cell(r, 4).value
        if tk and cli is not None:
            mapa.setdefault(tk, norm_cliente(cli))
    return mapa


def cargar_utilidad(path: Path) -> dict[str, float]:
    """ean_normalizado -> Costos (unitario c/imp). Columnas: A EAN, F Costos.
    Igual que BUSCARX: se conserva la PRIMERA coincidencia por EAN."""
    ws = _ws(path)
    mapa: dict[str, float] = {}
    for r in range(2, ws.max_row + 1):
        ean = norm_ean(ws.cell(r, 1).value)
        costo = ws.cell(r, 6).value
        if ean and isinstance(costo, (int, float)):
            mapa.setdefault(ean, float(costo))
    return mapa


def _es_fila_total(ticket, articulo) -> bool:
    if ticket in (None, ""):
        return True
    if isinstance(articulo, str) and "totales del" in articulo.lower():
        return True
    return False


def leer_ventas_diarias(path: Path) -> list[dict]:
    """Lee Ventas Diarias como tabla base. Una fila por linea de ticket.
    Columnas: A Ticket, B Fecha, C EAN, D Articulo, E Cantidad, F Total.
    Descarta la fila final de 'Totales del periodo'."""
    ws = _ws(path)
    filas: list[dict] = []
    for r in range(2, ws.max_row + 1):
        ticket = ws.cell(r, 1).value
        articulo = ws.cell(r, 4).value
        if _es_fila_total(ticket, articulo):
            continue
        parsed = parse_fecha(ws.cell(r, 2).value)
        if parsed is None:
            continue
        fecha_dt, hora = parsed
        filas.append(
            {
                "ticket": norm_ticket(ticket),
                "fecha": fecha_dt,
                "hora": hora,
                "ean": norm_ean(ws.cell(r, 3).value),
                "articulo": articulo,
                "cantidad": ws.cell(r, 5).value,
                "total": ws.cell(r, 6).value,
            }
        )
    return filas


# --------------------------------------------------------------------------- #
#  Escritura en el consolidado
# --------------------------------------------------------------------------- #
def _ultima_fila_y_fecha(ws) -> tuple[int, date | None]:
    """Ultima fila con dato en col B (FECHA VENTA) y su fecha maxima."""
    last = 1
    max_fecha: date | None = None
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, 2).value
        if v in (None, ""):
            continue
        last = r
        d = v.date() if isinstance(v, datetime) else None
        if d and (max_fecha is None or d > max_fecha):
            max_fecha = d
    return last, max_fecha


def actualizar_consolidado(
    consolidado_path: Path,
    ventas: list[dict],
    consulta: dict[str, str],
    utilidad: dict[str, float],
    fecha_corte: date | None = None,
    backups_dir: Path | None = None,
) -> Resultado:
    """Anexa las ventas nuevas al consolidado. Idempotente: solo agrega filas
    con fecha estrictamente posterior a la ultima registrada y no mayor a la
    fecha de corte (ayer)."""
    res = Resultado()
    res.filas_leidas = len(ventas)
    if fecha_corte is None:
        fecha_corte = date.today()  # se filtra con '<' hoy => hasta ayer
    res.fecha_corte = fecha_corte - _one_day() if fecha_corte == date.today() else fecha_corte

    # Backup antes de tocar el archivo
    if backups_dir is not None:
        backups_dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        bkp = backups_dir / f"{consolidado_path.stem}_{stamp}{consolidado_path.suffix}"
        shutil.copy2(consolidado_path, bkp)
        res.backup = bkp

    wb = openpyxl.load_workbook(consolidado_path, data_only=False)
    ws = wb.worksheets[0]

    last_row, last_fecha = _ultima_fila_y_fecha(ws)
    res.ultima_fecha_previa = last_fecha

    # Formatos de numero de la ultima fila existente (para replicar apariencia)
    fmt = {c: ws.cell(last_row, c).number_format for c in range(1, 16)}

    # Filtro de idempotencia: fecha > ultima registrada y fecha <= ayer
    corte = fecha_corte  # exclusivo por debajo, ver comparacion
    nuevas = []
    for v in ventas:
        d = v["fecha"].date()
        if last_fecha is not None and d <= last_fecha:
            continue
        if d >= corte:  # corte == hoy -> excluye hoy (toma hasta ayer)
            continue
        nuevas.append(v)
    nuevas.sort(key=lambda x: (x["fecha"], x["ticket"] or ""))

    r = last_row
    for v in nuevas:
        r += 1
        canal = consulta.get(v["ticket"], "")
        if not canal:
            res.avisos.append(f"Ticket {v['ticket']} sin canal en ConsultaVentas (fila {r}).")
        costo = utilidad.get(v["ean"])
        if costo is None:
            res.avisos.append(f"EAN {v['ean']} sin costo en Utilidad (ticket {v['ticket']}, fila {r}).")

        comision = es_canal_comision(canal)
        vals = {
            1: canal,                                  # A CLIENTE
            2: v["fecha"],                             # B FECHA VENTA (datetime)
            3: v["hora"],                              # C HORA
            4: f'=TEXT(B{r},"mmmm")',                  # D MES
            5: f"=_xlfn.ISOWEEKNUM(B{r})",             # E SEMANA
            6: _num_or_text(v["ticket"]),              # F TICKET
            7: _num_or_text(v["ean"]),                 # G CLAVE ARTICULO
            8: v["articulo"],                          # H ARTICULO
            9: v["cantidad"],                          # I CANTIDAD
            10: v["total"],                            # J TOTAL VENTA
            11: costo,                                 # K COSTO
            12: f"=K{r}*I{r}",                         # L COSTO TOTAL
            13: (f"=+J{r}*14%" if comision else 0),    # M COMISIONES
            14: (f"=+M{r}*16%" if comision else 0),    # N IVA
            15: f"=+J{r}-L{r}-M{r}-N{r}",              # O UTILIDAD
        }
        for c, val in vals.items():
            cell = ws.cell(r, c, val)
            if fmt.get(c):
                cell.number_format = fmt[c]

    if nuevas:
        res.filas_nuevas = len(nuevas)
        res.primera_fila_nueva = last_row + 1
        res.ultima_fila_nueva = r
        wb.save(consolidado_path)
    else:
        # Nada que anexar: se elimina el backup redundante para no acumular basura
        if res.backup and res.backup.exists():
            res.backup.unlink()
            res.backup = None
    wb.close()
    return res


def _num_or_text(s: str | None):
    """Ticket y EAN se guardan como numero cuando es posible (como en el
    consolidado existente), si no, como texto."""
    if s is None:
        return None
    try:
        return int(s)
    except (ValueError, TypeError):
        return s


def _one_day():
    from datetime import timedelta
    return timedelta(days=1)
