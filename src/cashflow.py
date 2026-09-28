"""Reconstruye el cash flow de Nexiu a partir de los movimientos bancarios.

Replica la logica del archivo manual (Cash_Flow_report_2608.xlsx):

  movimientos crudos -> normalizacion -> diccionario de merchants
  -> ajustes manuales -> filtro de estado -> asignacion de mes
  -> agregacion por Type 2 -> consolidacion -> conciliacion contra extractos.

Hay dos vistas del mismo mes:

  caja    : cada movimiento cae en el mes de su fecha bancaria real.
            Es la que concilia contra el saldo del extracto.
  devengo : algunos pagos se reasignan al mes que les corresponde
            (sueldos de fundadores y renta se pagan al mes siguiente),
            segun config/allocations.csv. Es la vista del reporte manual.

Uso:
    python3 src/cashflow.py 2026-01            # ambas vistas
    python3 src/cashflow.py 2026-01 --base caja
"""

from __future__ import annotations

import csv
import re
import sys
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import openpyxl
import yaml

ROOT = Path(__file__).resolve().parent.parent
CONFIG = ROOT / "config"


# --------------------------------------------------------------------------
# Modelo
# --------------------------------------------------------------------------
@dataclass
class Movement:
    """Un movimiento bancario ya normalizado y clasificado."""

    account: str
    date: datetime
    description: str        # clave de busqueda en el diccionario
    raw_description: str    # texto original del banco, para auditoria
    amount: float           # siempre en USD; negativo = salida
    status: str             # Sent | Failed
    card_last_four: str = ""
    statement: str = ""     # P&L | BS | Intercompany
    type_2: str = ""
    type_2_b: str = ""
    type_3: str = ""
    adjustment: str = ""    # id del ajuste manual aplicado, si hubo
    matched: bool = False   # si el diccionario lo encontro
    allocation: str = ""    # motivo de la reasignacion de mes, si hubo
    period_alloc: str = ""  # mes de devengo, si difiere del mes de la fecha

    @property
    def period_caja(self) -> str:
        return f"{self.date:%Y-%m}"

    @property
    def period_devengo(self) -> str:
        return self.period_alloc or self.period_caja

    def period(self, basis: str) -> str:
        return self.period_caja if basis == "caja" else self.period_devengo

    @property
    def is_cash(self) -> bool:
        """Movio plata de verdad (afecta el saldo de la cuenta)."""
        return self.status == "Sent"

    @property
    def is_pl(self) -> bool:
        return self.is_cash and self.statement == "P&L"


@dataclass
class Report:
    period: str
    basis: str = "caja"
    movements: list[Movement] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


# --------------------------------------------------------------------------
# Configuracion
# --------------------------------------------------------------------------
def load_config() -> tuple[dict, dict, dict]:
    accounts = yaml.safe_load((CONFIG / "accounts.yaml").read_text())
    mapping = yaml.safe_load((CONFIG / "mapping.yaml").read_text())
    dictionary: dict[tuple[str, str], tuple[str, str, str]] = {}
    with (CONFIG / "dictionary.csv").open() as fh:
        for row in csv.DictReader(fh):
            key = (row["account"], normalize(row["key"]))
            dictionary[key] = (row["statement"], row["type_2"], row["type_3"])
    return accounts, mapping, dictionary


def load_overrides() -> dict[tuple[str, str, str, float], dict]:
    """Correcciones manuales fila por fila (lo que Ivan hacia editando celdas).

    Permiten reasignar el mes (period_alloc) y/o pisar la clasificacion
    (statement / type_2 / type_3). Se indexan por cuenta, fecha real, merchant
    e importe, asi una correccion nunca se aplica a un movimiento parecido de
    otro mes.
    """
    path = CONFIG / "overrides.csv"
    if not path.exists():
        return {}
    overrides = {}
    with path.open() as fh:
        for row in csv.DictReader(fh):
            key = (row["account"], row["date_real"], normalize(row["description"]),
                   round(float(row["amount"]), 2))
            overrides[key] = row
    return overrides


def normalize(text: str) -> str:
    """Normaliza una clave de busqueda.

    El archivo manual tiene el mismo merchant escrito con y sin acentos (y a
    veces con mojibake del export), asi que comparamos sin acentos, sin dobles
    espacios y en minusculas.
    """
    if text is None:
        return ""
    text = str(text).strip()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in text if not unicodedata.combining(c))
    text = re.sub(r"\s+", " ", text)
    return text.lower()


# --------------------------------------------------------------------------
# Parsers de los movimientos crudos
# --------------------------------------------------------------------------
def parse_mercury_xlsx(path: Path, account: str) -> list[Movement]:
    """Export de transacciones de Mercury (xlsx)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = ws.iter_rows(values_only=True)
    header = [str(c).strip() if c else "" for c in next(rows)]
    idx = {name: i for i, name in enumerate(header)}

    movements = []
    for row in rows:
        if row[idx["Date (UTC)"]] is None:
            continue
        movements.append(
            Movement(
                account=account,
                date=row[idx["Date (UTC)"]],
                description=str(row[idx["Description"]] or "").strip(),
                raw_description=str(row[idx.get("Bank Description", 0)] or "").strip(),
                amount=float(row[idx["Amount"]] or 0),
                status=str(row[idx["Status"]] or "").strip(),
                card_last_four=str(row[idx["Last Four Digits"]] or "").strip(),
            )
        )
    return movements


def parse_dolarapp_csv(path: Path, account: str) -> list[Movement]:
    """Export de transacciones de DolarApp (csv).

    El importe en USD es `base_amount` (ya convertido al fx del dia); la
    clave del diccionario es `name`, igual que en el archivo manual.
    """
    movements = []
    with path.open(encoding="utf-8-sig") as fh:
        for row in csv.DictReader(fh):
            if not row.get("created_on"):
                continue
            movements.append(
                Movement(
                    account=account,
                    date=datetime.strptime(row["created_on"], "%b %d, %Y, %I:%M %p"),
                    description=(row.get("name") or "").strip(),
                    raw_description=(row.get("description") or "").strip(),
                    amount=float(row["base_amount"]),
                    status=(row.get("transaction_state") or "").strip(),
                    card_last_four=(row.get("card_last_four") or "").strip(),
                )
            )
    return movements


PARSERS = {"mercury_xlsx": parse_mercury_xlsx, "dolarapp_csv": parse_dolarapp_csv}


# --------------------------------------------------------------------------
# Clasificacion
# --------------------------------------------------------------------------
def classify(mv: Movement, mapping: dict, dictionary: dict) -> None:
    mv.status = mapping["status_map"].get(mv.status, mv.status)

    hit = dictionary.get((mv.account, normalize(mv.description)))
    if hit:
        mv.statement, mv.type_2, mv.type_3 = hit
        mv.matched = True
    mv.type_2_b = mapping["type_2_b_map"].get(mv.type_2, mv.type_2)

    for adj in mapping.get("adjustments", []):
        if _matches(mv, adj["when"]):
            for attr, value in adj["set"].items():
                setattr(mv, attr, value)
            mv.adjustment = adj["id"]
            mv.matched = True
            break


def _matches(mv: Movement, cond: dict) -> bool:
    if "description_matches" in cond:
        if not re.search(cond["description_matches"], normalize(mv.description)):
            return False
    if "sign" in cond:
        if (cond["sign"] == "+") != (mv.amount > 0):
            return False
    if "description_contains" in cond:
        needle = normalize(cond["description_contains"])
        haystack = normalize(mv.description) + " " + normalize(mv.raw_description)
        if needle not in haystack:
            return False
    if "card_last_four" in cond and mv.card_last_four != str(cond["card_last_four"]):
        return False
    if "abs_amount" in cond and round(abs(mv.amount), 2) != round(float(cond["abs_amount"]), 2):
        return False
    return True


# --------------------------------------------------------------------------
# Construccion del reporte
# --------------------------------------------------------------------------
def build(period: str, basis: str = "caja") -> tuple[Report, dict, dict]:
    accounts, mapping, dictionary = load_config()
    overrides = load_overrides()
    report = Report(period=period, basis=basis)

    # Se leen todos los meses disponibles: en la vista devengo un movimiento de
    # otro mes puede caer en este (y uno de este mes puede irse a otro).
    everything: list[Movement] = []
    for acct in accounts["accounts"]:
        if period not in acct["files"]:
            report.warnings.append(f"{acct['id']}: sin archivo de movimientos para {period}")
        for src in dict.fromkeys(acct["files"].values()):
            for mv in PARSERS[acct["parser"]](ROOT / src, acct["id"]):
                classify(mv, mapping, dictionary)
                hit = overrides.get(
                    (mv.account, f"{mv.date:%Y-%m-%d}", normalize(mv.description),
                     round(mv.amount, 2))
                )
                if hit:
                    mv.period_alloc = hit["period_alloc"]
                    mv.allocation = hit["motivo"]
                    for attr in ("statement", "type_2", "type_3"):
                        if hit[attr]:
                            setattr(mv, attr, hit[attr])
                            mv.matched = True
                    mv.type_2_b = mapping["type_2_b_map"].get(mv.type_2, mv.type_2)
                everything.append(mv)

    report.movements = [mv for mv in everything if mv.period(basis) == period]

    if basis == "devengo":
        pendientes = [
            f"{a['period_alloc']} <- movimiento de {a['date_real']} ({a['description']}, "
            f"{a['amount']}) no encontrado en los archivos cargados"
            for a in _missing_allocations(overrides, everything, period)
        ]
        report.warnings += pendientes

    for mv in report.movements:
        if not mv.matched and mv.is_cash:
            report.warnings.append(
                f"SIN DICCIONARIO [{mv.account}] {mv.date:%Y-%m-%d} "
                f"{mv.description!r} {mv.amount:,.2f}"
            )
    return report, accounts, mapping


def _missing_allocations(overrides: dict, loaded: list[Movement], period: str) -> list[dict]:
    """Reasignaciones que apuntan a este mes pero cuyo movimiento no esta cargado.

    Pasa cuando el pago cae en un mes cuyo export todavia no se sumo (p.ej. el
    sueldo de enero pagado en febrero, sin los movimientos de febrero).
    """
    seen = {
        (mv.account, f"{mv.date:%Y-%m-%d}", normalize(mv.description), round(mv.amount, 2))
        for mv in loaded
    }
    return [row for key, row in overrides.items()
            if row["period_alloc"] == period and key not in seen]


def by_type_2(movements: list[Movement], mapping: dict) -> dict[str, float]:
    totals: dict[str, float] = {}
    for mv in movements:
        if mv.is_pl:
            totals[mv.type_2] = totals.get(mv.type_2, 0.0) + mv.amount
    order = mapping["type_2_order"]
    return {k: round(totals[k], 2) for k in order if k in totals}


def reconcile(report: Report, accounts: dict) -> list[dict]:
    """Compara el movimiento neto calculado contra los saldos del extracto."""
    declared = accounts.get("statement_balances", {}).get(report.period, {})
    out = []
    for acct in accounts["accounts"]:
        aid = acct["id"]
        net = round(sum(m.amount for m in report.movements if m.account == aid and m.is_cash), 2)
        bal = declared.get(aid)
        row = {
            "account": aid,
            "name": acct["name"],
            "net_calculado": net,
            "opening": bal["opening"] if bal else None,
            "closing": bal["closing"] if bal else None,
        }
        if bal:
            row["closing_calculado"] = round(bal["opening"] + net, 2)
            row["diferencia"] = round(row["closing_calculado"] - bal["closing"], 2)
        out.append(row)
    return out


# --------------------------------------------------------------------------
# Salida
# --------------------------------------------------------------------------
def render(report: Report, accounts: dict, mapping: dict) -> str:
    lines = [f"# Cash flow Nexiu — {report.period} (base: {report.basis})", ""]

    cash = [m for m in report.movements if m.is_cash]
    lines += [
        f"Movimientos: {len(report.movements)} "
        f"({len(cash)} efectivos, {len(report.movements) - len(cash)} fallidos/revertidos)",
        "",
    ]

    if report.basis == "caja":
        lines += [
            "## Conciliacion contra extractos bancarios",
            "",
            "| Cuenta | Saldo inicial | Neto calculado | Saldo final calculado "
            "| Saldo final extracto | Dif |",
            "|---|---:|---:|---:|---:|---:|",
        ]
        total_open = total_net = total_close = 0.0
        for r in reconcile(report, accounts):
            lines.append(
                f"| {r['name']} | {r['opening']:,.2f} | {r['net_calculado']:,.2f} | "
                f"{r['closing_calculado']:,.2f} | {r['closing']:,.2f} | {r['diferencia']:,.2f} |"
            )
            total_open += r["opening"]
            total_net += r["net_calculado"]
            total_close += r["closing"]
        lines.append(
            f"| **Consolidado** | **{total_open:,.2f}** | **{total_net:,.2f}** | "
            f"**{total_open + total_net:,.2f}** | **{total_close:,.2f}** | "
            f"**{total_open + total_net - total_close:,.2f}** |"
        )
    else:
        lines += [
            "_La vista devengo no concilia contra el extracto por definicion: "
            "reasigna pagos a otros meses. La conciliacion va en la vista caja._",
        ]

    lines += ["", "## Cash flow consolidado por Type 2 (solo P&L)", "",
              "| Type 2 | Mercury | DolarApp | Consolidado |", "|---|---:|---:|---:|"]
    per_acct = {
        a["id"]: by_type_2([m for m in report.movements if m.account == a["id"]], mapping)
        for a in accounts["accounts"]
    }
    consolidated = by_type_2(report.movements, mapping)
    for t2 in mapping["type_2_order"]:
        if t2 not in consolidated:
            continue
        m = per_acct["mercury_8107"].get(t2, 0.0)
        d = per_acct["dolarapp_mx"].get(t2, 0.0)
        lines.append(f"| {t2} | {m:,.2f} | {d:,.2f} | {consolidated[t2]:,.2f} |")
    lines.append(f"| **Total P&L** | | | **{sum(consolidated.values()):,.2f}** |")

    non_pl = [m for m in report.movements if m.is_cash and m.statement in mapping["non_pl_statements"]]
    if non_pl:
        lines += ["", "## Fuera del P&L (BS / Intercompany)", "",
                  "| Fecha | Cuenta | Descripcion | Statement | Importe |", "|---|---|---|---|---:|"]
        for mv in sorted(non_pl, key=lambda m: m.date):
            lines.append(
                f"| {mv.date:%Y-%m-%d} | {mv.account} | {mv.description} | "
                f"{mv.statement} | {mv.amount:,.2f} |"
            )
        lines.append(f"| | | | **Total** | **{sum(m.amount for m in non_pl):,.2f}** |")

    realloc = [m for m in report.movements if m.allocation and m.is_cash]
    if realloc and report.basis == "devengo":
        lines += ["", "## Movimientos reasignados a este mes", "",
                  "| Fecha real | Descripcion | Importe | Motivo |", "|---|---|---:|---|"]
        for mv in sorted(realloc, key=lambda m: m.date):
            lines.append(
                f"| {mv.date:%Y-%m-%d} | {mv.description} | {mv.amount:,.2f} | {mv.allocation} |"
            )

    adjusted = [m for m in report.movements if m.adjustment and m.is_cash]
    lines += ["", "## Ajustes manuales aplicados", ""]
    if adjusted:
        lines += ["| Fecha | Descripcion | Tarjeta | Importe | Regla | Type 2 / Type 3 |",
                  "|---|---|---|---:|---|---|"]
        for mv in sorted(adjusted, key=lambda m: m.date):
            lines.append(
                f"| {mv.date:%Y-%m-%d} | {mv.description} | {mv.card_last_four or '-'} | "
                f"{mv.amount:,.2f} | {mv.adjustment} | {mv.type_2} / {mv.type_3} |"
            )
    else:
        lines.append("_Ninguno se disparo este mes._")

    if report.warnings:
        lines += ["", "## Avisos", ""] + [f"- {w}" for w in report.warnings]
    return "\n".join(lines) + "\n"


def write_detail(report: Report, path: Path) -> None:
    with path.open("w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["fecha", "mes_caja", "mes_devengo", "cuenta", "descripcion",
                    "descripcion_banco", "tarjeta", "importe_usd", "estado", "statement",
                    "type_2", "type_2_b", "type_3", "ajuste", "reasignacion",
                    "en_diccionario"])
        for mv in sorted(report.movements, key=lambda m: (m.date, m.account)):
            w.writerow([f"{mv.date:%Y-%m-%d}", mv.period_caja, mv.period_devengo, mv.account,
                        mv.description, mv.raw_description, mv.card_last_four,
                        f"{mv.amount:.2f}", mv.status, mv.statement, mv.type_2, mv.type_2_b,
                        mv.type_3, mv.adjustment, mv.allocation,
                        "si" if mv.matched else "NO"])


def main() -> int:
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    period = args[0] if args else "2026-01"
    if "--base" in sys.argv:
        bases = [sys.argv[sys.argv.index("--base") + 1]]
    else:
        bases = ["caja", "devengo"]

    out_dir = ROOT / "output"
    out_dir.mkdir(exist_ok=True)
    for basis in bases:
        report, accounts, mapping = build(period, basis)
        md = render(report, accounts, mapping)
        (out_dir / f"cashflow-{period}-{basis}.md").write_text(md)
        write_detail(report, out_dir / f"detalle-{period}-{basis}.csv")
        print(md)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
