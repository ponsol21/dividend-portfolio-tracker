"""Checks for the tracker. Run from the project root:  .venv/bin/python -m unittest discover -s tools -v

These checks need no personal data. Checks against your own broker export, workbook and holdings live in
tools/test_private.py, which stays on your computer (see .gitignore) and runs with the same command.
"""
from __future__ import annotations

import re
import unittest
from datetime import date
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook
from pydantic import ValidationError

from models import BrokerPosition, Transaction
from reconcile import _safe_text

ROOT = Path(__file__).resolve().parent.parent
BROKER = ROOT / "inputs" / "Investments Positions.csv"
WORKBOOK = ROOT / "Portfolio_Tracker_v2.xlsx"
FIXTURES = ROOT / "apps_script" / "tests" / "fixtures"
PREFILL = ROOT / "build" / "prefill.json"
PERSONAL = ROOT / "inputs" / "personal.json"


def seeded_transactions() -> list[Transaction]:
    ws = load_workbook(WORKBOOK)["Transactions"]
    out = []
    for r in range(2, ws.max_row + 1):
        d, t, typ, act, amt, sh = (ws.cell(r, c).value for c in range(1, 7))
        if not d:
            continue
        out.append(Transaction(date=d.date(), ticker=t, asset_type=typ, action=act, amount=str(amt),
                               shares=None if sh is None else str(sh), source=ws.cell(r, 9).value or "manual",
                               note=ws.cell(r, 11).value or ""))
    return out


class BrokerData(unittest.TestCase):
    def test_rejects_inconsistent_row(self):
        with self.assertRaises(ValidationError):
            BrokerPosition(ticker="XYZ", quantity="1", price="10", value="99", cost="5", as_of="10/19/2026")


class TransactionValidation(unittest.TestCase):
    base = dict(date=date(2026, 10, 19), asset_type="Stock", action="Buy", amount="100", shares="1")

    def test_rejects_formula_ticker(self):
        for bad in ("=HYPERLINK(1)", "XYZ;DROP", "xyz xyz", ""):
            with self.assertRaises(ValidationError):
                Transaction(ticker=bad, **self.base)

    def test_rejects_negative_buy(self):
        with self.assertRaises(ValidationError):
            Transaction(ticker="XYZ", **{**self.base, "amount": "-5"})

    def test_adjustment_may_be_negative(self):
        t = Transaction(ticker="XYZ", **{**self.base, "action": "Adjustment", "amount": "-1.5", "shares": "-0.01"})
        self.assertEqual(t.shares_delta, Decimal("-0.01"))

    def test_csv_injection_neutralised(self):
        self.assertEqual(_safe_text("=cmd|' /C calc'!A0"), "'=cmd|' /C calc'!A0")
        self.assertEqual(_safe_text("Reconcile vs broker"), "Reconcile vs broker")


def evaluate_engine(months: int, values: dict[str, float]) -> dict[str, float]:
    """Evaluates the ProjEngine formulas exactly as written in the workbook (simple arithmetic only)."""
    ws = load_workbook(WORKBOOK)["ProjEngine"]
    cells: dict[str, float] = {}
    ref = re.compile(r"\b([A-Z]{1,2})(\d+)\b")
    for r in range(2, months + 3):
        for c in range(1, 20):
            cell = ws.cell(r, c)
            v = cell.value
            if isinstance(v, str) and v.startswith("="):
                expr = v[1:].replace("^", "**")
                for nm, val in values.items():
                    expr = re.sub(rf"\b{nm}\b", repr(val), expr)
                expr = ref.sub(lambda m: f"cells['{m.group(1)}{m.group(2)}']", expr)
                v = eval(expr, {"cells": cells, "IF": lambda c, a, b: a if c else b})  # noqa: S307 — expressions come from our own generated workbook
            cells[cell.coordinate] = v
    return cells


def python_model(v0, y0, gp, gd, contrib, months, stage_years=40):
    p, d, s, cum = 1.0, y0, v0, 0.0
    for m in range(1, months + 1):
        div = s * d / 12
        p *= (1 + gp) ** (1 / 12)
        d_new = d * (1 + (gd if m <= 12 * stage_years else gp)) ** (1 / 12)
        s += (div + contrib) / p
        cum += div
        d = d_new
    return s * p, s * d, cum


class Projections(unittest.TestCase):
    inputs = dict(StartYield=0.025, TotalValue=5000.0, MonthlyContribution=250.0, GpCons=0.04, GdCons=0.03,
                  GpBase=0.06, GdBase=0.05, GpOpt=0.08, GdOpt=0.07, DivGrowthYears=5)

    def test_closed_form_no_contributions(self):
        # gp = gd = g, no contributions: value grows by ((1+g)^(1/12) + y/12) each month.
        v, _, _ = python_model(1000, 0.03, 0.06, 0.06, 0, 120)
        self.assertAlmostEqual(v, 1000 * ((1.06) ** (1 / 12) + 0.03 / 12) ** 120, delta=1e-6)

    def test_old_sheet_bug_reproduced(self):
        """Old 'Future Portfolio Outlook' J48: divides the dividend term by PriceGrowth (example numbers)."""
        pv, div, gp, gd = 1000, 30, 0.06, 0.05
        old = lambda n: pv * (1 + gp) ** n + (div / gp) * ((1 + gp) ** n - (1 + gd) ** n) / (gp - gd)
        fixed = lambda n: pv * (1 + gp) ** n + div * ((1 + gp) ** n - (1 + gd) ** n) / (gp - gd)
        self.assertAlmostEqual(old(5), 4435.426333, places=3)    # the old formula overstates ~3x at 5 years
        self.assertAlmostEqual(old(40), 172572.1792, places=2)
        self.assertAlmostEqual(fixed(5), 1524.06, delta=0.01)
        self.assertAlmostEqual(fixed(40), 20022.9, delta=0.1)


if __name__ == "__main__":
    unittest.main()
