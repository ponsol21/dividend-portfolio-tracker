#!/usr/bin/env python3
"""Compares the tracker's Transactions tab with a fresh broker positions export.

1. In the Google Sheet, open the Transactions tab → File → Download → Comma-separated values.
2. On your broker's website, export Positions to CSV.
3. Run:
     .venv/bin/python tools/reconcile.py --transactions ~/Downloads/Transactions.csv \
         --broker ~/Downloads/"Investments Positions.csv" --out adjustments.csv

Prints the share/cost differences per ticker and writes the Adjustment rows that fix them.
Paste those rows at the bottom of the Transactions tab (Source = reconcile). Old rows are never
edited, so every correction stays visible (audit trail). Reads local files only; no network.
"""
from __future__ import annotations

import argparse
import csv
from datetime import date
from decimal import Decimal
from pathlib import Path

from models import AssetType, BrokerPosition, Position, ledger, load_broker_positions, load_transactions_csv

SHARE_TOL = Decimal("0.00001")
COST_TOL = Decimal("0.01")


def differences(book: dict[str, Position], broker: list[BrokerPosition]) -> list[dict]:
    out = []
    held = {t: p for t, p in book.items() if p.shares != 0 or p.cost_basis != 0}
    for b in broker:
        p = held.pop(b.ticker, Position(ticker=b.ticker))
        ds, dc = b.quantity - p.shares, b.cost - p.cost_basis
        if abs(ds) > SHARE_TOL or abs(dc) > COST_TOL:
            out.append({"ticker": b.ticker, "sheet_shares": p.shares, "broker_shares": b.quantity, "d_shares": ds,
                        "sheet_cost": p.cost_basis, "broker_cost": b.cost, "d_cost": dc, "as_of": b.as_of})
    for t, p in held.items():  # in the sheet but not at the broker (sold, or held elsewhere)
        out.append({"ticker": t, "sheet_shares": p.shares, "broker_shares": Decimal(0), "d_shares": -p.shares,
                    "sheet_cost": p.cost_basis, "broker_cost": Decimal(0), "d_cost": -p.cost_basis, "as_of": None})
    return out


def _safe_text(s: str) -> str:
    """Neutralises spreadsheet formula injection in text cells (OWASP CSV injection)."""
    return "'" + s if s[:1] in ("=", "+", "-", "@", "\t", "\r") else s


def write_adjustments(diffs: list[dict], path: Path, today: date) -> int:
    rows = [d for d in diffs if d["as_of"] is not None]
    with path.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Date", "Ticker", "Type", "Action", "Amount (USD)", "Shares (optional)", "", "", "Source", "Key", "Note"])
        for d in rows:
            note = (f"Reconcile vs broker {d['as_of']:%m/%d/%Y}: shares {d['sheet_shares']:.5f} -> {d['broker_shares']:.5f}, "
                    f"cost {d['sheet_cost']:.2f} -> {d['broker_cost']:.2f}")
            w.writerow([today.isoformat(), d["ticker"], "Stock", "Adjustment", f"{d['d_cost']:.2f}", f"{d['d_shares']:.5f}",
                        "", "", "reconcile", "", _safe_text(note)])
    return len(rows)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--transactions", type=Path, required=True, help="CSV download of the Transactions tab")
    ap.add_argument("--broker", type=Path, required=True, help="broker positions CSV")
    ap.add_argument("--out", type=Path, default=Path("adjustments.csv"))
    args = ap.parse_args()

    txns = [t for t in load_transactions_csv(args.transactions) if t.asset_type is AssetType.STOCK]
    broker = load_broker_positions(args.broker)
    diffs = differences(ledger(txns), broker)
    if not diffs:
        print(f"OK — all {len(broker)} positions match the broker (shares ±{SHARE_TOL}, cost ±${COST_TOL}).")
        return
    print(f"{'Ticker':7}{'Sheet sh':>12}{'Broker sh':>12}{'Δ sh':>11}{'Sheet $':>10}{'Broker $':>10}{'Δ $':>8}")
    for d in diffs:
        print(f"{d['ticker']:7}{d['sheet_shares']:>12.5f}{d['broker_shares']:>12.5f}{d['d_shares']:>11.5f}"
              f"{d['sheet_cost']:>10.2f}{d['broker_cost']:>10.2f}{d['d_cost']:>8.2f}")
    n = write_adjustments(diffs, args.out, date.today())
    print(f"\nWrote {n} Adjustment rows to {args.out}. Paste them at the bottom of the Transactions tab.")
    not_at_broker = [d["ticker"] for d in diffs if d["as_of"] is None]
    if not_at_broker:
        print(f"Not at the broker (check manually, no rows written): {', '.join(not_at_broker)}")


if __name__ == "__main__":
    main()
