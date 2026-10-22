"""Validated data models for the portfolio tracker (pydantic v2).

Everything that enters the tracker from a file goes through these models, so bad rows
fail loudly instead of silently becoming zeros in the spreadsheet.
"""
from __future__ import annotations

import csv
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from enum import Enum
from pathlib import Path
from typing import Iterable

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator

TICKER_PATTERN = r"^[A-Z][A-Z0-9.\-]{0,9}$"
MAX_INPUT_BYTES = 5_000_000  # refuse unexpectedly large input files


class AssetType(str, Enum):
    STOCK = "Stock"
    CRYPTO = "Crypto"


class Action(str, Enum):
    BUY = "Buy"
    SELL = "Sell"
    DIVIDEND_REINVEST = "Dividend Reinvest"
    DIVIDEND_CASH = "Dividend Cash"
    ADJUSTMENT = "Adjustment"


def _decimal(value: object) -> Decimal | None:
    if value is None:
        return None
    text = str(value).strip().replace(",", "").replace("$", "")
    if text in ("", "None", "--"):
        return None
    try:
        return Decimal(text)
    except InvalidOperation as exc:
        raise ValueError(f"not a number: {value!r}") from exc


class Transaction(BaseModel):
    """One row of the Transactions tab."""

    model_config = ConfigDict(extra="forbid", frozen=True, str_strip_whitespace=True)

    date: date
    ticker: str = Field(pattern=TICKER_PATTERN)
    asset_type: AssetType
    action: Action
    amount: Decimal
    shares: Decimal | None = None
    source: str = Field(default="manual", max_length=40)
    key: str = Field(default="", max_length=40)
    note: str = Field(default="", max_length=300)

    @field_validator("ticker", mode="before")
    @classmethod
    def _upper(cls, v: object) -> object:
        return v.strip().upper() if isinstance(v, str) else v

    @field_validator("amount", "shares", mode="before")
    @classmethod
    def _num(cls, v: object) -> object:
        return _decimal(v)

    @model_validator(mode="after")
    def _check(self) -> "Transaction":
        if self.action is not Action.ADJUSTMENT:
            if self.amount < 0:
                raise ValueError("amount must be >= 0 (only Adjustment rows may be negative)")
            if self.shares is not None and self.shares < 0:
                raise ValueError("shares must be >= 0 (only Adjustment rows may be negative)")
        if self.action in (Action.BUY, Action.SELL) and self.amount == 0:
            raise ValueError("Buy/Sell rows need an amount")
        return self

    @property
    def shares_delta(self) -> Decimal:
        s = self.shares or Decimal(0)
        if self.action is Action.SELL:
            return -s
        if self.action is Action.DIVIDEND_CASH:
            return Decimal(0)
        return s


class Position(BaseModel):
    ticker: str
    shares: Decimal = Decimal(0)
    cost_basis: Decimal = Decimal(0)
    money_in: Decimal = Decimal(0)
    dividends: Decimal = Decimal(0)
    realized: Decimal = Decimal(0)


def ledger(transactions: Iterable[Transaction]) -> dict[str, Position]:
    """Average-cost ledger, same rules as the spreadsheet formulas (rows in date order)."""
    book: dict[str, Position] = {}
    for t in sorted(transactions, key=lambda x: x.date):
        p = book.setdefault(t.ticker, Position(ticker=t.ticker))
        if t.action is Action.SELL:
            sold = t.shares or Decimal(0)
            avg = p.cost_basis / p.shares if p.shares else Decimal(0)
            removed = avg * sold
            p.cost_basis -= removed
            p.realized += t.amount - removed
            p.money_in -= t.amount
        elif t.action is Action.DIVIDEND_CASH:
            p.dividends += t.amount
        else:
            p.cost_basis += t.amount
            if t.action is Action.BUY:
                p.money_in += t.amount
            if t.action is Action.DIVIDEND_REINVEST:
                p.dividends += t.amount
        p.shares += t.shares_delta
    return book


class BrokerPosition(BaseModel):
    """One holding from the broker's 'Investments Positions' CSV export."""

    model_config = ConfigDict(extra="ignore", frozen=True)

    ticker: str = Field(pattern=TICKER_PATTERN)
    description: str = ""
    quantity: Decimal = Field(gt=0)
    price: Decimal = Field(gt=0)
    value: Decimal = Field(ge=0)
    cost: Decimal = Field(ge=0)
    est_annual_income: Decimal | None = None
    income_earned: Decimal | None = None
    isin: str = ""
    as_of: date

    @field_validator("quantity", "price", "value", "cost", "est_annual_income", "income_earned", mode="before")
    @classmethod
    def _num(cls, v: object) -> object:
        return _decimal(v)

    @field_validator("as_of", mode="before")
    @classmethod
    def _date(cls, v: object) -> object:
        if isinstance(v, str):
            return datetime.strptime(v.strip(), "%m/%d/%Y").date()
        return v

    @model_validator(mode="after")
    def _consistent(self) -> "BrokerPosition":
        # Processing-integrity check: the broker's value must equal quantity x price.
        if abs(self.quantity * self.price - self.value) > Decimal("0.05"):
            raise ValueError(f"{self.ticker}: value {self.value} != quantity x price {self.quantity * self.price:.2f}")
        return self


BROKER_COLUMNS = {
    "ticker": "Ticker", "description": "Description", "quantity": "Quantity", "price": "Price", "value": "Value",
    "cost": "Cost", "est_annual_income": "Est. Annual Income", "income_earned": "Accrued/Income Earned",
    "isin": "ISIN", "as_of": "As of",
}


def _read_text(path: Path) -> str:
    if path.stat().st_size > MAX_INPUT_BYTES:
        raise ValueError(f"{path} is larger than {MAX_INPUT_BYTES} bytes")
    return path.read_text(encoding="utf-8-sig")


def load_broker_positions(path: Path) -> list[BrokerPosition]:
    lines = _read_text(path).splitlines()
    rows = []
    reader = csv.DictReader(lines)
    missing = [c for c in BROKER_COLUMNS.values() if c not in (reader.fieldnames or [])]
    if missing:
        raise ValueError(f"{path.name}: missing columns {missing}")
    for raw in reader:
        if not raw.get("Ticker") or not raw.get("Quantity"):
            break  # blank line before the FOOTNOTES section
        rows.append(BrokerPosition(**{k: raw[v] for k, v in BROKER_COLUMNS.items()}))
    tickers = [r.ticker for r in rows]
    if len(tickers) != len(set(tickers)):
        raise ValueError(f"{path.name}: duplicate tickers")
    return rows


TX_COLUMNS = {
    "date": "Date", "ticker": "Ticker", "asset_type": "Type", "action": "Action", "amount": "Amount (USD)",
    "shares": "Shares (used)", "source": "Source", "key": "Key", "note": "Note",
}


def load_transactions_csv(path: Path) -> list[Transaction]:
    """Reads a CSV download of the Transactions tab (File → Download → CSV while on that tab)."""
    reader = csv.DictReader(_read_text(path).splitlines())
    missing = [c for c in TX_COLUMNS.values() if c not in (reader.fieldnames or [])]
    if missing:
        raise ValueError(f"{path.name}: missing columns {missing}")
    out = []
    for n, raw in enumerate(reader, start=2):
        if not (raw.get("Date") or "").strip():
            continue
        values = {k: raw[v] for k, v in TX_COLUMNS.items()}
        values["date"] = _parse_sheet_date(values["date"])
        if values["shares"] in ("price?", ""):
            raise ValueError(f"row {n}: shares could not be determined ({values['ticker']})")
        try:
            out.append(Transaction(**values))
        except ValueError as exc:
            raise ValueError(f"row {n}: {exc}") from exc
    return out


def _parse_sheet_date(text: str) -> date:
    text = text.strip()
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%b %d, %Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"unrecognised date {text!r}")
