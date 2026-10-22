#!/usr/bin/env python3
"""Builds Portfolio_Tracker_v2.xlsx, to be imported into Google Sheets (File → Import).

Formulas are written in Google Sheets syntax (GOOGLEFINANCE, FILTER, QUERY, ...), so the file is
meant for Google Sheets, not Excel.

Inputs
  --broker   your broker's "Investments Positions" CSV (seeds the Transactions tab)
  --prefill  build/prefill.json from apps_script/tests/export_prefill.js (dividend data, same code as the live script)
  --payout   build/payout.json (Finnhub payout ratios)
  --old      the exported old workbook (PV History is carried over)
  --personal inputs/personal.json (your own buy overrides, BTC value, monthly contribution and sectors; never on GitHub)

Usage
  .venv/bin/python tools/build_workbook.py --broker "inputs/Investments Positions.csv" \
      --prefill build/prefill.json --payout build/payout.json \
      --old "inputs/Final Portfolio Tracking.xlsx" --personal inputs/personal.json --out Portfolio_Tracker_v2.xlsx
"""
from __future__ import annotations

import argparse
import json
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils.datetime import from_excel
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

from models import TICKER_PATTERN, load_broker_positions

# ── Decisions confirmed with the user on 2026-10-19 ─────────────────────────────
AS_OF = date(2026, 10, 19)
# The user's own amounts (buy overrides, BTC value, monthly contribution) are in inputs/personal.json, kept off GitHub.
BUY_OVERRIDES: dict[str, Decimal] = {}   # ticker -> confirmed buy amount, where it differs from broker cost − income earned
BTC_VALUE_ON_AS_OF = Decimal(0)          # market value, NOT cost — flagged TEMP until the user enters cost + quantity
MONTHLY_CONTRIBUTION = 0
SECTORS = {"BTC": "Crypto"}  # GICS sector per ticker; your stocks are added from inputs/personal.json
ISIN_COUNTRY = {"US": "United States", "CA": "Canada", "IE": "Ireland"}

TX_LAST = 1001      # Transactions formula rows 2..1001
HOLD_LAST = 41      # Holdings rows 2..41
GROWTH_LAST = 61    # Dividend Growth rows 2..61
CAL_FIRST, CAL_LAST = 4, 203
GRID_LAST = 43      # Calendar monthly grid rows 4..43
ENGINE_LAST = 482   # 481 months (0..480)

MONEY = '"$"#,##0.00'
MONEY0 = '"$"#,##0'
PCT = "0.00%"
PCT1 = "0.0%"
SHARES = "0.00000"
DATE = "yyyy-mm-dd"

NAVY = PatternFill("solid", fgColor="1F2A44")
INPUT = PatternFill("solid", fgColor="FFF7D6")
SECTION = PatternFill("solid", fgColor="E8EEF7")
RED = PatternFill("solid", fgColor="F8D7DA")
AMBER = PatternFill("solid", fgColor="FFF3CD")
GREEN_FONT = Font(color="1E7E34")
RED_FONT = Font(color="B02A37")
WHITE_BOLD = Font(color="FFFFFF", bold=True)
BOLD = Font(bold=True)
TITLE = Font(bold=True, size=16)
MUTED = Font(italic=True, color="6C757D")


def header(ws, row, labels, widths=None, notes=None):
    for i, label in enumerate(labels, start=1):
        c = ws.cell(row=row, column=i, value=label)
        c.fill, c.font = NAVY, WHITE_BOLD
        c.alignment = Alignment(wrap_text=True, vertical="center")
        if notes and label in notes:
            c.comment = Comment(notes[label], "tracker")
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    ws.row_dimensions[row].height = 32


def fmt(ws, col, first, last, number_format):
    for r in range(first, last + 1):
        ws[f"{col}{r}"].number_format = number_format


def name(wb, label, ref):
    wb.defined_names[label] = DefinedName(label, attr_text=ref)


def iso(d: str):
    return datetime.strptime(d, "%Y-%m-%d").date() if d else None


# ── Tabs ─────────────────────────────────────────────────────────────────────────

def build_transactions(wb, positions):
    ws = wb.create_sheet("Transactions")
    header(ws, 1, ["Date", "Ticker", "Type", "Action", "Amount (USD)", "Shares (optional)", "Shares (used)",
                   "Price/share", "Source", "Key", "Note", "Quote symbol", "Shares Δ", "Cost Δ", "Realized gain",
                   "Cash flow"],
           [12, 9, 8, 17, 13, 14, 13, 11, 12, 16, 60, 16, 11, 11, 12, 11],
           {"Date": "Trade date (or pay date for dividends).",
            "Amount (USD)": "Dollars spent (Buy), received (Sell) or reinvested (Dividend Reinvest). Adjustment rows may be negative.",
            "Shares (optional)": "Leave blank to estimate from that day's closing price; type the exact number from your broker confirmation when you have it.",
            "Shares (used)": "Formula — the exact shares if entered, otherwise Amount ÷ closing price on the date.",
            "Source": "manual / broker-seed / auto-drip (added by the script) / reconcile (from tools/reconcile.py).",
            "Key": "TICKER|PAYDATE for automatic dividend rows — prevents double-logging.",
            "Cost Δ": "Average-cost method (same as the broker): a sale removes shares × average cost."})
    ws.freeze_panes = "A2"

    rows = []
    for p in sorted(positions, key=lambda x: x.ticker):
        buy = BUY_OVERRIDES.get(p.ticker, p.cost - (p.income_earned or Decimal(0)))
        reinvested = p.cost - buy
        note = f"Opening balance from broker positions as of {p.as_of:%m/%d/%Y}; shares include reinvested dividends"
        rows.append([p.as_of, p.ticker, "Stock", "Buy", float(buy), float(p.quantity), "broker-seed", "", note])
        rnote = "Dividends reinvested before the opening balance (broker cost − original buy); the shares are on the Buy row"
        if p.ticker in BUY_OVERRIDES:
            rnote += f". Broker 'income earned' shows ${p.income_earned}; ${p.income_earned - reinvested:.2f} of it is not in the cost basis"
        rows.append([p.as_of, p.ticker, "Stock", "Dividend Reinvest", float(reinvested), 0, "broker-seed", "", rnote])
    rows.append([AS_OF, "BTC", "Crypto", "Buy", float(BTC_VALUE_ON_AS_OF), None, "manual", "",
                 f"TEMP: ${BTC_VALUE_ON_AS_OF} is the market value on {AS_OF:%m/%d/%Y}, not your cost. Replace Amount with what you paid "
                 "and Shares with your exact BTC quantity, then delete this note."])

    for i, (d, t, typ, act, amt, sh, src, key, note) in enumerate(rows, start=2):
        ws.cell(i, 1, d), ws.cell(i, 2, t), ws.cell(i, 3, typ), ws.cell(i, 4, act), ws.cell(i, 5, amt)
        if sh is not None:
            ws.cell(i, 6, sh)
        ws.cell(i, 9, src), ws.cell(i, 10, key), ws.cell(i, 11, note)

    for r in range(2, TX_LAST + 1):
        gf = f'GOOGLEFINANCE(L{r},IF(C{r}="Crypto","price","close"),A{r}-7,A{r})'
        ws[f"G{r}"] = (f'=IF(A{r}="","",IF(F{r}<>"",F{r},IF(OR(D{r}="Dividend Cash",D{r}="Adjustment"),0,'
                       f'IFERROR(E{r}/INDEX({gf},ROWS({gf}),2),"price?"))))')
        ws[f"H{r}"] = f'=IF(OR(A{r}="",NOT(ISNUMBER(G{r}))),"",IF(G{r}=0,"",E{r}/G{r}))'
        ws[f"L{r}"] = f'=IF(B{r}="","",IF(C{r}="Crypto","CURRENCY:"&B{r}&"USD",B{r}))'
        ws[f"M{r}"] = f'=IF(OR(A{r}="",NOT(ISNUMBER(G{r}))),0,IF(D{r}="Sell",-G{r},IF(D{r}="Dividend Cash",0,G{r})))'
        ws[f"N{r}"] = (f'=IF(A{r}="",0,IF(D{r}="Sell",-IFERROR(G{r}*SUMIFS(N$1:N{r-1},B$1:B{r-1},B{r})'
                       f'/SUMIFS(M$1:M{r-1},B$1:B{r-1},B{r}),0),IF(D{r}="Dividend Cash",0,E{r})))')
        ws[f"O{r}"] = f'=IF(D{r}="Sell",E{r}+N{r},"")'
        ws[f"P{r}"] = f'=IF(A{r}="","",IF(D{r}="Buy",-E{r},IF(D{r}="Sell",E{r},0)))'
        for col in "ABCDEFK":
            ws[f"{col}{r}"].fill = INPUT
    fmt(ws, "A", 2, TX_LAST, DATE)
    for col in "EHNOP":
        fmt(ws, col, 2, TX_LAST, MONEY)
    for col in "FGM":
        fmt(ws, col, 2, TX_LAST, SHARES)

    validations = [
        ("A", 'AND(ISNUMBER(A2),A2>=DATE(2000,1,1),A2<=TODAY()+7)', "Enter a real date (YYYY-MM-DD)."),
        ("B", f'REGEXMATCH(B2,"{TICKER_PATTERN}")', "Ticker in capitals, e.g. ABC, BRK.B, BTC."),
        ("E", 'OR(D2="Adjustment",AND(ISNUMBER(E2),E2>=0))', "Amount must be a positive number (only Adjustment rows may be negative)."),
        ("F", 'OR(F2="",AND(ISNUMBER(F2),OR(D2="Adjustment",F2>=0)))', "Shares must be a positive number or blank."),
    ]
    for col, formula, msg in validations:
        dv = DataValidation(type="custom", formula1=formula, allow_blank=True, showErrorMessage=True,
                            errorTitle="Invalid entry", error=msg)
        dv.add(f"{col}2:{col}{TX_LAST}")
        ws.add_data_validation(dv)
    for col, items in (("C", "Stock,Crypto"), ("D", "Buy,Sell,Dividend Reinvest,Dividend Cash,Adjustment")):
        dv = DataValidation(type="list", formula1=f'"{items}"', allow_blank=True, showErrorMessage=True)
        dv.add(f"{col}2:{col}{TX_LAST}")
        ws.add_data_validation(dv)
    ws.conditional_formatting.add(f"G2:G{TX_LAST}", FormulaRule(formula=['G2="price?"'], fill=RED))
    ws.conditional_formatting.add(f"A2:P{TX_LAST}", FormulaRule(formula=['LEFT($K2,4)="TEMP"'], fill=AMBER))
    return ws


def build_holdings(wb):
    ws = wb.create_sheet("Holdings")
    header(ws, 1, ["Ticker", "Type", "Name", "Sector", "Shares", "Cost basis", "Avg cost", "Price", "Value",
                   "Unrealized $", "Unrealized %", "Weight", "Day %", "Fwd div/share", "Fwd annual income",
                   "Yield", "Yield on cost", "Dividends received", "Realized gain", "Money in", "Country"],
           [9, 8, 26, 22, 11, 12, 11, 11, 12, 12, 11, 9, 8, 12, 13, 8, 10, 12, 11, 11, 14],
           {"Cost basis": "Everything you paid, including reinvested dividends (matches the broker's 'Cost').",
            "Money in": "Only your own money (Buy − Sell). Dividends reinvested are not counted here.",
            "Price": "Live GOOGLEFINANCE price (~20 min delay). #N/A = price unavailable; totals skip it and the Dashboard warns.",
            "Fwd div/share": "Current regular dividend × payments per year (Dividend Growth tab)."})
    ws.freeze_panes = "B2"
    ws["A2"] = '=IFERROR(SORT(UNIQUE(FILTER(Transactions!B2:B1001,Transactions!B2:B1001<>""))),"")'
    tx = "Transactions!"
    for r in range(2, HOLD_LAST + 1):
        g = f'IF(A{r}="","",'
        ws[f"B{r}"] = f'={g}IFERROR(VLOOKUP(A{r},{tx}$B$2:$C${TX_LAST},2,FALSE),"Stock"))'
        ws[f"C{r}"] = f'={g}IF(B{r}="Crypto",A{r},IFERROR(GOOGLEFINANCE(A{r},"name"),A{r})))'
        ws[f"D{r}"] = f'={g}IFERROR(VLOOKUP(A{r},Sectors!$A:$C,2,FALSE),"⚠ add to Sectors tab"))'
        ws[f"E{r}"] = f'={g}SUMIFS({tx}$M$2:$M${TX_LAST},{tx}$B$2:$B${TX_LAST},A{r}))'
        ws[f"F{r}"] = f'={g}SUMIFS({tx}$N$2:$N${TX_LAST},{tx}$B$2:$B${TX_LAST},A{r}))'
        ws[f"G{r}"] = f'={g}IFERROR(F{r}/E{r},""))'
        ws[f"H{r}"] = (f'={g}IFERROR(IF(B{r}="Crypto",GOOGLEFINANCE("CURRENCY:"&A{r}&"USD"),'
                       f'GOOGLEFINANCE(A{r},"price")),NA()))')
        ws[f"I{r}"] = f'={g}IFERROR(E{r}*H{r},NA()))'
        ws[f"J{r}"] = f'={g}IFERROR(I{r}-F{r},""))'
        ws[f"K{r}"] = f'={g}IFERROR(J{r}/F{r},""))'
        ws[f"L{r}"] = f'={g}IFERROR(I{r}/(StocksValue+CryptoValue),""))'
        ws[f"M{r}"] = f'=IF(OR(A{r}="",B{r}="Crypto"),"",IFERROR(GOOGLEFINANCE(A{r},"changepct")/100,""))'
        ws[f"N{r}"] = f"={g}IFERROR(VLOOKUP(A{r},'Dividend Growth'!$A:$E,5,FALSE),0))"
        ws[f"O{r}"] = f'={g}E{r}*N{r})'
        ws[f"P{r}"] = f'={g}IFERROR(N{r}/H{r},""))'
        ws[f"Q{r}"] = f'={g}IFERROR(O{r}/F{r},""))'
        ws[f"R{r}"] = f'={g}SUMIFS({tx}$E$2:$E${TX_LAST},{tx}$B$2:$B${TX_LAST},A{r},{tx}$D$2:$D${TX_LAST},"Dividend*"))'
        ws[f"S{r}"] = f'={g}SUMIFS({tx}$O$2:$O${TX_LAST},{tx}$B$2:$B${TX_LAST},A{r}))'
        ws[f"T{r}"] = (f'={g}SUMIFS({tx}$E$2:$E${TX_LAST},{tx}$B$2:$B${TX_LAST},A{r},{tx}$D$2:$D${TX_LAST},"Buy")'
                       f'-SUMIFS({tx}$E$2:$E${TX_LAST},{tx}$B$2:$B${TX_LAST},A{r},{tx}$D$2:$D${TX_LAST},"Sell"))')
        ws[f"U{r}"] = f'={g}IFERROR(VLOOKUP(A{r},Sectors!$A:$C,3,FALSE),""))'
    for col in "FGHIJNORST":
        fmt(ws, col, 2, HOLD_LAST, MONEY)
    for col in "KLMPQ":
        fmt(ws, col, 2, HOLD_LAST, PCT)
    fmt(ws, "E", 2, HOLD_LAST, SHARES)
    ws.conditional_formatting.add(f"J2:K{HOLD_LAST}", FormulaRule(formula=["AND(ISNUMBER(J2),J2<0)"], font=RED_FONT))
    ws.conditional_formatting.add(f"J2:K{HOLD_LAST}", FormulaRule(formula=["AND(ISNUMBER(J2),J2>0)"], font=GREEN_FONT))
    ws.conditional_formatting.add(f"H2:I{HOLD_LAST}", FormulaRule(formula=["ISERROR(H2)"], fill=RED))
    return ws


GROWTH_HEADERS = ["Ticker", "Frequency", "Payments/yr", "Current $/payment", "Fwd annual $/share",
                  "Last raise (declared)", "Last raise %", "Next raise (est.)", "1Y growth", "3Y CAGR", "5Y CAGR",
                  "10Y CAGR", "Raise streak (yrs)", "Streak note", "Status", "Specials excluded", "Stats as of",
                  "Price", "Yield", "Yield bucket", "My shares", "My fwd income", "Yield on cost", "Payout ratio"]


def build_growth(wb, prefill):
    ws = wb.create_sheet("Dividend Growth")
    header(ws, 1, GROWTH_HEADERS,
           [9, 11, 10, 12, 12, 13, 10, 12, 10, 10, 10, 10, 11, 30, 42, 10, 11, 10, 9, 15, 11, 12, 11, 11],
           {"Last raise %": "Change in the annualized regular dividend at the last raise.",
            "1Y growth": "Current regular rate vs the rate 1 year earlier (split-adjusted, special dividends excluded).",
            "5Y CAGR": "Compound annual growth of the regular dividend over 5 years. The Projections base case uses the value-weighted average of this column.",
            "Raise streak (yrs)": "Consecutive calendar years in which the average annualized dividend rose. History in the data starts in 1999, so long-time raisers show ≥27.",
            "Status": "CUT = a cut in the last 12 months. 'No raise in 15+ months' = possible freeze. Notes explain data caveats.",
            "Specials excluded": "One-off special dividends found in the history and left out of growth math.",
            "Payout ratio": "Finnhub, EPS-based. For REITs the EPS payout is not meaningful — they report payout vs AFFO, which is lower.",
            "Columns A–Q": "Written daily by the Apps Script (computeDividendStats). Columns R–X are formulas."})
    ws.freeze_panes = "B2"
    for i, s in enumerate(prefill["stats"], start=2):
        vals = [s["ticker"], s["frequency"], s["perYear"], s["current"], s["fwdAnnual"], iso(s["lastRaiseDate"]),
                s["lastRaisePct"], s["nextRaiseEst"], s["cagr1"], s["cagr3"], s["cagr5"], s["cagr10"], s["streak"],
                s["streakNote"], s["status"], s["specials"], iso(s["asOf"])]
        for j, v in enumerate(vals, start=1):
            ws.cell(i, j, None if v == "" else v)
    for r in range(2, GROWTH_LAST + 1):
        g = f'IF(A{r}="","",'
        ws[f"R{r}"] = f'={g}IFERROR(VLOOKUP(A{r},Holdings!$A:$H,8,FALSE),""))'
        ws[f"S{r}"] = f'={g}IFERROR(E{r}/R{r},""))'
        ws[f"T{r}"] = f'=IF(NOT(ISNUMBER(S{r})),"",IF(S{r}>=0.03,"High (≥3%)",IF(S{r}>=0.02,"Medium (2–3%)","Low (<2%)")))'
        ws[f"U{r}"] = f'={g}IFERROR(VLOOKUP(A{r},Holdings!$A:$E,5,FALSE),""))'
        ws[f"V{r}"] = f'={g}IFERROR(U{r}*E{r},""))'
        ws[f"W{r}"] = f'={g}IFERROR(V{r}/VLOOKUP(A{r},Holdings!$A:$F,6,FALSE),""))'
        ws[f"X{r}"] = f'={g}IFERROR(VLOOKUP(A{r},Fundamentals!$A:$B,2,FALSE),""))'
    for col in "DERV":
        fmt(ws, col, 2, GROWTH_LAST, MONEY)
    for col in "GIJKLSWX":
        fmt(ws, col, 2, GROWTH_LAST, PCT1)
    fmt(ws, "F", 2, GROWTH_LAST, DATE)
    fmt(ws, "Q", 2, GROWTH_LAST, DATE)
    fmt(ws, "U", 2, GROWTH_LAST, SHARES)
    ws.conditional_formatting.add(f"A2:X{GROWTH_LAST}", FormulaRule(formula=['ISNUMBER(SEARCH("CUT",$O2))'], fill=RED))
    ws.conditional_formatting.add(f"A2:X{GROWTH_LAST}", FormulaRule(formula=['ISNUMBER(SEARCH("No raise",$O2))'], fill=AMBER))
    ws.conditional_formatting.add(f"I2:L{GROWTH_LAST}", FormulaRule(formula=["AND(ISNUMBER(I2),I2<0)"], font=RED_FONT))
    return ws


def build_calendar(wb):
    ws = wb.create_sheet("Dividend Calendar")
    ws["A1"] = "Dividend Calendar"
    ws["A1"].font = TITLE
    ws["A2"] = ("Declared = announced by the company. est. = projected from last year's dates at the current rate "
                "(not confirmed). You must own the shares before the ex-date to get the payment. Source: Alpha Vantage, refreshed daily.")
    ws["A2"].font = MUTED
    labels = ["Ticker", "Ex-date", "Pay date", "$/share", "Status", "My shares", "My expected $", "Alert"]
    for i, l in enumerate(labels, start=1):
        c = ws.cell(3, i, l)
        c.fill, c.font = NAVY, WHITE_BOLD
    for col, w in zip("ABCDEFGH", [9, 12, 12, 10, 22, 11, 13, 44]):
        ws.column_dimensions[col].width = w
    ws["A4"] = ('=IFERROR(SORT(FILTER(DivSchedule!A2:E2000,DivSchedule!C2:C2000>=TODAY(),DivSchedule!C2:C2000<=EDATE(TODAY(),12)),3,TRUE),'
                '"No dividend data yet — run Portfolio → Refresh dividend data now")')
    for r in range(CAL_FIRST, CAL_LAST + 1):
        ws[f"F{r}"] = f'=IF(OR($A{r}="",NOT(ISNUMBER($B{r}))),"",IFERROR(VLOOKUP($A{r},Holdings!$A:$E,5,FALSE),0))'
        ws[f"G{r}"] = f'=IF(F{r}="","",F{r}*D{r})'
        ws[f"H{r}"] = (f'=IF(F{r}="","",IF(AND(B{r}>=TODAY(),B{r}-TODAY()<=7),"⚠ Ex-date in "&(B{r}-TODAY())&" days — '
                       f'own the shares before then",IF(C{r}-TODAY()<=30,"Pays within 30 days","")))')
    fmt(ws, "B", CAL_FIRST, CAL_LAST, DATE)
    fmt(ws, "C", CAL_FIRST, CAL_LAST, DATE)
    fmt(ws, "D", CAL_FIRST, CAL_LAST, "$0.0000")
    fmt(ws, "F", CAL_FIRST, CAL_LAST, SHARES)
    fmt(ws, "G", CAL_FIRST, CAL_LAST, MONEY)
    ws.conditional_formatting.add(f"A{CAL_FIRST}:H{CAL_LAST}", FormulaRule(formula=[f'LEFT($H{CAL_FIRST},1)="⚠"'], fill=AMBER))
    ws.conditional_formatting.add(f"A{CAL_FIRST}:H{CAL_LAST}", FormulaRule(formula=[f'$E{CAL_FIRST}="est."'], font=MUTED))
    ws.freeze_panes = "A4"

    # Monthly grid: expected income by stock for the next 12 months.
    ws["J2"] = "Expected income by month (next 12 months, your current shares)"
    ws["J2"].font = BOLD
    ws["J3"] = "Ticker"
    for k in range(12):
        col = chr(ord("K") + k)
        ws[f"{col}3"] = f"=EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),{k})"
        ws[f"{col}3"].number_format = "mmm yy"
        ws.column_dimensions[col].width = 9
    ws["W3"] = "12-mo total"
    for c in ["J3", "W3"] + [f"{chr(ord('K') + k)}3" for k in range(12)]:
        ws[c].fill, ws[c].font = NAVY, WHITE_BOLD
    ws.column_dimensions["J"].width = 14
    ws.column_dimensions["W"].width = 12
    ws["J4"] = '=IFERROR(FILTER(Holdings!A2:A41,Holdings!B2:B41="Stock",Holdings!A2:A41<>""),"")'
    for r in range(4, GRID_LAST + 1):
        for k in range(12):
            col = chr(ord("K") + k)
            ws[f"{col}{r}"] = (f'=IF($J{r}="","",SUMIFS(DivSchedule!$D:$D,DivSchedule!$A:$A,$J{r},'
                               f'DivSchedule!$C:$C,">="&{col}$3,DivSchedule!$C:$C,"<"&EDATE({col}$3,1))'
                               f'*IFERROR(VLOOKUP($J{r},Holdings!$A:$E,5,FALSE),0))')
            ws[f"{col}{r}"].number_format = MONEY
        ws[f"W{r}"] = f'=IF($J{r}="","",SUM(K{r}:V{r}))'
        ws[f"W{r}"].number_format = MONEY
    tr = GRID_LAST + 1
    ws[f"J{tr}"] = "Total expected"
    ws[f"J{tr}"].font = BOLD
    for k in range(13):
        col = chr(ord("K") + k)
        ws[f"{col}{tr}"] = f"=SUM({col}4:{col}{GRID_LAST})"
        ws[f"{col}{tr}"].number_format, ws[f"{col}{tr}"].font = MONEY, BOLD

    # Received (actual) in the last 12 months, from Transactions.
    hr = tr + 2
    ws[f"J{hr}"] = "Received (logged)"
    ws[f"J{hr}"].fill, ws[f"J{hr}"].font = NAVY, WHITE_BOLD
    for k in range(12):
        col = chr(ord("K") + k)
        ws[f"{col}{hr}"] = f"=EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),{k - 11})"
        ws[f"{col}{hr}"].number_format = "mmm yy"
        ws[f"{col}{hr}"].fill, ws[f"{col}{hr}"].font = NAVY, WHITE_BOLD
        ws[f"{col}{hr + 1}"] = (f'=SUMIFS(Transactions!$E$2:$E${TX_LAST},Transactions!$D$2:$D${TX_LAST},"Dividend*",'
                                f'Transactions!$I$2:$I${TX_LAST},"<>broker-seed",'
                                f'Transactions!$A$2:$A${TX_LAST},">="&{col}{hr},Transactions!$A$2:$A${TX_LAST},"<"&EDATE({col}{hr},1))')
        ws[f"{col}{hr + 1}"].number_format = MONEY
    ws[f"W{hr}"] = "12-mo total"
    ws[f"W{hr}"].fill, ws[f"W{hr}"].font = NAVY, WHITE_BOLD
    ws[f"J{hr + 1}"] = "Dividends"
    ws[f"W{hr + 1}"] = f"=SUM(K{hr + 1}:V{hr + 1})"
    ws[f"W{hr + 1}"].number_format = MONEY

    chart = BarChart()
    chart.title = "Expected dividend income — next 12 months"
    chart.y_axis.title = "USD"
    chart.add_data(Reference(ws, min_col=11, max_col=22, min_row=tr, max_row=tr), from_rows=True, titles_from_data=False)
    chart.set_categories(Reference(ws, min_col=11, max_col=22, min_row=3, max_row=3))
    chart.legend = None
    chart.height, chart.width = 8, 20
    ws.add_chart(chart, f"J{hr + 4}")
    return ws


def build_projections(wb):
    eng = wb.create_sheet("ProjEngine")
    header(eng, 1, ["Month"] + [f"{s} {k}" for s in ("Cons", "Base", "Opt")
                                for k in ("price", "div/share", "shares", "value", "annual income", "cum. dividends")])
    eng["A2"] = 0
    blocks = [("B", "GpCons", "GdCons"), ("H", "GpBase", "GdBase"), ("N", "GpOpt", "GdOpt")]
    for start, gp, gd in blocks:
        c = [chr(ord(start) + i) for i in range(6)]  # P, d, S, V, I, Cum
        P, D, S, V, I, CUM = c
        eng[f"{P}2"], eng[f"{D}2"], eng[f"{S}2"] = 1, "=StartYield", "=TotalValue"
        eng[f"{V}2"], eng[f"{I}2"], eng[f"{CUM}2"] = f"={S}2*{P}2", f"={S}2*{D}2", 0
        for r in range(3, ENGINE_LAST + 1):
            p = r - 1
            eng[f"{P}{r}"] = f"={P}{p}*(1+{gp})^(1/12)"
            eng[f"{D}{r}"] = f"={D}{p}*(1+IF(A{r}<=12*DivGrowthYears,{gd},{gp}))^(1/12)"
            eng[f"{S}{r}"] = f"={S}{p}+({S}{p}*{D}{p}/12+MonthlyContribution)/{P}{r}"
            eng[f"{V}{r}"] = f"={S}{r}*{P}{r}"
            eng[f"{I}{r}"] = f"={S}{r}*{D}{r}"
            eng[f"{CUM}{r}"] = f"={CUM}{p}+{S}{p}*{D}{p}/12"
    for r in range(3, ENGINE_LAST + 1):
        eng[f"A{r}"] = f"=A{r - 1}+1"

    ws = wb.create_sheet("Projections")
    ws["A1"] = "Projections — dividends reinvested, monthly contributions"
    ws["A1"].font = TITLE
    ws["A2"] = ("Model: the portfolio is treated as one asset. Each month the price grows at the scenario's price growth; "
                "the dividend per share grows at the scenario's dividend growth for the first N years (Settings), then at the "
                "price growth rate; every dividend plus your monthly contribution buys more shares. Assumptions, not forecasts.")
    ws["A2"].font = MUTED
    ws["A2"].alignment = Alignment(wrap_text=True)
    ws.merge_cells("A2:Q2")
    ws.row_dimensions[2].height = 45
    for i, (label, formula, f) in enumerate([("Starting value", "=TotalValue", MONEY),
                                             ("Starting forward dividend income", "=FwdIncome", MONEY),
                                             ("Starting yield", "=StartYield", PCT),
                                             ("Monthly contribution", "=MonthlyContribution", MONEY),
                                             ("Inflation (for today's $ columns)", "=Inflation", PCT),
                                             ("Years at stage-1 dividend growth", "=DivGrowthYears", "0")], start=4):
        ws[f"A{i}"], ws[f"C{i}"] = label, formula
        ws[f"C{i}"].number_format = f
    ws["F4"], ws["G4"], ws["H4"] = "Scenario", "Price growth/yr", "Dividend growth/yr"
    for i, (s, gp, gd) in enumerate([("Conservative", "GpCons", "GdCons"), ("Base", "GpBase", "GdBase"),
                                     ("Optimistic", "GpOpt", "GdOpt")], start=5):
        ws[f"F{i}"], ws[f"G{i}"], ws[f"H{i}"] = s, f"={gp}", f"={gd}"
        ws[f"G{i}"].number_format = ws[f"H{i}"].number_format = PCT
    for c in ("F4", "G4", "H4"):
        ws[c].font = BOLD

    groups = [("C", "Conservative", "E", "F", "G"), ("H", "Base", "K", "L", "M"), ("M", "Optimistic", "Q", "R", "S")]
    labels = ["Year", "Contributed (from today)"]
    for _, g, *_ in groups:
        labels += [f"{g}: value", f"{g}: value (today's $)", f"{g}: annual dividends", f"{g}: monthly dividends",
                   f"{g}: dividends received (total)"]
    header(ws, 12, labels, [7, 13] + [13] * 15)
    for start, g, vcol, icol, ccol in groups:
        ws[f"{start}11"] = g
        ws[f"{start}11"].font = BOLD
        ws[f"{start}11"].fill = SECTION
    for y in range(0, 41):
        r = 13 + y
        ws[f"A{r}"] = y
        ws[f"B{r}"] = f"=12*$A{r}*MonthlyContribution"
        ws[f"B{r}"].number_format = MONEY0
        for start, _, vcol, icol, ccol in groups:
            cols = [chr(ord(start) + i) for i in range(5)]
            idx = f"12*$A{r}+1"
            ws[f"{cols[0]}{r}"] = f"=INDEX(ProjEngine!${vcol}$2:${vcol}${ENGINE_LAST},{idx})"
            ws[f"{cols[1]}{r}"] = f"={cols[0]}{r}/(1+Inflation)^$A{r}"
            ws[f"{cols[2]}{r}"] = f"=INDEX(ProjEngine!${icol}$2:${icol}${ENGINE_LAST},{idx})"
            ws[f"{cols[3]}{r}"] = f"={cols[2]}{r}/12"
            ws[f"{cols[4]}{r}"] = f"=INDEX(ProjEngine!${ccol}$2:${ccol}${ENGINE_LAST},{idx})"
            for c in cols:
                ws[f"{c}{r}"].number_format = MONEY0
    ws.conditional_formatting.add("A13:Q53", FormulaRule(formula=["MOD($A13,5)=0"], font=BOLD, fill=SECTION))
    ws.freeze_panes = "C13"

    line = LineChart()
    line.title = "Portfolio value by scenario"
    line.x_axis.title, line.y_axis.title = "Years from today", "USD"
    for col in (3, 8, 13):
        line.add_data(Reference(ws, min_col=col, min_row=12, max_row=53), titles_from_data=True)
    line.set_categories(Reference(ws, min_col=1, min_row=13, max_row=53))
    line.height, line.width = 9, 20
    ws.add_chart(line, "A56")
    inc = LineChart()
    inc.title = "Monthly dividend income by scenario"
    inc.x_axis.title, inc.y_axis.title = "Years from today", "USD / month"
    for col in (6, 11, 16):
        inc.add_data(Reference(ws, min_col=col, min_row=12, max_row=53), titles_from_data=True)
    inc.set_categories(Reference(ws, min_col=1, min_row=13, max_row=53))
    inc.height, inc.width = 9, 20
    ws.add_chart(inc, "J56")
    return ws, eng


def build_dashboard(wb):
    """Placeholder: the Apps Script draws the real Dashboard (Portfolio → Rebuild dashboard), because
    Google Sheets map charts cannot be imported from .xlsx. It also defines the named totals on a Calc tab."""
    ws = wb.create_sheet("Dashboard", 0)
    ws["B2"] = "Investment Dashboard"
    ws["B2"].font = TITLE
    ws["B4"] = ("Finish setup: Extensions → Apps Script (paste apps_script/Code.gs), reload, then "
                "Portfolio → Set up automatic triggers. The dashboard, charts and map are drawn here.")
    ws["B4"].font = MUTED
    return ws


def build_settings(wb):
    ws = wb.create_sheet("Settings")
    ws["A1"] = "Settings"
    ws["A1"].font = TITLE
    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 90
    inputs = [("Monthly contribution (USD)", MONTHLY_CONTRIBUTION, MONEY, "MonthlyContribution",
               f"From your old sheet ('Investment MM' = {MONTHLY_CONTRIBUTION})."),
              ("Cash balance (USD)", 0, MONEY, "CashBalance", "Uninvested cash at the broker, if any."),
              ("Inflation for 'today's dollars' (0 = off)", 0, PCT, "Inflation", "Set e.g. 3% to see projected values in today's purchasing power."),
              ("Years dividends grow at the scenario rate", 5, "0", "DivGrowthYears",
               "Two-stage model: after these years dividend growth fades to the price growth rate, so the yield stops drifting. Set 40 to keep it the whole time.")]
    for i, (label, value, f, nm, note) in enumerate(inputs, start=3):
        ws[f"A{i}"], ws[f"B{i}"], ws[f"D{i}"] = label, value, note
        ws[f"B{i}"].number_format, ws[f"B{i}"].fill = f, INPUT
        ws[f"D{i}"].font = MUTED
        name(wb, nm, f"Settings!$B${i}")
    for c, l in zip("ABCD", ["Projection scenario", "Price growth / yr", "Dividend growth / yr (stage 1)", "Where the default comes from"]):
        ws[f"{c}7"] = l
        ws[f"{c}7"].fill, ws[f"{c}7"].font = NAVY, WHITE_BOLD
    scen = [
        ("Conservative", "=B9-0.02", "=MAX(0,C9-0.02)", "Base − 2 percentage points (dividend growth floored at 0%)."),
        ("Base", 0.06, "=IFERROR(ROUND(PortfolioDivCAGR5,4),0.05)",
         "Price 6% = your old sheet's assumption. Dividend growth = your portfolio's measured 5-year dividend CAGR, weighted "
         "by your dividend income (Dashboard). Type over either cell to override."),
        ("Optimistic", "=B9+0.02", "=C9+0.02", "Base + 2 percentage points."),
    ]
    for i, (s, gp, gd, note) in enumerate(scen, start=8):
        ws[f"A{i}"], ws[f"B{i}"], ws[f"C{i}"], ws[f"D{i}"] = s, gp, gd, note
        ws[f"B{i}"].number_format = ws[f"C{i}"].number_format = PCT
        ws[f"D{i}"].font = MUTED
        ws[f"D{i}"].alignment = Alignment(wrap_text=True)
    ws["B9"].fill = ws["C9"].fill = INPUT
    for nm, ref in (("GpCons", "B8"), ("GdCons", "C8"), ("GpBase", "B9"), ("GdBase", "C9"), ("GpOpt", "B10"), ("GdOpt", "C10")):
        name(wb, nm, f"Settings!${ref[0]}${ref[1:]}")
    ws["A11"] = ("Why two stages: your measured 5-year dividend growth (≈8%) is above the 6% price growth. If that gap lasted "
                 "40 years the yield would more than double, which rarely happens — fast dividend growth tends to slow down.")
    ws["A11"].font = MUTED
    ws.merge_cells("A11:D11")
    ws["A19"], ws["B19"], ws["D19"] = ("Target weight per sector", 0.1,
                                       "Used for the Dashboard over/underweight boxes (your old sheet used 10%).")
    ws["B19"].number_format, ws["B19"].fill = PCT, INPUT
    ws["D19"].font = MUTED
    name(wb, "SectorTarget", "Settings!$B$19")
    rows = [("Data sources", ""),
            ("Prices", "GOOGLEFINANCE (~20 minute delay), refreshes on its own while the file is open or shared."),
            ("Dividend history, ex/pay dates, splits", "Alpha Vantage DIVIDENDS + SPLITS, fetched daily by the Apps Script (free tier: 25 calls/day)."),
            ("Payout ratio", "Finnhub stock/metric (free tier), daily."),
            ("API keys", "Stored in Apps Script → Project Settings → Script Properties as ALPHAVANTAGE_KEY and FINNHUB_KEY. Never type keys into the sheet.")]
    for i, (a, b) in enumerate(rows, start=13):
        ws[f"A{i}"], ws[f"B{i}"] = a, b
    ws["A13"].fill, ws["A13"].font = NAVY, WHITE_BOLD
    return ws


def build_sectors(wb, positions):
    ws = wb.create_sheet("Sectors")
    header(ws, 1, ["Ticker", "Sector", "Country", "Note"], [9, 24, 16, 70])
    rows = [(p.ticker, SECTORS[p.ticker], ISIN_COUNTRY.get(p.isin[:2], p.isin[:2]), f"Country from ISIN {p.isin}")
            for p in sorted(positions, key=lambda x: x.ticker)]
    rows.append(("BTC", "Crypto", "—", ""))
    for i, r in enumerate(rows, start=2):
        for j, v in enumerate(r, start=1):
            ws.cell(i, j, v)
        ws.cell(i, 2).fill = ws.cell(i, 3).fill = INPUT
    ws.cell(len(rows) + 3, 1, "Add a row here when you buy a new ticker (GICS sector names keep the allocation chart consistent).").font = MUTED
    return ws


def build_pv_history(wb, old_path: Path | None):
    ws = wb.create_sheet("PV History")
    header(ws, 1, ["Date", "Total value", "Money in", "Cost basis", "Fwd annual income", "Note"], [12, 13, 12, 12, 15, 40])
    r = 2
    if old_path:
        old = load_workbook(old_path, read_only=True, data_only=True)["PV History"]
        for d, v in old.iter_rows(min_row=2, max_col=2, values_only=True):
            if isinstance(d, (int, float)):
                d = from_excel(d)
            if isinstance(d, datetime) and isinstance(v, (int, float)):
                ws.cell(r, 1, d.date()), ws.cell(r, 2, float(v)), ws.cell(r, 6, "imported from old sheet (stock value only)")
                r += 1
    fmt(ws, "A", 2, 2000, DATE)
    for col in "BCDE":
        fmt(ws, col, 2, 2000, MONEY)
    chart = LineChart()
    chart.title = "Portfolio value history"
    chart.add_data(Reference(ws, min_col=2, max_col=3, min_row=1, max_row=400), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=400))
    chart.height, chart.width = 9, 20
    ws.add_chart(chart, "H2")
    return ws


def build_data_tabs(wb, prefill, payout):
    fetched = prefill["fetchedAt"]
    tabs = {
        "DivHistory": (["Ticker", "Ex-date", "Declared", "Record", "Pay date", "Amount", "Fetched at"], prefill["divhistory"]),
        "Splits": (["Ticker", "Effective date", "Factor", "Fetched at"], prefill["splits"]),
        "DivSchedule": (["Ticker", "Ex-date", "Pay date", "$/share", "Status"],
                        [[s["ticker"], iso(s["ex"]), iso(s["pay"]), s["amount"], s["status"]] for s in prefill["schedule"]]),
        "Fundamentals": (["Ticker", "Payout ratio", "Fetched at"],
                         [[t, v if v is not None else "", fetched] for t, v in sorted(payout.items())]),
        "FetchMeta": (["Ticker", "Kind", "Last fetched"],
                      sorted(prefill["meta"] + [[t, "fund", fetched] for t in payout])),
    }
    for title, (labels, rows) in tabs.items():
        ws = wb.create_sheet(title)
        header(ws, 1, labels, [10] + [14] * (len(labels) - 1))
        for i, row in enumerate(rows, start=2):
            for j, val in enumerate(row, start=1):
                c = ws.cell(i, j, val)
                if title in ("DivHistory", "Splits") and 2 <= j <= 5 and isinstance(val, str):
                    c.number_format = "@"   # keep ISO dates as text; the script parses them
        if title == "DivSchedule":
            fmt(ws, "B", 2, len(rows) + 1, DATE)
            fmt(ws, "C", 2, len(rows) + 1, DATE)
        ws.sheet_state = "hidden"


def build_log(wb):
    ws = wb.create_sheet("Log")
    header(ws, 1, ["Timestamp", "Function", "Status", "Message"], [20, 24, 9, 110])
    ws["A2"] = datetime.now().replace(microsecond=0)
    ws["A2"].number_format = "yyyy-mm-dd hh:mm"
    ws["B2"], ws["C2"] = "build_workbook.py", "OK"
    ws["D2"] = f"Workbook generated from broker positions as of {AS_OF:%m/%d/%Y} and Alpha Vantage data fetched {AS_OF:%m/%d/%Y}"


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--broker", type=Path, required=True)
    ap.add_argument("--prefill", type=Path, required=True)
    ap.add_argument("--payout", type=Path, required=True)
    ap.add_argument("--old", type=Path)
    ap.add_argument("--personal", type=Path, default=Path("inputs/personal.json"))
    ap.add_argument("--out", type=Path, default=Path("Portfolio_Tracker_v2.xlsx"))
    args = ap.parse_args()

    global BTC_VALUE_ON_AS_OF, MONTHLY_CONTRIBUTION
    personal = json.loads(args.personal.read_text(encoding="utf-8"))
    BUY_OVERRIDES.update({t: Decimal(v) for t, v in personal["buy_overrides"].items()})
    BTC_VALUE_ON_AS_OF = Decimal(personal["btc_value_on_as_of"])
    MONTHLY_CONTRIBUTION = personal["monthly_contribution"]
    SECTORS.update(personal["sectors"])
    positions = load_broker_positions(args.broker)
    missing = [p.ticker for p in positions if p.ticker not in SECTORS]
    if missing:
        raise SystemExit(f"add a GICS sector for {missing} under \"sectors\" in {args.personal}")
    prefill = json.loads(args.prefill.read_text(encoding="utf-8"))
    payout = json.loads(args.payout.read_text(encoding="utf-8"))

    wb = Workbook()
    wb.remove(wb.active)
    build_transactions(wb, positions)
    build_holdings(wb)
    build_calendar(wb)
    build_growth(wb, prefill)
    build_projections(wb)
    build_dashboard(wb)
    build_pv_history(wb, args.old)
    build_settings(wb)
    build_sectors(wb, positions)
    build_log(wb)
    build_data_tabs(wb, prefill, payout)

    order = ["Dashboard", "Transactions", "Holdings", "Dividend Calendar", "Dividend Growth", "Projections",
             "PV History", "Settings", "Sectors", "Log", "DivHistory", "Splits", "DivSchedule", "Fundamentals",
             "FetchMeta", "ProjEngine"]
    wb._sheets = [wb[n] for n in order]
    wb["ProjEngine"].sheet_state = "hidden"
    wb.active = 0
    wb.save(args.out)
    print(f"wrote {args.out} ({len(positions)} positions + BTC, {len(prefill['schedule'])} scheduled payments)")


if __name__ == "__main__":
    main()
