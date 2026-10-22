# Dividend Portfolio Tracker

An automated Google Sheets tracker for a dividend-growth stock portfolio (plus Bitcoin), with live prices,
a dividend calendar, automatic logging of reinvested dividends (DRIP), long-term projections and a monthly
reconciliation against the broker's records.

I built it to replace my own spreadsheet after auditing it and finding **19 errors**, including a projection
formula that overstated the 40-year portfolio value by about **8×**, share counts that ignored every reinvested
dividend, and API keys stored in plain text.

> Personal tool and learning project, not financial advice. All personal data (holdings, amounts, broker exports)
> is kept out of this repository.

## Highlights

- **Audit first, then rebuild.** Every formula of the old sheet was checked against the broker's numbers; each error
  is fixed and covered by a test (for example, the old projection divided the dividend term by the price-growth rate).
- **Correct dividend math.** Split-adjusted dividend growth (1, 3, 5 and 10-year CAGR), special dividends detected and
  excluded, monthly vs quarterly payers, payers that changed frequency, and foreign payers whose USD amount moves with FX.
- **Automatic DRIP logging.** Each morning a Google Apps Script adds a `Dividend Reinvest` row for every paid dividend:
  shares held before the ex-date × dividend per share. Rows are keyed, so a payment is never logged twice.
- **Projections with three scenarios** (conservative / base / optimistic), monthly contributions, and a two-stage
  dividend-growth model. The sheet formulas are verified against an independent Python model.
- **Reconciliation.** A Python tool compares the sheet with the broker's positions export and writes the exact
  adjustment rows, so the tracker never drifts from the real account.
- **Data validation and security.** pydantic models reject inconsistent broker rows (value must equal quantity × price),
  tickers are validated before they reach a URL, CSV formula injection is neutralised, and API keys live only in
  Script Properties, never in code or logs.

## Tech

| Part | Tools |
|---|---|
| Spreadsheet | Google Sheets (GOOGLEFINANCE, QUERY, FILTER), generated with Python `openpyxl` |
| Automation | Google Apps Script (JavaScript): daily dividend data, DRIP rows, snapshots, dashboard, health checks |
| Data | Alpha Vantage (dividends, splits), Finnhub (payout ratios), GOOGLEFINANCE (prices) |
| Checks | Python `unittest` + `pydantic`, JavaScript tests run with macOS JavaScript for Automation |

## How it was built

I defined the requirements, audited my old spreadsheet against my broker statements, and decided the rules (for
example: never invent or estimate a number, reinvested dividends count toward cost basis but not money invested).
The code was written with an AI coding assistant (Claude), and I reviewed and tested the results against my real
account data.

**What I learned:** how dividend growth rates, yield on cost and DRIP affect long-term projections; why split
adjustments matter (an unadjusted 4:1 split looks like a 75% dividend cut); how small formula mistakes compound over
decades; and basic security practices for handling API keys and financial data.

---

## Files

| Path | What it is |
|---|---|
| `apps_script/Code.gs`, `apps_script/appsscript.json` | The script to paste into the sheet (Extensions → Apps Script). |
| `apps_script/tests/run_tests.js` | Tests for the script logic, using example data. |
| `tools/build_workbook.py` | Builds the tracker workbook (`.xlsx`) from a broker positions CSV. |
| `tools/reconcile.py` | Monthly check against a broker positions export. |
| `tools/models.py` | Validation rules (pydantic) for broker files and transactions. |
| `PROJECT_BRIEF.md` | The original project brief. |

**Kept private (not in this repository, see `.gitignore`):** the generated workbook, the audit report with my real
figures, broker exports and `inputs/personal.json` (my amounts and sector list), plus tests and test data that use my
actual holdings (`tools/test_private.py`, `apps_script/tests/run_private_tests.js`, `apps_script/tests/fixtures/`).

## One-time setup (about 15 minutes, on a computer)

1. **Get API keys.** Free keys from Alpha Vantage (alphavantage.co → "Get your free API key") and Finnhub.
2. **Import the workbook.**
   - Google Drive → New → File upload → the generated `.xlsx`.
   - Open it, then File → *Save as Google Sheets*.
   - In the new Google Sheet: File → Settings → set your **time zone** and locale **United States**.
3. **Add the script.**
   1. Open Extensions → Apps Script.
   2. Replace everything in `Code.gs` with `apps_script/Code.gs`.
   3. Project Settings (gear icon) → tick *Show "appsscript.json" manifest file*. Back in the editor, replace that file with `apps_script/appsscript.json`.
   4. Project Settings → **Script Properties** → add `ALPHAVANTAGE_KEY` and `FINNHUB_KEY` with your keys.
4. **Turn on the automation.**
   1. Reload the sheet. A **Portfolio** menu appears.
   2. Choose Portfolio → *Set up automatic triggers (run once)* and approve the permissions. It asks only for this spreadsheet, external requests to the two data providers, and triggers.
   3. Then run Portfolio → *Refresh dividend data now* once.
5. **Fix the Bitcoin row.** The highlighted Bitcoin row in Transactions holds a temporary value. Enter what you *paid* in Amount, your exact BTC quantity in Shares, and delete the "TEMP" note.

The **Dashboard** tab is drawn by the script, including the industry pie, the equity-vs-cost bars and the world map. It appears when you run *Set up automatic triggers*. To redraw it at any time, use Portfolio → **Rebuild dashboard**.

### Check after setup

| Where | Expect |
|---|---|
| Dashboard → Stocks | ≈ the total market value in your broker's positions export |
| Dashboard → Money in | your stock buys + the BTC row's amount (until you fix the BTC row) |
| Dashboard → Forward annual income | close to the broker's estimated annual income |
| Dividend Calendar | declared payments marked **Declared**; monthly payers appear every month |
| Health checks | 0 prices unavailable |

If a cell shows `#NAME?` after import, the file was opened in Excel or not converted. Use *Save as Google Sheets*.

## Everyday use (works in the Google Sheets iPhone app)

- **You bought something:** add a row at the bottom of **Transactions**.
  - Fill in Date, Ticker, Type (Stock/Crypto), Action **Buy**, and Amount (USD).
  - Leave *Shares* blank to estimate it from that day's closing price. Type the exact number from the broker's confirmation if you want it exact.
- **New ticker:** also add a row in **Sectors** (ticker, GICS sector, country).
- **You sold:** the same as a buy, with Action **Sell**, Amount = proceeds, and Shares = shares sold. Realized gain uses average cost, like the broker.
- **Dividends:** nothing to do (see below).

The Portfolio menu doesn't show in the iPhone app. You won't need it, because the triggers run on Google's servers.

## How reinvested dividends are handled

Every dividend is reinvested, so each payment adds shares and adds to your cost basis.
- **Every morning** the script checks each payment whose pay date has passed and that isn't logged yet.
- **It adds a `Dividend Reinvest` row:**
  - Amount = $/share × the shares you held before the ex-date.
  - Shares = Amount ÷ that day's closing price.
- **No duplicates:** each row carries a key (`TICKER|pay date`), so a payment can't be logged twice.
- **Dashboard figures:**
  - **Money in**: only your buys.
  - **Cost basis**: buys + reinvested dividends. This matches the broker's "Cost".
  - **Price gain %**: value vs cost basis.
  - **Total return %**: value vs money in. It includes dividends without counting them twice.

The broker buys the DRIP shares at its own price, so after a few months the share counts can differ by a few hundred-thousandths. The monthly reconcile fixes that.

## Monthly reconcile (5 minutes, on a computer)

1. In the sheet, open the **Transactions** tab → File → Download → **Comma-separated values**.
2. On the broker's website, export **Positions** to CSV.
3. Run:
   ```
   cd "Portfolio Tracking"
   python3 -m venv .venv && .venv/bin/pip install -r requirements.txt   # first time only
   .venv/bin/python tools/reconcile.py --transactions ~/Downloads/Transactions.csv \
       --broker ~/Downloads/"Investments Positions.csv" --out adjustments.csv
   ```
4. Paste the rows from `adjustments.csv` at the bottom of Transactions. Old rows are never edited, so every correction stays visible.

## Where the numbers come from

| Data | Source | Refresh |
|---|---|---|
| Prices, index moves | GOOGLEFINANCE | ~20 min |
| Dividend history, ex/pay dates, splits | Alpha Vantage `DIVIDENDS` / `SPLITS` (free: 25 calls/day; the script uses at most 20) | Daily, 6 am New York time |
| Payout ratio | Finnhub `stock/metric` (free) | Daily |
| Holdings, cost | Your Transactions tab, checked against the broker monthly | When you edit |

- **Dividend growth:** split-adjusted, with special dividends excluded.
- **Raise streak:** counts years in which the average annualized dividend rose. That's the "annual dividend went up" convention.
- **Calendar:** shows **Declared** payments as announced. Future payments not yet announced are projected from last year's dates at today's rate and marked **est.**

## Security and controls

These controls are *aligned* with SOC 2 trust criteria and the OWASP Top 10 for a personal tool. They are not a certification.

| Area | Control |
|---|---|
| Secrets (OWASP A02, SOC 2 CC6) | Keys only in Script Properties. Never in the sheet, the code or the log. Log messages are scrubbed of keys. |
| Least privilege (A01, A05) | `spreadsheets.currentonly` scope. URL allowlist of only alphavantage.co and finnhub.io. No redirects followed. |
| Injection (A03, A10 SSRF) | Tickers must match `^[A-Z][A-Z0-9.-]{0,9}$` before going into a URL. API values are validated as numbers or ISO dates before they're written, so nothing can become a formula. reconcile.py neutralises CSV formula injection. |
| Input validation / processing integrity (SOC 2 PI1) | Data-validation rules on every input column. pydantic models reject malformed broker rows (value must equal quantity × price). The monthly reconcile checks the sheet against the broker. |
| Change management (CC8) | Formula columns are protected (warning on edit). Corrections are added as new Adjustment rows; old rows aren't edited. Use Google Sheets version history. |
| Monitoring (CC7, A09) | The Log tab records every script run, warning and error. An hourly check flags missing prices or stale dividend data. The Dashboard shows a Health checks panel. |
| Availability | A daily API budget stays under the free quota. A failing step doesn't stop the others. |
| Components (A06, A08) | Python dependencies pinned in `requirements.txt`. The script loads no external libraries. |
| Account (A07) | Turn on 2-step verification for your Google account. Share the sheet with no one, or view-only. |

The OWASP Top 10 for LLM applications doesn't apply, because nothing in this tracker calls an AI model.

## Running the tests

```
osascript -l JavaScript apps_script/tests/run_tests.js "$PWD"     # Apps Script logic, example data (26 checks)
.venv/bin/python -m unittest discover -s tools -v                 # Python validation and projection math
```

The private test files (real holdings data) run with the same commands where they exist.

---

*Not financial advice. Past dividend growth does not guarantee future dividends or returns.*
