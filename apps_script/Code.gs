/**
 * Portfolio Tracker v2 — Apps Script bound to the tracker spreadsheet.
 *
 * What it does (all on time-driven triggers, see setupTriggers):
 *   daily  : refreshDividendData -> computeDividendStats -> logReinvestedDividends -> snapshotPortfolioValue
 *   hourly : hourlyHealthCheck (flags missing prices / stale data in the Log tab)
 * Prices themselves come from GOOGLEFINANCE formulas (~20 min refresh), not from this script.
 *
 * Security notes
 *   - API keys live in Script Properties (ALPHAVANTAGE_KEY, FINNHUB_KEY), never in code.
 *   - Only https://www.alphavantage.co/ and https://finnhub.io/ may be fetched (appsscript.json allowlist).
 *   - Tickers are validated before they are put in a URL; API values are validated before they
 *     are written to the sheet (numbers / ISO dates only, so nothing can become a formula).
 *   - Log messages are scrubbed of keys.
 *
 * @OnlyCurrentDoc
 */

const CFG = {
  SHEETS: {
    TX: 'Transactions', HOLD: 'Holdings', DIVHIST: 'DivHistory', SPLITS: 'Splits',
    SCHED: 'DivSchedule', GROWTH: 'Dividend Growth', FUND: 'Fundamentals',
    META: 'FetchMeta', PV: 'PV History', LOG: 'Log',
  },
  AV_DAILY_BUDGET: 20,          // Alpha Vantage free tier = 25 requests/day; keep headroom
  DIV_MAX_AGE_H: 20,            // refetch dividend history about once a day
  SPLIT_MAX_AGE_H: 24 * 7,      // splits change rarely
  FUND_MAX_AGE_H: 20,
  AV_SPACING_MS: 2500,
  FH_SPACING_MS: 1100,          // Finnhub free tier = 60 requests/minute
  DRIP_LOOKBACK_DAYS: 60,       // never back-fill older payments automatically
  LOG_MAX_ROWS: 2000,
  RAISE_TOL: 0.0005,            // declared USD amounts are exact; some monthly payers raise ~0.2% at a time
  RAISE_TOL_VARIABLE: 0.05,     // for payers whose USD amount moves with FX (e.g. Canadian companies)
  SPECIAL_RATIO: 1.75,          // payment > 1.75x its neighbours' median = special dividend
  SCHEDULE_HORIZON_DAYS: 400,
  HOLD_ROWS: 40,
  GROWTH_COLS: 17,
};

const TICKER_RE = /^[A-Z][A-Z0-9.\-]{0,9}$/;
const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;

// ───────────────────────────── Pure helpers (unit-tested) ─────────────────────────────

function isValidTicker_(t) {
  return typeof t === 'string' && TICKER_RE.test(t);
}

function isoToUtc_(iso) {
  const p = iso.split('-').map(Number);
  return Date.UTC(p[0], p[1] - 1, p[2]);
}

function utcToIso_(ms) {
  const d = new Date(ms);
  return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0') + '-' +
    String(d.getUTCDate()).padStart(2, '0');
}

function addDays_(iso, n) {
  return utcToIso_(isoToUtc_(iso) + n * 86400000);
}

function addMonths_(iso, n) {
  const p = iso.split('-').map(Number);
  const first = new Date(Date.UTC(p[0], p[1] - 1 + n, 1));
  const y = first.getUTCFullYear(), m = first.getUTCMonth();
  const lastDay = new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
  return utcToIso_(Date.UTC(y, m, Math.min(p[2], lastDay)));
}

function daysBetween_(a, b) {
  return Math.round((isoToUtc_(b) - isoToUtc_(a)) / 86400000);
}

function median_(arr) {
  const s = arr.slice().sort((a, b) => a - b);
  const n = s.length;
  if (!n) return NaN;
  return n % 2 ? s[(n - 1) / 2] : (s[n / 2 - 1] + s[n / 2]) / 2;
}

function apiError_(json) {
  if (!json || typeof json !== 'object') return 'empty response';
  const msg = json.Note || json.Information || json['Error Message'] || json.error;
  return msg ? String(msg).slice(0, 200) : '';
}

/** Alpha Vantage DIVIDENDS -> [{ex, decl, rec, pay, amount}] sorted by ex-date ascending. */
function parseAvDividends_(json, ticker) {
  const err = apiError_(json);
  if (err) throw new Error('Alpha Vantage: ' + err);
  if (!Array.isArray(json.data)) throw new Error('Alpha Vantage: unexpected response shape');
  if (json.symbol && String(json.symbol).toUpperCase() !== ticker) throw new Error('Alpha Vantage: symbol mismatch');
  const d = v => (typeof v === 'string' && DATE_RE.test(v)) ? v : '';
  const out = [];
  json.data.forEach(r => {
    const ex = d(r.ex_dividend_date);
    const amount = Number(r.amount);
    if (!ex || !isFinite(amount) || amount <= 0 || amount > 10000) return;
    out.push({ ex: ex, decl: d(r.declaration_date), rec: d(r.record_date), pay: d(r.payment_date), amount: amount });
  });
  out.sort((a, b) => (a.ex < b.ex ? -1 : a.ex > b.ex ? 1 : 0));
  return out;
}

/** Alpha Vantage SPLITS -> [{date, factor}]. */
function parseAvSplits_(json, ticker) {
  const err = apiError_(json);
  if (err) throw new Error('Alpha Vantage: ' + err);
  if (!Array.isArray(json.data)) throw new Error('Alpha Vantage: unexpected response shape');
  if (json.symbol && String(json.symbol).toUpperCase() !== ticker) throw new Error('Alpha Vantage: symbol mismatch');
  const out = [];
  json.data.forEach(r => {
    const date = (typeof r.effective_date === 'string' && DATE_RE.test(r.effective_date)) ? r.effective_date : '';
    const factor = Number(r.split_factor);
    if (!date || !isFinite(factor) || factor <= 0 || factor > 1000) return;
    out.push({ date: date, factor: factor });
  });
  return out;
}

/** Finnhub /stock/metric -> payout ratio as a fraction, or '' if not available. */
function parseFinnhubPayout_(json) {
  const err = apiError_(json);
  if (err) throw new Error('Finnhub: ' + err);
  const m = json && json.metric;
  if (!m || typeof m !== 'object') return '';
  const raw = (m.payoutRatioTTM !== undefined && m.payoutRatioTTM !== null) ? m.payoutRatioTTM : m.payoutRatioAnnual;
  const n = Number(raw);
  return (raw === null || raw === undefined || !isFinite(n)) ? '' : n / 100;
}

function freqFromGap_(gap) {
  return gap < 45 ? 12 : gap < 135 ? 4 : gap < 270 ? 2 : 1;
}

/**
 * Payments per year in force at each payment (median gap of nearby payments), and the
 * annualized rate. Comparing annualized rates keeps a switch from annual to quarterly
 * payments from looking like a cut or a special dividend.
 */
function annualize_(rows) {
  rows.forEach((r, i) => {
    const gaps = [];
    for (let j = Math.max(1, i - 3); j <= Math.min(rows.length - 1, i + 3); j++) gaps.push(daysBetween_(rows[j - 1].ex, rows[j].ex));
    r.freq = gaps.length ? freqFromGap_(median_(gaps)) : 4;
    r.ann = r.adj * r.freq;
  });
}

/** Flags special (one-off) dividends: annualized amount far above its neighbours' median. */
function markSpecials_(rows) {
  rows.forEach((r, i) => {
    const nb = [];
    for (let j = Math.max(0, i - 4); j <= Math.min(rows.length - 1, i + 4); j++) if (j !== i) nb.push(rows[j].ann);
    r.special = nb.length >= 3 && r.ann > CFG.SPECIAL_RATIO * median_(nb);
  });
}

/**
 * Dividend statistics for one ticker.
 * divs   : output of parseAvDividends_ (raw, NOT split-adjusted)
 * splits : output of parseAvSplits_
 * Returns { stats, rows } where rows are the split-adjusted payments with a `special` flag.
 */
function computeTickerStats_(ticker, divs, splits, todayIso, splitsKnown) {
  const rows = divs.map(r => {
    let f = 1;
    splits.forEach(s => { if (s.date > r.ex) f *= s.factor; });
    return { ex: r.ex, decl: r.decl, rec: r.rec, pay: r.pay, amount: r.amount, factor: f, adj: r.amount / f };
  });
  // A payment declared just before a split is sometimes reported on the post-split basis already
  // (seen in real data). Keep whichever reading matches the first payment after the split.
  splits.forEach(s => {
    const next = rows.find(r => r.ex >= s.date);
    if (!next) return;
    rows.forEach(r => {
      if (r.ex >= s.date || daysBetween_(r.ex, s.date) > 60) return;
      if (Math.abs(Math.log(r.adj * s.factor / next.adj)) < Math.abs(Math.log(r.adj / next.adj))) {
        r.factor /= s.factor;
        r.adj *= s.factor;
      }
    });
  });
  annualize_(rows);
  markSpecials_(rows);
  const reg = rows.filter(r => !r.special);

  const st = {
    ticker: ticker, frequency: '', perYear: '', current: '', fwdAnnual: '', lastRaiseDate: '', lastRaisePct: '',
    nextRaiseEst: '', cagr1: '', cagr3: '', cagr5: '', cagr10: '', streak: '', streakNote: '', status: '',
    specials: rows.length - reg.length, asOf: todayIso,
  };
  const notes = [];
  if (!splitsKnown) notes.push('split history pending');
  if (reg.length < 2) {
    st.status = ['Not enough dividend history'].concat(notes).join('; ');
    return { stats: st, rows: rows };
  }

  // Payment frequency from the median gap between recent regular payments.
  const recent = reg.slice(-9);
  const gaps = [];
  for (let i = 1; i < recent.length; i++) gaps.push(daysBetween_(recent[i - 1].ex, recent[i].ex));
  const gap = median_(gaps);
  const perYear = freqFromGap_(gap);
  st.perYear = perYear;
  st.frequency = { 12: 'Monthly', 4: 'Quarterly', 2: 'Semi-annual', 1: 'Annual' }[perYear];

  const last = reg[reg.length - 1];
  st.current = last.adj;
  st.fwdAnnual = last.adj * perYear;

  // Amounts that wobble up and down (foreign-currency payers converted to USD) need a wider tolerance.
  const tail = reg.slice(-12);
  let ups = 0, downs = 0;
  for (let i = 1; i < tail.length; i++) {
    const ch = tail[i].ann / tail[i - 1].ann - 1;
    if (ch > 0.0005) ups++; else if (ch < -0.0005) downs++;
  }
  const variable = ups >= 2 && downs >= 2;
  const tol = variable ? CFG.RAISE_TOL_VARIABLE : CFG.RAISE_TOL;
  if (variable) notes.push('USD amount varies (currency) — growth approximate');

  // Growth: current annualized regular rate vs the rate N years earlier.
  const tolDays = Math.round(gap / 2);
  const rateAt = refIso => {
    const lim = addDays_(refIso, tolDays);
    if (reg[0].ex > lim) return null;               // history does not reach back that far
    let found = null;
    for (let i = 0; i < reg.length && reg[i].ex <= lim; i++) found = reg[i];
    if (!found || daysBetween_(found.ex, refIso) > 2 * gap + tolDays) return null;  // hole in history
    return found;
  };
  [[1, 'cagr1'], [3, 'cagr3'], [5, 'cagr5'], [10, 'cagr10']].forEach(p => {
    const then = rateAt(addMonths_(last.ex, -12 * p[0]));
    if (then && then.ann > 0) st[p[1]] = Math.pow(last.ann / then.ann, 1 / p[0]) - 1;
  });

  // Raise / cut events between consecutive regular payments.
  const raises = [];
  let cutRecent = false;
  const yearAgo = addMonths_(todayIso, -12);
  for (let i = 1; i < reg.length; i++) {
    const ch = reg[i].ann / reg[i - 1].ann - 1;
    const when = reg[i].decl || reg[i].ex;
    if (ch > tol) raises.push({ date: when, pct: ch });
    else if (ch < -tol && reg[i].ex >= yearAgo) cutRecent = true;
  }
  if (raises.length) {
    const lr = raises[raises.length - 1];
    st.lastRaiseDate = lr.date;
    st.lastRaisePct = lr.pct;
    const recentRaises = raises.filter(r => r.date >= addMonths_(todayIso, -36));
    if (recentRaises.length >= 2) {
      const months = [];
      for (let i = 1; i < recentRaises.length; i++) {
        months.push(Math.max(1, Math.round(daysBetween_(recentRaises[i - 1].date, recentRaises[i].date) / 30.44)));
      }
      const step = Math.max(1, Math.round(median_(months)));
      let est = addMonths_(lr.date, step);
      while (est < todayIso) est = addMonths_(est, step);
      st.nextRaiseEst = est.slice(0, 7);
    }
  }

  // Consecutive years of increases: a year counts when its average annualized regular rate beat
  // the previous year's (the "annual dividend went up" convention, immune to payment timing).
  const sum = {}, cnt = {};
  reg.forEach(r => { const yy = Number(r.ex.slice(0, 4)); sum[yy] = (sum[yy] || 0) + r.ann; cnt[yy] = (cnt[yy] || 0) + 1; });
  const avg = yy => (cnt[yy] ? sum[yy] / cnt[yy] : undefined);
  const curY = Number(todayIso.slice(0, 4));
  const firstY = Number(reg[0].ex.slice(0, 4));
  const raisedIn = yy => avg(yy) !== undefined && avg(yy - 1) !== undefined && avg(yy) > avg(yy - 1) * (1 + tol);
  let y = raisedIn(curY) ? curY : curY - 1;
  let streak = 0;
  while (y > firstY && raisedIn(y)) { streak++; y--; }
  st.streak = streak;
  const sn = [];
  if (streak > 0 && y === firstY) sn.push('≥' + streak + ' (data starts ' + firstY + ')');
  if (!raises.some(r => Number(r.date.slice(0, 4)) === curY)) sn.push('no ' + curY + ' raise yet');
  st.streakNote = sn.join('; ');

  const status = [];
  if (cutRecent) status.push('CUT in last 12 months');
  if (!st.lastRaiseDate || st.lastRaiseDate < addMonths_(todayIso, -15)) status.push('No raise in 15+ months');
  st.status = (status.length ? status : ['OK']).concat(notes).join('; ');
  return { stats: st, rows: rows };
}

/**
 * Upcoming payments for one ticker over the next ~13 months.
 * Declared payments come straight from the data; the rest are projected from the last
 * year's pattern at the current rate and marked 'est.'.
 */
function buildSchedule_(ticker, rows, stats, todayIso) {
  const out = [];
  if (!stats.perYear) return out;
  const horizon = addDays_(todayIso, CFG.SCHEDULE_HORIZON_DAYS);
  const lags = rows.filter(r => r.pay && !r.special).slice(-8).map(r => daysBetween_(r.ex, r.pay));
  const lag = lags.length ? Math.round(median_(lags)) : 14;

  rows.forEach(r => {
    const pay = r.pay || addDays_(r.ex, lag);
    if (pay >= todayIso && r.ex <= horizon) {
      out.push({ ticker: ticker, ex: r.ex, pay: pay, amount: r.amount, status: r.special ? 'Special (declared)' : (r.pay ? 'Declared' : 'Declared (pay date est.)') });
    }
  });

  const reg = rows.filter(r => !r.special);
  const lastEx = reg[reg.length - 1].ex;
  const base = reg.slice(-stats.perYear);
  const gapDays = Math.round(365 / stats.perYear);
  for (let k = 1; k <= 3; k++) {
    base.forEach(b => {
      const ex = addMonths_(b.ex, 12 * k);
      if (ex <= addDays_(lastEx, Math.round(gapDays / 2)) || ex > horizon) return;
      const pay = b.pay ? addMonths_(b.pay, 12 * k) : addDays_(ex, lag);
      if (pay < todayIso) return;
      out.push({ ticker: ticker, ex: ex, pay: pay, amount: stats.current, status: 'est.' });
    });
  }
  out.sort((a, b) => (a.pay < b.pay ? -1 : a.pay > b.pay ? 1 : 0));
  return out;
}

/**
 * Dividend Reinvest rows that should exist but don't yet.
 * txns: [{date, ticker, type, action, shares (signed delta), source, key}]
 * divsByTicker: {TICKER: [{ex, pay, amount}]}
 * Rule: shares held on the ex-date = all rows dated before the ex-date, plus broker-seed rows
 * (the opening balance already holds the shares). Payments paid on or before the seed date are
 * already inside the broker numbers, so they are skipped.
 */
function computeDripRows_(txns, divsByTicker, todayIso, lookbackDays) {
  const keys = {};
  txns.forEach(t => { if (t.key) keys[t.key] = true; });
  const tickers = {};
  txns.forEach(t => { if (t.type === 'Stock') tickers[t.ticker] = true; });
  const out = [];
  const earliestPay = addDays_(todayIso, -lookbackDays);

  Object.keys(tickers).sort().forEach(tk => {
    const own = txns.filter(t => t.ticker === tk);
    const seedDates = own.filter(t => t.source === 'broker-seed').map(t => t.date).sort();
    const firstDate = own.map(t => t.date).sort()[0];
    const after = seedDates.length ? seedDates[seedDates.length - 1] : firstDate;
    (divsByTicker[tk] || []).forEach(d => {
      if (!d.pay || d.pay > todayIso || d.pay < earliestPay || d.pay <= after) return;
      const key = tk + '|' + d.pay;
      if (keys[key]) return;
      let shares = 0;
      own.forEach(t => { if (t.source === 'broker-seed' || t.date < d.ex) shares += t.shares; });
      if (!(shares > 0)) return;
      const amount = Math.round(shares * d.amount * 100) / 100;
      if (!(amount > 0)) return;
      keys[key] = true;
      out.push({
        date: d.pay, ticker: tk, type: 'Stock', action: 'Dividend Reinvest', amount: amount, source: 'auto-drip', key: key,
        note: '$' + d.amount + '/sh x ' + shares.toFixed(5) + ' sh (ex ' + d.ex + '); shares estimated from close',
      });
    });
  });
  return out;
}

function scrub_(msg, secrets) {
  let s = String(msg);
  (secrets || []).forEach(k => { if (k) s = s.split(k).join('***'); });
  return s.replace(/(apikey|token)=[^&\s"]+/gi, '$1=***').slice(0, 500);
}

// ───────────────────────────── Spreadsheet glue ─────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Portfolio')
    .addItem('Refresh dividend data now', 'refreshDividendData')
    .addItem('Recompute dividend stats & calendar', 'computeDividendStats')
    .addItem('Log reinvested dividends', 'logReinvestedDividends')
    .addItem('Snapshot portfolio value', 'snapshotPortfolioValue')
    .addItem('Rebuild dashboard', 'buildDashboard')
    .addSeparator()
    .addItem('Set up automatic triggers (run once)', 'setupTriggers')
    .addToUi();
}

function setupTriggers() {
  const handlers = ['dailyJob', 'hourlyHealthCheck'];
  ScriptApp.getProjectTriggers().forEach(t => {
    if (handlers.indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyJob').timeBased().everyDays(1).atHour(6).create();
  ScriptApp.newTrigger('hourlyHealthCheck').timeBased().everyHours(1).create();
  protectFormulaColumns_();
  buildDashboard();
  log_('setupTriggers', 'OK', 'daily job at 06:00 + hourly health check; formula columns protected (warning only)');
}

function dailyJob() {
  // Each step runs even if an earlier one fails (e.g. the dividend API is rate-limited).
  const failed = [];
  [refreshDividendData, computeDividendStats, logReinvestedDividends, snapshotPortfolioValue].forEach(step => {
    try { step(); } catch (e) { failed.push(step.name); }
  });
  log_('dailyJob', failed.length ? 'WARN' : 'OK', failed.length ? 'failed steps: ' + failed.join(', ') : 'finished');
}

function refreshDividendData() {
  withLock_('refreshDividendData', () => {
    const avKey = getSecret_('ALPHAVANTAGE_KEY');
    const fhKey = PropertiesService.getScriptProperties().getProperty('FINNHUB_KEY');
    const tickers = stockTickers_();
    const meta = readMeta_();
    const now = new Date();
    const ageH = (t, kind) => {
      const v = meta[t + '|' + kind];
      return v ? (now - new Date(v)) / 3.6e6 : Infinity;
    };

    const queue = [];
    tickers.forEach(t => {
      if (ageH(t, 'div') > CFG.DIV_MAX_AGE_H) queue.push({ t: t, kind: 'div', age: ageH(t, 'div'), pri: 0 });
      if (ageH(t, 'split') > CFG.SPLIT_MAX_AGE_H) queue.push({ t: t, kind: 'split', age: ageH(t, 'split'), pri: 1 });
    });
    queue.sort((a, b) => (b.age - a.age) || (a.pri - b.pri));

    let budget = avBudgetLeft_();
    let done = 0, failed = 0;
    for (let i = 0; i < queue.length && budget > 0; i++) {
      const q = queue[i];
      const fn = q.kind === 'div' ? 'DIVIDENDS' : 'SPLITS';
      const url = 'https://www.alphavantage.co/query?function=' + fn + '&symbol=' + encodeURIComponent(q.t) + '&apikey=' + encodeURIComponent(avKey);
      budget = avSpend_();
      try {
        const json = fetchJson_(url, [avKey]);
        if (q.kind === 'div') writeTickerRows_(CFG.SHEETS.DIVHIST, q.t, parseAvDividends_(json, q.t).map(r => [q.t, r.ex, r.decl, r.rec, r.pay, r.amount]));
        else writeTickerRows_(CFG.SHEETS.SPLITS, q.t, parseAvSplits_(json, q.t).map(r => [q.t, r.date, r.factor]));
        meta[q.t + '|' + q.kind] = now.toISOString();
        done++;
      } catch (e) {
        failed++;
        log_('refreshDividendData', 'WARN', q.t + ' ' + q.kind + ': ' + scrub_(e.message, [avKey]));
        if (/rate limit|requests per day|premium|Thank you for using/i.test(e.message)) break;
      }
      Utilities.sleep(CFG.AV_SPACING_MS);
    }

    let fund = 0;
    if (fhKey) {
      tickers.forEach(t => {
        if (ageH(t, 'fund') <= CFG.FUND_MAX_AGE_H) return;
        const url = 'https://finnhub.io/api/v1/stock/metric?symbol=' + encodeURIComponent(t) + '&metric=all&token=' + encodeURIComponent(fhKey);
        try {
          const payout = parseFinnhubPayout_(fetchJson_(url, [fhKey]));
          writeTickerRows_(CFG.SHEETS.FUND, t, [[t, payout]]);
          meta[t + '|fund'] = now.toISOString();
          fund++;
        } catch (e) {
          log_('refreshDividendData', 'WARN', t + ' payout: ' + scrub_(e.message, [fhKey]));
        }
        Utilities.sleep(CFG.FH_SPACING_MS);
      });
    }
    writeMeta_(meta);
    const pending = queue.length - done - failed;
    log_('refreshDividendData', failed ? 'WARN' : 'OK',
      done + ' Alpha Vantage updates, ' + failed + ' failed, ' + pending + ' waiting for tomorrow\'s quota; ' + fund + ' payout ratios');
  });
}

function computeDividendStats() {
  withLock_('computeDividendStats', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const today = todayIso_();
    const divs = groupByTicker_(readRows_(CFG.SHEETS.DIVHIST, 6), r => ({
      ex: isoCell_(r[1]), decl: isoCell_(r[2]), rec: isoCell_(r[3]), pay: isoCell_(r[4]), amount: Number(r[5]),
    }));
    const splits = groupByTicker_(readRows_(CFG.SHEETS.SPLITS, 3), r => ({ date: isoCell_(r[1]), factor: Number(r[2]) }));
    const meta = readMeta_();

    const statRows = [], schedRows = [];
    stockTickers_().forEach(t => {
      const d = (divs[t] || []).filter(x => x.ex && x.amount > 0).sort((a, b) => (a.ex < b.ex ? -1 : 1));
      if (!d.length) {
        const blank = new Array(CFG.GROWTH_COLS).fill('');
        blank[0] = t; blank[14] = 'No dividend data yet'; blank[16] = today;
        statRows.push(blank);
        return;
      }
      const res = computeTickerStats_(t, d, splits[t] || [], today, !!meta[t + '|split']);
      const s = res.stats;
      statRows.push([s.ticker, s.frequency, s.perYear, s.current, s.fwdAnnual, toDate_(s.lastRaiseDate), s.lastRaisePct,
        s.nextRaiseEst, s.cagr1, s.cagr3, s.cagr5, s.cagr10, s.streak, s.streakNote, s.status, s.specials, toDate_(s.asOf)]);
      buildSchedule_(t, res.rows, s, today).forEach(x => schedRows.push([x.ticker, toDate_(x.ex), toDate_(x.pay), x.amount, x.status]));
    });

    const g = ss.getSheetByName(CFG.SHEETS.GROWTH);
    g.getRange(2, 1, 60, CFG.GROWTH_COLS).clearContent();
    if (statRows.length) g.getRange(2, 1, statRows.length, CFG.GROWTH_COLS).setValues(statRows);

    const sc = ss.getSheetByName(CFG.SHEETS.SCHED);
    sc.getRange(2, 1, Math.max(sc.getMaxRows() - 1, 1), 5).clearContent();
    if (schedRows.length) {
      if (sc.getMaxRows() < schedRows.length + 1) sc.insertRowsAfter(sc.getMaxRows(), schedRows.length + 1 - sc.getMaxRows());
      sc.getRange(2, 1, schedRows.length, 5).setValues(schedRows);
    }
    log_('computeDividendStats', 'OK', statRows.length + ' tickers, ' + schedRows.length + ' scheduled payments');
  });
}

function logReinvestedDividends() {
  withLock_('logReinvestedDividends', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tx = ss.getSheetByName(CFG.SHEETS.TX);
    const last = tx.getLastRow();
    const vals = last >= 2 ? tx.getRange(2, 1, last - 1, 13).getValues() : [];
    const txns = [];
    let lastFilled = 1;
    vals.forEach((r, i) => {
      if (r[0] === '' || r[1] === '') return;
      lastFilled = i + 2;
      const shares = Number(r[12]);   // column M = signed shares delta
      txns.push({
        date: isoCell_(r[0]), ticker: String(r[1]).trim().toUpperCase(), type: String(r[2]), action: String(r[3]),
        shares: isFinite(shares) ? shares : 0, source: String(r[8]), key: String(r[9]),
      });
    });
    const divs = groupByTicker_(readRows_(CFG.SHEETS.DIVHIST, 6), r => ({ ex: isoCell_(r[1]), pay: isoCell_(r[4]), amount: Number(r[5]) }));
    const add = computeDripRows_(txns, divs, todayIso_(), CFG.DRIP_LOOKBACK_DAYS);
    if (!add.length) { log_('logReinvestedDividends', 'OK', 'nothing new'); return; }

    let row = lastFilled + 1;   // always append after the last real row, never into a gap
    add.forEach(a => {
      // Only input columns are written; shares (F) stay blank so column G estimates them from the pay-date close.
      tx.getRange(row, 1, 1, 5).setValues([[toDate_(a.date), a.ticker, a.type, a.action, a.amount]]);
      tx.getRange(row, 9, 1, 3).setValues([[a.source, a.key, a.note]]);
      ensureRowFormulas_(tx, row);
      row++;
    });
    log_('logReinvestedDividends', 'OK', 'added ' + add.length + ': ' + add.map(a => a.key + ' $' + a.amount).join(', '));
  });
}

function snapshotPortfolioValue() {
  withLock_('snapshotPortfolioValue', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(CFG.SHEETS.PV);
    const today = todayIso_();
    const lastRow = sh.getLastRow();
    if (lastRow >= 2 && isoCell_(sh.getRange(lastRow, 1).getValue()) === today) return;
    const v = name => { const r = ss.getRangeByName(name); return r ? r.getValue() : ''; };
    const total = v('TotalValue');
    if (typeof total !== 'number' || !isFinite(total)) {
      log_('snapshotPortfolioValue', 'WARN', 'TotalValue is not a number (price error?) — skipped');
      return;
    }
    sh.appendRow([toDate_(today), total, v('MoneyIn'), v('CostBasis'), v('FwdIncome'), 'daily snapshot']);
  });
}

function hourlyHealthCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const h = ss.getSheetByName(CFG.SHEETS.HOLD).getRange(2, 1, CFG.HOLD_ROWS, 8).getDisplayValues();
  const bad = h.filter(r => r[0] && /^#/.test(r[7])).map(r => r[0]);
  if (bad.length) log_('hourlyHealthCheck', 'WARN', 'no live price for: ' + bad.join(', '));
  const meta = readMeta_();
  const stale = stockTickers_().filter(t => !meta[t + '|div'] || (new Date() - new Date(meta[t + '|div'])) > 3 * 864e5);
  if (stale.length) log_('hourlyHealthCheck', 'WARN', 'dividend data older than 3 days: ' + stale.join(', '));
}

// ───────────────────────────── Internals ─────────────────────────────

function withLock_(name, fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { log_(name, 'WARN', 'another run in progress — skipped'); return; }
  try {
    fn();
  } catch (e) {
    log_(name, 'ERROR', scrub_(e && e.message ? e.message : e, secrets_()));
    throw e;
  } finally {
    lock.releaseLock();
  }
}

function secrets_() {
  const p = PropertiesService.getScriptProperties();
  return [p.getProperty('ALPHAVANTAGE_KEY'), p.getProperty('FINNHUB_KEY')];
}

function getSecret_(name) {
  const v = PropertiesService.getScriptProperties().getProperty(name);
  if (!v) throw new Error('Missing Script Property ' + name + ' (Apps Script → Project Settings → Script Properties)');
  return v;
}

function fetchJson_(url, secrets) {
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: false });
  const code = res.getResponseCode();
  const text = res.getContentText();
  if (code !== 200) throw new Error('HTTP ' + code + ': ' + scrub_(text.slice(0, 200), secrets));
  if (text.length > 2e6) throw new Error('response too large');
  try { return JSON.parse(text); } catch (e) { throw new Error('invalid JSON'); }
}

/** Daily request counter so the free Alpha Vantage quota is never exceeded. */
function avBudgetLeft_() {
  const u = JSON.parse(PropertiesService.getScriptProperties().getProperty('AV_USAGE') || '{}');
  const day = new Date().toISOString().slice(0, 10);
  return CFG.AV_DAILY_BUDGET - (u.day === day ? u.count : 0);
}

function avSpend_() {
  const p = PropertiesService.getScriptProperties();
  const u = JSON.parse(p.getProperty('AV_USAGE') || '{}');
  const day = new Date().toISOString().slice(0, 10);
  const count = (u.day === day ? u.count : 0) + 1;
  p.setProperty('AV_USAGE', JSON.stringify({ day: day, count: count }));
  return CFG.AV_DAILY_BUDGET - count;
}

function stockTickers_() {
  const v = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEETS.HOLD).getRange(2, 1, CFG.HOLD_ROWS, 2).getValues();
  const out = [];
  v.forEach(r => {
    const t = String(r[0]).trim().toUpperCase();
    if (r[1] === 'Stock' && isValidTicker_(t) && out.indexOf(t) < 0) out.push(t);
  });
  return out;
}

function readRows_(sheetName, cols) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, cols).getValues().filter(r => r[0] !== '') : [];
}

function groupByTicker_(rows, mapFn) {
  const out = {};
  rows.forEach(r => {
    const t = String(r[0]).trim().toUpperCase();
    (out[t] = out[t] || []).push(mapFn(r));
  });
  return out;
}

/** Replaces all rows of one ticker in a ticker-keyed data tab. */
function writeTickerRows_(sheetName, ticker, newRows) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const width = sh.getLastColumn();
  const n = sh.getLastRow() - 1;
  const keep = n > 0 ? sh.getRange(2, 1, n, width).getValues().filter(r => r[0] !== '' && String(r[0]).toUpperCase() !== ticker) : [];
  const fetched = new Date().toISOString();
  const rows = keep.concat(newRows.map(r => r.concat([fetched]).slice(0, width)));
  rows.sort((a, b) => (String(a[0]) + String(a[1]) < String(b[0]) + String(b[1]) ? -1 : 1));
  if (n > 0) sh.getRange(2, 1, n, width).clearContent();
  if (rows.length) {
    if (sh.getMaxRows() < rows.length + 1) sh.insertRowsAfter(sh.getMaxRows(), rows.length + 1 - sh.getMaxRows());
    sh.getRange(2, 1, rows.length, width).setValues(rows.map(r => { while (r.length < width) r.push(''); return r; }));
  }
}

function readMeta_() {
  const meta = {};
  readRows_(CFG.SHEETS.META, 3).forEach(r => { meta[String(r[0]) + '|' + String(r[1])] = String(r[2]); });
  return meta;
}

function writeMeta_(meta) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEETS.META);
  const rows = Object.keys(meta).sort().map(k => { const p = k.split('|'); return [p[0], p[1], meta[k]]; });
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 3).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 3).setValues(rows);
}

/** Copies the formula columns of row 2 into `row` if they are missing (rows beyond the pre-filled range). */
function ensureRowFormulas_(tx, row) {
  [7, 8, 12, 13, 14, 15, 16].forEach(c => {
    const cell = tx.getRange(row, c);
    if (!cell.getFormula()) tx.getRange(2, c).copyTo(cell);
  });
}

function protectFormulaColumns_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tx = ss.getSheetByName(CFG.SHEETS.TX);
  const ranges = [tx.getRange('G:H'), tx.getRange('L:P'), ss.getSheetByName(CFG.SHEETS.HOLD).getRange('A:U')];
  const desc = 'Formula columns — calculated automatically';
  ss.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => { if (p.getDescription() === desc) p.remove(); });
  ranges.forEach(r => r.protect().setDescription(desc).setWarningOnly(true));
}

function todayIso_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
}

function isoCell_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  }
  const s = String(v || '').trim();
  return DATE_RE.test(s) ? s : '';
}

function toDate_(iso) {
  if (!iso || !DATE_RE.test(iso)) return '';
  const p = iso.split('-').map(Number);
  return new Date(p[0], p[1] - 1, p[2]);
}

function log_(fn, status, msg) {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEETS.LOG);
    sh.appendRow([new Date(), fn, status, scrub_(msg, secrets_())]);
    const extra = sh.getLastRow() - 1 - CFG.LOG_MAX_ROWS;
    if (extra > 0) sh.deleteRows(2, extra);
  } catch (e) {
    console.error('log failed: ' + e);
  }
}

// ───────────────────────────── Dashboard (layout + charts) ─────────────────────────────
// Built by script because Google Sheets map (geo) charts cannot be imported from .xlsx.
// Portfolio → Rebuild dashboard re-creates it at any time; it only reads other tabs.

const DASH = {
  ORANGE: '#ff8c00', GREEN_HDR: '#d9ead3', YELLOW: '#ffd966', PEACH: '#fce5cd', RED_HDR: '#e06666',
  PINK: '#f4cccc', BLUE: '#a4c2f4', LIGHT_BLUE: '#c9daf8', HOLD_GREEN: '#b6d7a8', OUTLOOK: '#f9cb9c',
  OUTLOOK_HDR: '#ff9900', MONEY: '"$"#,##0.00', PCT: '0.00%', SIGNED_PCT: '+0.00%;-0.00%;0.00%',
};

const TX_RANGE = 'Transactions!$E$2:$E$1001', TX_ACTION = 'Transactions!$D$2:$D$1001';

/** Hidden "Calc" tab: every Dashboard number, each one a named formula. */
function calcDefinitions_() {
  const g = "'Dividend Growth'!", v = g + 'V2:V61';
  const weighted = col => '=IFERROR(SUMPRODUCT(FILTER(' + col + ',ISNUMBER(' + col + '),ISNUMBER(' + v + ')),FILTER(' + v +
    ',ISNUMBER(' + col + '),ISNUMBER(' + v + ')))/SUM(FILTER(' + v + ',ISNUMBER(' + col + '),ISNUMBER(' + v + '))),"")';
  const sumAction = a => 'SUMIFS(' + TX_RANGE + ',' + TX_ACTION + ',"' + a + '")';
  const byType = (col, type) => '=ARRAYFORMULA(SUM(IFERROR(IF(Holdings!B2:B41="' + type + '",VALUE(Holdings!' + col + '2:' + col + '41),0),0)))';
  const sorted = 'SORT(FILTER({E2:E30,H2:H30},E2:E30<>""),2,';
  return [
    ['StocksValue', byType('I', 'Stock')],
    ['CryptoValue', byType('I', 'Crypto')],
    ['TotalValue', '=StocksValue+CryptoValue+CashBalance'],
    ['StocksCost', byType('F', 'Stock')],
    ['CryptoCost', byType('F', 'Crypto')],
    ['CostBasis', '=SUM(Holdings!F2:F41)'],
    ['TotalBuys', '=' + sumAction('Buy')],
    ['MoneyIn', '=' + sumAction('Buy') + '-' + sumAction('Sell')],
    ['FwdIncome', '=SUM(Holdings!O2:O41)'],
    ['StartYield', '=IFERROR(FwdIncome/TotalValue,0)'],
    ['StockYield', '=IFERROR(FwdIncome/StocksValue,0)'],
    ['YieldOnCost', '=IFERROR(FwdIncome/StocksCost,0)'],
    ['DivReceived', '=SUMIFS(' + TX_RANGE + ',' + TX_ACTION + ',"Dividend*")'],
    ['DivCash', '=' + sumAction('Dividend Cash')],
    ['RealizedGain', '=SUM(Holdings!S2:S41)'],
    ['PriceGain', '=StocksValue+CryptoValue-CostBasis'],
    ['TotalReturn', '=StocksValue+CryptoValue+DivCash-MoneyIn'],
    ['XirrReturn', '=IF(COUNTIFS(Transactions!$I$2:$I$1001,"broker-seed",' + TX_ACTION + ',"Buy")>0,"needs buy history",' +
      'IFERROR(XIRR({FILTER(Transactions!P2:P1001,ISNUMBER(Transactions!P2:P1001),Transactions!P2:P1001<>0);StocksValue+CryptoValue},' +
      '{FILTER(Transactions!A2:A1001,ISNUMBER(Transactions!P2:P1001),Transactions!P2:P1001<>0);TODAY()}),"not enough history"))'],
    ['Positions', '=COUNTIFS(Holdings!A2:A41,"?*",Holdings!E2:E41,">0")'],
    ['PortfolioDivCAGR5', weighted(g + 'K2:K61')],
    ['PortfolioDivCAGR1', weighted(g + 'I2:I61')],
    ['UnderSector', '=IFERROR(INDEX(' + sorted + 'TRUE),1,1),"")'],
    ['UnderDiff', '=IFERROR(INDEX(' + sorted + 'TRUE),1,2),"")'],
    ['OverSector', '=IFERROR(INDEX(' + sorted + 'FALSE),1,1),"")'],
    ['OverDiff', '=IFERROR(INDEX(' + sorted + 'FALSE),1,2),"")'],
    ['PricesMissing', '=SUMPRODUCT(--ISERROR(Holdings!H2:H41))'],
    ['LastDailyRun', '=IFERROR(TEXT(INDEX(SORT(FILTER(Log!A2:A,Log!B2:B="dailyJob"),1,FALSE),1),"yyyy-mm-dd hh:mm"),"never")'],
  ];
}

function buildCalc_(ss) {
  ensureSectorTarget_(ss);
  let calc = ss.getSheetByName('Calc');
  if (!calc) calc = ss.insertSheet('Calc');
  calc.clear();
  const defs = calcDefinitions_();
  // Point the names at Calc first (re-pointing names that used to live on the old Dashboard), then write formulas.
  const existing = {};
  ss.getNamedRanges().forEach(n => { existing[n.getName()] = n; });
  defs.forEach((d, i) => {
    const cell = calc.getRange(i + 2, 2);
    if (existing[d[0]]) existing[d[0]].setRange(cell); else ss.setNamedRange(d[0], cell);
  });
  calc.getRange(1, 1, 1, 2).setValues([['Name', 'Value']]);
  calc.getRange(2, 1, defs.length, 1).setValues(defs.map(d => [d[0]]));
  calc.getRange(2, 2, defs.length, 1).setFormulas(defs.map(d => [d[1]]));

  // Chart source tables (stocks only).
  const held = 'Holdings!B2:B41="Stock",ISNUMBER(Holdings!E2:E41),Holdings!E2:E41>0';
  calc.getRange('E1:H1').setValues([['Sector', 'Value', 'Weight', 'vs target']]);
  // Each table is one spilled formula, so rows below the data stay truly empty (charts need that).
  // (SUMIFS does not expand inside ARRAYFORMULA in Google Sheets, so totals are grouped with QUERY over a FILTER.)
  calc.getRange('E2').setFormula('=IFERROR(QUERY(FILTER({Holdings!D2:D41,Holdings!I2:I41},' + held + '),' +
    '"select Col1, sum(Col2) group by Col1 order by sum(Col2) desc label sum(Col2) \'\'",0),"")');
  calc.getRange('G2').setFormula('=ARRAYFORMULA(IF(E2:E30="","",IFERROR(F2:F30/StocksValue,0)))');
  calc.getRange('H2').setFormula('=ARRAYFORMULA(IF(E2:E30="","",G2:G30-SectorTarget))');
  calc.getRange('J1:L1').setValues([['Ticker', 'Total Equity', 'Total Cost']]);
  calc.getRange('J2').setFormula('=IFERROR(FILTER({Holdings!A2:A41,Holdings!I2:I41,Holdings!F2:F41},' + held + '),"")');
  calc.getRange('N1:O1').setValues([['Country', 'Value']]);
  calc.getRange('N2').setFormula('=IFERROR(QUERY(FILTER({Holdings!U2:U41,Holdings!I2:I41},' + held + '),' +
    '"select Col1, sum(Col2) group by Col1 label sum(Col2) \'\'",0),"")');
  calc.hideSheet();
  return calc;
}

function ensureSectorTarget_(ss) {
  if (ss.getRangeByName('SectorTarget')) return;
  const st = ss.getSheetByName('Settings');
  st.getRange('A19:D19').setValues([['Target weight per sector', 0.1, '', 'Used for the Dashboard over/underweight boxes (your old sheet used 10%).']]);
  st.getRange('B19').setNumberFormat('0.00%').setBackground('#fff7d6');
  ss.setNamedRange('SectorTarget', st.getRange('B19'));
}

function style_(range, o) {
  if (o.merge) range.merge();
  if (o.bg) range.setBackground(o.bg);
  if (o.color) range.setFontColor(o.color);
  if (o.size) range.setFontSize(o.size);
  if (o.bold) range.setFontWeight('bold');
  if (o.nf) range.setNumberFormat(o.nf);
  if (o.wrap) range.setWrap(true);
  range.setHorizontalAlignment(o.align || 'center').setVerticalAlignment('middle');
  if (o.border) range.setBorder(true, true, true, true, o.inner || false, o.inner || false, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  return range;
}

function buildDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  buildCalc_(ss);
  // Repair: open-ended ranges like A2:E do not survive an .xlsx import, so write the calendar formula natively.
  ss.getSheetByName('Dividend Calendar').getRange('A4').setFormula(
    '=IFERROR(SORT(FILTER(DivSchedule!A2:E2000,DivSchedule!C2:C2000>=TODAY(),DivSchedule!C2:C2000<=EDATE(TODAY(),12)),3,TRUE),' +
    '"No dividend data yet — run Portfolio → Refresh dividend data now")');
  let sh = ss.getSheetByName('Dashboard');
  if (!sh) sh = ss.insertSheet('Dashboard', 0);
  sh.getCharts().forEach(c => sh.removeChart(c));
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart().clearDataValidations();
  sh.clear();
  sh.setConditionalFormatRules([]);
  if (sh.getMaxColumns() < 17) sh.insertColumnsAfter(sh.getMaxColumns(), 17 - sh.getMaxColumns());
  if (sh.getMaxRows() < 70) sh.insertRowsAfter(sh.getMaxRows(), 70 - sh.getMaxRows());
  sh.setHiddenGridlines(true);
  [20, 175, 95, 20, 125, 115, 60, 60, 60, 60, 60, 60, 20, 150, 150, 150].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  const R = a1 => sh.getRange(a1);

  // Title
  style_(R('B1:P2'), { merge: true, bg: DASH.ORANGE, size: 40, bold: true }).setValue('Investment Dashboard');
  sh.setRowHeights(1, 2, 38);

  // KPI boxes
  [['B4:C4', 'B5:C5', 'Expected Yearly Dividend Income', '=FwdIncome', DASH.MONEY],
   ['E4:G4', 'E5:G5', 'All Time Dividend Income', '=DivReceived', DASH.MONEY],
   ['I4:L4', 'I5:L5', 'Average Monthly Dividend', '=FwdIncome/12', DASH.MONEY],
   ['N4:P4', 'N5:P5', 'Portfolio Dividend Yield', '=StockYield', DASH.PCT]].forEach(k => {
    style_(R(k[0]), { merge: true, bg: DASH.GREEN_HDR, bold: true, size: 11, wrap: true, border: true }).setValue(k[2]);
    style_(R(k[1]), { merge: true, bg: '#ffffff', bold: true, size: 22, nf: k[4], border: true }).setFormula(k[3]);
  });
  sh.setRowHeight(4, 36);
  sh.setRowHeight(5, 58);

  // Left column: over/underweight, Stocks, Crypto, Cash, Total holdings
  const block = (row, title, bg, lines) => {
    style_(R('B' + row + ':C' + row), { merge: true, bg: bg, bold: true, size: lines ? 15 : 11, border: true, wrap: true }).setValue(title);
    (lines || []).forEach((l, i) => {
      const r = row + 1 + i;
      style_(R('B' + r), { bg: l[3] || DASH.PINK, bold: true, align: 'left', border: true }).setValue(l[0]);
      style_(R('C' + r), { bg: l[3] || DASH.PINK, nf: l[2], align: 'right', border: true }).setFormula(l[1]);
    });
  };
  block(7, 'Most underweight Industry', DASH.YELLOW);
  style_(R('B8'), { bg: DASH.PEACH, bold: true, border: true, wrap: true }).setFormula('=UnderSector');
  style_(R('C8'), { bg: DASH.PEACH, bold: true, border: true, nf: DASH.SIGNED_PCT }).setFormula('=UnderDiff');
  block(10, 'Most overweight Industry', DASH.YELLOW);
  style_(R('B11'), { bg: DASH.PEACH, bold: true, border: true, wrap: true }).setFormula('=OverSector');
  style_(R('C11'), { bg: DASH.PEACH, bold: true, border: true, nf: DASH.SIGNED_PCT }).setFormula('=OverDiff');
  block(13, 'Stocks:', DASH.RED_HDR, [
    ['Market Value:', '=StocksValue', DASH.MONEY], ['$ Gain/Loss:', '=StocksValue-StocksCost', DASH.MONEY],
    ['% Gain/Loss:', '=IFERROR((StocksValue-StocksCost)/StocksCost,0)', DASH.PCT],
    ['Dividend Income (yr):', '=FwdIncome', DASH.MONEY], ['Yield on Cost:', '=YieldOnCost', DASH.PCT]]);
  block(20, 'Crypto', DASH.YELLOW, [
    ['Market Value:', '=CryptoValue', DASH.MONEY, DASH.PEACH], ['$ Gain/Loss:', '=CryptoValue-CryptoCost', DASH.MONEY, DASH.PEACH],
    ['% Gain/Loss:', '=IFERROR((CryptoValue-CryptoCost)/CryptoCost,0)', DASH.PCT, DASH.PEACH]]);
  block(25, 'Cash', DASH.YELLOW, [['Cash Position:', '=CashBalance', DASH.MONEY, DASH.PEACH]]);
  style_(R('B28:C28'), { merge: true, bg: DASH.HOLD_GREEN, bold: true, size: 18, border: true }).setValue('Total Holdings:');
  style_(R('B29:C34'), { merge: true, bg: DASH.HOLD_GREEN, bold: true, size: 60, border: true }).setFormula('=Positions');

  // Summary table
  const rows = [
    ['Portfolio Value (stocks + crypto + cash):', '=TotalValue', DASH.MONEY],
    ['Cost Basis (incl. reinvested dividends):', '=CostBasis', DASH.MONEY],
    ['Money In (your own contributions):', '=MoneyIn', DASH.MONEY],
    ['$ Gain/Loss:', '=PriceGain', DASH.MONEY],
    ['% Gain/Loss:', '=IFERROR(PriceGain/CostBasis,0)', DASH.PCT],
    ['Realized Gains:', '=RealizedGain', DASH.MONEY],
    ['$ Total Return (dividends included & realized):', '=TotalReturn', DASH.MONEY],
    ['% Total Return (on your money):', '=IFERROR(TotalReturn/TotalBuys,0)', DASH.PCT],
    ['Annualized Return (XIRR):', '=XirrReturn', DASH.PCT],
    ['Dividend Growth (5Y, weighted by income):', '=PortfolioDivCAGR5', DASH.PCT],
  ];
  rows.forEach((r, i) => {
    const row = 7 + i;
    style_(R('E' + row + ':H' + row), { merge: true, bg: DASH.PINK, bold: true, align: 'right', border: true }).setValue(r[0]);
    style_(R('I' + row + ':L' + row), { merge: true, bg: DASH.PINK, nf: r[2], border: true }).setFormula(r[1]);
  });

  // Markets
  style_(R('E18:L18'), { merge: true, bg: DASH.BLUE, bold: true, border: true }).setValue('How Markets are Moving Today');
  [['Index:', null], ['Price:', '=IFERROR(GOOGLEFINANCE(F19,"price"),"")', '#,##0.00'],
   ['Day\'s Change $:', '=IFERROR(GOOGLEFINANCE(F19,"price")-GOOGLEFINANCE(F19,"closeyest"),"")', '+#,##0.00;-#,##0.00'],
   ['Day\'s Change %:', '=IFERROR(GOOGLEFINANCE(F19,"changepct")/100,"")', DASH.SIGNED_PCT],
   ['Volume:', '=IFERROR(GOOGLEFINANCE(F19,"volume"),"")', '#,##0']].forEach((m, i) => {
    const row = 19 + i;
    style_(R('E' + row), { bg: DASH.LIGHT_BLUE, align: 'left', border: true }).setValue(m[0]);
    const c = style_(R('F' + row), { bg: DASH.LIGHT_BLUE, align: 'right', border: true, nf: m[2] });
    if (m[1]) c.setFormula(m[1]);
  });
  R('F19').setValue('.IXIC').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['.IXIC', '.INX', '.DJI'], true).setHelpText('.IXIC Nasdaq · .INX S&P 500 · .DJI Dow Jones').build());
  style_(R('G19:L23'), { merge: true, bg: DASH.LIGHT_BLUE, border: true }).setFormula(
    '=IFERROR(SPARKLINE(QUERY(GOOGLEFINANCE(F19,"close",TODAY()-30,TODAY()),"select Col2 offset 1",0),{"charttype","column";"color","#0D47A1"}),"")');

  // Future outlook (Base scenario from the Projections tab)
  style_(R('E37:L37'), { merge: true, bg: DASH.OUTLOOK_HDR, bold: true, border: true }).setValue('Future Portfolio Outlook');
  [['E38:F38', 'Years'], ['G38:H38', 'Yearly Dividends'], ['I38:J38', 'Total Dividend Income'], ['K38:L38', 'Portfolio Value']]
    .forEach(h => style_(R(h[0]), { merge: true, bg: DASH.OUTLOOK, bold: true, border: true, wrap: true }).setValue(h[1]));
  [5, 10, 15, 20, 25, 30, 40].forEach((y, i) => {
    const row = 39 + i, idx = y + 1;
    style_(R('E' + row + ':F' + row), { merge: true, bg: DASH.OUTLOOK, bold: true, border: true }).setValue(y + ' Years');
    style_(R('G' + row + ':H' + row), { merge: true, bg: DASH.OUTLOOK, nf: DASH.MONEY, border: true }).setFormula('=INDEX(Projections!$J$13:$J$53,' + idx + ')');
    style_(R('I' + row + ':J' + row), { merge: true, bg: DASH.OUTLOOK, nf: DASH.MONEY, border: true }).setFormula('=INDEX(Projections!$L$13:$L$53,' + idx + ')');
    style_(R('K' + row + ':L' + row), { merge: true, bg: DASH.OUTLOOK, nf: DASH.MONEY, border: true }).setFormula('=INDEX(Projections!$H$13:$H$53,' + idx + ')');
  });
  style_(R('E46:L46'), { merge: true, color: '#666666', size: 9, wrap: true }).setFormula(
    '="Base scenario: "&TEXT(GpBase,"0.0%")&" price growth, "&TEXT(GdBase,"0.0%")&" dividend growth for "&DivGrowthYears&' +
    '" years, $"&MonthlyContribution&"/month, dividends reinvested. Other scenarios: Projections tab."');
  sh.setRowHeight(46, 30);

  // Upcoming dividends
  style_(R('E49:L49'), { merge: true, bg: DASH.GREEN_HDR, bold: true, border: true }).setValue('Upcoming Dividends');
  [['E50:F50', 'Ticker'], ['G50:H50', 'Ex-date'], ['I50:J50', 'Pay date'], ['K50:L50', 'My expected $']]
    .forEach(h => style_(R(h[0]), { merge: true, bg: DASH.GREEN_HDR, bold: true, border: true }).setValue(h[1]));
  for (let i = 1; i <= 6; i++) {
    const row = 50 + i, src = "'Dividend Calendar'!";
    const cell = (a1, col, nf) => style_(R(a1), { merge: true, nf: nf, border: true })
      .setFormula('=IFERROR(IF(ISNUMBER(' + src + '$B$' + (3 + i) + '),INDEX(' + src + '$' + col + '$4:$' + col + '$203,' + i + '),""),"")');
    cell('E' + row + ':F' + row, 'A');
    cell('G' + row + ':H' + row, 'B', 'mmm d, yyyy');
    cell('I' + row + ':J' + row, 'C', 'mmm d, yyyy');
    cell('K' + row + ':L' + row, 'G', DASH.MONEY);
  }
  style_(R('E58:P58'), { merge: true, color: '#666666', size: 9, align: 'left' }).setFormula(
    '="Prices missing: "&PricesMissing&"   ·   Last automatic update: "&LastDailyRun&"   ·   Prices: GOOGLEFINANCE (~20 min delay)"');

  // Green / red for gains and losses
  const gl = [R('C14:C16'), R('C21:C23'), R('I10:L11'), R('I13:L15'), R('F21:F22')];
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setFontColor('#38761d').setRanges(gl).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor('#cc0000').setRanges(gl).build(),
  ]);

  // Charts
  const calc = ss.getSheetByName('Calc');
  const chart = (name, fn) => {
    try { sh.insertChart(fn()); } catch (e) { log_('buildDashboard', 'WARN', name + ' chart: ' + e.message); }
  };
  chart('Industries', () => sh.newChart().setChartType(Charts.ChartType.PIE).addRange(calc.getRange('E1:F30'))
    .setNumHeaders(1).setPosition(7, 14, 0, 0).setOption('title', 'Industries').setOption('is3D', true)
    .setOption('pieSliceText', 'none').setOption('legend', { position: 'labeled' })
    .setOption('width', 460).setOption('height', 330).build());
  chart('Equity vs cost', () => sh.newChart().setChartType(Charts.ChartType.COLUMN).addRange(calc.getRange('J1:L41'))
    .setNumHeaders(1).setPosition(25, 5, 0, 0).setOption('colors', ['#4285f4', '#db4437'])
    .setOption('legend', { position: 'top' }).setOption('width', 1000).setOption('height', 230).build());
  chart('Map', () => sh.newChart().setChartType(Charts.ChartType.GEO).addRange(calc.getRange('N1:O30'))
    .setNumHeaders(1).setPosition(37, 14, 0, 0).setOption('colorAxis', { colors: ['#e06666', '#6aa84f'] })
    .setOption('backgroundColor', '#d9d2e9').setOption('datalessRegionColor', '#ffffff')
    .setOption('width', 460).setOption('height', 300).build());

  sh.activate();
  log_('buildDashboard', 'OK', 'dashboard rebuilt');
}
