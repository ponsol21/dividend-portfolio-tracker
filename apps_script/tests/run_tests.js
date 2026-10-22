// Unit tests for the pure functions in Code.gs, run locally with macOS JavaScript for Automation:
//   osascript -l JavaScript apps_script/tests/run_tests.js "$PWD"
// Uses made-up example data only. Tests on real dividend data for your own holdings are in
// run_private_tests.js, which stays on your computer (see .gitignore).
ObjC.import('Foundation');

function readFile_(p) {
  const s = $.NSString.stringWithContentsOfFileEncodingError($(p), $.NSUTF8StringEncoding, null);
  if (s.isNil()) throw new Error('cannot read ' + p);
  return s.js;
}

function run(argv) {
  const root = argv[0];
  (0, eval)(readFile_(root + '/apps_script/Code.gs'));

  const results = [];
  let failures = 0;
  const check = (name, cond, detail) => {
    results.push((cond ? 'PASS ' : 'FAIL ') + name + (cond || detail === undefined ? '' : '  -> ' + detail));
    if (!cond) failures++;
  };
  const near = (a, b, eps) => Math.abs(a - b) <= (eps || 1e-9);

  // Input validation
  ['XYZ', 'BRK.B', 'T', 'ABCD'].forEach(t => check('valid ticker ' + t, isValidTicker_(t)));
  ['ko', '=HYPERLINK("x")', 'ABCDEFGHIJK', '', 'A B', 'XYZ&apikey=x', '../etc'].forEach(t => check('rejects ticker ' + JSON.stringify(t), !isValidTicker_(t)));

  // Date helpers
  check('addMonths end of month', addMonths_('2026-01-31', 1) === '2026-02-28');
  check('addMonths leap day', addMonths_('2024-02-29', 12) === '2025-02-28');
  check('addMonths negative', addMonths_('2026-03-15', -14) === '2025-01-15');
  check('daysBetween', daysBetween_('2026-01-01', '2026-03-01') === 59);

  // Parsing hardening
  let threw = false;
  try { parseAvDividends_({ Information: 'rate limit reached' }, 'XYZ'); } catch (e) { threw = /rate limit/.test(e.message); }
  check('rate-limit response raises', threw);
  threw = false;
  try { parseAvDividends_({ symbol: 'ABC', data: [] }, 'XYZ'); } catch (e) { threw = /mismatch/.test(e.message); }
  check('symbol mismatch raises', threw);
  const dirty = parseAvDividends_({ symbol: 'XYZ', data: [
    { ex_dividend_date: '=IMPORTXML("x")', amount: '0.5' },
    { ex_dividend_date: '2026-01-02', amount: 'None' },
    { ex_dividend_date: '2026-01-03', amount: '-1' },
    { ex_dividend_date: '2026-01-04', payment_date: '=1+1', amount: '0.25' },
  ] }, 'XYZ');
  check('drops malformed rows', dirty.length === 1 && dirty[0].ex === '2026-01-04' && dirty[0].pay === '', JSON.stringify(dirty));
  check('finnhub payout parse', near(parseFinnhubPayout_({ metric: { payoutRatioTTM: 65.5 } }), 0.655) && parseFinnhubPayout_({ metric: {} }) === '');
  check('scrub removes keys', scrub_('GET https://x?apikey=SECRET123&a=1 token=abc', ['SECRET123']).indexOf('SECRET123') < 0 &&
    scrub_('token=abc').indexOf('abc') < 0);

  // DRIP auto-logging (made-up quarterly dividends for an example ticker)
  const xyzDivs = { XYZ: parseAvDividends_({ symbol: 'XYZ', data: [
    { ex_dividend_date: '2026-09-01', payment_date: '2026-09-15', amount: '0.51' },
    { ex_dividend_date: '2026-12-01', payment_date: '2026-12-15', amount: '0.53' },
    { ex_dividend_date: '2027-03-01', payment_date: '2027-03-15', amount: '0.53' },
    { ex_dividend_date: '2027-05-15', payment_date: '2027-06-01', amount: '0.53' },
  ] }, 'XYZ') };
  const seed = [
    { date: '2026-10-19', ticker: 'XYZ', type: 'Stock', action: 'Buy', shares: 2.0, source: 'broker-seed', key: '' },
    { date: '2026-10-19', ticker: 'XYZ', type: 'Stock', action: 'Dividend Reinvest', shares: 0, source: 'broker-seed', key: '' },
  ];
  const drip1 = computeDripRows_(seed, xyzDivs, '2026-12-20', 60);
  check('DRIP adds exactly the Dec 15 payment', drip1.length === 1 && drip1[0].key === 'XYZ|2026-12-15', JSON.stringify(drip1));
  check('DRIP amount = shares x dividend', drip1.length === 1 && near(drip1[0].amount, Math.round(2.0 * 0.53 * 100) / 100), drip1.length && drip1[0].amount);
  const withRow = seed.concat([{ date: '2026-12-15', ticker: 'XYZ', type: 'Stock', action: 'Dividend Reinvest', shares: 0.0102, source: 'auto-drip', key: 'XYZ|2026-12-15' }]);
  check('DRIP idempotent (second run adds nothing)', computeDripRows_(withRow, xyzDivs, '2026-12-20', 60).length === 0);
  check('DRIP skips payments already in broker snapshot', computeDripRows_(seed, xyzDivs, '2026-10-19', 60).length === 0);
  const lateBuy = seed.concat([{ date: '2026-12-05', ticker: 'XYZ', type: 'Stock', action: 'Buy', shares: 10, source: 'manual', key: '' }]);
  const drip2 = computeDripRows_(lateBuy, xyzDivs, '2026-12-20', 60);
  check('DRIP ignores shares bought after ex-date', drip2.length === 1 && near(drip2[0].amount, drip1[0].amount), drip2.length && drip2[0].amount);
  const late = computeDripRows_(seed, xyzDivs, '2027-06-30', 60);
  check('DRIP respects lookback window', late.length === 1 && late.every(r => r.date >= '2027-05-01'), JSON.stringify(late));

  results.push('');
  results.push(failures ? failures + ' FAILED' : 'ALL ' + (results.length - 1) + ' TESTS PASSED');
  return results.join('\n');
}
