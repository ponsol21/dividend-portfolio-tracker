// Runs the Code.gs dividend logic on the fixture data and writes build/prefill.json, which
// tools/build_workbook.py uses to pre-fill the new workbook (same code path as the live script).
//   osascript -l JavaScript apps_script/tests/export_prefill.js "$PWD" 2026-10-19 "2026-10-19T21:30:00Z"
ObjC.import('Foundation');

function readFile_(p) {
  const s = $.NSString.stringWithContentsOfFileEncodingError($(p), $.NSUTF8StringEncoding, null);
  if (s.isNil()) throw new Error('cannot read ' + p);
  return s.js;
}

function run(argv) {
  const root = argv[0], today = argv[1], fetchedAt = argv[2];
  (0, eval)(readFile_(root + '/apps_script/Code.gs'));
  const dir = root + '/apps_script/tests/fixtures/';
  const fm = $.NSFileManager.defaultManager;
  const files = ObjC.deepUnwrap(fm.contentsOfDirectoryAtPathError($(dir), null)) || [];
  const tickers = files.filter(f => /^[A-Z.]+\.json$/.test(f)).map(f => f.replace('.json', '')).sort();

  const out = { today: today, fetchedAt: fetchedAt, stats: [], schedule: [], divhistory: [], splits: [], meta: [] };
  const lines = [];
  tickers.forEach(t => {
    const d = parseAvDividends_(JSON.parse(readFile_(dir + t + '.json')), t);
    let sp = [], known = false;
    if (files.indexOf(t + '.splits.json') >= 0) {
      sp = parseAvSplits_(JSON.parse(readFile_(dir + t + '.splits.json')), t);
      known = true;
      out.meta.push([t, 'split', fetchedAt]);
      sp.forEach(s => out.splits.push([t, s.date, s.factor, fetchedAt]));
    }
    out.meta.push([t, 'div', fetchedAt]);
    d.forEach(r => out.divhistory.push([t, r.ex, r.decl, r.rec, r.pay, r.amount, fetchedAt]));
    const res = computeTickerStats_(t, d, sp, today, known);
    out.stats.push(res.stats);
    buildSchedule_(t, res.rows, res.stats, today).forEach(x => out.schedule.push(x));
    const s = res.stats, pct = v => (v === '' ? '   -  ' : (v * 100).toFixed(1).padStart(5) + '%');
    lines.push([t.padEnd(5), s.frequency.padEnd(9), String(s.current.toFixed ? s.current.toFixed(4) : s.current).padStart(7),
      String(s.fwdAnnual.toFixed ? s.fwdAnnual.toFixed(2) : s.fwdAnnual).padStart(6), (s.lastRaiseDate || '-').padEnd(10),
      pct(s.lastRaisePct), pct(s.cagr1), pct(s.cagr3), pct(s.cagr5), pct(s.cagr10), String(s.streak).padStart(3),
      (s.nextRaiseEst || '-').padEnd(7), s.specials, s.status + (s.streakNote ? ' | ' + s.streakNote : '')].join(' '));
  });
  $.NSString.alloc.initWithUTF8String(JSON.stringify(out)).writeToFileAtomicallyEncodingError($(root + '/build/prefill.json'), true, $.NSUTF8StringEncoding, null);
  return 'tkr   freq      current   fwd  lastRaise  raise%   1y     3y     5y     10y  strk next  sp status\n' + lines.join('\n') +
    '\nschedule rows: ' + out.schedule.length;
}
