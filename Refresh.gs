/**
 * Programme monitoring — weekly computation engine.
 *
 * Rules:
 *  - a solution applies to the programme of its branch, or to the whole group when it has no branch;
 *  - a programme branch (case × group × branch) uses its own Applied solutions, otherwise its group's;
 *  - it is closed when it has at least one Applied solution and all of them are closed (X/Y, threshold);
 *  - a case is open for a programme if at least one of its branches of this programme is open.
 * Each Applied solution still blocking is put in exactly one category (see CAT in Config.gs):
 *  - no implementation;
 *  - all implementations closed (column CU) = data inconsistency;
 *  - at least one open implementation with a manual type (column CL) = manual;
 *  - otherwise open implementations tracked in TOM = in progress;
 *  - unknown status or type = unclassified (to fix in Codes.gs).
 */

const SEP = '\u241F';
const WRITE_CHUNK = 20000;

function refreshAll() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('A refresh is already running.');
  const t0 = Date.now();
  try {
    const props = PropertiesService.getScriptProperties();
    const firstRun = !props.getProperty('LAST_REFRESH');
    const P = readColumns_(CFG.PROGRAMMES);
    const S = readColumns_(CFG.SOLUTIONS);
    const now = new Date();
    const week = isoWeek_(now);
    const extraction = Utilities.formatDate(now, tz_(), 'dd/MM/yyyy HH:mm');
    const prev = firstRun ? null : readOpenKeys_();
    const res = compute_(P, S, { week: week, extraction: extraction, prev: prev });
    writeOutputs_(res, week, Utilities.formatDate(now, tz_(), 'yyyy-MM-dd'));
    removeLegacySheets_();
    const info = {
      date: extraction, week: week,
      seconds: Math.round((Date.now() - t0) / 1000),
      programmesRows: P._n, solutionsRows: S._n, dataRows: res.data.length, programmes: res.summary.length,
      excludedDossiers: res.excluded,
      unknownStatus: res.unknown.status, unknownTypes: res.unknown.types
    };
    props.setProperty('LAST_REFRESH', JSON.stringify(info));
    console.log('Refresh done', JSON.stringify(info));
    return info;
  } finally {
    lock.releaseLock();
  }
}

function refreshAllFromMenu() {
  const ui = SpreadsheetApp.getUi();
  const info = refreshAll();
  const unknown = info.unknownStatus.length + info.unknownTypes.length;
  ui.alert('Refresh completed',
    info.programmes + ' programmes, ' + info.dataRows + ' rows computed in ' + info.seconds + ' s. ' +
    info.excludedDossiers + ' cases excluded (status).' +
    (unknown ? '\n\n' + unknown + ' unknown implementation code(s): see "Unclassified" in the dashboard and add them to Codes.gs.' : ''),
    ui.ButtonSet.OK);
}

function installWeeklyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(function (t) { return t.getHandlerFunction() === 'refreshAll'; })
    .forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('refreshAll').timeBased()
    .onWeekDay(ScriptApp.WeekDay[CFG.REFRESH.weekday]).atHour(CFG.REFRESH.hour).create();
  SpreadsheetApp.getUi().alert('Weekly refresh scheduled (' +
    CFG.REFRESH.weekday.charAt(0) + CFG.REFRESH.weekday.slice(1).toLowerCase() + ', around ' + CFG.REFRESH.hour + ':00).');
}

function removeLegacySheets_() {
  const ss = SpreadsheetApp.getActive();
  (CFG.SHEETS.legacy || []).forEach(function (n) {
    const sh = ss.getSheetByName(n);
    if (sh) ss.deleteSheet(sh);
  });
}

/** Reads only the useful columns of a source file, column by column. */
function readColumns_(src) {
  const fields = Object.keys(src.cols);
  const ranges = fields.map(function (f) {
    const c = src.cols[f];
    return "'" + src.sheet + "'!" + c + src.firstRow + ':' + c;
  });
  const res = Sheets.Spreadsheets.Values.batchGet(src.id, {
    ranges: ranges, majorDimension: 'COLUMNS', valueRenderOption: 'FORMATTED_VALUE'
  });
  const out = { _n: 0 };
  (res.valueRanges || []).forEach(function (vr, i) {
    const col = (vr.values && vr.values[0]) || [];
    out[fields[i]] = col;
    out._n = Math.max(out._n, col.length);
  });
  fields.forEach(function (f) { if (!out[f]) out[f] = []; });
  return out;
}

function val_(col, i) {
  const v = col[i];
  return v == null ? '' : String(v).trim();
}

function isClosed_(status) {
  const m = /^\s*(\d+)\s*\/\s*(\d+)/.exec(String(status || ''));
  if (!m) return false;
  const x = Number(m[1]), y = Number(m[2]);
  if (!y) return false;
  const threshold = CFG.CLOSED_THRESHOLDS[y] || y;
  return x >= threshold;
}

/* ---------- Implementation codes ---------- */

/** Normalised code: lower case, no accents, no tabs or quotes, no leading punctuation, single spaces. */
function normCode_(s) {
  return String(s == null ? '' : s).replace(/\t/g, ' ').trim().replace(/^"+|"+$/g, '').trim()
    .normalize('NFKD').replace(/[̀-ͯ]/g, '')
    .replace(/^[^A-Za-z0-9]+/, '').replace(/\s+/g, ' ').trim().toLowerCase();
}

/** Same as normCode_ with dates removed ("Go for delivery (12/03/26)" -> "go for delivery"). */
function looseCode_(s) {
  return normCode_(s)
    .replace(/\(?\b\d{1,2}[\/.]\d{1,2}[\/.]\d{2,4}\b\)?/g, ' ')
    .replace(/\(\s*\)/g, ' ')
    .replace(/\s+/g, ' ').trim()
    .replace(/\b(on|the)$/, '').trim()
    .replace(/^[^a-z0-9]+|[^a-z0-9)]+$/g, '').trim();
}

let CODE_MAPS_ = null;
function codeMaps_() {
  if (CODE_MAPS_) return CODE_MAPS_;
  const exact = function (open, closed) {
    const m = new Map();
    open.forEach(function (k) { m.set(normCode_(k), true); });
    closed.forEach(function (k) { m.set(normCode_(k), false); });
    return m;
  };
  const status = exact(CODES.IMPL_STATUS_OPEN, CODES.IMPL_STATUS_CLOSED);
  const loose = new Map(), conflict = new Set();
  status.forEach(function (v, k) {
    const l = looseCode_(k);
    if (!l) return;
    if (loose.has(l) && loose.get(l) !== v) conflict.add(l);
    loose.set(l, v);
  });
  conflict.forEach(function (k) { loose.delete(k); });
  CODE_MAPS_ = { type: exact(CODES.IMPL_TYPE_MANUAL, CODES.IMPL_TYPE_TOM), status: status, loose: loose };
  return CODE_MAPS_;
}

/** true = open, false = closed, null = unknown status. */
function implOpen_(status) {
  const n = normCode_(status);
  if (!n) return CODES.IMPL_STATUS_EMPTY === 'open' ? true : CODES.IMPL_STATUS_EMPTY === 'closed' ? false : null;
  const m = codeMaps_();
  if (m.status.has(n)) return m.status.get(n);
  const l = looseCode_(status);
  return m.loose.has(l) ? m.loose.get(l) : null;
}

/** true = manual, false = tracked in TOM, null = unknown type. */
function implManual_(code) {
  const n = normCode_(code);
  if (!n) return null;
  const m = codeMaps_();
  return m.type.has(n) ? m.type.get(n) : null;
}

/** Category of an Applied solution that is still blocking (see the rules at the top of the file). */
function classify_(sol) {
  const impls = Array.from(sol.impls.values());
  if (!impls.length) return CAT.NO_IMPL;
  if (impls.some(function (im) { return im.open === null; })) return CAT.UNCLASSIFIED;
  const open = impls.filter(function (im) { return im.open; });
  if (!open.length) return CAT.IMPL_CLOSED;
  if (open.some(function (im) { return im.manual === null; })) return CAT.UNCLASSIFIED;
  return open.some(function (im) { return im.manual; }) ? CAT.MANUAL : CAT.IN_PROGRESS;
}

const yn_ = function (b) { return b === null ? '?' : b ? 'Oui' : 'Non'; };

/** Core computation: pure, no sheet access (testable). */
function compute_(P, S, opts) {
  const unknownStatus = new Map(), unknownTypes = new Map();
  const bump = function (m, k) { m.set(k, (m.get(k) || 0) + 1); };

  // 1. "Applied" (closure) and "Proposed" (awaiting decision) solutions, by target then by ID.
  const applied = new Map(), proposed = new Map();
  for (let i = 0; i < S._n; i++) {
    const dossier = val_(S.dossier, i), solId = val_(S.solId, i);
    if (!dossier || !solId) continue;
    const state = val_(S.solState, i).toLowerCase();
    const bucket = state === CFG.SOLUTION_STATE.applied ? applied
      : state === CFG.SOLUTION_STATE.proposed ? proposed : null;
    if (!bucket) continue; // Rejected or empty: ignored
    const groupe = val_(S.groupe, i), branche = val_(S.branche, i);
    const target = branche ? 'I' + dossier + SEP + groupe + SEP + branche : 'G' + dossier + SEP + groupe;
    let m = bucket.get(target);
    if (!m) { m = new Map(); bucket.set(target, m); }
    let sol = m.get(solId);
    if (!sol) {
      const status = val_(S.solStatus, i);
      sol = { id: solId, name: val_(S.solName, i), type: val_(S.solType, i), status: status,
        closed: isClosed_(status), origin: branche ? 'Programme' : 'Groupe', impls: new Map() };
      m.set(solId, sol);
    }
    const implId = val_(S.implId, i);
    if (implId && !sol.impls.has(implId)) {
      const code = val_(S.implCode, i), st = val_(S.implStatus, i);
      sol.impls.set(implId, { id: implId, label: val_(S.implLabel, i), code: code, status: st, resp: val_(S.implResp, i),
        open: implOpen_(st), manual: implManual_(code) });
    }
  }

  // 2. Programme branches (excluding cancelled / closed / solved cases), with their effective solutions.
  const exStatus = new Set(CFG.EXCLUDED_DOSSIER_STATUS.map(function (s) { return s.toUpperCase(); }));
  const excluded = new Set();
  const progs = new Map();
  const seen = new Set();
  const values = function (m) { return m ? Array.from(m.values()) : []; };
  for (let i = 0; i < P._n; i++) {
    const dossier = val_(P.dossier, i), name = val_(P.programme, i);
    if (!dossier || !name) continue;
    const dStatus = val_(P.dossierStatus, i);
    if (exStatus.has(dStatus.toUpperCase())) { excluded.add(dossier); continue; }
    const groupe = val_(P.groupe, i), branche = val_(P.branche, i);
    const ik = dossier + SEP + groupe + SEP + branche;
    if (seen.has(ik)) continue;
    seen.add(ik);
    const gk = 'G' + dossier + SEP + groupe;
    const own = branche ? applied.get('I' + ik) : null;
    const sols = own && own.size ? values(own) : values(applied.get(gk));
    const props = (branche ? values(proposed.get('I' + ik)) : []).concat(values(proposed.get(gk)));
    const open = sols.length === 0 || sols.some(function (s) { return !s.closed; });
    let dm = progs.get(name);
    if (!dm) { dm = new Map(); progs.set(name, dm); }
    let e = dm.get(dossier);
    if (!e) {
      const first = val_(P.leaderFirst, i), last = val_(P.leaderLast, i);
      e = { dossier: dossier, urlId: val_(P.urlId, i) || dossier, open: false, inst: [],
        og: val_(P.og, i), status: dStatus, supplier: val_(P.supplier, i),
        leaderFirst: first, leaderLast: last, leader: (first + ' ' + last).trim() };
      dm.set(dossier, e);
    }
    e.inst.push({ groupe: groupe, groupName: val_(P.groupName, i), branche: branche, sols: sols, props: props, open: open });
    if (open) e.open = true;
  }

  // 3. Outputs: rows (sorted by programme), summary, open keys, case leaders.
  const data = [], summary = [], openKeys = [];
  const leaders = new Map();
  const coll = function (a, b) { return a.localeCompare(b, 'fr', { numeric: true }); };
  const empty = ['', '', '', '', '', '', '', '', '', ''];
  const solCols = function (nature, s, im) {
    return [nature, s.origin, s.id, s.name, s.type, s.status,
      im ? im.id : '', im ? im.label : '', im ? im.status : '', im ? im.resp : ''];
  };
  Array.from(progs.keys()).sort(coll).forEach(function (name) {
    const dm = progs.get(name);
    const start = data.length;
    let total = 0, nOpen = 0, nNew = 0, withProp = 0;
    const bloq = new Set(), her = new Set(), impls = new Set(), prop = new Set(), missing = new Set();
    const cats = {};
    Object.keys(CAT).forEach(function (k) { cats[k] = 0; });
    Array.from(dm.values()).sort(function (a, b) { return coll(a.dossier, b.dossier); }).forEach(function (e) {
      total++;
      if (e.leader) {
        const lk = normName_(e.leaderFirst) + '|' + normName_(e.leaderLast);
        let L = leaders.get(lk);
        if (!L) { L = { first: e.leaderFirst, last: e.leaderLast, name: e.leader, progs: {} }; leaders.set(lk, L); }
        const lp = L.progs[name] || (L.progs[name] = { open: 0, total: 0 });
        lp.total++;
        if (e.open) lp.open++;
      }
      const baseFor = function (ins) {
        return [name, e.dossier, e.urlId, e.og, e.status, e.leader, e.supplier, ins.groupe, ins.groupName, ins.branche];
      };
      const extra = function (cat, im) {
        return [cat, im ? im.code : '', im ? yn_(im.open) : '', im ? yn_(im.manual) : ''];
      };
      if (!e.open) {
        // One row per closed case: used for the filtered closure rate (OG, case leader, supplier).
        data.push(baseFor(e.inst[0]).concat(['Dossier soldé'], empty.slice(1), ['', '', '', opts.week], extra('', null)));
        return;
      }
      nOpen++;
      const key = name + SEP + e.dossier;
      openKeys.push([key]);
      const isNew = !!opts.prev && !opts.prev.has(key);
      if (isNew) nNew++;
      const all = new Map();
      e.inst.forEach(function (ins) {
        ins.sols.forEach(function (s) { all.set(solKey_(s, ins), s); });
      });
      let solDone = 0;
      all.forEach(function (s) { if (s.closed) solDone++; });
      const tail = [isNew ? 'Oui' : '', String(all.size), String(solDone), opts.week];
      const done = new Set();
      let hasProp = false;
      // Group level: a group is "missing" when none of its open branches has an applied or proposed solution.
      const gProps = {};
      e.inst.forEach(function (ins) { if (ins.open) gProps[ins.groupe] = (gProps[ins.groupe] || 0) + ins.props.length; });
      e.inst.forEach(function (ins) {
        const base = baseFor(ins);
        const push = function (nature, s, withImpls, cat) {
          const k = nature + SEP + solKey_(s, ins);
          if (done.has(k)) return false;
          done.add(k);
          const list = withImpls && s.impls.size ? Array.from(s.impls.values()) : [null];
          list.forEach(function (im) { data.push(base.concat(solCols(nature, s, im), tail, extra(cat, im))); });
          return true;
        };
        if (ins.open && !ins.sols.length) {
          const isMissing = !gProps[ins.groupe];
          const mk = 'M' + SEP + ins.groupe;
          if (!done.has(mk)) {
            done.add(mk);
            data.push(base.concat(['Aucune solution'], empty.slice(1), tail, extra(isMissing ? CAT.MISSING : '', null)));
          }
          if (isMissing) missing.add(e.dossier + SEP + ins.groupe);
        }
        ins.sols.forEach(function (s) {
          if (s.closed) { push('Soldée', s, false, CAT.CLOSED); return; }
          const cat = classify_(s);
          if (cat === CAT.UNCLASSIFIED) {
            s.impls.forEach(function (im) {
              if (im.open === null) bump(unknownStatus, im.status || '(empty)');
              else if (im.open && im.manual === null) bump(unknownTypes, im.code || '(empty)');
            });
          }
          if (push('Bloquante', s, true, cat)) {
            const sk = e.dossier + SEP + solKey_(s, ins);
            bloq.add(sk);
            cats[cat]++;
            if (s.origin === 'Groupe') her.add(sk);
          }
          s.impls.forEach(function (im) { impls.add(im.id); });
        });
        if (!ins.open) return; // proposals on a group already closed are not counted
        ins.props.forEach(function (s) {
          hasProp = true;
          if (push('Proposition', s, true, CAT.PROPOSED)) prop.add(e.dossier + SEP + solKey_(s, ins));
        });
      });
      if (hasProp) withProp++;
    });
    const len = data.length - start;
    summary.push([name, total, nOpen, nNew, bloq.size, her.size, impls.size,
      total ? Math.round((total - nOpen) / total * 100) : 0,
      len ? start + 2 : '', len, opts.week, opts.extraction, withProp,
      prop.size, missing.size, cats.NO_IMPL, cats.IMPL_CLOSED, cats.IN_PROGRESS, cats.MANUAL, cats.UNCLASSIFIED]);
  });
  const top = function (m) {
    return Array.from(m.entries()).sort(function (a, b) { return b[1] - a[1]; }).slice(0, 30);
  };
  const leaderRows = Array.from(leaders.values())
    .sort(function (a, b) { return coll(a.name, b.name); })
    .map(function (L) { return [L.first, L.last, L.name, JSON.stringify(L.progs)]; });
  return { data: data, summary: summary, openKeys: openKeys, leaders: leaderRows, excluded: excluded.size,
    unknown: { status: top(unknownStatus), types: top(unknownTypes) } };
}

/** Normalised person name: lower case, no accents, letters only ("Jean-Pierre" -> "jeanpierre"). */
function normName_(s) {
  return String(s || '').normalize('NFKD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z]/g, '');
}

function solKey_(s, ins) {
  return (s.origin === 'Groupe' ? 'G' + ins.groupe : 'I' + ins.groupe + SEP + ins.branche) + SEP + s.id;
}

function writeOutputs_(res, week, isoDate) {
  writeTable_(CFG.SHEETS.data, DATA_HEADER, res.data, true);
  writeTable_(CFG.SHEETS.progs, PROG_HEADER, res.summary, false);
  writeTable_(CFG.SHEETS.leaders, LEADER_HEADER, res.leaders, true);
  const w = HIST_HEADER.length;
  const hist = readTable_(CFG.SHEETS.hist)
    .map(function (r) { r[0] = normWeek_(r[0]); while (r.length < w) r.push(''); return r.slice(0, w); })
    .filter(function (r) { return r[0] !== week; });
  res.summary.forEach(function (r) {
    hist.push([week, isoDate, r[0], r[2], r[4], r[6], r[13], r[14], r[15], r[16], r[17], r[18], r[19]]);
  });
  writeTable_(CFG.SHEETS.hist, HIST_HEADER, hist, false);
  writeTable_(CFG.SHEETS.open, ['Cle'], res.openKeys, true);
}

function readOpenKeys_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEETS.open);
  if (!sh) return null;
  const n = sh.getLastRow() - 1;
  if (n < 1) return new Set();
  return new Set(sh.getRange(2, 1, n, 1).getDisplayValues().map(function (r) { return r[0]; }));
}

/* ---------- Sheet helpers ---------- */

function sheet_(name) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.hideSheet(); }
  return sh;
}

function ensureSheet_(name, header) {
  const sh = sheet_(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    sh.getRange(1, 1, sh.getMaxRows(), header.length).setNumberFormat('@');
    sh.setFrozenRows(1);
  }
  return sh;
}

function writeTable_(name, header, rows, asText) {
  const sh = sheet_(name);
  const nRows = Math.max(rows.length + 1, 2), nCols = header.length;
  sh.clearContents();
  fit_(sh, nRows, nCols);
  if (asText) sh.getRange(1, 1, nRows, nCols).setNumberFormat('@');
  sh.getRange(1, 1, 1, nCols).setValues([header]);
  for (let i = 0; i < rows.length; i += WRITE_CHUNK) {
    const part = rows.slice(i, i + WRITE_CHUNK);
    sh.getRange(2 + i, 1, part.length, nCols).setValues(part);
  }
  sh.setFrozenRows(1);
  SpreadsheetApp.flush();
}

function fit_(sh, rows, cols) {
  const mr = sh.getMaxRows(), mc = sh.getMaxColumns();
  if (mr < rows) sh.insertRowsAfter(mr, rows - mr);
  else if (mr > rows) sh.deleteRows(rows + 1, mr - rows);
  if (mc < cols) sh.insertColumnsAfter(mc, cols - mc);
  else if (mc > cols) sh.deleteColumns(cols + 1, mc - cols);
}

/** Always as displayed values (text) to avoid automatic conversions. */
function readTable_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getDisplayValues();
}

/* ---------- Dates ---------- */

function tz_() { return Session.getScriptTimeZone() || 'Europe/Paris'; }

function today_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }

function isoWeek_(d) {
  const t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const day = t.getUTCDay() || 7;
  t.setUTCDate(t.getUTCDate() + 4 - day);
  const y0 = new Date(Date.UTC(t.getUTCFullYear(), 0, 1));
  const w = Math.ceil(((t - y0) / 86400000 + 1) / 7);
  return t.getUTCFullYear() + '-W' + String(w).padStart(2, '0');
}

/** Older history rows used "2026-S40": read them as "2026-W40". */
function normWeek_(w) { return String(w || '').replace(/-S(\d+)$/, '-W$1'); }

function isIsoDate_(s) { return /^\d{4}-\d{2}-\d{2}$/.test(String(s || '')); }

/** Normalises a date (ISO text or dd/mm/yyyy) to yyyy-mm-dd. */
function dateStr_(v) {
  const s = String(v == null ? '' : v).trim();
  if (isIsoDate_(s)) return s;
  const m = /^(\d{2})\/(\d{2})\/(\d{4})/.exec(s);
  return m ? m[3] + '-' + m[2] + '-' + m[1] : s;
}

function fmtDate_(iso) {
  if (!isIsoDate_(iso)) return iso || '—';
  const p = iso.split('-');
  return p[2] + '/' + p[1] + '/' + p[0];
}
