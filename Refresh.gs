/**
 * Pilotage des clôtures — moteur de calcul hebdomadaire.
 *
 * Règles :
 *  - une solution s'applique au programme si sa clé contient la branche, sinon à tout le groupe ;
 *  - un programme (dossier × groupe × branche) utilise ses propres solutions s'il en a, sinon celles du groupe ;
 *  - il est fermé s'il a au moins une solution et que toutes sont soldées (X/Y avec X = Y, ou seuil spécial) ;
 *  - un dossier est ouvert pour un programme si au moins une de ses branches de ce programme est ouverte.
 */

const SEP = '␟';
const WRITE_CHUNK = 20000;

function refreshAll() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('Une actualisation est déjà en cours.');
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
    const info = {
      date: extraction, week: week,
      seconds: Math.round((Date.now() - t0) / 1000),
      programmesRows: P._n, solutionsRows: S._n, dataRows: res.data.length, programmes: res.summary.length,
      excludedDossiers: res.excluded
    };
    props.setProperty('LAST_REFRESH', JSON.stringify(info));
    console.log('Actualisation terminée', JSON.stringify(info));
    return info;
  } finally {
    lock.releaseLock();
  }
}

function refreshAllFromMenu() {
  const ui = SpreadsheetApp.getUi();
  const info = refreshAll();
  ui.alert('Actualisation terminée',
    info.programmes + ' programmes, ' + info.dataRows + ' lignes calculées en ' + info.seconds + ' s. ' +
    info.excludedDossiers + ' dossiers exclus (statut ou OG).',
    ui.ButtonSet.OK);
}

function installWeeklyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(function (t) { return t.getHandlerFunction() === 'refreshAll'; })
    .forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('refreshAll').timeBased()
    .onWeekDay(ScriptApp.WeekDay[CFG.REFRESH.weekday]).atHour(CFG.REFRESH.hour).create();
  SpreadsheetApp.getUi().alert('Actualisation programmée chaque semaine (' +
    CFG.REFRESH.weekday.toLowerCase() + ', vers ' + CFG.REFRESH.hour + ' h).');
}

/** Lit uniquement les colonnes utiles d'un fichier source, colonne par colonne. */
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

/** Cœur du calcul : pur, sans accès aux feuilles (testable). */
function compute_(P, S, opts) {
  // 1. Solutions « Applied » (calcul) et « Proposed » (information), par cible puis par ID.
  const applied = new Map(), proposed = new Map();
  for (let i = 0; i < S._n; i++) {
    const dossier = val_(S.dossier, i), solId = val_(S.solId, i);
    if (!dossier || !solId) continue;
    const state = val_(S.solState, i).toLowerCase();
    const bucket = state === CFG.SOLUTION_STATE.applied ? applied
      : state === CFG.SOLUTION_STATE.proposed ? proposed : null;
    if (!bucket) continue; // Rejected ou vide : ignoré
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
      sol.impls.set(implId, { id: implId, type: val_(S.implType, i), status: val_(S.implStatus, i), resp: val_(S.implResp, i) });
    }
  }

  // 2. Branches de programme (hors dossiers exclus), rattachées à leurs solutions effectives.
  const exStatus = new Set(CFG.EXCLUDED_DOSSIER_STATUS.map(function (s) { return s.toUpperCase(); }));
  const exOg = new Set(CFG.EXCLUDED_OG.map(normOg_));
  const excluded = new Set();
  const progs = new Map();
  const seen = new Set();
  const values = function (m) { return m ? Array.from(m.values()) : []; };
  for (let i = 0; i < P._n; i++) {
    const dossier = val_(P.dossier, i), name = val_(P.programme, i);
    if (!dossier || !name) continue;
    const dStatus = val_(P.dossierStatus, i), og = val_(P.og, i);
    if (exStatus.has(dStatus.toUpperCase()) || exOg.has(normOg_(og))) { excluded.add(dossier); continue; }
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
      e = { dossier: dossier, urlId: val_(P.urlId, i) || dossier, open: false, inst: [],
        og: og, status: dStatus, supplier: val_(P.supplier, i),
        leader: (val_(P.leaderFirst, i) + ' ' + val_(P.leaderLast, i)).trim() };
      dm.set(dossier, e);
    }
    e.inst.push({ groupe: groupe, groupName: val_(P.groupName, i), branche: branche, sols: sols, props: props, open: open });
    if (open) e.open = true;
  }

  // 3. Sorties : lignes bloquantes et propositions (triées par programme), synthèse, clés ouvertes.
  const data = [], summary = [], openKeys = [];
  const coll = function (a, b) { return a.localeCompare(b, 'fr', { numeric: true }); };
  const solCols = function (nature, s, im) {
    return [nature, s.origin, s.id, s.name, s.type, s.status,
      im ? im.id : '', im ? im.type : '', im ? im.status : '', im ? im.resp : ''];
  };
  Array.from(progs.keys()).sort(coll).forEach(function (name) {
    const dm = progs.get(name);
    const start = data.length;
    let total = 0, nOpen = 0, nNew = 0, withProp = 0;
    const bloq = new Set(), her = new Set(), impls = new Set();
    Array.from(dm.values()).sort(function (a, b) { return coll(a.dossier, b.dossier); }).forEach(function (e) {
      total++;
      if (!e.open) return;
      nOpen++;
      const key = name + SEP + e.dossier;
      openKeys.push([key]);
      const isNew = !!opts.prev && !opts.prev.has(key);
      if (isNew) nNew++;
      const all = new Map();
      e.inst.forEach(function (ins) {
        ins.sols.forEach(function (s) { all.set(solKey_(s, ins), s); });
      });
      const solTotal = all.size;
      let solDone = 0;
      all.forEach(function (s) { if (s.closed) solDone++; });
      const tail = [isNew ? 'Oui' : '', String(solTotal), String(solDone), opts.week];
      const propSeen = new Set();
      e.inst.forEach(function (ins) {
        if (!ins.open) return;
        const base = [name, e.dossier, e.urlId, e.og, e.status, e.leader, e.supplier, ins.groupe, ins.groupName, ins.branche];
        const push = function (nature, s) {
          const list = s.impls.size ? Array.from(s.impls.values()) : [null];
          list.forEach(function (im) { data.push(base.concat(solCols(nature, s, im), tail)); });
        };
        if (!ins.sols.length) {
          data.push(base.concat(['Aucune solution', '', '', '', '', '', '', '', '', ''], tail));
        }
        ins.sols.forEach(function (s) {
          if (s.closed) return;
          const sk = e.dossier + SEP + solKey_(s, ins);
          bloq.add(sk);
          if (s.origin === 'Groupe') her.add(sk);
          s.impls.forEach(function (im) { impls.add(im.id); });
          push('Bloquante', s);
        });
        ins.props.forEach(function (s) {
          const pk = solKey_(s, ins);
          if (propSeen.has(pk)) return;
          propSeen.add(pk);
          push('Proposition', s);
        });
      });
      if (propSeen.size) withProp++;
    });
    const len = data.length - start;
    summary.push([name, total, nOpen, nNew, bloq.size, her.size, impls.size,
      total ? Math.round((total - nOpen) / total * 100) : 0,
      len ? start + 2 : '', len, opts.week, opts.extraction, withProp]);
  });
  return { data: data, summary: summary, openKeys: openKeys, excluded: excluded.size };
}

function normOg_(s) { return String(s || '').replace(/\s+/g, '').toUpperCase(); }

function solKey_(s, ins) {
  return (s.origin === 'Groupe' ? 'G' + ins.groupe : 'I' + ins.groupe + SEP + ins.branche) + SEP + s.id;
}

function writeOutputs_(res, week, isoDate) {
  writeTable_(CFG.SHEETS.data, DATA_HEADER, res.data, true);
  writeTable_(CFG.SHEETS.progs, PROG_HEADER, res.summary, false);
  const hist = readTable_(CFG.SHEETS.hist).filter(function (r) { return r[0] !== week; });
  res.summary.forEach(function (r) { hist.push([week, isoDate, r[0], r[2], r[4], r[6]]); });
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

/* ---------- Utilitaires feuilles ---------- */

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

/** Toujours en valeurs affichées (texte) pour éviter les conversions automatiques. */
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
  return t.getUTCFullYear() + '-S' + String(w).padStart(2, '0');
}

function isIsoDate_(s) { return /^\d{4}-\d{2}-\d{2}$/.test(String(s || '')); }

/** Normalise une date lue (texte ISO ou jj/mm/aaaa) en aaaa-mm-jj. */
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
