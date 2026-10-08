/**
 * Programme monitoring — menu, API called by the dashboard, PDF report, Gmail and export.
 * api* functions are called from the browser through google.script.run.
 */

function onOpen() {
  SpreadsheetApp.getUi().createMenu(CFG.APP_NAME)
    .addItem('Open dashboard', 'openDashboard')
    .addSeparator()
    .addItem('Refresh data now', 'refreshAllFromMenu')
    .addItem('Schedule weekly refresh', 'installWeeklyTrigger')
    .addToUi();
}

function openDashboard() {
  const t = HtmlService.createTemplateFromFile('Dashboard');
  t.embed = false;
  t.appUrl = webAppUrl_();
  SpreadsheetApp.getUi().showModalDialog(t.evaluate().setWidth(1400).setHeight(880), CFG.APP_NAME);
}

/**
 * Web app entry point (link /exec, or embedded in Google Sites with /exec?embed=1).
 * Deployed as "Execute as: user accessing the web app".
 */
function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Dashboard');
  t.embed = !!(e && e.parameter && e.parameter.embed);
  t.appUrl = webAppUrl_();
  return t.evaluate()
    .setTitle(CFG.APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function webAppUrl_() {
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}

/** Saves a file in the shared reports folder, or in the user's own Drive if they cannot write there. */
function saveFile_(blob) {
  try { return reviewsFolder_().createFile(blob); } catch (e) { return DriveApp.createFile(blob); }
}

function include_(file) {
  return HtmlService.createHtmlOutputFromFile(file).getContent();
}

/** Shared logic (Logic.html), the same code as in the browser. */
let LOGIC_S_ = null;
function logic_() {
  if (!LOGIC_S_) LOGIC_S_ = new Function(include_('Logic') + '\nreturn LOGIC;')();
  return LOGIC_S_;
}

const int_ = function (v) { return parseInt(String(v).replace(/\s/g, ''), 10) || 0; };

/* ---------- Who is viewing ---------- */

/**
 * The e-mail address starts with firstname.lastname or firstname.letters.lastname:
 * it is matched against the case leader columns (first name AY, last name AZ).
 * This only adapts what is displayed; it is not a security barrier
 * (anyone with access to the spreadsheet can open the hidden tabs).
 */
function currentUser_() {
  const email = String(CFG.VIEW_AS_EMAIL || Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  const leaders = readTable_(CFG.SHEETS.leaders).map(function (r) {
    let progs = {};
    try { progs = JSON.parse(r[3] || '{}'); } catch (e) { progs = {}; }
    return { first: r[0], last: r[1], name: r[2], progs: progs };
  });
  const matches = matchLeaders_(email, leaders);
  const progs = {};
  matches.forEach(function (L) {
    Object.keys(L.progs).forEach(function (p) {
      const x = progs[p] || (progs[p] = { open: 0, total: 0 });
      x.open += L.progs[p].open; x.total += L.progs[p].total;
    });
  });
  return {
    email: email,
    isLeader: matches.length > 0,
    leaderNames: matches.map(function (L) { return L.name; }),
    progs: progs,
    leaders: leaders
  };
}

function matchLeaders_(email, leaders) {
  const local = String(email || '').split('@')[0];
  const tokens = local.split(/[._]+/).map(function (t) { return normName_(t); }).filter(Boolean);
  if (tokens.length < 2) return [];
  const first = tokens[0];
  const lastCandidates = [tokens[tokens.length - 1], tokens.slice(1).join(''), tokens.slice(2).join('')]
    .filter(Boolean);
  return leaders.filter(function (L) {
    return normName_(L.first) === first && lastCandidates.indexOf(normName_(L.last)) !== -1;
  });
}

/* ---------- Home ---------- */

function apiHome() {
  const user = currentUser_();
  const rows = readTable_(CFG.SHEETS.progs);
  let programmes = rows.map(function (r) { return { name: r[0], open: int_(r[2]) }; });
  if (user.isLeader) {
    programmes = programmes
      .filter(function (p) { return user.progs[p.name] && user.progs[p.name].open > 0; })
      .map(function (p) { return { name: p.name, open: user.progs[p.name].open }; });
  }
  const names = programmes.map(function (p) { return p.name; });
  return {
    appName: CFG.APP_NAME,
    user: { email: user.email, isLeader: user.isLeader, leaderNames: user.leaderNames },
    refresh: JSON.parse(PropertiesService.getScriptProperties().getProperty('LAST_REFRESH') || 'null'),
    programmes: programmes,
    leaders: user.isLeader ? [] : user.leaders
      .map(function (L) {
        const open = {};
        Object.keys(L.progs).forEach(function (p) { if (L.progs[p].open > 0) open[p] = L.progs[p].open; });
        return { name: L.name, progs: open };
      })
      .filter(function (L) { return Object.keys(L.progs).length; }),
    recent: getRecent_().filter(function (n) { return names.indexOf(n) !== -1; })
  };
}

function getRecent_() {
  return JSON.parse(PropertiesService.getUserProperties().getProperty('RECENT') || '[]');
}

function pushRecent_(name) {
  const list = getRecent_().filter(function (n) { return n !== name; });
  list.unshift(name);
  PropertiesService.getUserProperties().setProperty('RECENT', JSON.stringify(list.slice(0, 6)));
}

/* ---------- Programme view ---------- */

function apiProgramme(name) {
  const m = buildProgrammeModel_(String(name), currentUser_());
  pushRecent_(m.name);
  return m;
}

function buildProgrammeModel_(name, user) {
  const row = readTable_(CFG.SHEETS.progs).filter(function (r) { return r[0] === name; })[0];
  if (!row) throw new Error('Programme not found: ' + name);
  const start = int_(row[8]), len = int_(row[9]);
  const lines = len > 0
    ? sheet_(CFG.SHEETS.data).getRange(start, 1, len, DATA_HEADER.length).getDisplayValues()
    : [];
  const mine = user && user.isLeader ? user.leaderNames : null;

  const dmap = new Map(), closed = [];
  lines.forEach(function (l) {
    if (mine && mine.indexOf(l[D.CaseLeader]) === -1) return;
    const dossier = l[D.Dossier], groupe = l[D.Groupe], branche = l[D.Branche], origin = l[D.Origine];
    const nature = l[D.Nature], solId = l[D.Solution], implId = l[D.Implementation];
    if (nature === 'Dossier soldé') {
      closed.push({ id: dossier, og: l[D.OG], leader: l[D.CaseLeader], supplier: l[D.Fournisseur] });
      return;
    }
    let d = dmap.get(dossier);
    if (!d) {
      d = { id: dossier, url: toolUrl_(l[D.IdOutil] || dossier), isNew: l[D.Nouveau] === 'Oui',
        og: l[D.OG], status: l[D.StatutDossier], leader: l[D.CaseLeader], supplier: l[D.Fournisseur],
        groups: [], solTotal: int_(l[D.SolutionsTotal]), solDone: int_(l[D.SolutionsSoldees]), _g: {} };
      dmap.set(dossier, d);
    }
    let g = d._g[groupe];
    if (!g) {
      g = { name: l[D.NomGroupe] || groupe, blocking: [], proposals: [], done: [], noApplied: false, missing: false, _keys: {} };
      d._g[groupe] = g;
      d.groups.push(g);
    }
    if (nature === 'Aucune solution') {
      g.noApplied = true;
      if (l[D.Categorie] === CAT.MISSING) g.missing = true;
      return;
    }
    const sk = nature + SEP + (origin === 'Groupe' ? 'G' : 'I' + branche) + SEP + solId;
    let s = g._keys[sk];
    if (!s) {
      s = { key: solId, name: l[D.NomSolution], type: l[D.TypeSolution], status: l[D.StatutSolution],
        origin: origin, cat: l[D.Categorie], impls: [] };
      g._keys[sk] = s;
      (nature === 'Proposition' ? g.proposals : nature === 'Soldée' ? g.done : g.blocking).push(s);
    }
    if (implId && !s.impls.some(function (x) { return x.key === implId; })) {
      s.impls.push({ key: implId, code: l[D.CodeImplementation], label: l[D.TypeImplementation],
        status: l[D.StatutImplementation], resp: l[D.RefResponsable],
        open: l[D.ImplOuverte], manual: l[D.ImplManuelle] });
    }
  });

  const dossiers = Array.from(dmap.values()).map(function (d) {
    delete d._g;
    d.groups.forEach(function (g) { delete g._keys; });
    return d;
  });
  const all = dossiers.concat(closed);
  const uniq = function (k) {
    return all.map(function (d) { return d[k]; }).filter(function (v, i, a) { return v && a.indexOf(v) === i; })
      .sort(function (a, b) { return a.localeCompare(b, 'fr', { numeric: true }); });
  };

  // Weekly history of the programme (only meaningful without filters, and not for a case leader view).
  const history = mine ? [] : readTable_(CFG.SHEETS.hist)
    .filter(function (r) { return r[2] === name; })
    .map(function (r) {
      const has = String(r[6] || '') !== '';
      return {
        week: normWeek_(r[0]), open: int_(r[3]), applied: int_(r[4]),
        decision: has ? int_(r[6]) + int_(r[7]) : null,
        cats: has ? { NO_IMPL: int_(r[8]), IMPL_CLOSED: int_(r[9]), IN_PROGRESS: int_(r[10]), MANUAL: int_(r[11]), UNCLASSIFIED: int_(r[12]),
          PROPOSED: int_(r[6]), MISSING: int_(r[7]) } : null
      };
    })
    .sort(function (a, b) { return a.week.localeCompare(b.week); }).slice(-8);

  const sends = readTable_(CFG.SHEETS.envois).filter(function (r) { return r[2] === name; });
  let lastTo = [];
  if (sends.length) { try { lastTo = JSON.parse(sends[sends.length - 1][4]) || []; } catch (e) { lastTo = []; } }

  return {
    name: name, week: normWeek_(row[10]), extraction: row[11], today: today_(),
    dossiers: dossiers, closed: closed, history: history,
    filters: { ogs: uniq('og'), leaders: mine ? [] : uniq('leader'), suppliers: uniq('supplier') },
    leaderView: mine ? mine.join(', ') : '',
    lastTo: lastTo, toolName: CFG.TOOL_NAME
  };
}

function toolUrl_(id) {
  return CFG.TOOL_URL.replace('{ID}', encodeURIComponent(String(id)));
}

/** Programme model seen through the dashboard filters and the current indicator selection. */
function scopedModel_(name, f, sel) {
  const L = logic_();
  const user = currentUser_();
  const m = buildProgrammeModel_(String(name), user);
  f = sanitizeFilters_(f, user);
  sel = sanitizeSelection_(sel);
  const sc = L.scope(m, f, sel);
  return {
    name: m.name, week: m.week, extraction: m.extraction, today: m.today, toolName: m.toolName,
    filterLabel: L.filterLabel(f) || (m.leaderView ? 'Case leader: ' + m.leaderView : ''),
    selectionLabel: L.selectionLabel(sel),
    sel: sel, kpi: sc.kpi, breakdown: L.breakdown(sc.filtered, sel),
    dossiers: sc.selected.map(function (d) { return Object.assign({}, d, { view: L.annotate(d, sel) }); }),
    defs: { KPI: L.KPI, SUB: L.SUB }
  };
}

function sanitizeFilters_(f, user) {
  f = f || {};
  return {
    ogs: Array.isArray(f.ogs) ? f.ogs.map(String) : [],
    leader: user && user.isLeader ? '' : String(f.leader || ''),
    supplier: String(f.supplier || '')
  };
}

function sanitizeSelection_(sel) {
  sel = sel || {};
  const kpi = ['open', 'decision', 'applied'].indexOf(sel.kpi) !== -1 ? sel.kpi : 'open';
  return { kpi: kpi, sub: kpi === 'open' ? '' : String(sel.sub || ''), type: String(sel.type || '') };
}

/* ---------- Company directory ---------- */

function apiSearchPeople(q) {
  q = String(q || '').trim();
  if (q.length < 2) return [];
  const res = People.People.searchDirectoryPeople({
    query: q,
    readMask: 'names,emailAddresses,organizations',
    sources: ['DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE'],
    pageSize: 8
  });
  return (res.people || []).map(function (p) {
    const org = (p.organizations || [])[0] || {};
    return {
      name: ((p.names || [])[0] || {}).displayName || '',
      email: ((p.emailAddresses || [])[0] || {}).value || '',
      title: org.title || org.department || ''
    };
  }).filter(function (p) { return p.email; });
}

/* ---------- PDF report ---------- */

function renderFiche_(m) {
  const t = HtmlService.createTemplateFromFile('FicheTemplate');
  t.m = m;
  t.fmt = fmtDate_;
  t.logo = logoDataUri_();
  return t.evaluate().getContent();
}

function logoDataUri_() {
  if (!CFG.LOGO_URL) return '';
  const cache = CacheService.getScriptCache();
  const hit = cache.get('LOGO_URI');
  if (hit) return hit;
  try {
    const blob = UrlFetchApp.fetch(CFG.LOGO_URL).getBlob();
    const uri = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
    if (uri.length < 100000) cache.put('LOGO_URI', uri, 21600);
    return uri;
  } catch (e) { return ''; }
}

function ficheName_(m) {
  return m.name.replace(/[^\w.-]+/g, '-') + '_status_' + m.today + '.pdf';
}

function makePdf_(m) {
  return Utilities.newBlob(renderFiche_(m), 'text/html', 'report.html').getAs('application/pdf').setName(ficheName_(m));
}

function apiFicheHtml(name, f, sel) {
  return renderFiche_(scopedModel_(name, f, sel));
}

function apiPdfToDrive(name, f, sel) {
  const file = saveFile_(makePdf_(scopedModel_(name, f, sel)));
  return { url: file.getUrl(), name: file.getName() };
}

/* ---------- Gmail ---------- */

function apiSendFiche(p) {
  const isMail = function (e) { return /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(String(e || '')); };
  const to = (p.to || []).map(function (x) { return String(x.email).trim().toLowerCase(); }).filter(isMail);
  const ccIn = (p.cc || []).map(function (x) { return String(x.email).trim().toLowerCase(); }).filter(isMail);
  if (!to.length) throw new Error('Add at least one recipient.');
  if (to.length + ccIn.length > CFG.MAX_RECIPIENTS) throw new Error('Too many recipients (' + CFG.MAX_RECIPIENTS + ' maximum).');
  const subject = String(p.subject || '').trim().slice(0, 250);
  if (!subject) throw new Error('The subject is empty.');
  const body = String(p.body || '').slice(0, 10000);

  const me = Session.getActiveUser().getEmail().toLowerCase();
  const cc = Array.from(new Set([me].concat(ccIn))).filter(function (e) { return e && to.indexOf(e) === -1; });
  const m = scopedModel_(p.programme, p.filters, p.selection);
  const pdf = makePdf_(m);

  GmailApp.sendEmail(to.join(','), subject, body, { cc: cc.join(','), attachments: [pdf], name: CFG.APP_NAME });

  let link = '';
  if (p.saveDrive) { try { link = saveFile_(pdf).getUrl(); } catch (e) { link = ''; } }
  if (p.log) {
    // Needs edit access to the dashboard spreadsheet: skipped silently for read-only users.
    try {
      const sh = ensureSheet_(CFG.SHEETS.envois, ENVOI_HEADER);
      sh.appendRow([
        Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm'), me, m.name, m.today,
        JSON.stringify((p.to || []).map(function (x) { return { name: String(x.name || ''), email: String(x.email) }; })),
        cc.join(', '), subject, link
      ]);
    } catch (e) { console.warn('E-mail not logged: ' + e); }
  }
  return { to: to, cc: cc, link: link };
}

/* ---------- Export (Google Sheets): exactly the items of the current selection ---------- */

function apiExport(name, f, sel) {
  const m = scopedModel_(name, f, sel);
  const KIND = { blocking: 'Applied, blocking', proposal: 'Proposed', missing: 'Missing', closed: 'Applied, closed' };
  const header = ['Programme', 'Case', 'New', 'OG', 'Case status', 'Case leader', 'Supplier', 'Group',
    'Item', 'Category', 'Solution name', 'Solution type', 'Solution status', 'Origin',
    'Implementation type (CL)', 'Implementation label (CK)', 'Owner reference', 'Implementation status',
    'Implementation open', 'Link to ' + CFG.TOOL_NAME];
  const yn = function (v) { return v === 'Oui' ? 'Yes' : v === 'Non' ? 'No' : v === '?' ? 'Unknown' : ''; };
  const rows = [];
  m.dossiers.forEach(function (d) {
    d.view.groups.forEach(function (g) {
      g.main.forEach(function (it) {
        const base = [m.name, d.id, d.isNew ? 'Yes' : '', d.og, d.status, d.leader, d.supplier, g.name,
          KIND[it.kind], logic_().CAT_TITLE[it.cat] || ''];
        const s = it.sol;
        if (!s) { rows.push(base.concat(['', '', '', '', '', '', '', '', '', d.url])); return; }
        const sb = [s.name, s.type, s.status, s.origin === 'Groupe' ? 'Group' : 'Programme'];
        const list = s.impls.length ? s.impls : [null];
        list.forEach(function (im) {
          rows.push(base.concat(sb, im ? [im.code, im.label, im.resp, im.status, yn(im.open)] : ['', '', '', '', ''], [d.url]));
        });
      });
    });
  });
  const title = m.name + ' – ' + m.selectionLabel + (m.filterLabel ? ' (' + m.filterLabel + ')' : '') + ' – ' + fmtDate_(m.today);
  const ss = SpreadsheetApp.create(title.slice(0, 250));
  const sh = ss.getSheets()[0].setName('Export');
  sh.getRange(1, 1, rows.length + 1, header.length).setNumberFormat('@');
  sh.getRange(1, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#00205B');
  if (rows.length) sh.getRange(2, 1, rows.length, header.length).setValues(rows);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, header.length);
  try { DriveApp.getFileById(ss.getId()).moveTo(reviewsFolder_()); } catch (e) { /* stays in the user's Drive */ }
  return { url: ss.getUrl(), rows: rows.length };
}

function reviewsFolder_() {
  if (CFG.DRIVE_FOLDER_ID) return DriveApp.getFolderById(CFG.DRIVE_FOLDER_ID);
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('REVIEWS_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* recreated below */ } }
  const folder = DriveApp.createFolder(CFG.APP_NAME + ' – reports');
  props.setProperty('REVIEWS_FOLDER_ID', folder.getId());
  return folder;
}
