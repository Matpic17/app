/**
 * Pilotage des clôtures — menu, API appelée par l'interface, fiche PDF, envoi Gmail.
 * Les fonctions api* sont appelées depuis le navigateur via google.script.run.
 */

function onOpen() {
  SpreadsheetApp.getUi().createMenu(CFG.APP_NAME)
    .addItem('Ouvrir le tableau de bord', 'openDashboard')
    .addSeparator()
    .addItem('Actualiser les données maintenant', 'refreshAllFromMenu')
    .addItem("Programmer l'actualisation hebdomadaire", 'installWeeklyTrigger')
    .addToUi();
}

function openDashboard() {
  const html = HtmlService.createTemplateFromFile('Dashboard').evaluate()
    .setWidth(1400).setHeight(880);
  SpreadsheetApp.getUi().showModalDialog(html, CFG.APP_NAME);
}

function include_(file) {
  return HtmlService.createHtmlOutputFromFile(file).getContent();
}

/* ---------- Accueil ---------- */

function apiHome() {
  const progs = readTable_(CFG.SHEETS.progs).map(function (r) { return r[0]; });
  const reviews = lastReviewByProgramme_();
  const recent = getRecent_().filter(function (n) { return progs.indexOf(n) !== -1; });
  return {
    me: Session.getActiveUser().getEmail(),
    refresh: JSON.parse(PropertiesService.getScriptProperties().getProperty('LAST_REFRESH') || 'null'),
    programmes: progs,
    recent: recent.map(function (n) { return { name: n, lastReview: reviews[n] || '' }; })
  };
}

function lastReviewByProgramme_() {
  const out = {};
  readTable_(CFG.SHEETS.suivi).forEach(function (r) {
    const d = dateStr_(r[1]);
    if (!out[r[0]] || d > out[r[0]]) out[r[0]] = d;
  });
  return out;
}

function getRecent_() {
  return JSON.parse(PropertiesService.getUserProperties().getProperty('RECENT') || '[]');
}

function pushRecent_(name) {
  const list = getRecent_().filter(function (n) { return n !== name; });
  list.unshift(name);
  PropertiesService.getUserProperties().setProperty('RECENT', JSON.stringify(list.slice(0, 6)));
}

/* ---------- Vue programme ---------- */

function apiProgramme(name, meeting) {
  const m = buildProgrammeModel_(String(name), meeting);
  pushRecent_(m.name);
  return m;
}

function buildProgrammeModel_(name, meeting) {
  meeting = isIsoDate_(meeting) ? meeting : today_();
  const row = readTable_(CFG.SHEETS.progs).filter(function (r) { return r[0] === name; })[0];
  if (!row) throw new Error('Programme introuvable : ' + name);
  const int = function (v) { return parseInt(String(v).replace(/\s/g, ''), 10) || 0; };
  const start = int(row[8]), len = int(row[9]);

  const lines = len > 0
    ? sheet_(CFG.SHEETS.data).getRange(start, 1, len, DATA_HEADER.length).getDisplayValues()
    : [];

  const dmap = new Map();
  lines.forEach(function (l) {
    const dossier = l[D.Dossier], groupe = l[D.Groupe], branche = l[D.Branche], origin = l[D.Origine];
    const nature = l[D.Nature], solId = l[D.Solution], implId = l[D.Implementation];
    let d = dmap.get(dossier);
    if (!d) {
      d = { id: dossier, url: toolUrl_(l[D.IdOutil] || dossier), isNew: l[D.Nouveau] === 'Oui',
        og: l[D.OG], status: l[D.StatutDossier], leader: l[D.CaseLeader], supplier: l[D.Fournisseur],
        groups: [], sols: [], proposals: [], noSolution: false,
        solTotal: int(l[D.SolutionsTotal]), solDone: int(l[D.SolutionsSoldees]), _keys: {} };
      dmap.set(dossier, d);
    }
    const gName = l[D.NomGroupe] || groupe;
    if (d.groups.indexOf(gName) === -1) d.groups.push(gName);
    if (nature === 'Aucune solution') { d.noSolution = true; return; }
    const sk = nature + SEP + (origin === 'Groupe' ? 'G' + groupe : 'I' + groupe + SEP + branche) + SEP + solId;
    let s = d._keys[sk];
    if (!s) {
      s = { key: solId, name: l[D.NomSolution], type: l[D.TypeSolution], status: l[D.StatutSolution],
        origin: origin, group: gName, impls: [] };
      d._keys[sk] = s;
      (nature === 'Proposition' ? d.proposals : d.sols).push(s);
    }
    if (implId && !s.impls.some(function (x) { return x.key === implId; })) {
      s.impls.push({ key: implId, type: l[D.TypeImplementation], status: l[D.StatutImplementation], resp: l[D.RefResponsable] });
    }
  });

  const dossiers = Array.from(dmap.values()).map(function (d) { delete d._keys; return d; })
    .sort(function (a, b) { return (b.sols.length - a.sols.length) || a.id.localeCompare(b.id, 'fr', { numeric: true }); });
  const impls = aggregateImpls_(dossiers);
  const uniq = function (k) {
    return dossiers.map(function (d) { return d[k]; }).filter(function (v, i, a) { return v && a.indexOf(v) === i; })
      .sort(function (a, b) { return a.localeCompare(b, 'fr'); });
  };

  // Suivi de réunion
  const today = today_();
  const suivi = readTable_(CFG.SHEETS.suivi).filter(function (r) { return r[0] === name; }).map(function (r) {
    return { date: dateStr_(r[1]), dossier: r[2], vu: r[3] === 'Oui', decision: r[4] || '', due: dateStr_(r[5]), author: r[6] || '' };
  });
  const current = {};
  suivi.filter(function (s) { return s.date === meeting; }).forEach(function (s) {
    current[s.dossier] = { vu: s.vu, decision: s.decision, due: s.due };
  });
  const mmap = {};
  suivi.forEach(function (s) {
    const x = mmap[s.date] || (mmap[s.date] = { date: s.date, seen: 0, decisions: 0 });
    if (s.vu) x.seen++;
    if (s.decision) x.decisions++;
  });
  if (!mmap[meeting]) mmap[meeting] = { date: meeting, seen: 0, decisions: 0 };
  const meetings = Object.keys(mmap).sort().reverse().map(function (k) { return mmap[k]; });
  const decisions = suivi.filter(function (s) { return s.decision; })
    .sort(function (a, b) { return b.date.localeCompare(a.date); })
    .map(function (s) {
      const open = dmap.has(s.dossier);
      return { dossier: s.dossier, decision: s.decision, due: s.due, meeting: s.date, author: s.author,
        open: open, late: open && isIsoDate_(s.due) && s.due < today };
    });

  // Historique et dernier envoi
  const history = readTable_(CFG.SHEETS.hist).filter(function (r) { return r[2] === name; })
    .map(function (r) { return { week: r[0], open: int(r[3]) }; })
    .sort(function (a, b) { return a.week.localeCompare(b.week); }).slice(-8);
  const sends = readTable_(CFG.SHEETS.envois).filter(function (r) { return r[2] === name; });
  let lastTo = [];
  if (sends.length) { try { lastTo = JSON.parse(sends[sends.length - 1][4]) || []; } catch (e) { lastTo = []; } }

  const seen = dossiers.filter(function (d) { return current[d.id] && current[d.id].vu; }).length;
  const prevOpen = history.length > 1 ? history[history.length - 2].open : null;

  return {
    name: name, week: row[10], extraction: row[11], meeting: meeting, today: today,
    kpi: {
      total: int(row[1]), open: int(row[2]), nNew: int(row[3]), sols: int(row[4]), solsGroup: int(row[5]),
      impls: int(row[6]), rate: int(row[7]), seen: seen, withProp: int(row[12]),
      delta: prevOpen == null ? null : int(row[2]) - prevOpen
    },
    history: history, dossiers: dossiers, impls: impls,
    filters: { leaders: uniq('leader'), suppliers: uniq('supplier') },
    current: current, meetings: meetings, decisions: decisions, lastTo: lastTo,
    me: Session.getActiveUser().getEmail()
  };
}

/** Implémentations des solutions bloquantes, triées par nombre de dossiers débloqués. */
function aggregateImpls_(dossiers) {
  const map = new Map();
  dossiers.forEach(function (d) {
    d.sols.forEach(function (s) {
      s.impls.forEach(function (im) {
        let a = map.get(im.key);
        if (!a) { a = { key: im.key, type: im.type, status: im.status, resp: im.resp, sols: [], dossiers: [] }; map.set(im.key, a); }
        const label = s.name || s.type;
        if (label && a.sols.indexOf(label) === -1) a.sols.push(label);
        if (a.dossiers.indexOf(d.id) === -1) a.dossiers.push(d.id);
      });
    });
  });
  return Array.from(map.values()).sort(function (a, b) {
    return (b.dossiers.length - a.dossiers.length) || String(a.type).localeCompare(String(b.type), 'fr');
  });
}

function toolUrl_(id) {
  return CFG.TOOL_URL.replace('{ID}', encodeURIComponent(String(id)));
}

/* ---------- Suivi de réunion ---------- */

function apiSaveSuivi(p) {
  if (!p || !p.programme || !p.dossier || !isIsoDate_(p.date)) throw new Error('Données de suivi incomplètes.');
  if (p.due && !isIsoDate_(p.due)) throw new Error("Date d'échéance invalide.");
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const sh = ensureSheet_(CFG.SHEETS.suivi, SUIVI_HEADER);
    const last = sh.getLastRow();
    let target = -1;
    if (last >= 2) {
      const keys = sh.getRange(2, 1, last - 1, 3).getDisplayValues();
      for (let i = 0; i < keys.length; i++) {
        if (keys[i][0] === p.programme && dateStr_(keys[i][1]) === p.date && keys[i][2] === p.dossier) { target = i + 2; break; }
      }
    }
    if (target < 0) {
      target = last + 1;
      if (target > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 200);
    }
    const row = sh.getRange(target, 1, 1, SUIVI_HEADER.length);
    row.setNumberFormat('@');
    row.setValues([[
      String(p.programme), p.date, String(p.dossier), p.vu ? 'Oui' : '',
      String(p.decision || '').slice(0, 1000), p.due || '',
      Session.getActiveUser().getEmail(), Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm')
    ]]);
  } finally {
    lock.releaseLock();
  }
  return true;
}

/* ---------- Annuaire ---------- */

function apiSearchPeople(q) {
  q = String(q || '').trim();t.logo
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

/* ---------- Fiche PDF ---------- */

function renderFiche_(m) {
  const t = HtmlService.createTemplateFromFile('FicheTemplate');
  t.m = m;
  t.fmt = fmtDate_;
  t.logo = logoDataUri_();
  return t.evaluate().getContent();
}

function logoDataUri_() {
  if(!CFG.LOGO_URL) return '';
  const cache = CacheService.getScriptCache();
  const hit = cache.get('LOGO_URI');
  if(hit) return hit;
  try {
    const blob = UrlFetchApp.fetch(CFG.LOGO_URL).getBlob();
    const uri = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
    if(uri.length < 100000) cache.put('LOGO_URI', uri, 21600);
    return uri;
  } catch (e) {return ''; }
}

function ficheName_(m) {
  return 'Revue_' + m.name.replace(/[^\w.-]+/g, '-') + '_' + m.meeting + '.pdf';
}

function makePdf_(m) {
  return Utilities.newBlob(renderFiche_(m), 'text/html', 'fiche.html').getAs('application/pdf').setName(ficheName_(m));
}

function apiFicheHtml(name, meeting) {
  return renderFiche_(buildProgrammeModel_(String(name), meeting));
}

function apiPdfToDrive(name, meeting) {
  const file = reviewsFolder_().createFile(makePdf_(buildProgrammeModel_(String(name), meeting)));
  return { url: file.getUrl(), name: file.getName() };
}

/* ---------- Envoi Gmail ---------- */

function apiSendFiche(p) {
  const isMail = function (e) { return /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(String(e || '')); };
  const to = (p.to || []).map(function (x) { return String(x.email).trim().toLowerCase(); }).filter(isMail);
  const ccIn = (p.cc || []).map(function (x) { return String(x.email).trim().toLowerCase(); }).filter(isMail);
  if (!to.length) throw new Error('Ajoutez au moins un destinataire.');
  if (to.length + ccIn.length > CFG.MAX_RECIPIENTS) throw new Error('Trop de destinataires (' + CFG.MAX_RECIPIENTS + ' au maximum).');
  const subject = String(p.subject || '').trim().slice(0, 250);
  if (!subject) throw new Error("L'objet est vide.");
  const body = String(p.body || '').slice(0, 10000);

  const me = Session.getActiveUser().getEmail().toLowerCase();
  const cc = Array.from(new Set([me].concat(ccIn))).filter(function (e) { return e && to.indexOf(e) === -1; });
  const m = buildProgrammeModel_(String(p.programme), p.meeting);
  const pdf = makePdf_(m);

  GmailApp.sendEmail(to.join(','), subject, body, { cc: cc.join(','), attachments: [pdf], name: CFG.APP_NAME });

  let link = '';
  if (p.saveDrive) link = reviewsFolder_().createFile(pdf).getUrl();
  if (p.log) {
    const sh = ensureSheet_(CFG.SHEETS.envois, ENVOI_HEADER);
    sh.appendRow([
      Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm'), me, m.name, m.meeting,
      JSON.stringify((p.to || []).map(function (x) { return { name: String(x.name || ''), email: String(x.email) }; })),
      cc.join(', '), subject, link
    ]);
  }
  return { to: to, cc: cc, link: link };
}

/* ---------- Export du suivi (Google Sheets) ---------- */

function apiExportSuivi(name, meeting) {
  const m = buildProgrammeModel_(String(name), meeting);
  const header = ['Programme', 'Dossier', 'OG', 'Statut dossier', 'Case leader', 'Fournisseur', 'Groupes', 'Nouveau',
    'Nature', 'Nom solution', 'Type solution', 'Statut solution', 'Origine',
    'Type implémentation', 'Réf. responsable', 'Statut outil (info)',
    'Vu le ' + fmtDate_(m.meeting), 'Décision', 'Échéance', 'Lien outil'];
  const rows = [];
  m.dossiers.forEach(function (d) {
    const c = m.current[d.id] || {};
    const base = [m.name, d.id, d.og, d.status, d.leader, d.supplier, d.groups.join(' | '), d.isNew ? 'Oui' : ''];
    const end = [c.vu ? 'Oui' : '', c.decision || '', c.due ? fmtDate_(c.due) : '', d.url];
    if (d.noSolution) rows.push(base.concat(['Aucune solution appliquée', '', '', '', '', '', '', ''], end));
    const add = function (nature, s) {
      const list = s.impls.length ? s.impls : [null];
      list.forEach(function (im) {
        rows.push(base.concat([nature, s.name, s.type, s.status, s.origin,
          im ? im.type : '', im ? im.resp : '', im ? im.status : ''], end));
      });
    };
    d.sols.forEach(function (s) { add('Bloquante', s); });
    d.proposals.forEach(function (s) { add('Proposée (info)', s); });
  });
  const ss = SpreadsheetApp.create('Suivi ' + m.name + ' – revue du ' + fmtDate_(m.meeting));
  const sh = ss.getSheets()[0].setName('Suivi');
  sh.getRange(1, 1, rows.length + 1, header.length).setNumberFormat('@');
  sh.getRange(1, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#00205B');
  if (rows.length) sh.getRange(2, 1, rows.length, header.length).setValues(rows);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, header.length);
  DriveApp.getFileById(ss.getId()).moveTo(reviewsFolder_());
  return { url: ss.getUrl() };
}

function reviewsFolder_() {
  if (CFG.DRIVE_FOLDER_ID) return DriveApp.getFolderById(CFG.DRIVE_FOLDER_ID);
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('REVIEWS_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* recréé ci-dessous */ } }
  const folder = DriveApp.createFolder(CFG.APP_NAME + ' – revues');
  props.setProperty('REVIEWS_FOLDER_ID', folder.getId());
  return folder;
}
