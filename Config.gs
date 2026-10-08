/**
 * Programme monitoring — configuration.
 * Everything that may change from one deployment to another lives here.
 * (Hidden technical tabs keep their original French names.)
 */
const CFG = {
  APP_NAME: 'Programme Monitoring',

  // ID of THIS dashboard spreadsheet (the part of its URL between /d/ and /edit).
  // Required for the web app (link / Google Sites): there is no "open spreadsheet" there.
  DASHBOARD_ID: '',
  LOGO_URL: 'https://www.brand.airbus.com/sites/g/files/jlcbta121/files/styles/w1800/public/2021-06/logo_blue.webp?itok=0vqEp8ge',

  // "Programmes" extract: one row per case × group × branch.
  PROGRAMMES: {
    id: '1SBIPB9v2-2EOmWdWcVeqqunhumiK8pGGNSuIegJgm5A',
    sheet: 'Feuil1',
    firstRow: 2,
    cols: {
      urlId: 'A',          // case ID used in the TOM link (not displayed)
      dossier: 'B',        // case number displayed
      dossierStatus: 'AP', // case status (exclusion + display)
      og: 'AS',            // OG (display + filter)
      supplier: 'AV',      // supplier
      leaderFirst: 'AY',   // case leader: first name
      leaderLast: 'AZ',    // case leader: last name
      groupe: 'BB',        // group ID (join only)
      groupName: 'BC',     // group name displayed
      branche: 'BG',       // branch ID (join only)
      programme: 'BH'      // programme name
    }
  },

  // "Solutions" extract: one row per case × solution × implementation.
  SOLUTIONS: {
    id: '1F223VXpIsBNRlWpJoFMSYL1g5VtofwPP-BQwZrwCO84',
    sheet: 'Feuil1',
    firstRow: 2,
    cols: {
      dossier: 'B',        // join
      solType: 'BD',       // solution type displayed
      solName: 'BF',       // solution name displayed
      groupe: 'BM',        // group ID (join only)
      branche: 'BQ',       // branch ID (join only)
      solId: 'BT',         // solution ID (grouping only)
      solState: 'BU',      // Applied / Proposed / Rejected
      solStatus: 'BV',     // X/Y status, the reference for closure
      implId: 'CI',        // implementation ID (grouping only)
      implLabel: 'CK',     // implementation label (information)
      implCode: 'CL',      // implementation type: manual or tracked in TOM (see Codes.gs)
      implResp: 'CN',      // owner-side reference
      implStatus: 'CU'     // implementation status: open or closed (see Codes.gs)
    }
  },

  // Cases excluded entirely (programmes extract, column AP). All OGs are kept.
  EXCLUDED_DOSSIER_STATUS: ['CANCEL', 'CLOSED', 'SOLVED'],

  // Column BU: only "Applied" counts for closure, "Proposed" feeds "awaiting decision".
  SOLUTION_STATE: { applied: 'applied', proposed: 'proposed' },

  // Link to the tool: {ID} is replaced by column A of the case.
  TOOL_URL: 'https://tom.cr.eurocopter.corp/#/obso-case/{ID}/identification',
  TOOL_NAME: 'TOM',

  // An X/Y solution is closed if X >= Y, except: { Y: threshold }.
  CLOSED_THRESHOLDS: { 9: 7 },

  // Shared Drive folder for PDFs and exports. Empty = created automatically.
  DRIVE_FOLDER_ID: '1Ffxil6UIo1X0tDu5Jt0FpEPG3ENLT3gI',

  // Weekly automatic refresh.
  REFRESH: { weekday: 'MONDAY', hour: 7 },

  // Testing only: see the dashboard as this e-mail address would (e.g. a case leader).
  // Leave empty in production.
  VIEW_AS_EMAIL: '',

  // Hidden technical tabs of the dashboard spreadsheet.
  SHEETS: {
    data: '_Donnees',
    progs: '_Programmes',
    hist: '_Historique',
    open: '_Ouverts',
    leaders: '_CaseLeaders',
    envois: '_Envois',
    legacy: ['_Suivi']    // removed at the next refresh
  },

  MAX_RECIPIENTS: 20
};

// Solution categories (column "Categorie" of _Donnees).
const CAT = {
  NO_IMPL: 'NO_IMPL',             // applied, blocking, no implementation
  IMPL_CLOSED: 'IMPL_CLOSED',     // applied, blocking, all implementations closed (data inconsistency)
  IN_PROGRESS: 'IN_PROGRESS',     // applied, blocking, open implementation(s) tracked in TOM
  MANUAL: 'MANUAL',               // applied, blocking, at least one open manual implementation
  UNCLASSIFIED: 'UNCLASSIFIED',   // applied, blocking, unknown implementation code or status
  PROPOSED: 'PROPOSED',           // proposed solution on an open group
  MISSING: 'MISSING',             // open group with neither applied nor proposed solution
  CLOSED: 'CLOSED'                // applied solution already closed (X/Y)
};

// One row of _Donnees = programme × case × branch × item × implementation.
// Nature: Bloquante | Proposition | Soldée | Aucune solution | Dossier soldé
const DATA_HEADER = [
  'Programme', 'Dossier', 'IdOutil', 'OG', 'StatutDossier', 'CaseLeader', 'Fournisseur',
  'Groupe', 'NomGroupe', 'Branche',
  'Nature', 'Origine', 'Solution', 'NomSolution', 'TypeSolution', 'StatutSolution',
  'Implementation', 'TypeImplementation', 'StatutImplementation', 'RefResponsable',
  'Nouveau', 'SolutionsTotal', 'SolutionsSoldees', 'Semaine',
  'Categorie', 'CodeImplementation', 'ImplOuverte', 'ImplManuelle'
];
// Column index of _Donnees by name.
const D = DATA_HEADER.reduce(function (o, h, i) { o[h] = i; return o; }, {});

const PROG_HEADER = [
  'Programme', 'DossiersTotal', 'DossiersOuverts', 'Nouveaux', 'SolutionsBloquantes',
  'SolutionsHeritees', 'Implementations', 'TauxCloture', 'LigneDebut', 'NbLignes', 'Semaine', 'Extraction',
  'DossiersAvecProposition',
  'Proposees', 'Manquantes', 'SansImpl', 'ImplFermee', 'EnCours', 'Manuelles', 'NonClassees'
];
const HIST_HEADER = [
  'Semaine', 'Date', 'Programme', 'DossiersOuverts', 'SolutionsBloquantes', 'Implementations',
  'Proposees', 'Manquantes', 'SansImpl', 'ImplFermee', 'EnCours', 'Manuelles', 'NonClassees'
];
const LEADER_HEADER = ['Prenom', 'Nom', 'CaseLeader', 'Programmes'];
const ENVOI_HEADER = ['Horodatage', 'Auteur', 'Programme', 'DateReunion', 'Destinataires', 'Copie', 'Objet', 'LienPDF'];
