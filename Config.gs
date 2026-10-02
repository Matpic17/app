/**
 * Pilotage des clôtures — configuration.
 * Tout ce qui peut changer d'un déploiement à l'autre est ici.
 */
const CFG = {
  APP_NAME: 'Pilotage des programmes',
  LOGO_URL: 'https://www.brand.airbus.com/sites/g/files/jlcbta121/files/styles/w1800/public/2021-06/logo_blue.webp?itok=0vqEp8ge',

  // Fichier « programmes » : une ligne par dossier × groupe × branche.
  PROGRAMMES: {
    id: '1SBIPB9v2-2EOmWdWcVeqqunhumiK8pGGNSuIegJgm5A',
    sheet: 'Feuil1',
    firstRow: 2,
    cols: {
      urlId: 'A',          // ID du dossier pour le lien TOM (non affiché)
      dossier: 'B',        // numéro de dossier affiché
      dossierStatus: 'AP', // statut du dossier (filtre + affichage)
      og: 'AS',            // OG (filtre + affichage)
      supplier: 'AV',      // fournisseur
      leaderFirst: 'AY',   // case leader : prénom
      leaderLast: 'AZ',    // case leader : nom
      groupe: 'BB',        // ID groupe (jointure uniquement)
      groupName: 'BC',     // nom du groupe affiché
      branche: 'BG',       // ID branche (jointure uniquement)
      programme: 'BH'      // nom du programme
    }
  },

  // Fichier « solutions » : une ligne par dossier × solution × implémentation.
  SOLUTIONS: {
    id: '1F223VXpIsBNRlWpJoFMSYL1g5VtofwPP-BQwZrwCO84',
    sheet: 'Feuil1',
    firstRow: 2,
    cols: {
      dossier: 'B',        // jointure
      solType: 'BD',       // type de solution affiché
      solName: 'BF',       // nom de la solution affiché
      groupe: 'BM',        // ID groupe (jointure uniquement)
      branche: 'BQ',       // ID branche (jointure uniquement)
      solId: 'BT',         // ID solution (regroupement uniquement)
      solState: 'BU',      // Applied / Proposed / Rejected
      solStatus: 'BV',     // statut X/Y qui fait foi
      implId: 'CI',        // ID implémentation (regroupement uniquement)
      implType: 'CK',
      implResp: 'CN',      // référence côté responsable
      implStatus: 'CU'     // statut outil, information seulement
    }
  },

  // Filtres sur les dossiers (fichier programmes) : dossier exclu en entier.
  EXCLUDED_DOSSIER_STATUS: ['CANCEL', 'CLOSED', 'SOLVED'],
  EXCLUDED_OG: ['OG0', 'OG1', 'OG2', 'OG5'],

  // Colonne BU : seule « Applied » entre dans le calcul, « Proposed » est affichée pour info.
  SOLUTION_STATE: { applied: 'applied', proposed: 'proposed' },

  // Lien vers l'outil : {ID} est remplacé par la colonne A du dossier.
  TOOL_URL: 'https://tom.cr.eurocopter.corp/#/obso-case/{ID}/identification',

  // Une solution X/Y est soldée si X >= Y, sauf exceptions : { Y: seuil }.
  CLOSED_THRESHOLDS: { 9: 7 },

  // Dossier Drive partagé où ranger PDF et exports. Vide = créé automatiquement.
  DRIVE_FOLDER_ID: '1Ffxil6UIo1X0tDu5Jt0FpEPG3ENLT3gI',

  // Actualisation hebdomadaire automatique.
  REFRESH: { weekday: 'MONDAY', hour: 7 },

  // Onglets techniques (cachés) du fichier tableau de bord.
  SHEETS: {
    data: '_Donnees',
    progs: '_Programmes',
    hist: '_Historique',
    open: '_Ouverts',
    suivi: '_Suivi',
    envois: '_Envois'
  },

  MAX_RECIPIENTS: 20
};

// Une ligne de _Donnees = programme × dossier × branche × solution (bloquante ou proposée) × implémentation.
const DATA_HEADER = [
  'Programme', 'Dossier', 'IdOutil', 'OG', 'StatutDossier', 'CaseLeader', 'Fournisseur',
  'Groupe', 'NomGroupe', 'Branche',
  'Nature', 'Origine', 'Solution', 'NomSolution', 'TypeSolution', 'StatutSolution',
  'Implementation', 'TypeImplementation', 'StatutImplementation', 'RefResponsable',
  'Nouveau', 'SolutionsTotal', 'SolutionsSoldees', 'Semaine'
];
// Index des colonnes de _Donnees (évite les numéros en dur ailleurs).
const D = DATA_HEADER.reduce(function (o, h, i) { o[h] = i; return o; }, {});

const PROG_HEADER = [
  'Programme', 'DossiersTotal', 'DossiersOuverts', 'Nouveaux', 'SolutionsBloquantes',
  'SolutionsHeritees', 'Implementations', 'TauxCloture', 'LigneDebut', 'NbLignes', 'Semaine', 'Extraction',
  'DossiersAvecProposition'
];
const HIST_HEADER = ['Semaine', 'Date', 'Programme', 'DossiersOuverts', 'SolutionsBloquantes', 'Implementations'];
const SUIVI_HEADER = ['Programme', 'DateReunion', 'Dossier', 'Vu', 'Decision', 'Echeance', 'Auteur', 'MiseAJour'];
const ENVOI_HEADER = ['Horodatage', 'Auteur', 'Programme', 'DateReunion', 'Destinataires', 'Copie', 'Objet', 'LienPDF'];
