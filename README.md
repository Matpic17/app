# Programme Monitoring — Apps Script

Tableau de bord de pilotage des programmes, ouvert depuis le menu du fichier Google Sheets.
Interface, fiche PDF, e-mail et export en anglais ; onglets techniques cachés inchangés.

## Fichiers du projet

| Fichier (éditeur Apps Script) | Rôle |
|---|---|
| `appsscript.json` | Manifeste (services Sheets et People, autorisations) |
| `Config.gs` | Réglages : IDs des fichiers, colonnes, exclusions, lien TOM, dossier Drive |
| `Codes.gs` | Listes de codes : type d'implémentation (CL) manuel / TOM, statut (CU) ouvert / fermé |
| `Refresh.gs` | Calcul hebdomadaire et écriture des onglets cachés |
| `Server.gs` | Menu, API du tableau de bord, identification du case leader, PDF, Gmail, export |
| `Logic.html` | Logique partagée navigateur / serveur : indicateurs, sélection, ventilation |
| `Dashboard.html`, `App.html`, `Styles.html` | Interface |
| `FicheTemplate.html` | Fiche PDF |

`Codes` (sans extension) est la liste source fournie ; `Codes.gs` en est généré.

## Règles de calcul

- Cas exclus : statut AP = `CANCEL`, `CLOSED` ou `SOLVED`. Tous les OG sont conservés.
- Solutions : seules les « Applied » (BU) comptent pour la clôture ; les « Proposed » alimentent
  « awaiting decision » ; les « Rejected » sont ignorées.
- Une solution avec branche s'applique au programme de cette branche, sinon à tout le groupe ;
  la solution du programme prime sur celle du groupe.
- Une solution est soldée si son statut (BV) commence par X/Y avec X ≥ Y (ou X ≥ 7 sur 9).
- Les indicateurs « awaiting decision » et « applied » comptent des **solutions**, dédoublonnées par cas.

| Indicateur | Définition |
|---|---|
| Open cases | Cas où le programme est encore ouvert |
| Awaiting decision › Proposed | Solution proposée sur un groupe encore ouvert |
| Awaiting decision › Missing | Groupe ouvert sans aucune solution, ni appliquée ni proposée (compte 1 par groupe) |
| Applied blocking › No implementation | Aucune implémentation liée |
| Applied blocking › Implementation closed | Toutes les implémentations fermées (CU) : incohérence de données |
| Applied blocking › Manual implementation | Au moins une implémentation ouverte de type manuel (CL) |
| Applied blocking › Implementation in progress | Implémentations ouvertes, toutes suivies dans TOM |
| Applied blocking › Unclassified | Statut CU ou type CL inconnu : à ajouter dans `Codes.gs` |

Les codes sont comparés sans tenir compte des majuscules, accents, tabulations et espaces ;
un statut daté (« Go for delivery (12/03/26) ») est aussi comparé sans sa date.
Les valeurs inconnues trouvées lors de l'actualisation sont listées sur l'accueil.

## Case leaders

L'adresse e-mail de la personne connectée (`prénom.nom@…` ou `prénom.lettres.nom@…`) est comparée
aux colonnes AY (prénom) et AZ (nom). Un case leader ne voit que ses programmes et ses cas, sans
filtre case leader. Les autres voient tout, avec un filtre case leader sur l'accueil et dans la vue.
Ce n'est pas une barrière de sécurité : toute personne ayant accès au fichier peut ouvrir les onglets cachés.
Pour tester la vue d'un case leader : `CFG.VIEW_AS_EMAIL` dans `Config.gs` (à vider ensuite).

## Onglets cachés

`_Donnees` (une ligne par élément, colonne `Categorie`), `_Programmes`, `_Historique`, `_Ouverts`,
`_CaseLeaders`, `_Envois`. L'ancien onglet `_Suivi` est supprimé à la première actualisation.

## Mise à jour depuis la version précédente

1. Remplacer tous les fichiers, et créer `Codes.gs` et `Logic.html` (fichier HTML nommé `Logic`).
2. Reporter vos IDs dans `Config.gs` s'ils ont changé.
3. Lancer « Refresh data now » avant d'ouvrir le tableau de bord (la structure des données change).

## Application web et Google Sites

1. Renseigner `DASHBOARD_ID` dans `Config.gs` (ID du fichier tableau de bord).
2. Partager le fichier tableau de bord avec les utilisateurs : **lecteur** suffit pour consulter ;
   **éditeur** pour que l'envoi d'e-mail soit consigné dans `_Envois` (sinon l'envoi marche, sans journal).
3. Partager le dossier Drive des rapports en **éditeur** (sinon PDF et exports restent dans le Drive de l'utilisateur).
4. Déployer › Nouveau déploiement › Application web : *Exécuter en tant que* « l'utilisateur qui accède »,
   *Qui a accès* « tous les utilisateurs du domaine ». Lien direct : l'URL `/exec`.
5. Google Sites : Insérer › Intégrer › Par URL avec `…/exec?embed=1` (le paramètre affiche le bouton « Open full page »).
6. Nouvelle version : Gérer les déploiements › Modifier › Nouvelle version (l'URL ne change pas).
