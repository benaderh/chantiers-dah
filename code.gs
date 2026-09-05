/**
 * SAISIE CHANTIER MOBILE — Google Sheets
 * Application de saisie mobile avec thème sombre et menus déroulants filtrés.
 *
 * FONCTIONNEMENT :
 * - Chaque feuille = un chantier (sauf la feuille cachée "Listes")
 * - La feuille "Listes" (cachée) contient les libellés et leur chantier_id
 * - Le menu déroulant "libellé" est filtré par chantier
 * - Libellés sans chantier_id = communs à tous les chantiers
 * - Thème sombre optimisé pour mobile
 * - Seules les colonnes A-F sont visibles
 * - 10 lignes vides après la dernière saisie
 *
 * INSTALLATION :
 * 1. Extensions > Apps Script > coller ce code > Exécuter "configurerTout"
 * 2. Menu Chantiers > "Modifier les libellés" pour accéder à la feuille Listes
 *
 * AJOUT CHANTIER : créer la feuille + relancer "configurerTout"
 */

const HEADERS = ['date', 'libelle', 'charges', 'produits', 'reglement', 'obs'];
const FEUILLE_LISTES = 'Listes';
const NB_COLS = 6;

// ─── Thème sombre ───
const THEME = {
  headerBg:   '#0f3460',
  headerText: '#ffffff',
  rowDark:    '#1a1a2e',
  rowDarkAlt: '#16213e',
  text:       '#e0e0e0',
  textMuted:  '#8888aa',
  accent:     '#e94560',
  border:     '#2a2a4a'
};

// Libellés de départ (chantier_id vide = communs à tous)
const LIBELLES_DEFAUT = [
  'Achat matériaux', 'Location engin', 'Carburant', 'Main d\'œuvre',
  'Transport', 'Paiement client', 'Avance client', 'Facture fournisseur',
  'Frais divers'
];

// ═══════════════════════════════════════════════
//  FONCTIONS PRINCIPALES
// ═══════════════════════════════════════════════

function configurerTout() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  creerFeuilleListes_(ss);
  const listeData = lireListesData_(ss);

  const chantierSheets = [];
  ss.getSheets().forEach(sh => {
    if (sh.getName() === FEUILLE_LISTES) return;
    configurerFeuilleChantier_(sh, ss, listeData);
    chantierSheets.push(sh);
  });

  // Masquer la feuille Listes (accessible via menu Chantiers)
  if (chantierSheets.length > 0) {
    ss.setActiveSheet(chantierSheets[0]);
    const listeSh = ss.getSheetByName(FEUILLE_LISTES);
    if (listeSh) listeSh.hideSheet();
  }

  SpreadsheetApp.getUi().alert(
    '✅ Configuration terminée.\n\n' +
    '• Thème sombre activé\n' +
    '• Menus déroulants filtrés par chantier\n' +
    '• Feuille Listes masquée\n' +
    '  → Menu Chantiers > Modifier les libellés\n\n' +
    'Pense à remplir la colonne chantier_id dans Listes.'
  );
}

// ═══════════════════════════════════════════════
//  LECTURE / FILTRAGE DES LIBELLÉS
// ═══════════════════════════════════════════════

/**
 * Lit les données de la feuille Listes (libellé + chantier_id + CP)
 */
function lireListesData_(ss) {
  const sh = ss.getSheetByName(FEUILLE_LISTES);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const numCols = Math.min(sh.getLastColumn(), 3);
  const data = sh.getRange(2, 1, lastRow - 1, numCols).getValues();
  return data
    .filter(row => row[0] !== '')
    .map(row => ({
      libelle: String(row[0]).trim(),
      chantierId: String(row[1] || '').trim(),
      cp: String(row[2] || '').trim().toUpperCase()
    }));
}

/**
 * Retourne les libellés pour un chantier, filtré par type CP optionnel.
 * cpType = 'C' (charges), 'P' (produits), ou vide (tous)
 */
function filtrerLibellesPourChantier_(listeData, nomChantier, cpType) {
  return listeData
    .filter(item => {
      const okChantier = item.chantierId === '' || item.chantierId === nomChantier;
      const okType = !cpType || item.cp === '' || item.cp === cpType;
      return okChantier && okType;
    })
    .map(item => item.libelle);
}

// ═══════════════════════════════════════════════
//  FEUILLE LISTES (cachée)
// ═══════════════════════════════════════════════

function creerFeuilleListes_(ss) {
  let sh = ss.getSheetByName(FEUILLE_LISTES);
  if (!sh) {
    // Création initiale
    sh = ss.insertSheet(FEUILLE_LISTES);
    sh.getRange('A1').setValue('Libellé');
    sh.getRange('B1').setValue('chantier_id');
    sh.getRange('C1').setValue('CP');
    sh.getRange(2, 1, LIBELLES_DEFAUT.length, 1)
      .setValues(LIBELLES_DEFAUT.map(l => [l]));
    sh.setColumnWidth(1, 260);
    sh.setColumnWidth(2, 160);
    sh.setColumnWidth(3, 60);
    sh.getRange('A1:C1').setFontWeight('bold');
  } else {
    // Migration : ajouter colonnes manquantes
    if (sh.getRange('B1').getValue() !== 'chantier_id') {
      sh.getRange('B1').setValue('chantier_id');
      sh.setColumnWidth(2, 160);
    }
    if (sh.getRange('C1').getValue() !== 'CP') {
      sh.getRange('C1').setValue('CP');
      sh.setColumnWidth(3, 60);
    }
    sh.getRange('A1:C1').setFontWeight('bold');
  }
  // Placer en dernier
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(ss.getNumSheets());
  return sh;
}

// ═══════════════════════════════════════════════
//  CONFIGURATION D'UNE FEUILLE CHANTIER
// ═══════════════════════════════════════════════

function configurerFeuilleChantier_(sh, ss, listeData) {
  const nomChantier = sh.getName();

  // En-têtes si feuille vide
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
  }

  // ── Limiter les lignes : dernière donnée + 10 ──
  const lastDataRow = Math.max(sh.getLastRow(), 1);
  const targetRows = lastDataRow + 10;
  const maxRows = sh.getMaxRows();
  if (maxRows > targetRows) {
    sh.deleteRows(targetRows + 1, maxRows - targetRows);
  } else if (maxRows < targetRows) {
    sh.insertRowsAfter(maxRows, targetRows - maxRows);
  }
  const nbLignes = targetRows - 1;

  sh.setFrozenRows(1);

  // ── Nettoyer l'ancienne colonne H helper si elle existe ──
  if (sh.getMaxColumns() >= 8) {
    try {
      sh.showColumns(8);
      if (sh.getRange('H1').getValue() === '_libelles') {
        sh.getRange('H1:H200').clear();
      }
    } catch (e) { /* ignore */ }
  }

  // ── Masquer colonnes au-delà de F ──
  const maxCols = sh.getMaxColumns();
  if (maxCols > NB_COLS) {
    sh.hideColumns(NB_COLS + 1, maxCols - NB_COLS);
  }

  // ── Largeurs adaptées mobile ──
  sh.setColumnWidth(1, 130);  // date (jjj dd/mm/aaaa)
  sh.setColumnWidth(2, 180);  // libelle
  sh.setColumnWidth(3, 90);   // charges
  sh.setColumnWidth(4, 90);   // produits
  sh.setColumnWidth(5, 90);   // reglement
  sh.setColumnWidth(6, 180);  // obs

  // ── THÈME SOMBRE ──
  appliquerThemeSombre_(sh, nbLignes);

  // Tab color
  sh.setTabColor(THEME.accent);

  // ── VALIDATIONS ──
  // Date (colonne A)
  const rangeDate = sh.getRange(2, 1, nbLignes, 1);
  rangeDate.setNumberFormat('ddd dd/MM/yyyy');
  rangeDate.setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireDate()
      .setAllowInvalid(true)
      .setHelpText('Date')
      .build()
  );

  // Libellés filtrés (colonne B)
  const libellesFiltres = filtrerLibellesPourChantier_(listeData, nomChantier);
  const rangeLibelle = sh.getRange(2, 2, nbLignes, 1);
  if (libellesFiltres.length > 0) {
    rangeLibelle.setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(libellesFiltres, true)
        .setAllowInvalid(true)
        .setHelpText('Choisis ou tape un libellé')
        .build()
    );
  } else {
    rangeLibelle.clearDataValidations();
  }

  // Nombres (colonnes C, D, E) — entiers, espace milliers
  [3, 4, 5].forEach(col => {
    const r = sh.getRange(2, col, nbLignes, 1);
    r.setNumberFormat('# ##0');
    r.setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireNumberGreaterThanOrEqualTo(0)
        .setAllowInvalid(true)
        .setHelpText('Montant (entier)')
        .build()
    );
  });
}

// ═══════════════════════════════════════════════
//  THÈME SOMBRE
// ═══════════════════════════════════════════════

function appliquerThemeSombre_(sh, nbLignes) {
  // ── Header ──
  const headerRange = sh.getRange(1, 1, 1, NB_COLS);
  headerRange
    .setBackground(THEME.headerBg)
    .setFontColor(THEME.headerText)
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, false, false, THEME.border, SpreadsheetApp.BorderStyle.SOLID);

  // ── Data rows : alternance de 2 teintes sombres ──
  if (nbLignes > 0) {
    const bgs = [];
    const fontColors = [];
    for (let i = 0; i < nbLignes; i++) {
      const bg = (i % 2 === 0) ? THEME.rowDark : THEME.rowDarkAlt;
      bgs.push(Array(NB_COLS).fill(bg));
      fontColors.push(Array(NB_COLS).fill(THEME.text));
    }
    const dataRange = sh.getRange(2, 1, nbLignes, NB_COLS);
    dataRange
      .setBackgrounds(bgs)
      .setFontColors(fontColors)
      .setFontSize(11)
      .setBorder(false, false, false, false, true, true, THEME.border, SpreadsheetApp.BorderStyle.SOLID);

    // Colonne "obs" en texte atténué
    sh.getRange(2, 6, nbLignes, 1).setFontColor(THEME.textMuted);
  }
}

// ═══════════════════════════════════════════════
//  MENU & UTILITAIRES
// ═══════════════════════════════════════════════

/** Affiche la feuille Listes (cachée) pour modification */
function afficherFeuilleListes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(FEUILLE_LISTES);
  if (sh) {
    sh.showSheet();
    ss.setActiveSheet(sh);
    SpreadsheetApp.getUi().alert(
      '📝 Feuille Listes visible.\n\n' +
      'Modifie les libellés et chantier_id,\n' +
      'puis relance : Chantiers > Configurer'
    );
  }
}

/** Menu PC (les menus ne fonctionnent pas sur mobile) */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏗 Chantiers')
    .addItem('⚙️ Configurer / mettre à jour', 'configurerTout')
    .addItem('📝 Modifier les libellés', 'afficherFeuilleListes')
    .addToUi();
}

// ═══════════════════════════════════════════════
//  AUTO-DATE : remplit la date du jour si vide
// ═══════════════════════════════════════════════

/**
 * Trigger onEdit : quand on saisit dans une ligne vide,
 * la colonne A est remplie automatiquement avec la date du jour.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();

  // Ignorer la feuille Listes et la ligne d'en-tête
  if (sh.getName() === FEUILLE_LISTES) return;
  const row = e.range.getRow();
  if (row < 2) return;

  // Si la cellule date (colonne A) de cette ligne est vide, y mettre aujourd'hui
  const cellDate = sh.getRange(row, 1);
  if (cellDate.getValue() === '') {
    cellDate.setValue(new Date());
  }
}

// ═══════════════════════════════════════════════
//  WEB APP PWA — API
// ═══════════════════════════════════════════════

/** Point d'entrée Web App — sert la page HTML ou répond aux appels API */
function doGet(e) {
  // Si un paramètre "action" est présent → c'est un appel API depuis Netlify
  if (e && e.parameter && e.parameter.action) {
    return handleAPI_(e.parameter.action, e.parameter.params);
  }
  // Sinon → servir la page HTML (comportement original)
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Saisie Chantier')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Retourne la liste des noms de chantiers (feuilles visibles) */
function getChantiers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets()
    .filter(sh => sh.getName() !== FEUILLE_LISTES)
    .map(sh => sh.getName());
}

/**
 * Retourne les libellés filtrés pour un chantier et un type CP.
 * cpType = 'C' (charges), 'P' (produits), ou '' (tous)
 */
function getLibellesForChantier(nomChantier, cpType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const listeData = lireListesData_(ss);
  return filtrerLibellesPourChantier_(listeData, nomChantier, cpType || '');
}

/** Retourne les saisies d'un chantier avec numéro de ligne */
function getSaisies(nomChantier) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(nomChantier);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const data = sh.getRange(2, 1, lastRow - 1, 6).getValues();
  const tz = Session.getScriptTimeZone();
  return data
    .map((row, idx) => ({
      rowNum: idx + 2,
      date: (row[0] instanceof Date)
        ? Utilities.formatDate(row[0], tz, 'yyyy-MM-dd')
        : String(row[0] || ''),
      libelle: String(row[1] || ''),
      charges: Number(row[2]) || 0,
      produits: Number(row[3]) || 0,
      reglement: Number(row[4]) || 0,
      obs: String(row[5] || '')
    }))
    .filter(r => r.date !== '' || r.libelle !== '')
    .reverse();
}

/** Ajoute une saisie dans la feuille du chantier */
function ajouterSaisie(nomChantier, saisie) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(nomChantier);
  if (!sh) throw new Error('Chantier "' + nomChantier + '" introuvable');

  const dateValue = saisie.date ? new Date(saisie.date + 'T12:00:00') : new Date();
  sh.appendRow([
    dateValue,
    saisie.libelle || '',
    saisie.charges ? Number(saisie.charges) : '',
    saisie.produits ? Number(saisie.produits) : '',
    '',
    saisie.obs || ''
  ]);
  return { success: true };
}

/** Modifie une saisie existante (par numéro de ligne) */
function modifierSaisie(nomChantier, rowNum, saisie) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(nomChantier);
  if (!sh) throw new Error('Chantier introuvable');

  const dateValue = saisie.date ? new Date(saisie.date + 'T12:00:00') : new Date();
  sh.getRange(rowNum, 1, 1, 6).setValues([[
    dateValue,
    saisie.libelle || '',
    saisie.charges ? Number(saisie.charges) : '',
    saisie.produits ? Number(saisie.produits) : '',
    '',
    saisie.obs || ''
  ]]);
  return { success: true };
}

/** Supprime une saisie (par numéro de ligne) */
function supprimerSaisie(nomChantier, rowNum) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(nomChantier);
  if (!sh) throw new Error('Chantier introuvable');
  sh.deleteRow(rowNum);
  return { success: true };
}

// ═══════════════════════════════════════════════
//  API POUR NETLIFY (handleAPI_)
// ═══════════════════════════════════════════════

function handleAPI_(action, paramsStr) {
  try {
    var params = {};
    if (paramsStr) {
      params = JSON.parse(paramsStr);
    }
    var result = null;

    if (action === 'getChantiers') {
      result = getChantiers();
    } else if (action === 'getSaisies') {
      result = getSaisies(params.chantier);
    } else if (action === 'getLibellesForChantier') {
      result = getLibellesForChantier(params.chantier, params.mode);
    } else if (action === 'ajouterSaisie') {
      result = ajouterSaisie(params.chantier, params.saisie);
    } else if (action === 'modifierSaisie') {
      result = modifierSaisie(params.chantier, params.rowNum, params.saisie);
    } else if (action === 'supprimerSaisie') {
      result = supprimerSaisie(params.chantier, params.rowNum);
    } else {
      throw new Error("Action inconnue: " + action);
    }

    return ContentService.createTextOutput(JSON.stringify({ result: result }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ error: error.message || String(error) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    return handleAPI_(body.action, JSON.stringify(body.params || {}));
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ error: error.message || String(error) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


