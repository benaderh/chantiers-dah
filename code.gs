/**
 * SAISIE CHANTIER MOBILE — Google Sheets
 * v3 — Module Versements chef de chantier
 *
 * FONCTIONNEMENT :
 * - Chaque feuille = un chantier (sauf "Listes" et "Versements")
 * - Feuille "Listes" : libellés + chantier_id + CP
 * - Feuille "Versements" : suivi des sommes versées au chef
 * - Col G (cachée) dans feuilles chantier = Réf versement (V-001...)
 *
 * INSTALLATION :
 * 1. Extensions > Apps Script > coller ce code > Exécuter "configurerTout"
 * 2. Menu Chantiers > "Modifier les libellés"
 */

const HEADERS           = ['date', 'libelle', 'charges', 'produits', 'reglement', 'obs'];
const FEUILLE_LISTES    = 'Listes';
const FEUILLE_VERSEMENTS = 'Versements';
const NB_COLS           = 6;
const COL_REF_VERS      = 7; // Col G dans les feuilles chantier = référence versement

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
  creerFeuilleVersements_(ss);
  const listeData = lireListesData_(ss);

  const chantierSheets = [];
  ss.getSheets().forEach(sh => {
    if (sh.getName() === FEUILLE_LISTES)     return;
    if (sh.getName() === FEUILLE_VERSEMENTS) return;
    configurerFeuilleChantier_(sh, ss, listeData);
    chantierSheets.push(sh);
  });

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
    '• Feuille Versements créée/mise à jour\n\n' +
    'Pense à remplir la colonne chantier_id dans Listes.'
  );
}

// ═══════════════════════════════════════════════
//  LECTURE / FILTRAGE DES LIBELLÉS
// ═══════════════════════════════════════════════

function lireListesData_(ss) {
  const sh = ss.getSheetByName(FEUILLE_LISTES);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const numCols = Math.min(sh.getLastColumn(), 3);
  const data = sh.getRange(2, 1, lastRow - 1, numCols).getValues();
  return data
    .filter(row => row[0] !== '')
    .map(row => ({
      libelle:    String(row[0]).trim(),
      chantierId: String(row[1] || '').trim(),
      cp:         String(row[2] || '').trim().toUpperCase()
    }));
}

function filtrerLibellesPourChantier_(listeData, nomChantier, cpType) {
  return listeData
    .filter(item => {
      const okChantier = item.chantierId === '' || item.chantierId === nomChantier;
      const okType     = !cpType || item.cp === '' || item.cp === cpType;
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
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(ss.getNumSheets());
  return sh;
}

// ═══════════════════════════════════════════════
//  FEUILLE VERSEMENTS
// ═══════════════════════════════════════════════

function creerFeuilleVersements_(ss) {
  let sh = ss.getSheetByName(FEUILLE_VERSEMENTS);
  if (!sh) {
    sh = ss.insertSheet(FEUILLE_VERSEMENTS);
    const hdrs = [['ID', 'Date', 'Chantier', 'Montant', 'Obs']];
    sh.getRange(1, 1, 1, 5)
      .setValues(hdrs)
      .setBackground(THEME.headerBg)
      .setFontColor(THEME.headerText)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    sh.setColumnWidth(1, 80);
    sh.setColumnWidth(2, 120);
    sh.setColumnWidth(3, 180);
    sh.setColumnWidth(4, 110);
    sh.setColumnWidth(5, 240);
    sh.getRange(2, 2, 200, 1).setNumberFormat('dd/MM/yyyy');
    sh.getRange(2, 4, 200, 1).setNumberFormat('# ##0');
    sh.setTabColor('#f39c12');
    sh.setFrozenRows(1);
  }
  return sh;
}

// ═══════════════════════════════════════════════
//  CONFIGURATION D'UNE FEUILLE CHANTIER
// ═══════════════════════════════════════════════

function configurerFeuilleChantier_(sh, ss, listeData) {
  const nomChantier = sh.getName();

  if (sh.getLastRow() === 0) sh.appendRow(HEADERS);

  const lastDataRow = Math.max(sh.getLastRow(), 1);
  const targetRows  = lastDataRow + 10;
  const maxRows     = sh.getMaxRows();
  if (maxRows > targetRows)      sh.deleteRows(targetRows + 1, maxRows - targetRows);
  else if (maxRows < targetRows) sh.insertRowsAfter(maxRows, targetRows - maxRows);
  const nbLignes = targetRows - 1;

  sh.setFrozenRows(1);

  // Nettoyer ancienne col H helper
  if (sh.getMaxColumns() >= 8) {
    try {
      sh.showColumns(8);
      if (sh.getRange('H1').getValue() === '_libelles') sh.getRange('H1:H200').clear();
    } catch (e) {}
  }

  // S'assurer que la col G existe (Réf_Vers), puis masquer tout au-delà
  const maxCols = sh.getMaxColumns();
  if (maxCols < COL_REF_VERS) {
    sh.insertColumnsAfter(maxCols, COL_REF_VERS - maxCols);
  }
  // Masquer col G et au-delà de G
  const totalCols = sh.getMaxColumns();
  if (totalCols > COL_REF_VERS) sh.hideColumns(COL_REF_VERS + 1, totalCols - COL_REF_VERS);
  // Masquer col G (Réf_Vers — usage interne)
  try { sh.showColumns(COL_REF_VERS); sh.hideColumns(COL_REF_VERS); } catch(e) {}

  // Largeurs
  sh.setColumnWidth(1, 130);
  sh.setColumnWidth(2, 180);
  sh.setColumnWidth(3, 90);
  sh.setColumnWidth(4, 90);
  sh.setColumnWidth(5, 90);
  sh.setColumnWidth(6, 180);

  appliquerThemeSombre_(sh, nbLignes);
  sh.setTabColor(THEME.accent);

  // ── Validations ──
  const rangeDate = sh.getRange(2, 1, nbLignes, 1);
  rangeDate.setNumberFormat('ddd dd/MM/yyyy');
  rangeDate.setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireDate().setAllowInvalid(true).setHelpText('Date').build()
  );

  const libellesFiltres = filtrerLibellesPourChantier_(listeData, nomChantier);
  const rangeLibelle    = sh.getRange(2, 2, nbLignes, 1);
  if (libellesFiltres.length > 0) {
    rangeLibelle.setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(libellesFiltres, true)
        .setAllowInvalid(true).setHelpText('Choisis ou tape un libellé').build()
    );
  } else {
    rangeLibelle.clearDataValidations();
  }

  [3, 4, 5].forEach(col => {
    const r = sh.getRange(2, col, nbLignes, 1);
    r.setNumberFormat('# ##0');
    r.setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireNumberGreaterThanOrEqualTo(0)
        .setAllowInvalid(true).setHelpText('Montant (entier)').build()
    );
  });
}

// ═══════════════════════════════════════════════
//  THÈME SOMBRE
// ═══════════════════════════════════════════════

function appliquerThemeSombre_(sh, nbLignes) {
  const headerRange = sh.getRange(1, 1, 1, NB_COLS);
  headerRange
    .setBackground(THEME.headerBg)
    .setFontColor(THEME.headerText)
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, false, false, THEME.border, SpreadsheetApp.BorderStyle.SOLID);

  if (nbLignes > 0) {
    const bgs = [], fontColors = [];
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
    sh.getRange(2, 6, nbLignes, 1).setFontColor(THEME.textMuted);
  }
}

// ═══════════════════════════════════════════════
//  MENU & UTILITAIRES
// ═══════════════════════════════════════════════

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

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏗 Chantiers')
    .addItem('⚙️ Configurer / mettre à jour', 'configurerTout')
    .addItem('📝 Modifier les libellés',       'afficherFeuilleListes')
    .addToUi();
}

// ═══════════════════════════════════════════════
//  AUTO-DATE
// ═══════════════════════════════════════════════

function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() === FEUILLE_LISTES)     return;
  if (sh.getName() === FEUILLE_VERSEMENTS) return;
  const row = e.range.getRow();
  if (row < 2) return;
  const cellDate = sh.getRange(row, 1);
  if (cellDate.getValue() === '') cellDate.setValue(new Date());
}

// ═══════════════════════════════════════════════
//  WEB APP PWA — API
// ═══════════════════════════════════════════════

function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    return handleAPI_(e.parameter.action, e.parameter.params);
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Saisie Chantier')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Liste des chantiers (feuilles visibles, hors Listes et Versements) */
function getChantiers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets()
    .filter(sh => sh.getName() !== FEUILLE_LISTES && sh.getName() !== FEUILLE_VERSEMENTS)
    .map(sh => sh.getName());
}

function getLibellesForChantier(nomChantier, cpType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const listeData = lireListesData_(ss);
  return filtrerLibellesPourChantier_(listeData, nomChantier, cpType || '');
}

/** Retourne les saisies d'un chantier avec numéro de ligne et refVers */
function getSaisies(nomChantier) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(nomChantier);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  // Lire jusqu'à la col G (Réf_Vers) si elle existe
  const nCols = Math.min(Math.max(sh.getLastColumn(), 6), COL_REF_VERS);
  const data = sh.getRange(2, 1, lastRow - 1, nCols).getValues();
  const tz   = Session.getScriptTimeZone();
  return data
    .map((row, idx) => ({
      rowNum:   idx + 2,
      date:     (row[0] instanceof Date)
                  ? Utilities.formatDate(row[0], tz, 'yyyy-MM-dd')
                  : String(row[0] || ''),
      libelle:  String(row[1] || ''),
      charges:  Number(row[2]) || 0,
      produits: Number(row[3]) || 0,
      reglement:Number(row[4]) || 0,
      obs:      String(row[5] || ''),
      refVers:  String(row[6] || '')  // V-001 si versé, '' sinon
    }))
    .filter(r => r.date !== '' || r.libelle !== '')
    .reverse();
}

/** Ajoute une saisie (ne touche pas la col G) */
function ajouterSaisie(nomChantier, saisie) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(nomChantier);
  if (!sh) throw new Error('Chantier "' + nomChantier + '" introuvable');
  const dateValue = saisie.date ? new Date(saisie.date + 'T12:00:00') : new Date();
  sh.appendRow([
    dateValue,
    saisie.libelle  || '',
    saisie.charges  ? Number(saisie.charges)  : '',
    saisie.produits ? Number(saisie.produits) : '',
    '',
    saisie.obs || ''
  ]);
  return { success: true };
}

/** Modifie une saisie existante (ne touche pas la col G) */
function modifierSaisie(nomChantier, rowNum, saisie) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(nomChantier);
  if (!sh) throw new Error('Chantier introuvable');
  const dateValue = saisie.date ? new Date(saisie.date + 'T12:00:00') : new Date();
  sh.getRange(rowNum, 1, 1, 6).setValues([[
    dateValue,
    saisie.libelle  || '',
    saisie.charges  ? Number(saisie.charges)  : '',
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
//  VERSEMENTS — FONCTIONS API
// ═══════════════════════════════════════════════

/**
 * Retourne tous les versements d'un chantier avec calcul des charges couvertes.
 */
function getVersements(nomChantier) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const shV = ss.getSheetByName(FEUILLE_VERSEMENTS);
  if (!shV || shV.getLastRow() < 2) return [];

  const rows = shV.getRange(2, 1, shV.getLastRow() - 1, 5).getValues();
  const tz   = Session.getScriptTimeZone();

  // Lire les données du chantier (col G incluse)
  const chSh = ss.getSheetByName(nomChantier);
  let chData = [];
  if (chSh && chSh.getLastRow() >= 2) {
    const nCols = Math.max(chSh.getLastColumn(), COL_REF_VERS);
    chData = chSh.getRange(2, 1, chSh.getLastRow() - 1, nCols).getValues();
  }

  return rows
    .filter(r => String(r[2]) === nomChantier && r[0] !== '')
    .map(r => {
      const id = String(r[0]);

      // Charges couvertes par ce versement
      const chargesCouvertes = chData
        .filter(cr => String(cr[COL_REF_VERS - 1] || '') === id)
        .reduce((s, cr) => s + (Number(cr[2]) || 0), 0);

      // Détail des charges (libellé + montant)
      const details = chData
        .filter(cr => String(cr[COL_REF_VERS - 1] || '') === id && Number(cr[2]) > 0)
        .map(cr => ({ libelle: String(cr[1] || ''), montant: Number(cr[2]) || 0 }));

      return {
        id,
        date:             (r[1] instanceof Date) ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : String(r[1] || ''),
        chantier:         String(r[2] || ''),
        montant:          Number(r[3]) || 0,
        obs:              String(r[4] || ''),
        chargesCouvertes,
        solde:            (Number(r[3]) || 0) - chargesCouvertes,
        details
      };
    });
}

/**
 * Retourne les charges (col C > 0) sans versement associé.
 * Utilisé pour remplir la checklist du formulaire Nouveau Versement.
 */
function getChargesSansVersement(nomChantier) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(nomChantier);
  if (!sh || sh.getLastRow() < 2) return [];

  const nCols = Math.max(sh.getLastColumn(), COL_REF_VERS);
  const data  = sh.getRange(2, 1, sh.getLastRow() - 1, nCols).getValues();
  const tz    = Session.getScriptTimeZone();

  return data
    .map((row, idx) => ({
      rowNum:  idx + 2,
      date:    (row[0] instanceof Date) ? Utilities.formatDate(row[0], tz, 'yyyy-MM-dd') : String(row[0] || ''),
      libelle: String(row[1] || ''),
      charges: Number(row[2]) || 0,
      refVers: String(row[COL_REF_VERS - 1] || '')
    }))
    .filter(r => r.charges > 0 && r.libelle !== '' && r.refVers === '');
}

/**
 * Crée un versement et marque les lignes sélectionnées dans la feuille chantier.
 * versement = { date, montant, obs, rowNums: [2,5,7] }
 */
function ajouterVersement(nomChantier, versement) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let   shV = ss.getSheetByName(FEUILLE_VERSEMENTS);
  if (!shV) shV = creerFeuilleVersements_(ss);

  // Générer l'ID : V-XXX basé sur nombre de lignes actuelles
  const lastRow = shV.getLastRow(); // 1 = seulement header → prochain = V-001
  const id = 'V-' + String(lastRow).padStart(3, '0');

  const dateVal = versement.date ? new Date(versement.date + 'T12:00:00') : new Date();
  shV.appendRow([id, dateVal, nomChantier, Number(versement.montant) || 0, versement.obs || '']);

  // Écrire la Réf_Vers dans col G des lignes concernées
  if (versement.rowNums && versement.rowNums.length > 0) {
    const chSh = ss.getSheetByName(nomChantier);
    if (!chSh) throw new Error('Chantier introuvable');

    // S'assurer que la col G existe
    while (chSh.getMaxColumns() < COL_REF_VERS) {
      chSh.insertColumnAfter(chSh.getMaxColumns());
    }
    // Maintenir la col G cachée
    try { chSh.showColumns(COL_REF_VERS); chSh.hideColumns(COL_REF_VERS); } catch(e) {}

    versement.rowNums.forEach(rn => {
      if (rn >= 2) chSh.getRange(rn, COL_REF_VERS).setValue(id);
    });
  }

  return { success: true, id };
}

/**
 * Supprime un versement et efface les références dans la feuille chantier.
 */
function supprimerVersement(nomChantier, versementId) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const shV = ss.getSheetByName(FEUILLE_VERSEMENTS);
  if (!shV) return { success: false };

  // Supprimer la ligne dans Versements
  const data = shV.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === versementId) {
      shV.deleteRow(i + 1);
      break;
    }
  }

  // Effacer les Réf_Vers dans la feuille chantier
  const chSh = ss.getSheetByName(nomChantier);
  if (chSh && chSh.getLastRow() >= 2 && chSh.getMaxColumns() >= COL_REF_VERS) {
    const refs = chSh.getRange(2, COL_REF_VERS, chSh.getLastRow() - 1, 1).getValues();
    refs.forEach((row, idx) => {
      if (String(row[0]) === versementId) {
        chSh.getRange(idx + 2, COL_REF_VERS).setValue('');
      }
    });
  }

  return { success: true };
}

// ═══════════════════════════════════════════════
//  API DISPATCHER
// ═══════════════════════════════════════════════

function handleAPI_(action, paramsStr) {
  try {
    var params = {};
    if (paramsStr) params = JSON.parse(paramsStr);
    var result = null;

    if      (action === 'getChantiers')            result = getChantiers();
    else if (action === 'getSaisies')              result = getSaisies(params.chantier);
    else if (action === 'getLibellesForChantier')  result = getLibellesForChantier(params.chantier, params.mode);
    else if (action === 'ajouterSaisie')           result = ajouterSaisie(params.chantier, params.saisie);
    else if (action === 'modifierSaisie')          result = modifierSaisie(params.chantier, params.rowNum, params.saisie);
    else if (action === 'supprimerSaisie')         result = supprimerSaisie(params.chantier, params.rowNum);
    else if (action === 'getVersements')           result = getVersements(params.chantier);
    else if (action === 'getChargesSansVersement') result = getChargesSansVersement(params.chantier);
    else if (action === 'ajouterVersement')        result = ajouterVersement(params.chantier, params.versement);
    else if (action === 'supprimerVersement')      result = supprimerVersement(params.chantier, params.versementId);
    else throw new Error('Action inconnue: ' + action);

    return ContentService.createTextOutput(JSON.stringify({ result }))
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
