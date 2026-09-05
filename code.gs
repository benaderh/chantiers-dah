/**
 * SAISIE CHANTIER MOBILE — Google Sheets
 * v4 — Versements + Règlements + Tranches
 *
 * Structure feuilles chantier (colonnes A-G) :
 *   A: Date  B: Libellé  C: Charges  D: Produits  E: Reglement  F: Obs  G: Réf_Vers (cachée)
 *
 * Feuille "Versements" (colonnes A-E) :
 *   A: ID  B: Date  C: Chantier  D: Montant  E: Obs
 *
 * INSTALLATION :
 *   1. Extensions > Apps Script > coller ce code > Exécuter "configurerTout"
 *   2. Déployer > Nouveau déploiement > Application Web > Tout le monde
 *   3. Copier l'URL dans index.html (SCRIPT_URL)
 */

const HEADERS            = ['date', 'libelle', 'charges', 'produits', 'reglement', 'obs'];
const FEUILLE_LISTES     = 'Listes';
const FEUILLE_VERSEMENTS = 'Versements';
const NB_COLS            = 6;
const COL_REF_VERS       = 7; // Col G : références versements (ex: "V-001" ou "V-001:5000")

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
//  CONFIGURATION PRINCIPALE
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
    '✅ Configuration terminée (v4).\n\n' +
    '• Col G (Réf_Vers) ajoutée et masquée dans chaque feuille\n' +
    '• Feuille "Versements" créée/mise à jour\n' +
    '• Thème sombre + menus filtrés\n\n' +
    'N\'oublie pas de Déployer > Nouvelle version après toute modification.'
  );
}

/**
 * VERSION RAPIDE — uniquement pour mettre à jour le module Versements.
 * ✅ Exécuter CETTE fonction si "configurerTout" dépasse le délai.
 * Elle ajoute/crée uniquement :
 *   • La feuille "Versements"
 *   • La colonne G (Réf_Vers, cachée) dans chaque feuille chantier
 * Elle NE refait PAS le thème ni les validations (plus rapide).
 */
function ajouterModuleVersements() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Créer la feuille Versements
  creerFeuilleVersements_(ss);

  // 2. Pour chaque feuille chantier, ajouter col G si absente
  const modifiees = [];
  ss.getSheets().forEach(sh => {
    const nom = sh.getName();
    if (nom === FEUILLE_LISTES || nom === FEUILLE_VERSEMENTS) return;

    // Ajouter col G si besoin
    while (sh.getMaxColumns() < COL_REF_VERS) {
      sh.insertColumnAfter(sh.getMaxColumns());
    }
    // Masquer col G
    try { sh.showColumns(COL_REF_VERS); sh.hideColumns(COL_REF_VERS); } catch(e) {}
    // Écrire l'en-tête en G1 si vide
    const hdrCell = sh.getRange(1, COL_REF_VERS);
    if (!hdrCell.getValue()) hdrCell.setValue('Réf_Vers');

    modifiees.push(nom);
  });

  // 3. Résultat
  SpreadsheetApp.getUi().alert(
    '✅ Module Versements ajouté !\n\n' +
    '• Feuille "Versements" : prête\n' +
    '• Col G (Réf_Vers) ajoutée dans :\n  - ' + modifiees.join('\n  - ') + '\n\n' +
    'Maintenant : Déployer > Gérer les déploiements > Nouvelle version.'
  );
}

/**
 * TEST RAPIDE — exécuter depuis Apps Script pour vérifier que tout fonctionne.
 * Affiche un résumé de l'état de la configuration.
 */
function testerAPI() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Vérifier feuille Versements
  const shV = ss.getSheetByName(FEUILLE_VERSEMENTS);
  const versementsOk = !!shV;

  // Vérifier col G dans chaque feuille chantier
  const colGStatus = [];
  ss.getSheets().forEach(sh => {
    const nom = sh.getName();
    if (nom === FEUILLE_LISTES || nom === FEUILLE_VERSEMENTS) return;
    const hasColG = sh.getMaxColumns() >= COL_REF_VERS;
    const isHidden = hasColG ? sh.isColumnHiddenByUser(COL_REF_VERS) : false;
    colGStatus.push(nom + ': Col G ' + (hasColG ? (isHidden ? '✅ cachée' : '⚠️ visible') : '❌ absente'));
  });

  // Test écriture versement
  let writeTest = '❌';
  try {
    const testResult = getChantiers();
    writeTest = '✅ ' + testResult.length + ' chantier(s) trouvé(s): ' + testResult.join(', ');
  } catch(e) { writeTest = '❌ ' + e.message; }

  SpreadsheetApp.getUi().alert(
    '🔍 DIAGNOSTIC API\n\n' +
    '1. Feuille Versements : ' + (versementsOk ? '✅ OK' : '❌ ABSENTE → exécuter ajouterModuleVersements') + '\n\n' +
    '2. Chantiers :\n   ' + (colGStatus.join('\n   ') || 'Aucun chantier') + '\n\n' +
    '3. API getChantiers : ' + writeTest + '\n\n' +
    '─────────────────────\n' +
    'Si tout est ✅ ici mais le bouton ne fonctionne pas dans l\'app :\n' +
    '→ Déployer > Gérer les déploiements > ✏️ Modifier > Nouvelle version > Déployer'
  );
}

// ═══════════════════════════════════════════════
//  FEUILLE LISTES
// ═══════════════════════════════════════════════

function lireListesData_(ss) {
  const sh = ss.getSheetByName(FEUILLE_LISTES);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const numCols = Math.min(sh.getLastColumn(), 3);
  return sh.getRange(2, 1, lastRow - 1, numCols).getValues()
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
      const okC = item.chantierId === '' || item.chantierId === nomChantier;
      const okT = !cpType || item.cp === '' || item.cp === cpType;
      return okC && okT;
    })
    .map(item => item.libelle);
}

function creerFeuilleListes_(ss) {
  let sh = ss.getSheetByName(FEUILLE_LISTES);
  if (!sh) {
    sh = ss.insertSheet(FEUILLE_LISTES);
    sh.getRange('A1').setValue('Libellé');
    sh.getRange('B1').setValue('chantier_id');
    sh.getRange('C1').setValue('CP');
    sh.getRange(2, 1, LIBELLES_DEFAUT.length, 1).setValues(LIBELLES_DEFAUT.map(l => [l]));
    sh.setColumnWidth(1, 260); sh.setColumnWidth(2, 160); sh.setColumnWidth(3, 60);
    sh.getRange('A1:C1').setFontWeight('bold');
  } else {
    if (sh.getRange('B1').getValue() !== 'chantier_id') { sh.getRange('B1').setValue('chantier_id'); sh.setColumnWidth(2,160); }
    if (sh.getRange('C1').getValue() !== 'CP')          { sh.getRange('C1').setValue('CP');          sh.setColumnWidth(3,60);  }
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
    sh.getRange(1, 1, 1, 5).setValues([['ID', 'Date', 'Chantier', 'Montant', 'Obs']])
      .setBackground(THEME.headerBg).setFontColor(THEME.headerText)
      .setFontWeight('bold').setHorizontalAlignment('center');
    sh.setColumnWidth(1,80); sh.setColumnWidth(2,120);
    sh.setColumnWidth(3,180); sh.setColumnWidth(4,110); sh.setColumnWidth(5,240);
    sh.getRange(2,2,500,1).setNumberFormat('dd/MM/yyyy');
    sh.getRange(2,4,500,1).setNumberFormat('# ##0');
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

  // S'assurer que col G existe (Réf_Vers)
  while (sh.getMaxColumns() < COL_REF_VERS) sh.insertColumnAfter(sh.getMaxColumns());
  // Masquer col G et tout ce qui va au-delà
  const totalCols = sh.getMaxColumns();
  try { sh.showColumns(COL_REF_VERS); sh.hideColumns(COL_REF_VERS); } catch(e) {}
  if (totalCols > COL_REF_VERS) try { sh.hideColumns(COL_REF_VERS + 1, totalCols - COL_REF_VERS); } catch(e) {}

  sh.setColumnWidth(1, 130); sh.setColumnWidth(2, 180); sh.setColumnWidth(3, 90);
  sh.setColumnWidth(4, 90);  sh.setColumnWidth(5, 90);  sh.setColumnWidth(6, 180);

  appliquerThemeSombre_(sh, nbLignes);
  sh.setTabColor(THEME.accent);

  // Validations
  const rangeDate = sh.getRange(2, 1, nbLignes, 1);
  rangeDate.setNumberFormat('ddd dd/MM/yyyy');
  rangeDate.setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build());

  const libellesFiltres = filtrerLibellesPourChantier_(listeData, nomChantier);
  const rangeLibelle    = sh.getRange(2, 2, nbLignes, 1);
  if (libellesFiltres.length > 0) {
    rangeLibelle.setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(libellesFiltres, true).setAllowInvalid(true).build()
    );
  } else { rangeLibelle.clearDataValidations(); }

  [3, 4, 5].forEach(col => {
    const r = sh.getRange(2, col, nbLignes, 1);
    r.setNumberFormat('# ##0');
    r.setDataValidation(SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(true).build());
  });
}

// ═══════════════════════════════════════════════
//  THÈME SOMBRE
// ═══════════════════════════════════════════════

function appliquerThemeSombre_(sh, nbLignes) {
  sh.getRange(1, 1, 1, NB_COLS)
    .setBackground(THEME.headerBg).setFontColor(THEME.headerText)
    .setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center')
    .setBorder(true,true,true,true,false,false, THEME.border, SpreadsheetApp.BorderStyle.SOLID);

  if (nbLignes > 0) {
    const bgs = [], fcs = [];
    for (let i = 0; i < nbLignes; i++) {
      const bg = (i % 2 === 0) ? THEME.rowDark : THEME.rowDarkAlt;
      bgs.push(Array(NB_COLS).fill(bg));
      fcs.push(Array(NB_COLS).fill(THEME.text));
    }
    sh.getRange(2, 1, nbLignes, NB_COLS)
      .setBackgrounds(bgs).setFontColors(fcs).setFontSize(11)
      .setBorder(false,false,false,false,true,true, THEME.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(2, 6, nbLignes, 1).setFontColor(THEME.textMuted);
  }
}

// ═══════════════════════════════════════════════
//  MENU
// ═══════════════════════════════════════════════

function afficherFeuilleListes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(FEUILLE_LISTES);
  if (sh) { sh.showSheet(); ss.setActiveSheet(sh); }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏗 Chantiers')
    .addItem('⚡ Ajouter module Versements (RAPIDE)', 'ajouterModuleVersements')
    .addSeparator()
    .addItem('⚙️ Configurer tout (lent, 1ère fois)', 'configurerTout')
    .addItem('📝 Modifier libellés', 'afficherFeuilleListes')
    .addToUi();
}


function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() === FEUILLE_LISTES || sh.getName() === FEUILLE_VERSEMENTS) return;
  const row = e.range.getRow();
  if (row < 2) return;
  const cellDate = sh.getRange(row, 1);
  if (cellDate.getValue() === '') cellDate.setValue(new Date());
}

// ═══════════════════════════════════════════════
//  WEB APP
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

// ═══════════════════════════════════════════════
//  SAISIES API
// ═══════════════════════════════════════════════

function getChantiers() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .filter(sh => sh.getName() !== FEUILLE_LISTES && sh.getName() !== FEUILLE_VERSEMENTS)
    .map(sh => sh.getName());
}

function getLibellesForChantier(nomChantier, cpType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return filtrerLibellesPourChantier_(lireListesData_(ss), nomChantier, cpType || '');
}

/** Retourne les saisies avec refVers (col G) */
function getSaisies(nomChantier) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(nomChantier);
  if (!sh || sh.getLastRow() < 2) return [];
  const nCols = Math.min(Math.max(sh.getLastColumn(), 6), COL_REF_VERS);
  const data  = sh.getRange(2, 1, sh.getLastRow() - 1, nCols).getValues();
  const tz    = Session.getScriptTimeZone();
  return data
    .map((row, idx) => ({
      rowNum:    idx + 2,
      date:      (row[0] instanceof Date) ? Utilities.formatDate(row[0], tz, 'yyyy-MM-dd') : String(row[0]||''),
      libelle:   String(row[1]||''),
      charges:   Number(row[2])||0,
      produits:  Number(row[3])||0,
      reglement: Number(row[4])||0,
      obs:       String(row[5]||''),
      refVers:   String(row[6]||'')
    }))
    .filter(r => r.date !== '' || r.libelle !== '')
    .reverse();
}

/**
 * Ajoute une saisie (charge, produit, ou règlement direct)
 * saisie = { date, libelle, charges, produits, reglement, obs }
 */
function ajouterSaisie(nomChantier, saisie) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(nomChantier);
  if (!sh) throw new Error('Chantier "' + nomChantier + '" introuvable');
  const dateVal = saisie.date ? new Date(saisie.date + 'T12:00:00') : new Date();
  sh.appendRow([
    dateVal,
    saisie.libelle   || '',
    saisie.charges   ? Number(saisie.charges)   : '',
    saisie.produits  ? Number(saisie.produits)  : '',
    saisie.reglement ? Number(saisie.reglement) : '',
    saisie.obs       || ''
  ]);
  return { success: true };
}

/**
 * Modifie une saisie (ne touche pas la col G = Réf_Vers)
 */
function modifierSaisie(nomChantier, rowNum, saisie) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(nomChantier);
  if (!sh) throw new Error('Chantier introuvable');
  const dateVal = saisie.date ? new Date(saisie.date + 'T12:00:00') : new Date();
  sh.getRange(rowNum, 1, 1, 6).setValues([[
    dateVal,
    saisie.libelle   || '',
    saisie.charges   ? Number(saisie.charges)   : '',
    saisie.produits  ? Number(saisie.produits)  : '',
    saisie.reglement ? Number(saisie.reglement) : '',
    saisie.obs       || ''
  ]]);
  return { success: true };
}

function supprimerSaisie(nomChantier, rowNum) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(nomChantier);
  if (!sh) throw new Error('Chantier introuvable');
  sh.deleteRow(rowNum);
  return { success: true };
}

// ═══════════════════════════════════════════════
//  VERSEMENTS API
// ═══════════════════════════════════════════════

/**
 * Retourne les versements du chantier avec détail des charges couvertes.
 * La col G peut contenir "V-001" ou "V-001:5000" (versement:montant_tranche)
 */
function getVersements(nomChantier) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const shV = ss.getSheetByName(FEUILLE_VERSEMENTS);
  if (!shV || shV.getLastRow() < 2) return [];

  const rows = shV.getRange(2, 1, shV.getLastRow() - 1, 5).getValues();
  const tz   = Session.getScriptTimeZone();

  // Lire données chantier (col G incluse)
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

      // Chercher les charges avec cette référence dans col G
      // Format col G: "V-001" ou "V-001:5000" (montant tranche)
      let chargesCouvertes = 0;
      const details = [];
      chData.forEach(cr => {
        const refRaw = String(cr[COL_REF_VERS - 1] || '');
        // Une charge peut avoir plusieurs refs: "V-001,V-002"
        const refs = refRaw.split(',').map(s => s.trim());
        refs.forEach(ref => {
          const parts   = ref.split(':');
          const refId   = parts[0].trim();
          const tranche = parts[1] ? Number(parts[1]) : 0;
          if (refId === id && Number(cr[2]) > 0) {
            const montantDetail = tranche || Number(cr[2]) || 0;
            chargesCouvertes += montantDetail;
            details.push({
              libelle: String(cr[1] || ''),
              charge:  Number(cr[2]) || 0,
              tranche: tranche || null
            });
          }
        });
      });

      return {
        id,
        date:    (r[1] instanceof Date) ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : String(r[1]||''),
        chantier: String(r[2]||''),
        montant:  Number(r[3])||0,
        obs:      String(r[4]||''),
        chargesCouvertes,
        solde:    (Number(r[3])||0) - chargesCouvertes,
        details
      };
    });
}

/**
 * Retourne les charges (col C > 0) du chantier.
 * Inclut le règlement déjà effectué et les refs versements.
 * Utilisé pour la checklist du formulaire Versement.
 */
function getChargesPourVersement(nomChantier) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(nomChantier);
  if (!sh || sh.getLastRow() < 2) return [];
  const nCols = Math.max(sh.getLastColumn(), COL_REF_VERS);
  const data  = sh.getRange(2, 1, sh.getLastRow() - 1, nCols).getValues();
  const tz    = Session.getScriptTimeZone();
  return data
    .map((row, idx) => ({
      rowNum:    idx + 2,
      date:      (row[0] instanceof Date) ? Utilities.formatDate(row[0], tz, 'yyyy-MM-dd') : String(row[0]||''),
      libelle:   String(row[1]||''),
      charges:   Number(row[2])||0,
      reglement: Number(row[4])||0,  // déjà payé directement
      refVers:   String(row[COL_REF_VERS-1]||'')
    }))
    .filter(r => r.charges > 0 && r.libelle !== '');
}

/**
 * Crée un versement et marque les lignes concernées en col G.
 * versement.lignes = [{ rowNum, montant }, ...] (montant = tranche ou charge entière)
 */
function ajouterVersement(nomChantier, versement) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let   shV = ss.getSheetByName(FEUILLE_VERSEMENTS);
  if (!shV) shV = creerFeuilleVersements_(ss);

  // ID auto: V-NNN
  const lastRow = shV.getLastRow();
  const id      = 'V-' + String(lastRow).padStart(3, '0');

  const dateVal = versement.date ? new Date(versement.date + 'T12:00:00') : new Date();
  shV.appendRow([id, dateVal, nomChantier, Number(versement.montant)||0, versement.obs||'']);

  // Marquer les charges en col G
  if (versement.lignes && versement.lignes.length > 0) {
    const chSh = ss.getSheetByName(nomChantier);
    if (!chSh) throw new Error('Chantier introuvable');
    while (chSh.getMaxColumns() < COL_REF_VERS) chSh.insertColumnAfter(chSh.getMaxColumns());
    try { chSh.showColumns(COL_REF_VERS); chSh.hideColumns(COL_REF_VERS); } catch(e) {}

    versement.lignes.forEach(ligne => {
      if (ligne.rowNum < 2) return;
      const cell    = chSh.getRange(ligne.rowNum, COL_REF_VERS);
      const existing = String(cell.getValue() || '').trim();
      // Format: "V-001:5000" si tranche, sinon "V-001"
      const newRef  = ligne.montant ? id + ':' + ligne.montant : id;
      // Ajouter à la liste existante (virgule-séparée)
      const updated = existing ? existing + ',' + newRef : newRef;
      cell.setValue(updated);
    });
  }

  return { success: true, id };
}

/**
 * Supprime un versement et efface ses références dans col G.
 */
function supprimerVersement(nomChantier, versementId) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const shV = ss.getSheetByName(FEUILLE_VERSEMENTS);
  if (!shV) return { success: false };

  // Supprimer la ligne
  const data = shV.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === versementId) { shV.deleteRow(i + 1); break; }
  }

  // Effacer la référence dans col G du chantier
  const chSh = ss.getSheetByName(nomChantier);
  if (chSh && chSh.getLastRow() >= 2 && chSh.getMaxColumns() >= COL_REF_VERS) {
    const refs = chSh.getRange(2, COL_REF_VERS, chSh.getLastRow() - 1, 1).getValues();
    refs.forEach((row, idx) => {
      const refRaw = String(row[0] || '');
      if (!refRaw) return;
      // Supprimer les segments contenant cet ID
      const newRef = refRaw.split(',')
        .filter(s => !s.trim().startsWith(versementId))
        .join(',');
      chSh.getRange(idx + 2, COL_REF_VERS).setValue(newRef);
    });
  }

  return { success: true };
}

// ═══════════════════════════════════════════════
//  API DISPATCHER
// ═══════════════════════════════════════════════

function handleAPI_(action, paramsStr) {
  try {
    const params = paramsStr ? JSON.parse(paramsStr) : {};
    let result = null;

    switch (action) {
      case 'getChantiers':            result = getChantiers();                                           break;
      case 'getSaisies':              result = getSaisies(params.chantier);                              break;
      case 'getLibellesForChantier':  result = getLibellesForChantier(params.chantier, params.mode);    break;
      case 'ajouterSaisie':           result = ajouterSaisie(params.chantier, params.saisie);            break;
      case 'modifierSaisie':          result = modifierSaisie(params.chantier, params.rowNum, params.saisie); break;
      case 'supprimerSaisie':         result = supprimerSaisie(params.chantier, params.rowNum);          break;
      case 'getVersements':           result = getVersements(params.chantier);                           break;
      case 'getChargesPourVersement': result = getChargesPourVersement(params.chantier);                 break;
      case 'ajouterVersement':        result = ajouterVersement(params.chantier, params.versement);      break;
      case 'supprimerVersement':      result = supprimerVersement(params.chantier, params.versementId);  break;
      default: throw new Error('Action inconnue: ' + action);
    }

    return ContentService.createTextOutput(JSON.stringify({ result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message || String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    return handleAPI_(body.action, JSON.stringify(body.params || {}));
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message || String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
