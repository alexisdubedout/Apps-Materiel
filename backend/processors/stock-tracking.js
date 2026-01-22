// ===================================
// 📊 Stock Tracking Processor
// Version Node.js avec ExcelJS - Optimisé pour mémoire limitée
// ===================================

const ExcelJS = require('exceljs');
const path = require('path');

/**
 * Traite le suivi des stocks
 * @param {string} trackingPath - Chemin du fichier de suivi
 * @param {string} exportPath - Chemin du fichier d'export
 * @param {string} exportDateStr - Date au format 'YYYY-MM-DD' ou 'DD/MM/YYYY'
 * @returns {Promise<string>} - Chemin du fichier traité
 */
async function processStockTracking(trackingPath, exportPath, exportDateStr) {
  console.log('═══════════════════════════════════════════════════');
  console.log('🚀 DÉBUT DU TRAITEMENT STOCK TRACKING');
  console.log('═══════════════════════════════════════════════════');
  console.log('📁 Fichier tracking:', trackingPath);
  console.log('📁 Fichier export:', exportPath);
  console.log('📅 Date brute:', exportDateStr);

  // Forcer le garbage collector si disponible
  if (global.gc) {
    global.gc();
  }

  // Parser la date (accepter les deux formats)
  const exportDate = parseDate(exportDateStr);
  const exportDateFormatted = formatDate(exportDate, 'DD/MM/YYYY');
  console.log('✅ Date formatée:', exportDateFormatted);

  // Charger les deux workbooks UNE SEULE FOIS
  const trackingWb = new ExcelJS.Workbook();
  const exportWb = new ExcelJS.Workbook();
  
  await trackingWb.xlsx.readFile(trackingPath);
  await exportWb.xlsx.readFile(exportPath);
  
  console.log('📖 Workbooks chargés en mémoire');

  try {
    // Étape 1: Mise à jour du tracking principal
    console.log('📊 Étape 1/3: Update tracking...');
    await updateTracking(trackingWb, exportWb, exportDateFormatted);
    console.log('✅ Update tracking terminé');

    // Libérer la mémoire du workbook export
    exportWb.worksheets.forEach(sheet => {
      sheet.destroy && sheet.destroy();
    });

    // Forcer le garbage collector
    if (global.gc) {
      global.gc();
    }

    // Étape 2: Mise à jour suivi mensuel
    console.log('📊 Étape 2/3: Update monthly tracking...');
    await updateMonthlyTracking(trackingWb, exportDateFormatted);
    console.log('✅ Update monthly tracking terminé');

    if (global.gc) {
      global.gc();
    }

    // Étape 3: Mise à jour suivi semestriel
    console.log('📊 Étape 3/3: Update semestrial tracking...');
    await updateSemestrialTracking(trackingWb, exportDateFormatted);
    console.log('✅ Update semestrial tracking terminé');

    // Sauvegarder UNE SEULE FOIS
    console.log('💾 Sauvegarde du fichier...');
    await trackingWb.xlsx.writeFile(trackingPath);
    console.log('💾 Fichier sauvegardé');

    console.log('═══════════════════════════════════════════════════');
    console.log('🎉 TRAITEMENT TERMINÉ AVEC SUCCÈS');
    console.log('═══════════════════════════════════════════════════');

    return trackingPath;

  } catch (error) {
    console.error('❌ ERREUR:', error);
    throw error;
  }
}

// ===================================
// ÉTAPE 1: UPDATE TRACKING - OPTIMISÉ
// ===================================

async function updateTracking(trackingWb, exportWb, exportDate) {
  const stockSheet = trackingWb.getWorksheet('Liste de Stock');
  const exportSheet = exportWb.worksheets[0];

  if (!stockSheet) {
    throw new Error('Feuille "Liste de Stock" introuvable');
  }

  // Vérifier si la date existe déjà
  const headerRow = stockSheet.getRow(1);
  let dateColIndex = null;
  
  headerRow.eachCell((cell, colNum) => {
    if (cell.value === exportDate) {
      dateColIndex = colNum;
    }
  });

  if (dateColIndex) {
    throw new Error(`Les données pour la date ${exportDate} ont déjà été importées`);
  }

  // Extraire et regrouper les données d'export
  const exportData = extractExportData(exportSheet);
  console.log(`📦 ${exportData.size} articles uniques dans l'export`);

  // Ajouter la nouvelle colonne de date
  const newColIndex = headerRow.cellCount + 1;
  const newColCell = headerRow.getCell(newColIndex);
  newColCell.value = exportDate;
  styleHeaderCell(newColCell);

  // Mettre à jour les lignes existantes
  let updatedRows = 0;
  const existingKeys = new Set();

  stockSheet.eachRow((row, rowNum) => {
    if (rowNum === 1) return; // Skip header

    const codif = row.getCell(1).value;
    const magasin = row.getCell(3).value;
    const key = `${codif}|${magasin}`;
    
    existingKeys.add(key);

    if (exportData.has(key)) {
      row.getCell(newColIndex).value = exportData.get(key).quantite;
      updatedRows++;
    } else {
      row.getCell(newColIndex).value = 0;
    }
  });

  console.log(`✏️ ${updatedRows} lignes mises à jour`);

  // Ajouter les nouvelles lignes PAR LOTS (OPTIMISATION)
  const newRowsToAdd = [];
  exportData.forEach((data, key) => {
    if (!existingKeys.has(key)) {
      newRowsToAdd.push({ data, key, newColIndex });
    }
  });

  console.log(`➕ Ajout de ${newRowsToAdd.length} nouvelles lignes par lots...`);

  const BATCH_SIZE = 1000;
  let totalAdded = 0;

  for (let i = 0; i < newRowsToAdd.length; i += BATCH_SIZE) {
    const batch = newRowsToAdd.slice(i, i + BATCH_SIZE);
    
    batch.forEach(item => {
      const newRow = stockSheet.addRow([
        item.data.codeArticle,
        item.data.description,
        item.data.emplacement,
        item.data.descEmplacement
      ]);
      newRow.getCell(item.newColIndex).value = item.data.quantite;
    });

    totalAdded += batch.length;
    console.log(`  ✓ ${totalAdded}/${newRowsToAdd.length} lignes ajoutées...`);

    // Forcer le garbage collector tous les 5000 lignes
    if (totalAdded % 5000 === 0 && global.gc) {
      global.gc();
    }
  }

  console.log(`➕ ${totalAdded} nouvelles lignes ajoutées au total`);

  // Ajuster les largeurs de colonnes
  console.log('📐 Ajustement des colonnes...');
  adjustColumnWidths(stockSheet);
}

// ===================================
// ÉTAPE 2: UPDATE MONTHLY TRACKING
// ===================================

async function updateMonthlyTracking(workbook, exportDate) {
  const stockSheet = workbook.getWorksheet('Liste de Stock');
  let monthlySheet = findSheetByPrefix(workbook, 'suivi mensuel');

  if (!monthlySheet) {
    console.warn('⚠️ Feuille "Suivi Mensuel" introuvable, création...');
    monthlySheet = workbook.addWorksheet('Suivi Mensuel');
  } else {
    // Recréer la feuille proprement
    const index = workbook.worksheets.indexOf(monthlySheet);
    workbook.removeWorksheet(monthlySheet.id);
    monthlySheet = workbook.addWorksheet('Suivi Mensuel');
  }

  // Récupérer les headers
  const headers = getHeaders(stockSheet);
  const currentIndex = headers.indexOf(exportDate);

  if (currentIndex < 1) {
    addNoDataMessage(monthlySheet, 'Pas de données disponibles pour le mois', exportDate);
    return;
  }

  const previousDate = headers[currentIndex - 1];

  // Vérifier que c'est une date valide
  if (!isValidDate(previousDate)) {
    addNoDataMessage(monthlySheet, 'Pas de données disponibles pour le mois', exportDate);
    return;
  }

  // Calculer les variations
  const variations = calculateVariations(stockSheet, currentIndex + 1, currentIndex);

  // Écrire les résultats
  writeVariationsSheet(
    monthlySheet,
    variations,
    `Variation entre le ${previousDate} et le ${exportDate}`,
    ['Codification DSNA', 'Désignation', 'Magasin', 'Description', 'Variation', 'Quantité actuelle']
  );
}

// ===================================
// ÉTAPE 3: UPDATE SEMESTRIAL TRACKING
// ===================================

async function updateSemestrialTracking(workbook, exportDate) {
  const stockSheet = workbook.getWorksheet('Liste de Stock');
  let semestrialSheet = findSheetByPrefix(workbook, 'suivi semestriel');

  if (!semestrialSheet) {
    console.warn('⚠️ Feuille "Suivi Semestriel" introuvable, création...');
    semestrialSheet = workbook.addWorksheet('Suivi Semestriel');
  } else {
    // Recréer la feuille proprement
    const index = workbook.worksheets.indexOf(semestrialSheet);
    workbook.removeWorksheet(semestrialSheet.id);
    semestrialSheet = workbook.addWorksheet('Suivi Semestriel');
  }

  const headers = getHeaders(stockSheet);
  const currentIndex = headers.indexOf(exportDate);

  if (currentIndex < 6) {
    addNoDataMessage(semestrialSheet, 'Pas de données disponibles pour le semestre', exportDate);
    return;
  }

  const previousDate = headers[currentIndex - 6];

  if (!isValidDate(previousDate)) {
    addNoDataMessage(semestrialSheet, 'Pas de données disponibles pour le semestre', exportDate);
    return;
  }

  const variations = calculateVariations(stockSheet, currentIndex + 1, currentIndex - 5);

  writeVariationsSheet(
    semestrialSheet,
    variations,
    `Variation entre le ${previousDate} et le ${exportDate}`,
    ['Codification DSNA', 'Désignation', 'Magasin', 'Description', 'Variation', 'Quantité actuelle']
  );
}

// ===================================
// FONCTIONS UTILITAIRES
// ===================================

function extractExportData(exportSheet) {
  const data = new Map();
  
  exportSheet.eachRow((row, rowNum) => {
    if (rowNum === 1) return; // Skip header

    const codeArticle = getCellValue(row, 1);
    const emplacement = getCellValue(row, 2);
    const description = getCellValue(row, 3);
    const descEmplacement = getCellValue(row, 4);

    if (!codeArticle || !emplacement) return;

    const key = `${codeArticle}|${emplacement}`;

    if (!data.has(key)) {
      data.set(key, {
        codeArticle,
        emplacement,
        description: description || '',
        descEmplacement: descEmplacement || '',
        quantite: 0
      });
    }

    data.get(key).quantite++;
  });

  return data;
}

function calculateVariations(stockSheet, currentColIndex, previousColIndex) {
  const variations = [];

  stockSheet.eachRow((row, rowNum) => {
    if (rowNum === 1) return; // Skip header

    const currentQty = getCellValue(row, currentColIndex) || 0;
    const previousQty = getCellValue(row, previousColIndex) || 0;
    const variation = currentQty - previousQty;

    if (variation !== 0) {
      variations.push({
        codif: getCellValue(row, 1),
        designation: getCellValue(row, 2),
        magasin: getCellValue(row, 3),
        description: getCellValue(row, 4),
        variation,
        qtyActuelle: currentQty
      });
    }
  });

  return variations;
}

function writeVariationsSheet(sheet, variations, title, headers) {
  // Titre
  sheet.mergeCells('A1:F1');
  const titleCell = sheet.getCell('A1');
  titleCell.value = title;
  titleCell.alignment = { horizontal: 'center' };
  titleCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFD3D3D3' }
  };

  // Headers
  const headerRow = sheet.getRow(2);
  headers.forEach((header, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = header;
    styleHeaderCell(cell);
  });

  // Données
  if (variations.length > 0) {
    variations.forEach((item, i) => {
      const row = sheet.getRow(i + 3);
      row.getCell(1).value = item.codif;
      row.getCell(2).value = item.designation;
      row.getCell(3).value = item.magasin;
      row.getCell(4).value = item.description;
      row.getCell(5).value = item.variation;
      row.getCell(6).value = item.qtyActuelle;

      // Coloration selon quantité
      const qtyCell = row.getCell(6);
      if (item.qtyActuelle <= 5) {
        qtyCell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFCCCC' }
        };
      } else if (item.qtyActuelle <= 10) {
        qtyCell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFDAB9' }
        };
      }
    });
  } else {
    sheet.mergeCells('A3:F3');
    const cell = sheet.getCell('A3');
    cell.value = 'Aucune variation pour cette période';
    cell.alignment = { horizontal: 'center' };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFD3D3D3' }
    };
  }

  adjustColumnWidths(sheet);
}

function addNoDataMessage(sheet, message, dateContext) {
  sheet.mergeCells('A1:F1');
  const cell = sheet.getCell('A1');
  cell.value = message;
  cell.alignment = { horizontal: 'center' };
  cell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFD3D3D3' }
  };
}

function styleHeaderCell(cell) {
  cell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF003366' }
  };
  cell.font = {
    color: { argb: 'FFFFFFFF' },
    bold: true
  };
  cell.alignment = { horizontal: 'center' };
  cell.border = {
    top: { style: 'thin', color: { argb: 'FFFFFFFF' } },
    left: { style: 'thin', color: { argb: 'FFFFFFFF' } },
    bottom: { style: 'thin', color: { argb: 'FFFFFFFF' } },
    right: { style: 'thin', color: { argb: 'FFFFFFFF' } }
  };
}

function clearWorksheet(sheet) {
  // Supprimer la feuille et la recréer (plus propre)
  const workbook = sheet.workbook;
  const sheetName = sheet.name;
  const index = workbook.worksheets.indexOf(sheet);
  
  workbook.removeWorksheet(sheet.id);
  return workbook.addWorksheet(sheetName, { views: [{}] }, index);
}

function adjustColumnWidths(sheet) {
  sheet.columns.forEach(column => {
    let maxLength = 10;
    column.eachCell({ includeEmpty: false }, cell => {
      const length = cell.value ? cell.value.toString().length : 10;
      if (length > maxLength) {
        maxLength = length;
      }
    });
    column.width = Math.min(maxLength + 2, 50);
  });
}

function getHeaders(sheet) {
  const headers = [];
  const headerRow = sheet.getRow(1);
  headerRow.eachCell(cell => {
    headers.push(cell.value);
  });
  return headers;
}

function findSheetByPrefix(workbook, prefix) {
  const lowerPrefix = prefix.toLowerCase();
  return workbook.worksheets.find(sheet => 
    sheet.name.toLowerCase().startsWith(lowerPrefix)
  );
}

function getCellValue(row, colIndex) {
  const cell = row.getCell(colIndex);
  return cell.value;
}

function parseDate(dateStr) {
  // Essayer format ISO (YYYY-MM-DD)
  let match = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (match) {
    return new Date(match[1], match[2] - 1, match[3]);
  }

  // Essayer format FR (DD/MM/YYYY)
  match = dateStr.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (match) {
    return new Date(match[3], match[2] - 1, match[1]);
  }

  throw new Error(`Format de date invalide: ${dateStr}. Utilisez DD/MM/YYYY ou YYYY-MM-DD`);
}

function formatDate(date, format) {
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();

  if (format === 'DD/MM/YYYY') {
    return `${day}/${month}/${year}`;
  }
  return `${year}-${month}-${day}`;
}

function isValidDate(dateStr) {
  try {
    parseDate(dateStr);
    return true;
  } catch {
    return false;
  }
}

module.exports = { processStockTracking };
