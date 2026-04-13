const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs').promises;

const MAPPING_FILE = path.join(__dirname, '../data/mapping-emplacement.json');

async function processTriMateriel(exportPath) {
  console.log('═══════════════════════════════════════════════════');
  console.log('🚀 DÉBUT DU TRAITEMENT TRI MATÉRIEL');
  console.log('═══════════════════════════════════════════════════');
  console.log('📁 Fichier export:', exportPath);

  if (global.gc) global.gc();

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(exportPath);
  
  const dataSheet = wb.worksheets[0];
  if (!dataSheet) {
    throw new Error('Aucune feuille trouvée dans le fichier');
  }

  console.log('📖 Fichier chargé');

  const mapping = await loadMapping();
  console.log(`📋 ${Object.keys(mapping).length} emplacements dans le mapping`);

  const dateStr = formatDate(new Date(), 'DD-MM-YY');
  let sheetName = `Export ${dateStr}`;
  let suffix = 1;
  
  while (wb.getWorksheet(sheetName)) {
    suffix++;
    sheetName = `Export ${dateStr} (${suffix})`;
  }

  const newSheet = wb.addWorksheet(sheetName);
  console.log(`📄 Nouvel onglet créé: ${sheetName}`);

  await copyAndEnrichData(dataSheet, newSheet, mapping);

  newSheet.columns.forEach((col, idx) => {
    if (idx < 9) {
      col.outlineLevel = 1;
    }
  });


  const outputPath = exportPath.replace('.xlsx', `_traite_${Date.now()}.xlsx`);
  await wb.xlsx.writeFile(outputPath);

  console.log('═══════════════════════════════════════════════════');
  console.log('🎉 TRAITEMENT TERMINÉ');
  console.log('═══════════════════════════════════════════════════');

  return outputPath;
}

async function copyAndEnrichData(sourceSheet, targetSheet, mapping) {
  console.log('📊 Copie et enrichissement des données...');

  const headers = [
    'Code actif',
    'Description de l\'actif',
    'Code article',
    'Code produit',
    'N° de série',
    'Emplacement',
    'Description de l\'emplacement',
    'Type emplacement',
    'Détail'
  ];

  const headerRow = targetSheet.addRow(headers);
  headerRow.eachCell(cell => {
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
  });

  let processedRows = 0;
  const BATCH_SIZE = 1000;
  const rowsToAdd = [];

  sourceSheet.eachRow((row, rowNum) => {
    if (rowNum === 1) return;

    const codeActif = getCellValue(row, 1);
    const descActif = getCellValue(row, 2);
    const codeArticle = getCellValue(row, 3);
    const codeProduit = getCellValue(row, 5);
    const numSerie = getCellValue(row, 6);
    const emplacement = getCellValue(row, 7);
    const descEmplacement = getCellValue(row, 8);

    const mappingData = mapping[emplacement] || { typeEmplacement: '', detail: '' };

    rowsToAdd.push([
      codeActif,
      descActif,
      codeArticle,
      codeProduit,
      numSerie,
      emplacement,
      descEmplacement,
      mappingData.typeEmplacement || '',
      mappingData.detail || ''
    ]);

    if (rowsToAdd.length >= BATCH_SIZE) {
      rowsToAdd.forEach(rowData => targetSheet.addRow(rowData));
      processedRows += rowsToAdd.length;
      console.log(`  ✓ ${processedRows} lignes traitées...`);
      rowsToAdd.length = 0;
      if (global.gc) global.gc();
    }
  });

  if (rowsToAdd.length > 0) {
    rowsToAdd.forEach(rowData => targetSheet.addRow(rowData));
    processedRows += rowsToAdd.length;
  }

  console.log(`✅ ${processedRows} lignes traitées au total`);

  targetSheet.columns.forEach(col => {
    col.width = 20;
  });
}


async function loadMapping() {
  try {
    const data = await fs.readFile(MAPPING_FILE, 'utf8');
    return JSON.parse(data);
  } catch (error) {
    console.warn('⚠️ Mapping non trouvé');
    return {};
  }
}

function getCellValue(row, colIndex) {
  const cell = row.getCell(colIndex);
  return cell.value || '';
}

function formatDate(date, format) {
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = String(date.getFullYear()).slice(-2);

  if (format === 'DD-MM-YY') {
    return `${day}-${month}-${year}`;
  }
  return date.toISOString();
}

module.exports = { processTriMateriel };