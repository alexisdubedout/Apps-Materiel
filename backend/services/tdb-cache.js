const axios = require('axios');
const ExcelJS = require('exceljs');
const { chromium } = require('playwright');
const config = require('../config');

let tdbCache = {
  data: null,
  timestamp: null,
  ttl: 10 * 60 * 1000,
};

function isBusinessHours() {
  const now = new Date();
  const hour = now.getHours();
  const day = now.getDay();
  return day >= 1 && day <= 5 && hour >= 8 && hour < 18;
}

async function downloadSharePointFile(sharePointUrl) {
  let browser;
  try {
    console.log('📥 Téléchargement SharePoint via Playwright...');
    browser = await chromium.launch({ headless: true, args: ['--no-sandbox', '--disable-setuid-sandbox'] });
    const context = await browser.newContext({ acceptDownloads: true });
    const page = await context.newPage();
    const downloadPromise = page.waitForEvent('download');
    await page.goto(sharePointUrl, { waitUntil: 'commit' }).catch(() => {});
    const download = await downloadPromise;
    const stream = await download.createReadStream();
    const chunks = [];
    for await (const chunk of stream) chunks.push(chunk);
    const fileBuffer = Buffer.concat(chunks);
    console.log('✅ Fichier téléchargé :', fileBuffer.length, 'bytes');
    return fileBuffer;
  } finally {
    if (browser) await browser.close();
  }
}

async function parseExcelBuffer(buffer, sheetName = null, skipRows = 0) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const sheet = sheetName ? workbook.getWorksheet(sheetName) : workbook.worksheets[0];
  if (!sheet) throw new Error(`Onglet "${sheetName}" non trouvé`);

  const data = [];
  const headers = {};
  const headerRowIndex = skipRows + 1;
  const headerRow = sheet.getRow(headerRowIndex);

  headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    const headerValue = cell.value?.toString().trim() || '';
    headers[colNumber] = headerValue.replace(/\n/g, ' ').replace(/\s+/g, ' ').trim() || `Col${colNumber}`;
  });

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber <= headerRowIndex) return;
    const rowData = {};
    let hasData = false;

    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const value = cell.value;
      const header = headers[colNumber] || `Col${colNumber}`;
      if (value !== null && value !== undefined && value !== '') hasData = true;

      let cleanValue = '';
      if (value !== null && value !== undefined && value !== '') {
        if (value instanceof Date) {
          cleanValue = value.toLocaleDateString('fr-FR');
        } else if (typeof value === 'object') {
          if (value.result !== undefined) {
            const result = value.result;
            if (typeof result === 'number' && (header.toLowerCase().includes('montant') || header.toLowerCase().includes('prix'))) {
              cleanValue = (Math.round(result * 100) / 100).toFixed(2);
            } else {
              cleanValue = result.toString().trim();
            }
          } else if (value.text !== undefined) {
            cleanValue = value.text.toString().trim();
          } else {
            cleanValue = JSON.stringify(value);
          }
        } else if (typeof value === 'number' && (header.toLowerCase().includes('montant') || header.toLowerCase().includes('prix'))) {
          cleanValue = (Math.round(value * 100) / 100).toFixed(2);
        } else {
          cleanValue = value.toString().trim();
        }
      }
      rowData[header] = cleanValue;
    });

    if (hasData) data.push(rowData);
  });

  return data;
}

async function refreshTDBCache() {
  try {
    console.log('🔄 Rafraîchissement automatique du cache TDB...');

    const [devisBuffer, revueHebdoBuffer] = await Promise.all([
      downloadSharePointFile(config.SHAREPOINT_DEVIS_URL),
      downloadSharePointFile(config.SHAREPOINT_REVUE_HEBDO_URL),
    ]);

    const devisData = await parseExcelBuffer(devisBuffer);
    const revueHebdoData = await parseExcelBuffer(revueHebdoBuffer, 'DSM en cours', 4);

    const dsmData = revueHebdoData.filter(row => row['ticketid'] && row['ticketid'].toString().trim() !== '');
    const opalesData = revueHebdoData.filter(row => row['N° LandesK'] && row['N° LandesK'].toString().trim() !== '');

    let redmineData = [];
    try {
      const redmineResponse = await axios.get(`${config.REDMINE_URL}/issues.json`, {
        params: { project_id: 'mco', tracker_id: 12, limit: 100 },
        headers: { 'X-Redmine-API-Key': config.REDMINE_API_KEY },
      });
      redmineData = redmineResponse.data.issues.map(issue => ({
        'ID': issue.id.toString(),
        'Titre': issue.subject,
        'Statut': issue.status.name,
        'Assigné à': issue.assigned_to?.name || 'Non assigné',
        'Echéance': issue.due_date || '',
        '_redmine_url': `${config.REDMINE_URL}/issues/${issue.id}`,
      }));
    } catch (error) {
      console.error('⚠️ Erreur Redmine:', error.message);
    }

    tdbCache.data = { devis: devisData, dsm: dsmData, opales: opalesData, redmine: redmineData, timestamp: new Date().toISOString() };
    tdbCache.timestamp = Date.now();

    console.log(`✅ Cache TDB rafraîchi — Devis: ${devisData.length} | DSM: ${dsmData.length} | OPALES: ${opalesData.length} | Redmine: ${redmineData.length}`);
  } catch (error) {
    console.error('❌ Erreur rafraîchissement cache:', error.message);
  }
}

function startCacheRefresh() {
  setInterval(() => {
    if (isBusinessHours()) refreshTDBCache();
    else console.log('⏸️ Hors heures de bureau - pas de rafraîchissement auto');
  }, 10 * 60 * 1000);

  if (isBusinessHours()) {
    console.log('🚀 Rafraîchissement initial du cache...');
    refreshTDBCache();
  }
}

module.exports = { tdbCache, refreshTDBCache, downloadSharePointFile, startCacheRefresh };
