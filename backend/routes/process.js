const express = require('express');
const router = express.Router();
const multer = require('multer');
const path = require('path');
const fs = require('fs').promises;
const { authenticateToken, rejectClient } = require('../middleware/auth');
const { processStockTracking } = require('../processors/stock-tracking');
const { processTriMateriel } = require('../processors/tri-materiel');

const upload = multer({
  dest: 'uploads/',
  limits: { fileSize: 50 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext === '.xlsx' || ext === '.xls') cb(null, true);
    else cb(new Error('Seuls les fichiers Excel (.xlsx, .xls) sont acceptés'));
  },
});

router.get('/', (req, res) => {
  res.json({
    treatments: [
      {
        id: 'stock-tracking',
        name: 'Suivi des Stocks',
        description: 'Mise à jour automatique du suivi mensuel et semestriel des stocks',
        status: 'active',
        files: [
          { id: 'tracking', label: 'Fichier de suivi', accept: '.xlsx,.xls' },
          { id: 'export', label: "Fichier d'export", accept: '.xlsx,.xls' },
        ],
        params: [{ id: 'export_date', label: "Date d'export", type: 'date', placeholder: '' }],
      },
      {
        id: 'tri-materiel',
        name: 'Tri Matériel',
        description: 'Enrichissement automatique des données avec mapping',
        status: 'active',
        files: [{ id: 'export', label: 'Fichier export brut', accept: '.xlsx,.xls' }],
        params: [],
      },
    ],
  });
});

router.post('/:treatmentId', authenticateToken, rejectClient, upload.any(), async (req, res) => {
  let outputPath = null;
  const uploadedFiles = [];

  try {
    const { treatmentId } = req.params;
    const files = req.files;
    const params = JSON.parse(req.body.params || '{}');

    console.log('🚀 Traitement:', treatmentId);
    console.log('📁 Fichiers:', files.map(f => f.originalname));

    const fileMap = {};
    files.forEach(file => {
      const fileId = file.fieldname.replace('file_', '');
      fileMap[fileId] = file.path;
      uploadedFiles.push(file.path);
    });

    if (treatmentId === 'stock-tracking') {
      if (!fileMap.tracking || !fileMap.export) throw new Error('Fichiers manquants');
      if (!params.export_date) throw new Error('Date manquante');
      outputPath = await processStockTracking(fileMap.tracking, fileMap.export, params.export_date);
      const safeDateStr = params.export_date.replace(/[\/\\:]/g, '-');
      const fileBuffer = await fs.readFile(outputPath);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', `attachment; filename="resultat_stock_tracking_${safeDateStr}.xlsx"`);
      res.send(fileBuffer);

    } else if (treatmentId === 'tri-materiel') {
      if (!fileMap.export) throw new Error('Fichier export manquant');
      outputPath = await processTriMateriel(fileMap.export);
      const fileBuffer = await fs.readFile(outputPath);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', `attachment; filename="tri_materiel_${Date.now()}.xlsx"`);
      res.send(fileBuffer);

    } else {
      throw new Error('Traitement inconnu');
    }

    setTimeout(async () => {
      try {
        for (const filePath of uploadedFiles) await fs.unlink(filePath);
        if (outputPath) await fs.unlink(outputPath);
        console.log('🗑️ Fichiers nettoyés');
      } catch (err) {
        console.error('Erreur nettoyage:', err);
      }
    }, 1000);

  } catch (error) {
    console.error('❌ Erreur:', error.message);
    try {
      for (const filePath of uploadedFiles) await fs.unlink(filePath).catch(() => {});
      if (outputPath) await fs.unlink(outputPath).catch(() => {});
    } catch {}
    res.status(500).json({ detail: error.message || 'Erreur lors du traitement' });
  }
});

module.exports = router;
