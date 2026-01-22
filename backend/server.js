const express = require('express');
const cors = require('cors');
const multer = require('multer');
const path = require('path');
const fs = require('fs').promises;

// Import des processeurs
const { processStockTracking } = require('./processors/stock-tracking');

const app = express();
const PORT = process.env.PORT || 10000;
const FRONTEND_URL = process.env.FRONTEND_URL || 'http://localhost:5173';

// Configuration CORS
app.use(cors({
  origin: [
    FRONTEND_URL,
    'http://localhost:5173',
    'http://localhost:3000'
  ],
  credentials: true
}));

app.use(express.json());

// Configuration Multer pour upload de fichiers
const upload = multer({
  dest: 'uploads/',
  limits: { fileSize: 50 * 1024 * 1024 }, // 50MB
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext === '.xlsx' || ext === '.xls') {
      cb(null, true);
    } else {
      cb(new Error('Seuls les fichiers Excel (.xlsx, .xls) sont acceptés'));
    }
  }
});

// ===================================
// ROUTES
// ===================================

app.get('/', (req, res) => {
  res.json({
    message: 'Excel Processing API - Node.js',
    version: '2.0.0',
    status: 'running',
    endpoints: {
      treatments: '/api/treatments',
      process: '/api/process/{treatment_id}'
    }
  });
});

// Liste des traitements disponibles
app.get('/api/treatments', (req, res) => {
  res.json({
    treatments: [
      {
        id: 'stock-tracking',
        name: 'Suivi des Stocks',
        description: 'Mise à jour automatique du suivi mensuel et semestriel des stocks',
        status: 'active',
        files: [
          { id: 'tracking', label: 'Fichier de suivi', accept: '.xlsx,.xls' },
          { id: 'export', label: "Fichier d'export", accept: '.xlsx,.xls' }
        ],
        params: [
          { id: 'export_date', label: "Date d'export", type: 'date', placeholder: '' }
        ]
      },
      {
        id: 'sales-analysis',
        name: 'Analyse des Ventes',
        description: 'Génération de rapports et analyses de ventes mensuelles',
        status: 'coming_soon',
        files: [
          { id: 'sales', label: 'Fichier des ventes', accept: '.xlsx,.xls' }
        ],
        params: [
          { id: 'period', label: 'Période', type: 'text', placeholder: 'Ex: Q1 2024' }
        ]
      },
      {
        id: 'data-merge',
        name: 'Fusion de Données',
        description: 'Consolidation de plusieurs fichiers Excel en un seul',
        status: 'coming_soon',
        files: [
          { id: 'file1', label: 'Premier fichier', accept: '.xlsx,.xls' },
          { id: 'file2', label: 'Deuxième fichier', accept: '.xlsx,.xls' }
        ],
        params: []
      }
    ]
  });
});

// Traitement des fichiers
app.post('/api/process/:treatmentId', upload.any(), async (req, res) => {
  let outputPath = null;
  
  try {
    const { treatmentId } = req.params;
    const files = req.files;
    const params = JSON.parse(req.body.params || '{}');

    console.log('🚀 Traitement demandé:', treatmentId);
    console.log('📁 Fichiers reçus:', files.map(f => f.originalname));
    console.log('⚙️ Paramètres:', params);

    // Vérifier que le traitement existe et est actif
    if (treatmentId !== 'stock-tracking') {
      return res.status(400).json({
        detail: 'Ce traitement est en cours de développement. Seul "Suivi des Stocks" est disponible pour le moment.'
      });
    }

    // Mapper les fichiers uploadés par ID
    const fileMap = {};
    files.forEach(file => {
      // Le nom du champ est "file_tracking", "file_export", etc.
      const fileId = file.fieldname.replace('file_', '');
      fileMap[fileId] = file.path;
    });

    console.log('📋 Mapping des fichiers:', fileMap);

    // Vérifier que tous les fichiers requis sont présents
    const requiredFiles = ['tracking', 'export'];
    for (const fileId of requiredFiles) {
      if (!fileMap[fileId]) {
        throw new Error(`Fichier manquant: ${fileId}`);
      }
    }

    // Vérifier les paramètres requis
    if (!params.export_date) {
      throw new Error('Paramètre manquant: export_date');
    }

    // Exécuter le traitement
    outputPath = await processStockTracking(
      fileMap.tracking,
      fileMap.export,
      params.export_date
    );

    console.log('✅ Traitement terminé, fichier:', outputPath);

    // Générer le nom du fichier résultat
    const resultFilename = `resultat_stock_tracking_${params.export_date.replace(/\//g, '-')}.xlsx`;

    // Envoyer le fichier
    res.download(outputPath, resultFilename, async (err) => {
      // Nettoyer les fichiers temporaires après envoi
      try {
        await cleanupFiles([...files.map(f => f.path), outputPath]);
      } catch (cleanupErr) {
        console.error('Erreur nettoyage:', cleanupErr);
      }

      if (err) {
        console.error('Erreur envoi fichier:', err);
      }
    });

  } catch (error) {
    console.error('❌ Erreur traitement:', error);
    
    // Nettoyer les fichiers en cas d'erreur
    if (req.files) {
      try {
        await cleanupFiles(req.files.map(f => f.path));
        if (outputPath) await fs.unlink(outputPath);
      } catch (cleanupErr) {
        console.error('Erreur nettoyage après erreur:', cleanupErr);
      }
    }

    res.status(500).json({
      detail: error.message || 'Erreur lors du traitement'
    });
  }
});

// ===================================
// UTILITAIRES
// ===================================

async function cleanupFiles(filePaths) {
  for (const filePath of filePaths) {
    try {
      await fs.unlink(filePath);
      console.log('🗑️ Fichier supprimé:', filePath);
    } catch (err) {
      if (err.code !== 'ENOENT') {
        console.error('Erreur suppression:', filePath, err);
      }
    }
  }
}

// ===================================
// DÉMARRAGE
// ===================================

app.listen(PORT, () => {
  console.log('🚀 Serveur Node.js démarré');
  console.log(`📡 Port: ${PORT}`);
  console.log(`🌍 Frontend autorisé: ${FRONTEND_URL}`);
  console.log(`✅ Prêt à traiter des fichiers Excel`);
});