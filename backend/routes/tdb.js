const express = require('express');
const router = express.Router();
const { authenticateToken } = require('../middleware/auth');
const { tdbCache, refreshTDBCache, downloadSharePointFile } = require('../services/tdb-cache');
const { getNote, saveNote } = require('../services/notes');
const { validate } = require('../middleware/validate');
const { noteSchema } = require('../schemas/index');
const config = require('../config');

// Fetch parsed TDB data (with cache)
router.get('/fetch-and-parse', authenticateToken, async (req, res) => {
  try {
    if (tdbCache.data && tdbCache.timestamp) {
      const age = Math.floor((Date.now() - tdbCache.timestamp) / 1000 / 60);
      console.log(`✅ Données servies depuis le cache (${age} min)`);
      return res.json(tdbCache.data);
    }

    console.log('⚠️ Cache vide - rafraîchissement immédiat...');
    await refreshTDBCache();

    if (tdbCache.data) {
      res.json(tdbCache.data);
    } else {
      throw new Error('Impossible de récupérer les données');
    }
  } catch (error) {
    console.error('❌ Erreur TDB:', error);
    res.status(500).json({ error: error.message, detail: 'Erreur lors de la récupération des données TDB' });
  }
});

// Legacy endpoint kept for compatibility
router.get('/fetch-excel', authenticateToken, async (req, res) => {
  try {
    const fileBuffer = await downloadSharePointFile(config.SHAREPOINT_DEVIS_URL);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(fileBuffer);
  } catch (error) {
    console.error('❌ Erreur:', error);
    res.status(500).json({ error: error.message });
  }
});

// Get notes
router.get('/notes', authenticateToken, async (req, res) => {
  try {
    const note = await getNote();
    res.json(note || { content: '', updated_at: null, updated_by: null });
  } catch (error) {
    console.error('❌ Erreur récupération notes:', error);
    res.status(500).json({ error: 'Erreur lors de la récupération des notes' });
  }
});

// Save notes
router.post('/notes', authenticateToken, validate(noteSchema), async (req, res) => {
  try {
    const { content } = req.body;
    const updatedBy = req.user?.name || req.user?.login || 'Utilisateur';
    const result = await saveNote(content, updatedBy);
    res.json(result);
  } catch (error) {
    console.error('❌ Erreur sauvegarde notes:', error);
    res.status(500).json({ error: 'Erreur lors de la sauvegarde des notes' });
  }
});

module.exports = router;
