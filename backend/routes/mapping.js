const express = require('express');
const router = express.Router();
const fs = require('fs').promises;
const path = require('path');
const { addHistoryEntry, getHistory, getHistoryByCode } = require('../services/database');
const { validate } = require('../middleware/validate');
const { mappingSchema } = require('../schemas/index');

const MAPPING_FILE = path.join(__dirname, '../data/mapping-emplacement.json');

let mapping = {};

async function loadMapping() {
  try {
    const data = await fs.readFile(MAPPING_FILE, 'utf8');
    mapping = JSON.parse(data);
    console.log(`📋 Mapping chargé: ${Object.keys(mapping).length} emplacements`);
  } catch (error) {
    console.warn('⚠️ Mapping non trouvé, création d\'un mapping vide');
    mapping = {};
  }
}

async function saveMapping(data) {
  await fs.writeFile(MAPPING_FILE, JSON.stringify(data, null, 2), 'utf8');
}

loadMapping();

router.get('/emplacements', async (req, res) => {
  try {
    res.json(mapping);
  } catch (error) {
    console.error('Erreur récupération mapping:', error);
    res.status(500).json({ detail: 'Erreur serveur' });
  }
});

router.post('/emplacements', validate(mappingSchema), async (req, res) => {
  try {
    const { code, description, type, typeEmplacement, detail, user } = req.body;
    
    const oldMapping = mapping[code] || null;
    
    mapping[code] = {
      description: description || '',
      type: type || '',
      typeEmplacement: typeEmplacement || '',
      detail: detail || ''
    };

    await saveMapping(mapping);

    const userName = req.user ? req.user.name : user;
    const userLogin = req.user ? req.user.login : 'unknown';

    if (!oldMapping) {
      await addHistoryEntry({
        emplacementCode: code,
        action: 'CREATE',
        fieldChanged: null,
        oldValue: null,
        newValue: JSON.stringify(mapping[code]),
        userLogin: userLogin,
        userName: userName
      });
    } else {
      const fields = ['description', 'type', 'typeEmplacement', 'detail'];
      
      for (const field of fields) {
        if (oldMapping[field] !== mapping[code][field]) {
          await addHistoryEntry({
            emplacementCode: code,
            action: 'UPDATE',
            fieldChanged: field,
            oldValue: oldMapping[field] || '',
            newValue: mapping[code][field] || '',
            userLogin: userLogin,
            userName: userName
          });
        }
      }
    }

    res.json({ success: true, data: mapping[code] });
  } catch (error) {
    console.error('Erreur sauvegarde mapping:', error);
    res.status(500).json({ detail: 'Erreur serveur' });
  }
});

router.delete('/emplacements/:code', async (req, res) => {
  try {
    const { code } = req.params;
    const { user } = req.body;

    if (!mapping[code]) {
      return res.status(404).json({ detail: 'Emplacement non trouvé' });
    }

    const oldValue = { ...mapping[code] };
    delete mapping[code];
    await saveMapping(mapping);

    const userName = req.user ? req.user.name : user;
    const userLogin = req.user ? req.user.login : 'unknown';

    await addHistoryEntry({
      emplacementCode: code,
      action: 'DELETE',
      fieldChanged: null,
      oldValue: JSON.stringify(oldValue),
      newValue: null,
      userLogin: userLogin,
      userName: userName
    });

    res.json({ success: true });
  } catch (error) {
    console.error('Erreur suppression:', error);
    res.status(500).json({ detail: 'Erreur serveur' });
  }
});

router.get('/values/:field', async (req, res) => {
  try {
    const { field } = req.params;
    const values = new Set();

    Object.values(mapping).forEach(item => {
      if (item[field]) {
        values.add(item[field]);
      }
    });

    res.json(Array.from(values).sort());
  } catch (error) {
    console.error('Erreur récupération valeurs:', error);
    res.status(500).json({ detail: 'Erreur serveur' });
  }
});

router.get('/history', async (req, res) => {
  try {
    const limit = parseInt(req.query.limit) || 100;
    const history = await getHistory(limit);
    res.json(history);
  } catch (error) {
    console.error('Erreur récupération historique:', error);
    res.status(500).json({ detail: 'Erreur serveur' });
  }
});

router.get('/history/:code', async (req, res) => {
  try {
    const { code } = req.params;
    const limit = parseInt(req.query.limit) || 50;
    const history = await getHistoryByCode(code, limit);
    res.json(history);
  } catch (error) {
    console.error('Erreur récupération historique:', error);
    res.status(500).json({ detail: 'Erreur serveur' });
  }
});

module.exports = router;