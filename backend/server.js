const express = require('express');
const cors = require('cors');
const rateLimit = require('express-rate-limit');
const config = require('./config');

const authRoutes = require('./routes/auth');
const mappingRoutes = require('./routes/mapping');
const tdbRoutes = require('./routes/tdb');
const processRoutes = require('./routes/process');

const { authenticateToken, rejectClient } = require('./middleware/auth');
const { initDatabase, initHistoryTable } = require('./services/database');
const { initNotesTable, getNote, saveNote } = require('./services/notes');
const { startCacheRefresh } = require('./services/tdb-cache');

// ─── App setup ─────────────────────────────────────────────────────────────

const app = express();
app.set('trust proxy', 1);

const FRONTEND_URL = process.env.VITE_API_URL || 'http://localhost:5173';

app.use(cors({
  origin: [FRONTEND_URL, ...config.CORS_ORIGINS],
  credentials: true,
}));

app.use(express.json());

// Rate limiting
const apiLimiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 100,
  message: { detail: 'Trop de requêtes, veuillez réessayer dans 15 minutes' },
  standardHeaders: true,
  legacyHeaders: false,
});
app.use('/api/', apiLimiter);

// ─── Routes ────────────────────────────────────────────────────────────────

app.get('/', (req, res) => res.json({ message: 'Excel Processing API - Node.js', version: '2.0.0', status: 'running' }));

app.use('/api/auth', authRoutes);
app.use('/api/mapping', authenticateToken, rejectClient, mappingRoutes);
app.use('/api/tdb', tdbRoutes);
app.use('/api/process', processRoutes);
app.use('/api/treatments', (req, res) => res.redirect('/api/process'));

// ─── Notes (legacy /api/notes endpoints) ───────────────────────────────────

app.get('/api/notes', authenticateToken, async (req, res) => {
  try {
    const note = await getNote();
    res.json(note || { content: '', updated_by: null, updated_at: null });
  } catch (error) {
    res.status(500).json({ error: 'Erreur lors de la récupération de la note' });
  }
});

app.post('/api/notes', authenticateToken, async (req, res) => {
  try {
    const { content } = req.body;
    const username = req.user ? `${req.user.firstname} ${req.user.lastname}` : 'Anonyme';
    const updatedNote = await saveNote(content, username);
    res.json(updatedNote);
  } catch (error) {
    res.status(500).json({ error: 'Erreur lors de la sauvegarde de la note' });
  }
});

// ─── Init DB & start ───────────────────────────────────────────────────────

async function start() {
  await initDatabase();
  await initHistoryTable();
  await initNotesTable();
  startCacheRefresh();

  const PORT = config.PORT;
  app.listen(PORT, () => {
    console.log('🚀 Serveur Node.js démarré');
    console.log(`📡 Port: ${PORT}`);
    console.log('✅ Prêt à traiter des fichiers Excel');
  });
}

start().catch(err => {
  console.error('❌ Erreur démarrage:', err);
  process.exit(1);
});
