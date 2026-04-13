const { dbRun, dbGet } = require('./database');

async function initNotesTable() {
  await dbRun(`
    CREATE TABLE IF NOT EXISTS tdb_notes (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      content TEXT NOT NULL,
      updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
      updated_by TEXT
    )
  `);
  console.log('✅ Table notes prête');

  const row = await dbGet('SELECT COUNT(*) as count FROM tdb_notes');
  if (row.count === 0) {
    await dbRun(
      'INSERT INTO tdb_notes (content, updated_by) VALUES (?, ?)',
      ['Tableau de bord MCO Matériel - Suivi en temps réel', 'System']
    );
    console.log('✅ Note par défaut créée');
  }
}

async function getNote() {
  return dbGet('SELECT * FROM tdb_notes ORDER BY id DESC LIMIT 1');
}

async function saveNote(content, updatedBy) {
  const result = await dbRun(
    'INSERT INTO tdb_notes (content, updated_by) VALUES (?, ?)',
    [content, updatedBy]
  );
  return { id: result.lastID, content, updated_by: updatedBy };
}

module.exports = { initNotesTable, getNote, saveNote };
