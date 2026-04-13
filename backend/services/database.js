const sqlite3 = require('sqlite3').verbose();
const path = require('path');

const DB_PATH = path.join(__dirname, '../database/users.db');

// Persistent connection with WAL mode
const db = new sqlite3.Database(DB_PATH, (err) => {
  if (err) {
    console.error('Erreur ouverture DB:', err);
    process.exit(1);
  }
});

db.run('PRAGMA journal_mode = WAL');
db.run('PRAGMA foreign_keys = ON');

// Promise wrappers
function dbRun(sql, params = []) {
  return new Promise((resolve, reject) => {
    db.run(sql, params, function (err) {
      if (err) reject(err);
      else resolve({ lastID: this.lastID, changes: this.changes });
    });
  });
}

function dbGet(sql, params = []) {
  return new Promise((resolve, reject) => {
    db.get(sql, params, (err, row) => {
      if (err) reject(err);
      else resolve(row);
    });
  });
}

function dbAll(sql, params = []) {
  return new Promise((resolve, reject) => {
    db.all(sql, params, (err, rows) => {
      if (err) reject(err);
      else resolve(rows);
    });
  });
}

// Graceful shutdown
function closeDb() {
  db.close((err) => {
    if (err) console.error('Erreur fermeture DB:', err);
    else console.log('✅ DB fermée proprement');
  });
}
process.on('SIGINT', () => { closeDb(); process.exit(0); });
process.on('SIGTERM', () => { closeDb(); process.exit(0); });

// ─── Init tables ───────────────────────────────────────────────────────────

async function initDatabase() {
  await dbRun(`
    CREATE TABLE IF NOT EXISTS users (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      redmine_id INTEGER UNIQUE NOT NULL,
      login TEXT UNIQUE NOT NULL,
      email TEXT,
      firstname TEXT,
      lastname TEXT,
      last_login DATETIME,
      created_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )
  `);
  console.log('✅ Table users prête');
}

async function initHistoryTable() {
  await dbRun(`
    CREATE TABLE IF NOT EXISTS mapping_history (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      emplacement_code TEXT NOT NULL,
      action TEXT NOT NULL,
      field_changed TEXT,
      old_value TEXT,
      new_value TEXT,
      user_login TEXT NOT NULL,
      user_name TEXT,
      created_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )
  `);
  console.log('✅ Table mapping_history prête');
}

// ─── User functions ────────────────────────────────────────────────────────

async function upsertUser(userData) {
  return dbRun(
    `INSERT INTO users (redmine_id, login, email, firstname, lastname, last_login)
     VALUES (?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
     ON CONFLICT(redmine_id) DO UPDATE SET
       login = excluded.login, email = excluded.email,
       firstname = excluded.firstname, lastname = excluded.lastname,
       last_login = CURRENT_TIMESTAMP`,
    [userData.id, userData.login, userData.email, userData.firstname, userData.lastname]
  );
}

async function getUserByRedmineId(redmineId) {
  return dbGet('SELECT * FROM users WHERE redmine_id = ?', [redmineId]);
}

// ─── History functions ─────────────────────────────────────────────────────

async function addHistoryEntry(entry) {
  return dbRun(
    `INSERT INTO mapping_history
     (emplacement_code, action, field_changed, old_value, new_value, user_login, user_name)
     VALUES (?, ?, ?, ?, ?, ?, ?)`,
    [
      entry.emplacementCode, entry.action,
      entry.fieldChanged || null, entry.oldValue || null,
      entry.newValue || null, entry.userLogin, entry.userName
    ]
  );
}

async function getHistory(limit = 100) {
  return dbAll(
    'SELECT * FROM mapping_history ORDER BY created_at DESC LIMIT ?',
    [limit]
  );
}

async function getHistoryByCode(code, limit = 50) {
  return dbAll(
    'SELECT * FROM mapping_history WHERE emplacement_code = ? ORDER BY created_at DESC LIMIT ?',
    [code, limit]
  );
}

module.exports = {
  db, dbRun, dbGet, dbAll,
  initDatabase, initHistoryTable,
  upsertUser, getUserByRedmineId,
  addHistoryEntry, getHistory, getHistoryByCode,
};
