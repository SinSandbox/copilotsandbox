const path = require('path');
const express = require('express');
const sqlite3 = require('sqlite3').verbose();
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;
const DB_DIR = path.join(__dirname, 'data');
const DB_PATH = path.join(DB_DIR, 'users.db');

// ensure data directory exists
if (!fs.existsSync(DB_DIR)) fs.mkdirSync(DB_DIR, { recursive: true });

// open (or create) sqlite database
const db = new sqlite3.Database(DB_PATH, (err) => {
  if (err) {
    console.error('Failed to open DB', err);
    process.exit(1);
  }
  console.log('Connected to SQLite DB at', DB_PATH);
});

// create table if it doesn't exist
db.run(`
  CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    email TEXT NOT NULL,
    message TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
  )
`, (err) => {
  if (err) {
    console.error('Failed to create table', err);
    process.exit(1);
  }
});

app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// endpoint to submit user data
app.post('/submit', (req, res) => {
  const { name, email, message } = req.body;
  if (!name || !email) return res.status(400).json({ success: false, error: 'Name and email required' });

  const sql = 'INSERT INTO users (name, email, message) VALUES (?, ?, ?)';
  db.run(sql, [name.trim(), email.trim(), message || null], function(err) {
    if (err) {
      console.error('DB insert error', err);
      return res.status(500).json({ success: false, error: 'DB error' });
    }
    res.json({ success: true, id: this.lastID });
  });
});

// list users (for demo)
app.get('/users', (req, res) => {
  db.all('SELECT id, name, email, message, created_at FROM users ORDER BY created_at DESC LIMIT 100', [], (err, rows) => {
    if (err) return res.status(500).json({ success: false, error: 'DB error' });
    res.json({ success: true, users: rows });
  });
});

app.listen(PORT, () => {
  console.log(`Server listening on http://localhost:${PORT}`);
});
