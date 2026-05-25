const express = require('express');
const cors = require('cors');
const Database = require('better-sqlite3');

const app = express();
const port = 3000;

// Middleware
app.use(cors());
app.use(express.json({ limit: '10mb' }));

// Initialize SQLite Database
const db = new Database('database.sqlite');

// Create tables
db.exec(`
  CREATE TABLE IF NOT EXISTS scholarships (
    id TEXT PRIMARY KEY,
    name TEXT,
    country TEXT,
    provider TEXT,
    amount TEXT,
    deadline TEXT,
    status TEXT,
    requirements TEXT,
    requiredDocuments TEXT,
    link TEXT,
    notes TEXT
  );

  CREATE TABLE IF NOT EXISTS applications (
    id TEXT PRIMARY KEY,
    university TEXT,
    country TEXT,
    program TEXT,
    degree TEXT,
    deadline TEXT,
    status TEXT,
    notes TEXT,
    link TEXT,
    docs TEXT
  );

  CREATE TABLE IF NOT EXISTS emails (
    id TEXT PRIMARY KEY,
    professorName TEXT,
    university TEXT,
    email TEXT,
    department TEXT,
    researchArea TEXT,
    mailedDate TEXT,
    responseDate TEXT,
    followUpDate TEXT,
    followUpResponseDate TEXT,
    verdict TEXT,
    remarks TEXT
  );
`);

// --- Helper Functions ---
const createCrudEndpoints = (tableName) => {
  // GET all
  app.get(`/api/${tableName}`, (req, res) => {
    try {
      const rows = db.prepare(`SELECT * FROM ${tableName}`).all();
      // Parse JSON docs for applications
      if (tableName === 'applications') {
        rows.forEach(row => {
          row.docs = row.docs ? JSON.parse(row.docs) : [];
        });
      }
      res.json(rows);
    } catch (err) {
      console.error(err);
      res.status(500).json({ error: 'Internal Server Error' });
    }
  });

  // POST create
  app.post(`/api/${tableName}`, (req, res) => {
    try {
      const data = { ...req.body };
      if (tableName === 'applications' && Array.isArray(data.docs)) {
        data.docs = JSON.stringify(data.docs);
      }
      const keys = Object.keys(data);
      const values = Object.values(data);
      const placeholders = keys.map(() => '?').join(',');
      const stmt = db.prepare(`INSERT INTO ${tableName} (${keys.join(',')}) VALUES (${placeholders})`);
      stmt.run(values);
      res.json({ success: true, id: data.id });
    } catch (err) {
      console.error(err);
      res.status(500).json({ error: 'Internal Server Error' });
    }
  });

  // PUT update
  app.put(`/api/${tableName}/:id`, (req, res) => {
    try {
      const id = req.params.id;
      const data = { ...req.body };
      if (tableName === 'applications' && Array.isArray(data.docs)) {
        data.docs = JSON.stringify(data.docs);
      }
      // Remove id from update if present
      delete data.id;
      
      const keys = Object.keys(data);
      const values = Object.values(data);
      if (keys.length === 0) return res.json({ success: true });
      
      const setClause = keys.map(k => `${k} = ?`).join(', ');
      const stmt = db.prepare(`UPDATE ${tableName} SET ${setClause} WHERE id = ?`);
      const info = stmt.run([...values, id]);
      if (info.changes > 0) {
        res.json({ success: true });
      } else {
        res.status(404).json({ error: 'Not found' });
      }
    } catch (err) {
      console.error(err);
      res.status(500).json({ error: 'Internal Server Error' });
    }
  });

  // DELETE
  app.delete(`/api/${tableName}/:id`, (req, res) => {
    try {
      const stmt = db.prepare(`DELETE FROM ${tableName} WHERE id = ?`);
      const info = stmt.run(req.params.id);
      if (info.changes > 0) {
        res.json({ success: true });
      } else {
        res.status(404).json({ error: 'Not found' });
      }
    } catch (err) {
      console.error(err);
      res.status(500).json({ error: 'Internal Server Error' });
    }
  });
};

createCrudEndpoints('scholarships');
createCrudEndpoints('applications');
createCrudEndpoints('emails');

app.listen(port, () => {
  console.log(`Backend server listening at http://localhost:${port}`);
});
