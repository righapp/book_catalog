// ============================================================
// קטלוג הספרייה - server.js
// ============================================================
'use strict';

require('dotenv').config();
const express  = require('express');
const { google } = require('googleapis');
const path     = require('path');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ---- Config ----
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const BOOKS_SHEET    = 'ספרים';
const LOC_SHEET      = 'מיקומים';

if (!SPREADSHEET_ID || !process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
  console.error('❌  חסרים SPREADSHEET_ID או GOOGLE_SERVICE_ACCOUNT_JSON ב-.env');
  process.exit(1);
}

// ============================================================
// Google Sheets client
// ============================================================

let _sheets = null;

async function getSheets() {
  if (_sheets) return _sheets;
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  _sheets = google.sheets({ version: 'v4', auth });
  return _sheets;
}

// ---- Cache sheet IDs to avoid repeated metadata calls ----
const _sheetIds = {};

async function getSheetId(name) {
  if (_sheetIds[name] !== undefined) return _sheetIds[name];
  const s    = await getSheets();
  const meta = await s.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  meta.data.sheets.forEach(sh => { _sheetIds[sh.properties.title] = sh.properties.sheetId; });
  return _sheetIds[name];
}

// ---- Low-level helpers ----

async function sheetGet(sheetName) {
  const s   = await getSheets();
  const res = await s.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${sheetName}'`,
  });
  return res.data.values || [];
}

async function sheetAppend(sheetName, rows) {
  const s = await getSheets();
  await s.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${sheetName}'`,
    valueInputOption: 'RAW',
    resource: { values: rows },
  });
}

async function sheetUpdate(sheetName, rowNum, values) {
  // rowNum is 1-based (e.g. row 2 = first data row after header)
  const s = await getSheets();
  await s.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${sheetName}'!A${rowNum}`,
    valueInputOption: 'RAW',
    resource: { values: [values] },
  });
}

async function sheetDeleteRow(sheetName, arrayIndex) {
  // arrayIndex is 0-based index from the values array
  const s       = await getSheets();
  const sheetId = await getSheetId(sheetName);
  await s.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    resource: {
      requests: [{
        deleteDimension: {
          range: { sheetId, dimension: 'ROWS', startIndex: arrayIndex, endIndex: arrayIndex + 1 },
        },
      }],
    },
  });
}

// ---- Parse helpers ----

function parseBooks(rows) {
  if (rows.length <= 1) return [];
  return rows.slice(1)
    .map(r => ({
      id:        parseInt(r[0]),
      name:      r[1] || '',
      author:    r[2] || '',
      cabinetId: parseInt(r[3]) || null,
      shelfId:   parseInt(r[4]) || null,
      rowId:     parseInt(r[5]) || null,
    }))
    .filter(b => b.id && b.name);
}

function parseLocations(rows) {
  const out = { cabinets: [], shelves: [], rows: [] };
  if (rows.length <= 1) return out;
  rows.slice(1).forEach(r => {
    const type = r[0], id = parseInt(r[1]), name = r[2], pid = parseInt(r[3]) || null;
    if (!id || !name) return;
    if (type === 'ארון')  out.cabinets.push({ id, name });
    if (type === 'מדף')   out.shelves.push({ id, cabinetId: pid, name });
    if (type === 'שורה')  out.rows.push({ id, shelfId: pid, name });
  });
  return out;
}

function maxId(items) {
  if (!items.length) return 0;
  return Math.max(...items.map(i => i.id));
}

// ---- Sheet initialisation ----

async function ensureSheets() {
  const s    = await getSheets();
  const meta = await s.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const existing = meta.data.sheets.map(sh => sh.properties.title);
  meta.data.sheets.forEach(sh => { _sheetIds[sh.properties.title] = sh.properties.sheetId; });

  const toAdd = [BOOKS_SHEET, LOC_SHEET].filter(n => !existing.includes(n));
  if (toAdd.length) {
    const res = await s.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      resource: { requests: toAdd.map(title => ({ addSheet: { properties: { title } } })) },
    });
    res.data.replies.forEach((r, i) => {
      _sheetIds[toAdd[i]] = r.addSheet.properties.sheetId;
    });
  }

  // Add headers if sheets are empty
  const booksRows = await sheetGet(BOOKS_SHEET);
  if (!booksRows.length) {
    await sheetAppend(BOOKS_SHEET, [['id', 'שם ספר', 'שם סופר', 'ארון_id', 'מדף_id', 'שורה_id']]);
  }
  const locRows = await sheetGet(LOC_SHEET);
  if (!locRows.length) {
    await sheetAppend(LOC_SHEET, [['סוג', 'id', 'שם', 'parent_id']]);
  }
}

// ============================================================
// API Routes
// ============================================================

// GET /api/data  – all books + locations
app.get('/api/data', async (req, res) => {
  try {
    const [booksRows, locRows] = await Promise.all([sheetGet(BOOKS_SHEET), sheetGet(LOC_SHEET)]);
    res.json({ books: parseBooks(booksRows), locations: parseLocations(locRows) });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// POST /api/books  – add one book
app.post('/api/books', async (req, res) => {
  try {
    const { name, author, cabinetId, shelfId, rowId } = req.body;
    const rows   = await sheetGet(BOOKS_SHEET);
    const nextId = maxId(parseBooks(rows)) + 1;
    await sheetAppend(BOOKS_SHEET, [[nextId, name, author, cabinetId ?? '', shelfId ?? '', rowId ?? '']]);
    res.json({ id: nextId, name, author, cabinetId, shelfId, rowId });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// PUT /api/books/:id  – update a book
app.put('/api/books/:id', async (req, res) => {
  try {
    const targetId = parseInt(req.params.id);
    const rows     = await sheetGet(BOOKS_SHEET);
    const idx      = rows.findIndex((r, i) => i > 0 && parseInt(r[0]) === targetId);
    if (idx === -1) return res.status(404).json({ error: 'לא נמצא' });
    const { name, author, cabinetId, shelfId, rowId } = req.body;
    await sheetUpdate(BOOKS_SHEET, idx + 1, [targetId, name, author, cabinetId ?? '', shelfId ?? '', rowId ?? '']);
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// DELETE /api/books/:id
app.delete('/api/books/:id', async (req, res) => {
  try {
    const targetId = parseInt(req.params.id);
    const rows     = await sheetGet(BOOKS_SHEET);
    const idx      = rows.findIndex((r, i) => i > 0 && parseInt(r[0]) === targetId);
    if (idx === -1) return res.status(404).json({ error: 'לא נמצא' });
    await sheetDeleteRow(BOOKS_SHEET, idx);
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// POST /api/books/bulk  – import multiple books + auto-create locations
app.post('/api/books/bulk', async (req, res) => {
  try {
    const incoming = req.body.books || [];
    const [booksRows, locRows] = await Promise.all([sheetGet(BOOKS_SHEET), sheetGet(LOC_SHEET)]);

    let locs     = parseLocations(locRows);
    let nextLocId  = maxId([...locs.cabinets, ...locs.shelves, ...locs.rows]) + 1;
    let nextBookId = maxId(parseBooks(booksRows)) + 1;

    const newBookRows = [];
    const newLocRows  = [];
    const newBooks    = [];

    for (const b of incoming) {
      let cabinet = locs.cabinets.find(c => c.name === b.cabinet) || null;
      if (b.cabinet && !cabinet) {
        cabinet = { id: nextLocId++, name: b.cabinet };
        locs.cabinets.push(cabinet);
        newLocRows.push(['ארון', cabinet.id, cabinet.name, '']);
      }

      let shelf = null;
      if (b.shelf && cabinet) {
        shelf = locs.shelves.find(s => s.name === b.shelf && s.cabinetId === cabinet.id) || null;
        if (!shelf) {
          shelf = { id: nextLocId++, cabinetId: cabinet.id, name: b.shelf };
          locs.shelves.push(shelf);
          newLocRows.push(['מדף', shelf.id, shelf.name, cabinet.id]);
        }
      }

      let row = null;
      if (b.row && shelf) {
        row = locs.rows.find(r => r.name === b.row && r.shelfId === shelf.id) || null;
        if (!row) {
          row = { id: nextLocId++, shelfId: shelf.id, name: b.row };
          locs.rows.push(row);
          newLocRows.push(['שורה', row.id, row.name, shelf.id]);
        }
      }

      const book = { id: nextBookId++, name: b.name, author: b.author,
        cabinetId: cabinet?.id ?? null, shelfId: shelf?.id ?? null, rowId: row?.id ?? null };
      newBooks.push(book);
      newBookRows.push([book.id, book.name, book.author, book.cabinetId ?? '', book.shelfId ?? '', book.rowId ?? '']);
    }

    if (newLocRows.length)  await sheetAppend(LOC_SHEET,   newLocRows);
    if (newBookRows.length) await sheetAppend(BOOKS_SHEET, newBookRows);

    res.json({ books: newBooks, locations: locs });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// POST /api/locations  – add one location (cabinet / shelf / row)
app.post('/api/locations', async (req, res) => {
  try {
    const { type, name, parentId } = req.body; // type: 'ארון'|'מדף'|'שורה'
    const rows   = await sheetGet(LOC_SHEET);
    const locs   = parseLocations(rows);
    const nextId = maxId([...locs.cabinets, ...locs.shelves, ...locs.rows]) + 1;
    await sheetAppend(LOC_SHEET, [[type, nextId, name, parentId ?? '']]);
    res.json({ id: nextId, name, parentId });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// DELETE /api/locations/:id
app.delete('/api/locations/:id', async (req, res) => {
  try {
    const targetId = parseInt(req.params.id);
    const rows     = await sheetGet(LOC_SHEET);
    const idx      = rows.findIndex((r, i) => i > 0 && parseInt(r[1]) === targetId);
    if (idx === -1) return res.status(404).json({ error: 'לא נמצא' });
    await sheetDeleteRow(LOC_SHEET, idx);
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ---- Fallback: serve index.html for any non-API route ----
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

// ============================================================
// Start
// ============================================================
const PORT = process.env.PORT || 3000;
app.listen(PORT, async () => {
  console.log(`\n📚  קטלוג הספרייה פועל על  http://localhost:${PORT}\n`);
  try {
    await ensureSheets();
    console.log('✅  Google Sheets מוכן\n');
  } catch (err) {
    console.error('❌  שגיאה בחיבור ל-Google Sheets:', err.message, '\n');
  }
});
