// ============================================================
// קטלוג הספרייה - app.js
// ============================================================

// ---- STATE ----
const state = {
  view: 'grid',        // 'grid' | 'list'
  search: '',
  sort: 'name-asc',   // 'name-asc' | 'name-desc' | 'author-asc' | 'author-desc' | 'location'
  filter: { cabinetId: null, shelfId: null, rowId: null },
  mobileTab: 'catalog', // 'catalog' | 'manage'
  editingBookId: null,
  deletingBookId: null,
};

const SORT_LABELS = {
  'name-asc':   'שם ספר (א→ת)',
  'name-desc':  'שם ספר (ת→א)',
  'author-asc': 'שם סופר (א→ת)',
  'author-desc':'שם סופר (ת→א)',
  'location':   'מיקום',
};

// ============================================================
// API LAYER
// ============================================================

let db = { books: [], locations: { cabinets: [], shelves: [], rows: [] } };

async function apiFetch(method, path, body) {
  const opts = { method, headers: { 'Content-Type': 'application/json' } };
  if (body !== undefined) opts.body = JSON.stringify(body);
  const res = await fetch(path, opts);
  if (!res.ok) {
    const err = await res.json().catch(() => ({ error: res.statusText }));
    throw new Error(err.error || res.statusText);
  }
  return res.json();
}

async function loadData() {
  const data = await apiFetch('GET', '/api/data');
  db.books     = data.books;
  db.locations = data.locations;
}

function showLoadingOverlay(show) {
  document.getElementById('loadingOverlay').classList.toggle('visible', show);
}

// ---- Helpers ----
function getCabinet(id)  { return db.locations.cabinets.find(c => c.id === id); }
function getShelf(id)    { return db.locations.shelves.find(s => s.id === id); }
function getRow(id)      { return db.locations.rows.find(r => r.id === id); }

function getLocationLabel(book) {
  const parts = [];
  if (book.cabinetId) { const c = getCabinet(book.cabinetId); if (c) parts.push(c.name); }
  if (book.shelfId)   { const s = getShelf(book.shelfId);     if (s) parts.push(s.name); }
  if (book.rowId)     { const r = getRow(book.rowId);          if (r) parts.push(r.name); }
  return parts;
}

// ---- Sort ----
function sortBooks(books) {
  const collator = new Intl.Collator('he', { sensitivity: 'base' });
  return [...books].sort((a, b) => {
    switch (state.sort) {
      case 'name-asc':   return collator.compare(a.name,   b.name);
      case 'name-desc':  return collator.compare(b.name,   a.name);
      case 'author-asc': return collator.compare(a.author, b.author);
      case 'author-desc':return collator.compare(b.author, a.author);
      case 'location': {
        const ca = getCabinet(a.cabinetId), cb = getCabinet(b.cabinetId);
        const sa = getShelf(a.shelfId),     sb = getShelf(b.shelfId);
        const ra = getRow(a.rowId),         rb = getRow(b.rowId);
        return collator.compare(ca?.name ?? '', cb?.name ?? '') ||
               collator.compare(sa?.name ?? '', sb?.name ?? '') ||
               collator.compare(ra?.name ?? '', rb?.name ?? '') ||
               collator.compare(a.name,         b.name);
      }
      default: return 0;
    }
  });
}

// ---- Filter & Search ----
function getFilteredBooks() {
  return db.books.filter(book => {
    if (state.filter.rowId     && book.rowId     !== state.filter.rowId)     return false;
    if (state.filter.shelfId   && book.shelfId   !== state.filter.shelfId)   return false;
    if (state.filter.cabinetId && book.cabinetId !== state.filter.cabinetId) return false;
    if (state.search) {
      const q = state.search.toLowerCase();
      const locationParts = getLocationLabel(book).join(' ').toLowerCase();
      if (!book.name.toLowerCase().includes(q) &&
          !book.author.toLowerCase().includes(q) &&
          !locationParts.includes(q)) return false;
    }
    return true;
  });
}

// ---- Count books per location ----
function countBooksFor(type, id) {
  return db.books.filter(b => b[type + 'Id'] === id).length;
}

// ============================================================
// RENDERING
// ============================================================

function render() {
  renderStats();
  renderLocationTree();
  renderBooks();
}

// ---- Stats ----
function renderStats() {
  const filtered = getFilteredBooks();
  document.getElementById('statTotalBooks').textContent    = db.books.length;
  document.getElementById('statTotalCabinets').textContent = db.locations.cabinets.length;
  document.getElementById('statFilteredBooks').textContent = filtered.length;
  document.getElementById('mobileCount').textContent       = `${filtered.length} ספרים`;

  // Filter badge
  const hasFilter = state.filter.cabinetId || state.filter.shelfId || state.filter.rowId || state.search;
  const badge = document.getElementById('filterBadge');
  if (hasFilter) { badge.textContent = ''; badge.classList.add('visible'); }
  else           { badge.classList.remove('visible'); }
}

// ---- Mobile Tab Switching ----
function switchMobileTab(tab) {
  state.mobileTab = tab;
  const isCatalog = tab === 'catalog';

  document.getElementById('catalogView').style.display    = isCatalog ? '' : 'none';
  document.getElementById('managementView').style.display = isCatalog ? 'none' : 'block';
  document.getElementById('statsBar').style.display       = isCatalog ? '' : 'none';

  document.querySelectorAll('.bottom-tab').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.tab === tab);
  });

  document.body.classList.toggle('tab-manage', !isCatalog);

  // Scroll to top when switching
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ---- Sidebar (mobile) ----
function openSidebar() {
  document.getElementById('sidebar').classList.add('open');
  document.getElementById('sidebarBackdrop').classList.add('open');
  document.body.style.overflow = 'hidden';
}

function closeSidebar() {
  document.getElementById('sidebar').classList.remove('open');
  document.getElementById('sidebarBackdrop').classList.remove('open');
  document.body.style.overflow = '';
}

// ---- Location Tree (Sidebar) ----
function renderLocationTree() {
  const tree = document.getElementById('locationTree');
  const f = state.filter;

  let html = `<div class="tree-all ${!f.cabinetId && !f.shelfId && !f.rowId ? 'active' : ''}" data-action="filter-all">
    📚 כל הספרים
    <span class="tree-count">${db.books.length}</span>
  </div>`;

  for (const cab of db.locations.cabinets) {
    const cabShelves = db.locations.shelves.filter(s => s.cabinetId === cab.id);
    const cabBooks   = db.books.filter(b => b.cabinetId === cab.id).length;
    const cabActive  = f.cabinetId === cab.id && !f.shelfId;
    const cabOpen    = f.cabinetId === cab.id;

    html += `<div class="tree-cabinet">
      <div class="tree-cabinet-header ${cabActive ? 'active' : ''}" data-action="filter-cabinet" data-id="${cab.id}">
        🗄️ ${cab.name}
        <span class="tree-count">${cabBooks}</span>
        <span class="tree-toggle ${cabOpen ? 'open' : ''}">▶</span>
      </div>
      <div class="tree-shelves ${cabOpen ? 'open' : ''}">`;

    for (const shelf of cabShelves) {
      const shelfRows  = db.locations.rows.filter(r => r.shelfId === shelf.id);
      const shelfBooks = db.books.filter(b => b.shelfId === shelf.id).length;
      const shelfActive = f.shelfId === shelf.id && !f.rowId;
      const shelfOpen   = f.shelfId === shelf.id;

      html += `<div class="tree-shelf">
        <div class="tree-shelf-header ${shelfActive ? 'active' : ''}" data-action="filter-shelf" data-id="${shelf.id}" data-cabinet="${cab.id}">
          📋 ${shelf.name}
          <span class="tree-count">${shelfBooks}</span>
          <span class="tree-toggle ${shelfOpen ? 'open' : ''}">▶</span>
        </div>
        <div class="tree-rows ${shelfOpen ? 'open' : ''}">`;

      for (const row of shelfRows) {
        const rowBooks = db.books.filter(b => b.rowId === row.id).length;
        const rowActive = f.rowId === row.id;
        html += `<div class="tree-row-item ${rowActive ? 'active' : ''}" data-action="filter-row" data-id="${row.id}" data-shelf="${shelf.id}" data-cabinet="${cab.id}">
          • ${row.name} <span class="tree-count">${rowBooks}</span>
        </div>`;
      }

      html += `</div></div>`;
    }

    html += `</div></div>`;
  }

  tree.innerHTML = html;
}

// ---- Books ----
function renderBooks() {
  const books = sortBooks(getFilteredBooks());
  const container = document.getElementById('booksContainer');
  const empty = document.getElementById('emptyState');

  if (books.length === 0) {
    container.innerHTML = '';
    empty.classList.remove('hidden');
    return;
  }

  empty.classList.add('hidden');
  container.className = state.view === 'grid' ? 'books-grid' : 'books-list';

  if (state.view === 'grid') {
    container.innerHTML = books.map(renderBookCard).join('');
  } else {
    container.innerHTML = books.map(renderBookRow).join('');
  }
}

function renderBookCard(book) {
  const loc = getLocationLabel(book);
  const badges = loc.map(l => `<span class="location-badge">${l}</span>`).join('');
  return `<div class="book-card">
    <div class="book-card-top">
      <span class="book-card-title">${esc(book.name)}</span>
      <span class="book-card-author">${esc(book.author)}</span>
    </div>
    <div class="book-card-location">${badges || '<span class="location-badge" style="opacity:.5">ללא מיקום</span>'}</div>
    <div class="book-card-actions">
      <button class="btn-card-edit" data-action="edit" data-id="${book.id}">✏️ עריכה</button>
      <button class="btn-card-delete" data-action="delete" data-id="${book.id}">🗑️ מחק</button>
    </div>
  </div>`;
}

function renderBookRow(book) {
  const loc = getLocationLabel(book);
  const badges = loc.map(l => `<span class="location-badge">${l}</span>`).join('');
  return `<div class="book-row">
    <div class="book-row-main">
      <div class="book-card-top">
        <span class="book-card-title">${esc(book.name)}</span>
        <span class="book-card-author">${esc(book.author)}</span>
      </div>
      <div class="book-row-location">${badges || '<span class="location-badge" style="opacity:.5">ללא מיקום</span>'}</div>
    </div>
    <div class="book-row-actions">
      <button class="btn-card-edit" data-action="edit" data-id="${book.id}">✏️ עריכה</button>
      <button class="btn-card-delete" data-action="delete" data-id="${book.id}">🗑️</button>
    </div>
  </div>`;
}

function esc(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ============================================================
// MODALS
// ============================================================

function openModal(id) {
  document.getElementById(id).classList.add('open');
}

function closeModal(id) {
  document.getElementById(id).classList.remove('open');
}

// ---- Book Modal ----
function switchBookModalTab(tab) {
  document.querySelectorAll('.modal-tab').forEach(t =>
    t.classList.toggle('active', t.dataset.tab === tab));
  document.getElementById('tabManual').classList.toggle('active', tab === 'manual');
  document.getElementById('tabExcel').classList.toggle('active',  tab === 'excel');

  const saveBtn   = document.getElementById('bookModalSave');
  const importBtn = document.getElementById('importModalConfirm');
  if (tab === 'manual') {
    saveBtn.style.display   = '';
    importBtn.style.display = 'none';
  } else {
    saveBtn.style.display   = 'none';
    // Import button is shown only after a file is loaded — handled in showImportResult
    importBtn.style.display = 'none';
  }
}

function openAddBookModal() {
  state.editingBookId = null;
  document.getElementById('bookModalTitle').textContent = 'הוסף ספר חדש';
  document.getElementById('bookModalSave').textContent  = '💾 שמור ספר';
  document.getElementById('bookModalTabs').style.display = '';
  resetBookForm();
  resetExcelTab();
  switchBookModalTab('manual');
  openModal('bookModal');
  document.getElementById('bookName').focus();
}

function openEditBookModal(id) {
  const book = db.books.find(b => b.id === id);
  if (!book) return;
  state.editingBookId = id;
  document.getElementById('bookModalTitle').textContent  = 'עריכת ספר';
  document.getElementById('bookModalSave').textContent   = '💾 שמור שינויים';
  document.getElementById('bookModalTabs').style.display = 'none';
  resetBookForm();

  document.getElementById('bookId').value     = book.id;
  document.getElementById('bookName').value   = book.name;
  document.getElementById('bookAuthor').value = book.author;

  populateCabinetSelect(book.cabinetId);
  if (book.cabinetId) {
    populateShelfSelect(book.cabinetId, book.shelfId);
    if (book.shelfId) populateRowSelect(book.shelfId, book.rowId);
  }

  // Force manual tab visible, import button hidden
  document.getElementById('tabManual').classList.add('active');
  document.getElementById('tabExcel').classList.remove('active');
  document.getElementById('bookModalSave').style.display   = '';
  document.getElementById('importModalConfirm').style.display = 'none';

  openModal('bookModal');
  document.getElementById('bookName').focus();
}

function resetBookForm() {
  document.getElementById('bookId').value     = '';
  document.getElementById('bookName').value   = '';
  document.getElementById('bookAuthor').value = '';
  document.getElementById('bookNameError').textContent   = '';
  document.getElementById('bookAuthorError').textContent = '';

  populateCabinetSelect(null);
  populateShelfSelect(null, null);
  populateRowSelect(null, null);

  hideNewRow('newCabinetRow');
  hideNewRow('newShelfRow');
  hideNewRow('newRowRow');
}

function resetExcelTab() {
  pendingImportBooks = [];
  document.getElementById('dropZone').style.display     = '';
  document.getElementById('importResult').classList.add('hidden');
  document.getElementById('importModalConfirm').style.display = 'none';
  document.getElementById('excelFileInput').value = '';
}

// ---- Cascading Dropdowns ----
function populateCabinetSelect(selectedId) {
  const sel = document.getElementById('cabinetSelect');
  sel.innerHTML = '<option value="">-- בחר ארון --</option>';
  db.locations.cabinets.forEach(c => {
    const opt = new Option(c.name, c.id, false, c.id === selectedId);
    sel.appendChild(opt);
  });
  sel.appendChild(new Option('＋ הוסף ארון חדש...', 'NEW'));
}

function populateShelfSelect(cabinetId, selectedId) {
  const sel = document.getElementById('shelfSelect');
  sel.innerHTML = '<option value="">-- בחר מדף --</option>';
  sel.disabled  = !cabinetId;

  if (cabinetId) {
    const shelves = db.locations.shelves.filter(s => s.cabinetId === cabinetId);
    shelves.forEach(s => {
      const opt = new Option(s.name, s.id, false, s.id === selectedId);
      sel.appendChild(opt);
    });
    sel.appendChild(new Option('＋ הוסף מדף חדש...', 'NEW'));
  }
}

function populateRowSelect(shelfId, selectedId) {
  const sel = document.getElementById('rowSelect');
  sel.innerHTML = '<option value="">-- בחר שורה --</option>';
  sel.disabled  = !shelfId;

  if (shelfId) {
    const rows = db.locations.rows.filter(r => r.shelfId === shelfId);
    rows.forEach(r => {
      const opt = new Option(r.name, r.id, false, r.id === selectedId);
      sel.appendChild(opt);
    });
    sel.appendChild(new Option('＋ הוסף שורה חדשה...', 'NEW'));
  }
}

function showNewRow(rowId) {
  const row = document.getElementById(rowId);
  row.classList.add('visible');
  row.querySelector('input').value = '';
  row.querySelector('input').focus();
}

function hideNewRow(rowId) {
  document.getElementById(rowId).classList.remove('visible');
}

// ---- Save Book ----
async function saveBook() {
  const name   = document.getElementById('bookName').value.trim();
  const author = document.getElementById('bookAuthor').value.trim();
  let valid = true;

  document.getElementById('bookNameError').textContent   = '';
  document.getElementById('bookAuthorError').textContent = '';

  if (!name)   { document.getElementById('bookNameError').textContent   = 'שדה חובה'; valid = false; }
  if (!author) { document.getElementById('bookAuthorError').textContent = 'שדה חובה'; valid = false; }
  if (!valid)  return;

  const cabinetVal = document.getElementById('cabinetSelect').value;
  const shelfVal   = document.getElementById('shelfSelect').value;
  const rowVal     = document.getElementById('rowSelect').value;

  const bookData = {
    name,
    author,
    cabinetId: cabinetVal && cabinetVal !== 'NEW' ? parseInt(cabinetVal) : null,
    shelfId:   shelfVal   && shelfVal   !== 'NEW' ? parseInt(shelfVal)   : null,
    rowId:     rowVal     && rowVal     !== 'NEW' ? parseInt(rowVal)     : null,
  };

  showLoadingOverlay(true);
  try {
    if (state.editingBookId) {
      await apiFetch('PUT', `/api/books/${state.editingBookId}`, bookData);
      const book = db.books.find(b => b.id === state.editingBookId);
      Object.assign(book, bookData);
      showToast('הספר עודכן בהצלחה ✓', 'success');
    } else {
      const result = await apiFetch('POST', '/api/books', bookData);
      db.books.push(result);
      showToast('הספר נוסף בהצלחה ✓', 'success');
    }
    closeModal('bookModal');
    render();
  } catch (e) {
    showToast('שגיאה: ' + e.message, 'error');
  } finally {
    showLoadingOverlay(false);
  }
}

// ---- Delete Modal ----
function openDeleteModal(id) {
  const book = db.books.find(b => b.id === id);
  if (!book) return;
  state.deletingBookId = id;
  document.getElementById('deleteBookName').textContent = `"${book.name}"`;
  openModal('deleteModal');
}

async function confirmDelete() {
  if (!state.deletingBookId) return;
  showLoadingOverlay(true);
  try {
    await apiFetch('DELETE', `/api/books/${state.deletingBookId}`);
    db.books = db.books.filter(b => b.id !== state.deletingBookId);
    state.deletingBookId = null;
    closeModal('deleteModal');
    render();
    showToast('הספר נמחק', 'success');
  } catch (e) {
    showToast('שגיאה: ' + e.message, 'error');
  } finally {
    showLoadingOverlay(false);
  }
}

// ---- Locations Modal ----
function openLocationsModal() {
  renderLocationsManager();
  openModal('locationsModal');
}

function renderLocationsManager() {
  // Populate cabinet select in shelves tab
  const shelfCabSel = document.getElementById('newShelfCabinet');
  shelfCabSel.innerHTML = '<option value="">-- בחר ארון --</option>';
  db.locations.cabinets.forEach(c => shelfCabSel.appendChild(new Option(c.name, c.id)));

  // Populate shelf select in rows tab (all shelves with cabinet name)
  const rowShelfSel = document.getElementById('newRowShelf');
  rowShelfSel.innerHTML = '<option value="">-- בחר מדף --</option>';
  db.locations.shelves.forEach(s => {
    const cab = getCabinet(s.cabinetId);
    rowShelfSel.appendChild(new Option(`${cab ? cab.name + ' / ' : ''}${s.name}`, s.id));
  });

  // Cabinets list
  const cabList = document.getElementById('cabinetsList');
  if (db.locations.cabinets.length === 0) {
    cabList.innerHTML = '<div style="color:var(--color-muted);padding:10px">אין ארונות</div>';
  } else {
    cabList.innerHTML = db.locations.cabinets.map(c => {
      const booksCount = db.books.filter(b => b.cabinetId === c.id).length;
      return `<div class="loc-item">
        <div><div class="loc-item-name">🗄️ ${esc(c.name)}</div>
        <div class="loc-item-meta">${booksCount} ספרים</div></div>
        <button class="loc-item-delete" data-action="del-cabinet" data-id="${c.id}">מחק</button>
      </div>`;
    }).join('');
  }

  // Shelves list
  const shelvesList = document.getElementById('shelvesList');
  if (db.locations.shelves.length === 0) {
    shelvesList.innerHTML = '<div style="color:var(--color-muted);padding:10px">אין מדפים</div>';
  } else {
    shelvesList.innerHTML = db.locations.shelves.map(s => {
      const cab = getCabinet(s.cabinetId);
      const booksCount = db.books.filter(b => b.shelfId === s.id).length;
      return `<div class="loc-item">
        <div><div class="loc-item-name">📋 ${esc(s.name)}</div>
        <div class="loc-item-meta">${cab ? cab.name : ''} · ${booksCount} ספרים</div></div>
        <button class="loc-item-delete" data-action="del-shelf" data-id="${s.id}">מחק</button>
      </div>`;
    }).join('');
  }

  // Rows list
  const rowsList = document.getElementById('rowsList');
  if (db.locations.rows.length === 0) {
    rowsList.innerHTML = '<div style="color:var(--color-muted);padding:10px">אין שורות</div>';
  } else {
    rowsList.innerHTML = db.locations.rows.map(r => {
      const shelf = getShelf(r.shelfId);
      const cab   = shelf ? getCabinet(shelf.cabinetId) : null;
      const booksCount = db.books.filter(b => b.rowId === r.id).length;
      return `<div class="loc-item">
        <div><div class="loc-item-name">• ${esc(r.name)}</div>
        <div class="loc-item-meta">${cab ? cab.name + ' / ' : ''}${shelf ? shelf.name : ''} · ${booksCount} ספרים</div></div>
        <button class="loc-item-delete" data-action="del-row" data-id="${r.id}">מחק</button>
      </div>`;
    }).join('');
  }
}

// ============================================================
// TOAST
// ============================================================
let toastTimer;
function showToast(msg, type = '') {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.className   = 'toast show' + (type ? ' ' + type : '');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { el.className = 'toast'; }, 3000);
}

// ============================================================
// EXCEL IMPORT / EXPORT
// ============================================================

let pendingImportBooks = [];

function downloadTemplate() {
  const wb = XLSX.utils.book_new();
  const wsData = [
    ['שם ספר', 'שם סופר', 'ארון', 'מדף', 'שורה'],
    ['הארי פוטר ואבן החכמים', "ג'יי קיי רולינג",        'ארון 1', 'מדף 1', 'שורה 1'],
    ['1984',                  "ג'ורג' אורוול",            'ארון 1', 'מדף 1', 'שורה 2'],
    ['הנסיך הקטן',            'אנטואן דה סנט-אקזופרי',   'ארון 2', 'מדף 3', 'שורה 1'],
    ['ספר לדוגמה',            'סופר לדוגמה',              'ארון 2', 'מדף 3', 'שורה 2'],
  ];

  const ws = XLSX.utils.aoa_to_sheet(wsData);
  ws['!cols'] = [{ wch: 32 }, { wch: 26 }, { wch: 12 }, { wch: 12 }, { wch: 12 }];

  // Bold header
  ['A1','B1','C1','D1','E1'].forEach(cell => {
    if (!ws[cell]) return;
    ws[cell].s = {
      font: { bold: true, color: { rgb: 'FFFFFF' } },
      fill: { fgColor: { rgb: '6B3F26' } },
      alignment: { horizontal: 'center', readingOrder: 2 },
    };
  });

  XLSX.utils.book_append_sheet(wb, ws, 'ספרים');
  XLSX.writeFile(wb, 'תבנית_קטלוג_ספרים.xlsx');
  showToast('התבנית הורדה בהצלחה ✓', 'success');
}

function handleImportFile(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows  = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      parseImportRows(rows);
    } catch {
      showToast('שגיאה בקריאת הקובץ', 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

function parseImportRows(rows) {
  if (rows.length < 2) { showToast('הקובץ ריק', 'error'); return; }

  const HEADER_MAP = {
    'שם ספר': 'name', 'שם סופר': 'author',
    'ארון': 'cabinet', 'מדף': 'shelf', 'שורה': 'row',
  };

  const headers = rows[0].map(h => String(h).trim());
  const colIdx  = {};
  headers.forEach((h, i) => { if (HEADER_MAP[h]) colIdx[HEADER_MAP[h]] = i; });

  if (colIdx.name === undefined || colIdx.author === undefined) {
    showToast('חסרות עמודות חובה: "שם ספר" ו-"שם סופר"', 'error');
    return;
  }

  const books  = [];
  const errors = [];

  rows.slice(1).forEach((row, i) => {
    if (row.every(c => !String(c).trim())) return; // skip empty rows
    const line   = i + 2;
    const name   = String(row[colIdx.name]   ?? '').trim();
    const author = String(row[colIdx.author] ?? '').trim();
    if (!name)   { errors.push(`שורה ${line}: חסר שם ספר`);  return; }
    if (!author) { errors.push(`שורה ${line}: חסר שם סופר`); return; }
    books.push({
      name, author,
      cabinet: colIdx.cabinet !== undefined ? String(row[colIdx.cabinet] ?? '').trim() : '',
      shelf:   colIdx.shelf   !== undefined ? String(row[colIdx.shelf]   ?? '').trim() : '',
      row:     colIdx.row     !== undefined ? String(row[colIdx.row]     ?? '').trim() : '',
    });
  });

  if (books.length === 0) { showToast('לא נמצאו ספרים תקינים לייבוא', 'error'); return; }

  pendingImportBooks = books;
  showImportResult(books, errors);
}

function showImportResult(books, errors) {
  // Detect new locations
  const newCabNames  = new Set();
  const newShelfKeys = new Set();
  const newRowKeys   = new Set();

  books.forEach(b => {
    if (b.cabinet) {
      const existCab = db.locations.cabinets.find(c => c.name === b.cabinet);
      if (!existCab) newCabNames.add(b.cabinet);
      if (b.shelf) {
        const cab = existCab || { id: -1 };
        const existShelf = db.locations.shelves.find(
          s => s.name === b.shelf && (s.cabinetId === cab.id || newCabNames.has(b.cabinet))
        );
        if (!existShelf) newShelfKeys.add(`${b.cabinet}/${b.shelf}`);
        if (b.row) {
          const existRow = db.locations.rows.find(r => r.name === b.row);
          if (!existRow) newRowKeys.add(`${b.cabinet}/${b.shelf}/${b.row}`);
        }
      }
    }
  });

  // Summary
  const parts = [`<strong>${books.length}</strong> ספרים`];
  if (newCabNames.size)  parts.push(`<span class="import-new-badge">חדש</span>${newCabNames.size} ארונות`);
  if (newShelfKeys.size) parts.push(`<span class="import-new-badge">חדש</span>${newShelfKeys.size} מדפים`);
  if (newRowKeys.size)   parts.push(`<span class="import-new-badge">חדש</span>${newRowKeys.size} שורות`);

  document.getElementById('importSummary').innerHTML =
    `<div class="import-summary-box">📊 נמצאו: ${parts.join(' &nbsp;·&nbsp; ')}</div>`;

  // Errors
  const errBox = document.getElementById('importErrorsBox');
  if (errors.length) {
    errBox.classList.remove('hidden');
    errBox.innerHTML = `<strong>⚠️ ${errors.length} שורות עם שגיאות (ידולגו):</strong>
      <ul>${errors.map(e => `<li>${e}</li>`).join('')}</ul>`;
  } else {
    errBox.classList.add('hidden');
  }

  // Preview table
  const preview  = books.slice(0, 15);
  const moreRows = books.length > 15
    ? `<tr><td colspan="6" style="text-align:center;color:var(--color-muted);padding:10px">...ועוד ${books.length - 15} ספרים</td></tr>`
    : '';

  document.getElementById('importPreviewTable').innerHTML = `
    <table class="import-table">
      <thead><tr><th>#</th><th>שם ספר</th><th>שם סופר</th><th>ארון</th><th>מדף</th><th>שורה</th></tr></thead>
      <tbody>
        ${preview.map((b, i) => `<tr>
          <td style="color:var(--color-muted)">${i + 1}</td>
          <td><strong>${esc(b.name)}</strong></td>
          <td>${esc(b.author)}</td>
          <td>${esc(b.cabinet)}</td>
          <td>${esc(b.shelf)}</td>
          <td>${esc(b.row)}</td>
        </tr>`).join('')}
        ${moreRows}
      </tbody>
    </table>`;

  // Show result panel, hide drop zone, show import button in footer
  document.getElementById('dropZone').style.display = 'none';
  document.getElementById('importResult').classList.remove('hidden');
  const importBtn = document.getElementById('importModalConfirm');
  importBtn.textContent   = `📥 ייבא ${books.length} ספרים`;
  importBtn.style.display = '';
}

async function confirmImport() {
  const count = pendingImportBooks.length;
  showLoadingOverlay(true);
  try {
    const result = await apiFetch('POST', '/api/books/bulk', { books: pendingImportBooks });
    db.books.push(...result.books);
    db.locations = result.locations;
    pendingImportBooks = [];
    closeModal('bookModal');
    render();
    showToast(`${count} ספרים יובאו בהצלחה ✓`, 'success');
  } catch (e) {
    showToast('שגיאה: ' + e.message, 'error');
  } finally {
    showLoadingOverlay(false);
  }
}

// ============================================================
// EVENT LISTENERS
// ============================================================

document.addEventListener('DOMContentLoaded', () => {

  // ---- Navbar ----
  document.getElementById('addBookBtn').addEventListener('click', openAddBookModal);
  document.getElementById('emptyAddBtn').addEventListener('click', openAddBookModal);

  // Search
  const searchInput = document.getElementById('searchInput');
  const searchClear = document.getElementById('searchClear');

  searchInput.addEventListener('input', () => {
    state.search = searchInput.value;
    searchClear.classList.toggle('visible', state.search.length > 0);
    render();
  });

  searchClear.addEventListener('click', () => {
    searchInput.value = '';
    state.search = '';
    searchClear.classList.remove('visible');
    render();
  });

  // View toggle
  document.getElementById('viewGridBtn').addEventListener('click', () => {
    state.view = 'grid';
    document.getElementById('viewGridBtn').classList.add('active');
    document.getElementById('viewListBtn').classList.remove('active');
    render();
  });

  document.getElementById('viewListBtn').addEventListener('click', () => {
    state.view = 'list';
    document.getElementById('viewListBtn').classList.add('active');
    document.getElementById('viewGridBtn').classList.remove('active');
    render();
  });

  // Manage locations
  document.getElementById('manageLocationsBtn').addEventListener('click', openLocationsModal);

  // Clear filter
  document.getElementById('clearFilterBtn').addEventListener('click', () => {
    state.filter = { cabinetId: null, shelfId: null, rowId: null };
    render();
  });

  // ---- Location Tree (delegated) ----
  document.getElementById('locationTree').addEventListener('click', e => {
    const el = e.target.closest('[data-action]');
    if (!el) return;
    const action = el.dataset.action;

    if (action === 'filter-all') {
      state.filter = { cabinetId: null, shelfId: null, rowId: null };
    } else if (action === 'filter-cabinet') {
      const id = parseInt(el.dataset.id);
      if (state.filter.cabinetId === id && !state.filter.shelfId) {
        state.filter = { cabinetId: null, shelfId: null, rowId: null };
      } else {
        state.filter = { cabinetId: id, shelfId: null, rowId: null };
      }
    } else if (action === 'filter-shelf') {
      const shelfId   = parseInt(el.dataset.id);
      const cabinetId = parseInt(el.dataset.cabinet);
      if (state.filter.shelfId === shelfId && !state.filter.rowId) {
        state.filter = { cabinetId, shelfId: null, rowId: null };
      } else {
        state.filter = { cabinetId, shelfId, rowId: null };
      }
    } else if (action === 'filter-row') {
      const rowId     = parseInt(el.dataset.id);
      const shelfId   = parseInt(el.dataset.shelf);
      const cabinetId = parseInt(el.dataset.cabinet);
      if (state.filter.rowId === rowId) {
        state.filter = { cabinetId, shelfId, rowId: null };
      } else {
        state.filter = { cabinetId, shelfId, rowId };
      }
    }
    render();
  });

  // ---- Books container (delegated) ----
  document.getElementById('booksContainer').addEventListener('click', e => {
    const el = e.target.closest('[data-action]');
    if (!el) return;
    if (el.dataset.action === 'edit')   openEditBookModal(parseInt(el.dataset.id));
    if (el.dataset.action === 'delete') openDeleteModal(parseInt(el.dataset.id));
  });

  // ---- Book Modal ----
  document.getElementById('bookModalClose').addEventListener('click',  () => closeModal('bookModal'));
  document.getElementById('bookModalCancel').addEventListener('click', () => closeModal('bookModal'));
  document.getElementById('bookModalSave').addEventListener('click', saveBook);

  // Cabinet select change
  document.getElementById('cabinetSelect').addEventListener('change', e => {
    const val = e.target.value;
    if (val === 'NEW') {
      showNewRow('newCabinetRow');
      e.target.value = '';
    } else {
      hideNewRow('newCabinetRow');
      populateShelfSelect(val ? parseInt(val) : null, null);
      populateRowSelect(null, null);
    }
  });

  // Confirm / Cancel new cabinet
  document.getElementById('confirmNewCabinet').addEventListener('click', async () => {
    const name = document.getElementById('newCabinetName').value.trim();
    if (!name) { showToast('הכנס שם לארון', 'error'); return; }
    showLoadingOverlay(true);
    try {
      const result = await apiFetch('POST', '/api/locations', { type: 'ארון', name });
      const newCab = { id: result.id, name: result.name };
      db.locations.cabinets.push(newCab);
      hideNewRow('newCabinetRow');
      populateCabinetSelect(newCab.id);
      populateShelfSelect(newCab.id, null);
      populateRowSelect(null, null);
      showToast(`ארון "${name}" נוסף ✓`, 'success');
      render();
    } catch (e) {
      showToast('שגיאה: ' + e.message, 'error');
    } finally {
      showLoadingOverlay(false);
    }
  });

  document.getElementById('cancelNewCabinet').addEventListener('click', () => {
    hideNewRow('newCabinetRow');
    document.getElementById('cabinetSelect').value = '';
  });

  // Shelf select change
  document.getElementById('shelfSelect').addEventListener('change', e => {
    const val = e.target.value;
    const cabinetId = parseInt(document.getElementById('cabinetSelect').value);
    if (val === 'NEW') {
      showNewRow('newShelfRow');
      e.target.value = '';
    } else {
      hideNewRow('newShelfRow');
      populateRowSelect(val ? parseInt(val) : null, null);
    }
  });

  // Confirm / Cancel new shelf
  document.getElementById('confirmNewShelf').addEventListener('click', async () => {
    const name = document.getElementById('newShelfName').value.trim();
    const cabinetId = parseInt(document.getElementById('cabinetSelect').value);
    if (!name) { showToast('הכנס שם למדף', 'error'); return; }
    if (!cabinetId) { showToast('בחר ארון תחילה', 'error'); return; }
    showLoadingOverlay(true);
    try {
      const result = await apiFetch('POST', '/api/locations', { type: 'מדף', name, parentId: cabinetId });
      const newShelf = { id: result.id, cabinetId, name: result.name };
      db.locations.shelves.push(newShelf);
      hideNewRow('newShelfRow');
      populateShelfSelect(cabinetId, newShelf.id);
      populateRowSelect(newShelf.id, null);
      showToast(`מדף "${name}" נוסף ✓`, 'success');
      render();
    } catch (e) {
      showToast('שגיאה: ' + e.message, 'error');
    } finally {
      showLoadingOverlay(false);
    }
  });

  document.getElementById('cancelNewShelf').addEventListener('click', () => {
    hideNewRow('newShelfRow');
    document.getElementById('shelfSelect').value = '';
  });

  // Row select change
  document.getElementById('rowSelect').addEventListener('change', e => {
    const val = e.target.value;
    if (val === 'NEW') {
      showNewRow('newRowRow');
      e.target.value = '';
    } else {
      hideNewRow('newRowRow');
    }
  });

  // Confirm / Cancel new row
  document.getElementById('confirmNewRow').addEventListener('click', async () => {
    const name = document.getElementById('newRowName').value.trim();
    const shelfId = parseInt(document.getElementById('shelfSelect').value);
    if (!name) { showToast('הכנס שם לשורה', 'error'); return; }
    if (!shelfId) { showToast('בחר מדף תחילה', 'error'); return; }
    showLoadingOverlay(true);
    try {
      const result = await apiFetch('POST', '/api/locations', { type: 'שורה', name, parentId: shelfId });
      const newRow = { id: result.id, shelfId, name: result.name };
      db.locations.rows.push(newRow);
      hideNewRow('newRowRow');
      populateRowSelect(shelfId, newRow.id);
      showToast(`שורה "${name}" נוספה ✓`, 'success');
      render();
    } catch (e) {
      showToast('שגיאה: ' + e.message, 'error');
    } finally {
      showLoadingOverlay(false);
    }
  });

  document.getElementById('cancelNewRow').addEventListener('click', () => {
    hideNewRow('newRowRow');
    document.getElementById('rowSelect').value = '';
  });

  // ---- Delete Modal ----
  document.getElementById('deleteModalClose').addEventListener('click',   () => closeModal('deleteModal'));
  document.getElementById('deleteModalCancel').addEventListener('click',  () => closeModal('deleteModal'));
  document.getElementById('deleteModalConfirm').addEventListener('click', confirmDelete);

  // ---- Locations Modal ----
  document.getElementById('locationsModalClose').addEventListener('click',  () => closeModal('locationsModal'));
  document.getElementById('locationsModalClose2').addEventListener('click', () => closeModal('locationsModal'));

  // Tabs
  document.querySelectorAll('.loc-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.loc-tab').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.loc-panel').forEach(p => p.classList.remove('active'));
      tab.classList.add('active');
      document.getElementById('panel' + tab.dataset.tab.charAt(0).toUpperCase() + tab.dataset.tab.slice(1)).classList.add('active');
    });
  });

  // Add cabinet from manager
  document.getElementById('addCabinetBtn').addEventListener('click', async () => {
    const name = document.getElementById('newCabinetNameMgr').value.trim();
    if (!name) { showToast('הכנס שם לארון', 'error'); return; }
    showLoadingOverlay(true);
    try {
      const result = await apiFetch('POST', '/api/locations', { type: 'ארון', name });
      db.locations.cabinets.push({ id: result.id, name: result.name });
      document.getElementById('newCabinetNameMgr').value = '';
      renderLocationsManager();
      render();
      showToast(`ארון "${name}" נוסף ✓`, 'success');
    } catch (e) {
      showToast('שגיאה: ' + e.message, 'error');
    } finally {
      showLoadingOverlay(false);
    }
  });

  // Add shelf from manager
  document.getElementById('addShelfBtn').addEventListener('click', async () => {
    const name      = document.getElementById('newShelfNameMgr').value.trim();
    const cabinetId = parseInt(document.getElementById('newShelfCabinet').value);
    if (!name)      { showToast('הכנס שם למדף', 'error');   return; }
    if (!cabinetId) { showToast('בחר ארון תחילה', 'error'); return; }
    showLoadingOverlay(true);
    try {
      const result = await apiFetch('POST', '/api/locations', { type: 'מדף', name, parentId: cabinetId });
      db.locations.shelves.push({ id: result.id, cabinetId, name: result.name });
      document.getElementById('newShelfNameMgr').value = '';
      renderLocationsManager();
      render();
      showToast(`מדף "${name}" נוסף ✓`, 'success');
    } catch (e) {
      showToast('שגיאה: ' + e.message, 'error');
    } finally {
      showLoadingOverlay(false);
    }
  });

  // Add row from manager
  document.getElementById('addRowBtn').addEventListener('click', async () => {
    const name    = document.getElementById('newRowNameMgr').value.trim();
    const shelfId = parseInt(document.getElementById('newRowShelf').value);
    if (!name)    { showToast('הכנס שם לשורה', 'error');   return; }
    if (!shelfId) { showToast('בחר מדף תחילה', 'error');   return; }
    showLoadingOverlay(true);
    try {
      const result = await apiFetch('POST', '/api/locations', { type: 'שורה', name, parentId: shelfId });
      db.locations.rows.push({ id: result.id, shelfId, name: result.name });
      document.getElementById('newRowNameMgr').value = '';
      renderLocationsManager();
      render();
      showToast(`שורה "${name}" נוספה ✓`, 'success');
    } catch (e) {
      showToast('שגיאה: ' + e.message, 'error');
    } finally {
      showLoadingOverlay(false);
    }
  });

  // Delete location (delegated from manager lists)
  ['cabinetsList', 'shelvesList', 'rowsList'].forEach(listId => {
    document.getElementById(listId).addEventListener('click', async e => {
      const btn = e.target.closest('[data-action]');
      if (!btn) return;
      const id = parseInt(btn.dataset.id);

      if (btn.dataset.action === 'del-cabinet') {
        const usedByShelf = db.locations.shelves.some(s => s.cabinetId === id);
        const usedByBook  = db.books.some(b => b.cabinetId === id);
        if (usedByShelf || usedByBook) { showToast('לא ניתן למחוק - יש מדפים או ספרים בארון זה', 'error'); return; }
      } else if (btn.dataset.action === 'del-shelf') {
        const usedByRow  = db.locations.rows.some(r => r.shelfId === id);
        const usedByBook = db.books.some(b => b.shelfId === id);
        if (usedByRow || usedByBook) { showToast('לא ניתן למחוק - יש שורות או ספרים במדף זה', 'error'); return; }
      } else if (btn.dataset.action === 'del-row') {
        const usedByBook = db.books.some(b => b.rowId === id);
        if (usedByBook) { showToast('לא ניתן למחוק - יש ספרים בשורה זו', 'error'); return; }
      }

      showLoadingOverlay(true);
      try {
        await apiFetch('DELETE', `/api/locations/${id}`);
        if (btn.dataset.action === 'del-cabinet') {
          db.locations.cabinets = db.locations.cabinets.filter(c => c.id !== id);
        } else if (btn.dataset.action === 'del-shelf') {
          db.locations.shelves = db.locations.shelves.filter(s => s.id !== id);
        } else if (btn.dataset.action === 'del-row') {
          db.locations.rows = db.locations.rows.filter(r => r.id !== id);
        }
        renderLocationsManager();
        render();
        showToast('המיקום נמחק ✓', 'success');
      } catch (e) {
        showToast('שגיאה: ' + e.message, 'error');
      } finally {
        showLoadingOverlay(false);
      }
    });
  });

  // Close modal on overlay click
  ['bookModal', 'deleteModal', 'locationsModal'].forEach(id => {
    document.getElementById(id).addEventListener('click', e => {
      if (e.target === document.getElementById(id)) closeModal(id);
    });
  });

  // Enter key in book form fields
  ['bookName', 'bookAuthor'].forEach(fieldId => {
    document.getElementById(fieldId).addEventListener('keydown', e => {
      if (e.key === 'Enter') saveBook();
    });
  });

  // Enter key in new location inputs
  [['newCabinetName', 'confirmNewCabinet'],
   ['newShelfName',   'confirmNewShelf'],
   ['newRowName',     'confirmNewRow']].forEach(([inputId, btnId]) => {
    document.getElementById(inputId).addEventListener('keydown', e => {
      if (e.key === 'Enter')  document.getElementById(btnId).click();
      if (e.key === 'Escape') document.getElementById('cancel' + btnId.replace('confirm', '')).click();
    });
  });

  // ---- Bottom Nav Tabs ----
  document.querySelectorAll('.bottom-tab').forEach(btn => {
    btn.addEventListener('click', () => switchMobileTab(btn.dataset.tab));
  });

  // Management view action cards
  document.getElementById('mgmtAddBook').addEventListener('click', openAddBookModal);

  document.getElementById('mgmtImportExcel').addEventListener('click', () => {
    openAddBookModal();
    switchBookModalTab('excel');
  });

  document.getElementById('mgmtLocations').addEventListener('click', openLocationsModal);

  // ---- Sidebar toggle (mobile) ----
  document.getElementById('filterToggleBtn').addEventListener('click', openSidebar);
  document.getElementById('sidebarCloseBtn').addEventListener('click', closeSidebar);
  document.getElementById('sidebarDoneBtn').addEventListener('click', closeSidebar);
  document.getElementById('sidebarBackdrop').addEventListener('click', closeSidebar);

  // Also close sidebar after selecting a filter on mobile
  document.getElementById('locationTree').addEventListener('click', () => {
    if (window.innerWidth < 768) setTimeout(closeSidebar, 180);
  });

  // ---- Sort ----
  const sortBtn      = document.getElementById('sortBtn');
  const sortDropdown = document.getElementById('sortDropdown');
  const sortBackdrop = document.getElementById('sortBackdrop');

  function openSortDropdown() {
    sortDropdown.classList.add('open');
    sortBackdrop.classList.add('open');
    sortDropdown.querySelectorAll('.sort-option').forEach(opt => {
      opt.classList.toggle('active', opt.dataset.sort === state.sort);
    });
    document.body.style.overflow = 'hidden';
  }

  function closeSortDropdown() {
    sortDropdown.classList.remove('open');
    sortBackdrop.classList.remove('open');
    document.body.style.overflow = '';
  }

  sortBtn.addEventListener('click', e => {
    e.stopPropagation();
    sortDropdown.classList.contains('open') ? closeSortDropdown() : openSortDropdown();
  });

  sortDropdown.addEventListener('click', e => {
    const opt = e.target.closest('.sort-option');
    if (!opt) return;
    state.sort = opt.dataset.sort;
    closeSortDropdown();
    sortBtn.innerHTML = `↕ ${SORT_LABELS[state.sort]}`;
    render();
  });

  sortBackdrop.addEventListener('click', closeSortDropdown);

  document.addEventListener('click', e => {
    if (!sortBtn.contains(e.target) && !sortDropdown.contains(e.target)) {
      closeSortDropdown();
    }
  });

  // ---- Book Modal Tabs ----
  document.getElementById('bookModalTabs').addEventListener('click', e => {
    const tab = e.target.closest('.modal-tab');
    if (!tab) return;
    switchBookModalTab(tab.dataset.tab);
    if (tab.dataset.tab === 'manual') document.getElementById('bookName').focus();
  });

  // ---- Excel Tab ----
  document.getElementById('downloadTemplateBtn').addEventListener('click', downloadTemplate);

  document.getElementById('pickFileBtn').addEventListener('click', () => {
    document.getElementById('excelFileInput').value = '';
    document.getElementById('excelFileInput').click();
  });

  // Also clicking anywhere on the drop zone opens file picker
  document.getElementById('dropZone').addEventListener('click', e => {
    if (e.target.closest('.btn-link') || e.target.closest('.btn-primary')) return;
    document.getElementById('excelFileInput').click();
  });

  document.getElementById('excelFileInput').addEventListener('change', e => {
    const file = e.target.files[0];
    if (file) handleImportFile(file);
  });

  // Drag and drop
  const dropZone = document.getElementById('dropZone');
  dropZone.addEventListener('dragover',  e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
  dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) handleImportFile(file);
  });

  document.getElementById('changeFileBtn').addEventListener('click', resetExcelTab);

  document.getElementById('importModalConfirm').addEventListener('click', confirmImport);

  // ---- Initial load ----
  initApp();
});

async function initApp() {
  showLoadingOverlay(true);
  try {
    await loadData();
    render();
  } catch (e) {
    showToast('שגיאה בטעינת הנתונים: ' + e.message, 'error');
  } finally {
    showLoadingOverlay(false);
  }
}
