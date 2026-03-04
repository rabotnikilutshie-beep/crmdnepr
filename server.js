const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const fs = require('fs');
const app = express();
const path = require('path');
const UPLOAD_DIR = path.join(__dirname, 'uploads');
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, UPLOAD_DIR),
  filename: (req, file, cb) => {
    const ext = path.extname(file.originalname || '').toLowerCase();
    const base = Date.now() + '_' + Math.random().toString(16).slice(2);
    cb(null, base + ext);
  }
});
const upload = multer({ storage });

app.use((req, res, next) => {
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
  res.setHeader('Pragma', 'no-cache');
  res.setHeader('Expires', '0');
  next();
});

app.use(express.json());
app.use(express.static(__dirname));
app.use('/uploads', express.static(UPLOAD_DIR));

const DB_FILES = {
  users: 'users.json',
  history: 'history.json',
  orders: 'orders.json',
  kupat: 'kupat_db.json',
  db_list: 'databases.json',
  deleted_db_list: 'deleted_databases.json',
  notes: 'notes.json',
  used_phones: 'used_phones.json',
};

function dbPath(fileName) {
  return path.join(__dirname, String(fileName || ''));
}

DB_FILES.export_auto_ndz = 'export_auto_ndz.xlsx';
DB_FILES.export_kupat = 'export_kupat.xlsx';


DB_FILES.stats_log = 'stats_log.json';
DB_FILES.stats_xlsx = 'stats_by_day.xlsx';

const CALLBACKS_FILE = dbPath(DB_FILES.callbacks || 'callbacks.json');

function loadData(file, def) {
  if (fs.existsSync(file)) {
    try {
      return JSON.parse(fs.readFileSync(file, 'utf8'));
    } catch (e) {
      return def;
    }
  }
  return def;
}
function saveData(file, data) {
  const payload = JSON.stringify(data, null, 2);
  try {
    fs.writeFileSync(file, payload);
  } catch (e) {
    if (fs.existsSync(file)) {
      try {
        const st = fs.statSync(file);
        if (st.isDirectory()) {
          fs.rmSync(file, { recursive: true, force: true });
          fs.writeFileSync(file, payload);
          return;
        }
      } catch (_) {}
    }
    throw e;
  }
}

function loadArrayData(file) {
  const rows = loadData(file, []);
  return Array.isArray(rows) ? rows : [];
}

if (!fs.existsSync(DB_FILES.notes)) saveData(DB_FILES.notes, []);
if (!fs.existsSync(DB_FILES.deleted_db_list)) saveData(DB_FILES.deleted_db_list, []);



function ensureExportWorkbook() {
  const file = DB_FILES.export_auto_ndz;
  if (!fs.existsSync(file)) {
    const wb = xlsx.utils.book_new();
    const autoSheet = xlsx.utils.aoa_to_sheet([["Дата","Время","Оператор","Имя","Телефон","ТЗ","Адрес","Возраст","Доп.Инфа","Заметка"]]);
    const ndzSheet  = xlsx.utils.aoa_to_sheet([["Дата","Время","Оператор","Имя","Телефон","ТЗ","Адрес","Возраст","Доп.Инфа","Заметка"]]);
    xlsx.utils.book_append_sheet(wb, autoSheet, "AUTO");
    xlsx.utils.book_append_sheet(wb, ndzSheet,  "НДЗ");
    xlsx.writeFile(wb, file);
  }
}

function appendAutoNdzToExcel(kind, payload) {
  ensureExportWorkbook();
  const file = DB_FILES.export_auto_ndz;
  const wb = xlsx.readFile(file);
  const sheetName = (kind === "AUTO") ? "AUTO" : "НДЗ";
  let ws = wb.Sheets[sheetName];
  if (!ws) {
    ws = xlsx.utils.aoa_to_sheet([["Дата","Время","Оператор","Имя","Телефон","ТЗ","Адрес","Возраст","Доп.Инфа","Заметка"]]);
    xlsx.utils.book_append_sheet(wb, ws, sheetName);
  }
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1 });
  const now = new Date();
  const d = now.toISOString().slice(0,10);
  const t = now.toTimeString().slice(0,8);
  rows.push([
    d,
    t,
    payload.operator || "",
    payload.name || "",
    payload.phone || "",
    payload.tz || "",
    payload.address || "",
    payload.age || "",
    payload.extra || "",
    payload.note || ""
  ]);
  const newWs = xlsx.utils.aoa_to_sheet(rows);
  wb.Sheets[sheetName] = newWs;
  xlsx.writeFile(wb, file);
}

function ensureKupatWorkbook() {
  const file = DB_FILES.export_kupat;
  if (!fs.existsSync(file)) {
    const wb = xlsx.utils.book_new();
    const ws = xlsx.utils.aoa_to_sheet([["Дата","Время","Оператор","Имя","Телефон","Комментарий"]]);
    xlsx.utils.book_append_sheet(wb, ws, "KUPAT");
    xlsx.writeFile(wb, file);
  }
}
function appendKupatToExcel(payload) {
  ensureKupatWorkbook();
  const file = DB_FILES.export_kupat;
  const wb = xlsx.readFile(file);
  const ws = wb.Sheets["KUPAT"] || xlsx.utils.aoa_to_sheet([["Дата","Время","Оператор","Имя","Телефон","Комментарий"]]);
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1 });
  const t = nowUtc2Str();
  rows.push([t.slice(0,10), t.slice(11,19), payload.operator||"", payload.name||"", payload.phone||"", payload.note||""]);
  wb.Sheets["KUPAT"] = xlsx.utils.aoa_to_sheet(rows);
  if (!wb.SheetNames.includes("KUPAT")) wb.SheetNames.push("KUPAT");
  xlsx.writeFile(wb, file);
}

function nowIso() { return new Date().toISOString(); }

function nowUtc2Str(){
  const d = new Date(Date.now() + 2*60*60*1000);
  return d.toISOString().replace('T',' ').slice(0,19);
}

function formatUtc2DateTime(dateObj){
  const d = dateObj instanceof Date ? dateObj : new Date(dateObj);
  if (Number.isNaN(d.getTime())) return nowUtc2Str();
  const shifted = new Date(d.getTime() + 2*60*60*1000);
  return shifted.toISOString().replace('T',' ').slice(0,19);
}

function parseUtc2InputToDate(value) {
  const raw = String(value || '').trim();
  const m = raw.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})$/);
  if (!m) return null;
  const [, y, mo, d, h, mi] = m;
  const utcMs = Date.UTC(Number(y), Number(mo) - 1, Number(d), Number(h) - 2, Number(mi), 0, 0);
  const dt = new Date(utcMs);
  return Number.isNaN(dt.getTime()) ? null : dt;
}

function isoDate(dIso) { return String(dIso || '').slice(0, 10); }
function isoTime(dIso) { const s=String(dIso||''); return s.includes('T') ? s.slice(11,19) : s.slice(11,19); }

const STATS_FLUSH_DEBOUNCE_MS = 1500;
let statsLogCache = loadData(DB_FILES.stats_log, []);
if (!Array.isArray(statsLogCache)) statsLogCache = [];
let statsRowsBySheet = new Map();
let statsWbLoaded = false;
let statsDirty = false;
let statsFlushTimer = null;

function ensureStatsWorkbookLoaded() {
  if (statsWbLoaded) return;
  const wb = fs.existsSync(DB_FILES.stats_xlsx) ? xlsx.readFile(DB_FILES.stats_xlsx) : xlsx.utils.book_new();
  for (const sheetName of (wb.SheetNames || [])) {
    const rows = xlsx.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' });
    statsRowsBySheet.set(sheetName, Array.isArray(rows) ? rows : []);
  }
  statsWbLoaded = true;
}

function flushStatsToDisk() {
  if (!statsDirty) return;
  statsDirty = false;
  if (statsFlushTimer) {
    clearTimeout(statsFlushTimer);
    statsFlushTimer = null;
  }

  try {
    saveData(DB_FILES.stats_log, statsLogCache);
  } catch (e) {}

  try {
    ensureStatsWorkbookLoaded();
    const wb = xlsx.utils.book_new();
    const headers = ['Time', 'Login', 'Role', 'Action', 'Phone', 'Extra'];
    for (const [sheetName, rows] of statsRowsBySheet.entries()) {
      const safeRows = Array.isArray(rows) ? rows : [];
      const ws = xlsx.utils.json_to_sheet(safeRows, { header: headers });
      xlsx.utils.book_append_sheet(wb, ws, sheetName);
    }
    xlsx.writeFile(wb, DB_FILES.stats_xlsx);
  } catch (e) {}
}

function scheduleStatsFlush() {
  if (statsFlushTimer) return;
  statsFlushTimer = setTimeout(() => {
    statsFlushTimer = null;
    flushStatsToDisk();
  }, STATS_FLUSH_DEBOUNCE_MS);
}

function appendStatsEvent(evt) {
  statsLogCache.push(evt);
  ensureStatsWorkbookLoaded();
  const sheetName = isoDate(evt.ts) || 'unknown';
  const rows = statsRowsBySheet.get(sheetName) || [];
  rows.push({
    Time: isoTime(evt.ts),
    Login: evt.login || '',
    Role: evt.role || '',
    Action: evt.action || '',
    Phone: evt.phone || '',
    Extra: evt.extra || ''
  });
  statsRowsBySheet.set(sheetName, rows);
  statsDirty = true;
  scheduleStatsFlush();
}

function isAdmin(login, role) {
  if (role !== 'admin') return false;
  const u = (users || []).find(x => x.login === login);
  return !!u && u.role === 'admin';
}

let users = loadData(DB_FILES.users, [{ login: 'admin', pass: 'admin123', role: 'admin', balance: 0 }]);
users = users.map(u => ({...u, balance: Number(u.balance || 0)}));
let callResults = loadData(DB_FILES.history, []);
let orders = loadData(DB_FILES.orders, []); 
let kupatOrders = loadData(DB_FILES.kupat, []); 

let usedPhones = new Set();
let usedIndices = new Set();

function normPhone(p) {
  const digits = String(p || '').replace(/\D/g, '');
  if (!digits) return '';

  if (digits.startsWith('972')) {
    let local = digits.slice(3);
    local = local.replace(/^0+/, '');
    return local ? ('0' + local) : '';
  }

  if (digits.startsWith('0')) {
    return digits.replace(/^0+/, '0');
  }

  return digits;
}

function loadUsedPhones() {
  const arr = loadData(DB_FILES.used_phones, []);
  usedPhones = new Set((arr || []).map(normPhone).filter(Boolean));
  (callResults || []).forEach((r) => {
    const ph = normPhone(r.phone);
    if (ph) usedPhones.add(ph);
  });
  (orders || []).forEach((o) => {
    const ph = normPhone(o.phone);
    if (ph) usedPhones.add(ph);
  });
  (kupatOrders || []).forEach((o) => {
    const ph = normPhone(o.phone);
    if (ph) usedPhones.add(ph);
  });
}

function saveUsedPhones() {
  try {
    saveData(DB_FILES.used_phones, Array.from(usedPhones));
  } catch (e) {}
}

loadUsedPhones();

let databases = loadData(DB_FILES.db_list, []);
let deletedDatabases = loadData(DB_FILES.deleted_db_list, []);
let activeDbId = null;
const DB_LIST_CACHE_TTL_MS = 3000;
const dbListCache = new Map();

function dbListCacheKey(view, page, limit) {
  return `${String(view || 'full')}|${Number(page || 1)}|${Number(limit || 20)}`;
}

function clearDbListCache() {
  dbListCache.clear();
}

function getCachedDbList(view, page, limit) {
  const key = dbListCacheKey(view, page, limit);
  const cached = dbListCache.get(key);
  if (!cached) return null;
  if (Date.now() > cached.expiresAt) {
    dbListCache.delete(key);
    return null;
  }
  return cached.payload;
}

function setCachedDbList(view, page, limit, payload) {
  const key = dbListCacheKey(view, page, limit);
  dbListCache.set(key, {
    expiresAt: Date.now() + DB_LIST_CACHE_TTL_MS,
    payload
  });
}

function dbAvailableRowsCount(db) {
  if (!db || !Array.isArray(db.rows)) return 0;
  let count = 0;
  for (let i = 0; i < db.rows.length; i++) {
    const row = db.rows[i];
    const ph = normPhone(row && row[1]);
    if (!ph) continue;
    if (usedPhones.has(ph)) continue;
    count++;
  }
  return count;
}

function normalizeDbMeta(db) {
  if (!db || typeof db !== 'object') return db;
  if (!db.uploadedAt) db.uploadedAt = nowUtc2Str();
  if (!db.lastActivatedAt) db.lastActivatedAt = null;
  return db;
}

databases = (Array.isArray(databases) ? databases : []).map(normalizeDbMeta);
deletedDatabases = Array.isArray(deletedDatabases) ? deletedDatabases : [];


// вход дебилов
app.post('/api/login', (req, res) => {
  const user = users.find((u) => u.login === req.body.login && u.pass === req.body.pass);
  if (user) res.json({ success: true, user });
  else res.status(401).json({ success: false });
});
app.get('/api/notes', (req, res) => {
  const login = (req.query.login || '').toString();
  if (!login) return res.status(400).send('login required');
  const notes = loadData(DB_FILES.notes, []);
  const mine = (Array.isArray(notes) ? notes : []).filter(n => n && n.owner === login);
  res.json(mine);
});

app.post('/api/notes/add', upload.single('file'), (req, res) => {
  const login = (req.body.login || '').toString();
  const note = req.body.note ? (typeof req.body.note === 'string' ? JSON.parse(req.body.note) : req.body.note) : req.body;

  if (!login) {
    if (req.file) { try { fs.unlinkSync(req.file.path); } catch (e) {} }
    return res.status(400).send('login required');
  }

  const comment = (note.comment || '').toString().trim();
  const hasFile = !!req.file;
  if (!comment && !hasFile) {
    if (req.file) { try { fs.unlinkSync(req.file.path); } catch (e) {} }
    return res.status(400).send('comment or file required');
  }

  const role = (req.body.role || note.role || '').toString();
  if (hasFile && !(role === 'tech' || role === 'closer' || role === 'admin')) {
    try { fs.unlinkSync(req.file.path); } catch (e) {}
    return res.status(403).send('forbidden_file');
  }

  const notes = loadData(DB_FILES.notes, []);
  const item = {
    id: Date.now() + '_' + Math.random().toString(16).slice(2),
    owner: login,
    source: (note.source || '').toString(),
    orderId: note.orderId || null,
    clientName: (note.clientName || '').toString(),
    phone: (note.phone || '').toString(),
    text: (note.text || '').toString(),
    comment: (note.comment || '').toString(),
    createdAt: note.createdAt || new Date().toISOString()
  };
  if (hasFile) item.file = '/uploads/' + req.file.filename;

  notes.push(item);
  saveData(DB_FILES.notes, notes);
  res.json({ ok: true, id: item.id });
});

app.put('/api/notes/update', upload.single('file'), (req, res) => {
  const login = (req.body.login || '').toString();
  const id = (req.body.id || '').toString();
  const comment = (req.body.comment || '').toString();
  const role = (req.body.role || '').toString();

  if(!login) {
    if (req.file) { try { fs.unlinkSync(req.file.path); } catch (e) {} }
    return res.status(400).send('login required');
  }
  if(!id) {
    if (req.file) { try { fs.unlinkSync(req.file.path); } catch (e) {} }
    return res.status(400).send('id required');
  }

  const notes = loadData(DB_FILES.notes, []);
  const idx = (Array.isArray(notes) ? notes : []).findIndex(n => n && String(n.id) === String(id));
  if(idx < 0) {
    if (req.file) { try { fs.unlinkSync(req.file.path); } catch (e) {} }
    return res.status(404).send('not found');
  }
  if(notes[idx].owner !== login) {
    if (req.file) { try { fs.unlinkSync(req.file.path); } catch (e) {} }
    return res.status(403).send('forbidden');
  }

  if (req.file && !(role === 'tech' || role === 'closer' || role === 'admin')) {
    try { fs.unlinkSync(req.file.path); } catch (e) {}
    return res.status(403).send('forbidden_file');
  }

  notes[idx].comment = comment;
  if (req.file) {
    if (notes[idx].file) {
      const p = path.join(__dirname, notes[idx].file.replace('/uploads/', 'uploads/'));
      try { fs.unlinkSync(p); } catch (e) {}
    }
    notes[idx].file = '/uploads/' + req.file.filename;
  }
  notes[idx].updatedAt = new Date().toISOString();
  saveData(DB_FILES.notes, notes);
  res.json({ ok: true });
});

app.post('/api/notes/delete', (req, res) => {
  const login = (req.body.login || '').toString();
  const id = (req.body.id || '').toString();
  if(!login) return res.status(400).send('login required');
  if(!id) return res.status(400).send('id required');

  const notes = loadData(DB_FILES.notes, []);
  const before = Array.isArray(notes) ? notes.length : 0;
  const filtered = (Array.isArray(notes) ? notes : []).filter(n => !(n && String(n.id) === String(id) && n.owner === login));
  if(filtered.length === before) return res.status(404).send('not found');
  saveData(DB_FILES.notes, filtered);
  res.json({ ok: true });
});

app.post('/api/save-result', (req, res) => {
  const { status, user, role, phone, note, name, skipOrderCreate } = req.body;
  const time = nowUtc2Str();

  callResults.push({ id: Date.now(), status, user, phone, name, note: note ?? '', time });
  saveData(DB_FILES.history, callResults);

const st = (status || '').toString().toUpperCase();
if (st === 'AUTO' || st === 'АВТО') {
  appendAutoNdzToExcel('AUTO', { operator: user, name, phone, tz: '', address: '', age: '', extra: '', note: note ?? '' });
} else if (st === 'НДЗ' || st === 'NDZ') {
  appendAutoNdzToExcel('НДЗ', { operator: user, name, phone, tz: '', address: '', age: '', extra: '', note: note ?? '' });
}

  const nph = normPhone(phone);
  if(nph){ usedPhones.add(nph); saveUsedPhones(); }

  if (status === 'ПЕРЕДАЛ' && !skipOrderCreate) {
    orders.push({
      id: Date.now(),
      operator: user,
      opToTechAt: nowUtc2Str(),
      clientName: name,
      phone,
      details: note ?? '',
      time,
      status: 'new',
      tech: null,
      closer: null,
    });
    saveData(DB_FILES.orders, orders);
  } else if (status === 'НА КУПАТ') {
    appendKupatToExcel({ operator: user, name, phone, note: note ?? '' });

    kupatOrders.push({
      id: Date.now(),
      operatorFrom: user,
      opToKupatAt: nowUtc2Str(),
      clientName: name,
      phone,
      details: note ?? '',
      time,
      status: 'kupat_new',
      kupatUser: null,
      closer: null,
    });
    saveData(DB_FILES.kupat, kupatOrders);
  } else if (status === 'НА БАНК' && role === 'kupat') {
    orders.push({
      id: Date.now(),
      operator: user,
      clientName: name,
      phone,
      details: note ?? '',
      time,
      status: 'closer',
      tech: null,
      closer: null,
    });
    saveData(DB_FILES.orders, orders);
  }

  res.json({ success: true });
});



// ебучие перезвоны
app.get('/api/callbacks', (req, res) => {
  const login = String((req.query || {}).login || '').trim();
  if (!login) return res.status(400).json({ error: 'login_required' });
  const rows = loadArrayData(CALLBACKS_FILE);
  const mine = (Array.isArray(rows) ? rows : []).filter(x => x && x.owner === login);
  res.json(mine);
});

app.post('/api/callbacks/add', (req, res) => {
  const { login, role, source, sourceId, clientName, phone, text, card, callbackAtUtc2 } = req.body || {};
  if (!login) return res.status(400).json({ success: false, error: 'login_required' });
  if (!['user','kupat','tech','admin'].includes(String(role || ''))) return res.status(403).json({ success: false, error: 'forbidden' });
  const item = {
    id: Date.now().toString(36) + '_' + Math.random().toString(36).slice(2,7),
    owner: String(login),
    role: String(role || ''),
    source: String(source || ''),
    sourceId: sourceId != null ? String(sourceId) : '',
    clientName: String(clientName || ''),
    phone: String(phone || ''),
    text: String(text || ''),
    callbackAtUtc2: String(callbackAtUtc2 || ''),
    createdAt: nowUtc2Str()
  };
  if (card && typeof card === 'object') item.card = card;
  const rows = loadArrayData(CALLBACKS_FILE);
  rows.push(item);
  saveData(CALLBACKS_FILE, rows);
  res.json({ success: true, item });
});

app.post('/api/callbacks/update', (req, res) => {
  const { login, id, clientName, phone, text } = req.body || {};
  if (!login || !id) return res.status(400).json({ success: false, error: 'bad_request' });
  const rows = loadArrayData(CALLBACKS_FILE);
  const idx = rows.findIndex(x => x && String(x.id) === String(id));
  if (idx < 0) return res.status(404).json({ success: false, error: 'not_found' });
  if (rows[idx].owner !== login) return res.status(403).json({ success: false, error: 'forbidden' });
  rows[idx].clientName = String(clientName || '');
  rows[idx].phone = String(phone || '');
  rows[idx].text = String(text || '');
  rows[idx].updatedAt = nowUtc2Str();
  saveData(CALLBACKS_FILE, rows);
  res.json({ success: true, item: rows[idx] });
});

app.post('/api/callbacks/delete', (req, res) => {
  const { login, id } = req.body || {};
  if (!login || !id) return res.status(400).json({ success: false, error: 'bad_request' });
  const rows = loadArrayData(CALLBACKS_FILE);
  const filtered = rows.filter(x => !(x && String(x.id) === String(id) && x.owner === login));
  saveData(CALLBACKS_FILE, filtered);
  res.json({ success: true });
});

app.post('/api/callbacks/transfer-tech', (req, res) => {
  const { login, role, id, callbackAtUtc2 } = req.body || {};
  if (!login || !id) return res.status(400).json({ success: false, error: 'bad_request' });
  if (!['user','admin'].includes(String(role || ''))) return res.status(403).json({ success: false, error: 'forbidden' });

  const rows = loadArrayData(CALLBACKS_FILE);
  const idx = rows.findIndex(x => x && String(x.id) === String(id));
  if (idx < 0) return res.status(404).json({ success: false, error: 'not_found' });
  const c = rows[idx];
  if (c.owner !== login) return res.status(403).json({ success: false, error: 'forbidden' });

  const selectedCallbackAt = parseUtc2InputToDate(callbackAtUtc2);
  if (!selectedCallbackAt) return res.status(400).json({ success: false, error: 'callback_time_required' });

  const diffMinutes = Math.floor((selectedCallbackAt.getTime() - Date.now()) / 60000);
  const delayedForTech = diffMinutes > 30;
  const techAvailableAt = delayedForTech ? selectedCallbackAt : new Date();

  orders.push({
    id: Date.now(),
    operator: login,
    opToTechAt: nowUtc2Str(),
    clientName: c.clientName || '',
    phone: c.phone || '',
    details: c.text || '',
    time: nowUtc2Str(),
    callbackAtUtc2: String(callbackAtUtc2),
    availableAt: formatUtc2DateTime(techAvailableAt),
    status: 'new',
    tech: null,
    closer: null,
  });
  saveData(DB_FILES.orders, orders);

  rows.splice(idx, 1);
  saveData(CALLBACKS_FILE, rows);
  res.json({ success: true, delayedForTech, availableAt: formatUtc2DateTime(techAvailableAt) });
});

app.post('/api/callbacks/transfer-closer', (req, res) => {
  const { login, role, id } = req.body || {};
  if (!login || !id) return res.status(400).json({ success: false, error: 'bad_request' });
  if (!['kupat','admin'].includes(String(role || ''))) return res.status(403).json({ success: false, error: 'forbidden' });

  const rows = loadArrayData(CALLBACKS_FILE);
  const idx = rows.findIndex(x => x && String(x.id) === String(id));
  if (idx < 0) return res.status(404).json({ success: false, error: 'not_found' });
  const c = rows[idx];
  if (c.owner !== login) return res.status(403).json({ success: false, error: 'forbidden' });

  let target = null;
  if (c.source === 'kupat' && c.sourceId) {
    target = kupatOrders.find(o => String(o.id) === String(c.sourceId));
  }
  if (!target) {
    target = {
      id: Date.now(),
      operatorFrom: c.owner,
      opToKupatAt: nowUtc2Str(),
      clientName: c.clientName || '',
      phone: c.phone || '',
      details: c.text || '',
      time: nowUtc2Str(),
      status: 'kupat_to_closer',
      kupatUser: login,
      closer: null,
    };
    kupatOrders.push(target);
  }
  target.status = 'kupat_to_closer';
  target.kupatUser = login;
  if (!target.kupatToCloserAt) target.kupatToCloserAt = nowUtc2Str();
  if (c.text) target.kupatNote = c.text;
  saveData(DB_FILES.kupat, kupatOrders);

  rows.splice(idx, 1);
  saveData(CALLBACKS_FILE, rows);
  res.json({ success: true });
});


app.get('/api/callback', (req, res) => {
  const login = String((req.query || {}).login || '').trim();
  if (!login) return res.status(400).json({ error: 'login_required' });
  const rows = loadArrayData(CALLBACKS_FILE);
  const mine = (Array.isArray(rows) ? rows : []).filter(x => x && x.owner === login);
  res.json(mine);
});
app.post('/api/callback/add', (req, res) => res.redirect(307, '/api/callbacks/add'));
app.post('/api/callback/update', (req, res) => res.redirect(307, '/api/callbacks/update'));
app.post('/api/callback/delete', (req, res) => res.redirect(307, '/api/callbacks/delete'));
app.post('/api/callback/transfer-tech', (req, res) => res.redirect(307, '/api/callbacks/transfer-tech'));
app.post('/api/callback/transfer-closer', (req, res) => res.redirect(307, '/api/callbacks/transfer-closer'));

app.get('/api/call-logs', (req, res) => {
  const { role } = req.query;
  if (role === 'admin') return res.json(callResults);
  return res.json([]);
});

app.get('/api/orders', (req, res) => {
  const { login, role, view } = req.query;
  const norm = (v)=> (v||'').toString().trim();
  const normalizedOrders = orders.map(o=>{
    const techs = Array.isArray(o.techs) ? o.techs.map(norm).filter(Boolean) : (o.tech ? [norm(o.tech)] : []);
    const closers = Array.isArray(o.closers) ? o.closers.map(norm).filter(Boolean) : (o.closer ? [norm(o.closer)] : []);
    return { ...o, techs, closers };
  });


  if (role === 'admin') return res.json(normalizedOrders);

  if (role === 'user') {
    if (view === 'returns') return res.json(orders.filter((o) => o.status === 'return'));
    return res.json(orders.filter((o) => o.operator === login));
  }

  if (role === 'tech') {
    if (view === 'tech-my') {
      return res.json(
        normalizedOrders.filter(
          (o) =>
            o.status === 'tech_work' &&
            (o.tech === login || (Array.isArray(o.techs) && o.techs.includes(login)))
        )
      );
    }
    if (view === 'my-orders') {
      return res.json(
        normalizedOrders.filter(
          (o) => o.tech === login || (Array.isArray(o.techs) && o.techs.includes(login))
        )
      );
    }
    const nowTs = Date.now();
    return res.json(
      orders.filter((o) => {
        const inQueue = o.status === 'new' || o.status === 're-work' || o.status === 'return';
        if (!inQueue) return false;
        if (!o.availableAt) return true;
        const availableDate = parseUtc2InputToDate(String(o.availableAt).slice(0,16).replace(' ', 'T'));
        if (!availableDate) return true;
        return availableDate.getTime() <= nowTs;
      })
    );
  }

  if (role === 'closer') {
    return res.json(orders.filter((o) => o.status === 'closer' || o.status === 'final' || o.closer === login));
  }

  if (role === 'kupat') {
    return res.json(
      orders.filter(
        (o) =>
          o.operator === login &&
          (o.status === 'closer' || o.status === 'final')
      )
    );
  }

  return res.json([]);
});

app.get('/api/my-clients', (req, res) => {
  const { login, role } = req.query;
  if (!login) return res.status(400).json({ error: 'login_required' });
  if (role !== 'user' && role !== 'admin') return res.status(403).json({ error: 'forbidden' });

  const mine = orders.filter((o) => {
    if (role === 'admin') return !!o.closer;
    return o.operator === login && !!o.closer;
  });

  res.json(mine);
});

app.post('/api/orders/action', (req, res) => {
  const { id, status, techLogin, info } = req.body;
  const order = orders.find((o) => String(o.id) === String(id));
  if (order) {
    order.status = status; // 'return', 'refuse', 'closer'
    if (techLogin) {
      order.tech = techLogin;
      if (!Array.isArray(order.techs)) order.techs = [];
      if (!order.techs.includes(techLogin)) order.techs.push(techLogin);
    }
    if (status === 'closer' && !order.techToCloserAt) order.techToCloserAt = nowUtc2Str();
    if (info) order.techInfo = info;
    saveData(DB_FILES.orders, orders);
  }
  res.json({ success: true });
});

app.post('/api/orders/delete-self', (req, res) => {
  const { id, login, role } = req.body || {};
  if (!id || !login || !role) return res.status(400).json({ success: false, error: 'bad_request' });

  const idx = orders.findIndex((o) => String(o.id) === String(id));
  if (idx < 0) return res.status(404).json({ success: false, error: 'not_found' });

  const order = orders[idx];
  const assignedTechs = Array.isArray(order.techs) ? order.techs : (order.tech ? [order.tech] : []);
  const isAssignedTech = assignedTechs.includes(login) || order.tech === login;
  const allowed = role === 'admin' || (role === 'tech' && isAssignedTech);
  if (!allowed) return res.status(403).json({ success: false, error: 'forbidden' });

  orders.splice(idx, 1);
  saveData(DB_FILES.orders, orders);
  return res.json({ success: true });
});
app.post('/api/orders/assign', (req, res) => {
  const { id, listType, assignee, requesterLogin, requesterRole } = req.body || {};
  if (!id || !listType || !assignee) return res.status(400).json({ success:false, error:'bad_request' });

  if (listType === 'tech' && !(requesterRole === 'tech' || requesterRole === 'admin')) {
    return res.status(403).json({ success:false, error:'forbidden' });
  }
  if (listType === 'closer' && !(requesterRole === 'closer' || requesterRole === 'admin')) {
    return res.status(403).json({ success:false, error:'forbidden' });
  }

  const order = orders.find(o => String(o.id) === String(id));
  if (!order) return res.status(404).json({ success:false, error:'not_found' });

  const norm = (v)=> (v||'').toString().trim();
  const a = norm(assignee);
  if (!a) return res.status(400).json({ success:false, error:'bad_assignee' });

  if (listType === 'tech') {
    const arr = Array.isArray(order.techs) ? order.techs.map(norm).filter(Boolean) : (order.tech ? [norm(order.tech)] : []);
    if (!arr.includes(a)) arr.push(a);
    order.techs = arr;
    order.tech = arr[0] || order.tech || '';
  } else if (listType === 'closer') {
    const arr = Array.isArray(order.closers) ? order.closers.map(norm).filter(Boolean) : (order.closer ? [norm(order.closer)] : []);
    if (!arr.includes(a)) arr.push(a);
    order.closers = arr;
    order.closer = arr[0] || order.closer || '';
  } else {
    return res.status(400).json({ success:false, error:'bad_list' });
  }

  saveData(DB_FILES.orders, orders);
  res.json({ success:true, order });
});



app.post('/api/orders/create-direct', (req, res) => {
  const { login, role, clientName, phone, tz, address, age, extra, details } = req.body || {};
  if (!login || !role) return res.status(400).json({ error: 'bad_request' });
  if (role === 'kupat') return res.status(403).json({ error: 'forbidden' });
  if (!phone) return res.status(400).json({ error: 'phone_required' });

  const time = nowUtc2Str();

  const textParts = [];
  if (details) textParts.push(String(details));
  if (tz) textParts.push(`ТЗ: ${tz}`);
  if (address) textParts.push(`Адрес: ${address}`);
  if (age) textParts.push(`Возраст: ${age}`);
  if (extra) textParts.push(`Доп.Инфа: ${extra}`);
  const fullDetails = textParts.join('\n');

  const id = Date.now();
  orders.push({
    id,
    operator: login,
    opToTechAt: time,
    clientName: clientName || '',
    phone,
    details: fullDetails,
    time,
    status: 'new',
    tech: null,
    closer: null,
  });
  saveData(DB_FILES.orders, orders);

  // анти-дубль
  const nph = normPhone(phone);
  if (nph) { usedPhones.add(nph); saveUsedPhones(); }

  try {
    appendStatsEvent({ ts: nowIso(), login: String(login), role: String(role), action: 'CREATE_ORDER', phone: String(phone), extra: '' });
  } catch (e) {}

  res.json({ success: true, id });
});

app.post('/api/orders/to-tech', (req, res) => {
  const order = orders.find((o) => String(o.id) === String(req.body.id));
  if (order) {
    order.status = 're-work';
    order.operatorComment = req.body.comment;
    saveData(DB_FILES.orders, orders);
  }
  res.json({ success: true });
});


app.post('/api/orders/comment', upload.single('file'), (req, res) => {
  const { id, role, login } = req.body;
  const text = (req.body.text ?? '').toString();

  const order = orders.find((o) => String(o.id) === String(id));
  if (!order) return res.status(404).json({ success: false, error: 'not_found' });

  const clean = String(text ?? '').trim();
  const hasFile = !!req.file;

  if (!clean && !hasFile) return res.json({ success: false, error: 'empty' });

  const status = order.status;

  let allowed = false;
  if (role === 'admin') allowed = true;

  if (role === 'user') {
    allowed = status === 'new' || status === 're-work' || status === 'return';
  }

  if (role === 'tech') {
    allowed = status !== 'final' && status !== 'refuse';
    if (allowed && !order.tech) order.tech = login;
  }

  if (role === 'closer') {
    allowed = status === 'closer' || status === 'final' || order.closer === login;
    if (allowed && !order.closer && status === 'closer') order.closer = login;
  }

  if (hasFile && !(role === 'tech' || role === 'closer' || role === 'admin')) {
    try { fs.unlinkSync(req.file.path); } catch (e) {}
    return res.status(403).json({ success: false, error: 'forbidden_file' });
  }

  if (!allowed) {
    if (hasFile) { try { fs.unlinkSync(req.file.path); } catch (e) {} }
    return res.status(403).json({ success: false, error: 'forbidden' });
  }

  if (!Array.isArray(order.comments)) order.comments = [];
  const isDuplicateComment = order.comments.some((c) => {
    if (!c) return false;
    const sameAuthor = String(c.by || '') === String(login || '');
    const sameRole = String(c.role || '') === String(role || '');
    const sameText = String(c.text || '') === clean;
    const sameFile = String(c.file || '') === String(hasFile ? ('/uploads/' + req.file.filename) : '');
    const createdTs = Date.parse(String(c.createdAt || c.time || ''));
    const isRecent = Number.isFinite(createdTs) ? (Date.now() - createdTs) < 15000 : false;
    return sameAuthor && sameRole && sameText && sameFile && isRecent;
  });
  if (isDuplicateComment) {
    return res.json({ success: true, deduplicated: true });
  }

  const item = {
    by: login,
    role,
    text: clean,
    time: nowUtc2Str(),
	createdAt: new Date().toISOString(),
  };
  if (hasFile) item.file = '/uploads/' + req.file.filename;
  order.comments.push(item);

  saveData(DB_FILES.orders, orders);
  return res.json({ success: true });
});

app.post('/api/orders/final', (req, res) => {
  const order = orders.find((o) => String(o.id) === String(req.body.id));
  if (order) {
    order.status = 'final';
    order.closer = req.body.closerLogin;
    order.finalColor = req.body.color;
    order.finalText = req.body.text;
    saveData(DB_FILES.orders, orders);
  }
  res.json({ success: true });
});

app.get('/api/kupat/list', (req, res) => {
  const { login, role, view } = req.query;

  if (role === 'admin') return res.json(kupatOrders);

  if (role === 'kupat') {
    if (view === 'incoming') return res.json(kupatOrders.filter((o) => o.status === 'kupat_new'));
    return res.json(kupatOrders); // для "всё/мои" на фронте
  }

  // Закрыв видит только тех старичков, которых купат пометил "передать"
  if (role === 'closer') {
    return res.json(
      kupatOrders.filter((o) => o.status === 'kupat_to_closer' || o.status === 'kupat_final' || o.closer === login)
    );
  }

  return res.json([]);
});

app.post('/api/kupat/action', (req, res) => {
  const { id, status, note, user } = req.body;
  const order = kupatOrders.find((o) => String(o.id) === String(id));
  if (!order) return res.status(404).json({ success: false, error: 'not_found' });

  if (status === 'kupat_refuse') {
    const refusalNote = String(note || '').trim();
    if (!refusalNote) {
      return res.status(400).json({ success: false, error: 'refuse_note_required' });
    }
  }

  order.status = status;
  order.kupatUser = user;
  if (status === 'kupat_to_closer' && !order.kupatToCloserAt) order.kupatToCloserAt = nowUtc2Str();
  if (typeof note !== 'undefined' && String(note).trim()) order.kupatNote = String(note).trim();
  saveData(DB_FILES.kupat, kupatOrders);

  return res.json({ success: true });
});

app.post('/api/kupat/assign-closer', (req, res) => {
  const { id, closer, requesterRole } = req.body || {};
  if (!(requesterRole === 'closer' || requesterRole === 'admin')) {
    return res.status(403).json({ success: false, error: 'forbidden' });
  }
  const order = kupatOrders.find((o) => String(o.id) === String(id));
  if (!order) return res.status(404).json({ success: false, error: 'not_found' });

  order.closer = closer || null;
  saveData(DB_FILES.kupat, kupatOrders);
  return res.json({ success: true, order });
});


app.post('/api/kupat/comment', upload.single('file'), (req, res) => {
  const body = req.body || {};
  const { id, login, role } = body;
  const cleanText = String(body.text || '').trim();
  if (!id || !login || (!cleanText && !req.file)) return res.status(400).json({ success: false, error: 'bad_request' });

  const order = kupatOrders.find((o) => String(o.id) === String(id));
  if (!order) return res.status(404).json({ success: false, error: 'not_found' });

  const isAllowed = role === 'admin' || (role === 'kupat' && order.kupatUser === login);
  if (!isAllowed) return res.status(403).json({ success: false, error: 'forbidden' });

  if (!Array.isArray(order.comments)) order.comments = [];
  const isDuplicateComment = order.comments.some((c) => {
    if (!c) return false;
    const sameAuthor = String(c.by || '') === String(login || '');
    const sameRole = String(c.role || '') === String(role || 'kupat');
    const sameText = String(c.text || '') === cleanText.slice(0, 2000);
    const sameFile = String(c.fileUrl || '') === String(req.file && req.file.filename ? `/uploads/${req.file.filename}` : '');
    const createdTs = Date.parse(String(c.createdAt || c.time || ''));
    const isRecent = Number.isFinite(createdTs) ? (Date.now() - createdTs) < 15000 : false;
    return sameAuthor && sameRole && sameText && sameFile && isRecent;
  });
  if (isDuplicateComment) {
    return res.json({ success: true, deduplicated: true, order });
  }

  const item = {
    by: String(login),
    role: String(role || 'kupat'),
    text: cleanText.slice(0, 2000),
    time: nowUtc2Str(),
    createdAt: new Date().toISOString()
  };
  if (req.file && req.file.filename) item.fileUrl = `/uploads/${req.file.filename}`;
  order.comments.push(item);

  saveData(DB_FILES.kupat, kupatOrders);
  return res.json({ success: true, order });
});

app.post('/api/kupat/final', (req, res) => {
  const { id, text, color, closer } = req.body;
  const order = kupatOrders.find((o) => String(o.id) === String(id));
  if (!order) return res.status(404).json({ success: false, error: 'not_found' });

  order.status = 'kupat_final';
  order.closer = closer;
  order.finalText = text;
  order.finalColor = color;
  saveData(DB_FILES.kupat, kupatOrders);

  return res.json({ success: true });
});

app.post('/api/kupat/delete-self', (req, res) => {
  const { id, login, role } = req.body || {};
  if (!id || !login || !role) return res.status(400).json({ success: false, error: 'bad_request' });

  const idx = kupatOrders.findIndex((o) => String(o.id) === String(id));
  if (idx < 0) return res.status(404).json({ success: false, error: 'not_found' });

  const order = kupatOrders[idx];
  const allowed = role === 'admin' || (role === 'kupat' && String(order.kupatUser || '') === String(login));
  if (!allowed) return res.status(403).json({ success: false, error: 'forbidden' });

  kupatOrders.splice(idx, 1);
  saveData(DB_FILES.kupat, kupatOrders);
  return res.json({ success: true });
});

// админка олега любимого
app.post('/api/delete-item', (req, res) => {
  const { id, type } = req.body;
  if (type === 'orders') {
    orders = orders.filter((o) => String(o.id) !== String(id));
    saveData(DB_FILES.orders, orders);
  } else {
    kupatOrders = kupatOrders.filter((o) => String(o.id) !== String(id));
    saveData(DB_FILES.kupat, kupatOrders);
  }
  res.json({ success: true });
});


app.get('/api/users', (req,res)=>{
  res.json({users: users.map(u=>({login:u.login, role:u.role, balance: Number(u.balance||0)}))});
});

app.post('/api/users/delete', (req, res) => {
  const { login, requesterRole } = req.body || {};
  if (requesterRole !== 'admin') return res.status(403).json({ success: false, error: 'forbidden' });
  if (!login) return res.status(400).json({ success: false, error: 'no_login' });

  // защита на дебила
  const adminCount = users.filter(u => u.role === 'admin').length;
  const target = users.find(u => u.login === login);
  if (target?.role === 'admin' && adminCount <= 1) {
    return res.status(400).json({ success: false, error: 'cannot_delete_last_admin' });
  }

  users = users.filter((u) => u.login !== login);
  saveData(DB_FILES.users, users);
  res.json({ success: true });
});


app.post('/api/admin/set-balance', (req,res)=>{
  const {login, balance, requesterRole} = req.body || {};
  if (requesterRole !== 'admin') return res.status(403).json({success:false, error:'forbidden'});
  users = users.map(u => u.login===login ? ({...u, balance: Number(balance||0)}) : u);
  saveData(DB_FILES.users, users);
  res.json({success:true});
});

app.post('/api/users/add', (req, res) => {
  users.push(req.body);
  saveData(DB_FILES.users, users);
  res.json({ success: true });
});


app.get('/api/me-stats', (req,res)=>{
  const login = req.query.login;
  const role = req.query.role;
  const u = users.find(x=>x.login===login);
  const balance = Number(u?.balance || 0);

  const mainMine = orders.filter(o=>{
    if(role==='user') return o.operator===login;
    if(role==='tech') return o.tech===login;
    if(role==='closer') return o.closer===login;
    if(role==='admin') return true;
    return false;
  });
  const main = {
    total: mainMine.length,
    new: mainMine.filter(o=>o.status==='new' || o.status==='re-work' || o.status==='return').length,
    work: mainMine.filter(o=>o.status==='work' || o.status==='tech').length,
    closer: mainMine.filter(o=>o.status==='closer').length,
    final: mainMine.filter(o=>o.status==='final').length,
    refuse: mainMine.filter(o=>o.status==='refuse').length,
  };

  const kupMine = kupatOrders.filter(o=>{
    if(role==='user') return o.operatorFrom===login;
    if(role==='kupat') return o.kupatUser===login;
    if(role==='closer') return o.closer===login;
    if(role==='admin') return true;
    return false;
  });
  const kupat = {
    total: kupMine.length,
    incoming: kupMine.filter(o=>o.status==='kupat_new').length,
    work: kupMine.filter(o=>o.status==='kupat_work').length,
    closer: kupMine.filter(o=>o.status==='closer').length,
    final: kupMine.filter(o=>o.status==='final').length,
    refuse: kupMine.filter(o=>o.status==='refuse' || o.status==='kupat_refuse').length,
  };

  const uCalls = callResults.filter(r => r.user === login);
  const calls = {
    total: uCalls.length,
    passedToTech: orders.filter(o => o.operator === login && !!o.opToTechAt).length,
    sentToKupat: uCalls.filter(r => r.status === 'НА КУПАТ').length,
  };

  const operatorToCloser = orders.filter(o => o.operator === login && !!o.closer).length;
  const operatorFinalByCloser = orders.filter(o => o.operator === login && !!o.closer && o.status === 'final').length;

  const closerWorked = orders.filter(o => o.closer === login).length;
  const closerFinal = orders.filter(o => o.closer === login && o.status === 'final').length;

  res.json({ balance, calls, closer: { operatorToCloser, operatorFinalByCloser, closerWorked, closerFinal }, stats: { main, kupat } });
});

app.get('/api/stats', (req, res) => {
  const stats = users.map((u) => {
    const login = u.login;
    const uCalls = callResults.filter((r) => r.user === login);

    const calls = {
      total: uCalls.length,
      passedToTech: orders.filter((o) => o.operator === login && !!o.opToTechAt).length,
      sentToKupat: uCalls.filter((r) => r.status === 'НА КУПАТ').length,
    };

    const main = {
      created: orders.filter((o) => o.operator === login).length,
      techTook: orders.filter((o) => o.tech === login).length,
      closerTook: orders.filter((o) => o.closer === login).length,
      closerFinal: orders.filter((o) => o.closer === login && o.status === 'final').length,
      closerRefuse: orders.filter((o) => o.closer === login && o.status === 'refuse').length,
    };

    const kupat = {
      fromOperator: kupatOrders.filter((o) => o.operatorFrom === login).length,
      kupatTook: kupatOrders.filter((o) => o.kupatUser === login).length,
      kupatToCloser: kupatOrders.filter((o) => o.kupatUser === login && o.status === 'kupat_to_closer').length,
      kupatFinalByCloser: kupatOrders.filter((o) => o.closer === login && o.status === 'kupat_final').length,
    };

    return {
      login,
      role: u.role,
      balance: Number(u.balance || 0),
      calls,
      main,
      kupat,
    };
  });
  res.json({ byUsers: stats });
});

app.post('/api/upload', upload.single('file'), (req, res) => {
  if (!req.file) return res.json({ error: 'Нет файла' });

  const workbook = xlsx.readFile(req.file.path);
  const rows = xlsx.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1, defval: '' });

  databases.push({
    id: Date.now(),
    name: req.file.originalname,
    uploadedAt: nowUtc2Str(),
    lastActivatedAt: null,
    rows: rows.filter((r) => r.length > 0 && r[0]),
  });
  saveData(DB_FILES.db_list, databases);
  clearDbListCache();

  res.json({ success: true });
});

app.get('/api/databases', (req, res) => {
  const view = String(req.query.view || 'full').toLowerCase() === 'compact' ? 'compact' : 'full';
  const page = Math.max(parseInt(req.query.page, 10) || 1, 1);
  const limit = Math.min(Math.max(parseInt(req.query.limit, 10) || 20, 1), 100);

  const cached = getCachedDbList(view, page, limit);
  if (cached) return res.json(cached);

  const activeBase = (Array.isArray(databases) ? databases : []).map((db) => ({
    id: db.id,
    name: db.name,
    isActive: String(db.id) === String(activeDbId),
    availableRows: dbAvailableRowsCount(db)
  }));

  if (view === 'compact') {
    const payload = { active: activeBase };
    setCachedDbList(view, page, limit, payload);
    return res.json(payload);
  }

  const active = (Array.isArray(databases) ? databases : []).map((db) => ({
    id: db.id,
    name: db.name,
    allowDuplicates: !!db.allowDuplicates,
    uploadedAt: db.uploadedAt || null,
    lastActivatedAt: db.lastActivatedAt || null,
    totalRows: Array.isArray(db.rows) ? db.rows.length : 0,
    availableRows: dbAvailableRowsCount(db),
    isActive: String(db.id) === String(activeDbId)
  }));

  const deletedStart = (page - 1) * limit;
  const deletedSlice = deletedDatabases.slice(deletedStart, deletedStart + limit);
  const deleted = deletedSlice.map((db) => ({
    id: db.id,
    name: db.name,
    uploadedAt: db.uploadedAt || null,
    lastActivatedAt: db.lastActivatedAt || null,
    deletedAt: db.deletedAt || null,
    totalRows: Number(db.totalRows || 0),
    availableRows: Number(db.availableRows || 0),
    deletedBy: db.deletedBy || ''
  }));

  const payload = {
    active,
    deleted,
    pagination: {
      page,
      limit,
      total: deletedDatabases.length,
      totalPages: Math.max(Math.ceil(deletedDatabases.length / limit), 1)
    }
  };
  setCachedDbList(view, page, limit, payload);
  res.json(payload);
});

app.post('/api/databases/delete', (req, res) => {
  const { dbId, login, role } = req.body || {};
  if (!dbId) return res.status(400).json({ success: false, error: 'db_required' });
  if (!(role === 'admin')) return res.status(403).json({ success: false, error: 'forbidden' });

  const idx = databases.findIndex((d) => String(d.id) === String(dbId));
  if (idx < 0) return res.status(404).json({ success: false, error: 'not_found' });

  const db = databases[idx];
  const archiveItem = {
    id: db.id,
    name: db.name,
    uploadedAt: db.uploadedAt || null,
    lastActivatedAt: db.lastActivatedAt || null,
    deletedAt: nowUtc2Str(),
    deletedBy: String(login || ''),
    totalRows: Array.isArray(db.rows) ? db.rows.length : 0,
    availableRows: dbAvailableRowsCount(db)
  };
  deletedDatabases.push(archiveItem);
  saveData(DB_FILES.deleted_db_list, deletedDatabases);

  databases.splice(idx, 1);
  if (String(activeDbId) === String(dbId)) {
    activeDbId = null;
    usedIndices.clear();
  }
  saveData(DB_FILES.db_list, databases);
  clearDbListCache();
  res.json({ success: true, archived: archiveItem });
});

app.post('/api/set-active-db', (req, res) => {
  const dbId = req.body && req.body.dbId;
  activeDbId = dbId;

  const allowDuplicates = !!(req.body && req.body.allowDuplicates);
  activeDbAllowDuplicates = allowDuplicates;

  const db = databases.find((d) => String(d.id) === String(dbId));
  if (db) {
    db.allowDuplicates = allowDuplicates;
    db.lastActivatedAt = nowUtc2Str();
    saveData(DB_FILES.db_list, databases);
  }

  usedIndices.clear();
  clearDbListCache();
  res.json({ success: true, allowDuplicates: activeDbAllowDuplicates });
});

app.get('/api/get-row', (req, res) => {
  if (!activeDbId) return res.json({ error: 'Админ не выбрал базу!' });
  const db = databases.find((d) => d.id == activeDbId);
  if (!db) return res.json({ error: 'База не найдена' });

  const totalRows = Array.isArray(db.rows) ? db.rows.length : 0;
  if (!totalRows) return res.json({ error: 'Нет новых номеров (все уже были или база закончилась)' });

  function isAvailable(i) {
    if (usedIndices.has(i)) return false;
    const row = db.rows[i];
    const ph = normPhone(row && row[1]);
    if (!ph) return false;
    if (!activeDbAllowDuplicates && usedPhones.has(ph)) return false;
    return true;
  }

  let idx = -1;
  const randomAttempts = Math.min(totalRows, 400);
  for (let i = 0; i < randomAttempts; i++) {
    const probe = Math.floor(Math.random() * totalRows);
    if (isAvailable(probe)) {
      idx = probe;
      break;
    }
  }

  if (idx === -1) {
    const start = Math.floor(Math.random() * totalRows);
    for (let offset = 0; offset < totalRows; offset++) {
      const probe = (start + offset) % totalRows;
      if (isAvailable(probe)) {
        idx = probe;
        break;
      }
    }
  }

  if (idx === -1) return res.json({ error: 'Нет новых номеров (все уже были или база закончилась)' });
  usedIndices.add(idx);

  const chosen = db.rows[idx];
  const ph = normPhone(chosen && chosen[1]);
  if (ph) { usedPhones.add(ph); saveUsedPhones(); }

  res.json({ row: chosen });
});


app.post('/api/stats/event', (req, res) => {
  const { login, role, action, phone, extra } = req.body || {};
  if (!login || !action) return res.status(400).json({ error: 'bad_request' });

  const u = (users || []).find(x => x.login === login);
  if (!u) return res.status(403).json({ error: 'unknown_user' });

  const evt = {
    ts: nowUtc2Str(),
    login: String(login),
    role: String(role || u.role || ''),
    action: String(action),
    phone: String(phone || ''),
    extra: extra ? String(extra).slice(0, 500) : ''
  };
  appendStatsEvent(evt);
  res.json({ ok: true });
});

app.get('/api/admin/stats', (req, res) => {
  const { login, role, date } = req.query || {};
  if (!isAdmin(login, role)) return res.status(403).json({ error: 'forbidden' });

  const day = String(date || '').slice(0, 10);
  if (!day) return res.status(400).json({ error: 'date_required' });

  const log = Array.isArray(statsLogCache) ? statsLogCache : [];
  const events = (log || []).filter(e => isoDate(e.ts) === day);

  const byUser = {};
  function ensure(user) {
    if (!byUser[user]) byUser[user] = { login: user, calls: 0, auto: 0, ndz: 0, refuse: 0, ivrit: 0, kupat: 0, passed: 0, closed: 0, first: '', last: '' };
    return byUser[user];
  }
  function bump(user, key) { ensure(user)[key] = (ensure(user)[key] || 0) + 1; }

  for (const e of events) {
    const user = e.login || 'unknown';
    const row = ensure(user);
    const t = isoTime(e.ts);
    if (!row.first || t < row.first) row.first = t;
    if (!row.last || t > row.last) row.last = t;

    const a = String(e.action || '').toUpperCase();
    if (a === 'CALL' || a === 'ПОЗВОНИТЬ') bump(user, 'calls');
    else if (a === 'АВТО') bump(user, 'auto');
    else if (a === 'НДЗ') bump(user, 'ndz');
    else if (a.includes('ОТКАЗ') || a === 'REFUSE') bump(user, 'refuse');
    else if (a.includes('ИВРИТ') || a.includes('IVRIT') || a === 'HEBREW') bump(user, 'ivrit');
    else if (a.includes('КУПАТ')) bump(user, 'kupat');     // НА КУПАТ
    else if (a.includes('ПЕРЕД') || a === 'CREATE_ORDER') bump(user, 'passed');    // ПЕРЕДАЛ / СОЗДАТЬ ЗАЯВКУ
    else if (a.includes('ЗАКР') || a === 'FINAL') bump(user, 'closed'); // ЗАКРЫЛ/FINAL
  }

  for (const u of (users || [])) ensure(u.login);

  res.json({ date: day, rows: Object.values(byUser).sort((a,b)=>a.login.localeCompare(b.login)) });
});

app.get('/api/admin/stats/download', (req, res) => {
  const { login, role } = req.query || {};
  if (!isAdmin(login, role)) return res.status(403).send('forbidden');

  flushStatsToDisk();
  if (!fs.existsSync(DB_FILES.stats_xlsx)) {
    const wb = xlsx.utils.book_new();
    xlsx.writeFile(wb, DB_FILES.stats_xlsx);
  }
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="stats_by_day.xlsx"');
  fs.createReadStream(DB_FILES.stats_xlsx).pipe(res);
});

app.get('/api/export/kupat', (req, res) => {
  const { login, role } = req.query || {};
  if (!isAdmin(login, role)) return res.status(403).send('forbidden');
  ensureKupatWorkbook();
  res.download(DB_FILES.export_kupat);
});



app.get('/api/export/auto-ndz', (req, res) => {
  const { login, role } = req.query || {};
  if (!isAdmin(login, role)) return res.status(403).json({ error: 'forbidden' });
  ensureExportWorkbook();
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="export_auto_ndz.xlsx"');
  fs.createReadStream(DB_FILES.export_auto_ndz).pipe(res);
});

app.listen(3000, '0.0.0.0', () => console.log('CRM v10.0 STABLE is running...'));

process.on('SIGINT', () => {
  flushStatsToDisk();
  process.exit(0);
});

process.on('SIGTERM', () => {
  flushStatsToDisk();
  process.exit(0);
});