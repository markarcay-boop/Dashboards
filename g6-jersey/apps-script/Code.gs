/**
 * G6 Jersey Configurator – Google Sheet backend.
 *
 * Paste this whole file into Extensions > Apps Script of the Google Sheet
 * that should hold the orders, then follow SETUP.md in the same folder.
 *
 * The web page talks to this script:
 *   GET  -> list of all entries (PINs are never sent back)
 *   POST -> {action: "create" | "update" | "delete" | "verify" | "admin" | "setPaid" | "uploadProof" | "proofs", ...}
 *
 * Several people can submit at the same time: every write takes a script
 * lock, so rows are never lost or overwritten.
 */

const SHEET_NAME = 'Orders';
// Members can't add or change entries after this moment (UAE time).
// The organiser (admin key) still can.
const DEADLINE = new Date('2026-11-17T23:59:59+04:00');

const FIELDS = [
  'id', 'fullName', 'sex', 'kitStyle', 'items', 'sleeves', 'jerseyName', 'jerseyNumber',
  'pockets', 'upperSize', 'shortsSize', 'notes', 'createdAt', 'updatedAt', 'pinHash', 'paid', 'proofId', 'proofAt'
];
// Payment proofs are saved privately in this Drive folder (created automatically).
const PROOF_FOLDER_NAME = 'G6 Jersey Payment Proofs';
const PROOF_TYPES = { 'image/jpeg': 'jpg', 'image/png': 'png', 'image/webp': 'webp', 'image/heic': 'heic', 'image/heif': 'heif', 'application/pdf': 'pdf' };
const PROOF_MAX_BYTES = 8 * 1024 * 1024;
// Password used by the Settings (entries) page to mark entries as paid.
// You can override it with a Script Property named SETTINGS_PASSWORD.
const DEFAULT_SETTINGS_PASSWORD = '0000';
const ENUMS = {
  sex: ['Male', 'Female'],
  kitStyle: ['basketball', 'volleyball', 'volleyball_men'],
  items: ['upper', 'shorts', 'full'],
  sleeves: ['', 'with', 'without'],
  pockets: ['', 'with', 'without'],
  upperSize: ['', 'XS', 'S', 'M', 'L', 'XL', '2XL', '3XL', '4XL'],
  shortsSize: ['', 'XS', 'S', 'M', 'L', 'XL', '2XL', '3XL', '4XL']
};

// ---------- helpers ----------
function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, FIELDS.length).setValues([FIELDS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    // Store everything as plain text so "07" and "00" stay as typed.
    sh.getRange(1, 1, sh.getMaxRows(), FIELDS.length).setNumberFormat('@');
  }
  // Keep the header row up to date when new columns are added (e.g. "paid").
  const head = sh.getRange(1, 1, 1, FIELDS.length).getDisplayValues()[0];
  if (head.join('|') !== FIELDS.join('|')) {
    sh.getRange(1, 1, 1, FIELDS.length).setNumberFormat('@').setValues([FIELDS]).setFontWeight('bold');
  }
  return sh;
}

function rows_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(2, 1, last - 1, FIELDS.length).getDisplayValues();
  return values.map((r, i) => {
    const o = { _row: i + 2 };
    FIELDS.forEach((f, j) => (o[f] = r[j] == null ? '' : String(r[j])));
    return o;
  }).filter(o => o.id);
}

function public_(o) {
  const c = {};
  FIELDS.forEach(f => { if (f !== 'pinHash' && f !== 'proofId') c[f] = o[f] || ''; });
  c.hasProof = !!o.proofId;
  return c;
}

function proofFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('PROOF_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const folder = DriveApp.createFolder(PROOF_FOLDER_NAME);
  props.setProperty('PROOF_FOLDER_ID', folder.getId());
  return folder;
}

// Run this once from the Apps Script editor (select it, click Run) if
// Google needs you to grant Drive permission for payment proofs.
function authorizeDrive() {
  proofFolder_();
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function hash_(pin) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'g6-jersey:' + pin, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(bytes);
}

function isSettings_(pw) {
  const real = PropertiesService.getScriptProperties().getProperty('SETTINGS_PASSWORD') || DEFAULT_SETTINGS_PASSWORD;
  return typeof pw === 'string' && pw === real;
}

function isAdmin_(key) {
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  return !!real && typeof key === 'string' && key === real;
}

// Text that Sheets would treat as a formula is stored with a leading apostrophe.
function safe_(v) {
  const s = String(v == null ? '' : v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function clean_(e) {
  e = e || {};
  const str = (v, max) => String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max);
  const c = {
    fullName: str(e.fullName, 60),
    sex: str(e.sex, 10),
    kitStyle: str(e.kitStyle, 20),
    items: str(e.items, 10),
    sleeves: str(e.sleeves, 10),
    jerseyName: str(e.jerseyName, 14).toUpperCase(),
    jerseyNumber: str(e.jerseyNumber, 2),
    pockets: str(e.pockets, 10),
    upperSize: str(e.upperSize, 4),
    shortsSize: str(e.shortsSize, 4),
    notes: str(e.notes, 200)
  };
  if (!c.fullName) return null;
  for (const k in ENUMS) if (ENUMS[k].indexOf(c[k]) < 0) return null;
  const hasU = c.items === 'upper' || c.items === 'full';
  const hasS = c.items === 'shorts' || c.items === 'full';
  if (hasU && (!c.jerseyName || !/^\d{1,2}$/.test(c.jerseyNumber) || !c.upperSize || !c.sleeves)) return null;
  if (hasS && (!c.shortsSize || !c.pockets)) return null;
  if (!hasU) { c.jerseyName = ''; c.jerseyNumber = ''; c.upperSize = ''; c.sleeves = ''; }
  if (!hasS) { c.shortsSize = ''; c.pockets = ''; }
  return c;
}

function writeRow_(sh, row, obj) {
  const range = sh.getRange(row, 1, 1, FIELDS.length);
  range.setNumberFormat('@');
  range.setValues([FIELDS.map(f => safe_(obj[f]))]);
}

function list_() { return rows_().map(public_); }

// ---------- web endpoints ----------
function doGet() {
  return out_({ ok: true, entries: list_() });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'invalid' }); }
  const action = req.action;
  const admin = isAdmin_(req.adminKey);

  if (action === 'admin') return out_(admin ? { ok: true } : { ok: false, error: 'not_admin' });

  if (action === 'verify') {
    const r = rows_().find(x => x.id === req.id);
    if (!r) return out_({ ok: false, error: 'not_found' });
    if (!r.pinHash || r.pinHash !== hash_(String(req.pin || ''))) return out_({ ok: false, error: 'bad_pin' });
    return out_({ ok: true });
  }

  if (action === 'proofs') {
    if (!admin && !isSettings_(req.password)) return out_({ ok: false, error: 'not_admin' });
    const map = {};
    rows_().forEach(r => { if (r.proofId) map[r.id] = 'https://drive.google.com/file/d/' + r.proofId + '/view'; });
    return out_({ ok: true, proofs: map });
  }

  if (action === 'uploadProof') {
    const r0 = rows_().find(x => x.id === req.id);
    if (!r0) return out_({ ok: false, error: 'not_found' });
    const pinOk = !!r0.pinHash && r0.pinHash === hash_(String(req.pin || ''));
    if (!admin && !pinOk && !isSettings_(req.password)) return out_({ ok: false, error: 'bad_pin' });
    const mime = String(req.mimeType || '').toLowerCase();
    const ext = PROOF_TYPES[mime];
    if (!ext || !req.data) return out_({ ok: false, error: 'bad_file' });
    let bytes;
    try { bytes = Utilities.base64Decode(String(req.data)); } catch (e) { return out_({ ok: false, error: 'bad_file' }); }
    if (!bytes.length || bytes.length > PROOF_MAX_BYTES) return out_({ ok: false, error: 'too_large' });
    const label = (r0.jerseyName || r0.fullName || 'entry').replace(/[^\w .-]/g, '').slice(0, 40);
    const stamp = Utilities.formatDate(new Date(), 'Asia/Dubai', 'yyyy-MM-dd HHmm');
    let file;
    try {
      file = proofFolder_().createFile(Utilities.newBlob(bytes, mime, `${label} ${r0.jerseyNumber ? '#' + r0.jerseyNumber + ' ' : ''}- ${r0.fullName} - ${stamp}.${ext}`));
    } catch (e) {
      return out_({ ok: false, error: 'drive', message: String(e) });
    }
    const lk = LockService.getScriptLock();
    if (!lk.tryLock(20000)) return out_({ ok: false, error: 'busy' });
    try {
      const r = rows_().find(x => x.id === req.id);
      if (!r) return out_({ ok: false, error: 'not_found' });
      if (r.proofId) { try { DriveApp.getFileById(r.proofId).setTrashed(true); } catch (e) {} }
      const now = new Date().toISOString();
      writeRow_(sheet_(), r._row, Object.assign({}, r, { proofId: file.getId(), proofAt: now, paid: r.paid || now }));
      return out_({ ok: true, entries: list_() });
    } finally {
      lk.releaseLock();
    }
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return out_({ ok: false, error: 'busy' });
  try {
    const sh = sheet_();
    const now = new Date().toISOString();
    const closed = new Date() > DEADLINE;

    if (action === 'create') {
      if (closed && !admin) return out_({ ok: false, error: 'closed' });
      const c = clean_(req.entry);
      if (!c) return out_({ ok: false, error: 'invalid' });
      const pin = String(req.pin || '');
      if (!admin && !/^\d{4,6}$/.test(pin)) return out_({ ok: false, error: 'invalid' });
      if (pin && !/^\d{4,6}$/.test(pin)) return out_({ ok: false, error: 'invalid' });
      const id = Utilities.getUuid();
      const row = Object.assign({ id: id, createdAt: now, updatedAt: now, pinHash: pin ? hash_(pin) : '' }, c);
      writeRow_(sh, sh.getLastRow() + 1, row);
      return out_({ ok: true, id: id, entries: list_() });
    }

    if (action === 'update') {
      const r = rows_().find(x => x.id === req.id);
      if (!r) return out_({ ok: false, error: 'not_found' });
      if (!admin) {
        if (closed) return out_({ ok: false, error: 'closed' });
        if (!r.pinHash || r.pinHash !== hash_(String(req.pin || ''))) return out_({ ok: false, error: 'bad_pin' });
      }
      const c = clean_(req.entry);
      if (!c) return out_({ ok: false, error: 'invalid' });
      const row = Object.assign({}, r, c, { updatedAt: now });
      writeRow_(sh, r._row, row);
      return out_({ ok: true, entries: list_() });
    }

    if (action === 'setPaid') {
      if (!admin && !isSettings_(req.password)) return out_({ ok: false, error: 'not_admin' });
      const r = rows_().find(x => x.id === req.id);
      if (!r) return out_({ ok: false, error: 'not_found' });
      const row = Object.assign({}, r, { paid: req.paid ? now : '' });
      writeRow_(sh, r._row, row);
      return out_({ ok: true, entries: list_() });
    }

    if (action === 'delete') {
      if (!admin) return out_({ ok: false, error: 'not_admin' });
      const r = rows_().find(x => x.id === req.id);
      if (!r) return out_({ ok: false, error: 'not_found' });
      sh.deleteRow(r._row);
      return out_({ ok: true, entries: list_() });
    }

    return out_({ ok: false, error: 'invalid' });
  } finally {
    lock.releaseLock();
  }
}
