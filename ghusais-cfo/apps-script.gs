/* Ghusais CFO activity dashboard – Google Sheet + Drive backend
   Stores every activity, memo and uploaded file in ONE place so all browsers see the same data.

   FIRST-TIME SETUP
   1. Create a new Google Sheet (e.g. "Ghusais CFO Dashboard Data").
   2. Extensions > Apps Script. Delete any code there, paste this whole file, click Save.
   3. In the function dropdown pick "setup" and click Run. Approve the permissions it asks for
      (Sheets + Drive). This creates the Activities / News / Files tabs and a Drive folder.
   4. Deploy > New deployment > type "Web app".
        Execute as: Me      Who has access: Anyone
      Click Deploy and copy the Web app URL (ends with /exec).
   5. Put that URL in the dashboard (const API_URL = '...') – or send it to Claude to do it.

   UPDATING THIS CODE LATER
   Paste over the old code, Save, then Deploy > Manage deployments > Edit (pencil)
   > Version: New version > Deploy. The web app URL stays the same.

   Passwords are checked here on the server. Only salted SHA-256 hashes are stored, never the
   plain passwords, and the hashes are not sent to the public web page. */

const ACCOUNTS = {
  buklod: { user: 'Buklod', role: 'dept',  salt: 'fd345fb6ade4ee18', hash: '3e9235c9c6f7c8fb9f0a9741490c95db6de72636522fe6dd2eb300bd5962270c' },
  kadiwa: { user: 'KADIWA', role: 'dept',  salt: '4e7172f53cd16066', hash: 'a2e02d97952b98652cff5106ea213f00e500881d4c1b3b36e1380f9d8db43604' },
  binhi:  { user: 'Binhi',  role: 'dept',  salt: '9cfcd9c652d403ad', hash: '801bc32d51773e3f21269d7ebe2604d2c4a77f8e11507e71e8236c6ce2e2dfa8' },
  pnk:    { user: 'PNK',    role: 'dept',  salt: '594805d474adb5a6', hash: '0c85346c6efac1ad403b4ec290a879083cdfb487e9b457404e6e65f389118543' },
  scan:   { user: 'SCAN',   role: 'dept',  salt: 'a0e4b47c5c0c7df6', hash: 'f90c91eb43e6c593b894a189ddcf19bafc03a0d4694613582c000bfc7a9e3072' },
  admin:  { user: 'admin',  role: 'admin', salt: 'f8a9e0eab9a3797f', hash: '9e47091e6cad2cdd85db0a51dbf7ca6400a63d1bd8ab380ebaf3bd8d90837d4c' }
};
const DEPT_NAMES = { buklod: 'Buklod', kadiwa: 'KADIWA', binhi: 'Binhi', pnk: 'PNK', scan: 'SCAN' };
const COMMITTEES = ['EPC', 'SDSC', 'LEC', 'ACC', 'ELC', 'SC'];
const STATUSES = ['planned', 'arf', 'hardcopy', 'review', 'returned', 'approved', 'executed', 'reported', 'custom'];
const STATUS_NAMES = {
  planned: 'Planned', arf: 'Activity Request Form submitted', hardcopy: 'Activity Request hard copy sent to district',
  review: 'Activity under review', returned: 'Activity returned for revision', approved: 'Activity approved with UAN number',
  executed: 'Activity executed', reported: 'Activity reported with attendance and photo attachments', custom: 'Custom'
};
const FOLDER_NAME = 'Ghusais CFO Dashboard Files';
const SESSION_SECONDS = 6 * 60 * 60;          // a login lasts 6 hours
const MAX_FILE_BYTES = 10 * 1024 * 1024;      // 10 MB per uploaded file

const SHEETS = {
  Activities: ['ID', 'Department', 'Title', 'Committee', 'Date', 'Week', 'Expenses', 'Amount (AED)', 'Venue',
               'Status', 'UAN', 'Attendance', 'Notes', 'Attachments', 'Created', 'Last updated', 'JSON (do not edit)'],
  News:       ['ID', 'Date', 'Title', 'Reference', 'For', 'Pinned', 'Message', 'Attachment', 'Last updated', 'JSON (do not edit)'],
  Files:      ['ID', 'Section', 'Title', 'File name', 'Description', 'Size (KB)', 'Uploaded', 'Link', 'JSON (do not edit)']
};

/* ---------------- helpers ---------------- */
function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function fail(msg) { throw new Error(msg); }
function sha256Hex(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}
function newId() { return Date.now().toString(36) + Utilities.getUuid().replace(/-/g, '').slice(0, 8); }
function fmtTime(ms) { return ms ? Utilities.formatDate(new Date(ms), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : ''; }
function str(v, max) { return String(v == null ? '' : v).slice(0, max || 2000); }

function sheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(SHEETS[name]);
    sh.getRange(1, 1, 1, SHEETS[name].length).setFontWeight('bold').setBackground('#DDF2E4');
    sh.setFrozenRows(1);
  }
  return sh;
}
function readAll(name) {
  const sh = sheet(name), n = sh.getLastRow(), cols = SHEETS[name].length;
  if (n < 2) return [];
  return sh.getRange(2, cols, n - 1, 1).getValues()
    .map(function (r) { try { return JSON.parse(r[0]); } catch (e) { return null; } })
    .filter(function (x) { return x && x.id; });
}
function findRow(name, id) {
  const sh = sheet(name), n = sh.getLastRow();
  if (n < 2) return -1;
  const ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}
function readOne(name, id) {
  const row = findRow(name, id);
  if (row < 0) return null;
  const cols = SHEETS[name].length;
  try { return JSON.parse(sheet(name).getRange(row, cols).getValue()); } catch (e) { return null; }
}
function upsert(name, rec, rowValues) {
  const sh = sheet(name), row = findRow(name, rec.id);
  const values = [rowValues.concat([JSON.stringify(rec)])];
  if (row > 0) sh.getRange(row, 1, 1, values[0].length).setValues(values);
  else sh.getRange(sh.getLastRow() + 1, 1, 1, values[0].length).setValues(values);
}
function removeRow(name, id) { const row = findRow(name, id); if (row > 0) sheet(name).deleteRow(row); }

/* ---------------- Drive ---------------- */
function rootFolder() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder(FOLDER_NAME);
  props.setProperty('FOLDER_ID', f.getId());
  return f;
}
function subFolder(path) {          // e.g. 'Activity attachments/Buklod'
  let f = rootFolder();
  String(path).split('/').forEach(function (name) {
    const it = f.getFoldersByName(name);
    f = it.hasNext() ? it.next() : f.createFolder(name);
  });
  return f;
}
function trashFile(fileId) { if (!fileId) return; try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {} }
function saveUpload(folderName, d) {
  const bytes = Utilities.base64Decode(String(d.b64 || ''));
  if (!bytes.length) fail('Empty file');
  if (bytes.length > MAX_FILE_BYTES) fail('File is larger than 10 MB');
  const name = str(d.name, 200) || 'file';
  const blob = Utilities.newBlob(bytes, str(d.type, 100) || 'application/octet-stream', name);
  const file = subFolder(folderName).createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const id = file.getId();
  return { fileId: id, name: name, type: blob.getContentType(), size: bytes.length,
           url: 'https://drive.google.com/uc?export=download&id=' + id };
}

/* ---------------- sessions ---------------- */
function startSession(acct) {
  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('tok_' + token, JSON.stringify({ id: acct, role: ACCOUNTS[acct].role }), SESSION_SECONDS);
  return token;
}
function session(token) {
  if (!token) fail('auth');
  const v = CacheService.getScriptCache().get('tok_' + token);
  if (!v) fail('auth');
  return JSON.parse(v);
}
function passwordRecord(acct) {
  const o = PropertiesService.getScriptProperties().getProperty('pw_' + acct);
  return o ? JSON.parse(o) : { salt: ACCOUNTS[acct].salt, hash: ACCOUNTS[acct].hash };
}
function checkPassword(acct, pass) {
  const p = passwordRecord(acct);
  return sha256Hex(p.salt + String(pass || '')) === p.hash;
}

/* ---------------- record cleaning (never trust the browser) ---------------- */
function cleanAttachment(a) {
  return { name: str(a.name, 200), type: str(a.type, 100), size: Number(a.size) || 0,
           fileId: str(a.fileId, 100), url: str(a.url, 300) };
}
function cleanActivity(r, dept, prev) {
  const status = STATUSES.indexOf(r.status) >= 0 ? r.status : 'planned';
  const rec = {
    id: prev ? prev.id : (str(r.id, 60) || newId()),
    dept: dept,
    title: str(r.title, 200).trim(),
    committee: COMMITTEES.indexOf(r.committee) >= 0 ? r.committee : fail('Choose a committee'),
    date: /^\d{4}-\d{2}-\d{2}$/.test(r.date) ? r.date : fail('Enter a valid date'),
    week: Math.min(5, Math.max(1, Number(r.week) || 1)),
    expense: r.expense === 'with' ? 'with' : 'without',
    amount: r.expense === 'with' ? str(r.amount, 20) : '',
    venue: str(r.venue, 200), status: status,
    customStatus: status === 'custom' ? str(r.customStatus, 120) : '',
    uan: str(r.uan, 60), attendance: str(r.attendance, 200), notes: str(r.notes, 3000),
    attachments: (Array.isArray(r.attachments) ? r.attachments : []).filter(function (a) { return a && a.fileId; }).map(cleanAttachment),
    createdAt: prev ? prev.createdAt : Date.now(), updatedAt: Date.now(), createdBy: dept
  };
  if (!rec.title) fail('Enter a title');
  return rec;
}
function activityRow(a) {
  return [a.id, DEPT_NAMES[a.dept] || a.dept, a.title, a.committee, a.date, 'Week ' + a.week,
          a.expense === 'with' ? 'With expenses' : 'Without expenses', a.amount || '', a.venue || '',
          a.status === 'custom' ? a.customStatus : STATUS_NAMES[a.status], a.uan || '', a.attendance || '', a.notes || '',
          (a.attachments || []).map(function (f) { return f.url; }).join('\n'), fmtTime(a.createdAt), fmtTime(a.updatedAt)];
}
function newsRow(n) {
  return [n.id, n.date, n.title, n.ref || '', n.audience || 'All', n.pinned ? 'Yes' : '', n.body,
          n.attachment ? n.attachment.url : '', fmtTime(n.updatedAt)];
}
function fileRow(f) {
  return [f.id, f.category === 'forms' ? 'Official forms & documents' : 'Logos & images', f.title || '', f.name,
          f.desc || '', Math.round((f.size || 0) / 1024), fmtTime(f.uploadedAt), f.url];
}

/* ---------------- web app ---------------- */
function doGet(e) {
  try {
    return json({ ok: true, activities: readAll('Activities'), news: readAll('News'), files: readAll('Files'), serverTime: Date.now() });
  } catch (err) { return json({ ok: false, error: String(err.message || err) }); }
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json({ ok: false, error: 'Bad request' }); }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    return json(handle(d));
  } catch (err) {
    return json({ ok: false, error: String(err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function handle(d) {
  const action = d.action;

  if (action === 'login') {
    const acct = String(d.acct || '');
    if (!ACCOUNTS[acct]) fail('Unknown account');
    const cache = CacheService.getScriptCache(), key = 'fail_' + acct;
    const fails = Number(cache.get(key) || 0);
    if (fails >= 10) fail('Too many wrong attempts. Try again in 10 minutes.');
    const okUser = String(d.user || '').trim().toLowerCase() === ACCOUNTS[acct].user.toLowerCase();
    if (!okUser || !checkPassword(acct, d.pass)) { cache.put(key, String(fails + 1), 600); fail('Incorrect username or password.'); }
    cache.remove(key);
    return { ok: true, token: startSession(acct), id: acct, role: ACCOUNTS[acct].role };
  }

  const s = session(d.token);   // everything below needs a login
  const isAdmin = s.role === 'admin', isDept = s.role === 'dept';

  switch (action) {
    case 'whoami':
      return { ok: true, id: s.id, role: s.role };

    case 'logout':
      CacheService.getScriptCache().remove('tok_' + d.token);
      return { ok: true };

    case 'changePassword': {
      if (!checkPassword(s.id, d.oldPass)) fail('Current password is incorrect');
      if (String(d.newPass || '').length < 6) fail('New password must be at least 6 characters');
      const salt = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
      PropertiesService.getScriptProperties().setProperty('pw_' + s.id, JSON.stringify({ salt: salt, hash: sha256Hex(salt + d.newPass) }));
      return { ok: true };
    }

    case 'upload': {
      const kind = d.kind;
      if (kind === 'attachment') {
        if (!isDept) fail('Only departments can attach files to activities');
        return { ok: true, file: saveUpload('Activity attachments/' + DEPT_NAMES[s.id], d) };
      }
      if (kind === 'memo') {
        if (!isAdmin) fail('Admin only');
        return { ok: true, file: saveUpload('Memos', d) };
      }
      if (kind === 'forms' || kind === 'logos') {
        if (!isAdmin) fail('Admin only');
        const f = saveUpload(kind === 'forms' ? 'Official forms and documents' : 'Logos and images', d);
        const rec = { id: newId(), category: kind, name: f.name, title: str(d.title, 200), desc: str(d.desc, 300),
                      type: f.type, size: f.size, fileId: f.fileId, url: f.url, uploadedAt: Date.now(), updatedAt: Date.now() };
        upsert('Files', rec, fileRow(rec));
        return { ok: true, record: rec };
      }
      fail('Unknown upload type');
    }

    case 'saveActivity': {
      if (!isDept) fail('Only departments can add activities');
      const r = d.record || {};
      const prev = r.id ? readOne('Activities', r.id) : null;
      if (prev && prev.dept !== s.id) fail('You can only edit your own department\'s activities');
      const rec = cleanActivity(r, s.id, prev);
      if (prev) {   // remove Drive files that were taken off the activity
        const keep = rec.attachments.map(function (a) { return a.fileId; });
        (prev.attachments || []).forEach(function (a) { if (keep.indexOf(a.fileId) < 0) trashFile(a.fileId); });
      }
      upsert('Activities', rec, activityRow(rec));
      return { ok: true, record: rec };
    }

    case 'deleteActivity': {
      if (!isDept) fail('Not allowed');
      const prev = readOne('Activities', d.id);
      if (!prev) return { ok: true };
      if (prev.dept !== s.id) fail('You can only delete your own department\'s activities');
      (prev.attachments || []).forEach(function (a) { trashFile(a.fileId); });
      removeRow('Activities', d.id);
      return { ok: true };
    }

    case 'saveNews': {
      if (!isAdmin) fail('Admin only');
      const r = d.record || {};
      const prev = r.id ? readOne('News', r.id) : null;
      const rec = {
        id: prev ? prev.id : newId(), title: str(r.title, 200).trim(),
        date: /^\d{4}-\d{2}-\d{2}$/.test(r.date) ? r.date : fail('Enter a valid date'),
        ref: str(r.ref, 100), audience: str(r.audience, 40), body: str(r.body, 10000), pinned: !!r.pinned,
        attachment: r.attachment && r.attachment.fileId ? cleanAttachment(r.attachment) : null,
        createdAt: prev ? prev.createdAt : Date.now(), updatedAt: Date.now()
      };
      if (!rec.title) fail('Enter a title');
      if (prev && prev.attachment && (!rec.attachment || rec.attachment.fileId !== prev.attachment.fileId)) trashFile(prev.attachment.fileId);
      upsert('News', rec, newsRow(rec));
      return { ok: true, record: rec };
    }

    case 'deleteNews': {
      if (!isAdmin) fail('Admin only');
      const prev = readOne('News', d.id);
      if (prev && prev.attachment) trashFile(prev.attachment.fileId);
      removeRow('News', d.id);
      return { ok: true };
    }

    case 'deleteFile': {
      if (!isAdmin) fail('Admin only');
      const prev = readOne('Files', d.id);
      if (prev) trashFile(prev.fileId);
      removeRow('Files', d.id);
      return { ok: true };
    }
  }
  fail('Unknown action');
}

/* Run once from the editor: creates the tabs and the Drive folder. */
function setup() {
  Object.keys(SHEETS).forEach(sheet);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const first = ss.getSheetByName('Sheet1');
  if (first && ss.getSheets().length > 1 && first.getLastRow() === 0) ss.deleteSheet(first);
  const f = rootFolder();
  Logger.log('Ready. Files will be saved in Drive folder: ' + f.getUrl());
}
