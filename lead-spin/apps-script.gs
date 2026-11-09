/* KO COACH Lead Spin – Google Sheet backend (version 3)
   One Google Sheet is the single source of truth for EVERY device and browser:
   - "Settings" tab : theme, business name, event name, wheel rules, admin password (hashed)
   - "Gifts" tab    : the gift list and remaining stock
   - "Leads" tab    : one row per sign-up
   The server also does the duplicate check and draws the prize, so stock is never oversold
   and the same WhatsApp number / email cannot enter twice from different phones.

   UPDATE STEPS (keeps the same web-app URL):
   1. Open the Google Sheet > Extensions > Apps Script
   2. Select all the old code, paste this file over it, click Save
   3. Deploy > Manage deployments > Edit (pencil) > Version: "New version" > Deploy
      (Execute as: Me · Who has access: Anyone)                                            */

const LEADS_SHEET = 'Leads', GIFTS_SHEET = 'Gifts', SETTINGS_SHEET = 'Settings';
const HEADERS = ['Received','Lead ID','Date/time','Name','WhatsApp','Email',
                 'City','Gym/Club','Profile','Interest','Prize','Device','Consent'];
const GIFT_HEADERS = ['ID','Gift','Quantity left','Given','Color'];
const DEFAULT_PASS_HASH = '725506babd15cf507ac203ef34f3e25e2027dbefe999115dd4298935d97c1595'; // "0000"
const SHARED_KEYS = ['theme','brand','eventName','resetSec','weighted','consol','consolLabel','consolPct'];

/* ---------- helpers ---------- */
function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet(name, headers) {
  let sh = ss().getSheetByName(name);
  if (!sh) { sh = ss().insertSheet(name); if (headers) sh.appendRow(headers); }
  return sh;
}
function normPhone(p) { return String(p || '').replace(/[^\d]/g, '').replace(/^00/, '').replace(/^971/, '').replace(/^0/, ''); }
function normEmail(e) { return String(e || '').trim().toLowerCase(); }

/* ---------- settings (key / value rows) ---------- */
function readSettingsRaw() {
  const sh = sheet(SETTINGS_SHEET, ['Key', 'Value']), n = sh.getLastRow(), o = {};
  if (n > 1) sh.getRange(2, 1, n - 1, 2).getValues().forEach(r => {
    if (r[0] !== '') { try { o[r[0]] = JSON.parse(r[1]); } catch (e) { o[r[0]] = r[1]; } }
  });
  return o;
}
function writeSettingsRaw(o) {
  const sh = sheet(SETTINGS_SHEET, ['Key', 'Value']), n = sh.getLastRow();
  if (n > 1) sh.getRange(2, 1, n - 1, 2).clearContent();
  const rows = Object.keys(o).map(k => [k, JSON.stringify(o[k])]);
  if (rows.length) sh.getRange(2, 1, rows.length, 2).setValues(rows);
}
function publicSettings(raw) {
  const o = {};
  SHARED_KEYS.forEach(k => { if (raw[k] !== undefined) o[k] = raw[k]; });
  return o;
}
function passHash(raw) { return raw.passHash || DEFAULT_PASS_HASH; }
function authed(raw, h) { return !!h && h === passHash(raw); }

/* ---------- gifts ---------- */
function readGifts() {
  const sh = sheet(GIFTS_SHEET, GIFT_HEADERS), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 5).getValues()
    .filter(r => String(r[1]).trim() !== '')
    .map(r => ({ id: String(r[0] || ''), name: String(r[1]), qty: Math.max(0, Number(r[2]) || 0),
                 given: Math.max(0, Number(r[3]) || 0), color: String(r[4] || '') }));
}
function writeGifts(list) {
  const sh = sheet(GIFTS_SHEET, GIFT_HEADERS), n = sh.getLastRow();
  if (n > 1) sh.getRange(2, 1, n - 1, 5).clearContent();
  if (list.length) sh.getRange(2, 1, list.length, 5).setValues(
    list.map(g => [g.id || '', g.name || '', Math.max(0, Number(g.qty) || 0), Math.max(0, Number(g.given) || 0), g.color || '']));
}

/* ---------- leads ---------- */
function leadsSheet() {
  const sh = sheet(LEADS_SHEET, HEADERS);
  if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
  if (sh.getLastColumn() < HEADERS.length) sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]); // adds "Consent"
  return sh;
}
function readLeads() {
  const sh = leadsSheet(), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, HEADERS.length).getValues().map((r, i) => ({
    row: i + 2, ts: r[0] instanceof Date ? r[0].getTime() : (Date.parse(r[0]) || 0), id: String(r[1] || ''),
    name: String(r[3] || ''), mobile: String(r[4] || '').replace(/^'/, ''), email: String(r[5] || ''), city: String(r[6] || ''),
    gym: String(r[7] || ''), profile: String(r[8] || ''), interest: String(r[9] || ''), prize: String(r[10] || ''),
    device: String(r[11] || ''), consent: String(r[12] || '') !== 'No' }));
}
function findDup(leads, mobile, email) {
  const ph = normPhone(mobile), em = normEmail(email);
  for (const l of leads) {
    if (ph.length >= 7 && normPhone(l.mobile) === ph) return { field: 'mobile', ts: l.ts };
    if (em.indexOf('@') > 0 && normEmail(l.email) === em) return { field: 'email', ts: l.ts };
  }
  return null;
}
function appendLead(d, prize) {
  leadsSheet().appendRow([new Date(), d.id || '', d.time || '', d.name || '', "'" + (d.mobile || ''),
    d.email || '', d.city || '', d.gym || d.club || '', d.profile || '', d.interest || '', prize || d.prize || '',
    d.device || '', d.consent === false ? 'No' : 'Yes']);
}

/* ---------- prize draw (same rules as the page) ---------- */
function drawPrize(gifts, s) {
  const pool = gifts.filter(g => g.qty > 0);
  if (!pool.length) return null;
  const items = pool.map(g => ({ key: g.id, w: s.weighted === false ? 1 : g.qty }));
  if (s.consol) {
    const pct = Math.min(90, Math.max(1, Number(s.consolPct) || 15)) / 100;
    const sum = items.reduce((a, b) => a + b.w, 0);
    items.push({ key: '__consol', w: sum * pct / (1 - pct) });
  }
  const total = items.reduce((a, b) => a + b.w, 0);
  let r = Math.random() * total;
  for (const it of items) { if ((r -= it.w) <= 0) return it.key; }
  return items[items.length - 1].key;
}

function state(raw) { return { ok: true, v: 3, settings: publicSettings(raw), gifts: readGifts(), passSet: !!raw.passHash }; }

/* ---------- GET ---------- */
function doGet(e) {
  const p = (e && e.parameter) || {}, raw = readSettingsRaw();
  if (p.action === 'leads') {
    if (!authed(raw, p.auth)) return json({ ok: false, error: 'auth', passSet: !!raw.passHash });
    return json({ ok: true, v: 3, leads: readLeads() });
  }
  if (p.action === 'auth') return json({ ok: authed(raw, p.auth), passSet: !!raw.passHash });
  if (p.action === 'check') return json({ ok: true, dup: findDup(readLeads(), p.mobile, p.email) });
  return json(state(raw));                                   // default: settings + gifts
}

/* ---------- POST ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock();                  // one request at a time
  lock.waitLock(25000);
  try {
    const d = JSON.parse(e.postData.contents), raw = readSettingsRaw();

    /* visitor registers: duplicate check + prize draw happen here */
    if (d.type === 'register') {
      const lead = d.lead || {};
      const dup = findDup(readLeads(), lead.mobile, lead.email);
      if (dup) return json({ ok: false, dup: dup });
      const gifts = readGifts(), before = JSON.parse(JSON.stringify(gifts));
      const key = drawPrize(gifts, raw);
      let prize = '', prizeName = '';
      if (key === '__consol') { prizeName = raw.consolLabel || 'Thank You!'; prize = '— (' + prizeName + ')'; }
      else if (key) {
        const g = gifts.find(x => x.id === key); g.qty -= 1; g.given += 1;
        writeGifts(gifts); prize = prizeName = g.name;
      } else prize = 'No gifts left';
      appendLead(lead, prize);
      return json({ ok: true, key: key, prize: prize, prizeName: prizeName, gifts: before, settings: publicSettings(raw) });
    }

    /* lead recorded while offline (or imported) — skip if already in the sheet */
    if (d.type === 'lead' || (!d.type && d.name)) {
      const leads = readLeads();
      if (d.id && leads.some(l => l.id === d.id)) return json({ ok: true, skipped: true });
      appendLead(d, d.prize);
      if (d.prize) {
        const gifts = readGifts(), g = gifts.find(x => x.name === d.prize);
        if (g) { g.qty = Math.max(0, g.qty - 1); g.given += 1; writeGifts(gifts); }
      }
      return json({ ok: true });
    }

    /* first-time password claim (only while no password has been saved to the sheet) */
    if (d.type === 'claimPass') {
      if (raw.passHash) return json({ ok: false, error: 'set' });
      if (!d.newPassHash) return json({ ok: false });
      raw.passHash = d.newPassHash; writeSettingsRaw(raw);
      return json({ ok: true });
    }

    /* everything below needs the admin password */
    if (!authed(raw, d.auth)) return json({ ok: false, error: 'auth' });

    if (d.type === 'save') {
      if (d.settings) SHARED_KEYS.forEach(k => { if (d.settings[k] !== undefined) raw[k] = d.settings[k]; });
      if (d.newPassHash) raw.passHash = d.newPassHash;
      writeSettingsRaw(raw);
      if (Array.isArray(d.gifts)) {
        // stock edits are applied as a change (+/-) so spins made meanwhile are not lost
        const cur = readGifts(), byId = {};
        cur.forEach(g => byId[g.id] = g);
        const out = d.gifts.filter(g => g && String(g.name || '').trim() !== '').map(g => {
          const c = byId[g.id];
          if (!c) return { id: g.id, name: g.name, qty: Number(g.qty) || 0, given: Number(g.given) || 0, color: g.color };
          const delta = (Number(g.qty) || 0) - (g.baseQty === undefined ? c.qty : Number(g.baseQty) || 0);
          return { id: g.id, name: g.name, qty: Math.max(0, c.qty + delta), given: c.given, color: g.color };
        });
        writeGifts(out);
      }
      return json(state(readSettingsRaw()));
    }
    if (d.type === 'resetGiven') {
      writeGifts(readGifts().map(g => (g.given = 0, g)));
      return json(state(raw));
    }
    if (d.type === 'deleteLead') {
      const l = readLeads().find(x => x.id === d.id);
      if (l) leadsSheet().deleteRow(l.row);
      return json({ ok: true });
    }
    if (d.type === 'clearLeads') {
      const sh = leadsSheet(), n = sh.getLastRow();
      if (n > 1) sh.deleteRows(2, n - 1);
      return json({ ok: true });
    }
    return json({ ok: false, error: 'unknown' });
  } finally {
    lock.releaseLock();
  }
}
