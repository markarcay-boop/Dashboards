/* KO COACH Lead Spin – Google Sheet backend (version 2)
   - "Leads" tab: one row per sign-up (unchanged from version 1)
   - "Gifts" tab: the shared "Up for grabs" list and stock, used by every device
   Update: paste over the old code, Save, then Deploy > Manage deployments > Edit (pencil)
   > Version: New version > Deploy. The web app URL stays the same. */

const SHEET_NAME = 'Leads';
const GIFTS_SHEET = 'Gifts';
const HEADERS = ['Received','Lead ID','Date/time','Name','WhatsApp','Email',
                 'City','Gym/Club','Profile','Interest','Prize','Device'];
const GIFT_HEADERS = ['ID','Gift','Quantity left','Given','Color'];

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
                       .setMimeType(ContentService.MimeType.JSON);
}

function giftSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(GIFTS_SHEET);
  if (!sh) { sh = ss.insertSheet(GIFTS_SHEET); sh.appendRow(GIFT_HEADERS); }
  return sh;
}

function readGifts() {
  const sh = giftSheet(), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 5).getValues()
    .filter(r => String(r[1]).trim() !== '')
    .map(r => ({ id: String(r[0] || ''), name: String(r[1]), qty: Number(r[2]) || 0,
                 given: Number(r[3]) || 0, color: String(r[4] || '') }));
}

function writeGifts(list) {
  const sh = giftSheet(), n = sh.getLastRow();
  if (n > 1) sh.getRange(2, 1, n - 1, 5).clearContent();
  if (list.length) sh.getRange(2, 1, list.length, 5).setValues(
    list.map(g => [g.id || '', g.name || '', Number(g.qty) || 0, Number(g.given) || 0, g.color || '']));
}

function doGet(e) {
  return json({ ok: true, gifts: readGifts() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();   // queues simultaneous requests
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);

    if (d.type === 'gifts') {                   // admin saved the gift list
      writeGifts(Array.isArray(d.gifts) ? d.gifts : []);
      return json({ ok: true });
    }

    // a new lead (same as version 1)
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
    sh.appendRow([new Date(), d.id||'', d.time||'', d.name||'', d.mobile||'',
                  d.email||'', d.city||'', d.club||'', d.profile||'',
                  d.interest||'', d.prize||'', d.device||'']);

    // take the won gift off the shared stock
    if (d.prize) {
      const g = giftSheet(), n = g.getLastRow();
      if (n > 1) {
        const v = g.getRange(2, 1, n - 1, 5).getValues();
        for (let i = 0; i < v.length; i++) {
          if (String(v[i][1]) === String(d.prize)) {
            g.getRange(i + 2, 3, 1, 2).setValues([[Math.max(0, (Number(v[i][2]) || 0) - 1),
                                                   (Number(v[i][3]) || 0) + 1]]);
            break;
          }
        }
      }
    }
    return json({ ok: true });
  } finally {
    lock.releaseLock();
  }
}
