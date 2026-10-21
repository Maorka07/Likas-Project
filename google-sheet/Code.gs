/**
 * Lika's Project — online storage in Google Sheets
 * ------------------------------------------------
 * Paste this whole file into: the Google Sheet → Extensions → Apps Script.
 * Then: Deploy → New deployment → type "Web app"
 *       Execute as: Me   ·   Who has access: Anyone
 * Copy the Web app URL and send it to Claude (or paste it in the site's settings).
 *
 * Tabs it creates:
 *   "פניות"   — notes, 👍/👎 flags and ⚠️ error reports from Lika
 *   "יומן"    — every practice session (type, level, score, minutes)
 *   "סיכום"   — current status (XP, streak, words learned...)
 *   "הודעה"   — write a message in cell A2 and it appears on her home screen
 *   "_state"  — the app's saved progress (do not edit)
 */

// ⬇️ IMPORTANT: replace XXXX with Lika's PIN (keep the quotes). Requests without it are rejected.
const PIN = 'XXXX';

const TZ = 'Asia/Jerusalem';
const CHUNK = 40000; // max characters per cell is 50,000

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet_(name, header) {
  const ss = ss_();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    if (header) {
      s.appendRow(header);
      s.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#DBEAFE');
      s.setFrozenRows(1);
    }
    s.setRightToLeft(true);
  }
  return s;
}

function setup_() {
  sheet_('פניות', ['זמן', 'סוג', 'פירוט', 'שאלה', 'תשובה נכונה', 'תרגול', 'רמה', 'נושא']);
  sheet_('יומן', ['זמן', 'תרגול', 'רמה', 'נכונות', 'אחוז', 'דקות', 'ציון משוער']);
  sheet_('סיכום', ['נתון', 'ערך']);
  const m = sheet_('הודעה');
  if (!m.getRange('A1').getValue()) {
    m.getRange('A1').setValue('ההודעה בתא A2 תופיע לליקה במסך הבית 💌').setFontWeight('bold');
    m.getRange('A2').setValue('חיים שלי זה קטן עלייך, מאמין בך ואוהב אותך! ❤️');
    m.setColumnWidth(1, 520);
  }
  const st = sheet_('_state');
  st.hideSheet();
}

function now_() { return Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm'); }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function message_() { return String(sheet_('הודעה').getRange('A2').getValue() || ''); }

function loadState_() {
  const st = sheet_('_state');
  const n = st.getLastRow();
  if (!n) return null;
  const parts = st.getRange(1, 1, n, 1).getValues().map(r => r[0]).join('');
  try { return JSON.parse(parts); } catch (e) { return null; }
}

function saveState_(state) {
  const st = sheet_('_state');
  const txt = JSON.stringify(state);
  const rows = [];
  for (let i = 0; i < txt.length; i += CHUNK) rows.push([txt.slice(i, i + CHUNK)]);
  st.clear();
  if (rows.length) st.getRange(1, 1, rows.length, 1).setNumberFormat('@').setValues(rows);
}

/* GET: returns the saved progress + the message for Lika */
function doGet(e) {
  setup_();
  if (String((e.parameter || {}).pin || '') !== PIN) return json_({ ok: false, error: 'bad pin' });
  return json_({ ok: true, state: loadState_(), message: message_() });
}

/* POST (text/plain JSON): { state, events:[...], summary:{...} } */
function doPost(e) {
  setup_();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    if (String(body.pin || '') !== PIN) return json_({ ok: false, error: 'bad pin' });
    if (body.state) {
      const cur = loadState_();
      // last write wins — only overwrite with a newer copy
      if (!cur || !cur.upd || (body.state.upd || 0) >= cur.upd) saveState_(body.state);
    }
    (body.events || []).forEach(ev => {
      const t = ev.time ? Utilities.formatDate(new Date(ev.time), TZ, 'dd/MM/yyyy HH:mm') : now_();
      if (ev.kind === 'session') {
        sheet_('יומן').appendRow([t, ev.type || '', ev.lv || '', (ev.c || 0) + '/' + (ev.n || 0),
          ev.n ? Math.round(ev.c / ev.n * 100) + '%' : '', ev.mins || '', ev.est || '']);
      } else {
        const label = { note: '💌 הערה', more: '👍 עוד כאלה', less: '👎 פחות כאלה', error: '⚠️ טעות / לא ברור' }[ev.kind] || ev.kind;
        sheet_('פניות').appendRow([t, label, ev.text || '', ev.q || '', ev.a || '', ev.type || '', ev.lv || '', ev.topic || '']);
      }
    });
    if (body.summary) {
      const s = sheet_('סיכום');
      const rows = Object.keys(body.summary).map(k => [k, body.summary[k]]);
      rows.unshift(['עודכן לאחרונה', now_()]);
      s.getRange(2, 1, Math.max(s.getLastRow(), 2), 2).clearContent();
      s.getRange(2, 1, rows.length, 2).setValues(rows);
    }
    return json_({ ok: true, message: message_() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
