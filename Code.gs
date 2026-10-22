/**
 * QRC Class of '06 Reunion Lime - signup backend
 *
 * Paste this into Extensions > Apps Script from inside the Google Sheet
 * you want to use as the database. See the setup steps in the chat.
 *
 * - POST  : appends one signup row
 * - GET   : returns public totals only (no names)
 */

// ---- Settings (keep PRICES in sync with index.html) -----------------------
const SHEET_NAME = 'Signups';
const PRICES = { lime: 100, jersey: 200 };
const PAID_LAST_CREDIT = 100; // off the total for "I paid for the last lime"
const PAID_LAST_NOTE = 'Total decreased for previously paid amount';
const MAX_CONTRIBUTION = 50000;
// false = the public totals count every pledge.
// true  = only rows where you type Y / Yes / Paid in the "Paid" column count.
const COUNT_ONLY_PAID = false;
// ---------------------------------------------------------------------------

const HEADERS = [
  'Timestamp', 'ID', 'Name', 'Contact',
  'Lime', 'Jersey', 'Jersey size', 'Contribution',
  'Amount due', 'Paid', 'Notes'
];
const COL = { LIME: 5, JERSEY: 6, CONTRIB: 8, PAID: 10 };
const SIZES = ['S', 'M', 'L', 'XL', '2XL', '3XL'];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function computeTotals_() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  const totals = { attending: 0, jerseys: 0, contributed: 0, updated: new Date().toISOString() };
  if (last < 2) return totals;

  const rows = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  rows.forEach(function (r) {
    if (COUNT_ONLY_PAID) {
      const paid = String(r[COL.PAID - 1]).trim().toLowerCase();
      if (['y', 'yes', 'paid'].indexOf(paid) === -1) return;
    }
    if (r[COL.LIME - 1] === 'Yes') totals.attending++;
    if (r[COL.JERSEY - 1] === 'Yes') totals.jerseys++;
    totals.contributed += Number(r[COL.CONTRIB - 1]) || 0;
  });
  return totals;
}

function getTotals_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('totals');
  if (hit) return JSON.parse(hit);
  const t = computeTotals_();
  cache.put('totals', JSON.stringify(t), 30); // seconds
  return t;
}

function doGet(e) {
  return json_({ ok: true, totals: getTotals_() });
}

// Stops spreadsheet formulas from being injected through the form.
function clean_(value, max) {
  let s = String(value == null ? '' : value).replace(/[\r\n\t]+/g, ' ').trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);

    // Honeypot: real people never fill this in. Pretend it worked.
    if (d.website) return json_({ ok: true, totals: getTotals_() });

    const name = clean_(d.name, 100);
    const contact = clean_(d.contact, 120);
    const id = clean_(d.id, 60);
    const lime = d.lime === true;
    const jersey = d.jersey === true;
    const paidLast = lime && d.paidLast === true;
    const size = jersey ? clean_(d.size, 6) : '';
    let contribution = d.contribute === true ? Math.round(Number(d.contribution) * 100) / 100 : 0;

    if (!name) return json_({ ok: false, error: 'Please enter your name.' });
    if (!contact) return json_({ ok: false, error: 'Please enter an email or phone number.' });
    if (!id) return json_({ ok: false, error: 'Missing submission id.' });
    if (!lime && !jersey && !(d.contribute === true)) {
      return json_({ ok: false, error: 'Pick at least one option.' });
    }
    if (jersey && SIZES.indexOf(size) === -1) {
      return json_({ ok: false, error: 'Please choose a jersey size.' });
    }
    if (d.contribute === true) {
      if (!(contribution > 0) || contribution > MAX_CONTRIBUTION) {
        return json_({ ok: false, error: 'Enter a contribution amount greater than 0.' });
      }
    } else {
      contribution = 0;
    }

    const due = (lime ? PRICES.lime : 0) + (jersey ? PRICES.jersey : 0) + contribution
      - (paidLast ? PAID_LAST_CREDIT : 0);

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const sh = getSheet_();
      const last = sh.getLastRow();

      // Same submission sent twice (double tap, retry): ignore the repeat.
      if (last >= 2) {
        const dup = sh.getRange(2, 2, last - 1, 1)
          .createTextFinder(id).matchEntireCell(true).findNext();
        if (dup) return json_({ ok: true, totals: computeTotals_() });
      }

      sh.appendRow([
        new Date(), id, name, contact,
        lime ? 'Yes' : '', jersey ? 'Yes' : '', size, contribution || '',
        due, '', paidLast ? PAID_LAST_NOTE : ''
      ]);
    } finally {
      lock.releaseLock();
    }

    CacheService.getScriptCache().remove('totals');
    return json_({ ok: true, totals: computeTotals_() });
  } catch (err) {
    return json_({ ok: false, error: 'Something went wrong. Please try again.' });
  }
}
