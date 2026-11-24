/**
 * Health Checkup Book / สมุดผลตรวจสุขภาพ
 * Google Apps Script web app that reads and writes the "Health Checkup Book" Google Sheet.
 * Sheet tabs:
 *   Results — one row per test; columns A–I describe the test, then one column per year (header = 4-digit year)
 *   Years   — Year | Date (Thai) | Date (English) | Source
 */
const SHEET_ID = '1lIh4jAl_q8_oKxXboYNnJH7N1BrwC1R3fN-SahBguUY';
const RESULTS = 'Results';
const YEARS = 'Years';
const FIRST_YEAR_COL = 10; // column J

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Health Checkup Book')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function book_() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function getData() {
  const ss = book_();
  return {
    results: ss.getSheetByName(RESULTS).getDataRange().getDisplayValues(),
    years: ss.getSheetByName(YEARS).getDataRange().getDisplayValues(),
    url: ss.getUrl()
  };
}

/** o = {k:'2026', dth:'15 พ.ย. 69', den:'15 Nov 2026', v:{ldl:150, bp:'115/75', ...}} */
function addYear(o) {
  const k = String(o && o.k || '').trim();
  if (!/^\d{4}$/.test(k)) throw new Error('Year must be 4 digits');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = book_();
    const rs = ss.getSheetByName(RESULTS);
    const ys = ss.getSheetByName(YEARS);
    const head = rs.getRange(1, 1, 1, rs.getLastColumn()).getDisplayValues()[0];
    if (head.some(function (h) { return String(h).trim() === k; })) throw new Error('Year ' + k + ' already exists');

    // keep year columns in order: insert before the first later year, else append
    let col = rs.getLastColumn() + 1;
    for (let i = FIRST_YEAR_COL - 1; i < head.length; i++) {
      if (/^\d{4}$/.test(String(head[i]).trim()) && String(head[i]).trim() > k) { col = i + 1; break; }
    }
    if (col <= rs.getLastColumn()) rs.insertColumnBefore(col);

    const ids = rs.getRange(2, 1, Math.max(rs.getLastRow() - 1, 1), 1).getDisplayValues().map(function (r) { return String(r[0]).trim(); });
    const vals = ids.map(function (id) {
      const x = o.v && Object.prototype.hasOwnProperty.call(o.v, id) ? o.v[id] : '';
      if (x === '' || x == null) return [''];
      return [typeof x === 'number' ? x : "'" + String(x)];
    });
    rs.getRange(1, col).setValue("'" + k).setFontWeight('bold');
    rs.getRange(2, col, vals.length, 1).setValues(vals).setHorizontalAlignment('right');

    ys.appendRow(["'" + k, o.dth || k, o.den || k, 'Added from the app']);
    const n = ys.getLastRow() - 1;
    if (n > 1) ys.getRange(2, 1, n, ys.getLastColumn()).sort(1);
  } finally {
    lock.releaseLock();
  }
  return getData();
}

function deleteYear(k) {
  k = String(k || '').trim();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = book_();
    const rs = ss.getSheetByName(RESULTS);
    const ys = ss.getSheetByName(YEARS);
    const head = rs.getRange(1, 1, 1, rs.getLastColumn()).getDisplayValues()[0];
    const i = head.findIndex(function (h, j) { return j >= FIRST_YEAR_COL - 1 && String(h).trim() === k; });
    if (i >= 0) rs.deleteColumn(i + 1);
    const yv = ys.getDataRange().getDisplayValues();
    for (let r = yv.length - 1; r >= 1; r--) if (String(yv[r][0]).trim() === k) ys.deleteRow(r + 1);
  } finally {
    lock.releaseLock();
  }
  return getData();
}
