/* Scorte di Essenza su Google Sheet.
 *
 * Foglio "Scorte" (riga 1 = intestazioni):
 *   id                     | totale | rimasti
 *   mystery-box-natalizia  | 24     | 22
 *   gnomi-natalizi         | 24     | 23
 *   pallina-natalizia      | 20     | 18
 *
 * GET  → restituisce le scorte di tutti i prodotti in JSON.
 * POST → corpo = id del prodotto: scala "rimasti" di 1 e annota l'ordine nel foglio "Ordini".
 */
const SHEET = 'Scorte';
const LOG = 'Ordini';

function doGet() {
  return json(readStock());
}

function doPost(e) {
  const id = String((e.postData && e.postData.contents) || '').trim();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const book = SpreadsheetApp.getActive();
    const sheet = book.getSheetByName(SHEET);
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() !== id) continue;
      const left = Number(rows[i][2]);
      if (left > 0) {
        sheet.getRange(i + 1, 3).setValue(left - 1);
        const log = book.getSheetByName(LOG) || book.insertSheet(LOG);
        log.appendRow([new Date(), id, left - 1]);
      }
      break;
    }
  } finally {
    lock.releaseLock();
  }
  return json(readStock());
}

function readStock() {
  const rows = SpreadsheetApp.getActive().getSheetByName(SHEET).getDataRange().getValues();
  const out = {};
  rows.slice(1).forEach(r => {
    const id = String(r[0]).trim();
    if (id) out[id] = { total: Number(r[1]), left: Math.max(0, Number(r[2])) };
  });
  return out;
}

function json(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
