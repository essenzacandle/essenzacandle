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

// La scheda "Scorte" senza badare a maiuscole e spazi; se non c'è, la prima scheda del file.
function stockSheet(book) {
  const sheets = book.getSheets();
  return sheets.find(s => s.getName().trim().toLowerCase() === SHEET.toLowerCase()) || sheets[0];
}

function doGet() {
  return json(readStock());
}

function doPost(e) {
  const id = String((e.postData && e.postData.contents) || '').trim();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const book = SpreadsheetApp.getActive();
    const sheet = stockSheet(book);
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
  const rows = stockSheet(SpreadsheetApp.getActive()).getDataRange().getValues();
  const out = {};
  rows.slice(1).forEach(r => {
    const id = String(r[0]).trim();
    // Righe con totale o rimasti vuoti = prodotto senza limite di pezzi: il sito le ignora.
    if (id && r[1] !== '' && r[2] !== '') out[id] = { total: Number(r[1]), left: Math.max(0, Number(r[2])) };
  });
  return out;
}

function json(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
