// ================================================================
// MagMa — Backend Google Apps Script
// ================================================================
//
// ─── ISTRUZIONI DI PUBBLICAZIONE ────────────────────────────────
//
//  1. Vai su https://script.google.com e crea un nuovo progetto.
//
//  2. Incolla questo codice nel file "Code.gs" (sostituendo tutto).
//
//  3. Sostituisci il valore di SPREADSHEET_ID qui sotto con l'ID
//     del tuo Google Sheet.
//     Come trovarlo: apri il foglio nel browser, l'URL sarà:
//     https://docs.google.com/spreadsheets/d/[ID-QUI]/edit
//     Copia la stringa tra /d/ e /edit.
//
//  4. Assicurati che il foglio si chiami "Magazzino" (oppure
//     cambia la costante SHEET_NAME).
//     La riga 1 deve essere l'intestazione:
//     A: Numero scatola | B: Codice Articolo | C: Codice a barre
//     D: Marca | E: Nome Prodotto | F: Colore | G: Taglia
//     H: Quantità | I: Scadenza | J: Prezzo | K: Data controllo
//
//  5. Clic su "Distribuisci" → "Nuova distribuzione".
//
//  6. Impostazioni:
//       - Tipo:          App web
//       - Esegui come:   Me (il tuo account Google)
//       - Chi ha accesso: Chiunque  ← IMPORTANTE per l'app
//
//  7. Clic su "Distribuisci". Copia l'URL "Web app" che appare.
//
//  8. Incolla quell'URL nella costante API_URL in app.js.
//
//  9. Al primo accesso autorizza l'app (richiesta una tantum).
//
// ─── NOTA IMPORTANTE SUI PERMESSI ──────────────────────────────
//  Se imposti "Chiunque, inclusi gli utenti anonimi" chiunque
//  conosca l'URL può scrivere nel foglio. Per maggiore sicurezza
//  usa "Chiunque" (richiede account Google) o aggiungi un campo
//  "token" segreto nel payload e verificalo qui.
//
// ================================================================

const SPREADSHEET_ID = 'IL_TUO_SPREADSHEET_ID_QUI';
const SHEET_NAME     = 'Magazzino';

// Indici colonne (0-based, per getValues())
const COL = {
  BOX:      0,   // A: Numero scatola
  ARTICLE:  1,   // B: Codice Articolo
  BARCODE:  2,   // C: Codice a barre
  BRAND:    3,   // D: Marca
  NAME:     4,   // E: Nome Prodotto
  COLOR:    5,   // F: Colore
  SIZE:     6,   // G: Taglia
  QUANTITY: 7,   // H: Quantità
  EXPIRY:   8,   // I: Scadenza
  PRICE:    9,   // J: Prezzo
  DATE:     10,  // K: Data controllo
};

// ================================================================
// PUNTO DI INGRESSO — riceve le richieste POST dall'app
// ================================================================
function doPost(e) {
  try {
    const data   = JSON.parse(e.postData.contents);
    const action = (data.action || '').toLowerCase().trim();

    let result;
    switch (action) {
      case 'search': result = searchProduct(data.barcode);          break;
      case 'add':    result = addUnit(data.barcode, data.expiry, data.date);  break;
      case 'remove': result = removeUnit(data.barcode, data.date);  break;
      case 'create': result = createProduct(data);                  break;
      default:
        result = { success: false, message: 'Azione sconosciuta: ' + action };
    }

    return buildResponse(result);

  } catch (err) {
    return buildResponse({
      success: false,
      message: 'Errore interno del server: ' + err.message,
    });
  }
}

// Costruisce risposta JSON con MIME type corretto
function buildResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ================================================================
// HELPERS
// ================================================================

// Restituisce il foglio di lavoro configurato
function getSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Foglio "' + SHEET_NAME + '" non trovato.');
  return sheet;
}

// Cerca la riga per codice a barre (colonna C).
// Restituisce { rowIndex (1-based, per setRange), rowData } oppure null.
function findByBarcode(sheet, barcode) {
  const allData = sheet.getDataRange().getValues();
  const target  = String(barcode).trim();

  for (let i = 1; i < allData.length; i++) {   // i=1: salta riga intestazioni
    if (String(allData[i][COL.BARCODE]).trim() === target) {
      return {
        rowIndex: i + 1,   // 1-based per l'API Sheets
        rowData:  allData[i],
      };
    }
  }
  return null;
}

// Formatta data (oggetto Date o stringa ISO YYYY-MM-DD) in GG/MM/AAAA
function formatDateIT(value) {
  if (!value) return '';
  try {
    const d = (value instanceof Date) ? value : new Date(value + 'T00:00:00');
    if (isNaN(d.getTime())) return String(value);
    const dd = String(d.getDate()).padStart(2, '0');
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    return `${dd}/${mm}/${d.getFullYear()}`;
  } catch (_) {
    return String(value);
  }
}

// Pulisce una stringa da caratteri pericolosi per le celle
function sanitize(v) {
  if (v === null || v === undefined) return '';
  const s = String(v).trim();
  // Impedisce formula injection nelle celle Google Sheets
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

// ================================================================
// AZIONE: CERCA
// ================================================================
function searchProduct(barcode) {
  if (!barcode) return { success: false, message: 'Codice a barre non fornito.' };

  const sheet = getSheet();
  const found = findByBarcode(sheet, barcode);

  if (!found) {
    return {
      success: false,
      message: 'Nessun prodotto trovato per il barcode: ' + barcode,
    };
  }

  const r = found.rowData;
  return {
    success:     true,
    boxNumber:   r[COL.BOX],
    article:     r[COL.ARTICLE],
    barcode:     r[COL.BARCODE],
    brand:       r[COL.BRAND],
    productName: r[COL.NAME],
    color:       r[COL.COLOR],
    size:        r[COL.SIZE],
    quantity:    Number(r[COL.QUANTITY]) || 0,
    expiry:      r[COL.EXPIRY]  ? String(r[COL.EXPIRY])  : '',
    price:       r[COL.PRICE]   ? Number(r[COL.PRICE])   : 0,
    date:        r[COL.DATE]    ? String(r[COL.DATE])     : '',
  };
}

// ================================================================
// AZIONE: AGGIUNGI (+1 unità)
// ================================================================
function addUnit(barcode, expiry, dateISO) {
  if (!barcode) return { success: false, message: 'Codice a barre non fornito.' };

  const sheet = getSheet();
  const found = findByBarcode(sheet, barcode);

  if (!found) {
    return {
      success: false,
      message: 'Prodotto non trovato. Usa "Nuovo" per registrarlo.',
    };
  }

  const { rowIndex, rowData } = found;
  const newQty = (Number(rowData[COL.QUANTITY]) || 0) + 1;

  // Aggiorna Quantità (col H → indice Sheets = COL.QUANTITY + 1 = 8+1 = 9... attento:
  // COL è 0-based, Sheets è 1-based → colonna Sheets = COL + 1)
  sheet.getRange(rowIndex, COL.QUANTITY + 1).setValue(newQty);

  // Aggiorna Data controllo
  const today = dateISO ? formatDateIT(dateISO) : formatDateIT(new Date());
  sheet.getRange(rowIndex, COL.DATE + 1).setValue(today);

  // Aggiorna Scadenza solo se fornita
  if (expiry && expiry.trim()) {
    sheet.getRange(rowIndex, COL.EXPIRY + 1).setValue(sanitize(expiry));
  }

  return {
    success:     true,
    productName: rowData[COL.NAME],
    boxNumber:   rowData[COL.BOX],
    quantity:    newQty,
  };
}

// ================================================================
// AZIONE: PRELEVA (−1 unità)
// ================================================================
function removeUnit(barcode, dateISO) {
  if (!barcode) return { success: false, message: 'Codice a barre non fornito.' };

  const sheet = getSheet();
  const found = findByBarcode(sheet, barcode);

  if (!found) {
    return { success: false, message: 'Prodotto non trovato nel magazzino.' };
  }

  const { rowIndex, rowData } = found;
  const currentQty = Number(rowData[COL.QUANTITY]) || 0;

  if (currentQty <= 0) {
    return {
      success: false,
      message: 'Quantità già a zero per: ' + rowData[COL.NAME],
    };
  }

  const newQty = currentQty - 1;

  // Aggiorna Quantità
  sheet.getRange(rowIndex, COL.QUANTITY + 1).setValue(newQty);

  // Aggiorna Data controllo
  const today = dateISO ? formatDateIT(dateISO) : formatDateIT(new Date());
  sheet.getRange(rowIndex, COL.DATE + 1).setValue(today);

  return {
    success:     true,
    productName: rowData[COL.NAME],
    boxNumber:   rowData[COL.BOX],
    quantity:    newQty,
    isLastItem:  newQty === 0,
  };
}

// ================================================================
// AZIONE: CREA NUOVO PRODOTTO
// ================================================================
function createProduct(data) {
  if (!data.barcode) return { success: false, message: 'Barcode obbligatorio.' };
  if (!data.name)    return { success: false, message: 'Nome prodotto obbligatorio.' };

  const sheet = getSheet();

  // Controlla duplicati per barcode
  const existing = findByBarcode(sheet, data.barcode);
  if (existing) {
    return {
      success: false,
      message: 'Barcode già presente per: "' + existing.rowData[COL.NAME] +
               '". Usa "Aggiungi" per incrementare la quantità.',
    };
  }

  const today = data.date ? formatDateIT(data.date) : formatDateIT(new Date());

  // Riga nell'ordine esatto delle colonne A→K
  const newRow = [
    sanitize(data.box),       // A: Numero scatola
    sanitize(data.article),   // B: Codice Articolo
    sanitize(data.barcode),   // C: Codice a barre
    sanitize(data.brand),     // D: Marca
    sanitize(data.name),      // E: Nome Prodotto
    sanitize(data.color),     // F: Colore
    sanitize(data.size),      // G: Taglia
    Number(data.quantity) || 0,             // H: Quantità
    sanitize(data.expiry),    // I: Scadenza
    Number(data.price) || 0,               // J: Prezzo
    today,                    // K: Data controllo
  ];

  sheet.appendRow(newRow);

  return {
    success:     true,
    message:     'Prodotto creato con successo.',
    productName: data.name,
  };
}

// ================================================================
// FUNZIONI DI TEST — eseguibili manualmente dalla dashboard GAS
// ================================================================

/** Test: cerca un prodotto (sostituisci il barcode) */
function _testSearch() {
  const fake = { postData: { contents: JSON.stringify({ action: 'search', barcode: '8001000000001' }) } };
  Logger.log(doPost(fake).getContent());
}

/** Test: aggiunge un'unità */
function _testAdd() {
  const fake = { postData: { contents: JSON.stringify({
    action: 'add', barcode: '8001000000001', expiry: '2026-12-31', date: '2026-05-25'
  }) } };
  Logger.log(doPost(fake).getContent());
}

/** Test: preleva un'unità */
function _testRemove() {
  const fake = { postData: { contents: JSON.stringify({
    action: 'remove', barcode: '8001000000001', date: '2026-05-25'
  }) } };
  Logger.log(doPost(fake).getContent());
}

/** Test: crea nuovo prodotto */
function _testCreate() {
  const fake = { postData: { contents: JSON.stringify({
    action:   'create',
    barcode:  '8001000099999',
    box:      'SC-042',
    article:  'ART099',
    brand:    'TestBrand',
    name:     'Prodotto di Test',
    color:    'Blu',
    size:     'M',
    quantity: 5,
    expiry:   '2027-06-30',
    price:    12.50,
    date:     '2026-05-25',
  }) } };
  Logger.log(doPost(fake).getContent());
}

// ================================================================
// PING — risponde alle richieste GET per verificare che il backend sia online
// ================================================================
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({
      status: "ok",
      message: "Backend magazzino attivo",
      timestamp: new Date().toISOString()
    }))
    .setMimeType(ContentService.MimeType.JSON);
}