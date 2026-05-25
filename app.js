/* ================================================================
   MagMa — app.js
   Modifica API_URL con l'URL del tuo Google Apps Script Web App
================================================================ */
const API_URL = 'https://script.google.com/macros/s/AKfycbwfh_TcxlKvAzCm5so6Z-mGJXp4Xv6thW2BbwI906-b/dev/exec';


/* ================================================================
   STATO APPLICAZIONE
================================================================ */
let currentAction   = null;   // 'add' | 'remove' | 'create' | 'search'
let scanControls    = null;   // IScannerControls di ZXing
let alertTimer      = null;   // timer per auto-dismiss dell'alert
let pendingBarcode  = null;   // barcode in attesa (usato per AGGIUNGI + popup)
let codeReader      = null;   // istanza BrowserMultiFormatReader


/* ================================================================
   RIFERIMENTI DOM
================================================================ */
const $ = id => document.getElementById(id);

const dom = {
  // Schermate
  screenHome:     $('screen-home'),
  screenScanner:  $('screen-scanner'),

  // Scanner
  scannerVideo:   $('scanner-video'),
  scannerTitle:   $('scanner-title'),
  btnCancelScan:  $('btn-cancel-scan'),

  // Spinner
  spinnerOverlay: $('spinner-overlay'),

  // Alert
  alertBanner:    $('alert-banner'),
  alertMessage:   $('alert-message'),
  alertClose:     $('alert-close'),

  // Home: pulsanti azione
  btnAdd:         $('btn-add'),
  btnRemove:      $('btn-remove'),
  btnNew:         $('btn-new'),
  btnSearch:      $('btn-search'),

  // Modal scadenza (AGGIUNGI)
  modalExpiry:    $('modal-expiry'),
  expiryBackdrop: $('expiry-backdrop'),
  inputExpiry:    $('input-expiry'),
  btnExpirySkip:  $('btn-expiry-skip'),
  btnExpiryConfirm: $('btn-expiry-confirm'),

  // Modal form (NUOVO)
  modalForm:      $('modal-form'),
  formBackdrop:   $('form-backdrop'),
  formNewProduct: $('form-new-product'),
  fBarcode:       $('f-barcode'),
  fBox:           $('f-box'),
  fArticle:       $('f-article'),
  fBrand:         $('f-brand'),
  fName:          $('f-name'),
  fColor:         $('f-color'),
  fSize:          $('f-size'),
  fQty:           $('f-qty'),
  fExpiry:        $('f-expiry'),
  fPrice:         $('f-price'),
  btnFormCancel:  $('btn-form-cancel'),

  // Modal risultato (CERCA / PRELEVA ultimo)
  modalResult:    $('modal-result'),
  resultBackdrop: $('result-backdrop'),
  resultIcon:     $('result-icon'),
  resultTitle:    $('result-title'),
  resultCard:     $('result-card'),
  btnResultClose: $('btn-result-close'),
};


/* ================================================================
   INIZIALIZZAZIONE
================================================================ */
document.addEventListener('DOMContentLoaded', () => {
  initScanner();
  setupEventListeners();
  registerServiceWorker();
});

function initScanner() {
  if (typeof ZXingBrowser === 'undefined') {
    console.warn('ZXing non caricato. La scansione non sarà disponibile.');
    return;
  }
  codeReader = new ZXingBrowser.BrowserMultiFormatReader();
}

function setupEventListeners() {
  // Pulsanti home
  dom.btnAdd.addEventListener('click',    () => startScan('add'));
  dom.btnRemove.addEventListener('click', () => startScan('remove'));
  dom.btnNew.addEventListener('click',    () => startScan('create'));
  dom.btnSearch.addEventListener('click', () => startScan('search'));

  // Scanner — annulla
  dom.btnCancelScan.addEventListener('click', stopScanner);

  // Alert — chiudi
  dom.alertClose.addEventListener('click', hideAlert);

  // Modal scadenza
  dom.btnExpirySkip.addEventListener('click', () => resolveExpiry(null));
  dom.btnExpiryConfirm.addEventListener('click', () => {
    resolveExpiry(dom.inputExpiry.value.trim() || null);
  });
  dom.expiryBackdrop.addEventListener('click', () => resolveExpiry(null));

  // Modal form nuovo prodotto
  dom.btnFormCancel.addEventListener('click', closeForm);
  dom.formBackdrop.addEventListener('click', closeForm);
  dom.formNewProduct.addEventListener('submit', handleFormSubmit);

  // Modal risultato
  dom.btnResultClose.addEventListener('click', closeResult);
  dom.resultBackdrop.addEventListener('click', closeResult);
}

async function registerServiceWorker() {
  if ('serviceWorker' in navigator) {
    try {
      await navigator.serviceWorker.register('./service-worker.js');
    } catch (err) {
      console.warn('Service Worker non registrato:', err);
    }
  }
}


/* ================================================================
   SCANNER
================================================================ */
const SCAN_TITLES = {
  add:    '➕  Aggiungi unità',
  remove: '➖  Preleva unità',
  create: '🆕  Nuovo prodotto',
  search: '🔍  Cerca prodotto',
};

async function startScan(action) {
  if (!codeReader) {
    showAlert('Libreria scanner non disponibile. Ricarica la pagina.', 'error');
    return;
  }

  currentAction = action;
  dom.scannerTitle.textContent = SCAN_TITLES[action];

  // Mostra schermata scanner
  dom.screenHome.classList.add('hidden');
  dom.screenScanner.classList.remove('hidden');

  try {
    // Usa la fotocamera posteriore se disponibile (ideale per barcode)
    const constraints = {
      video: {
        facingMode: { ideal: 'environment' },
        width:  { ideal: 1280 },
        height: { ideal: 720 },
      }
    };

    let resolved = false;

    scanControls = await codeReader.decodeFromConstraints(
      constraints,
      dom.scannerVideo,
      (result, error) => {
        if (result && !resolved) {
          resolved = true;
          const barcode = result.getText();
          stopScanner();
          handleScan(barcode);
        }
      }
    );
  } catch (err) {
    stopScanner();
    let msg = 'Impossibile accedere alla fotocamera.';
    if (err.name === 'NotAllowedError')
      msg = 'Permesso fotocamera negato. Abilitalo nelle impostazioni.';
    else if (err.name === 'NotFoundError')
      msg = 'Nessuna fotocamera trovata sul dispositivo.';
    showAlert(msg, 'error');
  }
}

function stopScanner() {
  if (scanControls) {
    try { scanControls.stop(); } catch (_) {}
    scanControls = null;
  }
  dom.screenScanner.classList.add('hidden');
  dom.screenHome.classList.remove('hidden');
}


/* ================================================================
   ROUTING SCAN → AZIONE
================================================================ */
async function handleScan(barcode) {
  if (!barcode || !barcode.trim()) {
    showAlert('Codice a barre non valido. Riprova.', 'warning');
    return;
  }

  switch (currentAction) {
    case 'add':    await handleAdd(barcode.trim());    break;
    case 'remove': await handleRemove(barcode.trim()); break;
    case 'create': await handleCreate(barcode.trim()); break;
    case 'search': await handleSearch(barcode.trim()); break;
  }
}


/* ================================================================
   AZIONE: AGGIUNGI (+1 unità)
================================================================ */
async function handleAdd(barcode) {
  pendingBarcode = barcode;
  dom.inputExpiry.value = '';
  dom.modalExpiry.classList.remove('hidden');
}

function resolveExpiry(expiry) {
  dom.modalExpiry.classList.add('hidden');
  if (pendingBarcode) {
    performAdd(pendingBarcode, expiry);
    pendingBarcode = null;
  }
}

async function performAdd(barcode, expiry) {
  const res = await apiCall({
    action:  'add',
    barcode,
    expiry:  expiry || '',
    date:    todayISO(),
  });
  if (!res) return;

  if (res.success) {
    showAlert(
      `✓ Aggiunto! ${res.productName || barcode} — Qtà: ${res.quantity}`,
      'success'
    );
  } else {
    showAlert(res.message || 'Prodotto non trovato nel magazzino.', 'error');
  }
}


/* ================================================================
   AZIONE: PRELEVA (−1 unità)
================================================================ */
async function handleRemove(barcode) {
  const res = await apiCall({
    action:  'remove',
    barcode,
    date:    todayISO(),
  });
  if (!res) return;

  if (res.success) {
    if (res.quantity === 0) {
      // Quantità a zero → mostra scheda di avviso prominente
      showResult('warning', '⚠️', `Ultimo pezzo della scatola N.${res.boxNumber || '?'}!`, [
        { label: 'Prodotto',  value: res.productName || '—' },
        { label: 'Scatola',   value: `N. ${res.boxNumber || '—'}` },
        { label: 'Quantità',  value: '0 — ESAURITO', cls: 'qty-zero' },
        { label: 'Barcode',   value: barcode },
      ]);
    } else {
      const suffix = res.quantity <= 3 ? ` ⚠️ Rimangono solo ${res.quantity}!` : `— Qtà: ${res.quantity}`;
      showAlert(`✓ Prelevato! ${res.productName || barcode} ${suffix}`, 'success');
    }
  } else {
    showAlert(res.message || 'Prodotto non trovato nel magazzino.', 'error');
  }
}


/* ================================================================
   AZIONE: NUOVO PRODOTTO
================================================================ */
async function handleCreate(barcode) {
  // Resetta e precompila il form
  dom.formNewProduct.reset();
  dom.fBarcode.value = barcode;
  dom.fQty.value     = '1';

  // Rimuovi eventuali errori precedenti
  dom.formNewProduct.querySelectorAll('.invalid').forEach(el => el.classList.remove('invalid'));

  dom.modalForm.classList.remove('hidden');

  // Focus sul primo campo editabile
  setTimeout(() => dom.fBox.focus(), 350);
}

async function handleFormSubmit(e) {
  e.preventDefault();

  // Validazione minima
  if (!dom.fName.value.trim()) {
    dom.fName.classList.add('invalid');
    dom.fName.focus();
    showAlert('Il campo "Nome prodotto" è obbligatorio.', 'warning');
    return;
  }

  const qty = parseInt(dom.fQty.value, 10);
  if (isNaN(qty) || qty < 0) {
    dom.fQty.classList.add('invalid');
    dom.fQty.focus();
    showAlert('La quantità deve essere un numero valido.', 'warning');
    return;
  }

  const res = await apiCall({
    action:   'create',
    barcode:  dom.fBarcode.value.trim(),
    box:      dom.fBox.value.trim(),
    article:  dom.fArticle.value.trim(),
    brand:    dom.fBrand.value.trim(),
    name:     dom.fName.value.trim(),
    color:    dom.fColor.value.trim(),
    size:     dom.fSize.value.trim(),
    quantity: qty,
    expiry:   dom.fExpiry.value || '',
    price:    parseFloat(dom.fPrice.value) || 0,
    date:     todayISO(),
  });

  if (!res) return;

  if (res.success) {
    closeForm();
    showAlert(`✓ Prodotto "${dom.fName.value.trim()}" creato con successo!`, 'success');
  } else {
    showAlert(res.message || 'Errore durante la creazione del prodotto.', 'error');
  }
}

function closeForm() {
  dom.modalForm.classList.add('hidden');
  dom.formNewProduct.reset();
}


/* ================================================================
   AZIONE: CERCA
================================================================ */
async function handleSearch(barcode) {
  const res = await apiCall({ action: 'search', barcode });
  if (!res) return;

  if (res.success) {
    const qty = parseInt(res.quantity, 10) || 0;
    let qtyCls = 'qty-ok';
    if (qty === 0)     qtyCls = 'qty-zero';
    else if (qty <= 3) qtyCls = 'qty-low';

    showResult('success', '📦', res.productName || 'Prodotto trovato', [
      { label: 'Scatola N.',   value: res.boxNumber  || '—' },
      { label: 'Marca',        value: res.brand       || '—' },
      { label: 'Articolo',     value: res.article     || '—' },
      { label: 'Colore',       value: res.color       || '—' },
      { label: 'Taglia',       value: res.size        || '—' },
      { label: 'Quantità',     value: String(qty),   cls: qtyCls },
      { label: 'Scadenza',     value: res.expiry      || '—' },
      { label: 'Prezzo',       value: res.price ? `€ ${parseFloat(res.price).toFixed(2)}` : '—' },
      { label: 'Ult. controllo', value: res.date      || '—' },
    ]);
  } else {
    showAlert(res.message || 'Prodotto non trovato per questo barcode.', 'error');
  }
}


/* ================================================================
   UI HELPERS
================================================================ */
function showSpinner() {
  dom.spinnerOverlay.classList.remove('hidden');
}

function hideSpinner() {
  dom.spinnerOverlay.classList.add('hidden');
}

function showAlert(message, type = 'info', durationMs = 4500) {
  clearTimeout(alertTimer);
  dom.alertBanner.className = `alert-${type}`;
  dom.alertMessage.textContent = message;
  dom.alertBanner.classList.remove('hidden');
  alertTimer = setTimeout(hideAlert, durationMs);
}

function hideAlert() {
  dom.alertBanner.classList.add('hidden');
  clearTimeout(alertTimer);
  alertTimer = null;
}

function showResult(type, icon, title, rows) {
  dom.resultIcon.className   = `result-icon icon-${type}`;
  dom.resultIcon.textContent = icon;
  dom.resultTitle.textContent = title;

  dom.resultCard.innerHTML = rows.map(row => `
    <div class="result-row">
      <span class="result-label">${escHtml(row.label)}</span>
      <span class="result-value ${row.cls || ''}">${escHtml(String(row.value))}</span>
    </div>
  `).join('');

  dom.modalResult.classList.remove('hidden');
}

function closeResult() {
  dom.modalResult.classList.add('hidden');
}

function escHtml(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}


/* ================================================================
   CHIAMATA API
================================================================ */
async function apiCall(payload) {
  showSpinner();
  try {
    const response = await fetch(API_URL, {
      method:  'POST',
      // Google Apps Script richiede text/plain per evitare CORS preflight
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body:    JSON.stringify(payload),
    });

    if (!response.ok) {
      throw new Error(`Risposta HTTP ${response.status}: ${response.statusText}`);
    }

    const data = await response.json();
    return data;

  } catch (err) {
    const isOffline = !navigator.onLine || err instanceof TypeError;
    if (isOffline) {
      showAlert('Nessuna connessione. Controlla la rete e riprova.', 'error');
    } else {
      showAlert(`Errore di comunicazione: ${err.message}`, 'error');
    }
    return null;
  } finally {
    hideSpinner();
  }
}


/* ================================================================
   UTILITÀ
================================================================ */
function todayISO() {
  return new Date().toISOString().split('T')[0];
}
