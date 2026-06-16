/* STEP Dashboard v37.74 - Pesquisa, visualização, download e impressão Zebra ZD230 203dpi de QR Codes por ISO. */
const isoQrModalEl = document.getElementById('iso-qr-modal');
const isoQrCloseEl = document.getElementById('iso-qr-close');
const isoQrSearchEl = document.getElementById('iso-qr-search');
const isoQrSearchButtonEl = document.getElementById('iso-qr-search-button');
const isoQrPrintSelectedEl = document.getElementById('iso-qr-print-selected');
const isoQrDownloadSelectedEl = document.getElementById('iso-qr-download-selected');
const isoQrFeedbackEl = document.getElementById('iso-qr-feedback');
const isoQrPreviewEl = document.getElementById('iso-qr-preview');
const isoQrResultsEl = document.getElementById('iso-qr-results');

const isoQrState = {
  items: [],
  selected: new Set(),
  loading: false,
};

function canOpenIsoQrModule(user = state.user) {
  return Boolean(user && !isClientUser(user));
}

function isoQrImageUrl(item, width = 360, download = false) {
  const token = encodeURIComponent(String(item?.qrToken || ''));
  const params = new URLSearchParams({ token, w: String(width) });
  if (download) params.set('download', '1');
  // URLSearchParams codifica o token de novo quando recebe já escapado; por isso montamos direto.
  return `/api/iso-qr-image?token=${token}&w=${encodeURIComponent(String(width))}${download ? '&download=1' : ''}`;
}

function normalizeIsoQrFileName(value = '') {
  return String(value || 'iso-qrcode')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9._-]+/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '')
    .slice(0, 120) || 'iso-qrcode';
}

function getIsoQrItemByToken(token = '') {
  return isoQrState.items.find((item) => String(item.qrToken || '') === String(token || '')) || null;
}

function getSelectedIsoQrItems() {
  return Array.from(isoQrState.selected)
    .map((token) => getIsoQrItemByToken(token))
    .filter(Boolean);
}

function updateIsoQrActionButtons() {
  const count = isoQrState.selected.size;
  if (isoQrPrintSelectedEl) {
    isoQrPrintSelectedEl.disabled = count === 0;
    isoQrPrintSelectedEl.textContent = count ? `Imprimir selecionados (${count})` : 'Imprimir selecionados';
  }
  if (isoQrDownloadSelectedEl) {
    isoQrDownloadSelectedEl.disabled = count === 0;
    isoQrDownloadSelectedEl.textContent = count ? `Baixar selecionados (${count})` : 'Baixar selecionados';
  }
}

function renderIsoQrPreview(item) {
  if (!isoQrPreviewEl) return;
  if (!item) {
    isoQrPreviewEl.classList.add('hidden');
    isoQrPreviewEl.innerHTML = '';
    return;
  }
  const isoName = item.isoFullName || item.iso || 'ISO';
  isoQrPreviewEl.classList.remove('hidden');
  isoQrPreviewEl.innerHTML = `
    <div class="iso-qr-preview-card">
      <div class="iso-qr-label-preview">
        <img src="${escapeHtml(isoQrImageUrl(item, 360))}" alt="QR Code ${escapeHtml(isoName)}" />
        <strong>${escapeHtml(isoName)}</strong>
      </div>
      <div class="iso-qr-preview-actions">
        <button class="primary-button" type="button" data-iso-qr-print="${escapeHtml(item.qrToken)}">Imprimir</button>
        <button class="ghost-button" type="button" data-iso-qr-download="${escapeHtml(item.qrToken)}">Baixar SVG</button>
      </div>
    </div>`;
}

function renderIsoQrResults() {
  if (!isoQrResultsEl) return;
  const items = Array.isArray(isoQrState.items) ? isoQrState.items : [];
  updateIsoQrActionButtons();

  if (isoQrState.loading) {
    isoQrResultsEl.className = 'iso-qr-results empty-state';
    isoQrResultsEl.textContent = 'Carregando QR Codes...';
    return;
  }

  if (!items.length) {
    isoQrResultsEl.className = 'iso-qr-results empty-state';
    isoQrResultsEl.textContent = 'Nenhum QR Code encontrado. Se a BSP/ISO for nova, aguarde a atualização do cache ou clique em Atualizar agora.';
    return;
  }

  isoQrResultsEl.className = 'iso-qr-results';
  isoQrResultsEl.innerHTML = items.map((item) => {
    const token = String(item.qrToken || '');
    const checked = isoQrState.selected.has(token) ? 'checked' : '';
    const isoName = item.isoFullName || item.iso || 'ISO';
    return `
      <article class="iso-qr-card">
        <label class="iso-qr-select-line">
          <input type="checkbox" data-iso-qr-select="${escapeHtml(token)}" ${checked} />
          <span>
            <strong>${escapeHtml(isoName)}</strong>
            <small>${escapeHtml([item.bsp, item.client, item.vessel].filter(Boolean).join(' • ') || 'QR Code automático')}</small>
          </span>
        </label>
        <div class="iso-qr-card-actions">
          <button class="ghost-button ghost-button--compact" type="button" data-iso-qr-preview="${escapeHtml(token)}">Visualizar</button>
          <button class="ghost-button ghost-button--compact" type="button" data-iso-qr-print="${escapeHtml(token)}">Imprimir</button>
          <button class="ghost-button ghost-button--compact" type="button" data-iso-qr-download="${escapeHtml(token)}">Baixar</button>
        </div>
      </article>`;
  }).join('');
}

async function loadIsoQrCodes() {
  if (!isoQrResultsEl || !canOpenIsoQrModule()) return;
  const q = String(isoQrSearchEl?.value || '').trim();
  isoQrState.loading = true;
  isoQrState.selected.clear();
  renderIsoQrPreview(null);
  renderIsoQrResults();
  if (isoQrFeedbackEl) isoQrFeedbackEl.textContent = q ? `Pesquisando “${q}”...` : 'Carregando QR Codes recentes...';

  try {
    const params = new URLSearchParams({ limit: '150' });
    if (q) params.set('q', q);
    const response = await fetch(`/api/iso-qr-codes?${params.toString()}`, {
      credentials: 'same-origin',
      cache: 'no-store',
      headers: { 'Cache-Control': 'no-store' },
    });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) throw new Error(data?.error || 'Falha ao consultar QR Codes.');
    isoQrState.items = Array.isArray(data.items) ? data.items : [];
    if (isoQrFeedbackEl) {
      isoQrFeedbackEl.textContent = isoQrState.items.length
        ? `${isoQrState.items.length} QR Code(s) encontrado(s). Selecione para imprimir ou baixar.`
        : 'Nenhum QR Code encontrado para essa busca.';
    }
  } catch (error) {
    isoQrState.items = [];
    if (isoQrFeedbackEl) isoQrFeedbackEl.textContent = error.message || 'Falha ao consultar QR Codes.';
  } finally {
    isoQrState.loading = false;
    renderIsoQrResults();
  }
}

function openIsoQrModal() {
  if (!isoQrModalEl || !canOpenIsoQrModule()) return;
  isoQrModalEl.classList.remove('hidden');
  isoQrModalEl.setAttribute('aria-hidden', 'false');
  document.body.classList.add('modal-open');
  if (isoQrSearchEl) {
    isoQrSearchEl.focus();
    isoQrSearchEl.select?.();
  }
  if (!isoQrState.items.length) loadIsoQrCodes();
}

function closeIsoQrModal() {
  if (!isoQrModalEl) return;
  isoQrModalEl.classList.add('hidden');
  isoQrModalEl.setAttribute('aria-hidden', 'true');
  document.body.classList.remove('modal-open');
}

function chunkIsoQrItemsForZebra(items = [], chunkSize = 3) {
  const list = Array.isArray(items) ? items.filter(Boolean) : [];
  const chunks = [];
  for (let i = 0; i < list.length; i += chunkSize) chunks.push(list.slice(i, i + chunkSize));
  return chunks.length ? chunks : [];
}

function printIsoQrItems(items = []) {
  const list = Array.isArray(items) ? items.filter(Boolean) : [];
  if (!list.length) return;
  const printWindow = window.open('', '_blank', 'width=520,height=760');
  if (!printWindow) {
    window.alert('O navegador bloqueou a janela de impressão. Permita pop-ups para imprimir as etiquetas.');
    return;
  }

  const pages = chunkIsoQrItemsForZebra(list, 3).map((chunk) => {
    const slots = [0, 1, 2].map((slotIndex) => {
      const item = chunk[slotIndex];
      if (!item) return '<section class="qr-slot qr-slot--empty" aria-hidden="true"></section>';
      const isoName = item.isoFullName || item.iso || 'ISO';
      return `
        <section class="qr-slot">
          <div class="qr-box"><img src="${escapeHtml(isoQrImageUrl(item, 900))}" alt="QR Code ${escapeHtml(isoName)}" /></div>
          <strong>${escapeHtml(isoName)}</strong>
        </section>`;
    }).join('');
    return `<main class="zebra-sheet">${slots}</main>`;
  }).join('');

  printWindow.document.open();
  printWindow.document.write(`<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8" />
<title>Etiquetas QR Code ISO</title>
<style>
  /* v37.74 - Zebra ZD230 203dpi ZPL
     IMPORTANTE: o tamanho físico fica no driver ZDesigner (Paper Size = Custom da etiqueta).
     O HTML não força 50x90mm para não brigar com o Custom da Zebra/Chrome.
     A página sempre tem 3 quadros fixos: se selecionar 1 ou 2, os demais saem em branco. */
  @page { margin: 0; }
  * { box-sizing: border-box; }
  html,
  body {
    margin: 0 !important;
    padding: 0 !important;
    width: 100% !important;
    height: 100% !important;
    overflow: hidden !important;
    background: #fff;
    color: #000;
    font-family: Arial, Helvetica, sans-serif;
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
  }
  .zebra-sheet {
    width: 100vw;
    height: 100vh;
    min-height: 100vh;
    max-height: 100vh;
    display: grid;
    grid-template-rows: repeat(3, 33.333333vh);
    overflow: hidden;
    break-after: page;
    page-break-after: always;
    background: #fff;
  }
  .zebra-sheet:last-child { break-after: auto; page-break-after: auto; }
  .qr-slot {
    width: 100%;
    height: 33.333333vh;
    display: grid;
    grid-template-rows: 1fr auto;
    align-items: center;
    justify-items: center;
    gap: 0.7mm;
    padding: 1.4mm 1mm 1mm;
    overflow: hidden;
    background: #fff;
  }
  .qr-slot--empty { color: transparent; }
  .qr-box {
    width: 22mm;
    height: 22mm;
    max-width: 88%;
    max-height: calc(33.333333vh - 7mm);
    display: grid;
    place-items: center;
    overflow: hidden;
  }
  .qr-box img {
    width: 100%;
    height: 100%;
    object-fit: contain;
    display: block;
    image-rendering: pixelated;
    image-rendering: crisp-edges;
  }
  .qr-slot strong {
    display: block;
    width: 96%;
    max-height: 5.5mm;
    overflow: hidden;
    text-align: center;
    font-size: 5.6pt;
    line-height: 1.05;
    font-weight: 700;
    word-break: break-word;
    white-space: normal;
  }
  @media screen {
    body { padding: 10px !important; background: #e5e7eb; overflow: auto !important; }
    .zebra-sheet {
      width: 50mm;
      height: 90mm;
      min-height: 90mm;
      max-height: none;
      grid-template-rows: repeat(3, 30mm);
      margin: 0 auto 12px;
      border: 1px solid #cfcfcf;
      box-shadow: 0 8px 28px rgba(0,0,0,.16);
    }
    .qr-slot { height: 30mm; }
  }
  @media print {
    .zebra-sheet { position: relative; }
    .qr-slot { page-break-inside: avoid; break-inside: avoid; }
  }
</style>
</head>
<body>
  ${pages}
  <script>
    (function(){
      function waitForImages(timeoutMs){
        var imgs = Array.prototype.slice.call(document.images || []);
        if (!imgs.length) return Promise.resolve();
        var done = imgs.map(function(img){
          if (img.complete && img.naturalWidth > 0) return Promise.resolve();
          return new Promise(function(resolve){
            var finished = false;
            function end(){ if (!finished) { finished = true; resolve(); } }
            img.addEventListener('load', end, { once: true });
            img.addEventListener('error', end, { once: true });
            setTimeout(end, timeoutMs || 3500);
          });
        });
        return Promise.all(done).then(function(){ return undefined; });
      }
      window.addEventListener('load', function(){
        waitForImages(4500).then(function(){
          setTimeout(function(){ window.focus(); window.print(); }, 350);
        });
      });
    })();
  <\/script>
</body>
</html>`);
  printWindow.document.close();
}

function downloadIsoQrItems(items = []) {
  const list = Array.isArray(items) ? items.filter(Boolean) : [];
  list.forEach((item, index) => {
    window.setTimeout(() => {
      const isoName = item.isoFullName || item.iso || 'ISO';
      const link = document.createElement('a');
      link.href = isoQrImageUrl(item, 600, true);
      link.download = `${normalizeIsoQrFileName(isoName)}.svg`;
      document.body.appendChild(link);
      link.click();
      link.remove();
    }, index * 250);
  });
}

function handleIsoQrClick(event) {
  const closeTarget = event.target.closest('[data-close-iso-qr]');
  if (closeTarget) {
    closeIsoQrModal();
    return;
  }

  const select = event.target.closest('[data-iso-qr-select]');
  if (select) {
    const token = select.getAttribute('data-iso-qr-select') || '';
    if (select.checked) isoQrState.selected.add(token);
    else isoQrState.selected.delete(token);
    updateIsoQrActionButtons();
    return;
  }

  const preview = event.target.closest('[data-iso-qr-preview]');
  if (preview) {
    renderIsoQrPreview(getIsoQrItemByToken(preview.getAttribute('data-iso-qr-preview') || ''));
    return;
  }

  const print = event.target.closest('[data-iso-qr-print]');
  if (print) {
    const item = getIsoQrItemByToken(print.getAttribute('data-iso-qr-print') || '');
    if (item) printIsoQrItems([item]);
    return;
  }

  const download = event.target.closest('[data-iso-qr-download]');
  if (download) {
    const item = getIsoQrItemByToken(download.getAttribute('data-iso-qr-download') || '');
    if (item) downloadIsoQrItems([item]);
  }
}

function bindIsoQrEvents() {
  if (openIsoQrButtonEl) openIsoQrButtonEl.addEventListener('click', openIsoQrModal);
  if (isoQrCloseEl) isoQrCloseEl.addEventListener('click', closeIsoQrModal);
  if (isoQrModalEl) isoQrModalEl.addEventListener('click', handleIsoQrClick);
  if (isoQrSearchButtonEl) isoQrSearchButtonEl.addEventListener('click', loadIsoQrCodes);
  if (isoQrSearchEl) {
    isoQrSearchEl.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') {
        event.preventDefault();
        loadIsoQrCodes();
      }
      if (event.key === 'Escape') closeIsoQrModal();
    });
  }
  if (isoQrPrintSelectedEl) {
    isoQrPrintSelectedEl.addEventListener('click', () => printIsoQrItems(getSelectedIsoQrItems()));
  }
  if (isoQrDownloadSelectedEl) {
    isoQrDownloadSelectedEl.addEventListener('click', () => downloadIsoQrItems(getSelectedIsoQrItems()));
  }
}

bindIsoQrEvents();
