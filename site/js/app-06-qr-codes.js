/* STEP Dashboard v37.75 - Pesquisa, visualização, download e impressão Zebra ZD230 203dpi de QR Codes por ISO. */
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

function isoQrSvgEscape(value = '') {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function isoQrDataUrlFromSvg(svg = '') {
  const encoded = encodeURIComponent(String(svg || ''))
    .replace(/'/g, '%27')
    .replace(/"/g, '%22');
  return `data:image/svg+xml;charset=UTF-8,${encoded}`;
}

async function fetchIsoQrSvgDataUrl(item, width = 700) {
  const response = await fetch(isoQrImageUrl(item, width), {
    credentials: 'same-origin',
    cache: 'no-store',
    headers: { 'Cache-Control': 'no-store' },
  });
  if (!response.ok) throw new Error('Falha ao carregar imagem do QR Code.');
  const svg = await response.text();
  return isoQrDataUrlFromSvg(svg);
}

function wrapIsoQrLabelForZebra(label = '') {
  const text = String(label || 'ISO').replace(/\s+/g, ' ').trim() || 'ISO';
  const maxChars = 27;
  const words = text.split(' ');
  const lines = [];
  let current = '';

  for (const word of words) {
    if (!current) {
      current = word;
      continue;
    }
    if ((current + ' ' + word).length <= maxChars) {
      current += ' ' + word;
    } else {
      lines.push(current);
      current = word;
    }
  }
  if (current) lines.push(current);

  const normalizedLines = [];
  for (const line of lines) {
    if (line.length <= maxChars + 8) {
      normalizedLines.push(line);
      continue;
    }
    for (let i = 0; i < line.length; i += maxChars) {
      normalizedLines.push(line.slice(i, i + maxChars));
    }
  }

  if (normalizedLines.length <= 3) return normalizedLines;
  return [normalizedLines[0], normalizedLines[1], normalizedLines.slice(2).join(' ')];
}

async function buildIsoQrZebraSheetSvg(chunk = []) {
  const slots = [chunk[0] || null, chunk[1] || null, chunk[2] || null];
  const qrImages = await Promise.all(slots.map((item) => item ? fetchIsoQrSvgDataUrl(item, 760) : Promise.resolve('')));
  const slotHeight = 30;
  const pageWidth = 50;
  const qrSize = 20.5;
  const qrX = (pageWidth - qrSize) / 2;

  const slotMarkup = slots.map((item, index) => {
    const y = index * slotHeight;
    if (!item) {
      return `<g transform="translate(0 ${y})"><rect x="0" y="0" width="50" height="30" fill="#fff"/></g>`;
    }
    const isoName = item.isoFullName || item.iso || 'ISO';
    const lines = wrapIsoQrLabelForZebra(isoName);
    const longest = Math.max(...lines.map((line) => line.length), 1);
    const fontSize = longest > 44 ? 1.55 : (longest > 34 ? 1.75 : 2.0);
    const lineGap = fontSize + 0.35;
    const firstLineY = 24.0 - ((lines.length - 1) * lineGap / 2);
    const textLines = lines.map((line, lineIndex) => {
      const textLengthAttr = line.length > 32 ? ' textLength="46" lengthAdjust="spacingAndGlyphs"' : '';
      return `<tspan x="25" y="${(firstLineY + lineIndex * lineGap).toFixed(2)}"${textLengthAttr}>${isoQrSvgEscape(line)}</tspan>`;
    }).join('');

    return `<g transform="translate(0 ${y})">
      <rect x="0" y="0" width="50" height="30" fill="#fff"/>
      <image href="${qrImages[index]}" x="${qrX.toFixed(2)}" y="1.6" width="${qrSize}" height="${qrSize}" preserveAspectRatio="xMidYMid meet"/>
      <text x="25" y="24" text-anchor="middle" font-family="Arial, Helvetica, sans-serif" font-weight="700" font-size="${fontSize}" fill="#000">${textLines}</text>
    </g>`;
  }).join('');

  return `<svg xmlns="http://www.w3.org/2000/svg" width="50mm" height="90mm" viewBox="0 0 50 90">
    <rect x="0" y="0" width="50" height="90" fill="#fff"/>
    ${slotMarkup}
  </svg>`;
}

function writeIsoQrPrintLoading(printWindow) {
  printWindow.document.open();
  printWindow.document.write(`<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><title>Etiquetas QR Code ISO</title><style>body{margin:0;font-family:Arial,sans-serif;display:grid;place-items:center;min-height:100vh;color:#0f2740}.box{padding:18px;text-align:center}</style></head><body><div class="box"><strong>Preparando etiquetas Zebra...</strong><br><small>Aguarde o carregamento dos QR Codes.</small></div></body></html>`);
  printWindow.document.close();
}

async function printIsoQrItems(items = []) {
  const list = Array.isArray(items) ? items.filter(Boolean) : [];
  if (!list.length) return;
  const printWindow = window.open('', '_blank', 'width=520,height=760');
  if (!printWindow) {
    window.alert('O navegador bloqueou a janela de impressão. Permita pop-ups para imprimir as etiquetas.');
    return;
  }

  writeIsoQrPrintLoading(printWindow);

  try {
    const chunks = chunkIsoQrItemsForZebra(list, 3);
    const sheets = await Promise.all(chunks.map((chunk) => buildIsoQrZebraSheetSvg(chunk)));
    const pages = sheets.map((svg) => `<main class="zebra-page"><img class="zebra-page-image" src="${isoQrDataUrlFromSvg(svg)}" alt="Etiquetas QR Code ISO" /></main>`).join('');

    printWindow.document.open();
    printWindow.document.write(`<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8" />
<title>Etiquetas QR Code ISO</title>
<style>
  /* v37.75 - Zebra ZD230-203dpi ZPL
     A impressão selecionada agora vira uma única arte SVG por faixa, com 3 quadros fixos.
     Isso evita a Zebra/Chrome imprimir somente o primeiro QR quando há 3 imagens soltas. */
  @page { size: 50mm 90mm; margin: 0; }
  * { box-sizing: border-box; }
  html, body {
    margin: 0 !important;
    padding: 0 !important;
    width: 50mm;
    min-height: 90mm;
    background: #fff;
    color: #000;
    font-family: Arial, Helvetica, sans-serif;
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
  }
  .zebra-page {
    width: 50mm;
    height: 90mm;
    margin: 0;
    padding: 0;
    overflow: hidden;
    break-after: page;
    page-break-after: always;
    background: #fff;
  }
  .zebra-page:last-child { break-after: auto; page-break-after: auto; }
  .zebra-page-image {
    display: block;
    width: 50mm;
    height: 90mm;
    margin: 0;
    padding: 0;
    object-fit: fill;
    image-rendering: pixelated;
    image-rendering: crisp-edges;
  }
  @media screen {
    html, body { width: auto; min-height: 100vh; background: #e5e7eb; display: grid; place-items: start center; padding: 10px !important; }
    .zebra-page { border: 1px solid #cfcfcf; box-shadow: 0 8px 28px rgba(0,0,0,.18); margin: 0 auto 12px; }
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
        return Promise.all(imgs.map(function(img){
          if (img.complete && img.naturalWidth > 0) return Promise.resolve();
          return new Promise(function(resolve){
            var done = false;
            function finish(){ if (!done) { done = true; resolve(); } }
            img.addEventListener('load', finish, { once: true });
            img.addEventListener('error', finish, { once: true });
            setTimeout(finish, timeoutMs || 3500);
          });
        })).then(function(){ return undefined; });
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
  } catch (error) {
    printWindow.document.open();
    printWindow.document.write(`<!doctype html><html><body style="font-family:Arial,sans-serif;padding:18px"><strong>Falha ao preparar impressão Zebra.</strong><br>${escapeHtml(error.message || 'Erro desconhecido')}</body></html>`);
    printWindow.document.close();
  }
}

function downloadSingleIsoQrItem(item) {
  if (!item) return;
  const isoName = item.isoFullName || item.iso || 'ISO';
  const link = document.createElement('a');
  link.href = isoQrImageUrl(item, 600, true);
  link.download = `${normalizeIsoQrFileName(isoName)}.svg`;
  document.body.appendChild(link);
  link.click();
  link.remove();
}

async function downloadIsoQrItems(items = []) {
  const list = Array.isArray(items) ? items.filter(Boolean) : [];
  if (!list.length) return;

  // Individual continua baixando somente o QR daquele ISO.
  if (list.length === 1) {
    downloadSingleIsoQrItem(list[0]);
    return;
  }

  // Selecionados baixa uma arte Zebra consolidada por grupo de 3, igual ao que será impresso.
  const chunks = chunkIsoQrItemsForZebra(list, 3);
  for (let index = 0; index < chunks.length; index += 1) {
    const svg = await buildIsoQrZebraSheetSvg(chunks[index]);
    const link = document.createElement('a');
    link.href = isoQrDataUrlFromSvg(svg);
    const suffix = chunks.length > 1 ? `_${String(index + 1).padStart(2, '0')}` : '';
    link.download = `STEP_QR_Zebra_ZD230_${list.length}_selecionados${suffix}.svg`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    await new Promise((resolve) => setTimeout(resolve, 180));
  }
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
