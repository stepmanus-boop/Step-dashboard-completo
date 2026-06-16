const { getIsoQrByToken } = require('./_isoQrCodes');

function textResponse(statusCode, body, contentType = 'text/plain; charset=utf-8', extraHeaders = {}) {
  return {
    statusCode,
    headers: {
      'content-type': contentType,
      'cache-control': 'no-store',
      ...extraHeaders,
    },
    body,
  };
}

function safeFilename(value = '') {
  return String(value || 'iso-qrcode')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9._-]+/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '')
    .slice(0, 120) || 'iso-qrcode';
}

exports.handler = async (event = {}) => {
  if (event.httpMethod && event.httpMethod !== 'GET') {
    return textResponse(405, 'Método não permitido.');
  }

  const query = event.queryStringParameters || {};
  const token = String(query.token || '').trim();
  if (!token) return textResponse(400, 'Token do QR Code não informado.');

  try {
    const item = await getIsoQrByToken(token);
    if (!item) return textResponse(404, 'QR Code não encontrado.');
    const QRCode = require('qrcode');
    const qrPayload = item.qrUrl || `${event.headers?.['x-forwarded-proto'] || 'https'}://${event.headers?.host || ''}/qr-tracking.html?token=${encodeURIComponent(token)}`;
    const width = Math.min(1200, Math.max(160, Number(query.w || query.width || 420)));
    const svg = await QRCode.toString(qrPayload, {
      type: 'svg',
      errorCorrectionLevel: 'M',
      margin: 2,
      width,
      color: { dark: '#000000', light: '#FFFFFF' },
    });
    const headers = {};
    if (String(query.download || '') === '1') {
      headers['content-disposition'] = `attachment; filename="${safeFilename(item.isoFullName || item.iso)}.svg"`;
    }
    return textResponse(200, svg, 'image/svg+xml; charset=utf-8', headers);
  } catch (error) {
    return textResponse(500, error.message || 'Falha ao gerar imagem do QR Code.');
  }
};
