import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const { handler } = require('./scheduled-sync-tracking-cache.js');

export default async function scheduledCacheHeartbeat(request) {
  const headers = {};
  try {
    for (const [key, value] of request.headers.entries()) headers[key] = value;
  } catch {}

  const result = await handler({
    httpMethod: 'POST',
    headers,
    queryStringParameters: { scheduled: '1', source: 'scheduled-cache-heartbeat-modern' },
  });

  return new Response(result?.body || '', {
    status: result?.statusCode || 200,
    headers: result?.headers || { 'content-type': 'application/json', 'cache-control': 'no-store' },
  });
}

export const config = {
  schedule: '*/15 * * * *',
};
