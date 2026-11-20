import { readFile } from 'node:fs/promises';
import { createServer } from 'node:http';

const clientId = process.env.MICROSOFT_CLIENT_ID;
if (!/^[\da-f]{8}(?:-[\da-f]{4}){3}-[\da-f]{12}$/i.test(clientId ?? '')) {
  throw new Error('Set MICROSOFT_CLIENT_ID to the Application (client) ID of your Microsoft SPA registration.');
}

const js = 'text/javascript; charset=utf-8';
const html = 'text/html; charset=utf-8';
const json = 'application/json; charset=utf-8';
const files = [
  ['/', html, 'public/index.html'],
  ['/style.css', 'text/css; charset=utf-8', 'public/style.css'],
  ['/app.js', js, 'public/app.js'],
  ['/export.js', js, 'public/export.js'],
  ['/auth.html', html, 'public/auth.html'],
  ['/auth.js', js, 'public/auth.js'],
  ['/msal.js', js, 'node_modules/@azure/msal-browser/lib/msal-browser.min.js'],
  ['/msal-bridge.js', js, 'node_modules/@azure/msal-browser/lib/redirect-bridge/msal-redirect-bridge.min.js'],
];
const routes = new Map(await Promise.all(files.map(async ([path, type, file]) => [
  path, { type, body: await readFile(new URL(file, import.meta.url)) },
])));
routes.set('/config.json', { type: json, body: JSON.stringify({ clientId }) });
routes.set('/healthz', { type: 'text/plain; charset=utf-8', body: 'OK' });

createServer((request, response) => {
  // Use only fixed routes; never resolve request paths to files or process request bodies.
  const path = request.url.split('?', 1)[0];
  response.setHeader('Cache-Control', 'no-store');
  response.setHeader('Referrer-Policy', 'no-referrer');
  response.setHeader('X-Content-Type-Options', 'nosniff');
  response.setHeader('Content-Security-Policy', [
    "default-src 'none'",
    "script-src 'self'",
    "style-src 'self'",
    "connect-src 'self' https://login.microsoftonline.com https://graph.microsoft.com",
    "frame-src 'self' https://login.microsoftonline.com https://login.live.com",
    path === '/auth.html' ? "frame-ancestors 'self'" : "frame-ancestors 'none'",
    "base-uri 'none'",
    "form-action 'none'",
  ].join('; '));

  if (request.method !== 'GET' && request.method !== 'HEAD') {
    response.writeHead(405, { Allow: 'GET, HEAD' }).end();
    return;
  }
  const asset = routes.get(path);
  if (!asset) {
    response.writeHead(404).end();
    return;
  }
  response.writeHead(200, {
    'Content-Type': asset.type,
    'Content-Length': Buffer.byteLength(asset.body),
  }).end(request.method === 'HEAD' ? undefined : asset.body);
}).listen(Number(process.env.PORT ?? 8080), '0.0.0.0');
