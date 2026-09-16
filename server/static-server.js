// Dependency-free static file server for local use.
// Serves server/public without needing MongoDB, so the browser-side app
// (Packing List Builder, Match Helper, ...) can be exercised locally.
// The config API lives in server.js and still requires MONGO_URI.
const http  = require('http');
const https = require('https');
const fs    = require('fs');
const path  = require('path');

const PORT = process.env.PORT || 4173;
const ROOT = path.join(__dirname, 'public');

// Pages such as the Container Match Helper call same-origin /api paths.
// Without MongoDB credentials locally, forward them to the deployed backend
// (the same one js/db.js targets by default). Override with API_TARGET, or
// set it to a local `npm start` instance once MONGO_URI is available.
const API_TARGET = process.env.API_TARGET
  || 'https://pl-builder-helper-server-production.up.railway.app';

const TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js':   'text/javascript; charset=utf-8',
  '.css':  'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg':  'image/svg+xml',
  '.ico':  'image/x-icon',
  '.png':  'image/png',
};

function proxyApi(req, res) {
  const target = new URL(req.url, API_TARGET);
  const agent  = target.protocol === 'https:' ? https : http;
  const headers = { ...req.headers, host: target.host };
  delete headers['accept-encoding'];

  const up = agent.request(target, { method: req.method, headers }, upRes => {
    res.writeHead(upRes.statusCode, upRes.headers);
    upRes.pipe(res);
  });
  up.on('error', err => {
    res.writeHead(502, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({ error: `API proxy to ${API_TARGET} failed: ${err.message}` }));
  });
  req.pipe(up);
}

http.createServer((req, res) => {
  const urlPath = decodeURIComponent(req.url.split('?')[0]);

  if (urlPath.startsWith('/api/')) { proxyApi(req, res); return; }

  const rel     = urlPath === '/' ? 'index.html' : urlPath.replace(/^\/+/, '');
  const file    = path.join(ROOT, rel);

  // Refuse anything resolving outside the public directory.
  if (!file.startsWith(ROOT + path.sep) && file !== path.join(ROOT, 'index.html')) {
    res.writeHead(403).end('Forbidden');
    return;
  }

  fs.readFile(file, (err, buf) => {
    if (err) {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('Not found: ' + rel);
      return;
    }
    res.writeHead(200, { 'Content-Type': TYPES[path.extname(file).toLowerCase()] || 'application/octet-stream' });
    res.end(buf);
  });
}).listen(PORT, () => console.log(`Static server on http://localhost:${PORT}`));
