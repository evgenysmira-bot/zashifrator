const https = require('https');
const fs = require('fs');
const path = require('path');
const url = require('url');
const { getHttpsServerOptions } = require('office-addin-dev-certs');

const PORT = 3000;
const ROOT = __dirname;

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js':   'application/javascript; charset=utf-8',
  '.css':  'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.xml':  'application/xml; charset=utf-8',
  '.png':  'image/png',
  '.jpg':  'image/jpeg',
  '.svg':  'image/svg+xml',
  '.ico':  'image/x-icon',
};

(async () => {
  const opts = await getHttpsServerOptions();
  const server = https.createServer(opts, (req, res) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Cache-Control', 'no-cache');

    const urlPath = decodeURIComponent(url.parse(req.url).pathname);
    const safePath = path.normalize(urlPath).replace(/^(\.\.[\/\\])+/, '');
    const filePath = path.join(ROOT, safePath === '/' ? '/index.html' : safePath);

    if (!filePath.startsWith(ROOT)) {
      res.writeHead(403); res.end('Forbidden'); return;
    }

    fs.readFile(filePath, (err, data) => {
      if (err) {
        console.log('404', urlPath);
        res.writeHead(404);
        res.end('Not found: ' + urlPath);
        return;
      }
      const ext = path.extname(filePath).toLowerCase();
      res.writeHead(200, { 'Content-Type': MIME[ext] || 'application/octet-stream' });
      res.end(data);
    });
  });

  server.listen(PORT, () => {
    console.log(`Dev server: https://localhost:${PORT}`);
    console.log(`Manifest:   https://localhost:${PORT}/manifest.xml`);
    console.log(`Task pane:  https://localhost:${PORT}/src/taskpane.html`);
  });
})().catch(err => {
  console.error('Failed to start server:', err);
  console.error('Run "npm run install-certs" first to install local HTTPS certificates.');
  process.exit(1);
});
