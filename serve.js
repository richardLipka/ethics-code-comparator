const http = require('http');
const fs = require('fs');
const path = require('path');

// Port: `node serve.js 8080`, or PORT=8080 node serve.js (default 3000)
const PORT = parseInt(process.argv[2] || process.env.PORT || '3000', 10);
// Only this computer can connect; HOST=0.0.0.0 makes the app reachable from the network
const HOST = process.env.HOST || '127.0.0.1';
const ROOT = path.resolve(__dirname);
const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
  '.md': 'text/markdown; charset=utf-8',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
};

function send(res, status, text) {
  res.writeHead(status, { 'Content-Type': 'text/plain; charset=utf-8' });
  res.end(text);
}

const server = http.createServer((req, res) => {
  let pathname;
  try {
    pathname = decodeURIComponent(new URL(req.url, 'http://localhost').pathname);
  } catch (e) {
    send(res, 400, '400 Bad Request');
    return;
  }

  // Normalize path and map root to index.html
  let safePath = path.normalize(pathname).replace(/^(\.\.[/\\])+/, '');
  if (safePath === '/' || safePath === '\\' || safePath === '' || safePath === '.') {
    safePath = '/index.html';
  }

  let fullPath = path.join(ROOT, safePath);

  // Security check: never serve anything outside project folder, nor hidden files (.git, …)
  const rel = path.relative(ROOT, fullPath);
  if (rel.startsWith('..') || path.isAbsolute(rel) || rel.split(path.sep).some((part) => part.startsWith('.'))) {
    send(res, 403, '403 Forbidden');
    return;
  }

  fs.stat(fullPath, (err, stat) => {
    if (err) {
      send(res, 404, '404 Not Found');
      return;
    }

    if (stat.isDirectory()) {
      fullPath = path.join(fullPath, 'index.html');
      try {
        stat = fs.statSync(fullPath);
      } catch (e) {
        send(res, 404, '404 Not Found');
        return;
      }
    }

    if (!stat || !stat.isFile()) {
      send(res, 404, '404 Not Found');
      return;
    }

    const ext = path.extname(fullPath).toLowerCase();
    const mime = MIME_TYPES[ext] || 'application/octet-stream';
    res.writeHead(200, { 'Content-Type': mime, 'Cache-Control': 'no-cache', 'X-Content-Type-Options': 'nosniff' });

    const stream = fs.createReadStream(fullPath);
    stream.on('error', (streamErr) => {
      console.error('File stream error:', streamErr);
      if (!res.headersSent) {
        send(res, 500, '500 Server Error');
      } else {
        res.destroy();
      }
    });
    stream.pipe(res);
  });
});

server.on('error', (err) => {
  if (err.code === 'EADDRINUSE') {
    console.error(`Port ${PORT} je obsazený. Spusťte server na jiném portu: node serve.js 8080`);
    process.exit(1);
  }
  throw err;
});

server.listen(PORT, HOST, () => {
  console.log(`Server běží na http://localhost:${PORT}`);
});
