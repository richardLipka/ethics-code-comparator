const http = require('http');
const fs = require('fs');
const path = require('path');

// Port: `node serve.js 8080`, or PORT=8080 node serve.js (default 3000)
const PORT = parseInt(process.argv[2] || process.env.PORT || '3000', 10);
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
  let reqPath;
  try {
    reqPath = decodeURIComponent(req.url.split('?')[0]);
  } catch (e) {
    send(res, 400, '400 Bad Request');
    return;
  }
  if (reqPath === '/' || reqPath === '') reqPath = '/index.html';

  // Never serve anything outside the project folder
  let fullPath = path.resolve(ROOT, '.' + path.sep + reqPath);
  if (fullPath !== ROOT && !fullPath.startsWith(ROOT + path.sep)) {
    send(res, 403, '403 Forbidden');
    return;
  }

  fs.stat(fullPath, (err, stat) => {
    if (!err && stat.isDirectory()) {
      fullPath = path.join(fullPath, 'index.html');
      stat = fs.existsSync(fullPath) ? fs.statSync(fullPath) : null;
    }
    if (err || !stat || !stat.isFile()) {
      send(res, 404, '404 Not Found');
      return;
    }

    const ext = path.extname(fullPath).toLowerCase();
    const mime = MIME_TYPES[ext] || 'application/octet-stream';
    res.writeHead(200, { 'Content-Type': mime, 'Cache-Control': 'no-cache' });
    fs.createReadStream(fullPath).pipe(res);
  });
});

server.on('error', (err) => {
  if (err.code === 'EADDRINUSE') {
    console.error(`Port ${PORT} je obsazený. Spusťte server na jiném portu: node serve.js 8080`);
    process.exit(1);
  }
  throw err;
});

server.listen(PORT, () => {
  console.log(`Server běží na http://localhost:${PORT}`);
});
