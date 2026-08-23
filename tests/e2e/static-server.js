// Zero-dependency static file server for docs/, used only to serve the
// add-in to Playwright during tests. Not part of the deployed site.
const http = require('http');
const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..', '..', 'docs');
const port = process.env.PORT || 4173;

const MIME = {
  '.html': 'text/html',
  '.js': 'application/javascript',
  '.css': 'text/css',
  '.png': 'image/png',
};

const server = http.createServer((req, res) => {
  const urlPath = decodeURIComponent(req.url.split('?')[0]);
  const relative = urlPath === '/' ? '/commands.html' : urlPath;
  const filePath = path.join(root, relative);

  if (!filePath.startsWith(root)) {
    res.writeHead(403);
    res.end('Forbidden');
    return;
  }

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404);
      res.end('Not found');
      return;
    }
    res.writeHead(200, { 'Content-Type': MIME[path.extname(filePath)] || 'application/octet-stream' });
    res.end(data);
  });
});

server.listen(port, () => {
  console.log(`Static server serving docs/ at http://127.0.0.1:${port}`);
});
