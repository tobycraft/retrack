// HTTPS static server for docs/, used only for fast local iteration on the
// live-Word test (tests/real-office-e2e/). Word requires HTTPS even for a
// sideloaded local add-in, so this uses the trusted dev cert from
// office-addin-dev-certs (`npx office-addin-dev-certs install`).
//
// Serves with Cache-Control: no-store — Word's add-in WebView was found to
// aggressively cache commands.html across runs (a fix wasn't picked up
// until Word was fully quit and relaunched), which made iterating against
// the deployed GitHub Pages copy painfully slow. No-store keeps every test
// run honest about what's actually in docs/ right now.
const https = require('https');
const fs = require('fs');
const path = require('path');
const { getHttpsServerOptions } = require('office-addin-dev-certs');

const root = path.join(__dirname, '..', '..', 'docs');
const port = process.env.PORT || 3000;

const MIME = {
  '.html': 'text/html',
  '.js': 'application/javascript',
  '.css': 'text/css',
  '.png': 'image/png',
  '.xml': 'application/xml',
};

async function main() {
  const options = await getHttpsServerOptions();

  const server = https.createServer(options, (req, res) => {
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
      res.writeHead(200, {
        'Content-Type': MIME[path.extname(filePath)] || 'application/octet-stream',
        'Cache-Control': 'no-store',
      });
      res.end(data);
    });
  });

  server.listen(port, () => {
    console.log(`Local HTTPS server serving docs/ at https://localhost:${port}`);
  });
}

main();
