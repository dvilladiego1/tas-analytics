const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = 8081;
const ROOT = __dirname;

http.createServer((req, res) => {
  let filePath = path.join(ROOT, req.url === '/' ? '/TAS_Pitch_Deck.html' : req.url);
  const ext = path.extname(filePath);
  const mime = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css' };
  fs.readFile(filePath, (err, data) => {
    if (err) { res.writeHead(404); res.end('Not found'); return; }
    res.writeHead(200, { 'Content-Type': mime[ext] || 'text/plain' });
    res.end(data);
  });
}).listen(PORT, () => console.log(`Serving on http://localhost:${PORT}`));
