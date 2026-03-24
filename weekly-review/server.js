const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = 5174;
const FILE = path.join('/Users/danielvilladiego/Downloads/TAS_Weekly_Review_18Mar2026.html');

http.createServer((req, res) => {
  fs.readFile(FILE, (err, data) => {
    if (err) { res.writeHead(500); res.end('Error'); return; }
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(data);
  });
}).listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));
