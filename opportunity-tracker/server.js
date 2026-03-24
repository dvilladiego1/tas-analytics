const express = require('express');
const multer = require('multer');
const { parse } = require('csv-parse/sync');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = 3000;
const DATA_FILE = path.join(__dirname, 'data.json');
const upload = multer({ storage: multer.memoryStorage() });

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

function readData() {
  if (!fs.existsSync(DATA_FILE)) return { opportunities: [], feedback: {}, documents: {} };
  return JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
}

function writeData(data) {
  fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2), 'utf8');
}

// Upload CSV
app.post('/api/upload', upload.single('file'), (req, res) => {
  try {
    const content = req.file.buffer.toString('utf8');
    const records = parse(content, { columns: true, skip_empty_lines: true, trim: true, bom: true });

    const opportunities = records.map((r, i) => ({
      id: r['Transaction Id'] || `opp-${i}`,
      eventId: r['EventIdSyncOlimpo'] || '',
      prospecto: r['Prospecto'] || '',
      correo: r['Correo'] || '',
      telefono: r['Telefono'] || '',
      marca: r['Marca'] || '',
      modelo: r['Modelo'] || '',
      ano: r['Año Modelo'] || r['Ano Modelo'] || '',
      vin: r['Auto Cotizado: VIN'] || '',
      etapa: r['Etapa'] || '',
      tipoOferta: r['Tipo de Oferta'] || '',
      fechaInspeccion: r['Fecha Inspección'] || r['Fecha Inspeccion'] || '',
      agencia: r['Agencia'] || '',
      ofertaInicial: r['Oferta Inicial'] || '',
      ofertaFinal: r['Oferta Final'] || '',
      pctComision: r['% Comisión por Grupo'] || r['% Comision por Grupo'] || '',
      montoComision: r['$ Comisión por Grupo'] || r['$ Comision por Grupo'] || '',
      estatus: r['Estatus de la Oportunidad'] || 'En proceso',
      asesor: r['Asesor del Comentario'] || '',
      cotizadoPor: r['Cotizado por'] || ''
    }));

    const data = readData();
    data.opportunities = opportunities;
    // preserve existing feedback
    if (!data.feedback) data.feedback = {};
    writeData(data);

    res.json({ ok: true, count: opportunities.length });
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
});

// Get all opportunities
app.get('/api/opportunities', (req, res) => {
  const data = readData();
  res.json(data);
});

// Add feedback to an opportunity
app.post('/api/feedback', (req, res) => {
  const { opportunityId, role, author, comment } = req.body;
  if (!opportunityId || !role || !comment) {
    return res.status(400).json({ error: 'Missing fields' });
  }

  const data = readData();
  if (!data.feedback[opportunityId]) data.feedback[opportunityId] = [];
  data.feedback[opportunityId].push({
    role,
    author: author || role,
    comment,
    timestamp: new Date().toISOString()
  });
  writeData(data);
  res.json({ ok: true });
});

// Update opportunity status (kavak only)
app.patch('/api/status/:id', (req, res) => {
  const { estatus } = req.body;
  const data = readData();
  const opp = data.opportunities.find(o => o.id === req.params.id);
  if (!opp) return res.status(404).json({ error: 'Not found' });
  opp.estatus = estatus;
  writeData(data);
  res.json({ ok: true });
});

// Update document checklist for an opportunity
app.patch('/api/documents/:id', (req, res) => {
  const { documents } = req.body;
  const data = readData();
  if (!data.documents) data.documents = {};
  data.documents[req.params.id] = documents;
  writeData(data);
  res.json({ ok: true });
});

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
