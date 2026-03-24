// Run: node generate-sample.js
const fs = require('fs');

const statuses = ['Finished', 'Finished', 'Finished', 'Finished', 'Failed', 'Closed'];
const serviceTypes = ['Compras'];
const subserviceTypes = ['Compra en domicilio', 'Compra en tienda', 'Compra express'];
const providers = ['LogiExpress MX', 'EnvioRapido', 'TransMex', 'FleetGo', 'Rappi Logistics', 'MercadoEnvios'];
const drivers = [
  { name: 'Carlos Martinez', email: 'carlos.m@logistic.com' },
  { name: 'Ana Rodriguez', email: 'ana.r@logistic.com' },
  { name: 'Jorge Lopez', email: 'jorge.l@logistic.com' },
  { name: 'Maria Fernandez', email: 'maria.f@logistic.com' },
  { name: 'Pedro Sanchez', email: 'pedro.s@logistic.com' },
  { name: 'Laura Garcia', email: 'laura.g@logistic.com' },
  { name: 'Roberto Diaz', email: 'roberto.d@logistic.com' },
  { name: 'Sofia Herrera', email: 'sofia.h@logistic.com' },
  { name: 'Miguel Torres', email: 'miguel.t@logistic.com' },
  { name: 'Andrea Morales', email: 'andrea.m@logistic.com' },
  { name: 'Fernando Ruiz', email: 'fernando.r@logistic.com' },
  { name: 'Gabriela Castro', email: 'gabriela.c@logistic.com' },
];
const failedReasons = [
  'Cliente no disponible', 'Direccion incorrecta', 'Producto agotado',
  'Tienda cerrada', 'Zona de riesgo', 'Vehiculo averiado', 'Clima adverso',
];
const userEmails = ['admin@tas.com', 'ops1@tas.com', 'ops2@tas.com', 'supervisor@tas.com'];
const categories = [1, 2, 3, 4];

function rand(arr) { return arr[Math.floor(Math.random() * arr.length)]; }
function randInt(min, max) { return Math.floor(Math.random() * (max - min + 1)) + min; }

function addMinutes(d, m) {
  return new Date(d.getTime() + m * 60000);
}

function fmt(d) {
  if (!d) return '';
  return d.toISOString().replace('T', ' ').slice(0, 19);
}

const rows = [];
const startDate = new Date('2025-03-25');
const endDate = new Date('2025-12-31');
const daySpan = Math.floor((endDate - startDate) / 86400000);

for (let i = 1; i <= 2500; i++) {
  const dayOffset = randInt(0, daySpan);
  const baseDate = new Date(startDate.getTime() + dayOffset * 86400000);
  baseDate.setHours(randInt(6, 22), randInt(0, 59), 0);

  const status = rand(statuses);
  const driver = rand(drivers);
  const provider = rand(providers);

  const creationDate = new Date(baseDate);
  const logServiceDate = fmt(addMinutes(creationDate, randInt(0, 30)));
  const confirmedDate = fmt(addMinutes(creationDate, randInt(5, 60)));
  const startedDate = fmt(addMinutes(creationDate, randInt(30, 120)));
  const processDate = fmt(addMinutes(creationDate, randInt(60, 180)));

  let finalizedDate = '';
  let failedDate = '';
  let cancelReason = '';

  if (status === 'Finished' || status === 'Closed') {
    finalizedDate = fmt(addMinutes(creationDate, randInt(90, 360)));
  }
  if (status === 'Failed') {
    failedDate = fmt(addMinutes(creationDate, randInt(60, 240)));
    cancelReason = rand(failedReasons);
  }

  const shippedDate = fmt(addMinutes(creationDate, randInt(45, 150)));

  rows.push({
    service_id: 100000 + i,
    stock_id: 200000 + randInt(1, 5000),
    status,
    service_type: rand(serviceTypes),
    subservice_type: rand(subserviceTypes),
    los_log_service_date: logServiceDate,
    lsh_date: logServiceDate,
    origin_dep_id: randInt(1, 50),
    destiny_dep_id: randInt(1, 50),
    direccion_origen: `Calle ${randInt(1, 200)} #${randInt(1, 999)}, Col. Centro`,
    'dirección_destino': `Av. ${randInt(1, 100)} #${randInt(1, 500)}, Col. Norte`,
    finalized_date: finalizedDate,
    process_date: processDate,
    started_date: startedDate,
    confirmed_date: confirmedDate,
    failed_date: failedDate,
    Pre_cancelled_date: '',
    Ready_to_ship_date: '',
    shipped_date: shippedDate,
    creation_date: fmt(creationDate),
    user_created: rand(userEmails),
    user_confirmed: rand(userEmails),
    user_started: rand(userEmails),
    user_finalized: status === 'Finished' ? rand(userEmails) : '',
    user_cancelled: '',
    user_failed: status === 'Failed' ? rand(userEmails) : '',
    user_precancelled: '',
    los_url_qr: `https://example.com/qr/${100000 + i}`,
    categoria_driver: rand(categories),
    ste_provider_name: provider,
    driver: driver.name,
    email_driver: driver.email,
    cancel_or_failed_reason: cancelReason,
    filter_date: finalizedDate || failedDate || logServiceDate,
    attended_date: processDate || startedDate || finalizedDate || failedDate,
  });
}

const headers = Object.keys(rows[0]);
const csv = [
  headers.join(','),
  ...rows.map((r) =>
    headers.map((h) => {
      const v = String(r[h] ?? '');
      return v.includes(',') || v.includes('"') ? `"${v.replace(/"/g, '""')}"` : v;
    }).join(',')
  ),
].join('\n');

fs.writeFileSync('public/sample_data.csv', csv);
console.log(`Generated ${rows.length} rows -> public/sample_data.csv`);
