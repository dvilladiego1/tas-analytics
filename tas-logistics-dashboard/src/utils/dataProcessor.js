import Papa from 'papaparse';
import { format, differenceInMinutes, differenceInHours, differenceInCalendarDays, subDays, startOfDay, endOfDay } from 'date-fns';

export function parseCSV(file, skipFirstRow = false) {
  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: !skipFirstRow,
      skipEmptyLines: true,
      dynamicTyping: false,
      complete: (results) => {
        if (skipFirstRow) {
          // First row was metadata, second row is headers, rest is data
          const rows = results.data;
          if (rows.length < 2) return resolve([]);
          const headers = rows[1];
          const data = rows.slice(2).map((row) => {
            const obj = {};
            headers.forEach((h, i) => { obj[h] = row[i] || ''; });
            return obj;
          });
          resolve(data);
        } else {
          resolve(results.data);
        }
      },
      error: (error) => reject(error),
    });
  });
}

// Parse dates in multiple formats: "21/02/2026, 16:22" or "23/2/2026 13:25:15" or "21/02/2026"
function safeDate(val) {
  if (!val || val === '' || val === 'null' || val === 'NULL' || val === '[no field found]') return null;
  try {
    // Try "DD/MM/YYYY, HH:MM" (BBDD format)
    let m = val.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4}),?\s*(\d{1,2}):(\d{2})(?::(\d{2}))?/);
    if (m) {
      const d = new Date(+m[3], +m[2] - 1, +m[1], +m[4], +m[5], m[6] ? +m[6] : 0);
      return isNaN(d.getTime()) ? null : d;
    }
    // Try "DD/MM/YYYY" only
    m = val.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) {
      const d = new Date(+m[3], +m[2] - 1, +m[1]);
      return isNaN(d.getTime()) ? null : d;
    }
    // Fallback
    const d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
  } catch {
    return null;
  }
}

// Normalize car brand names (inconsistent casing in Jira data)
function normalizeBrand(tag) {
  if (!tag || tag === '[no field found]') return null;
  const map = {
    kia: 'KIA', chevrolet: 'CHEVROLET', nissan: 'NISSAN', mazda: 'MAZDA',
    volkswagen: 'VOLKSWAGEN', hyundai: 'HYUNDAI', toyota: 'TOYOTA', honda: 'HONDA',
    ford: 'FORD', mg: 'MG', bmw: 'BMW', seat: 'SEAT', audi: 'AUDI',
    'mercedes-benz': 'MERCEDES-BENZ', mercedes: 'MERCEDES-BENZ', volvo: 'VOLVO',
    changan: 'CHANGAN', chirey: 'CHIREY', byd: 'BYD', gwm: 'GWM', omoda: 'OMODA',
    gac: 'GAC', renault: 'RENAULT', peugeot: 'PEUGEOT', suzuki: 'SUZUKI',
    subaru: 'SUBARU', jeep: 'JEEP', dodge: 'DODGE', ram: 'RAM', fiat: 'FIAT',
    mini: 'MINI', mitsubishi: 'MITSUBISHI', buick: 'BUICK', cadillac: 'CADILLAC',
    lincoln: 'LINCOLN', jaguar: 'JAGUAR', 'land rover': 'LAND ROVER', porsche: 'PORSCHE',
    acura: 'ACURA', infiniti: 'INFINITI', genesis: 'GENESIS',
  };
  const lower = tag.trim().toLowerCase();
  return map[lower] || tag.trim().toUpperCase();
}

// ============ BBDD Processing ============
export function processBBDD(rawData) {
  // Skip metadata row if present (line 1 has "Request status:")
  const data = rawData.filter((r) => r['Stock Id'] && r['Stock Id'] !== '');
  return data.map((row) => ({
    stockId: row['Stock Id']?.trim(),
    opportunity: row['Nombre de la oportunidad']?.trim() || '',
    poDate: safeDate(row['Fecha y Hora PO']),
    origin: row['Lugar origen de servicio logístico: Nombre']?.trim() || '',
    transactionId: row['Transaction Id']?.trim() || '',
    etapa: row['Etapa']?.trim() || '',
    destination: row['Lugar destino servicio logístico: Nombre']?.trim() || '',
    orderDate: safeDate(row['Fecha inicial del pedido']),
    estado: row['Estado']?.trim() || '',
    agent: row['Agente que realiza la compra: Nombre completo']?.trim() || '',
  }));
}

// ============ Jira Processing ============
export function processJira(rawData) {
  return rawData
    .filter((r) => r['Key'] && r['Key'].startsWith('LOG-'))
    .map((row) => ({
      key: row['Key']?.trim(),
      requestType: row['Request Type']?.trim() || '',
      summary: row['Summary']?.trim() || '',
      status: row['Status']?.trim() || '',
      reporter: row['Reporter']?.trim() || '',
      assignee: row['Assignee']?.trim() || '',
      created: safeDate(row['Created']),
      updated: safeDate(row['Updated']),
      resolved: safeDate(row['Resolved']),
      pickupDate: safeDate(row['Kavak Pickup Date']),
      pickupHour: row['Kavak Dropdown Hour (30 min)']?.trim() || '',
      region: row['Mexico Region Hub']?.trim() || '',
      stockId: row['Stock ID']?.trim() || '',
      brand: normalizeBrand(row['App Tag']),
      carModel: row['To Kavak Center']?.trim() || '',
      tramites: row['Kavak Requested Tramites']?.trim() || '',
      address: row['Kavak Google Maps Address - Long']?.trim() || '',
    }));
}

// ============ Publish CSV Processing ============
export function processPublish(rawData) {
  return rawData
    .filter((r) => r['stock_id'] && r['stock_id'] !== '')
    .map((row) => ({
      stockId: row['stock_id']?.trim(),
      publishDate: safeDate(row['publication_created_date_local']),
      contractType: row['contract_type']?.trim() || '',
      stockStatus: row['stock_status']?.trim() || '',
      publicationStatus: row['publication_status']?.trim() || '',
      publishRegion: row['region_final_compra']?.trim() || '',
      publishSite: row['site_final_compra']?.trim() || '',
    }));
}

// ============ Join datasets by Stock ID ============
export function joinData(bbddData, jiraData, publishData = []) {
  const jiraMap = new Map();
  jiraData.forEach((j) => {
    if (j.stockId) jiraMap.set(j.stockId, j);
  });

  const publishMap = new Map();
  publishData.forEach((p) => {
    if (p.stockId) publishMap.set(p.stockId, p);
  });

  return bbddData.map((b) => {
    const jira = jiraMap.get(b.stockId) || null;
    const pub = publishMap.get(b.stockId) || null;
    return {
      ...b,
      jiraKey: jira?.key || null,
      jiraStatus: jira?.status || 'Sin ticket',
      jiraReporter: jira?.reporter || '',
      jiraAssignee: jira?.assignee || '',
      jiraCreated: jira?.created || null,
      jiraResolved: jira?.resolved || null,
      jiraPickupDate: jira?.pickupDate || null,
      jiraPickupHour: jira?.pickupHour || '',
      jiraRegion: jira?.region || '',
      jiraBrand: jira?.brand || null,
      jiraRequestType: jira?.requestType || '',
      hasJiraTicket: !!jira,
      publishDate: pub?.publishDate || null,
      publishContractType: pub?.contractType || '',
      publishStockStatus: pub?.stockStatus || '',
      publicationStatus: pub?.publicationStatus || '',
      publishRegion: pub?.publishRegion || '',
      publishSite: pub?.publishSite || '',
      hasPublishDate: !!pub?.publishDate,
    };
  });
}

// ============ Filters ============
export function filterData(data, { startDate, endDate, status, region, agent }) {
  let filtered = data;
  if (startDate) filtered = filtered.filter((r) => r.poDate && r.poDate >= new Date(startDate));
  if (endDate) filtered = filtered.filter((r) => r.poDate && r.poDate <= new Date(endDate + 'T23:59:59'));
  if (status && status !== 'all') filtered = filtered.filter((r) => r.jiraStatus === status);
  if (region && region !== 'all') filtered = filtered.filter((r) => r.jiraRegion === region);
  if (agent && agent !== 'all') filtered = filtered.filter((r) => r.agent === agent);
  return filtered;
}

// ============ KPIs ============
export function getKPIs(data) {
  const total = data.length;
  const withTicket = data.filter((r) => r.hasJiraTicket).length;
  const done = data.filter((r) => r.jiraStatus === 'DONE').length;
  const todo = data.filter((r) => r.jiraStatus === 'TO DO').length;
  const canceled = data.filter((r) => r.jiraStatus === 'CANCELED').length;
  const pagada = data.filter((r) => r.estado === 'Pagada').length;
  const abierta = data.filter((r) => r.estado === 'Abierta').length;
  const matchRate = total > 0 ? ((withTicket / total) * 100).toFixed(1) : '0.0';
  const doneRate = withTicket > 0 ? ((done / withTicket) * 100).toFixed(1) : '0.0';

  return { total, withTicket, done, todo, canceled, pagada, abierta, matchRate, doneRate };
}

// ============ Jira Status Breakdown ============
export function getStatusBreakdown(data) {
  const counts = {};
  data.forEach((r) => {
    const s = r.jiraStatus || 'Sin ticket';
    counts[s] = (counts[s] || 0) + 1;
  });
  return Object.entries(counts).map(([name, value]) => ({ name, value }));
}

// ============ Estado (Pagada/Abierta) Breakdown ============
export function getEstadoBreakdown(data) {
  const counts = {};
  data.forEach((r) => {
    const e = r.estado || 'Sin estado';
    counts[e] = (counts[e] || 0) + 1;
  });
  return Object.entries(counts).map(([name, value]) => ({ name, value }));
}

// ============ Daily Volume ============
export function getDailyVolume(data) {
  const daily = {};
  data.forEach((row) => {
    const d = row.poDate;
    if (!d) return;
    const key = format(d, 'yyyy-MM-dd');
    if (!daily[key]) daily[key] = { date: key, DONE: 0, 'TO DO': 0, CANCELED: 0, 'Sin ticket': 0, Total: 0 };
    daily[key].Total += 1;
    const s = row.jiraStatus || 'Sin ticket';
    if (daily[key][s] !== undefined) daily[key][s] += 1;
  });
  return Object.values(daily).sort((a, b) => a.date.localeCompare(b.date));
}

// ============ Region Breakdown ============
export function getRegionBreakdown(data) {
  const counts = {};
  data.forEach((r) => {
    const region = r.jiraRegion || 'Sin region';
    if (!counts[region]) counts[region] = { name: region, DONE: 0, 'TO DO': 0, CANCELED: 0, Total: 0 };
    counts[region].Total += 1;
    const s = r.jiraStatus;
    if (s === 'DONE') counts[region].DONE += 1;
    else if (s === 'TO DO') counts[region]['TO DO'] += 1;
    else if (s === 'CANCELED') counts[region].CANCELED += 1;
  });
  return Object.values(counts).sort((a, b) => b.Total - a.Total);
}

// ============ Car Brand Breakdown ============
export function getBrandBreakdown(data) {
  const counts = {};
  data.forEach((r) => {
    if (!r.jiraBrand) return;
    counts[r.jiraBrand] = (counts[r.jiraBrand] || 0) + 1;
  });
  return Object.entries(counts)
    .map(([name, value]) => ({ name, value }))
    .sort((a, b) => b.value - a.value);
}

// ============ Request Type Breakdown ============
export function getRequestTypeBreakdown(data) {
  const counts = {};
  data.filter((r) => r.jiraRequestType).forEach((r) => {
    counts[r.jiraRequestType] = (counts[r.jiraRequestType] || 0) + 1;
  });
  return Object.entries(counts)
    .map(([name, value]) => ({ name, value }))
    .sort((a, b) => b.value - a.value);
}

// ============ Agent Performance ============
export function getAgentPerformance(data) {
  const agents = {};
  data.forEach((r) => {
    const name = r.agent || 'Sin agente';
    if (!agents[name]) {
      agents[name] = {
        agent: name,
        total: 0,
        pagada: 0,
        abierta: 0,
        withTicket: 0,
        done: 0,
        todo: 0,
        canceled: 0,
        resolutionMinutes: [],
      };
    }
    agents[name].total += 1;
    if (r.estado === 'Pagada') agents[name].pagada += 1;
    if (r.estado === 'Abierta') agents[name].abierta += 1;
    if (r.hasJiraTicket) agents[name].withTicket += 1;
    if (r.jiraStatus === 'DONE') agents[name].done += 1;
    if (r.jiraStatus === 'TO DO') agents[name].todo += 1;
    if (r.jiraStatus === 'CANCELED') agents[name].canceled += 1;

    if (r.jiraCreated && r.jiraResolved) {
      const mins = differenceInMinutes(r.jiraResolved, r.jiraCreated);
      if (mins > 0 && mins < 43200) agents[name].resolutionMinutes.push(mins);
    }
  });

  return Object.values(agents)
    .map((a) => ({
      ...a,
      ticketRate: a.total > 0 ? ((a.withTicket / a.total) * 100).toFixed(0) : '-',
      doneRate: a.withTicket > 0 ? ((a.done / a.withTicket) * 100).toFixed(0) : '-',
      avgResolution:
        a.resolutionMinutes.length > 0
          ? Math.round(a.resolutionMinutes.reduce((x, y) => x + y, 0) / a.resolutionMinutes.length)
          : null,
    }))
    .sort((a, b) => b.total - a.total);
}

// ============ Time Analysis (Jira Created → Resolved) ============
export function getTimeAnalysis(data) {
  const createdToResolved = [];
  const poToJiraCreated = [];

  data.forEach((r) => {
    if (r.jiraCreated && r.jiraResolved) {
      const mins = differenceInMinutes(r.jiraResolved, r.jiraCreated);
      if (mins > 0 && mins < 43200) createdToResolved.push(mins);
    }
    if (r.poDate && r.jiraCreated) {
      const mins = differenceInMinutes(r.jiraCreated, r.poDate);
      if (mins > -1440 && mins < 43200) poToJiraCreated.push(Math.max(0, mins));
    }
  });

  const avg = (arr) => (arr.length > 0 ? Math.round(arr.reduce((a, b) => a + b, 0) / arr.length) : null);
  const median = (arr) => {
    if (arr.length === 0) return null;
    const sorted = [...arr].sort((a, b) => a - b);
    const mid = Math.floor(sorted.length / 2);
    return sorted.length % 2 !== 0 ? sorted[mid] : Math.round((sorted[mid - 1] + sorted[mid]) / 2);
  };

  return {
    createdToResolved: { avg: avg(createdToResolved), median: median(createdToResolved), count: createdToResolved.length },
    poToJiraCreated: { avg: avg(poToJiraCreated), median: median(poToJiraCreated), count: poToJiraCreated.length },
  };
}

export function formatMinutes(mins) {
  if (mins === null || mins === undefined) return '-';
  if (mins < 60) return `${mins} min`;
  const h = Math.floor(mins / 60);
  const m = mins % 60;
  if (h >= 24) {
    const d = Math.floor(h / 24);
    const rh = h % 24;
    return rh > 0 ? `${d}d ${rh}h` : `${d}d`;
  }
  return m > 0 ? `${h}h ${m}m` : `${h}h`;
}

// ============ Hourly Distribution (PO time) ============
export function getHourlyDistribution(data) {
  const hours = Array.from({ length: 24 }, (_, i) => ({ hour: `${String(i).padStart(2, '0')}:00`, count: 0 }));
  data.forEach((r) => {
    if (r.poDate) hours[r.poDate.getHours()].count += 1;
  });
  return hours;
}

// ============ Pickup Hour Distribution (scheduled) ============
export function getPickupHourDistribution(data) {
  const counts = {};
  data.forEach((r) => {
    if (r.jiraPickupHour) {
      counts[r.jiraPickupHour] = (counts[r.jiraPickupHour] || 0) + 1;
    }
  });
  return Object.entries(counts)
    .map(([name, value]) => ({ name, value }))
    .sort((a, b) => {
      const timeA = a.name.replace(/\s*(AM|PM)/, ' $1');
      const timeB = b.name.replace(/\s*(AM|PM)/, ' $1');
      return timeA.localeCompare(timeB);
    });
}

// ============ Destination Breakdown ============
export function getDestinationBreakdown(data) {
  const counts = {};
  data.forEach((r) => {
    if (r.destination) counts[r.destination] = (counts[r.destination] || 0) + 1;
  });
  return Object.entries(counts)
    .map(([name, value]) => ({ name, value }))
    .sort((a, b) => b.value - a.value);
}

// ============ Funnel Data ============
export function getFunnelData(data) {
  const total = data.length;
  const pagada = data.filter((r) => r.estado === 'Pagada').length;
  const withTicket = data.filter((r) => r.hasJiraTicket).length;
  const done = data.filter((r) => r.jiraStatus === 'DONE').length;
  return [
    { label: 'Cotizaciones (PO)', value: total, color: '#6366f1' },
    { label: 'Pagadas', value: pagada, color: '#3b82f6' },
    { label: 'Con Ticket Jira', value: withTicket, color: '#10b981' },
    { label: 'Completadas (DONE)', value: done, color: '#f59e0b' },
  ];
}

// ============ Route Breakdown ============
export function getRouteBreakdown(data) {
  const routes = {};
  data.forEach((r) => {
    const origin = r.origin || 'Sin origen';
    const dest = r.destination || 'Sin destino';
    const route = `${origin} > ${dest}`;
    if (!routes[route]) {
      routes[route] = { route, origin, destination: dest, total: 0, done: 0, todo: 0, canceled: 0, resolutionMinutes: [] };
    }
    routes[route].total += 1;
    if (r.jiraStatus === 'DONE') routes[route].done += 1;
    else if (r.jiraStatus === 'TO DO') routes[route].todo += 1;
    else if (r.jiraStatus === 'CANCELED') routes[route].canceled += 1;
    if (r.jiraCreated && r.jiraResolved) {
      const mins = differenceInMinutes(r.jiraResolved, r.jiraCreated);
      if (mins > 0 && mins < 43200) routes[route].resolutionMinutes.push(mins);
    }
  });
  return Object.values(routes)
    .map((r) => ({
      ...r,
      doneRate: r.total > 0 ? ((r.done / r.total) * 100).toFixed(0) : '-',
      avgResolution: r.resolutionMinutes.length > 0
        ? Math.round(r.resolutionMinutes.reduce((a, b) => a + b, 0) / r.resolutionMinutes.length)
        : null,
    }))
    .sort((a, b) => b.total - a.total);
}

// ============ SLA Distribution ============
export function getSLADistribution(data) {
  const buckets = [
    { label: '0-1 dia', min: 0, max: 1440, count: 0 },
    { label: '1-3 dias', min: 1440, max: 4320, count: 0 },
    { label: '3-5 dias', min: 4320, max: 7200, count: 0 },
    { label: '5-7 dias', min: 7200, max: 10080, count: 0 },
    { label: '7-14 dias', min: 10080, max: 20160, count: 0 },
    { label: '14+ dias', min: 20160, max: Infinity, count: 0 },
  ];
  const TARGET_MINUTES = 10080; // 7 days
  let withinSLA = 0;
  let totalMeasured = 0;
  let hasPublishData = false;

  data.forEach((r) => {
    // Prefer real publish date, fallback to jiraResolved as proxy
    const endDate = r.publishDate || r.jiraResolved;
    if (!endDate) return;
    if (r.publishDate) hasPublishData = true;

    if (r.poDate && endDate) {
      const mins = differenceInMinutes(endDate, r.poDate);
      if (mins > 0 && mins < 100000) {
        totalMeasured++;
        if (mins <= TARGET_MINUTES) withinSLA++;
        for (const b of buckets) {
          if (mins >= b.min && mins < b.max) { b.count++; break; }
        }
      }
    }
  });

  return {
    buckets,
    withinSLA,
    totalMeasured,
    slaRate: totalMeasured > 0 ? ((withinSLA / totalMeasured) * 100).toFixed(1) : null,
    hasPublishData,
  };
}

// ============ Comparison KPIs (MoM) ============
export function getComparisonKPIs(currentData, previousData) {
  const current = getKPIs(currentData);
  const previous = getKPIs(previousData);

  const delta = (cur, prev) => {
    const diff = cur - prev;
    return diff;
  };
  const ppDelta = (curRate, prevRate) => {
    const diff = parseFloat(curRate) - parseFloat(prevRate);
    return isNaN(diff) ? 0 : diff;
  };

  return {
    ...current,
    comparison: {
      total: { prev: previous.total, delta: delta(current.total, previous.total) },
      withTicket: { prev: previous.withTicket, delta: delta(current.withTicket, previous.withTicket) },
      done: { prev: previous.done, delta: delta(current.done, previous.done) },
      todo: { prev: previous.todo, delta: delta(current.todo, previous.todo) },
      canceled: { prev: previous.canceled, delta: delta(current.canceled, previous.canceled) },
      pagada: { prev: previous.pagada, delta: delta(current.pagada, previous.pagada) },
      matchRate: { prev: previous.matchRate, delta: ppDelta(current.matchRate, previous.matchRate) },
      doneRate: { prev: previous.doneRate, delta: ppDelta(current.doneRate, previous.doneRate) },
    },
    previousPeriodLabel: null,
  };
}

// ============ Filter by period dates ============
export function filterByPeriod(data, start, end) {
  if (!start && !end) return data;
  return data.filter((r) => {
    if (!r.poDate) return false;
    if (start && r.poDate < start) return false;
    if (end && r.poDate > end) return false;
    return true;
  });
}

// ============ Hub Mapping ============
export const HUB_MAPPING = [
  { id: 'lerma', label: 'Lerma', shortLabel: 'LER', color: '#6366f1', match: (val) => /lerma/i.test(val) },
  { id: 'gdl', label: 'Guadalajara', shortLabel: 'GDL', color: '#f59e0b', match: (val) => /^GDL/i.test(val) || /guadalajara/i.test(val) || /tlaquepaque/i.test(val) },
  { id: 'qro', label: 'Queretaro', shortLabel: 'QRO', color: '#10b981', match: (val) => /^QRO/i.test(val) || /quer[eé]taro/i.test(val) },
  { id: 'fortuna', label: 'Fortuna', shortLabel: 'FOR', color: '#3b82f6', match: (val) => /fortuna/i.test(val) },
  { id: 'sanangel', label: 'San Angel', shortLabel: 'SAN', color: '#8b5cf6', match: (val) => /san\s*[aá]ngel/i.test(val) },
  { id: 'otros', label: 'Otros', shortLabel: 'OTR', color: '#94a3b8', match: () => true },
];

export function resolveHub(record) {
  const candidates = [record.destination, record.jiraRegion, record.publishSite].filter(Boolean);
  for (const candidate of candidates) {
    for (const hub of HUB_MAPPING) {
      if (hub.id !== 'otros' && hub.match(candidate)) return hub;
    }
  }
  return HUB_MAPPING[HUB_MAPPING.length - 1];
}

// ============ Pending Cars (purchased but not at hub) ============
export function getPendingCars(data) {
  return data.filter((r) =>
    r.poDate &&
    !r.hasPublishDate &&
    r.jiraStatus !== 'CANCELED'
  );
}

// ============ Pending KPIs ============
export function getPendingKPIs(data) {
  const pending = getPendingCars(data);
  const now = new Date();
  const waitDays = pending.map((r) => differenceInCalendarDays(now, r.poDate));

  const total = pending.length;
  const withTicket = pending.filter((r) => r.hasJiraTicket).length;
  const withoutTicket = total - withTicket;
  const pagada = pending.filter((r) => r.estado === 'Pagada').length;
  const abierta = pending.filter((r) => r.estado === 'Abierta').length;
  const avgWaitDays = waitDays.length > 0
    ? parseFloat((waitDays.reduce((a, b) => a + b, 0) / waitDays.length).toFixed(1))
    : 0;
  const maxWaitDays = waitDays.length > 0 ? Math.max(...waitDays) : 0;
  const overSLA = waitDays.filter((d) => d > 7).length;
  const slaCompliance = total > 0
    ? parseFloat(((1 - overSLA / total) * 100).toFixed(1))
    : 100.0;

  return { total, withTicket, withoutTicket, pagada, abierta, avgWaitDays, maxWaitDays, overSLA, slaCompliance };
}

// ============ Hub Breakdown ============
export function getHubBreakdown(data) {
  const pending = getPendingCars(data);
  const now = new Date();
  const hubMap = {};

  HUB_MAPPING.forEach((h) => {
    hubMap[h.id] = { ...h, pending: 0, withTicket: 0, pagada: 0, waitDays: [], overSLA: 0 };
  });

  pending.forEach((r) => {
    const hub = resolveHub(r);
    const entry = hubMap[hub.id];
    entry.pending += 1;
    if (r.hasJiraTicket) entry.withTicket += 1;
    if (r.estado === 'Pagada') entry.pagada += 1;
    const days = differenceInCalendarDays(now, r.poDate);
    entry.waitDays.push(days);
    if (days > 7) entry.overSLA += 1;
  });

  return HUB_MAPPING.map((h) => {
    const e = hubMap[h.id];
    const avg = e.waitDays.length > 0
      ? parseFloat((e.waitDays.reduce((a, b) => a + b, 0) / e.waitDays.length).toFixed(1))
      : 0;
    const sla = e.pending > 0
      ? parseFloat(((1 - e.overSLA / e.pending) * 100).toFixed(1))
      : 100.0;
    return { ...e, avgWaitDays: avg, slaCompliance: sla, waitDays: undefined };
  }).filter((h) => h.pending > 0 || h.id !== 'otros');
}

// ============ Pending Aging Distribution ============
export function getPendingAgingDistribution(data) {
  const pending = getPendingCars(data);
  const now = new Date();
  const buckets = [
    { label: '0-3 dias', min: 0, max: 3, count: 0, color: '#10b981' },
    { label: '3-5 dias', min: 3, max: 5, count: 0, color: '#22d3ee' },
    { label: '5-7 dias', min: 5, max: 7, count: 0, color: '#f59e0b' },
    { label: '7-14 dias', min: 7, max: 14, count: 0, color: '#f97316' },
    { label: '14+ dias', min: 14, max: Infinity, count: 0, color: '#ef4444' },
  ];

  pending.forEach((r) => {
    const days = differenceInCalendarDays(now, r.poDate);
    for (const b of buckets) {
      if (b.max === Infinity) { if (days >= b.min) { b.count++; break; } }
      else if (days >= b.min && days < b.max) { b.count++; break; }
    }
  });

  return { buckets, total: pending.length };
}

// ============ Pending Route Breakdown ============
export function getPendingRouteBreakdown(data) {
  const pending = getPendingCars(data);
  const completed = data.filter((r) => r.hasPublishDate && r.poDate);
  const now = new Date();
  const routes = {};

  pending.forEach((r) => {
    const hub = resolveHub(r);
    const origin = r.origin || 'Sin origen';
    const route = `${origin} > ${hub.label}`;
    if (!routes[route]) {
      routes[route] = {
        route, origin, hubId: hub.id, hubLabel: hub.label, hubColor: hub.color,
        pending: 0, waitDays: [], overSLA: 0, withTicket: 0, withoutTicket: 0,
        completed: 0, completedDays: [],
      };
    }
    routes[route].pending += 1;
    const days = differenceInCalendarDays(now, r.poDate);
    routes[route].waitDays.push(days);
    if (days > 7) routes[route].overSLA += 1;
    if (r.hasJiraTicket) routes[route].withTicket += 1;
    else routes[route].withoutTicket += 1;
  });

  completed.forEach((r) => {
    const hub = resolveHub(r);
    const origin = r.origin || 'Sin origen';
    const route = `${origin} > ${hub.label}`;
    if (routes[route]) {
      routes[route].completed += 1;
      const days = differenceInCalendarDays(r.publishDate, r.poDate);
      if (days > 0 && days < 120) routes[route].completedDays.push(days);
    }
  });

  return Object.values(routes).map((r) => ({
    route: r.route,
    hubId: r.hubId,
    hubLabel: r.hubLabel,
    hubColor: r.hubColor,
    pending: r.pending,
    avgWaitDays: r.waitDays.length > 0
      ? parseFloat((r.waitDays.reduce((a, b) => a + b, 0) / r.waitDays.length).toFixed(1))
      : 0,
    maxWaitDays: r.waitDays.length > 0 ? Math.max(...r.waitDays) : 0,
    overSLA: r.overSLA,
    slaCompliance: r.pending > 0
      ? parseFloat(((1 - r.overSLA / r.pending) * 100).toFixed(1))
      : 100,
    withoutTicket: r.withoutTicket,
    avgHistoricalDays: r.completedDays.length > 0
      ? parseFloat((r.completedDays.reduce((a, b) => a + b, 0) / r.completedDays.length).toFixed(1))
      : null,
  })).sort((a, b) => b.pending - a.pending);
}

// ============ Pending Daily Trend (last 30 days) ============
export function getPendingDailyTrend(data) {
  const now = new Date();
  const days = [];
  for (let i = 29; i >= 0; i--) {
    const day = startOfDay(subDays(now, i));
    const dayEnd = endOfDay(day);
    const pendingOnDay = data.filter((r) =>
      r.poDate && r.poDate <= dayEnd &&
      (!r.publishDate || r.publishDate > dayEnd) &&
      r.jiraStatus !== 'CANCELED'
    ).length;
    days.push({ date: format(day, 'yyyy-MM-dd'), pending: pendingOnDay });
  }
  return days;
}

// ============ Bottleneck Alerts ============
export function getBottleneckAlerts(hubBreakdown, pendingKPIs) {
  const alerts = [];

  // Worst hub by SLA
  const worstHub = hubBreakdown
    .filter((h) => h.id !== 'otros' && h.pending > 0)
    .sort((a, b) => a.slaCompliance - b.slaCompliance)[0];

  if (worstHub && worstHub.slaCompliance < 100) {
    const severity = worstHub.slaCompliance < 50 ? 'red' : worstHub.slaCompliance < 80 ? 'yellow' : 'green';
    alerts.push({
      severity,
      title: `${worstHub.label}: ${worstHub.overSLA} autos fuera de SLA`,
      description: `${worstHub.slaCompliance}% cumplimiento. ${worstHub.pending} pendientes, promedio ${worstHub.avgWaitDays} dias.`,
    });
  }

  // Cars without Jira tickets
  if (pendingKPIs.withoutTicket > 0) {
    alerts.push({
      severity: pendingKPIs.withoutTicket > 10 ? 'red' : 'yellow',
      title: `${pendingKPIs.withoutTicket} autos sin ticket Jira`,
      description: `Comprados pero sin ticket de logistica. Requieren atencion inmediata.`,
    });
  }

  // Overall SLA
  if (pendingKPIs.slaCompliance < 100) {
    const severity = pendingKPIs.slaCompliance < 50 ? 'red' : pendingKPIs.slaCompliance < 80 ? 'yellow' : 'green';
    alerts.push({
      severity,
      title: `SLA global: ${pendingKPIs.slaCompliance}%`,
      description: `${pendingKPIs.overSLA} de ${pendingKPIs.total} autos pendientes superan 7 dias. Promedio: ${pendingKPIs.avgWaitDays} dias.`,
    });
  }

  // Oldest car
  if (pendingKPIs.maxWaitDays > 14) {
    alerts.push({
      severity: 'red',
      title: `Auto mas antiguo: ${pendingKPIs.maxWaitDays} dias`,
      description: `Al menos un auto lleva ${pendingKPIs.maxWaitDays} dias sin llegar a su hub destino.`,
    });
  }

  return alerts;
}
