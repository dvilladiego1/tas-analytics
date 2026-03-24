import { useState, useMemo, useCallback, useEffect } from 'react';
import { Upload, RefreshCw, FileSpreadsheet, CheckCircle2, Truck } from 'lucide-react';
import {
  parseCSV, processBBDD, processJira, processPublish, joinData, filterByPeriod,
  HUB_MAPPING, resolveHub, getPendingKPIs, getHubBreakdown,
  getPendingAgingDistribution, getPendingRouteBreakdown,
  getPendingDailyTrend, getBottleneckAlerts,
} from './utils/dataProcessor';
import Sidebar from './components/Sidebar';
import TopBar from './components/TopBar';
import PeriodSelector, { computePeriodDates } from './components/PeriodSelector';
import KPICards from './components/KPICards';
import HubCards from './components/HubCards';
import KeyIndicators from './components/KeyIndicators';
import AgingDistribution from './components/AgingDistribution';
import DailyVolumeChart from './components/DailyVolumeChart';
import RoutesBreakdown from './components/RoutesBreakdown';
import './App.css';

export default function App() {
  const [bbddRaw, setBbddRaw] = useState(null);
  const [jiraRaw, setJiraRaw] = useState(null);
  const [publishRaw, setPublishRaw] = useState(null);
  const [joinedData, setJoinedData] = useState(null);

  // Period-based filtering
  const [activePeriod, setActivePeriod] = useState('all');

  // Hub filter
  const [hubFilter, setHubFilter] = useState('all');
  const [loading, setLoading] = useState('');

  // Helper to rejoin data when any source changes
  const rejoin = useCallback((bbdd, jira, publish) => {
    if (bbdd && jira) {
      setJoinedData(joinData(bbdd, jira, publish || []));
    }
  }, []);

  // Auto-load CSVs from public folder on mount
  useEffect(() => {
    async function autoLoad() {
      try {
        setLoading('auto');
        const [bbddRes, jiraRes, publishRes] = await Promise.all([
          fetch('/bbdd.csv'),
          fetch('/jira.csv'),
          fetch('/publish.csv'),
        ]);
        if (bbddRes.ok && jiraRes.ok) {
          const bbddBlob = await bbddRes.blob();
          const jiraBlob = await jiraRes.blob();
          const bbddParsed = await parseCSV(bbddBlob, true);
          const jiraParsed = await parseCSV(jiraBlob);
          const bbddProcessed = processBBDD(bbddParsed);
          const jiraProcessed = processJira(jiraParsed);
          setBbddRaw(bbddProcessed);
          setJiraRaw(jiraProcessed);

          let publishProcessed = [];
          if (publishRes.ok) {
            const publishBlob = await publishRes.blob();
            const publishParsed = await parseCSV(publishBlob, true);
            publishProcessed = processPublish(publishParsed);
            setPublishRaw(publishProcessed);
          }

          setJoinedData(joinData(bbddProcessed, jiraProcessed, publishProcessed));
        }
      } catch (err) {
        console.log('Auto-load skipped:', err.message);
      }
      setLoading('');
    }
    autoLoad();
  }, []);

  const handleBBDD = useCallback(async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setLoading('bbdd');
    try {
      const parsed = await parseCSV(file, true);
      const processed = processBBDD(parsed);
      setBbddRaw(processed);
      rejoin(processed, jiraRaw, publishRaw);
    } catch (err) {
      alert('Error al leer BBDD CSV: ' + err.message);
    }
    setLoading('');
  }, [jiraRaw, publishRaw, rejoin]);

  const handleJira = useCallback(async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setLoading('jira');
    try {
      const parsed = await parseCSV(file);
      const processed = processJira(parsed);
      setJiraRaw(processed);
      rejoin(bbddRaw, processed, publishRaw);
    } catch (err) {
      alert('Error al leer Jira CSV: ' + err.message);
    }
    setLoading('');
  }, [bbddRaw, publishRaw, rejoin]);

  const handlePublish = useCallback(async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setLoading('publish');
    try {
      const parsed = await parseCSV(file, true);
      const processed = processPublish(parsed);
      setPublishRaw(processed);
      rejoin(bbddRaw, jiraRaw, processed);
    } catch (err) {
      alert('Error al leer Publish CSV: ' + err.message);
    }
    setLoading('');
  }, [bbddRaw, jiraRaw, rejoin]);

  // Compute period date range
  const periodDates = useMemo(() => computePeriodDates(activePeriod), [activePeriod]);

  // Filter data by period first
  const periodFilteredData = useMemo(() => {
    if (!joinedData) return [];
    return filterByPeriod(joinedData, periodDates.start, periodDates.end);
  }, [joinedData, periodDates]);

  // Apply hub filter
  const filteredData = useMemo(() => {
    let data = periodFilteredData;
    if (hubFilter !== 'all') {
      data = data.filter((r) => resolveHub(r).id === hubFilter);
    }
    return data;
  }, [periodFilteredData, hubFilter]);

  // Get hub options for dropdown (only show hubs that have data)
  const hubOptions = useMemo(() => {
    const seen = new Set();
    periodFilteredData.forEach((r) => {
      seen.add(resolveHub(r).id);
    });
    return HUB_MAPPING.filter((h) => seen.has(h.id));
  }, [periodFilteredData]);

  // All metrics — pending-focused
  const pendingKPIs = useMemo(() => getPendingKPIs(filteredData), [filteredData]);
  const hubBreakdown = useMemo(() => getHubBreakdown(filteredData), [filteredData]);
  const agingData = useMemo(() => getPendingAgingDistribution(filteredData), [filteredData]);
  const pendingRoutes = useMemo(() => getPendingRouteBreakdown(filteredData), [filteredData]);
  const pendingTrend = useMemo(() => getPendingDailyTrend(filteredData), [filteredData]);
  const alerts = useMemo(
    () => getBottleneckAlerts(hubBreakdown, pendingKPIs),
    [hubBreakdown, pendingKPIs]
  );

  // Upload screen
  if (!joinedData) {
    return (
      <div className="upload-screen">
        <div className="upload-box">
          <Truck size={48} strokeWidth={1.5} />
          <h1>TAS Logistics Dashboard</h1>
          <p>Sube los archivos CSV para generar el dashboard</p>

          <div className="upload-pair">
            <div className={`upload-item ${bbddRaw ? 'upload-done' : ''}`}>
              <label className="upload-btn">
                {bbddRaw ? <CheckCircle2 size={18} /> : <FileSpreadsheet size={18} />}
                {bbddRaw ? `BBDD cargado (${bbddRaw.length})` : 'CSV BBDD (Salesforce)'}
                <input type="file" accept=".csv" onChange={handleBBDD} hidden />
              </label>
              <span className="upload-hint">TAS performance - BBDD.csv</span>
            </div>

            <div className={`upload-item ${jiraRaw ? 'upload-done' : ''}`}>
              <label className="upload-btn">
                {jiraRaw ? <CheckCircle2 size={18} /> : <FileSpreadsheet size={18} />}
                {jiraRaw ? `Jira cargado (${jiraRaw.length})` : 'CSV Jira Data'}
                <input type="file" accept=".csv" onChange={handleJira} hidden />
              </label>
              <span className="upload-hint">TAS performance - Jira Data.csv</span>
            </div>

            <div className={`upload-item ${publishRaw ? 'upload-done' : ''}`}>
              <label className="upload-btn" style={{ background: publishRaw ? '#10b981' : '#8b5cf6' }}>
                {publishRaw ? <CheckCircle2 size={18} /> : <FileSpreadsheet size={18} />}
                {publishRaw ? `Publish cargado (${publishRaw.length})` : 'CSV Publish (Opcional)'}
                <input type="file" accept=".csv" onChange={handlePublish} hidden />
              </label>
              <span className="upload-hint">Content Analysis - Publish.csv</span>
            </div>
          </div>

          {loading && <p className="loading">Procesando {loading}...</p>}
          {bbddRaw && jiraRaw && !joinedData && <p className="loading">Uniendo datos...</p>}
        </div>
      </div>
    );
  }

  const publishCount = publishRaw?.length || 0;

  return (
    <div className="app-shell">
      <Sidebar activeItem="dashboard" />

      <div className="content-area">
        <TopBar
          title="Inventario Pendiente"
          subtitle="Autos comprados sin mover a hub de Kavak"
          bbddCount={bbddRaw?.length}
          jiraCount={jiraRaw?.length}
          recordCount={filteredData.length}
        >
          <PeriodSelector
            activePeriod={activePeriod}
            onPeriodChange={setActivePeriod}
          />
        </TopBar>

        <div className="filter-bar">
          <div className="filter-group">
            <label>Hub Destino</label>
            <select value={hubFilter} onChange={(e) => setHubFilter(e.target.value)}>
              <option value="all">Todos los Hubs</option>
              {hubOptions.map((h) => (
                <option key={h.id} value={h.id}>{h.label}</option>
              ))}
            </select>
          </div>
          <button className="reset-btn" onClick={() => setHubFilter('all')}>
            <RefreshCw size={14} /> Reset
          </button>
          {publishRaw && (
            <span className="top-bar-badge" style={{ background: '#f0fdf4', color: '#16a34a' }}>
              Publish: {publishCount.toLocaleString()} registros
            </span>
          )}
          <span className="record-count">{filteredData.length.toLocaleString()} registros totales</span>
        </div>

        <main className="main-content">
          <KPICards pendingKPIs={pendingKPIs} />

          <HubCards hubs={hubBreakdown} />

          <div className="section-gap" />

          <KeyIndicators alerts={alerts} />

          <div className="section-gap" />

          <AgingDistribution data={agingData} />

          <DailyVolumeChart data={pendingTrend} />

          <RoutesBreakdown data={pendingRoutes} />
        </main>
      </div>
    </div>
  );
}
