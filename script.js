/* ═══════════════════════════════════════════════════════════════
   BCS DASHBOARD · script.js
   Lee el Excel con SheetJS, limpia datos y genera analítica visual
═══════════════════════════════════════════════════════════════ */

'use strict';

// ── Estado global ──
const STATE = {
  raw: {},       // hojas crudas
  estructura: [],
  portafolio: [],
  radicaciones: [],
  renovaciones: [],
  puntos_inactivos: [],
  actas_sarlaft: [],
  brechas: [],
  informe: [],
  filtered: [],  // portafolio filtrado
  currentPage: 1,
  rowsPerPage: 15,
  charts: {}     // instancias ApexCharts
};

// Paleta BCS
const C = {
  blue:    '#005BAC',
  blueL:   '#1a74c4',
  red:     '#CC1F33',
  teal:    '#0d8478',
  green:   '#1a8c4e',
  orange:  '#d97706',
  purple:  '#6d28d9',
  indigo:  '#3730a3',
  slate:   '#1e3050',
  muted:   '#8ba3bf',
  palette: ['#005BAC','#CC1F33','#0d8478','#1a8c4e','#d97706','#6d28d9','#3730a3','#1e3050','#0891b2','#be185d']
};

// ── Logo BCS embebido (extraído del Excel) ──
// Intentamos cargar el logo del mismo directorio; si falla, ignoramos
(function loadLogo(){
  const img = document.getElementById('logo-img');
  if(!img) return;
  const src = 'logo_bcs.png';
  img.src = src;
  img.onerror = () => { img.style.display = 'none'; };
})();

// ── File inputs ──
document.getElementById('file-input').addEventListener('change', e => handleFile(e.target.files[0]));
document.getElementById('file-input2').addEventListener('change', e => handleFile(e.target.files[0]));

// ── Filtros ──
['filter-territorio','filter-zona','filter-gestor'].forEach(id => {
  document.getElementById(id).addEventListener('change', applyFilters);
});
document.getElementById('btn-reset').addEventListener('click', resetFilters);

// ── Tabla ──
document.getElementById('table-search').addEventListener('input', () => { STATE.currentPage = 1; renderTable(); });
document.getElementById('table-estado').addEventListener('change', () => { STATE.currentPage = 1; renderTable(); });

// ═══════════════════════════════════════════ FILE HANDLER ════
function handleFile(file) {
  if (!file) return;
  showLoader(true);
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
      processWorkbook(wb);
    } catch(err) {
      console.error(err);
      alert('Error al leer el archivo: ' + err.message);
      showLoader(false);
    }
  };
  reader.readAsArrayBuffer(file);
}

// ═══════════════════════════════════════════ WORKBOOK ════
function processWorkbook(wb) {
  // Leer hojas como array de arrays
  const readSheet = name => {
    const ws = wb.Sheets[name];
    if (!ws) return [];
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: false });
  };

  STATE.raw = {
    informe:        readSheet('INFORME'),
    estructura:     readSheet('DIM_ESTRUCTURA_CB'),
    portafolio:     readSheet('DIM_PORTAFOLIO'),
    radicaciones:   readSheet('RADICACIONES'),
    renovaciones:   readSheet('RENOVACIONES'),
    puntos_inactivos: readSheet('PUNTOS INACTIVOS'),
    actas_sarlaft:  readSheet('ACTAS_SARLAFT'),
    brechas:        readSheet('BRECHAS_VISITAS'),
  };

  // Fecha de corte
  const inf = STATE.raw.informe;
  let fechaCorte = '–';
  if (inf && inf[0]) {
    const cell1 = inf[0][1];
    if (cell1 && typeof cell1 === 'string' && cell1.includes('corte')) {
      fechaCorte = cell1.replace('Fecha de corte', '').trim();
    }
  }
  document.getElementById('fecha-corte-val').textContent = fechaCorte;

  // Parsear cada hoja
  STATE.estructura       = parseEstructura(STATE.raw.estructura);
  STATE.portafolio       = parsePortafolio(STATE.raw.portafolio);
  STATE.radicaciones     = parseRadicaciones(STATE.raw.radicaciones);
  STATE.renovaciones     = parseRenovaciones(STATE.raw.renovaciones);
  STATE.puntos_inactivos = parsePuntosInactivos(STATE.raw.puntos_inactivos);
  STATE.actas_sarlaft    = parseActasSarlaft(STATE.raw.actas_sarlaft);
  STATE.brechas          = parseBrechas(STATE.raw.brechas);

  // Enriquecer portafolio con zona/territorio
  const mapGestor = {};
  STATE.estructura.forEach(e => { mapGestor[String(e.id_gestor)] = e; });
  STATE.portafolio.forEach(p => {
    const m = mapGestor[String(p.id_gestor)];
    if (m) { p.zona = m.zona; p.territorio = m.territorio; p.gestor_nombre = m.gestor; }
  });

  STATE.filtered = [...STATE.portafolio];

  // Poblar filtros
  populateFilters();

  // UI
  showLoader(false);
  document.getElementById('empty-state').classList.add('hidden');
  document.getElementById('main-content').classList.remove('hidden');

  // Mostrar botón exportar
  const btnExport = document.getElementById('btn-export');
  if (btnExport) btnExport.classList.remove('hidden');

  // Renderizar todo
  renderKPIs();
  renderAllCharts();
  renderTable();
  renderInsights();
}

// ═══════════════════════════════════════════ PARSERS ════

function clean(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim().replace(/\s+/g, ' ');
}

function toNum(v) {
  const n = parseFloat(String(v).replace(/,/g,''));
  return isNaN(n) ? 0 : n;
}

function parseEstructura(rows) {
  if (!rows || rows.length < 2) return [];
  return rows.slice(1).map(r => ({
    territorio: clean(r[0]).toUpperCase(),
    zona:       clean(r[1]),
    gestor:     clean(r[2]),
    id_gestor:  String(clean(r[3]))
  })).filter(r => r.id_gestor);
}

function parsePortafolio(rows) {
  if (!rows || rows.length < 2) return [];
  return rows.slice(1).map(r => ({
    id_gestor:    clean(r[0]),
    cod_pds:      clean(r[1]),
    nombre:       clean(r[2]),
    id_corresponsal: clean(r[3]),
    titular:      clean(r[4]),
    direccion:    clean(r[5]),
    modelo:       clean(r[8]),
    estado:       clean(r[9]),
    cuenta:       clean(r[10]),
    dias_sobregiro: toNum(r[11]),
    estado_cartera: clean(r[12]) || 'Al día',
    valor_sobregiro: toNum(r[13]),
    total_tx:     toNum(r[14]),
    estrategia:   clean(r[15]),
    estado2:      clean(r[16]),
    tx_consultas:        toNum(r[17]),
    tx_depositos:        toNum(r[18]),
    tx_pago_pila:        toNum(r[19]),
    tx_pago_producto:    toNum(r[20]),
    tx_recaudos:         toNum(r[21]),
    tx_retiros:          toNum(r[22]),
    tx_transferencias:   toNum(r[23]),
    zona: '', territorio: '', gestor_nombre: ''
  })).filter(r => r.cod_pds);
}

function parseRadicaciones(rows) {
  if (!rows || rows.length < 2) return [];
  return rows.slice(1).map(r => ({
    id_gestor: clean(r[0]),
    cod_pds:   clean(r[1]),
    nombre:    clean(r[2]),
    modelo:    clean(r[8]),
    fecha_envio: clean(r[9]),
    tipo:      clean(r[11]),
    estado:    clean(r[12])
  })).filter(r => r.cod_pds);
}

function parseRenovaciones(rows) {
  if (!rows || rows.length < 2) return [];
  return rows.slice(1).map(r => ({
    id_gestor: clean(r[0]),
    cod_pds:   clean(r[1]),
    nombre:    clean(r[2]),
    titular:   clean(r[4]),
    fecha_renovacion: clean(r[8]),
    mes:       clean(r[9]),
    enviados:  clean(r[10]),
    estado:    clean(r[11]),
    fecha_aprobacion: clean(r[12])
  })).filter(r => r.cod_pds);
}

function parsePuntosInactivos(rows) {
  if (!rows || rows.length < 2) return [];
  return rows.slice(1).map(r => ({
    id_gestor: clean(r[0]),
    cod_pds:   clean(r[1]),
    nombre:    clean(r[2]),
    titular:   clean(r[4]),
    novedad:   clean(r[8])
  })).filter(r => r.cod_pds);
}

function parseActasSarlaft(rows) {
  if (!rows || rows.length < 2) return [];
  return rows.slice(1).map(r => ({
    id_gestor: clean(r[0]),
    cod_pds:   clean(r[1]),
    nombre:    clean(r[2]),
    acta:      clean(r[9]),
    evaluacion: clean(r[10]),
    mes:       clean(r[11]),
    fecha_limite: clean(r[12]),
    tipo_operacion: clean(r[13])
  })).filter(r => r.cod_pds);
}

function parseBrechas(rows) {
  if (!rows || rows.length < 2) return [];
  return rows.slice(1).map(r => ({
    gestor:           clean(r[0]),
    total_visitas:    toNum(r[1]),
    realizadas:       toNum(r[2]),
    ejecucion:        toNum(r[3]),
    meta:             toNum(r[4]),
    pct_cumpl:        toNum(r[5]),
    brecha:           toNum(r[6])
  })).filter(r => r.gestor);
}

// ═══════════════════════════════════════════ FILTROS ════
function populateFilters() {
  const territorios = [...new Set(STATE.portafolio.map(p => p.territorio).filter(Boolean))].sort();
  const zonas       = [...new Set(STATE.portafolio.map(p => p.zona).filter(Boolean))].sort();
  const gestores    = [...new Set(STATE.portafolio.map(p => p.gestor_nombre).filter(Boolean))].sort();

  fillSelect('filter-territorio', territorios);
  fillSelect('filter-zona', zonas);
  fillSelect('filter-gestor', gestores);
}

function fillSelect(id, values) {
  const sel = document.getElementById(id);
  const current = sel.value;
  sel.innerHTML = '<option value="all">Todos</option>';
  values.forEach(v => {
    const opt = document.createElement('option');
    opt.value = v; opt.textContent = v;
    sel.appendChild(opt);
  });
  if (values.includes(current)) sel.value = current;
}

function applyFilters() {
  const ter = document.getElementById('filter-territorio').value;
  const zon = document.getElementById('filter-zona').value;
  const ges = document.getElementById('filter-gestor').value;

  STATE.filtered = STATE.portafolio.filter(p => {
    if (ter !== 'all' && p.territorio !== ter) return false;
    if (zon !== 'all' && p.zona !== zon) return false;
    if (ges !== 'all' && p.gestor_nombre !== ges) return false;
    return true;
  });

  STATE.currentPage = 1;
  renderKPIs();
  renderAllCharts();
  renderTable();
  renderInsights();
}

function resetFilters() {
  ['filter-territorio','filter-zona','filter-gestor'].forEach(id => {
    document.getElementById(id).value = 'all';
  });
  STATE.filtered = [...STATE.portafolio];
  STATE.currentPage = 1;
  renderKPIs();
  renderAllCharts();
  renderTable();
  renderInsights();
}

// ═══════════════════════════════════════════ KPIs ════
function renderKPIs() {
  const d = STATE.filtered;

  const terFil = document.getElementById('filter-territorio').value;
  const zonFil = document.getElementById('filter-zona').value;
  const gesFil = document.getElementById('filter-gestor').value;

  const gIds = new Set(
    STATE.estructura.filter(e => {
      if (terFil !== 'all' && e.territorio !== terFil) return false;
      if (zonFil !== 'all' && e.zona !== zonFil) return false;
      if (gesFil !== 'all' && e.gestor !== gesFil) return false;
      return true;
    }).map(e => e.id_gestor)
  );

  const rads  = STATE.radicaciones.filter(r => gIds.has(r.id_gestor));
  const renov = STATE.renovaciones.filter(r => gIds.has(r.id_gestor));
  const sarla = STATE.actas_sarlaft.filter(r => gIds.has(r.id_gestor));
  const inact = STATE.puntos_inactivos.filter(r => gIds.has(r.id_gestor));
  const mora  = d.filter(p => (p.estado_cartera || '').toLowerCase().includes('mora'));

  const total   = d.length;
  const activos = d.filter(p => p.estado === 'Activo').length;
  const cancelados = d.filter(p => p.estado === 'Cancelado').length;
  const clientesUni = new Set(d.map(p => p.id_corresponsal).filter(Boolean)).size;
  const sinTx  = d.filter(p => p.total_tx === 0);
  const totalTx = d.reduce((s, p) => s + p.total_tx, 0);

  // ── KPI 1: Cobertura ──
  const metaCobertura = total > 0 ? Math.round(total * 1.047) : 0; // aproximación meta +4.7%
  const pctCobertura  = metaCobertura > 0 ? ((total / metaCobertura) * 100).toFixed(1) : '–';
  const pctActivacion = total > 0 ? ((activos / total) * 100).toFixed(1) : '0';
  setKPI('kpi-portafolio',     fmt(total));
  setText('kpi-cobertura-meta', `Meta: ${fmt(metaCobertura)} · ${pctCobertura}% cumplimiento`);
  setKPI('kpi-clientes',       fmt(clientesUni));
  setKPI('kpi-activacion-pct', pctActivacion + '%');
  setKPI('kpi-cancelados',     fmt(cancelados));
  setKPI('kpi-inactivos-foot', fmt(inact.length));

  // ── KPI 2: Transaccionalidad ──
  const metaTx    = 530000; // meta fija (puede leerse del Excel en futuro)
  const pctTx     = metaTx > 0 ? ((totalTx / metaTx) * 100).toFixed(1) : '–';
  const pctSinTx  = total > 0 ? ((sinTx.length / total) * 100).toFixed(1) : '0';
  setKPI('kpi-tx',          fmt(totalTx));
  setText('kpi-tx-meta',   `Meta: ${fmt(metaTx)} · ${pctTx}% ejec.`);
  setKPI('kpi-meta-tx-val', fmt(sinTx.length));
  setKPI('kpi-sin-tx-pct',  pctSinTx + '%');

  // ── KPI 3: Renovaciones ──
  const renovPend = renov.filter(r => (r.estado || '').toUpperCase().includes('PEND') || (r.estado || '').toUpperCase().includes('SIN RADIC')).length;
  const metaRenov = 154; // meta referencial
  const pctRenov  = metaRenov > 0 ? ((renov.length / metaRenov) * 100).toFixed(1) : '–';
  const pctRenovPend = renov.length > 0 ? ((renovPend / renov.length) * 100).toFixed(1) : '0';
  setKPI('kpi-renovaciones', fmt(renov.length));
  setText('kpi-renov-meta',  `Meta: ${metaRenov} · Avance: ${pctRenov}%`);
  setKPI('kpi-renov-pend',   fmt(renovPend));
  setKPI('kpi-renov-pct',    pctRenovPend + '%');

  // ── KPI 4: Brechas ──
  const totalVisitas    = STATE.brechas.reduce((s, b) => s + (b.meta || 0), 0);
  const realizadas      = STATE.brechas.reduce((s, b) => s + (b.realizadas || 0), 0);
  const pctBrechas      = totalVisitas > 0 ? ((realizadas / totalVisitas) * 100).toFixed(1) : '–';
  setKPI('kpi-brechas-total',     fmt(realizadas));
  setText('kpi-brechas-meta',     `Meta: ${fmt(totalVisitas)} · ${pctBrechas}% cumpl.`);
  setKPI('kpi-brechas-realizadas', fmt(realizadas));
  setKPI('kpi-brechas-pct',        pctBrechas + '%');

  // ── KPI 5: Morosidad ──
  const proxVto = renov.filter(r => {
    const e = (r.estado || '').toUpperCase();
    return e.includes('PRÓX') || e.includes('PROX') || e.includes('VENC');
  }).length;
  setKPI('kpi-mora',          fmt(mora.length));
  setText('kpi-mora-meta',   `En mora · ${fmt(proxVto)} próximos vencimiento`);
  setKPI('kpi-mora-count',    fmt(mora.length));
  setKPI('kpi-prox-vto-count', fmt(proxVto));
}

function setKPI(id, val) {
  const el = document.getElementById(id);
  if (el) el.textContent = val;
}

function setText(id, val) {
  const el = document.getElementById(id);
  if (el) el.textContent = val;
}

function fmt(n) {
  if (typeof n === 'number') return n.toLocaleString('es-CO');
  return n;
}

// ═══════════════════════════════════════════ CHARTS ════

function destroyChart(key) {
  if (STATE.charts[key]) { STATE.charts[key].destroy(); delete STATE.charts[key]; }
}

function renderAllCharts() {
  renderPortafolioTerritorio();
  renderCarteraEstado();
  renderTxZona();
  renderPortafolioZona();
  renderFunnelRadicaciones();
  renderFunnelRenovaciones();
  renderBrechas();
  renderMoraZona();
  renderSarlaft();
}

// --- 1. Portafolio por Territorio ---
function renderPortafolioTerritorio() {
  const key = 'portafolio-territorio';
  destroyChart(key);

  const byTer = groupBy(STATE.filtered, 'territorio');
  const labels = Object.keys(byTer).filter(Boolean).sort();
  const totales = labels.map(t => byTer[t].length);
  const activos = labels.map(t => byTer[t].filter(p => p.estado === 'Activo').length);

  const options = {
    chart: { type: 'bar', height: 270, toolbar: { show: false }, fontFamily: 'Sora, sans-serif' },
    series: [
      { name: 'Total Portafolio', data: totales },
      { name: 'Activos', data: activos }
    ],
    xaxis: { categories: labels, labels: { style: { fontSize: '11px', colors: '#5a7292' } } },
    yaxis: { labels: { style: { colors: '#5a7292' } } },
    colors: [C.blue, C.green],
    plotOptions: { bar: { borderRadius: 5, columnWidth: '55%' } },
    dataLabels: { enabled: false },
    legend: { position: 'top' },
    grid: { borderColor: '#e8f1fb', strokeDashArray: 3 },
    tooltip: { y: { formatter: v => v.toLocaleString('es-CO') } }
  };

  STATE.charts[key] = new ApexCharts(document.getElementById('chart-portafolio-territorio'), options);
  STATE.charts[key].render();
}

// --- 2. Estado de Cartera ---
function renderCarteraEstado() {
  const key = 'cartera-estado';
  destroyChart(key);

  const byEstado = groupBy(STATE.filtered, 'estado_cartera');
  const labels = Object.keys(byEstado).filter(Boolean);
  const values = labels.map(l => byEstado[l].length);

  const options = {
    chart: { type: 'donut', height: 270, toolbar: { show: false }, fontFamily: 'Sora, sans-serif' },
    series: values,
    labels: labels,
    colors: [C.green, C.red, C.orange, C.blue, C.muted],
    plotOptions: { pie: { donut: { size: '65%', labels: { show: true, total: { show: true, label: 'Total', formatter: () => fmt(values.reduce((a,b)=>a+b,0)) } } } } },
    dataLabels: { enabled: true, style: { fontSize: '11px' } },
    legend: { position: 'bottom', fontSize: '11px' }
  };

  STATE.charts[key] = new ApexCharts(document.getElementById('chart-cartera-estado'), options);
  STATE.charts[key].render();
}

// --- 3. TX por Zona ---
function renderTxZona() {
  const key = 'tx-zona';
  destroyChart(key);

  const byZona = groupBy(STATE.filtered.filter(p => p.zona), 'zona');
  const labels = Object.keys(byZona).sort();
  const values = labels.map(z => byZona[z].reduce((s,p) => s + p.total_tx, 0));

  // Sort descending
  const combined = labels.map((l,i) => ({ l, v: values[i] })).sort((a,b) => b.v - a.v);

  const options = {
    chart: { type: 'bar', height: 280, toolbar: { show: false }, fontFamily: 'Sora, sans-serif' },
    series: [{ name: 'Transacciones', data: combined.map(c => c.v) }],
    xaxis: { categories: combined.map(c => c.l), labels: { style: { fontSize: '10px', colors: '#5a7292' }, rotate: -35 } },
    yaxis: { labels: { style: { colors: '#5a7292' }, formatter: v => (v/1000).toFixed(0)+'K' } },
    colors: [C.teal],
    plotOptions: { bar: { borderRadius: 4, distributed: true } },
    dataLabels: { enabled: false },
    legend: { show: false },
    grid: { borderColor: '#e8f1fb', strokeDashArray: 3 },
    tooltip: { y: { formatter: v => v.toLocaleString('es-CO') } }
  };

  STATE.charts[key] = new ApexCharts(document.getElementById('chart-tx-zona'), options);
  STATE.charts[key].render();
}

// --- 4. Portafolio por Zona ---
function renderPortafolioZona() {
  const key = 'portafolio-zona';
  destroyChart(key);

  const byZona = groupBy(STATE.filtered.filter(p => p.zona), 'zona');
  const labels = Object.keys(byZona).sort();
  const totales = labels.map(z => byZona[z].length);
  const activos = labels.map(z => byZona[z].filter(p => p.estado === 'Activo').length);
  const mora    = labels.map(z => byZona[z].filter(p => (p.estado_cartera||'').includes('ora')).length);

  const options = {
    chart: { type: 'bar', height: 270, stacked: false, toolbar: { show: false }, fontFamily: 'Sora, sans-serif' },
    series: [
      { name: 'Total', data: totales },
      { name: 'Activos', data: activos },
      { name: 'En Mora', data: mora }
    ],
    xaxis: { categories: labels, labels: { style: { fontSize: '10px', colors: '#5a7292' }, rotate: -30 } },
    yaxis: { labels: { style: { colors: '#5a7292' } } },
    colors: [C.blue, C.green, C.red],
    plotOptions: { bar: { borderRadius: 4, columnWidth: '60%' } },
    dataLabels: { enabled: false },
    legend: { position: 'top' },
    grid: { borderColor: '#e8f1fb', strokeDashArray: 3 }
  };

  STATE.charts[key] = new ApexCharts(document.getElementById('chart-portafolio-zona'), options);
  STATE.charts[key].render();
}

// --- 5. Embudo Radicaciones ---
function renderFunnelRadicaciones() {
  const key = 'funnel-radicaciones';
  destroyChart(key);

  const byEstado = groupBy(STATE.radicaciones, 'estado');
  const data = Object.entries(byEstado).map(([e,v]) => ({ x: e || 'Sin Estado', y: v.length }))
    .sort((a,b) => b.y - a.y);

  const options = {
    chart: { type: 'bar', height: 270, toolbar: { show: false }, fontFamily: 'Sora, sans-serif' },
    series: [{ name: 'Radicaciones', data: data.map(d => d.y) }],
    xaxis: { categories: data.map(d => d.x), labels: { style: { fontSize: '10px', colors: '#5a7292' } } },
    yaxis: { labels: { style: { colors: '#5a7292' } } },
    colors: [C.blue, C.teal, C.green, C.orange, C.red, C.purple, C.indigo, C.muted],
    plotOptions: { bar: { borderRadius: 5, distributed: true, horizontal: true } },
    dataLabels: { enabled: true, style: { fontSize: '11px' } },
    legend: { show: false },
    grid: { borderColor: '#e8f1fb', strokeDashArray: 3 }
  };

  STATE.charts[key] = new ApexCharts(document.getElementById('chart-funnel-radicaciones'), options);
  STATE.charts[key].render();
}

// --- 6. Embudo Renovaciones ---
function renderFunnelRenovaciones() {
  const key = 'funnel-renovaciones';
  destroyChart(key);

  const byEstado = groupBy(STATE.renovaciones, 'estado');
  const data = Object.entries(byEstado).map(([e,v]) => ({ x: e || 'Sin Estado', y: v.length }))
    .sort((a,b) => b.y - a.y);

  const options = {
    chart: { type: 'bar', height: 270, toolbar: { show: false }, fontFamily: 'Sora, sans-serif' },
    series: [{ name: 'Renovaciones', data: data.map(d => d.y) }],
    xaxis: { categories: data.map(d => d.x), labels: { style: { fontSize: '10px', colors: '#5a7292' } } },
    yaxis: { labels: { style: { colors: '#5a7292' } } },
    colors: [C.indigo, C.teal, C.green, C.orange, C.red, C.purple, C.blue, C.muted],
    plotOptions: { bar: { borderRadius: 5, distributed: true, horizontal: true } },
    dataLabels: { enabled: true, style: { fontSize: '11px' } },
    legend: { show: false },
    grid: { borderColor: '#e8f1fb', strokeDashArray: 3 }
  };

  STATE.charts[key] = new ApexCharts(document.getElementById('chart-funnel-renovaciones'), options);
  STATE.charts[key].render();
}

// --- 7. Brechas de Visitas ---
function renderBrechas() {
  const key = 'brechas';
  destroyChart(key);

  const data = STATE.brechas.sort((a,b) => a.pct_cumpl - b.pct_cumpl);
  const pcts = data.map(d => +(d.pct_cumpl * 100).toFixed(1));
  const nombres = data.map(d => {
    const parts = d.gestor.split(' ');
    return parts[0] + (parts[1] ? ' ' + parts[1] : '');
  });

  const options = {
    chart: { type: 'bar', height: 280, toolbar: { show: false }, fontFamily: 'Sora, sans-serif' },
    series: [{ name: '% Cumplimiento', data: pcts }],
    xaxis: { categories: nombres, labels: { style: { fontSize: '10px', colors: '#5a7292' }, rotate: -30 } },
    yaxis: { max: 120, labels: { style: { colors: '#5a7292' }, formatter: v => v + '%' } },
    colors: pcts.map(p => p >= 100 ? C.green : p >= 80 ? C.orange : C.red),
    plotOptions: { bar: { borderRadius: 4, distributed: true, columnWidth: '55%' } },
    dataLabels: { enabled: true, formatter: v => v + '%', style: { fontSize: '10px' } },
    annotations: { yaxis: [{ y: 100, borderColor: C.blue, strokeDashArray: 4, label: { text: 'Meta 100%', style: { color: C.blue, fontSize: '10px' } } }] },
    legend: { show: false },
    grid: { borderColor: '#e8f1fb', strokeDashArray: 3 }
  };

  STATE.charts[key] = new ApexCharts(document.getElementById('chart-brechas'), options);
  STATE.charts[key].render();
}

// --- 8. Ranking Mora por Zona ---
function renderMoraZona() {
  const key = 'mora-zona';
  destroyChart(key);

  // Enriquecer renovaciones con zona
  const mapGestor = {};
  STATE.estructura.forEach(e => { mapGestor[e.id_gestor] = e; });

  // Mora en cartera por zona
  const moraByZona = {};
  const proxVencByZona = {};

  STATE.portafolio.forEach(p => {
    if (!p.zona) return;
    if (!moraByZona[p.zona]) moraByZona[p.zona] = 0;
    if ((p.estado_cartera || '').toLowerCase().includes('mora')) moraByZona[p.zona]++;
  });

  STATE.renovaciones.forEach(r => {
    const g = mapGestor[r.id_gestor];
    if (!g) return;
    const zona = g.zona;
    if (!proxVencByZona[zona]) proxVencByZona[zona] = 0;
    proxVencByZona[zona]++;
  });

  const zonas = [...new Set([...Object.keys(moraByZona), ...Object.keys(proxVencByZona)])].sort();
  const mora  = zonas.map(z => moraByZona[z] || 0);
  const prox  = zonas.map(z => proxVencByZona[z] || 0);

  const options = {
    chart: { type: 'bar', height: 280, toolbar: { show: false }, fontFamily: 'Sora, sans-serif' },
    series: [
      { name: 'En Mora', data: mora },
      { name: 'Próx. Vencimiento', data: prox }
    ],
    xaxis: { categories: zonas, labels: { style: { fontSize: '10px', colors: '#5a7292' }, rotate: -30 } },
    yaxis: { labels: { style: { colors: '#5a7292' } } },
    colors: [C.red, C.orange],
    plotOptions: { bar: { borderRadius: 4, columnWidth: '55%' } },
    dataLabels: { enabled: true, style: { fontSize: '10px' } },
    legend: { position: 'top' },
    grid: { borderColor: '#e8f1fb', strokeDashArray: 3 }
  };

  STATE.charts[key] = new ApexCharts(document.getElementById('chart-mora-zona'), options);
  STATE.charts[key].render();
}

// --- 9. SARLAFT por mes y tipo ---
function renderSarlaft() {
  const key = 'sarlaft';
  destroyChart(key);

  const byTipo = groupBy(STATE.actas_sarlaft, 'tipo_operacion');
  const labels = Object.keys(byTipo).filter(Boolean);
  const values = labels.map(l => byTipo[l].length);

  const options = {
    chart: { type: 'pie', height: 280, toolbar: { show: false }, fontFamily: 'Sora, sans-serif' },
    series: values,
    labels: labels,
    colors: [C.blue, C.red, C.orange, C.purple],
    dataLabels: { enabled: true, style: { fontSize: '11px' } },
    legend: { position: 'bottom', fontSize: '11px' },
    plotOptions: { pie: { expandOnClick: true } }
  };

  STATE.charts[key] = new ApexCharts(document.getElementById('chart-sarlaft'), options);
  STATE.charts[key].render();
}

// ═══════════════════════════════════════════ TABLE ════
function renderTable() {
  const search = document.getElementById('table-search').value.toLowerCase();
  const estadoFil = document.getElementById('table-estado').value;

  let rows = STATE.filtered.filter(p => {
    if (estadoFil !== 'all' && p.estado !== estadoFil) return false;
    if (search) {
      const hay = `${p.nombre} ${p.gestor_nombre} ${p.titular} ${p.zona}`.toLowerCase();
      if (!hay.includes(search)) return false;
    }
    return true;
  });

  document.getElementById('table-count').textContent = `${rows.length.toLocaleString('es-CO')} registros`;

  const total = Math.ceil(rows.length / STATE.rowsPerPage);
  STATE.currentPage = Math.min(STATE.currentPage, total || 1);
  const start = (STATE.currentPage - 1) * STATE.rowsPerPage;
  const page  = rows.slice(start, start + STATE.rowsPerPage);

  const tbody = document.getElementById('table-body');
  tbody.innerHTML = page.map(p => `
    <tr>
      <td title="${p.gestor_nombre}">${truncate(p.gestor_nombre, 22)}</td>
      <td>${p.zona || '–'}</td>
      <td>${p.territorio ? p.territorio.replace('TERRITORIO ','') : '–'}</td>
      <td title="${p.nombre}">${truncate(p.nombre, 30)}</td>
      <td><span class="badge ${badgeClass(p.estado)}">${p.estado || '–'}</span></td>
      <td><span class="badge ${cartBadge(p.estado_cartera)}">${p.estado_cartera || 'Al día'}</span></td>
      <td style="font-family:var(--mono)">${p.total_tx.toLocaleString('es-CO')}</td>
      <td>${p.estrategia || '–'}</td>
      <td><span class="badge ${estado2Badge(p.estado2)}">${p.estado2 || '–'}</span></td>
    </tr>
  `).join('');

  renderPagination(total, rows.length);
}

function truncate(str, n) {
  if (!str) return '–';
  return str.length > n ? str.slice(0, n) + '…' : str;
}

function badgeClass(e) {
  if (!e) return 'badge-default';
  const l = e.toLowerCase();
  if (l === 'activo') return 'badge-activo';
  if (l === 'inactivo') return 'badge-inactivo';
  if (l === 'cancelado') return 'badge-cancelado';
  return 'badge-default';
}

function cartBadge(e) {
  if (!e) return 'badge-al-dia';
  const l = e.toLowerCase();
  if (l.includes('mora')) return 'badge-mora';
  if (l.includes('día') || l.includes('dia')) return 'badge-al-dia';
  return 'badge-default';
}

function estado2Badge(e) {
  if (!e) return 'badge-default';
  const l = e.toLowerCase();
  if (l === 'mora') return 'badge-mora';
  if (l === 'negado') return 'badge-inactivo';
  if (l === 'activo') return 'badge-activo';
  return 'badge-default';
}

function renderPagination(total, count) {
  const pg = document.getElementById('pagination');
  if (total <= 1) { pg.innerHTML = ''; return; }

  const maxBtns = 7;
  let pages = [];
  if (total <= maxBtns) {
    pages = Array.from({ length: total }, (_, i) => i + 1);
  } else {
    const c = STATE.currentPage;
    pages = [1];
    if (c > 3) pages.push('…');
    for (let i = Math.max(2, c-1); i <= Math.min(total-1, c+1); i++) pages.push(i);
    if (c < total - 2) pages.push('…');
    pages.push(total);
  }

  pg.innerHTML = `
    <button class="page-btn" ${STATE.currentPage===1?'disabled':''} onclick="goPage(${STATE.currentPage-1})">‹</button>
    ${pages.map(p => p === '…'
      ? `<span style="padding:4px 6px;color:var(--text-muted)">…</span>`
      : `<button class="page-btn ${p===STATE.currentPage?'active':''}" onclick="goPage(${p})">${p}</button>`
    ).join('')}
    <button class="page-btn" ${STATE.currentPage===total?'disabled':''} onclick="goPage(${STATE.currentPage+1})">›</button>
  `;
}

window.goPage = function(p) {
  STATE.currentPage = p;
  renderTable();
};

// ═══════════════════════════════════════════ INSIGHTS ════
function renderInsights() {
  const d = STATE.filtered;
  const total  = d.length;
  const activos = d.filter(p => p.estado === 'Activo').length;
  const mora   = d.filter(p => (p.estado_cartera||'').toLowerCase().includes('mora'));
  const sinTx  = d.filter(p => p.total_tx === 0);
  const pctAct = total > 0 ? ((activos/total)*100).toFixed(1) : 0;
  const pctMora = total > 0 ? ((mora.length/total)*100).toFixed(1) : 0;
  const totalTx = d.reduce((s,p) => s + p.total_tx, 0);
  const avgTx   = total > 0 ? (totalTx / total).toFixed(0) : 0;

  // Zona con más mora
  const moraByZona = {};
  mora.forEach(p => { moraByZona[p.zona] = (moraByZona[p.zona]||0)+1; });
  const topMoraZona = Object.entries(moraByZona).sort((a,b)=>b[1]-a[1])[0];

  // Gestor con más tx
  const txByGestor = {};
  d.forEach(p => { txByGestor[p.gestor_nombre] = (txByGestor[p.gestor_nombre]||0)+p.total_tx; });
  const topTxGestor = Object.entries(txByGestor).sort((a,b)=>b[1]-a[1])[0];

  // Brecha más crítica
  const brechasCrit = STATE.brechas.filter(b => b.pct_cumpl < 0.8).sort((a,b)=>a.pct_cumpl-b.pct_cumpl);

  const cards = [];

  if (pctAct >= 90) {
    cards.push({ type: 'good', icon: '✅', text: `<strong>${pctAct}%</strong> del portafolio está activo (${activos.toLocaleString('es-CO')} de ${total.toLocaleString('es-CO')} puntos). Excelente cobertura.` });
  } else if (pctAct >= 75) {
    cards.push({ type: '', icon: '📊', text: `<strong>${pctAct}%</strong> de los puntos están activos. Hay oportunidad de mejora en activación.` });
  } else {
    cards.push({ type: 'warn', icon: '⚠️', text: `Solo <strong>${pctAct}%</strong> del portafolio está activo. Se recomienda revisar los puntos inactivos.` });
  }

  if (mora.length > 0) {
    cards.push({ type: pctMora > 10 ? 'danger' : 'warn', icon: '💳', text: `<strong>${mora.length}</strong> corresponsales en mora (${pctMora}%). ${topMoraZona ? `La zona con más mora es <strong>${topMoraZona[0]}</strong> con ${topMoraZona[1]} puntos.` : ''}` });
  } else {
    cards.push({ type: 'good', icon: '💚', text: `No se registran puntos en mora en la selección actual. Canal financieramente sano.` });
  }

  if (sinTx.length > 0) {
    cards.push({ type: 'warn', icon: '🔴', text: `<strong>${sinTx.length}</strong> puntos sin transacciones. Representan el ${((sinTx.length/total)*100).toFixed(1)}% del portafolio activo.` });
  }

  if (topTxGestor) {
    cards.push({ type: '', icon: '🏆', text: `El gestor con mayor transaccionalidad es <strong>${topTxGestor[0]}</strong> con <strong>${Number(topTxGestor[1]).toLocaleString('es-CO')}</strong> transacciones.` });
  }

  cards.push({ type: '', icon: '📈', text: `Promedio de <strong>${Number(avgTx).toLocaleString('es-CO')}</strong> transacciones por punto en el período de corte.` });

  if (brechasCrit.length > 0) {
    cards.push({ type: 'warn', icon: '👁️', text: `<strong>${brechasCrit.length}</strong> gestor(es) con cumplimiento de visitas inferior al 80%. El más crítico: <strong>${brechasCrit[0].gestor}</strong> al ${(brechasCrit[0].pct_cumpl*100).toFixed(0)}%.` });
  } else {
    cards.push({ type: 'good', icon: '🚀', text: `Todos los gestores superan el 80% de cumplimiento en visitas. Buen ritmo operativo.` });
  }

  const sarFalt = STATE.actas_sarlaft.length;
  if (sarFalt > 0) {
    cards.push({ type: 'warn', icon: '📋', text: `<strong>${sarFalt}</strong> actas SARLAFT pendientes de entrega. Requieren atención para evitar incumplimiento regulatorio.` });
  }

  if (STATE.renovaciones.length > 0) {
    const sinRad = STATE.renovaciones.filter(r => r.estado === 'SIN RADICAR').length;
    if (sinRad > 0) {
      cards.push({ type: 'warn', icon: '🔁', text: `<strong>${sinRad}</strong> renovaciones sin radicar. Riesgo de vencimiento de contratos sin gestionar.` });
    }
  }

  const grid = document.getElementById('insights-grid');
  grid.innerHTML = cards.map(c => `
    <div class="insight-card ${c.type}">
      <span class="insight-icon">${c.icon}</span>
      <span class="insight-text">${c.text}</span>
    </div>
  `).join('');
}

// ═══════════════════════════════════════════ UTILS ════
function groupBy(arr, key) {
  return arr.reduce((acc, item) => {
    const k = item[key] || '';
    (acc[k] = acc[k] || []).push(item);
    return acc;
  }, {});
}

function showLoader(show) {
  document.getElementById('loader').classList.toggle('hidden', !show);
}

// ─── Auto-load if Excel dropped on page ──
document.addEventListener('dragover', e => e.preventDefault());
document.addEventListener('drop', e => {
  e.preventDefault();
  const file = e.dataTransfer.files[0];
  if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls'))) {
    handleFile(file);
  }
});

// ═══════════════════════════════════════════ EXPORTAR INFORME ════
document.getElementById('btn-export').addEventListener('click', exportarInforme);

function exportarInforme() {
  const btn = document.getElementById('btn-export');
  btn.textContent = '⏳ Generando…';
  btn.disabled = true;

  // Esperar que los gráficos SVG estén listos
  setTimeout(() => {
    // 1. Capturar SVGs de ApexCharts como imágenes base64
    const chartEls = document.querySelectorAll('.apexcharts-canvas svg');
    const chartPromises = [];

    chartEls.forEach(svg => {
      chartPromises.push(new Promise(resolve => {
        const svgData = new XMLSerializer().serializeToString(svg);
        const svgBlob = new Blob([svgData], { type: 'image/svg+xml;charset=utf-8' });
        const url = URL.createObjectURL(svgBlob);
        const img = new Image();
        const canvas = document.createElement('canvas');
        canvas.width = svg.viewBox.baseVal.width || svg.clientWidth || 400;
        canvas.height = svg.viewBox.baseVal.height || svg.clientHeight || 300;
        img.onload = () => {
          canvas.getContext('2d').drawImage(img, 0, 0);
          URL.revokeObjectURL(url);
          resolve({ svg, dataUrl: canvas.toDataURL('image/png') });
        };
        img.onerror = () => { URL.revokeObjectURL(url); resolve(null); };
        img.src = url;
      }));
    });

    Promise.all(chartPromises).then(results => {
      // 2. Clonar el documento
      const clone = document.documentElement.cloneNode(true);

      // 3. Reemplazar cada SVG de ApexCharts por un <img> con la imagen PNG
      results.forEach(res => {
        if (!res) return;
        const id = res.svg.closest('[id]')?.id;
        if (!id) return;
        const container = clone.querySelector('#' + id);
        if (container) {
          container.innerHTML = `<img src="${res.dataUrl}" style="width:100%;max-height:320px;object-fit:contain;" />`;
        }
      });

      // 4. Quitar elementos innecesarios del clone
      clone.querySelectorAll('#file-input,#file-input2,.loader-overlay,#empty-state,#btn-export,#btn-reset,.upload-btn,.upload-btn-big,.pagination,#loader').forEach(el => el.remove());

      // 5. Convertir logo a base64
      const logoImg = document.getElementById('logo-img');
      if (logoImg && logoImg.src && logoImg.src.startsWith('http')) {
        // logo ya cargado — lo convertimos
        try {
          const c2 = document.createElement('canvas');
          c2.width = logoImg.naturalWidth || 140;
          c2.height = logoImg.naturalHeight || 40;
          c2.getContext('2d').drawImage(logoImg, 0, 0);
          const cloneLogo = clone.querySelector('#logo-img');
          if (cloneLogo) cloneLogo.src = c2.toDataURL('image/png');
        } catch(e) { /* CORS — dejamos src como está */ }
      }

      // 6. Incrustar CSS inline
      let cssText = '';
      Array.from(document.styleSheets).forEach(ss => {
        try {
          Array.from(ss.cssRules || []).forEach(rule => { cssText += rule.cssText + '\n'; });
        } catch(e) {}
      });

      // 7. Construir HTML final
      const fecha = document.getElementById('fecha-corte-val').textContent;
      const htmlContent = `<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1"/>
  <title>BCS · Dashboard Corresponsales · Corte ${fecha}</title>
  <link href="https://fonts.googleapis.com/css2?family=Sora:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet"/>
  <style>${cssText}</style>
</head>
<body>
${clone.querySelector('header').outerHTML}
${clone.querySelector('main').outerHTML}
</body>
</html>`;

      // 8. Descargar
      const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = `Dashboard_BCS_Corresponsales_${fecha.replace(/[\/\-\s]/g,'_')}.html`;
      a.click();
      URL.revokeObjectURL(a.href);

      btn.innerHTML = '<span>⬇</span> Descargar Informe';
      btn.disabled = false;
    });
  }, 600); // pequeño delay para que los gráficos terminen de renderizar
}
