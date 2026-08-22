/* ============================================================
   GALEGAL SAS — Dashboard de Facturación
   Lógica de carga de datos, normalización y gráficos.
   ============================================================ */

// -------------------- Paleta de gráficos --------------------
const COLORS = {
  navy:   '#182b65',
  navy2:  '#324a92',
  gold:   '#dcb970',
  gold2:  '#c69d59',
  goldLt: '#f0e194',
  red:    '#a8384a',
  green:  '#3c7a5e',
  slate:  '#7d88a8',
  orange: '#e08e3e',
};

const MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

// -------------------- Normalización "Tipo de Proceso" --------------------
// Mapa de valores (sin tildes, en mayúsculas, recortados) -> valor normalizado final
const TIPO_PROCESO_MAP = {
  'ACTUACIONES': 'ACTUACION',
  'AJUSTES': 'AJUSTE',
  'FISCAL': 'FISCALES',
  'JUDICIAL': 'JUDICIALES',
  'AUDIENCIA PREJUDICIAL': 'AUDIENCIAS PREJUDICIALES',
  'AUDIENCIA EXTRAJUDICIAL': 'AUDIENCIAS PREJUDICIALES',
  'CONCILIACION PREJUDICIAL': 'AUDIENCIAS PREJUDICIALES',
  'CONCILIACION EXTRAJUDICIAL': 'AUDIENCIAS PREJUDICIALES',
};

function stripAccents(str){
  return String(str).normalize('NFD').replace(/[\u0300-\u036f]/g,'');
}

function normalizeTipoProceso(raw){
  if(raw === null || raw === undefined || String(raw).trim() === '') return 'SIN CLASIFICAR';
  const key = stripAccents(String(raw)).trim().toUpperCase().replace(/\s+/g,' ');
  return TIPO_PROCESO_MAP[key] || key;
}

// -------------------- Estado de la factura desde "N. Factura" --------------------
// GA / GAXXXX  -> PAGADA
// NC / NCV...  -> DEVUELTA (cubre NC y NCV)
function estadoFromNFactura(nFactura){
  if(!nFactura) return null;
  const val = String(nFactura).trim().toUpperCase().replace(/\s+/g,'');
  if(val.startsWith('GA')) return 'PAGADA';
  if(val.startsWith('NC')) return 'DEVUELTA'; // cubre NC y NCV
  return null; // filas sin patrón reconocido se excluyen del conteo de estado
}

// -------------------- Estado global --------------------
let RAW_ROWS = [];       // filas crudas ya normalizadas y listas para filtrar
let charts = {};         // instancias Chart.js activas

// -------------------- Utilidades de formato --------------------
const fmtInt = new Intl.NumberFormat('es-CO');
const fmtMoney = new Intl.NumberFormat('es-CO', {maximumFractionDigits:0});

function money(n){
  return '$' + fmtMoney.format(Math.round(n || 0));
}

// ============================================================
// CARGA DE DATOS
// ============================================================

async function loadData(){
  setStatus('Cargando datos…', false);
  hideUploadPanel();

  const mode = DASHBOARD_CONFIG.DATA_SOURCE_MODE;

  try{
    let arrayBuffer;

    if(mode === 'sharepoint_rest'){
      arrayBuffer = await fetchFromSharePointRest();
    } else if(mode === 'direct_url'){
      arrayBuffer = await fetchDirectUrl(DASHBOARD_CONFIG.DIRECT_FILE_URL);
    } else {
      // modo manual: no intentamos fetch automático
      showUploadPanel();
      setStatus('Selecciona un archivo Excel para comenzar.', false);
      return;
    }

    processWorkbookArrayBuffer(arrayBuffer);
    setStatus('Datos cargados · ' + new Date().toLocaleString('es-CO'), true);

  } catch(err){
    console.error('Error cargando datos automáticamente:', err);
    setStatus('No se pudo conectar automáticamente. Carga el archivo manualmente.', false);
    showUploadPanel();
  }
}

async function fetchFromSharePointRest(){
  const site = DASHBOARD_CONFIG.SHAREPOINT_SITE_URL || window.location.origin;
  const relUrl = encodeURIComponent(DASHBOARD_CONFIG.SHAREPOINT_FILE_SERVER_RELATIVE_URL);
  const endpoint = `${site}/_api/web/GetFileByServerRelativeUrl('${DASHBOARD_CONFIG.SHAREPOINT_FILE_SERVER_RELATIVE_URL}')/$value`;

  const res = await fetch(endpoint, {
    method: 'GET',
    headers: { 'Accept': 'application/octet-stream' },
    credentials: 'same-origin', // usa la sesión de SharePoint ya autenticada
  });
  if(!res.ok) throw new Error('SharePoint REST respondió ' + res.status);
  return await res.arrayBuffer();
}

async function fetchDirectUrl(url){
  const res = await fetch(url);
  if(!res.ok) throw new Error('Fetch directo respondió ' + res.status);
  return await res.arrayBuffer();
}

function processWorkbookArrayBuffer(arrayBuffer){
  const data = new Uint8Array(arrayBuffer);
  const workbook = XLSX.read(data, {type:'array', cellDates:true});
  const sheetName = workbook.SheetNames.includes(DASHBOARD_CONFIG.SHEET_NAME)
    ? DASHBOARD_CONFIG.SHEET_NAME
    : workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const json = XLSX.utils.sheet_to_json(sheet, {defval:null, raw:true});
  RAW_ROWS = buildRows(json);
  populateFilterOptions(RAW_ROWS);
  applyFiltersAndRender();
}

// Toma las filas crudas del Excel y construye el modelo que usa el dashboard
function buildRows(json){
  const rows = [];
  for(const r of json){
    const nFactura = r['N. Factura'];
    if(!nFactura) continue; // ignora filas vacías/sin número de factura

    const estado = estadoFromNFactura(nFactura);
    if(!estado) continue; // fila con formato de N. Factura no reconocido

    let fecha = r['Fecha Factura'];
    if(fecha && !(fecha instanceof Date)) fecha = new Date(fecha);
    if(!fecha || isNaN(fecha.getTime())) continue; // sin fecha válida no se puede filtrar por año/mes

    const valor = Number(r['Valor a pagar']) || 0;
    const aseguradora = (r['Aseguradora'] || 'SIN ASEGURADORA').toString().trim().toUpperCase();
    const tipoProceso = normalizeTipoProceso(r['Clase de proceso']);

    rows.push({
      nFactura: String(nFactura).trim(),
      estado,
      anio: fecha.getFullYear(),
      mes: fecha.getMonth(), // 0-11
      aseguradora,
      tipoProceso,
      valor,
    });
  }
  return rows;
}

// ============================================================
// FILTROS
// ============================================================

function populateFilterOptions(rows){
  const aseguradoras = [...new Set(rows.map(r => r.aseguradora))].sort();
  const anios = [...new Set(rows.map(r => r.anio))].sort((a,b)=>b-a);

  fillSelect('f-aseguradora', aseguradoras, 'Todas');
  fillSelect('f-anio', anios, 'Todos');
  fillSelectMeses();
}

function fillSelect(id, values, allLabel){
  const sel = document.getElementById(id);
  const current = sel.value;
  sel.innerHTML = `<option value="">${allLabel}</option>` +
    values.map(v => `<option value="${v}">${v}</option>`).join('');
  if(values.includes(current)) sel.value = current;
}

function fillSelectMeses(){
  const sel = document.getElementById('f-mes');
  const current = sel.value;
  sel.innerHTML = '<option value="">Todos</option>' +
    MESES.map((m,i) => `<option value="${i}">${m}</option>`).join('');
  if(current !== '') sel.value = current;
}

function getActiveFilters(){
  return {
    aseguradora: document.getElementById('f-aseguradora').value,
    anio: document.getElementById('f-anio').value,
    mes: document.getElementById('f-mes').value,
    estado: document.getElementById('f-estado').value,
  };
}

function filterRows(rows, f){
  return rows.filter(r =>
    (f.aseguradora === '' || r.aseguradora === f.aseguradora) &&
    (f.anio === '' || String(r.anio) === String(f.anio)) &&
    (f.mes === '' || String(r.mes) === String(f.mes)) &&
    (f.estado === '' || r.estado === f.estado)
  );
}

// ============================================================
// AGREGACIONES + RENDER
// ============================================================

function applyFiltersAndRender(){
  const f = getActiveFilters();
  const rows = filterRows(RAW_ROWS, f);
  renderKPIs(rows);
  renderCharts(rows);
}

function renderKPIs(rows){
  const total = rows.length;
  const pagadas = rows.filter(r => r.estado === 'PAGADA').length;
  const devueltas = rows.filter(r => r.estado === 'DEVUELTA').length;
  const valorTotal = rows.reduce((s,r) => s + r.valor, 0);

  document.getElementById('kpi-facturas').textContent = fmtInt.format(total);
  document.getElementById('kpi-pagadas').textContent = fmtInt.format(pagadas);
  document.getElementById('kpi-devueltas').textContent = fmtInt.format(devueltas);
  document.getElementById('kpi-valor').textContent = money(valorTotal);
}

function renderCharts(rows){
  renderPorMes(rows);
  renderAseguradora(rows);
  renderEstado(rows);
  renderTipoProceso(rows);
}

function baseChartOptions(extra){
  return Object.assign({
    responsive:true,
    maintainAspectRatio:false,
    plugins:{
      legend:{labels:{color:'#1c2340', font:{family:'Inter', size:11}}},
    },
  }, extra || {});
}

function renderPorMes(rows){
  const counts = new Array(12).fill(0);
  rows.forEach(r => counts[r.mes]++);

  upsertChart('chart-mes', 'bar', {
    labels: MESES.map(m => m.slice(0,3)),
    datasets:[{
      label:'Facturas',
      data: counts,
      backgroundColor: COLORS.navy2,
      borderRadius:4,
      maxBarThickness:42,
    }]
  }, baseChartOptions({
    plugins:{legend:{display:false}},
    scales:{
      x:{ticks:{color:'#5b6688', font:{family:'Inter',size:10}}, grid:{display:false}},
      y:{ticks:{color:'#5b6688', precision:0}, grid:{color:'#eee8d8'}, beginAtZero:true},
    }
  }));
}

function renderAseguradora(rows){
  const map = {};
  rows.forEach(r => { map[r.aseguradora] = (map[r.aseguradora]||0) + 1; });
  const entries = Object.entries(map).sort((a,b)=>b[1]-a[1]);

  upsertChart('chart-aseguradora', 'doughnut', {
    labels: entries.map(e=>e[0]),
    datasets:[{
      data: entries.map(e=>e[1]),
      backgroundColor: palette(entries.length),
      borderColor:'#f8f6ef',
      borderWidth:2,
    }]
  }, baseChartOptions({
    plugins:{legend:{position:'right', labels:{boxWidth:12, font:{family:'Inter',size:10}}}}
  }));
}

function renderEstado(rows){
  const pagadas = rows.filter(r=>r.estado==='PAGADA').length;
  const devueltas = rows.filter(r=>r.estado==='DEVUELTA').length;

  upsertChart('chart-estado', 'doughnut', {
    labels:['Pagada','Devuelta'],
    datasets:[{
      data:[pagadas, devueltas],
      backgroundColor:[COLORS.green, COLORS.red],
      borderColor:'#f8f6ef',
      borderWidth:2,
    }]
  }, baseChartOptions({
    plugins:{legend:{position:'right'}}
  }));
}

function renderTipoProceso(rows){
  const map = {};
  rows.forEach(r => { map[r.tipoProceso] = (map[r.tipoProceso]||0) + 1; });
  const entries = Object.entries(map).sort((a,b)=>b[1]-a[1]);

  upsertChart('chart-tipo', 'bar', {
    labels: entries.map(e=>e[0]),
    datasets:[{
      label:'Facturas',
      data: entries.map(e=>e[1]),
      backgroundColor: palette(entries.length),
      borderRadius:4,
    }]
  }, baseChartOptions({
    indexAxis:'y',
    plugins:{legend:{display:false}},
    scales:{
      x:{ticks:{color:'#5b6688', precision:0}, grid:{color:'#eee8d8'}, beginAtZero:true},
      y:{ticks:{color:'#1c2340', font:{family:'Inter', size:10, weight:'600'}}, grid:{display:false}},
    }
  }));
}

function palette(n){
  const base = [COLORS.navy, COLORS.gold2, COLORS.red, COLORS.navy2, COLORS.orange, COLORS.green, COLORS.slate, COLORS.goldLt];
  const out = [];
  for(let i=0;i<n;i++) out.push(base[i % base.length]);
  return out;
}

function upsertChart(canvasId, type, data, options){
  const ctx = document.getElementById(canvasId).getContext('2d');
  if(charts[canvasId]){
    charts[canvasId].data = data;
    charts[canvasId].options = options;
    charts[canvasId].update();
  } else {
    charts[canvasId] = new Chart(ctx, {type, data, options});
  }
}

// ============================================================
// UI: estado de carga / panel de subida manual
// ============================================================

function setStatus(text, ok){
  document.getElementById('data-status-text').textContent = text;
  document.getElementById('data-status').style.color = ok ? '#f0e194' : '#e6b8b8';
  document.getElementById('btn-reload').style.display = 'inline-block';
  document.getElementById('footer-updated').textContent = ok ? ('Última actualización: ' + new Date().toLocaleString('es-CO')) : '';
}

function showUploadPanel(){ document.getElementById('upload-panel').style.display = 'flex'; }
function hideUploadPanel(){ document.getElementById('upload-panel').style.display = 'none'; }

function handleManualFile(file){
  const reader = new FileReader();
  reader.onload = (e) => {
    try{
      processWorkbookArrayBuffer(e.target.result);
      setStatus('Datos cargados desde archivo local · ' + file.name, true);
      hideUploadPanel();
    }catch(err){
      console.error(err);
      setStatus('El archivo no pudo leerse. Verifica el formato.', false);
    }
  };
  reader.readAsArrayBuffer(file);
}

// ============================================================
// EVENTOS
// ============================================================

document.addEventListener('DOMContentLoaded', () => {
  ['f-aseguradora','f-anio','f-mes','f-estado'].forEach(id => {
    document.getElementById(id).addEventListener('change', applyFiltersAndRender);
  });

  document.getElementById('btn-reset').addEventListener('click', () => {
    ['f-aseguradora','f-anio','f-mes','f-estado'].forEach(id => document.getElementById(id).value = '');
    applyFiltersAndRender();
  });

  document.getElementById('btn-reload').addEventListener('click', loadData);

  const fileInput = document.getElementById('file-input');
  fileInput.addEventListener('change', (e) => {
    if(e.target.files[0]) handleManualFile(e.target.files[0]);
  });

  const dropZone = document.getElementById('upload-drop');
  dropZone.addEventListener('dragover', (e)=>{ e.preventDefault(); dropZone.style.background = '#f1ecd8'; });
  dropZone.addEventListener('dragleave', ()=>{ dropZone.style.background = ''; });
  dropZone.addEventListener('drop', (e)=>{
    e.preventDefault();
    dropZone.style.background = '';
    if(e.dataTransfer.files[0]) handleManualFile(e.dataTransfer.files[0]);
  });

  loadData();

  if(DASHBOARD_CONFIG.DATA_SOURCE_MODE !== 'manual' && DASHBOARD_CONFIG.AUTO_REFRESH_MINUTES > 0){
    setInterval(loadData, DASHBOARD_CONFIG.AUTO_REFRESH_MINUTES * 60 * 1000);
  }
});
