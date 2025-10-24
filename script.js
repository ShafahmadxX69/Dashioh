// Unified Viewer for IN + ExportSchedule + ERP (improvised & robust)
// - Single config block
// - Loads multiple GIDs (skips empty ones)
// - Normalizes columns across sheets
// - Merges by Part No (if available) and exposes mergedData for charts/tables
// - Uses Chart.js (same rendering functions as you had)

// === CONFIG ===
const SPREADSHEET_ID = "1XoV7020NTZk1kzqn3F2ks3gOVFJ5arr5NVgUdewWPNQ";

// Sheet GIDs (fill the xxxxx with real values when ready)
const GID_IN = "1100244896";
const GID_EXPSCHED = "xxxxxxxx"; // isi nanti
const GID_ERP = "xxxxxxxx";      // isi nanti

// Palette & globals
const palette = ['#7c3aed','#00ffe1','#10b981','#ff7ab6','#f59e0b','#60a5fa'];
let prodChart = null, shiftChart = null, expChart = null;

// Data stores
let dataIN = [];
let dataExp = [];
let dataERP = [];
let mergedData = []; // final data used by UI

// Helper to build GViz URL
function gviz(gid){
  if(!gid) return null;
  return `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=out:json&gid=${gid}`;
}

// ====== INIT ======
document.addEventListener('DOMContentLoaded', async () => {
  try {
    await loadAllSheets();
    processMergedData();
    initUI();
    renderAll();
  } catch (err) {
    console.error('Init error:', err);
    alert('Gagal inisialisasi. Cek console untuk detail.');
  }

  // UI event bindings (assumes these elements exist in HTML)
  const search = document.getElementById('globalSearch');
  if(search) search.addEventListener('keyup', (e)=>{ if(e.key==='Enter') renderAll(); });

  const brandSelect = document.getElementById('brandSelect');
  if(brandSelect) brandSelect.addEventListener('change', renderAll);

  const timeRange = document.getElementById('timeRange');
  if(timeRange) timeRange.addEventListener('change', renderAll);

  const refreshBtn = document.getElementById('refreshBtn');
  if(refreshBtn) refreshBtn.addEventListener('click', () => { refreshAll(); });
});

// ====== LOAD SHEETS ======
async function loadAllSheets(){
  // load each if GID provided, otherwise set empty array
  dataIN = await safeFetchGViz(GID_IN);
  dataExp = await safeFetchGViz(GID_EXPSCHED);
  dataERP = await safeFetchGViz(GID_ERP);
}

// wrapper that returns [] on missing/failed fetch
async function safeFetchGViz(gid){
  if(!gid || gid.startsWith('xxx')) return []; // skip if placeholder
  try {
    return await fetchGViz(gid);
  } catch(e){
    console.warn('Failed to fetch gid', gid, e);
    return [];
  }
}

// universal GViz loader -> returns array of objects keyed by column label/id
async function fetchGViz(gid){
  const url = gviz(gid);
  if(!url) throw new Error('No GViz URL');
  const res = await fetch(url);
  const text = await res.text();
  const jsonText = text.replace(/^[^\(]*\(/,'').replace(/\);?$/,'');
  const obj = JSON.parse(jsonText);
  const cols = (obj.table.cols || []).map(c => (c.label || c.id || '').toString().trim());
  const rows = obj.table.rows || [];
  return rows.map(r => {
    const out = {};
    (r.c || []).forEach((cell, i) => {
      const key = cols[i] || `col${i}`;
      out[key] = (cell ? (cell.v !== undefined ? cell.v : (cell.f !== undefined ? cell.f : null)) : null);
    });
    return out;
  });
}

// ====== PROCESS / MERGE RULES ======
function processMergedData(){
  // Create initial normalized rows from dataIN (primary source)
  mergedData = dataIN.map(r => {
    const normalized = normalizeINRow(r);
    normalized.InvoiceData = {}; // placeholder
    normalized.ERPData = {};
    return normalized;
  });

  // If no rows from IN, but there are ERP or Exp rows, try to use them as base
  if(mergedData.length === 0){
    // fallback: try using dataExp
    if(dataExp.length) mergedData = dataExp.map(r => normalizeExpRow(r));
    else if(dataERP.length) mergedData = dataERP.map(r => normalizeERPRow(r));
  }

  // Merge Exp schedule by PART. NO. / PartNo / Part Number (flexible)
  dataExp.forEach(exp => {
    const expKey = (exp['PART. NO.'] ?? exp['PartNo'] ?? exp['Part No'] ?? exp['PART NO'] ?? '').toString().trim();
    if(!expKey) return;
    mergedData.forEach(row => {
      const rowKey = (row.PartNo ?? row.PartNoAlt ?? '').toString().trim();
      if(rowKey && rowKey === expKey){
        row.InvoiceData = Object.assign({}, row.InvoiceData || {}, exp);
      }
    });
  });

  // Merge ERP similarly
  dataERP.forEach(erp => {
    const erpKey = (erp['PART. NO.'] ?? erp['PartNo'] ?? erp['Part No'] ?? erp['PART NO'] ?? '').toString().trim();
    if(!erpKey) return;
    mergedData.forEach(row => {
      const rowKey = (row.PartNo ?? row.PartNoAlt ?? '').toString().trim();
      if(rowKey && rowKey === erpKey){
        row.ERPData = Object.assign({}, row.ERPData || {}, erp);
      }
    });
  });

  // As a final step normalize date strings and numeric fields
  mergedData.forEach(r => {
    r.DateISO = excelToISO(r.DateISO || r.Date || r['入庫日期'] || r['Date'] || r._Date || '');
    r.Qty = Number(r.Qty || 0) || 0;
    r.Rework = Number(r.Rework || 0) || 0;
  });
}

// Normalizers - attempt many common column names to make script tolerant
function normalizeINRow(r){
  return {
    DateISO: r['入庫日期'] ?? r['Date'] ?? r['date'] ?? r['Tanggal'] ?? r['tanggal'] ?? '',
    Brand: r['客戶'] ?? r['Customer'] ?? r['Brand'] ?? r['brand'] ?? r['Merk'] ?? r['merchant'] ?? '',
    PartNo: r['產品料號'] ?? r['PART NO'] ?? r['PART. NO.'] ?? r['PartNo'] ?? r['Part No'] ?? r['料號'] ?? '',
    PartNoAlt: r['SKU'] ?? r['SKU NO'] ?? '',
    Qty: r['入庫數'] ?? r['Qty'] ?? r['QTY'] ?? r['Quantity'] ?? r['Jumlah'] ?? 0,
    Rework: r['Rework QTY'] ?? r['Rework'] ?? r['rework'] ?? 0,
    Shift: r['Shift'] ?? r['shift'] ?? '',
    Line: r['Line'] ?? r['line'] ?? '',
    _raw: r
  };
}
function normalizeExpRow(r){
  return {
    DateISO: r['Date'] ?? r['Ship Date'] ?? '',
    Brand: r['Customer'] ?? r['Brand'] ?? '',
    PartNo: r['PART. NO.'] ?? r['PART NO'] ?? r['PartNo'] ?? '',
    Qty: r['Qty'] ?? r['Quantity'] ?? 0,
    Rework: 0,
    Shift: '',
    Line: '',
    InvoiceData: r,
    _raw: r
  };
}
function normalizeERPRow(r){
  return {
    DateISO: r['Date'] ?? '',
    Brand: r['Customer'] ?? r['Brand'] ?? '',
    PartNo: r['PART. NO.'] ?? r['PartNo'] ?? '',
    Qty: r['Qty'] ?? 0,
    Rework: 0,
    Shift: '',
    Line: '',
    ERPData: r,
    _raw: r
  };
}

// ====== DATE / UTILS ======
function excelToISO(val){
  if(!val && val !== 0) return '';
  if(typeof val === 'number'){
    // treat as excel serial
    const d = new Date(Math.round((val - 25569) * 86400 * 1000));
    if(!isNaN(d)) return d.toISOString().slice(0,10);
    return String(val);
  }
  // if it's already ISO-like
  if(typeof val === 'string'){
    const s = val.trim();
    // try parse common formats
    const parsed = new Date(s);
    if(!isNaN(parsed)) return parsed.toISOString().slice(0,10);
    // if string looks like dd/mm/yyyy or dd-mm-yyyy
    const m = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
    if(m){
      const dd = Number(m[1]), mm = Number(m[2]), yy = Number(m[3]);
      const y = yy < 100 ? (yy > 50 ? 1900+yy : 2000+yy) : yy;
      const d = new Date(y, mm-1, dd);
      if(!isNaN(d)) return d.toISOString().slice(0,10);
    }
    // fallback: first 10 chars
    return s.slice(0,10);
  }
  return '';
}

// ====== UI INIT & RENDER ALL ======
function initUI(){
  // populate brand select from mergedData
  const select = document.getElementById('brandSelect');
  if(!select) return;
  const brands = Array.from(new Set(mergedData.map(r => (r.Brand||'').toString().trim()).filter(Boolean))).sort();
  select.innerHTML = '<option value="">-- All Brands --</option>';
  brands.forEach(b => {
    const o = document.createElement('option'); o.value = b; o.textContent = b; select.appendChild(o);
  });
}

function renderAll(){
  const filters = getFilters();
  const filtered = filterData(filters);
  updateKPIs(filtered);
  drawProdChart(filtered);
  drawShiftChart(filtered);
  drawExpChart(filtered);
  populateTable(filtered);
}

// ====== FILTERS ======
function getFilters(){
  const select = document.getElementById('brandSelect');
  const brands = select ? (select.value ? [select.value] : []) : [];
  const qNode = document.getElementById('globalSearch');
  const q = qNode ? qNode.value.trim().toLowerCase() : '';
  const rangeNode = document.getElementById('timeRange');
  const range = rangeNode ? rangeNode.value : '1M';
  return { brands, q, range };
}

function filterData({brands, q, range}){
  let arr = mergedData.slice();
  if(brands && brands.length) arr = arr.filter(r => r.Brand && brands.includes(String(r.Brand).trim()));
  if(q) {
    arr = arr.filter(r => {
      const hay = (r.Brand || '') + ' ' + JSON.stringify(r._raw || r.InvoiceData || r.ERPData || {});
      return hay.toLowerCase().includes(q);
    });
  }
  // apply date range
  const today = new Date();
  if(range === '1D'){
    const iso = today.toISOString().slice(0,10);
    arr = arr.filter(r => (r.DateISO||'') === iso);
  } else if(range === '1W'){
    const then = new Date(); then.setDate(today.getDate()-6);
    arr = arr.filter(r => r.DateISO && new Date(r.DateISO) >= then);
  } else if(range === '1M'){
    const then = new Date(); then.setDate(today.getDate()-29);
    arr = arr.filter(r => r.DateISO && new Date(r.DateISO) >= then);
  } else if(range === '1Y'){
    const then = new Date(); then.setFullYear(today.getFullYear()-1);
    arr = arr.filter(r => r.DateISO && new Date(r.DateISO) >= then);
  }
  return arr;
}

// ====== KPIs ======
function updateKPIs(data){
  const todayISO = new Date().toISOString().slice(0,10);
  const todaySum = data.filter(d => d.DateISO === todayISO).reduce((s,i)=>s+(i.Qty||0),0) || 0;
  const monthSum = data.reduce((s,i)=>s+(i.Qty||0),0) || 0;
  const totalRework = data.reduce((s,i)=>s+(i.Rework||0),0);
  const reworkRate = monthSum > 0 ? (totalRework / monthSum * 100) : 0;
  const brandCounts = {};
  data.forEach(d => { if(d.Brand) brandCounts[d.Brand] = (brandCounts[d.Brand]||0) + (d.Qty||0); });
  const topBrand = Object.entries(brandCounts).sort((a,b)=>b[1]-a[1])[0];

  const kpiProd = document.getElementById('kpiProd');
  const kpiExp = document.getElementById('kpiExp');
  const kpiRework = document.getElementById('kpiRework');
  const kpiBrand = document.getElementById('kpiBrand');

  if(kpiProd) kpiProd.textContent = todaySum;
  if(kpiExp) kpiExp.textContent = monthSum;
  if(kpiRework) kpiRework.textContent = reworkRate.toFixed(2) + '%';
  if(kpiBrand) kpiBrand.textContent = topBrand ? topBrand[0] : '-';
}

// ====== CHARTS & TABLE (adapted from original) ======
function drawProdChart(data){
  const range = document.getElementById('timeRange') ? document.getElementById('timeRange').value : '1M';
  const labels = generateLabels(range);
  const brands = Array.from(new Set(data.map(d=>d.Brand).filter(Boolean))).slice(0,6);
  const datasets = brands.map((b, idx) => {
    const map = labels.map(l => data.filter(r => r.Brand===b && (r.DateISO||'').startsWith(l)).reduce((s,i)=>s+(i.Qty||0),0));
    return {
      label: b,
      data: map,
      borderColor: palette[idx%palette.length],
      backgroundColor: hexToRGBA(palette[idx%palette.length], 0.12),
      tension: 0.25,
      fill: true
    };
  });

  const ctx = document.getElementById('prodChart');
  if(!ctx) return;
  if(prodChart) prodChart.destroy();
  prodChart = new Chart(ctx, {
    type: 'line',
    data: { labels, datasets },
    options: { plugins:{legend:{position:'bottom'}}, scales:{y:{beginAtZero:true}}}
  });
}

function drawShiftChart(data){
  const days = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'];
  const labels = days;
  const shiftA = labels.map(()=>0);
  const shiftB = labels.map(()=>0);

  data.forEach(d => {
    if(!d.DateISO) return;
    const dt = new Date(d.DateISO);
    if(isNaN(dt)) return;
    const idx = (dt.getDay()+6)%7; // Monday->0 ... Sunday->6
    const qty = d.Qty || 0;
    if(String(d.Shift).toUpperCase().includes('A')) shiftA[idx] += qty;
    else shiftB[idx] += qty;
  });

  const ctx = document.getElementById('shiftChart');
  if(!ctx) return;
  if(shiftChart) shiftChart.destroy();
  shiftChart = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets: [
      { label: 'Shift A', data: shiftA, backgroundColor: palette[0] },
      { label: 'Shift B', data: shiftB, backgroundColor: palette[1] }
    ] },
    options: { responsive:true, plugins:{legend:{position:'bottom'}} }
  });
}

function drawExpChart(data){
  const brandMap = {};
  data.forEach(d => {
    const b = d.Brand || 'Unknown';
    brandMap[b] = (brandMap[b] || 0) + (d.Qty || 0);
  });
  const labels = Object.keys(brandMap);
  const values = labels.map(l => brandMap[l]);
  const ctx = document.getElementById('expChart');
  if(!ctx) return;
  if(expChart) expChart.destroy();
  expChart = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets: [{ label:'Export Qty', data: values, backgroundColor: labels.map((_,i)=>palette[i%palette.length]) }] },
    options: { indexAxis: 'y', plugins:{legend:{display:false}} }
  });
}

function populateTable(data){
  const tbody = document.querySelector('#dataTable tbody');
  if(!tbody) return;
  tbody.innerHTML = '';
  const rows = (data || mergedData).slice(0,200);
  rows.forEach(r => {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${r.DateISO||''}</td>
                    <td>${escapeHtml(r.Brand||'')}</td>
                    <td>${escapeHtml(r.Shift||'')}</td>
                    <td>${escapeHtml(r.Line||'')}</td>
                    <td>${r.Qty||0}</td>
                    <td>${r.Rework||0}</td>
                    <td>${escapeHtml(r.PartNo||'')}</td>`;
    tbody.appendChild(tr);
  });
}

// ====== Helpers ======
function generateLabels(range){
  const today = new Date();
  const out = [];
  if(range==='1D'){ out.push(today.toISOString().slice(0,10)); return out; }
  if(range==='1W'){ for(let i=6;i>=0;i--){ const d=new Date(); d.setDate(today.getDate()-i); out.push(d.toISOString().slice(0,10)); } return out; }
  if(range==='1M'){ for(let i=29;i>=0;i--){ const d=new Date(); d.setDate(today.getDate()-i); out.push(d.toISOString().slice(0,10)); } return out; }
  if(range==='1Y'){ for(let i=11;i>=0;i--){ const d=new Date(today.getFullYear(), today.getMonth()-i, 1); out.push(d.toISOString().slice(0,7)); } return out; }
  return out;
}

function hexToRGBA(hex, a=0.15){
  const c = hex.replace('#','');
  const bigint = parseInt(c,16);
  const r = (bigint>>16)&255, g=(bigint>>8)&255, b=bigint&255;
  return `rgba(${r},${g},${b},${a})`;
}

function escapeHtml(text){
  if(!text) return '';
  return String(text).replace(/[&<>"'`]/g, function(match){ return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;','`':'&#96;'})[match]; });
}

// UI helpers exposed to HTML buttons
function applyFilters(){ renderAll(); }
function refreshAll(){ loadAllSheets().then(()=>{ processMergedData(); initUI(); renderAll(); }).catch(e=>{ console.error(e); alert('Refresh gagal'); }); }
function toggleTheme(){ document.body.classList.toggle('dark'); }

