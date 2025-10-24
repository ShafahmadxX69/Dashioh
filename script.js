// CONFIG: change these if needed
const SPREADSHEET_ID = '1XoV7020NTZk1kzqn3F2ks3gOVFJ5arr5NVgUdewWPNQ';
const GID = '1100244896'; // sheet gid you provided
const SHEET_GVIZ_URL = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=out:json&gid=${GID}`;

// global chart refs
let prodChart, shiftChart, expChart;
let rawData = []; // array of objects from sheet
const palette = ['#7c3aed','#00ffe1','#10b981','#ff7ab6','#f59e0b','#60a5fa'];

document.addEventListener('DOMContentLoaded', () => {
  loadSheetData().then(() => {
    initUI();
    renderAll();
  }).catch(err=>{
    console.error('Failed to load sheet:', err);
    alert('Gagal memuat sheet. Pastikan spreadsheet sudah di-publish (File > Publish to web).');
  });

  document.getElementById('globalSearch').addEventListener('keyup', (e)=>{
    if(e.key === 'Enter') renderAll();
  });
});

async function loadSheetData(){
  // fetch the GViz JSON (it's wrapped in a function call, so we remove prefix/suffix)
  const res = await fetch(SHEET_GVIZ_URL);
  const text = await res.text();
  const jsonText = text.replace(/^[^\(]*\(/,'').replace(/\);?$/,'');
  const obj = JSON.parse(jsonText);
  const cols = obj.table.cols.map(c => (c.label || c.id || '').toString().trim());
  const rows = obj.table.rows || [];

  rawData = rows.map(r => {
    const out = {};
    r.c.forEach((cell, idx) => {
      const key = cols[idx] || `col${idx}`;
      out[key] = cell ? (cell.v !== undefined ? cell.v : (cell.f !== undefined ? cell.f : null)) : null;
    });
    return out;
  });

  // Normalize common column names if present
  rawData = rawData.map(r => {
    const norm = {};
    // try common keys and mapping fallback
    norm.Date = r['Date'] ?? r['date'] ?? r['Tanggal'] ?? r['tanggal'] ?? r['Col1'] ?? r[Object.keys(r)[0]];
    norm.Brand = r['Brand'] ?? r['brand'] ?? r['Merk'] ?? r['merk'] ?? r['Brand Name'] ?? r['invoice_brand'] ?? '';
    norm.Shift = r['Shift'] ?? r['shift'] ?? r['Shift A/B'] ?? '';
    norm.Line = r['Line'] ?? r['line'] ?? r['Line Number'] ?? '';
    // Qty might be in many columns
    norm.Qty = Number(r['Qty'] ?? r['qty'] ?? r['Quantity'] ?? r['Jumlah'] ?? r['jumlah'] ?? r['QTY'] ?? 0) || 0;
    norm.Rework = Number(r['Rework'] ?? r['rework'] ?? r['Repair'] ?? r['Perbaikan'] ?? 0) || 0;
    norm.ReworkFixed = Number(r['Rework Fixed'] ?? r['ReworkFixed'] ?? r['reworkfixed'] ?? r['Fixed'] ?? 0) || 0;
    // keep original
    norm._raw = r;
    return norm;
  });

  // convert Date to ISO-date string where possible
  rawData.forEach(r=>{
    let d = r.Date;
    if(typeof d === 'number'){ // Google may give serial number -> treat as date
      // Google sometimes returns epoch-ms; if it's big, try to parse
      const date = new Date(Math.round((d - 25569) * 86400 * 1000)); // excel serial -> js date (attempt)
      if(!isNaN(date)) r.DateISO = date.toISOString().slice(0,10);
      else r.DateISO = String(d);
    } else if (typeof d === 'string') {
      // try parse string
      const parsed = new Date(d);
      if(!isNaN(parsed)) r.DateISO = parsed.toISOString().slice(0,10);
      else r.DateISO = d.slice(0,10);
    } else {
      r.DateISO = '';
    }
  });
}

function initUI(){
  // fill brand select
  const brands = Array.from(new Set(rawData.map(r=>String(r.Brand || '').trim()).filter(Boolean))).sort();
  const select = document.getElementById('brandSelect');
  select.innerHTML = '';
  brands.forEach(b=>{
    const o = document.createElement('option'); o.value=b; o.textContent=b; select.appendChild(o);
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

function getFilters(){
  const brands = Array.from(document.getElementById('brandSelect').selectedOptions).map(o=>o.value);
  const q = document.getElementById('globalSearch').value.trim().toLowerCase();
  const range = document.getElementById('timeRange').value;
  return { brands, q, range };
}

function filterData({brands,q,range}){
  let arr = rawData.slice();
  if(brands && brands.length) arr = arr.filter(r=> r.Brand && brands.includes(String(r.Brand).trim()));
  if(q) arr = arr.filter(r=> {
    return (String(r.Brand||'') + ' ' + String(r._raw ? JSON.stringify(r._raw) : '')).toLowerCase().includes(q);
  });
  // apply range (approx)
  const today = new Date();
  if(range === '1D'){
    const iso = today.toISOString().slice(0,10); arr = arr.filter(r=>r.DateISO === iso);
  } else if(range === '1W'){
    const then = new Date(); then.setDate(today.getDate()-6);
    arr = arr.filter(r=>{
      if(!r.DateISO) return false;
      return new Date(r.DateISO) >= then;
    });
  } else if(range === '1M'){
    const then = new Date(); then.setDate(today.getDate()-29);
    arr = arr.filter(r=> r.DateISO && new Date(r.DateISO) >= then);
  } else if(range === '1Y'){
    const then = new Date(); then.setFullYear(today.getFullYear()-1);
    arr = arr.filter(r=> r.DateISO && new Date(r.DateISO) >= then);
  }
  return arr;
}

function updateKPIs(data){
  const todayISO = new Date().toISOString().slice(0,10);
  const todaySum = data.filter(d=>d.DateISO===todayISO).reduce((s,i)=>s+(i.Qty||0),0) || 0;
  const monthSum = data.reduce((s,i)=>s+(i.Qty||0),0) || 0;
  const totalRework = data.reduce((s,i)=>s+(i.Rework||0),0);
  const reworkRate = monthSum > 0 ? (totalRework / monthSum * 100) : 0;
  const brandCounts = {};
  data.forEach(d=>{ if(d.Brand) brandCounts[d.Brand] = (brandCounts[d.Brand]||0) + (d.Qty||0) });
  const topBrand = Object.entries(brandCounts).sort((a,b)=>b[1]-a[1])[0];
  document.getElementById('kpiProd').textContent = todaySum;
  document.getElementById('kpiExp').textContent = monthSum;
  document.getElementById('kpiRework').textContent = reworkRate.toFixed(2) + '%';
  document.getElementById('kpiBrand').textContent = topBrand ? topBrand[0] : '-';
}

function drawProdChart(data){
  // group by date and brand
  const range = document.getElementById('timeRange').value;
  const labels = generateLabels(range);
  const brands = Array.from(new Set(data.map(d=>d.Brand).filter(Boolean))).slice(0,6);
  const datasets = brands.map((b, idx)=>{
    const map = labels.map(l => {
      const sum = data.filter(r=>r.Brand===b && (r.DateISO||'').startsWith(l)).reduce((s,i)=>s+(i.Qty||0),0);
      return sum;
    });
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
  if(prodChart) prodChart.destroy();
  prodChart = new Chart(ctx, {
    type: 'line',
    data: { labels, datasets },
    options: { plugins:{legend:{position:'bottom'}}, scales:{y:{beginAtZero:true}}}
  });
}

function drawShiftChart(data){
  // group last 7 days by day with shift A/B totals
  const days = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'];
  const labels = days;
  const shiftA = labels.map((_,i)=>0);
  const shiftB = labels.map((_,i)=>0);
  // map day index from DateISO
  data.forEach(d=>{
    if(!d.DateISO) return;
    const dt = new Date(d.DateISO);
    const idx = (dt.getDay()+6)%7; // shift Sunday -> index 6
    const qty = d.Qty || 0;
    if(String(d.Shift).toUpperCase().includes('A')) shiftA[idx] += qty;
    else shiftB[idx] += qty;
  });

  const ctx = document.getElementById('shiftChart');
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
  // group total qty by brand
  const brandMap = {};
  data.forEach(d=>{
    const b = d.Brand || 'Unknown';
    brandMap[b] = (brandMap[b]||0) + (d.Qty||0);
  });
  const labels = Object.keys(brandMap);
  const values = labels.map(l=>brandMap[l]);
  const ctx = document.getElementById('expChart');
  if(expChart) expChart.destroy();
  expChart = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets: [{ label:'Export Qty', data: values, backgroundColor: labels.map((_,i)=>palette[i%palette.length]) }] },
    options: { indexAxis: 'y', plugins:{legend:{display:false}} }
  });
}

function populateTable(data){
  const tbody = document.querySelector('#dataTable tbody');
  tbody.innerHTML = '';
  const rows = (data || rawData).slice(0,200);
  rows.forEach(r=>{
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${r.DateISO||''}</td><td>${r.Brand||''}</td><td>${r.Shift||''}</td><td>${r.Line||''}</td><td>${r.Qty||0}</td><td>${r.Rework||0}</td><td>${r.ReworkFixed||0}</td>`;
    tbody.appendChild(tr);
  });
}

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

// UI helpers
function applyFilters(){ renderAll(); }
function refreshAll(){ loadSheetData().then(()=>{ initUI(); renderAll(); }); }
function toggleTheme(){ document.body.classList.toggle('dark'); }
