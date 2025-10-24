// CONFIG: change these if needed
// This script fetches data from Google Sheets GViz JSON and renders charts/tables.
// NOTE: SHEET_GVIZ_URL now uses getSheetUrl(gid) to select the correct sheet (GID).
const SPREADSHEET_ID = '1XoV7020NTZk1kzqn3F2ks3gOVFJ5arr5NVgUdewWPNQ';
const GID_IN = "1100244896";
const GID_EXPSCHED = "359974075"; // isi nanti, aku tunjukkan caranya di bawah
const GID_ERP = "1158274905";

function getSheetUrl(gid){
  return `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=out:json&gid=${gid}`;
}

// --- Sheet field mapping (acuan / 參考)
// Sheet "IN" (gid = GID_IN)
// 訂單號碼 / PO. NO. -> col A (mulai baris 6)
// 工單號碼 / NO. WO. -> col B (mulai baris 6)
// 產品料號 / PART. NO. -> col C (mulai baris 6)
// 客戶 / Customer -> col D (mulai baris 6)
// 產品型號 / ITEM Type -> col E (mulai baris 6)
// 尺寸 / Size -> col F (mulai baris 6)
// SML -> col G (mulai baris 6)
// 顏色 / Color -> col H (mulai baris 6)
// 訂單量 / PO QTY -> col I (mulai baris 6)
// 入庫數 / IN -> col J (mulai baris 6)
// 欠數 / ReMaining -> col K (mulai baris 6)
// Used for Shipment -> col L (mulai baris 6)
// Ready For Shipment -> col M (mulai baris 6)
// Rework QTY -> col N (mulai baris 6)
// Mulai kolom Q hingga seterusnya:
// baris 5: judul invoice
// baris 4: QTY Container
// baris 3: QTY Total per invoice
// baris 2: tanggal eksport
// baris 1: brand name
// baris 6,7,8,...: qty setiap invoice sesuai detailnya

// Sheet "ExpSched" (gid = GID_EXPSCHED)
// NO -> col A
// Customer -> col B
// Destination -> col D
// Invoice No. -> col F
// QTY Carton -> col I
// QTY Pcs -> col J
// Stuffing -> col R
// QTYCONTAINER -> col S

// Sheet "ERP" (gid = GID_ERP)
// 工單號碼 / NO. WO. -> col C
// 訂單號碼 / PO. NO. -> col D
// 產品料號 / PART. NO. -> col E
// Finish Goods QTY -> col J

// global chart refs
let prodChart, shiftChart, expChart;
let rawData = []; // array of objects from sheet
const palette = ['#7c3aed','#00ffe1','#10b981','#ff7ab6','#f59e0b','#60a5fa'];

document.addEventListener('DOMContentLoaded', () => {
  // Default to reading the IN sheet; change to GID_EXPSCHED or GID_ERP if needed.
  loadSheetData(GID_IN).then(() => {
    initUI();
    renderAll();
  }).catch(err=>{
    console.error('Failed to load sheet:', err);
    alert('Gagal memuat sheet. Pastikan spreadsheet sudah di-publish (File > Publish to web) dan GID benar.');
  });

  const gs = document.getElementById('globalSearch');
  if(gs) gs.addEventListener('keyup', (e)=>{
    if(e.key === 'Enter') renderAll();
  });
});

async function loadSheetData(gid){
  const SHEET_GVIZ_URL = getSheetUrl(gid);
  const res = await fetch(SHEET_GVIZ_URL);
  const text = await res.text();
  // GViz returns: /*O_o*/
google.visualization.Query.setResponse({...});
  const jsonText = text.replace(/^[^\(]*\(/,'').replace(/\);?\s*$/,'');
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
      // attempt Excel serial -> JS Date conversion
      const date = new Date(Math.round((d - 25569) * 86400 * 1000));
      if(!isNaN(date)) r.DateISO = date.toISOString().slice(0,10);
      else r.DateISO = String(d);
    } else if (typeof d === 'string') {
      const parsed = new Date(d);
      if(!isNaN(parsed)) r.DateISO = parsed.toISOString().slice(0,10);
      else r.DateISO = d.slice(0,10);
    } else {
      r.DateISO = '';
    }
  });
}

function initUI(){
  // fill brand select if present
  const brands = Array.from(new Set(rawData.map(r=>String(r.Brand || '').trim()).filter(Boolean))).sort();
  const select = document.getElementById('brandSelect');
  if(!select) return;
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
  const select = document.getElementById('brandSelect');
  const brands = select ? Array.from(select.selectedOptions).map(o=>o.value) : [];
  const qEl = document.getElementById('globalSearch');
  const q = qEl ? qEl.value.trim().toLowerCase() : '';
  const rangeEl = document.getElementById('timeRange');
  const range = rangeEl ? rangeEl.value : '';
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
  const kpiProd = document.getElementById('kpiProd');
  const kpiExp = document.getElementById('kpiExp');
  const kpiRework = document.getElementById('kpiRework');
  const kpiBrand = document.getElementById('kpiBrand');
  if(kpiProd) kpiProd.textContent = todaySum;
  if(kpiExp) kpiExp.textContent = monthSum;
  if(kpiRework) kpiRework.textContent = reworkRate.toFixed(2) + '%';
  if(kpiBrand) kpiBrand.textContent = topBrand ? topBrand[0] : '-';
}

function drawProdChart(data){
  const range = (document.getElementById('timeRange') || {}).value;
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
  // group total qty by brand
  const brandMap = {};
  data.forEach(d=>{
    const b = d.Brand || 'Unknown';
    brandMap[b] = (brandMap[b]||0) + (d.Qty||0);
  });
  const labels = Object.keys(brandMap);
  const values = labels.map(l=>brandMap[l]);
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
function refreshAll(){ loadSheetData(GID_IN).then(()=>{ initUI(); renderAll(); }); }
function toggleTheme(){ document.body.classList.toggle('dark'); }
