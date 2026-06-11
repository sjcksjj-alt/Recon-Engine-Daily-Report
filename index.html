<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>Recon Engine Daily Report</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
*{margin:0;padding:0;box-sizing:border-box;}
body{background:#f0f2f5;color:#1a1a2e;font-family:'Segoe UI',sans-serif;font-size:12px;padding:12px;}
#sync_banner{position:fixed;top:0;left:0;right:0;z-index:9999;padding:7px 14px;font-size:11px;font-weight:700;text-align:center;display:none;}
#sync_banner.saving{display:block;background:#fef3c7;color:#92400e;}
#sync_banner.saved{display:block;background:#d1fae5;color:#065f46;}
#sync_banner.error{display:block;background:#fee2e2;color:#991b1b;}

/* ── Setup overlay ── */
.setup-overlay{position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(15,23,42,0.88);z-index:10000;display:flex;align-items:center;justify-content:center;}
.setup-box{background:#fff;border-radius:14px;padding:28px 30px;max-width:480px;width:96%;box-shadow:0 24px 60px rgba(0,0,0,0.3);}
.setup-box h2{font-size:16px;font-weight:800;color:#1a1a2e;margin-bottom:6px;}
.setup-box p{font-size:11px;color:#64748b;line-height:1.7;margin-bottom:14px;}
.setup-steps{background:#f8fafc;border-radius:8px;padding:12px 14px;margin-bottom:16px;font-size:10px;color:#374151;line-height:2;}
.setup-steps b{color:#2563eb;}
.setup-steps a{color:#2563eb;}
.setup-box label{font-size:9px;font-weight:700;color:#374151;display:block;margin-bottom:3px;text-transform:uppercase;letter-spacing:.4px;}
.setup-box input{width:100%;border:1.5px solid #e2e8f0;border-radius:7px;padding:9px 12px;font-size:12px;color:#1a1a2e;outline:none;margin-bottom:10px;font-family:monospace;}
.setup-box input:focus{border-color:#2563eb;}
.btn-connect{width:100%;padding:11px;background:#2563eb;border:none;border-radius:8px;color:#fff;font-size:13px;font-weight:700;cursor:pointer;margin-top:2px;}
.btn-connect:hover{background:#1d4ed8;}
.setup-note{font-size:9px;color:#94a3b8;margin-top:10px;line-height:1.6;text-align:center;}

.header{display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;background:#fff;border-radius:10px;padding:10px 16px;border:1px solid #e2e8f0;box-shadow:0 1px 4px rgba(0,0,0,0.06);}
.header-left h1{font-size:15px;font-weight:700;color:#1a1a2e;}
.header-left .subtitle{font-size:10px;color:#94a3b8;margin-top:1px;}
.header-right{display:flex;align-items:center;gap:10px;}
#db_status{font-size:9px;color:#94a3b8;font-weight:600;}
.tab-btns{display:flex;gap:6px;}
.tab-btn{padding:5px 18px;border-radius:6px;border:1.5px solid #cbd5e1;background:#f8fafc;color:#64748b;cursor:pointer;font-size:11px;font-weight:600;transition:all 0.15s;}
.tab-btn.h-active{background:#1e3a5f;border-color:#2563eb;color:#fff;}
.tab-btn.p-active{background:#064e3b;border-color:#10b981;color:#fff;}
.top-row{display:grid;grid-template-columns:320px 1fr 1fr;gap:10px;margin-bottom:10px;}
.cal-card{background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:12px;box-shadow:0 1px 4px rgba(0,0,0,0.05);}
.cal-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:6px;}
.cal-month{font-size:13px;font-weight:700;color:#1a1a2e;}
.cal-nav{display:flex;gap:4px;align-items:center;}
.cal-btn{background:#f1f5f9;border:1px solid #e2e8f0;color:#64748b;border-radius:5px;padding:2px 8px;cursor:pointer;font-size:12px;line-height:1.4;}
.cal-btn:hover{background:#e2e8f0;}
.cal-grid{display:grid;grid-template-columns:repeat(7,1fr);gap:2px;}
.cal-day-name{text-align:center;font-size:9px;color:#94a3b8;padding:2px 0;font-weight:600;}
.cal-day{text-align:center;padding:3px 0;border-radius:5px;font-size:11px;font-weight:600;cursor:pointer;border:1.5px solid transparent;min-height:22px;display:flex;align-items:center;justify-content:center;transition:all 0.1s;color:#374151;}
.cal-day.empty{cursor:default;}
.cal-day.today{border-color:#2563eb;color:#2563eb;background:#eff6ff;}
.cal-day.has-data{background:#d1fae5;border-color:#10b981;color:#059669;}
.cal-day.selected{background:#dbeafe;border-color:#2563eb;color:#1d4ed8;}
.cal-day:not(.empty):hover{background:#f1f5f9;border-color:#cbd5e1;}
.entry-card{background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:12px;box-shadow:0 1px 4px rgba(0,0,0,0.05);}
.card-title{font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px;}
.selected-date-label{font-size:12px;font-weight:700;color:#1a1a2e;margin-bottom:8px;}
.field-group{margin-bottom:7px;}
.field-group label{display:block;font-size:9px;color:#94a3b8;margin-bottom:3px;font-weight:600;text-transform:uppercase;letter-spacing:0.3px;}
.field-group input{width:100%;background:#f8fafc;border:1.5px solid #e2e8f0;border-radius:6px;padding:7px 10px;color:#1a1a2e;font-size:13px;font-weight:700;outline:none;transition:border-color 0.15s;}
.field-group input:focus{border-color:#2563eb;background:#fff;}
.btn-save{width:100%;padding:8px;background:#2563eb;border:none;border-radius:7px;color:#fff;font-size:12px;font-weight:700;cursor:pointer;transition:background 0.15s;margin-top:4px;}
.btn-save:hover{background:#1d4ed8;}
.btn-save:disabled{background:#93c5fd;cursor:not-allowed;}
.changelog{background:#f8fafc;border:1px solid #e2e8f0;border-radius:5px;padding:6px;max-height:65px;overflow-y:auto;margin-top:6px;}
.cl-title{color:#f59e0b;font-size:8px;font-weight:700;margin-bottom:3px;}
.cl-time{color:#3b82f6;font-size:8px;}
.cl-change{font-size:9px;color:#374151;}
.summary-card{background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:12px;box-shadow:0 1px 4px rgba(0,0,0,0.05);}
.summary-title-row{display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;}
.summary-title{font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:0.5px;}
.badge{background:#2563eb;border-radius:4px;padding:2px 8px;font-size:9px;color:#fff;font-weight:700;}
.s-row{display:flex;justify-content:space-between;align-items:center;padding:4px 0;border-bottom:1px solid #f1f5f9;}
.s-row:last-child{border-bottom:none;}
.s-label{font-size:9px;color:#64748b;display:flex;align-items:center;gap:5px;}
.s-dot{width:7px;height:7px;border-radius:2px;flex-shrink:0;}
.s-val{font-size:12px;font-weight:700;font-family:'Courier New',monospace;}
.summary-bar{background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:10px 14px;margin-bottom:10px;box-shadow:0 1px 4px rgba(0,0,0,0.05);}
.summary-bar-title{font-size:10px;font-weight:700;color:#64748b;margin-bottom:7px;text-transform:uppercase;letter-spacing:0.5px;}
.legend-flex{display:flex;gap:16px;flex-wrap:wrap;align-items:center;}
.legend-item{display:flex;align-items:center;gap:5px;}
.leg-dot{width:10px;height:10px;border-radius:3px;flex-shrink:0;}
.leg-label{font-size:9px;color:#64748b;}
.leg-val{font-size:11px;font-weight:700;font-family:'Courier New',monospace;margin-left:2px;}
.charts-row{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;margin-bottom:10px;}
.chart-card{background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:12px;box-shadow:0 1px 4px rgba(0,0,0,0.05);}
.chart-title{font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px;}
.bottom-charts-row{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;margin-bottom:10px;}
.bottom-chart-card{background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:12px;box-shadow:0 1px 4px rgba(0,0,0,0.05);}
.bottom-chart-title{font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px;}
.dual-chart-labels{display:flex;justify-content:space-between;margin-bottom:4px;}
.dual-label{font-size:9px;font-weight:700;padding:1px 7px;border-radius:4px;}
.dual-label.h{background:#dbeafe;color:#1d4ed8;}
.dual-label.p{background:#d1fae5;color:#059669;}
.dual-chart-wrap{display:grid;grid-template-columns:1fr 1fr;gap:4px;}
.dual-chart-wrap canvas{width:100%!important;}
.monthly-card{background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:12px;margin-bottom:10px;box-shadow:0 1px 4px rgba(0,0,0,0.05);}
.monthly-title{font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:7px;}
.monthly-table{width:100%;border-collapse:collapse;font-size:10px;}
.monthly-table th{text-align:left;padding:4px 8px;font-size:9px;color:#94a3b8;border-bottom:1.5px solid #e2e8f0;font-weight:600;}
.monthly-table td{padding:4px 8px;border-bottom:1px solid #f1f5f9;font-family:'Courier New',monospace;}
.monthly-table tr:last-child td{border-bottom:none;}
.monthly-table tr:hover td{background:#f8fafc;}
.actions-row{display:flex;gap:10px;margin-bottom:10px;justify-content:center;flex-wrap:wrap;}
.btn-action{padding:10px 24px;border-radius:8px;font-size:12px;font-weight:700;cursor:pointer;border:none;transition:all 0.15s;}
.btn-view{background:#1e3a5f;color:#fff;}
.btn-view:hover{background:#2563eb;}
.btn-excel{background:#064e3b;color:#fff;}
.btn-excel:hover{background:#10b981;}
.btn-settings{background:#f1f5f9;color:#374151;border:1.5px solid #e2e8f0;}
.btn-settings:hover{background:#e2e8f0;}
.share-card{background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:12px;box-shadow:0 1px 4px rgba(0,0,0,0.05);}
.share-title{font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:7px;}
.share-preview{background:#f8fafc;border:1px solid #e2e8f0;border-radius:6px;padding:10px;font-family:'Courier New',monospace;font-size:10px;color:#374151;line-height:1.7;margin-bottom:8px;white-space:pre-wrap;max-height:160px;overflow-y:auto;}
.share-btns{display:grid;grid-template-columns:1fr 1fr;gap:8px;}
.btn-wa{padding:9px;background:#064e3b;border:1.5px solid #10b981;border-radius:7px;color:#10b981;font-size:12px;font-weight:700;cursor:pointer;text-align:center;transition:all 0.15s;}
.btn-wa:hover{background:#10b981;color:#fff;}
.btn-email{padding:9px;background:#dbeafe;border:1.5px solid #2563eb;border-radius:7px;color:#1d4ed8;font-size:12px;font-weight:700;cursor:pointer;text-align:center;transition:all 0.15s;}
.btn-email:hover{background:#2563eb;color:#fff;}
.modal-overlay{display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.5);z-index:1000;align-items:center;justify-content:center;}
.modal-overlay.show{display:flex;}
.modal{background:#fff;border-radius:12px;padding:20px;max-width:700px;width:95%;max-height:80vh;overflow-y:auto;box-shadow:0 20px 60px rgba(0,0,0,0.2);}
.modal-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:14px;padding-bottom:10px;border-bottom:1px solid #e2e8f0;}
.modal-title{font-size:14px;font-weight:700;color:#1a1a2e;}
.modal-close{background:#f1f5f9;border:none;border-radius:6px;padding:4px 10px;cursor:pointer;font-size:14px;color:#64748b;}
.modal-close:hover{background:#e2e8f0;}
.modal-table{width:100%;border-collapse:collapse;font-size:11px;}
.modal-table th{text-align:left;padding:6px 8px;font-size:9px;color:#94a3b8;border-bottom:1.5px solid #e2e8f0;font-weight:600;background:#f8fafc;}
.modal-table td{padding:5px 8px;border-bottom:1px solid #f1f5f9;font-family:'Courier New',monospace;}
.modal-table tr:last-child td{border-bottom:none;}
.modal-table tr:hover td{background:#f8fafc;}
.modal-section{font-size:10px;font-weight:700;color:#2563eb;background:#eff6ff;padding:5px 8px;border-radius:5px;margin:10px 0 5px;}
canvas{max-height:90px;}
</style>
</head>
<body>

<div id="sync_banner"></div>

<!-- Setup overlay -->
<div class="setup-overlay" id="setup_overlay" style="display:none;">
  <div class="setup-box">
    <h2>🔗 JSONBin Setup — One Time Only</h2>
    <p>உங்கள் data எல்லாரும் பார்க்க, ஒரு free JSONBin account வேணும். 2 நிமிஷம் மட்டும்!</p>
    <div class="setup-steps">
      <b>Step 1:</b> <a href="https://jsonbin.io" target="_blank">jsonbin.io</a> → Sign up (free, email போதும்)<br>
      <b>Step 2:</b> Login → Top right → <b>API Keys</b> → <b>Secret Key</b> copy பண்ணு<br>
      <b>Step 3:</b> அந்த key இங்க paste பண்ணு → Connect பண்ணு<br>
      <b>Step 4:</b> Done! எல்லாரும் same data பார்ப்பாங்க 🎉
    </div>
    <label>JSONBin Secret Key (starts with $2b$...)</label>
    <input id="jb_key" placeholder="$2b$10$..." autocomplete="off"/>
    <button class="btn-connect" onclick="connectJSONBin()">🔗 Connect & Start</button>
    <div class="setup-note">உங்கள் key இந்த browser's localStorage ல மட்டும் save ஆகும். JSONBin free plan: 10,000 requests/month — போதும்.</div>
  </div>
</div>

<div class="header">
  <div class="header-left">
    <h1>Recon Engine Daily Report</h1>
    <div class="subtitle" id="report_subtitle"></div>
  </div>
  <div class="header-right">
    <div id="db_status">⏳ Loading…</div>
    <div class="tab-btns">
      <div class="tab-btn h-active" id="tab_H" onclick="switchTab('H')">H Recon</div>
      <div class="tab-btn" id="tab_P15" onclick="switchTab('P15')">P15 Recon</div>
    </div>
  </div>
</div>

<div class="top-row">
  <div class="cal-card">
    <div class="cal-header">
      <div class="cal-month" id="cal_title"></div>
      <div class="cal-nav">
        <button class="cal-btn" onclick="changeMonth(-1)">‹</button>
        <button class="cal-btn" onclick="goToday()">•</button>
        <button class="cal-btn" onclick="changeMonth(1)">›</button>
      </div>
    </div>
    <div class="cal-grid" id="cal_grid"></div>
  </div>
  <div class="entry-card">
    <div class="card-title">Select a Date</div>
    <div class="selected-date-label" id="entry_title">—</div>
    <div class="field-group"><label>Core Receipt</label><input type="number" id="inp_core" placeholder="0" min="0"/></div>
    <div class="field-group"><label>Invoice Received</label><input type="number" id="inp_inv_r" placeholder="0" min="0"/></div>
    <div class="field-group"><label>Invoice Dispatched</label><input type="number" id="inp_inv_d" placeholder="0" min="0"/></div>
    <button class="btn-save" id="btn_save" onclick="saveDay()">💾 Save</button>
    <div id="modified_info"></div>
  </div>
  <div class="summary-card">
    <div class="summary-title-row">
      <div class="summary-title">Day Summary</div>
      <span class="badge" id="summary_date_badge">—</span>
    </div>
    <div id="day_summary_box"><div style="color:#94a3b8;font-size:10px;text-align:center;padding:20px;">Click a date to view</div></div>
  </div>
</div>

<div class="summary-bar">
  <div class="summary-bar-title">📊 Summary — <span id="legend_day_label">—</span></div>
  <div class="legend-flex" id="legend_box"></div>
</div>

<div class="charts-row">
  <div class="chart-card"><div class="chart-title">Core Receipt Trend</div><canvas id="bar_chart"></canvas></div>
  <div class="chart-card"><div class="chart-title">Invoice Dispatched Trend</div><canvas id="line_chart"></canvas></div>
  <div class="chart-card"><div class="chart-title">Invoice Received Trend</div><canvas id="invr_chart"></canvas></div>
</div>

<div class="bottom-charts-row">
  <div class="bottom-chart-card">
    <div class="bottom-chart-title">📈 Inv &amp; Dispatch Cumulative</div>
    <div class="dual-chart-labels"><span class="dual-label h">H Recon</span><span class="dual-label p">P15 Recon</span></div>
    <div class="dual-chart-wrap"><canvas id="h_cum_chart"></canvas><canvas id="p_cum_chart"></canvas></div>
  </div>
  <div class="bottom-chart-card">
    <div class="bottom-chart-title">📊 Inv &amp; Dispatch Average</div>
    <div class="dual-chart-labels"><span class="dual-label h">H Recon</span><span class="dual-label p">P15 Recon</span></div>
    <div class="dual-chart-wrap"><canvas id="h_avg_chart"></canvas><canvas id="p_avg_chart"></canvas></div>
  </div>
  <div class="bottom-chart-card">
    <div class="bottom-chart-title">💜 Core Average Receipt</div>
    <div class="dual-chart-labels"><span class="dual-label h">H Recon</span><span class="dual-label p">P15 Recon</span></div>
    <div class="dual-chart-wrap"><canvas id="h_coreavg_chart"></canvas><canvas id="p_coreavg_chart"></canvas></div>
  </div>
</div>

<div class="monthly-card">
  <div class="monthly-title">📅 Monthly Summary</div>
  <div style="overflow-x:auto"><table class="monthly-table" id="monthly_table"></table></div>
</div>

<div class="actions-row">
  <button class="btn-action btn-view" onclick="openAllData()">📄 View All Data</button>
  <button class="btn-action btn-excel" onclick="exportExcel()">⬇ Download Excel (.xlsx)</button>
  <button class="btn-action btn-settings" onclick="resetSetup()">⚙ Change API Key</button>
</div>

<div class="share-card">
  <div class="share-title">📤 Share Day Report</div>
  <div class="share-preview" id="share_preview"></div>
  <div class="share-btns">
    <div class="btn-wa" onclick="shareWA()">📱 WhatsApp</div>
    <div class="btn-email" onclick="shareEmail()">📧 Email</div>
  </div>
</div>

<div class="modal-overlay" id="modal_overlay" onclick="closeModalOutside(event)">
  <div class="modal">
    <div class="modal-header">
      <div class="modal-title">📄 All Data</div>
      <button class="modal-close" onclick="closeModal()">✕ Close</button>
    </div>
    <div id="modal_body"></div>
  </div>
</div>

<script>
/* ══════════════════════════════════════════
   CONSTANTS & STATE
══════════════════════════════════════════ */
const MONTHS=['January','February','March','April','May','June','July','August','September','October','November','December'];
const MONTHS_SHORT=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const JSONBIN_API='https://api.jsonbin.io/v3/b';

let activeTab='H', today=new Date(),
    viewYear=today.getFullYear(), viewMonth=today.getMonth(),
    selectedDate='';

// In-memory store — loaded from JSONBin
const store={ H:{}, P15:{} };
let BIN_ID=null, API_KEY=null;

/* ══════════════════════════════════════════
   JSONBIN HELPERS
══════════════════════════════════════════ */
function getConfig(){
  try{ return JSON.parse(localStorage.getItem('recon_jb_cfg')||'null'); }
  catch(e){ return null; }
}

async function connectJSONBin(){
  const key=document.getElementById('jb_key').value.trim();
  if(!key){ showBanner('error','⚠ Key பண்ணுங்க!'); return; }
  API_KEY=key;
  showBanner('saving','⏳ Testing connection & creating bin…',30000);
  try{
    // Create a new bin with empty data
    const res=await fetch(JSONBIN_API,{
      method:'POST',
      headers:{'Content-Type':'application/json','X-Master-Key':key,'X-Bin-Name':'recon-engine-data','X-Collection-Name':'ReconEngine'},
      body:JSON.stringify({H:{},P15:{}})
    });
    if(!res.ok) throw new Error('Invalid key or network error ('+res.status+')');
    const data=await res.json();
    BIN_ID=data.metadata.id;
    localStorage.setItem('recon_jb_cfg',JSON.stringify({key,binId:BIN_ID}));
    document.getElementById('setup_overlay').style.display='none';
    showBanner('saved','✅ Connected! Bin ID: '+BIN_ID);
    setStatus('✅ Cloud synced','#10b981');
    renderAll();
  }catch(err){
    showBanner('error','❌ '+err.message);
  }
}

async function loadFromCloud(){
  if(!BIN_ID||!API_KEY) return;
  try{
    const res=await fetch(`${JSONBIN_API}/${BIN_ID}/latest`,{
      headers:{'X-Master-Key':API_KEY}
    });
    if(!res.ok) throw new Error('Load failed: '+res.status);
    const data=await res.json();
    const record=data.record||{};
    store.H  =record.H  ||{};
    store.P15=record.P15||{};
    setStatus('✅ Cloud synced','#10b981');
  }catch(err){
    setStatus('⚠ Load error','#f59e0b');
    showBanner('error','Load error: '+err.message,4000);
    // Fallback to localStorage
    store.H  =JSON.parse(localStorage.getItem('recon_bk_H' )||'{}');
    store.P15=JSON.parse(localStorage.getItem('recon_bk_P15')||'{}');
  }
}

async function saveToCloud(){
  if(!BIN_ID||!API_KEY) return false;
  const res=await fetch(`${JSONBIN_API}/${BIN_ID}`,{
    method:'PUT',
    headers:{'Content-Type':'application/json','X-Master-Key':API_KEY},
    body:JSON.stringify({H:store.H,P15:store.P15})
  });
  if(!res.ok) throw new Error('Save failed: '+res.status);
  return true;
}

function resetSetup(){
  if(!confirm('API Key மாத்தணுமா? இந்த browser la setup reset ஆகும்.')) return;
  localStorage.removeItem('recon_jb_cfg');
  location.reload();
}

/* ══════════════════════════════════════════
   BANNER & STATUS
══════════════════════════════════════════ */
let bannerTimer=null;
function showBanner(type,msg,ms=2800){
  const b=document.getElementById('sync_banner');
  b.className=type; b.textContent=msg; b.style.display='block';
  clearTimeout(bannerTimer);
  bannerTimer=setTimeout(()=>{b.style.display='none';},ms);
}
function setStatus(txt,color){
  const el=document.getElementById('db_status');
  el.textContent=txt; el.style.color=color;
}

/* ══════════════════════════════════════════
   SAVE A DAY
══════════════════════════════════════════ */
async function saveDay(){
  if(!selectedDate) return;
  const dateKey=selectedDate;
  const core =parseInt(document.getElementById('inp_core').value)||0;
  const inv_r=parseInt(document.getElementById('inp_inv_r').value)||0;
  const inv_d=parseInt(document.getElementById('inp_inv_d').value)||0;
  const old  =store[activeTab][dateKey]||{core:null,inv_r:null,inv_d:null};
  const ts   =new Date().toISOString();
  const changes=[];
  if(old.core===null) changes.push(`New: Core=${core}, Rec=${inv_r}, Disp=${inv_d}`);
  else{
    if(old.core!==core)   changes.push(`Core: ${old.core}→${core}`);
    if(old.inv_r!==inv_r) changes.push(`Rec: ${old.inv_r}→${inv_r}`);
    if(old.inv_d!==inv_d) changes.push(`Disp: ${old.inv_d}→${inv_d}`);
  }
  const log=[...(old.log||[])];
  if(changes.length>0) log.push({ts,changes});
  store[activeTab][dateKey]={core,inv_r,inv_d,ts,log};

  // Local backup always
  localStorage.setItem('recon_bk_H',  JSON.stringify(store.H));
  localStorage.setItem('recon_bk_P15',JSON.stringify(store.P15));

  selectedDate='';
  renderAll();

  if(BIN_ID&&API_KEY){
    showBanner('saving','⏳ Saving to cloud…',60000);
    document.getElementById('btn_save').disabled=true;
    try{
      await saveToCloud();
      showBanner('saved','✅ Saved! எல்லாரும் இப்போ பார்க்கலாம்.');
      setStatus('✅ Cloud synced','#10b981');
    }catch(err){
      showBanner('error','⚠ Cloud save failed. Local backup ok. '+err.message,5000);
      setStatus('⚠ Sync error','#ef4444');
    }finally{
      document.getElementById('btn_save').disabled=false;
    }
  }else{
    showBanner('saving','📦 Saved locally (Setup பண்ணா cloud sync ஆகும்)',3000);
  }
}

/* ══════════════════════════════════════════
   UTILITIES
══════════════════════════════════════════ */
function formatDate(d){return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;}
function displayDate(s){if(!s)return '—';const p=s.split('-');return `${p[2]}/${p[1]}/${p[0]}`;}

function calcUpTo(tab,date){
  if(!date)return{cumCore:0,cumR:0,cumD:0,avgR:0,avgD:0,avgCore:0};
  const s=store[tab],keys=Object.keys(s).filter(k=>k<=date).sort();
  let cc=0,cr=0,cd=0;
  keys.forEach(k=>{cc+=s[k].core;cr+=s[k].inv_r;cd+=s[k].inv_d;});
  const n=keys.length||1;
  return{cumCore:cc,cumR:cr,cumD:cd,avgR:+(cr/n).toFixed(1),avgD:+(cd/n).toFixed(1),avgCore:+(cc/n).toFixed(1)};
}

function buildSeries(tab){
  const s=store[tab],keys=Object.keys(s).sort();
  const labels=[],core=[],invR=[],invDisp=[],invCum=[],invAvg=[],coreAvg=[];
  let cc=0,cr=0,cd=0;
  keys.forEach((k,i)=>{
    const p=k.split('-');labels.push(`${p[2]}/${p[1]}`);
    cc+=s[k].core;cr+=s[k].inv_r;cd+=s[k].inv_d;const n=i+1;
    core.push(s[k].core);invR.push(s[k].inv_r);invDisp.push(s[k].inv_d);
    invCum.push(cr+cd);invAvg.push(+((cr+cd)/(2*n)).toFixed(1));coreAvg.push(+(cc/n).toFixed(1));
  });
  return{labels,core,invR,invDisp,invCum,invAvg,coreAvg};
}

const FIELDS=[
  {key:'core',   label:'Core receipt',    color:'#2563eb'},
  {key:'invDisp',label:'Inv dispatched',  color:'#10b981'},
  {key:'coreCum',label:'Core cumulative', color:'#ef4444'},
  {key:'invCum', label:'Inv cumulative',  color:'#f97316'},
  {key:'invAvg', label:'Inv & disp avg',  color:'#eab308'},
  {key:'coreAvg',label:'Core avg',        color:'#a855f7'},
];

/* ══════════════════════════════════════════
   RENDER
══════════════════════════════════════════ */
function switchTab(t){
  activeTab=t; selectedDate='';
  document.getElementById('tab_H').className='tab-btn'+(t==='H'?' h-active':'');
  document.getElementById('tab_P15').className='tab-btn'+(t==='P15'?' p-active':'');
  renderAll();
}
function changeMonth(d){viewMonth+=d;if(viewMonth>11){viewMonth=0;viewYear++;}if(viewMonth<0){viewMonth=11;viewYear--;}renderAll();}
function goToday(){viewYear=today.getFullYear();viewMonth=today.getMonth();renderAll();}

function renderCalendar(){
  const s=store[activeTab];
  document.getElementById('cal_title').textContent=`${MONTHS[viewMonth]} ${viewYear}`;
  document.getElementById('report_subtitle').textContent=`${MONTHS[viewMonth]} ${viewYear} — Recon Engine`;
  const grid=document.getElementById('cal_grid'); grid.innerHTML='';
  ['Su','Mo','Tu','We','Th','Fr','Sa'].forEach(d=>{
    const el=document.createElement('div');el.className='cal-day-name';el.textContent=d;grid.appendChild(el);
  });
  const first=new Date(viewYear,viewMonth,1).getDay();
  const days=new Date(viewYear,viewMonth+1,0).getDate();
  for(let i=0;i<first;i++){const el=document.createElement('div');el.className='cal-day empty';grid.appendChild(el);}
  for(let d=1;d<=days;d++){
    const ds=`${viewYear}-${String(viewMonth+1).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
    const el=document.createElement('div');el.className='cal-day';el.textContent=d;
    if(s[ds])el.classList.add('has-data');
    if(ds===formatDate(today)&&ds!==selectedDate)el.classList.add('today');
    if(ds===selectedDate)el.classList.add('selected');
    el.onclick=()=>{selectedDate=(selectedDate===ds)?'':ds;renderAll();};
    grid.appendChild(el);
  }
}

function renderDaySummary(){
  const dateKey=selectedDate||formatDate(today);
  document.getElementById('summary_date_badge').textContent=displayDate(dateKey);
  const box=document.getElementById('day_summary_box');
  const ex=store[activeTab][dateKey];
  if(!ex){box.innerHTML=`<div style="color:#94a3b8;font-size:10px;text-align:center;padding:20px;">📅 ${displayDate(dateKey)}<br>No data yet</div>`;return;}
  const c=calcUpTo(activeTab,dateKey);
  const rows=[
    {label:'Core Receipt',  val:ex.core,               color:'#2563eb'},
    {label:'Inv Rec / Disp',val:`${ex.inv_r} / ${ex.inv_d}`,color:'#10b981'},
    {label:'Core Cumulative',val:c.cumCore,             color:'#ef4444'},
    {label:'Inv Cumulative', val:`${c.cumR} / ${c.cumD}`,color:'#f97316'},
    {label:'Inv & Disp Avg', val:`${c.avgR} / ${c.avgD}`,color:'#eab308'},
    {label:'Core Avg',       val:c.avgCore,             color:'#a855f7'},
  ];
  box.innerHTML=rows.map(r=>`<div class="s-row"><div class="s-label"><div class="s-dot" style="background:${r.color}"></div>${r.label}</div><div class="s-val" style="color:${r.color}">${r.val}</div></div>`).join('');
}

function renderEntryForm(){
  const ex=selectedDate?store[activeTab][selectedDate]:null;
  document.getElementById('entry_title').textContent=selectedDate?`📅 ${displayDate(selectedDate)}`:'—';
  document.getElementById('inp_core').value=ex?ex.core:'';
  document.getElementById('inp_inv_r').value=ex?ex.inv_r:'';
  document.getElementById('inp_inv_d').value=ex?ex.inv_d:'';
  if(ex&&ex.log&&ex.log.length>0){
    let html=`<div class="changelog"><div class="cl-title">📋 CHANGE LOG</div>`;
    [...ex.log].reverse().forEach(l=>{
      const ld=new Date(l.ts);
      html+=`<div style="margin-bottom:4px"><div class="cl-time">🕐 ${ld.toLocaleDateString('en-GB')} ${ld.toLocaleTimeString('en-GB',{hour:'2-digit',minute:'2-digit'})}</div>`;
      l.changes.forEach(ch=>{html+=`<div class="cl-change" style="color:${ch.startsWith('New')?'#10b981':'#374151'}">• ${ch}</div>`;});
      html+=`</div>`;
    });
    html+=`</div>`;
    document.getElementById('modified_info').innerHTML=html;
  }else document.getElementById('modified_info').innerHTML='';
}

function renderLegend(){
  const date=selectedDate||formatDate(today);
  const ex=store[activeTab][date]||{core:0,inv_r:0,inv_d:0};
  const c=calcUpTo(activeTab,date);
  const values={core:ex.core,invDisp:ex.inv_d,coreCum:c.cumCore,invCum:c.cumR,invAvg:c.avgR,coreAvg:c.avgCore};
  document.getElementById('legend_day_label').textContent=displayDate(date);
  document.getElementById('legend_box').innerHTML=FIELDS.map(f=>`
    <div class="legend-item">
      <div class="leg-dot" style="background:${f.color}"></div>
      <span class="leg-label">${f.label}</span>
      <span class="leg-val" style="color:${f.color}">${values[f.key]}</span>
    </div>`).join('');
}

const charts={};
const CHART_OPTS={responsive:true,plugins:{legend:{display:false}},scales:{x:{ticks:{color:'#94a3b8',font:{size:7}},grid:{color:'#f1f5f9'}},y:{ticks:{color:'#94a3b8',font:{size:7}},grid:{color:'#f1f5f9'},beginAtZero:true}}};
function makeChart(id,labels,data,color,type='line'){
  if(charts[id])charts[id].destroy();
  charts[id]=new Chart(document.getElementById(id),{type,data:{labels,datasets:[{data,borderColor:color,backgroundColor:color+'33',pointRadius:2,pointBackgroundColor:color,tension:0.4,fill:true,borderWidth:1.5,borderRadius:type==='bar'?3:0}]},options:{...CHART_OPTS,maintainAspectRatio:false}});
}

function renderTrends(){
  const s=buildSeries(activeTab),sH=buildSeries('H'),sP=buildSeries('P15');
  makeChart('bar_chart',s.labels,s.core,'#2563eb','line');
  makeChart('line_chart',s.labels,s.invDisp,'#10b981');
  makeChart('invr_chart',s.labels,s.invR,'#f97316');
  makeChart('h_cum_chart',sH.labels,sH.invCum,'#f97316');
  makeChart('h_avg_chart',sH.labels,sH.invAvg,'#eab308');
  makeChart('h_coreavg_chart',sH.labels,sH.coreAvg,'#a855f7');
  makeChart('p_cum_chart',sP.labels,sP.invCum,'#10b981');
  makeChart('p_avg_chart',sP.labels,sP.invAvg,'#eab308');
  makeChart('p_coreavg_chart',sP.labels,sP.coreAvg,'#a855f7');
}

function renderMonthly(){
  const s=store[activeTab],monthly={};
  Object.keys(s).sort().forEach(k=>{const ym=k.substring(0,7);if(!monthly[ym])monthly[ym]={core:0,inv_r:0,inv_d:0,days:0};monthly[ym].core+=s[k].core;monthly[ym].inv_r+=s[k].inv_r;monthly[ym].inv_d+=s[k].inv_d;monthly[ym].days++;});
  const keys=Object.keys(monthly).sort();
  if(!keys.length){document.getElementById('monthly_table').innerHTML='<tr><td style="color:#94a3b8;padding:8px">No data yet</td></tr>';return;}
  let html=`<thead><tr><th>Month</th><th>Days</th><th>Core</th><th>Rec</th><th>Disp</th><th>C.Avg</th><th>I.Avg</th></tr></thead><tbody>`;
  keys.forEach(ym=>{const m=monthly[ym];const[y,mo]=ym.split('-');html+=`<tr><td style="color:#1a1a2e;font-weight:700">${MONTHS_SHORT[parseInt(mo)-1]} ${y}</td><td style="color:#2563eb">${m.days}</td><td style="color:#10b981">${m.core}</td><td style="color:#f97316">${m.inv_r}</td><td style="color:#eab308">${m.inv_d}</td><td style="color:#a855f7">${(m.core/m.days).toFixed(1)}</td><td style="color:#ef4444">${((m.inv_r+m.inv_d)/2/m.days).toFixed(1)}</td></tr>`;});
  document.getElementById('monthly_table').innerHTML=html+'</tbody>';
}

function buildMessage(){
  const date=selectedDate||formatDate(today),dStr=displayDate(date);
  const sH=store['H'][date]||{core:0,inv_r:0,inv_d:0},cH=calcUpTo('H',date);
  const sP=store['P15'][date]||{core:0,inv_r:0,inv_d:0},cP=calcUpTo('P15',date);
  return `${dStr} Recon Core Receipt, Invoice Received & Dispatched :\n_ _ _ _ _ _ _ _ _ _ _ _\nH - Recon Engine\nCore Receipt : ${String(sH.core).padStart(2,'0')}\nInvoice Received & Dispatched : ${sH.inv_r}/${sH.inv_d}\nCore Receipt Cumulative : ${String(cH.cumCore).padStart(2,'0')}\nInvoice & Dispatch Cumulative : ${cH.cumR}\nInvoice & Dispatch Average : ${cH.avgR}\nCore Average Receipt : ${cH.avgCore}\n_ _ _ _ _ _ _ _ _ _ _ _\nP15 - Recon Engine\nCore Receipt : ${String(sP.core).padStart(2,'0')}\nInvoice Received & Dispatched : ${sP.inv_r}/${sP.inv_d}\nCore Receipt Cumulative : ${String(cP.cumCore).padStart(2,'0')}\nInvoice & Dispatch Cumulative : ${cP.cumR}\nInvoice & Dispatch Average : ${cP.avgR}\nCore Average Receipt : ${cP.avgCore}\n_ _ _ _ _ _ _ _ _ _ _ _\nWarm Regards!!`;
}
function renderShare(){document.getElementById('share_preview').textContent=buildMessage();}
function shareWA(){window.open(`https://wa.me/?text=${encodeURIComponent(buildMessage())}`,'_blank');}
function shareEmail(){window.open(`mailto:?subject=${encodeURIComponent(`Recon Engine Report — ${displayDate(selectedDate||formatDate(today))}`)}&body=${encodeURIComponent(buildMessage())}`,'_blank');}

function openAllData(){
  let html='';
  ['H','P15'].forEach(tab=>{
    const s=store[tab],keys=Object.keys(s).sort();
    html+=`<div class="modal-section">${tab} Recon — ${keys.length} records</div>`;
    if(!keys.length){html+='<p style="color:#94a3b8;font-size:11px;padding:4px 0">No data</p>';return;}
    html+=`<table class="modal-table"><thead><tr><th>Date</th><th>Core</th><th>Inv Rec</th><th>Inv Disp</th><th>Saved At</th></tr></thead><tbody>`;
    keys.forEach(k=>{const r=s[k],d=new Date(r.ts);html+=`<tr><td style="font-weight:700;color:#1a1a2e">${displayDate(k)}</td><td style="color:#2563eb">${r.core}</td><td style="color:#f97316">${r.inv_r}</td><td style="color:#10b981">${r.inv_d}</td><td style="color:#94a3b8;font-size:9px">${d.toLocaleDateString('en-GB')} ${d.toLocaleTimeString('en-GB',{hour:'2-digit',minute:'2-digit'})}</td></tr>`;});
    html+=`</tbody></table>`;
  });
  document.getElementById('modal_body').innerHTML=html;
  document.getElementById('modal_overlay').classList.add('show');
}
function closeModal(){document.getElementById('modal_overlay').classList.remove('show');}
function closeModalOutside(e){if(e.target===document.getElementById('modal_overlay'))closeModal();}

function exportExcel(){
  function buildSheet(tab){
    const s=store[tab],keys=Object.keys(s).sort();
    const rows=[['Date','Core Receipt','Inv Received','Inv Dispatched','Core Cumulative','Inv Cumulative','Inv & Disp Avg','Core Avg']];
    let cc=0,cr=0,cd=0;
    keys.forEach((k,i)=>{
      const r=s[k];cc+=r.core;cr+=r.inv_r;cd+=r.inv_d;const n=i+1;
      const p=k.split('-');
      rows.push([`${p[2]}/${p[1]}/${p[0]}`,r.core,r.inv_r,r.inv_d,cc,cr+cd,+((cr+cd)/(2*n)).toFixed(1),+(cc/n).toFixed(1)]);
    });
    return rows;
  }
  const wb=XLSX.utils.book_new();
  const wsH=XLSX.utils.aoa_to_sheet(buildSheet('H'));
  const wsP=XLSX.utils.aoa_to_sheet(buildSheet('P15'));
  const wscols=[{wch:12},{wch:14},{wch:14},{wch:16},{wch:18},{wch:16},{wch:16},{wch:16}];
  wsH['!cols']=wscols;wsP['!cols']=wscols;
  XLSX.utils.book_append_sheet(wb,wsH,'H Recon');
  XLSX.utils.book_append_sheet(wb,wsP,'P15 Recon');
  XLSX.writeFile(wb,'Recon_Engine_Report.xlsx');
}

function renderAll(){
  renderCalendar();renderDaySummary();renderEntryForm();
  renderLegend();renderTrends();renderMonthly();renderShare();
}

/* ══════════════════════════════════════════
   BOOT
══════════════════════════════════════════ */
(async function boot(){
  const cfg=getConfig();
  if(cfg){
    API_KEY=cfg.key; BIN_ID=cfg.binId;
    setStatus('⏳ Loading from cloud…','#94a3b8');
    renderAll(); // show UI immediately
    await loadFromCloud();
    renderAll(); // re-render with cloud data
    // Auto-refresh every 60 seconds so teammates' changes appear
    setInterval(async()=>{
      await loadFromCloud();
      renderAll();
    },60000);
  }else{
    // No config — load local backup if any, then show setup
    store.H  =JSON.parse(localStorage.getItem('recon_bk_H' )||'{}');
    store.P15=JSON.parse(localStorage.getItem('recon_bk_P15')||'{}');
    setStatus('⚠ Setup needed','#f59e0b');
    renderAll();
    document.getElementById('setup_overlay').style.display='flex';
  }
})();
</script>
</body>
</html>
