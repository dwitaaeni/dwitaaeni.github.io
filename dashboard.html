<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover">
<title>Dashboard Distribusi &mdash; Sales · AR Aging · AR Payment · Stock</title>

<!-- PWA: installable "Add to Home Screen" on Android & iOS -->
<link rel="manifest" href="manifest.json">
<meta name="theme-color" content="#0A4F48">
<meta name="mobile-web-app-capable" content="yes">
<link rel="icon" href="icon-192.png" type="image/png">
<!-- iOS Safari-specific tags (Safari doesn't read the manifest for these) -->
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="Dashboard">
<link rel="apple-touch-icon" href="icon-192.png">

<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.5.0/chart.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
:root{
  --bg:#F4F7F6;
  --surface:#FFFFFF;
  --surface-alt:#EBF1EF;
  --ink:#152B26;
  --ink-soft:#516560;
  --ink-faint:#8A9C97;
  --primary:#0F6E64;
  --primary-dark:#0A4F48;
  --primary-tint:#E2F0EE;
  --accent:#D98F4E;
  --accent-tint:#FBECDB;
  --line:#DCE6E2;
  --line-soft:#EAF0EE;
  --risk-0:#4C8C7A;
  --risk-1:#CBA43A;
  --risk-2:#E08A3C;
  --risk-3:#C94F3F;
  --risk-4:#7A2E2E;
  --risk-0-tint:#E7F1EC;
  --risk-1-tint:#FBF3DC;
  --risk-2-tint:#FBEADB;
  --risk-3-tint:#F8E1DD;
  --risk-4-tint:#EEDCDC;
  --font-display:'Space Grotesk', 'IBM Plex Sans', sans-serif;
  --font-body:'IBM Plex Sans', -apple-system, sans-serif;
  --font-mono:'IBM Plex Mono', monospace;
  --shadow-sm: 0 1px 2px rgba(21,43,38,0.06), 0 1px 1px rgba(21,43,38,0.04);
  --shadow-md: 0 4px 16px rgba(21,43,38,0.08);
  --radius: 12px;
}
*{box-sizing:border-box;}
html,body{margin:0;padding:0;}
body{
  background:var(--bg);
  color:var(--ink);
  font-family:var(--font-body);
  font-size:14px;
  line-height:1.5;
  -webkit-font-smoothing:antialiased;
}
::selection{background:var(--primary-tint);}
button{font-family:inherit;}
.app{max-width:1400px;margin:0 auto;padding:0 24px 60px;}

/* ===== Header ===== */
.topbar{
  position:sticky; top:0; z-index:40;
  background:linear-gradient(180deg, var(--primary-dark) 0%, var(--primary) 100%);
  color:#fff;
  padding:0;
  box-shadow: var(--shadow-md);
}
.topbar-inner{
  max-width:1400px;margin:0 auto;padding:18px 24px 0;
  display:flex; align-items:flex-start; justify-content:space-between; gap:20px; flex-wrap:wrap;
}
.brand{display:flex; align-items:center; gap:12px;}
.brand-mark{
  width:38px;height:38px;border-radius:9px;
  background:rgba(255,255,255,0.16);
  display:flex;align-items:center;justify-content:center;
  font-family:var(--font-display); font-weight:700; font-size:16px;
  border:1px solid rgba(255,255,255,0.25);
}
.brand-text h1{
  font-family:var(--font-display); font-size:19px; font-weight:600; margin:0; letter-spacing:0.2px;
}
.brand-text p{margin:1px 0 0; font-size:12.5px; color:rgba(255,255,255,0.72);}
.topbar-actions{display:flex; gap:10px; align-items:center; padding-top:2px;}
.status-pill{
  display:flex; align-items:center; gap:6px;
  background:rgba(255,255,255,0.12); border:1px solid rgba(255,255,255,0.2);
  padding:6px 12px; border-radius:20px; font-size:12px; color:rgba(255,255,255,0.9);
}
.status-dot{width:7px;height:7px;border-radius:50%; background:var(--ink-faint); flex-shrink:0;}
.status-dot.on{background:#7FD9B9; box-shadow:0 0 0 3px rgba(127,217,185,0.25);}
.btn-upload{
  background:var(--accent); color:#3A2409; border:none;
  font-weight:600; font-size:13px; padding:9px 16px; border-radius:9px;
  cursor:pointer; display:flex; align-items:center; gap:7px;
  box-shadow:var(--shadow-sm);
  transition:transform .12s ease, box-shadow .12s ease;
}
.btn-upload:hover{transform:translateY(-1px); box-shadow:var(--shadow-md);}

/* ===== Tabs ===== */
.tabs{
  max-width:1400px;margin:0 auto; padding:14px 24px 0;
  display:flex; gap:4px; overflow-x:auto;
}
.tab-btn{
  background:transparent; border:none; color:rgba(255,255,255,0.68);
  padding:11px 18px 14px; font-size:14px; font-weight:600; cursor:pointer;
  border-bottom:3px solid transparent; white-space:nowrap;
  display:flex; align-items:center; gap:8px;
  transition:color .15s ease;
}
.tab-btn:hover{color:#fff;}
.tab-btn.active{color:#fff; border-bottom-color:var(--accent);}
.tab-count{
  background:rgba(255,255,255,0.16); padding:1px 7px; border-radius:10px; font-size:11px; font-weight:600;
}

/* ===== Layout ===== */
.view{display:none; padding-top:22px;}
.view.active{display:block; animation:fadeIn .25s ease;}
@keyframes fadeIn{from{opacity:0; transform:translateY(4px);} to{opacity:1; transform:translateY(0);}}

.section-head{display:flex; align-items:baseline; justify-content:space-between; margin:28px 0 14px; flex-wrap:wrap; gap:8px;}
.section-head h2{font-family:var(--font-display); font-size:16px; font-weight:600; margin:0; color:var(--ink);}
.section-head .hint{font-size:12.5px; color:var(--ink-faint);}
.section-head:first-child{margin-top:0;}

/* ===== Empty state ===== */
.empty-state{
  background:var(--surface); border:1.5px dashed var(--line); border-radius:var(--radius);
  padding:48px 24px; text-align:center; color:var(--ink-soft);
}
.empty-state .icon{font-size:30px; margin-bottom:10px;}
.empty-state h3{font-family:var(--font-display); font-size:16px; margin:0 0 6px; color:var(--ink);}
.empty-state p{margin:0 0 16px; font-size:13px;}
.empty-state button{
  background:var(--primary); color:#fff; border:none; padding:10px 18px; border-radius:9px;
  font-weight:600; font-size:13px; cursor:pointer;
}

/* ===== KPI Cards ===== */
.kpi-row{display:grid; grid-template-columns:repeat(auto-fit,minmax(200px,1fr)); gap:12px;}
.kpi-card{
  background:var(--surface); border:1px solid var(--line-soft); border-radius:var(--radius);
  padding:16px 18px; box-shadow:var(--shadow-sm);
}
.kpi-card .label{font-size:12px; color:var(--ink-faint); font-weight:600; text-transform:uppercase; letter-spacing:0.4px;}
.kpi-card .value{font-family:var(--font-mono); font-size:23px; font-weight:600; margin-top:6px; color:var(--ink);}
.kpi-card .sub{font-size:12px; margin-top:5px; color:var(--ink-soft);}
.kpi-card .sub.up{color:var(--risk-0);}
.kpi-card .sub.down{color:var(--risk-3);}
.kpi-card.warn{border-color:#EAD2B0; background:var(--risk-2-tint);}
.kpi-card.warn .value{color:var(--risk-2);}
.kpi-card.danger{border-color:#E3B8B4; background:var(--risk-3-tint);}
.kpi-card.danger .value{color:var(--risk-3);}
.kpi-card.good{border-color:#B9DACE; background:var(--risk-0-tint);}
.kpi-card.good .value{color:var(--primary-dark);}

/* ===== Panels / Cards ===== */
.panel{
  background:var(--surface); border:1px solid var(--line-soft); border-radius:var(--radius);
  padding:18px 20px; box-shadow:var(--shadow-sm);
}
.grid-2{display:grid; grid-template-columns:1.3fr 1fr; gap:14px;}
.grid-3{display:grid; grid-template-columns:repeat(3,1fr); gap:14px;}
@media(max-width:900px){.grid-2,.grid-3{grid-template-columns:1fr;}}
.panel h3{font-family:var(--font-display); font-size:14.5px; margin:0 0 14px; font-weight:600;}
.panel h3 .tag{font-family:var(--font-body); font-weight:500; font-size:11.5px; color:var(--ink-faint); margin-left:6px;}

/* ===== Filter bar ===== */
.filterbar{
  background:var(--surface); border:1px solid var(--line-soft); border-radius:var(--radius);
  padding:14px 16px; display:flex; gap:10px; flex-wrap:wrap; align-items:flex-end;
}
.filter-field{display:flex; flex-direction:column; gap:5px; min-width:130px;}
.filter-field label{font-size:11px; font-weight:600; color:var(--ink-faint); text-transform:uppercase; letter-spacing:0.3px;}
.filter-field select, .filter-field input{
  border:1px solid var(--line); border-radius:7px; padding:7px 9px; font-size:13px;
  font-family:var(--font-body); background:var(--surface); color:var(--ink); min-width:120px;
}
.filter-field input[type="search"]{min-width:180px;}
.filter-field select:focus, .filter-field input:focus{outline:2px solid var(--primary-tint); border-color:var(--primary);}
.btn-reset{
  background:var(--surface-alt); border:1px solid var(--line); color:var(--ink-soft);
  padding:8px 14px; border-radius:7px; font-size:12.5px; font-weight:600; cursor:pointer; height:34px;
}
.btn-reset:hover{background:var(--line-soft);}

/* ===== Table ===== */
.table-wrap{overflow-x:auto; border-radius:10px; border:1px solid var(--line-soft);}
table{width:100%; border-collapse:collapse; font-size:12.8px; background:var(--surface);}
thead th{
  background:var(--surface-alt); text-align:left; padding:10px 12px; font-weight:600;
  color:var(--ink-soft); white-space:nowrap; position:sticky; top:0; cursor:pointer; user-select:none;
  border-bottom:1px solid var(--line);
}
thead th:hover{color:var(--primary);}
thead th.sort-asc::after{content:' \2191';}
thead th.sort-desc::after{content:' \2193';}
tbody td{padding:9px 12px; border-bottom:1px solid var(--line-soft); white-space:nowrap; font-family:var(--font-mono); font-size:12.3px;}
tbody td.text{font-family:var(--font-body); font-size:12.6px;}
tbody tr:hover{background:var(--primary-tint);}
tbody tr:last-child td{border-bottom:none;}
.mono{font-family:var(--font-mono);}
.link-cell{color:var(--primary); font-weight:600; cursor:pointer; text-decoration:underline; text-decoration-style:dotted;}
.link-cell:hover{color:var(--primary-dark);}
.table-foot{display:flex; justify-content:space-between; align-items:center; padding:10px 4px 0; font-size:12px; color:var(--ink-faint);}
.pager{display:flex; gap:6px;}
.pager button{border:1px solid var(--line); background:var(--surface); border-radius:6px; padding:4px 10px; cursor:pointer; font-size:12px;}
.pager button:disabled{opacity:0.4; cursor:default;}

/* ===== Badges ===== */
.badge{display:inline-flex; align-items:center; gap:4px; padding:3px 9px; border-radius:20px; font-size:11px; font-weight:600; font-family:var(--font-body);}
.badge.b0{background:var(--risk-0-tint); color:var(--risk-0);}
.badge.b1{background:var(--risk-1-tint); color:#8A6D1E;}
.badge.b2{background:var(--risk-2-tint); color:var(--risk-2);}
.badge.b3{background:var(--risk-3-tint); color:var(--risk-3);}
.badge.b4{background:var(--risk-4-tint); color:var(--risk-4);}
.badge.neutral{background:var(--surface-alt); color:var(--ink-soft);}

/* ===== Risk gradient bar (signature element) ===== */
.risk-bar{position:relative; width:84px; height:7px; border-radius:4px; background:linear-gradient(90deg, var(--risk-0), var(--risk-1) 35%, var(--risk-2) 60%, var(--risk-3) 82%, var(--risk-4) 100%); opacity:0.28;}
.risk-bar .marker{position:absolute; top:-3px; width:3px; height:13px; background:var(--ink); border-radius:2px; box-shadow:0 0 0 2px #fff;}
.risk-wrap{display:flex; align-items:center; gap:8px;}
.risk-label{font-size:11.5px; font-family:var(--font-mono); font-weight:600; min-width:52px;}

/* ===== Alert banner ===== */
.alert-banner{
  display:flex; align-items:center; gap:12px; padding:13px 16px; border-radius:10px;
  background:var(--risk-3-tint); border:1px solid #E3B8B4; color:#7A342B; font-size:13px; margin-bottom:16px;
}
.alert-banner.warn{background:var(--risk-2-tint); border-color:#EAD2B0; color:#8A5A22;}
.alert-banner strong{font-weight:700;}
.alert-banner .icon{font-size:18px;}

/* ===== Upload modal ===== */
.modal-overlay{
  display:none; position:fixed; inset:0; background:rgba(15,30,27,0.5); z-index:100;
  align-items:center; justify-content:center; padding:20px;
}
.modal-overlay.open{display:flex;}
.modal{
  background:var(--surface); border-radius:16px; max-width:640px; width:100%;
  max-height:85vh; overflow-y:auto; box-shadow:0 20px 60px rgba(0,0,0,0.3);
}
.modal-head{padding:20px 24px; border-bottom:1px solid var(--line-soft); display:flex; justify-content:space-between; align-items:center;}
.modal-head h2{font-family:var(--font-display); font-size:17px; margin:0;}
.modal-close{background:none; border:none; font-size:20px; cursor:pointer; color:var(--ink-faint); line-height:1;}
.modal-body{padding:22px 24px;}
.dropzone{
  border:2px dashed var(--primary); border-radius:12px; background:var(--primary-tint);
  padding:34px 20px; text-align:center; cursor:pointer; transition:background .15s ease;
}
.dropzone:hover, .dropzone.drag{background:#D6ECE8;}
.dropzone .icon{font-size:28px; margin-bottom:8px;}
.dropzone h4{font-family:var(--font-display); margin:0 0 4px; font-size:14.5px;}
.dropzone p{margin:0; font-size:12.5px; color:var(--ink-soft);}
.upload-log{margin-top:16px; display:flex; flex-direction:column; gap:8px; max-height:220px; overflow-y:auto;}
.upload-log-item{
  display:flex; align-items:center; gap:10px; padding:9px 11px; border-radius:8px; background:var(--surface-alt); font-size:12.5px;
}
.upload-log-item .dot{width:8px;height:8px;border-radius:50%; flex-shrink:0;}
.upload-log-item.ok .dot{background:var(--risk-0);}
.upload-log-item.err .dot{background:var(--risk-3);}
.upload-log-item .name{font-weight:600; flex:1; overflow:hidden; text-overflow:ellipsis;}
.upload-log-item .detail{color:var(--ink-faint); font-size:11.5px;}
.data-status-list{margin-top:18px; border-top:1px solid var(--line-soft); padding-top:16px;}
.data-status-row{display:flex; align-items:center; justify-content:space-between; padding:9px 0; border-bottom:1px solid var(--line-soft); font-size:12.8px;}
.data-status-row:last-child{border-bottom:none;}
.data-status-row .name{font-weight:600;}
.data-status-row .meta{color:var(--ink-faint); font-size:12px;}
.btn-remove{background:none; border:1px solid var(--line); color:var(--risk-3); border-radius:6px; padding:4px 9px; font-size:11.5px; cursor:pointer;}
.btn-remove:hover{background:var(--risk-3-tint);}

/* ===== Detail modal (invoice drilldown) ===== */
.drawer-overlay{display:none; position:fixed; inset:0; background:rgba(15,30,27,0.45); z-index:110; align-items:center; justify-content:center; padding:20px;}
.drawer-overlay.open{display:flex;}
.drawer{background:var(--surface); border-radius:16px; max-width:760px; width:100%; max-height:82vh; overflow-y:auto; box-shadow:0 20px 60px rgba(0,0,0,0.3);}
.drawer.wide{max-width:960px;}
.cust-result-item{padding:10px 14px; cursor:pointer; font-size:13px; border-bottom:1px solid var(--line-soft);}
.cust-result-item:last-child{border-bottom:none;}
.cust-result-item:hover{background:var(--primary-tint);}
.cust-result-item .sub{color:var(--ink-faint); font-size:11.5px; margin-top:2px;}
.c360-head{display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:10px; margin-bottom:16px; padding-bottom:14px; border-bottom:1px solid var(--line-soft);}
.c360-head h3{font-family:var(--font-display); font-size:18px; margin:0;}
.c360-head .meta{font-size:12.5px; color:var(--ink-faint); margin-top:3px;}
.c360-section{margin-bottom:20px;}
.c360-section h4{font-family:var(--font-display); font-size:13.5px; margin:0 0 10px; color:var(--ink-soft); text-transform:uppercase; letter-spacing:0.3px;}
.attention-card{background:var(--surface); border:1px solid var(--line-soft); border-radius:var(--radius); padding:16px; box-shadow:var(--shadow-sm);}
.attention-card h4{font-family:var(--font-display); font-size:13.5px; margin:0 0 10px; display:flex; align-items:center; gap:7px;}
.attention-card ul{margin:0; padding:0; list-style:none;}
.attention-card li{padding:7px 0; border-bottom:1px solid var(--line-soft); font-size:12.3px; display:flex; justify-content:space-between; gap:8px;}
.attention-card li:last-child{border-bottom:none;}
.attention-card .empty{color:var(--ink-faint); font-size:12px; padding:8px 0;}
.customer-link{cursor:pointer; text-decoration:underline; text-decoration-style:dotted; color:inherit;}
.customer-link:hover{color:var(--primary);}

/* ===== Cross-tab heatmap ===== */
#table-crosstab{border-collapse:collapse; width:100%; font-size:12.5px;}
#table-crosstab th, #table-crosstab td{padding:8px 12px; border:1px solid var(--line-soft); text-align:right; white-space:nowrap; font-family:var(--font-mono);}
#table-crosstab thead th{background:var(--surface-alt); font-family:var(--font-body); font-weight:600; color:var(--ink-soft); text-align:center; position:sticky; top:0;}
#table-crosstab tbody th{background:var(--surface-alt); font-family:var(--font-body); font-weight:600; text-align:left; color:var(--ink); position:sticky; left:0;}
#table-crosstab td.cell{cursor:pointer; transition:transform .1s ease;}
#table-crosstab td.cell:hover{outline:2px solid var(--primary); outline-offset:-2px;}
#table-crosstab td.total-cell{background:var(--surface-alt); font-weight:700; color:var(--ink); cursor:default;}
#table-crosstab tr.total-row th, #table-crosstab tr.total-row td{background:var(--surface-alt); font-weight:700;}
.heat-0{background:var(--surface);}
.heat-1{background:#E1F5EE; color:#04342C;}
.heat-2{background:#9FE1CB; color:#04342C;}
.heat-3{background:#5DCAA5; color:#04342C;}
.heat-4{background:#1D9E75; color:#FFFFFF;}
.heat-5{background:#0B4F48; color:#FFFFFF;}

/* ===== Chips (multi-select-like, expandable rows) ===== */
.expand-row{cursor:pointer;}
.expand-row td:first-child::before{content:'\25B8 '; color:var(--ink-faint); font-family:var(--font-body);}
.expand-row.open td:first-child::before{content:'\25BE ';}
.sub-panel{background:var(--surface-alt); padding:0;}
.sub-panel-inner{padding:12px 20px;}

/* ===== Chart wrap ===== */
.chart-wrap{position:relative; height:260px;}
.chart-wrap.tall{height:320px;}

/* ===== Legend note ===== */
.note{font-size:11.5px; color:var(--ink-faint); margin-top:10px; line-height:1.5;}
.note code{background:var(--surface-alt); padding:1px 5px; border-radius:4px; font-family:var(--font-mono);}

/* ===== Loader ===== */
.spinner{width:16px;height:16px;border:2px solid rgba(255,255,255,0.3); border-top-color:#fff; border-radius:50%; animation:spin .7s linear infinite; display:inline-block;}
@keyframes spin{to{transform:rotate(360deg);}}

/* ===== scrollbar ===== */
::-webkit-scrollbar{height:8px; width:8px;}
::-webkit-scrollbar-thumb{background:#C7D6D1; border-radius:4px;}

@media(max-width:640px){
  .grid-2, .grid-3{grid-template-columns:1fr;}
  .topbar-inner{padding:14px 16px 0;}
  .tabs{padding:10px 16px 0;}
  .app{padding:0 14px 40px;}
}

/* ===== Export buttons ===== */
.export-bar{display:flex; gap:8px; margin-left:auto;}
.target-manage-toggle{cursor:pointer; user-select:none;}
.target-manage-toggle:hover h2{color:var(--primary);}
#targetManageArrow{display:inline-block; color:var(--ink-faint); font-size:13px; transition:transform .15s ease;}
.target-manage-toggle.open #targetManageArrow{transform:rotate(90deg);}
.btn-export{
  background:var(--surface); border:1px solid var(--line); color:var(--ink-soft);
  padding:8px 14px; border-radius:8px; font-size:12.5px; font-weight:600; cursor:pointer;
  display:flex; align-items:center; gap:6px; height:34px;
}
.btn-export:hover{background:var(--primary-tint); border-color:var(--primary); color:var(--primary-dark);}

/* ===== Print ===== */
.print-only{display:none;}
@media print{
  body{background:#fff;}
  .topbar, .modal-overlay, .drawer-overlay, .export-bar, .btn-reset, .empty-state, #openUploadBtn{display:none !important;}
  .filterbar{display:none !important;}
  .view, .view.active{display:none !important;}
  .print-only{display:block !important; padding:0 0 16px;}
  .print-only h1{font-family:var(--font-display); font-size:20px; margin:0 0 6px;}
  .print-only p{font-size:12px; color:#444; margin:2px 0;}
  .print-only .print-generated{color:#888; font-size:10.5px; margin-top:6px;}
  .print-table{width:100%; border-collapse:collapse; margin-top:14px; font-size:11px;}
  .print-table th, .print-table td{border:1px solid #bbb; padding:5px 7px; text-align:left;}
  .print-table th{background:#f0f0f0; font-weight:700;}
  .print-table tr{break-inside:avoid;}
}
</style>
</head>
<body>
<div id="print-header" class="print-only"></div>
<script id="embedded-data-slot" type="text/plain">__EMBEDDED_DATA_PLACEHOLDER__</script>

<div class="topbar">
  <div class="topbar-inner">
    <div class="brand">
      <div class="brand-mark">DB</div>
      <div class="brand-text">
        <h1>Dashboard Distribusi</h1>
        <p>Sales &middot; AR Aging &middot; AR Payment &middot; Stock Produk</p>
      </div>
    </div>
    <div class="topbar-actions">
      <div class="status-pill" id="storageModePill" title=""><span class="status-dot" id="storageModeDot"></span> <span id="storageModeLabel">Mengecek server&hellip;</span></div>
      <div class="status-pill"><span class="status-dot" id="statusDotSales"></span> Sales</div>
      <div class="status-pill"><span class="status-dot" id="statusDotAging"></span> AR Aging</div>
      <div class="status-pill"><span class="status-dot" id="statusDotPayment"></span> AR Payment</div>
      <div class="status-pill"><span class="status-dot" id="statusDotStock"></span> Stock</div>
      <button class="btn-upload" id="openUploadBtn">&#8593; Upload / Update Data</button>
    </div>
  </div>
  <nav class="tabs">
    <button class="tab-btn active" data-tab="home">Ringkasan</button>
    <button class="tab-btn" data-tab="sales">Data Sales</button>
    <button class="tab-btn" data-tab="aging">AR Aging</button>
    <button class="tab-btn" data-tab="payment">AR Payment</button>
    <button class="tab-btn" data-tab="stock">Stock Produk</button>
    <button class="tab-btn" data-tab="retention">Retensi Pelanggan</button>
    <button class="tab-btn" data-tab="evaluasi">Evaluasi Tim</button>
    <button class="tab-btn" data-tab="crosstab">Analisis Silang</button>
  </nav>
</div>

<div class="app">

  <!-- ============ HOME / RINGKASAN VIEW ============ -->
  <div class="view active" id="view-home">
    <div class="panel" style="margin-bottom:18px;">
      <h3>&#128269; Cari Pelanggan <span class="tag">lihat profil gabungan: pembelian, piutang, dan pembayaran</span></h3>
      <div style="display:flex; gap:8px; position:relative;">
        <input type="search" id="customer360-search" placeholder="Ketik nama pelanggan..." style="flex:1; border:1px solid var(--line); border-radius:8px; padding:10px 12px; font-size:13.5px; font-family:var(--font-body);">
        <div id="customer360-results" style="display:none; position:absolute; top:44px; left:0; right:0; background:var(--surface); border:1px solid var(--line); border-radius:10px; box-shadow:var(--shadow-md); max-height:280px; overflow-y:auto; z-index:20;"></div>
      </div>
    </div>

    <div id="home-alert-slot"></div>

    <div class="section-head"><h2>Ringkasan Hari Ini &amp; Minggu Ini</h2><span class="hint" id="home-daily-hint"></span></div>
    <div class="kpi-row" id="home-daily-kpi"></div>

    <div class="section-head"><h2>Ringkasan Bisnis</h2><span class="hint" id="home-asof-hint"></span></div>
    <div class="kpi-row" id="home-kpi"></div>

    <div class="grid-2" style="margin-top:14px;">
      <div class="panel">
        <h3>Penjualan 6 Bulan Terakhir</h3>
        <div class="chart-wrap"><canvas id="chart-home-sales"></canvas></div>
      </div>
      <div class="panel">
        <h3>Komposisi Piutang</h3>
        <div class="chart-wrap"><canvas id="chart-home-aging"></canvas></div>
      </div>
    </div>

    <div class="section-head"><h2>Perlu Perhatian Segera</h2></div>
    <div class="grid-3" id="home-attention"></div>

    <div id="home-empty" class="empty-state" style="display:none; margin-top:18px;">
      <div class="icon">&#128075;</div>
      <h3>Mulai dengan upload data</h3>
      <p>Upload minimal satu file data (Sales, AR Aging, AR Payment, atau Stock) untuk melihat ringkasan di sini.</p>
      <button onclick="openUploadModal()">Upload Data</button>
    </div>
  </div>

  <!-- ============ SALES VIEW ============ -->
  <div class="view" id="view-sales">
    <div id="sales-empty" class="empty-state" style="display:none;">
      <div class="icon">&#128202;</div>
      <h3>Belum ada data penjualan</h3>
      <p>Upload file database sales (export dari sistem) untuk mulai melihat dashboard.</p>
      <button onclick="openUploadModal()">Upload Data Sales</button>
    </div>
    <div id="sales-content" style="display:none;">
      <div class="filterbar" id="sales-filterbar"></div>
      <div class="section-head"><h2>Ringkasan</h2><span class="hint" id="sales-period-hint"></span></div>
      <div class="kpi-row" id="sales-kpi"></div>
      <div class="section-head"><h2>Tren Penjualan &amp; Perbandingan YoY</h2><span class="hint">Berdasarkan bulan pada rentang yang difilter</span></div>
      <div class="panel"><div class="chart-wrap tall"><canvas id="chart-sales-trend"></canvas></div></div>
      <div class="section-head"><h2>Analisis Produk &amp; Sales</h2></div>
      <div class="grid-2">
        <div class="panel">
          <h3>10 Produk Teratas <span class="tag" id="produk-chart-tag">berdasarkan nilai penjualan</span></h3>
          <div class="chart-wrap"><canvas id="chart-top-produk"></canvas></div>
        </div>
        <div class="panel">
          <h3>Kontribusi per Principal <span class="tag" id="principal-chart-tag"></span></h3>
          <div class="chart-wrap"><canvas id="chart-principal"></canvas></div>
        </div>
      </div>
      <div class="section-head" style="margin-top:14px;"><h2>Performa Sales</h2></div>
      <div class="panel">
        <div class="table-wrap"><table id="table-sales-performer"><thead></thead><tbody></tbody></table></div>
      </div>

      <div class="section-head"><h2>Rincian Produk per Salesman</h2><span class="hint">Produk apa saja yang dijual &mdash; dan tidak dijual &mdash; oleh masing-masing sales</span></div>
      <div class="panel" style="margin-bottom:14px;">
        <div class="filter-field" style="max-width:280px;"><label>Pilih Salesman</label><select id="f-salesman-detail"></select></div>
        <p class="note" id="salesman-detail-scope-note" style="margin-top:10px;"></p>
      </div>
      <div class="grid-2">
        <div class="panel">
          <h3>&#9989; Produk yang Dijual <span class="tag" id="salesman-sold-count"></span></h3>
          <div class="table-wrap"><table id="table-salesman-sold"><thead></thead><tbody></tbody></table></div>
        </div>
        <div class="panel">
          <h3>&#10060; Produk yang TIDAK Dijual <span class="tag" id="salesman-notsold-count"></span></h3>
          <div class="table-wrap"><table id="table-salesman-notsold"><thead></thead><tbody></tbody></table></div>
        </div>
      </div>

      <div class="section-head"><h2>Detail Transaksi</h2><span class="hint" id="sales-table-count"></span><div class="export-bar"><button class="btn-export" id="btn-export-sales-xlsx">&#8681; Excel</button><button class="btn-export" id="btn-export-sales-pdf">&#128196; PDF</button></div></div>
      <div class="panel">
        <div class="table-wrap"><table id="table-sales-detail"><thead></thead><tbody></tbody></table></div>
        <div class="table-foot">
          <span id="sales-pager-info"></span>
          <div class="pager" id="sales-pager"></div>
        </div>
      </div>
    </div>
  </div>

  <!-- ============ AGING VIEW ============ -->
  <div class="view" id="view-aging">
    <div id="aging-empty" class="empty-state" style="display:none;">
      <div class="icon">&#9203;</div>
      <h3>Belum ada data AR Aging</h3>
      <p>Upload file AR Aging (tarikan terupdate dari sistem) untuk melihat status piutang.</p>
      <button onclick="openUploadModal()">Upload Data AR Aging</button>
    </div>
    <div id="aging-content" style="display:none;">
      <div id="aging-alert-slot"></div>
      <div class="filterbar" id="aging-filterbar"></div>
      <div class="section-head"><h2>Ringkasan Piutang</h2><span class="hint" id="aging-asof-hint"></span></div>
      <div class="kpi-row" id="aging-kpi"></div>
      <div class="section-head"><h2>Distribusi Umur Piutang</h2></div>
      <div class="grid-2">
        <div class="panel">
          <div class="chart-wrap"><canvas id="chart-aging-bucket"></canvas></div>
        </div>
        <div class="panel">
          <h3>Top 10 Pelanggan &mdash; Piutang Terbesar</h3>
          <div class="table-wrap"><table id="table-top-debtor"><thead></thead><tbody></tbody></table></div>
        </div>
      </div>
      <div class="section-head"><h2>Detail Piutang per Invoice</h2><span class="hint">Klik nomor invoice untuk lihat detail produk</span><div class="export-bar"><button class="btn-export" id="btn-export-aging-xlsx">&#8681; Excel</button><button class="btn-export" id="btn-export-aging-pdf">&#128196; PDF</button></div></div>
      <div class="panel">
        <div class="table-wrap"><table id="table-aging-detail"><thead></thead><tbody></tbody></table></div>
        <div class="table-foot"><span id="aging-pager-info"></span><div class="pager" id="aging-pager"></div></div>
      </div>
    </div>
  </div>

  <!-- ============ PAYMENT VIEW ============ -->
  <div class="view" id="view-payment">
    <div id="payment-empty" class="empty-state" style="display:none;">
      <div class="icon">&#128176;</div>
      <h3>Belum ada data AR Payment</h3>
      <p>Upload file AR Payment untuk melihat riwayat pembayaran pelanggan.</p>
      <button onclick="openUploadModal()">Upload Data AR Payment</button>
    </div>
    <div id="payment-content" style="display:none;">
      <div class="filterbar" id="payment-filterbar"></div>
      <div class="section-head"><h2>Ringkasan Pembayaran</h2><span class="hint" id="payment-asof-hint"></span></div>
      <div class="kpi-row" id="payment-kpi"></div>
      <div class="grid-2">
        <div class="panel">
          <h3>Komposisi Metode Pembayaran</h3>
          <div class="chart-wrap"><canvas id="chart-payment-method"></canvas></div>
        </div>
        <div class="panel">
          <h3>Pembayaran Diterima per Sales</h3>
          <div class="chart-wrap"><canvas id="chart-payment-sales"></canvas></div>
        </div>
      </div>
      <div class="section-head"><h2>Detail Pembayaran</h2><span class="hint" id="payment-table-count"></span><div class="export-bar"><button class="btn-export" id="btn-export-payment-xlsx">&#8681; Excel</button><button class="btn-export" id="btn-export-payment-pdf">&#128196; PDF</button></div></div>
      <div class="panel">
        <div class="table-wrap"><table id="table-payment-detail"><thead></thead><tbody></tbody></table></div>
        <div class="table-foot"><span id="payment-pager-info"></span><div class="pager" id="payment-pager"></div></div>
      </div>
    </div>
  </div>

  <!-- ============ STOCK VIEW ============ -->
  <div class="view" id="view-stock">
    <div id="stock-empty" class="empty-state" style="display:none;">
      <div class="icon">&#128230;</div>
      <h3>Belum ada data Stock</h3>
      <p>Upload file Data Stock untuk melihat posisi stock, produk slow moving, dan peringatan stock menipis.</p>
      <button onclick="openUploadModal()">Upload Data Stock</button>
    </div>
    <div id="stock-content" style="display:none;">
      <div id="stock-alert-slot"></div>
      <div class="filterbar" id="stock-filterbar"></div>
      <div class="section-head"><h2>Ringkasan Stock</h2><span class="hint" id="stock-asof-hint"></span></div>
      <div class="kpi-row" id="stock-kpi"></div>

      <div class="section-head"><h2>&#128680; Peringatan Stock Menipis <span class="hint">estimasi hari hingga stock habis, berdasar kecepatan jual 90 hari terakhir</span></h2></div>
      <div class="panel">
        <div class="table-wrap"><table id="table-reorder"><thead></thead><tbody></tbody></table></div>
      </div>

      <div class="grid-2" style="margin-top:14px;">
        <div class="panel">
          <h3>&#128207; Produk Slow Moving <span class="tag">bottom 20% penjualan, masih ada stock</span></h3>
          <div class="table-wrap"><table id="table-slowmoving"><thead></thead><tbody></tbody></table></div>
        </div>
        <div class="panel">
          <h3>&#128293; Produk Terlaris <span class="tag" id="fastmoving-tag">top penjualan vs sisa stock</span></h3>
          <div class="table-wrap"><table id="table-fastmoving"><thead></thead><tbody></tbody></table></div>
        </div>
      </div>

      <div class="section-head"><h2>Produk Mendekati / Sudah Kadaluwarsa</h2></div>
      <div class="panel">
        <div class="table-wrap"><table id="table-expiry"><thead></thead><tbody></tbody></table></div>
      </div>

      <div class="section-head"><h2>Detail Stock per Lot</h2><span class="hint" id="stock-table-count"></span><div class="export-bar"><button class="btn-export" id="btn-export-stock-xlsx">&#8681; Excel</button><button class="btn-export" id="btn-export-stock-pdf">&#128196; PDF</button></div></div>
      <div class="panel">
        <div class="table-wrap"><table id="table-stock-detail"><thead></thead><tbody></tbody></table></div>
        <div class="table-foot"><span id="stock-pager-info"></span><div class="pager" id="stock-pager"></div></div>
      </div>
      <p class="note">Catatan: perhitungan kecepatan jual &amp; slow-moving mengasumsikan satuan Qty pada data Sales sama dengan satuan Stok Akhir untuk kode produk yang sama.</p>
    </div>
  </div>

  <!-- ============ RETENTION VIEW ============ -->
  <div class="view" id="view-retention">
    <div id="retention-empty" class="empty-state" style="display:none;">
      <div class="icon">&#128101;</div>
      <h3>Belum ada data Sales</h3>
      <p>Upload data Sales terlebih dahulu untuk melihat pelanggan mana yang aktif dan mana yang berhenti membeli.</p>
      <button onclick="openUploadModal()">Upload Data Sales</button>
    </div>
    <div id="retention-content" style="display:none;">
      <div class="filterbar" id="retention-filterbar"></div>
      <div class="section-head"><h2>Ringkasan Retensi</h2><span class="hint" id="retention-asof-hint"></span></div>
      <div class="kpi-row" id="retention-kpi"></div>
      <div class="grid-3" style="margin-top:14px;">
        <div class="panel">
          <h3>Distribusi Lama Tidak Membeli</h3>
          <div class="chart-wrap"><canvas id="chart-retention-dist"></canvas></div>
        </div>
        <div class="panel">
          <h3>Ringkasan per Cabang</h3>
          <div class="table-wrap"><table id="table-retention-cabang"><thead></thead><tbody></tbody></table></div>
        </div>
        <div class="panel">
          <h3>Ringkasan per Tipe Outlet</h3>
          <div class="table-wrap"><table id="table-retention-outlet"><thead></thead><tbody></tbody></table></div>
        </div>
      </div>
      <div id="retention-wilayah-wrap" style="display:none; margin-top:14px;">
        <div class="section-head"><h2>Ringkasan per Wilayah</h2><span class="hint">Dari database alamat pelanggan yang diupload &mdash; top 15 wilayah dengan pelanggan tidak aktif terbanyak</span></div>
        <div class="panel">
          <div class="table-wrap"><table id="table-retention-wilayah"><thead></thead><tbody></tbody></table></div>
        </div>
      </div>
      <div class="section-head"><h2>Pelanggan Tidak Aktif</h2><span class="hint" id="retention-table-count"></span><div class="export-bar"><button class="btn-export" id="btn-export-retention-xlsx">&#8681; Excel</button><button class="btn-export" id="btn-export-retention-pdf">&#128196; PDF</button></div></div>
      <div class="panel">
        <div class="table-wrap"><table id="table-retention-detail"><thead></thead><tbody></tbody></table></div>
        <div class="table-foot"><span id="retention-pager-info"></span><div class="pager" id="retention-pager"></div></div>
      </div>
    </div>
  </div>

  <!-- ============ EVALUASI TIM VIEW (YoY / MoM) ============ -->
  <div class="view" id="view-evaluasi">
    <div id="evaluasi-empty" class="empty-state" style="display:none;">
      <div class="icon">&#128202;</div>
      <h3>Belum ada data Sales</h3>
      <p>Upload data Sales (minimal 2 periode berbeda untuk MoM, atau 2 tahun untuk YoY) untuk melihat perbandingan performa.</p>
      <button onclick="openUploadModal()">Upload Data Sales</button>
    </div>
    <div id="evaluasi-content" style="display:none;">
      <div class="filterbar" id="evaluasi-filterbar"></div>
      <div class="section-head"><h2>Ringkasan Perbandingan</h2><span class="hint" id="evaluasi-period-hint"></span></div>
      <div class="kpi-row" id="evaluasi-kpi"></div>

      <div class="section-head"><h2>Performa per Salesman</h2><span class="hint">Klik baris salesman untuk lihat rincian produk yang naik/turun</span></div>
      <div class="panel">
        <div class="table-wrap"><table id="table-evaluasi-salesman"><thead></thead><tbody></tbody></table></div>
      </div>

      <div class="grid-2" style="margin-top:14px;">
        <div class="panel">
          <h3>&#128200; Top Produk Naik</h3>
          <div class="table-wrap"><table id="table-evaluasi-naik"><thead></thead><tbody></tbody></table></div>
        </div>
        <div class="panel">
          <h3>&#128201; Top Produk Turun</h3>
          <div class="table-wrap"><table id="table-evaluasi-turun"><thead></thead><tbody></tbody></table></div>
        </div>
      </div>

      <div class="section-head"><h2>Target &amp; Pencapaian</h2><span class="hint" id="target-scope-hint"></span></div>
      <div class="kpi-row" id="target-kpi"></div>
      <div class="section-head" style="margin-top:14px;"><h2>Pencapaian per Salesman</h2></div>
      <div class="panel">
        <div class="table-wrap"><table id="table-target-salesman"><thead></thead><tbody></tbody></table></div>
      </div>

      <div class="section-head" style="margin-top:14px;"><h2>Tren Pencapaian vs Target per Salesman</h2><span class="hint" id="target-trend-hint"></span></div>
      <div class="panel">
        <div class="filter-field" style="max-width:280px; margin-bottom:14px;"><label>Pilih Salesman</label><select id="f-target-trend-salesman"></select></div>
        <div class="chart-wrap tall"><canvas id="chart-target-trend"></canvas></div>
      </div>

      <div class="section-head target-manage-toggle" id="targetManageToggle">
        <h2><span id="targetManageArrow">&#9656;</span> Kelola Target Penjualan</h2>
        <span class="hint">Klik untuk buka/tutup &middot; input manual, atau upload Excel (kolom: Nama Salesman, Target) lewat panel Upload Data &middot; daftar salesman ikut filter Cabang di atas</span>
      </div>
      <div class="panel" id="targetManagePanel" style="display:none;">
        <div class="table-wrap"><table id="table-target-manage"><thead></thead><tbody></tbody></table></div>
        <div style="display:flex; justify-content:flex-end; margin-top:12px;">
          <button class="btn-upload" id="btn-save-targets" style="background:var(--primary); color:#fff;">Simpan Target</button>
        </div>
      </div>
    </div>
  </div>

  <!-- ============ ANALISIS SILANG (CROSS-TAB) VIEW ============ -->
  <div class="view" id="view-crosstab">
    <div id="crosstab-empty" class="empty-state" style="display:none;">
      <div class="icon">&#128202;</div>
      <h3>Belum ada data Sales</h3>
      <p>Upload data Sales untuk mulai membedah data per salesman, outlet, dan produk.</p>
      <button onclick="openUploadModal()">Upload Data Sales</button>
    </div>
    <div id="crosstab-content" style="display:none;">
      <div class="filterbar" id="crosstab-filterbar"></div>
      <div class="section-head"><h2>Matrix Silang</h2><span class="hint" id="crosstab-hint">Klik sel untuk lihat detail transaksinya</span></div>
      <div class="panel">
        <div class="table-wrap"><table id="table-crosstab"></table></div>
      </div>
      <p class="note">Kalau salah satu dimensi punya terlalu banyak nilai unik (misalnya Produk), yang ditampilkan hanya top 15 berdasarkan nilai terbesar &mdash; sisanya digabung sebagai "Lainnya".</p>

      <div class="section-head" style="margin-top:28px;"><h2>Riwayat Produk per Pelanggan</h2><span class="hint" id="custprod-count"></span></div>
      <div class="filterbar" id="custprod-filterbar"></div>
      <div class="panel" style="margin-top:14px;">
        <div class="table-wrap"><table id="table-custprod"><thead></thead><tbody></tbody></table></div>
        <div class="table-foot"><span id="custprod-pager-info"></span><div class="pager" id="custprod-pager"></div></div>
      </div>
    </div>
  </div>

</div>

<!-- ============ UPLOAD MODAL ============ -->
<div class="modal-overlay" id="uploadModal">
  <div class="modal">
    <div class="modal-head">
      <h2>Upload / Update Data</h2>
      <button class="modal-close" onclick="closeUploadModal()">&times;</button>
    </div>
    <div class="modal-body">
      <div class="dropzone" id="dropzone">
        <div class="icon">&#128194;</div>
        <h4>Tarik file ke sini, atau klik untuk pilih file</h4>
        <p>Bisa upload beberapa file sekaligus &middot; tipe data (Sales / AR Aging / AR Payment / Stock) terdeteksi otomatis dari kolomnya</p>
        <input type="file" id="fileInput" multiple style="display:none" accept=".xls,.xlsx,.csv,.txt">
      </div>
      <div class="upload-log" id="uploadLog"></div>
      <div class="data-status-list" id="dataStatusList"></div>
      <div id="exportWithDataWrap" style="display:none; margin-top:16px; padding-top:16px; border-top:1px solid var(--line-soft);">
        <div style="display:flex; justify-content:space-between; align-items:center; gap:12px; flex-wrap:wrap;">
          <p class="note" style="margin:0; max-width:420px;">Buat file dashboard baru yang datanya sudah tertanam di dalamnya &mdash; kirim file itu ke komputer lain, begitu dibuka datanya langsung ada (tetap bisa dihapus/upload ulang seperti biasa di komputer tujuan).</p>
          <button class="btn-upload" id="btnExportWithData" style="background:var(--primary); color:#fff; white-space:nowrap;">&#128190; Export Dashboard + Data</button>
        </div>
      </div>
      <div id="resetAllWrap" style="display:none; margin-top:16px; padding-top:16px; border-top:1px solid var(--line-soft); text-align:right;">
        <button class="btn-remove" id="btnResetAll" style="padding:8px 14px;">Hapus SEMUA Data (Reset Total)</button>
      </div>
    </div>
  </div>
</div>

<!-- ============ INVOICE DETAIL DRAWER ============ -->
<div class="drawer-overlay" id="invoiceDrawer">
  <div class="drawer" id="drawerBox">
    <div class="modal-head">
      <h2 id="drawerTitle">Detail Invoice</h2>
      <button class="modal-close" onclick="closeDrawer()">&times;</button>
    </div>
    <div class="modal-body" id="drawerBody"></div>
  </div>
</div>

<script>
// Captured immediately, before any rendering/DOM mutation happens, so "Export Dashboard + Data"
// can rebuild a clean copy of this file with the current data baked in.
const ORIGINAL_HTML = document.documentElement.outerHTML;

/* =========================================================================
   STATE
========================================================================= */
const DB = {
  sales: [],          // flat array of sales line records
  salesFiles: [],      // [{filename, cabang, rowCount, minDate, maxDate, uploadedAt}]
  aging: { rows: [], filename: null, uploadedAt: null },
  payment: { rows: [], filename: null, uploadedAt: null },
  stock: { rows: [], filename: null, uploadedAt: null },
  targets: { rows: [], filename: null, uploadedAt: null }, // [{salesman, targetBulanan}]
  customers: { rows: [], filename: null, uploadedAt: null } // [{idPelanggan, namaPelanggan, alamat, wilayah, kontakPerson, telp, status, ...}]
};

const TODAY = new Date(); TODAY.setHours(0,0,0,0);

/* =========================================================================
   UTILITIES
========================================================================= */
function fmtIDR(n){
  if(n===null||n===undefined||isNaN(n)) return '-';
  const neg = n<0; n = Math.abs(n);
  let s = 'Rp ' + Math.round(n).toLocaleString('id-ID');
  return neg ? '-'+s : s;
}
function fmtNumber(n){ if(n===null||n===undefined||isNaN(n)) return '-'; return (Math.round(n*100)/100).toLocaleString('id-ID'); }
function fmtDate(d){
  if(!d) return '-';
  const dd=String(d.getDate()).padStart(2,'0'), mm=String(d.getMonth()+1).padStart(2,'0'), yy=d.getFullYear();
  return dd+'/'+mm+'/'+yy;
}
function daysBetween(a,b){ return Math.round((a-b)/(1000*60*60*24)); }
function uid(){ return Math.random().toString(36).slice(2,9); }
function esc(s){ return (s===null||s===undefined) ? '' : String(s).replace(/[&<>"]/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }

/* Parse "M/D/YYYY" or "M/D/YYYY H:MM:SS AM/PM" strings into Date objects */
function parseUSDateTime(s){
  if(!s || typeof s!=='string') return null;
  s = s.trim();
  if(!s) return null;
  const parts = s.split(' ').filter(Boolean);
  const datePart = parts[0];
  const dm = datePart.split('/');
  if(dm.length<3) return null;
  const m = parseInt(dm[0],10), d = parseInt(dm[1],10);
  let y = parseInt(dm[2],10);
  if(isNaN(m)||isNaN(d)||isNaN(y)) return null;
  if(y < 100) y += 2000; // guard against 2-digit year formats (e.g. "6/2/26")
  let hh=0, mi=0, ss=0;
  if(parts.length>=3){
    const tp = parts[1].split(':').map(x=>parseInt(x,10));
    let H = tp[0]||0, M = tp[1]||0, S = tp[2]||0;
    const ap = parts[2].toUpperCase();
    if(ap==='PM' && H<12) H+=12;
    if(ap==='AM' && H===12) H=0;
    hh=H; mi=M; ss=S;
  }
  const dt = new Date(y, m-1, d, hh, mi, ss);
  return isNaN(dt.getTime()) ? null : dt;
}
// Handles both a raw Excel date serial number (from XLSX.js reading with raw:true — locale-
// independent) and a text date string (from the tab-delimited text files, or a text-typed
// Excel cell). Always prefer this over calling parseUSDateTime directly on upload-derived data.
function parseDateValue(v){
  if(v===null || v===undefined || v==='') return null;
  if(typeof v === 'number'){
    if(!isFinite(v) || v<=0) return null;
    // Excel serial date -> JS Date. 25569 = days between the Excel epoch (1899-12-30) and
    // the Unix epoch (1970-01-01); this is the standard, widely-used conversion.
    const ms = Math.round((v - 25569) * 86400 * 1000);
    const d = new Date(ms);
    return isNaN(d.getTime()) ? null : d;
  }
  return parseUSDateTime(String(v));
}
function toFloat(v){
  if(v===null||v===undefined||v==='') return 0;
  if(typeof v==='number') return v;
  const f = parseFloat(String(v).replace(/,/g,''));
  return isNaN(f) ? 0 : f;
}

/* =========================================================================
   FILE READING & TYPE DETECTION
========================================================================= */
async function readFileAsRows(file){
  const buf = await file.arrayBuffer();
  const bytes = new Uint8Array(buf);

  // Real xlsx (zip, starts with PK) or legacy binary xls (OLE2, D0 CF 11 E0)
  const isZip = bytes[0]===0x50 && bytes[1]===0x4B;
  const isOle = bytes[0]===0xD0 && bytes[1]===0xCF && bytes[2]===0x11 && bytes[3]===0xE0;
  if(isZip || isOle){
    if(!HAS_XLSX) throw new Error('Library pembaca Excel (XLSX) gagal dimuat dari CDN, coba refresh halaman');
    const wb = XLSX.read(buf, {type:'array'});
    const ws = wb.Sheets[wb.SheetNames[0]];
    // raw:true reads the underlying cell value (numbers as numbers, Excel date-cells as their
    // serial number) instead of a locale-formatted display string. This matters because if this
    // file is later opened/re-saved in Excel on a different computer with different regional
    // settings, a formatted date string (raw:false) can come out differently (or get misread)
    // depending on that computer's locale — the raw serial number is locale-independent.
    const rows = XLSX.utils.sheet_to_json(ws, {header:1, raw:true, defval:''});
    if(!rows.length) return {header:[], rows:[]};
    const header = rows[0].map(h=>String(h).trim());
    const dataRows = rows.slice(1).filter(r=>r.some(c=>String(c).trim()!==''));
    return {header, rows: dataRows};
  }

  // Text-based: detect UTF-16 (BOM or heuristic) vs UTF-8/ASCII
  let text;
  if(bytes[0]===0xFF && bytes[1]===0xFE){
    text = new TextDecoder('utf-16le').decode(buf);
  } else if(bytes[0]===0xFE && bytes[1]===0xFF){
    text = new TextDecoder('utf-16be').decode(buf);
  } else {
    let nullCount = 0; const sampleLen = Math.min(bytes.length, 4000);
    for(let i=1;i<sampleLen;i+=2){ if(bytes[i]===0) nullCount++; }
    if(nullCount > sampleLen*0.3){
      text = new TextDecoder('utf-16le').decode(buf);
    } else {
      text = new TextDecoder('utf-8').decode(buf);
    }
  }
  const lines = text.split(/\r\n|\n|\r/).filter(l=>l.trim().length>0);
  if(!lines.length) return {header:[], rows:[]};
  const header = lines[0].split('\t').map(h=>h.trim());
  const dataRows = lines.slice(1).map(l=>{
    const cells = l.split('\t');
    // trim/pad to header length (source files sometimes carry one trailing empty column)
    const row = new Array(header.length);
    for(let i=0;i<header.length;i++) row[i] = cells[i]!==undefined ? cells[i].trim() : '';
    return row;
  });
  return {header, rows: dataRows};
}

function colIndex(header, name){
  return header.findIndex(h=>h.toLowerCase()===name.toLowerCase());
}
function getCell(row, header, name, fallback){
  const i = colIndex(header, name);
  if(i<0 || row[i]===undefined) return fallback!==undefined?fallback:'';
  return row[i];
}

function detectFileType(header){
  const has = n => colIndex(header,n)>=0;
  const containsAny = (patterns) => header.some(h => patterns.some(p => h.toUpperCase().includes(p)));
  if(has('JUMLAH NETTO FAKTUR') && has('NO INVOICE')) return 'sales';
  if(has('TANGGAL JATUH TEMPO') && has('NO INVOICE')) return 'aging';
  if(has('SISA PEMB')) return 'payment';
  if(has('STOK AKHIR')) return 'stock';
  if(containsAny(['TARGET']) && containsAny(['SALESMAN','SALES'])) return 'target';
  if(has('Id Pelanggan') && has('Kelurahan') && has('Kecamatan')) return 'customer';
  return null;
}
function findColContains(header, patterns){
  return header.findIndex(h => patterns.some(p=>h.toUpperCase().includes(p)));
}
function buildTargetRecords(header, rows){
  const salesmanIdx = findColContains(header, ['SALESMAN','SALES']);
  const targetIdx = findColContains(header, ['TARGET']);
  if(salesmanIdx<0 || targetIdx<0) return [];
  return rows.map(r=>({
    salesman: (r[salesmanIdx]||'').trim(),
    targetBulanan: toFloat(r[targetIdx])
  })).filter(r=>r.salesman);
}

// Junk placeholder values seen in the Kelurahan column of the customer master data
// (e.g. "kl", "0") — excluded when composing a clean display address.
const ADDR_JUNK = new Set(['', '0', 'kl', '-', 'n/a', 'na', '.']);
function isJunkAddrPart(v){ return ADDR_JUNK.has((v||'').trim().toLowerCase()); }
function titleCase(s){
  let out = (s||'').trim().toLowerCase().replace(/\b\w/g, c=>c.toUpperCase());
  out = out.replace(/\bDki\b/g, 'DKI'); // keep "DKI Jakarta" as a proper acronym
  return out;
}
// Fixes "KOTABANDAR LAMPUNG" -> "Kota Bandar Lampung", "KAB.BOGOR" -> "Kab. Bogor",
// and normalizes casing/whitespace for the Kota (city/regency) field.
function cleanKotaLabel(raw){
  let s = (raw||'').trim();
  if(!s || s==='0') return '';
  s = s.replace(/\s+/g,' ');
  s = s.replace(/^(KOTA|KAB\.?)(?=[A-Za-z])/i, (m)=> m.toUpperCase().replace(/\.?$/,'.').replace('KOTA.','KOTA')+' ');
  s = titleCase(s);
  return s.trim();
}
function composeAlamat(parts){
  // parts: {alamat1, kelurahan, kecamatan, kota, propinsi, kodePos}
  const bits = [];
  if(parts.alamat1 && !isJunkAddrPart(parts.alamat1)) bits.push(titleCase(parts.alamat1));
  if(parts.kelurahan && !isJunkAddrPart(parts.kelurahan)) bits.push('Kel. '+titleCase(parts.kelurahan));
  if(parts.kecamatan && !isJunkAddrPart(parts.kecamatan)) bits.push('Kec. '+titleCase(parts.kecamatan));
  const kota = cleanKotaLabel(parts.kota);
  if(kota) bits.push(kota);
  if(parts.propinsi && !isJunkAddrPart(parts.propinsi)) bits.push(titleCase(parts.propinsi));
  if(parts.kodePos && parts.kodePos.trim() && parts.kodePos.trim()!=='0') bits.push(parts.kodePos.trim());
  return bits.join(', ');
}

function buildCustomerRecords(header, rows){
  return rows.map(r=>{
    const idPelanggan = getCell(r,header,'Id Pelanggan').trim();
    const kota = getCell(r,header,'Kota');
    const telp = getCell(r,header,'No HP').trim() || getCell(r,header,'No Telp 1').trim() || getCell(r,header,'No Telp 2').trim();
    const parts = {
      alamat1: getCell(r,header,'Alamat 1'),
      kelurahan: getCell(r,header,'Kelurahan'),
      kecamatan: getCell(r,header,'Kecamatan'),
      kota: kota,
      propinsi: getCell(r,header,'Propinsi'),
      kodePos: getCell(r,header,'Kode Pos')
    };
    return {
      idPelanggan,
      idPelangganKey: normProdCode(idPelanggan),
      namaPelanggan: getCell(r,header,'Nama Pelanggan').trim(),
      kontakPerson: getCell(r,header,'Kontak Person').trim(),
      telp,
      alamat: composeAlamat(parts),
      wilayah: cleanKotaLabel(kota) || '(Tidak diketahui)',
      kecamatan: titleCase(getCell(r,header,'Kecamatan')),
      status: getCell(r,header,'Status').trim(),
      latitude: toFloat(getCell(r,header,'Latitude')),
      longitude: toFloat(getCell(r,header,'Longitude'))
    };
  }).filter(r=>r.idPelanggan && r.namaPelanggan);
}

/* =========================================================================
   RECORD BUILDERS
========================================================================= */
// Groups product-name variants that only differ by a trailing carton/pack-size
// descriptor (e.g. "...60 GR" vs "...60 GR 70 BOTOL") into one display name.
// This NEVER touches raw namaProduk/idProduk/qty/value fields — it only adds
// an extra display label used for grouping in charts/tables.
function stripPackSuffix(name){
  let n = (name||'').trim();
  n = n.replace(/\s*\(\s*\d+\s*Lusin\s*\)\s*$/i, '');
  n = n.replace(/\s+\d+\s*(BOTOL|BTL|BOX|PCS|CTN)\s*$/i, '');
  return n.trim() || (name||'').trim();
}
function normProdCode(code){ return (code||'').trim().toUpperCase(); }

function buildSalesRecords(header, rows, filename){
  return rows.map(r=>{
    const tgl = parseDateValue(getCell(r,header,'TANGGAL'));
    const namaProduk = getCell(r,header,'NAMA PRODUK');
    return {
      cabang: getCell(r,header,'TEMPAT KERJA'),
      doId: getCell(r,header,'DO ID'),
      invoice: getCell(r,header,'NO INVOICE'),
      invoiceKey: getCell(r,header,'NO INVOICE').toUpperCase(),
      tanggal: tgl,
      idPelanggan: getCell(r,header,'ID PELANGGAN'),
      namaPelanggan: getCell(r,header,'NAMA PELANGGAN'),
      tipeOutlet: getCell(r,header,'JUMLAH OUTLET'),
      salesman: getCell(r,header,'Salesman'),
      principal: getCell(r,header,'PRINCIPAL'),
      groupProduk: getCell(r,header,'GROUP PRODUK'),
      idProduk: getCell(r,header,'ID PRODUK'),
      idProdukKey: normProdCode(getCell(r,header,'ID PRODUK')),
      namaProduk: namaProduk,
      namaProdukGroup: stripPackSuffix(namaProduk),
      qty: toFloat(getCell(r,header,'Qty')),
      hna: toFloat(getCell(r,header,'HNA')),
      discount: toFloat(getCell(r,header,'Rp Discount')),
      dpp: toFloat(getCell(r,header,'DASAR PENGENAAN PAJAK (DPP)')),
      ppn: toFloat(getCell(r,header,'PPN')),
      netto: toFloat(getCell(r,header,'JUMLAH NETTO FAKTUR')),
      year: tgl ? tgl.getFullYear() : null,
      month: tgl ? tgl.getMonth()+1 : null,
      sourceFile: filename
    };
  }).filter(r=>r.tanggal);
}

function buildAgingRecords(header, rows){
  return rows.map(r=>{
    const inv = getCell(r,header,'NO INVOICE');
    const tglInv = parseDateValue(getCell(r,header,'TANGGAL INVOICE'));
    const tglJt = parseDateValue(getCell(r,header,'TANGGAL JATUH TEMPO'));
    const totalSaldo = toFloat(getCell(r,header,'Rp Total Saldo'));
    const agingDays = tglJt ? daysBetween(TODAY, tglJt) : null;
    let bucket, bucketIdx;
    if(totalSaldo < 0){ bucket='Saldo Kredit'; bucketIdx=-1; }
    else if(agingDays===null){ bucket='Tanpa Tgl Jatuh Tempo'; bucketIdx=0; }
    else if(agingDays<=0){ bucket='Belum Jatuh Tempo'; bucketIdx=0; }
    else if(agingDays<=14){ bucket='1-14 Hari'; bucketIdx=1; }
    else if(agingDays<=30){ bucket='15-30 Hari'; bucketIdx=2; }
    else if(agingDays<=60){ bucket='31-60 Hari'; bucketIdx=3; }
    else if(agingDays<=90){ bucket='61-90 Hari'; bucketIdx=4; }
    else { bucket='> 90 Hari'; bucketIdx=5; }
    return {
      invoice: inv,
      invoiceKey: inv.toUpperCase(),
      tglInvoice: tglInv,
      tglJatuhTempo: tglJt,
      idPelanggan: getCell(r,header,'ID PELANGGAN'),
      namaPelanggan: getCell(r,header,'NAMA PELANGGAN'),
      currentDue: toFloat(getCell(r,header,'Rp Current Due')),
      b1_30: toFloat(getCell(r,header,'1-30')),
      b31_60: toFloat(getCell(r,header,'31-60')),
      b61_90: toFloat(getCell(r,header,'61-90')),
      bOver90: toFloat(getCell(r,header,'> 90')),
      notDue: toFloat(getCell(r,header,'Rp Not Due')),
      totalDue: toFloat(getCell(r,header,'Rp Total Due')),
      totalSaldo: totalSaldo,
      agingDays: agingDays,
      bucket, bucketIdx
    };
  }).filter(r=>r.invoice);
}

function buildPaymentRecords(header, rows){
  return rows.map(r=>{
    const sisa = toFloat(getCell(r,header,'SISA PEMB'));
    return {
      cabang: getCell(r,header,'NAMA TEMPAT KERJA'),
      sales: getCell(r,header,'NAMA SALES'),
      namaPelanggan: getCell(r,header,'NAMA PELANGGAN'),
      invoice: getCell(r,header,'INVOICE'),
      invoiceKey: getCell(r,header,'INVOICE').toUpperCase(),
      tglInvoice: parseDateValue(getCell(r,header,'TGL INVOICE')),
      nilaiInvoice: toFloat(getCell(r,header,'NILAI INVOICE')),
      pembTunai: toFloat(getCell(r,header,'PEMB TUNAI')),
      pembDeposit: toFloat(getCell(r,header,'PEMB DEPOSIT')),
      pembCek: toFloat(getCell(r,header,'PEMB CEK')),
      tglPembTerakhir: parseDateValue(getCell(r,header,'TGL PEMB TERAKHIR')),
      sisaPemb: sisa,
      lunas: sisa <= 1
    };
  }).filter(r=>r.invoice);
}

function buildStockRecords(header, rows){
  return rows.map(r=>{
    return {
      cabang: getCell(r,header,'NAMA TEMPAT KERJA'),
      idProduk: getCell(r,header,'ID PRODUK'),
      idProdukKey: normProdCode(getCell(r,header,'ID PRODUK')),
      namaProduk: getCell(r,header,'NAMA PRODUK'),
      satuanKomposit: getCell(r,header,'SATUAN KOMPOSIT'),
      satuan: getCell(r,header,'SATUAN'),
      locationId: getCell(r,header,'LOCATION ID'),
      idLot: getCell(r,header,'ID LOT'),
      tglKadaluwarsa: parseDateValue(getCell(r,header,'TANGGAL KADALUWARSA')),
      dist: toFloat(getCell(r,header,'DIST')),
      stokAkhir: toFloat(getCell(r,header,'STOK AKHIR'))
    };
  }).filter(r=>r.idProduk);
}

/* =========================================================================
   STORAGE LAYER (IndexedDB-backed)
   Uses the browser's own IndexedDB instead of window.storage, since this
   file may be opened as a standalone HTML file (double-clicked / downloaded)
   where the Artifacts-only window.storage API is not injected.
========================================================================= */
const LocalKV = (function(){
  const DB_NAME = 'dashboard_distribusi_db';
  const STORE = 'kv';
  let dbPromise = null;
  function openDB(){
    if(dbPromise) return dbPromise;
    dbPromise = new Promise((resolve,reject)=>{
      if(!('indexedDB' in window)){ reject(new Error('IndexedDB tidak didukung browser ini')); return; }
      const req = indexedDB.open(DB_NAME, 1);
      req.onupgradeneeded = e=>{
        const db = e.target.result;
        if(!db.objectStoreNames.contains(STORE)) db.createObjectStore(STORE);
      };
      req.onsuccess = e=>resolve(e.target.result);
      req.onerror = e=>reject(e.target.error);
    });
    return dbPromise;
  }
  async function get(key){
    const db = await openDB();
    return new Promise((resolve,reject)=>{
      const tx = db.transaction(STORE,'readonly');
      const req = tx.objectStore(STORE).get(key);
      req.onsuccess = ()=> resolve(req.result!==undefined ? {key, value:req.result} : null);
      req.onerror = ()=>reject(req.error);
    });
  }
  async function set(key, value){
    const db = await openDB();
    return new Promise((resolve,reject)=>{
      const tx = db.transaction(STORE,'readwrite');
      tx.objectStore(STORE).put(value, key);
      tx.oncomplete = ()=>resolve({key,value});
      tx.onerror = ()=>reject(tx.error);
    });
  }
  async function del(key){
    const db = await openDB();
    return new Promise((resolve,reject)=>{
      const tx = db.transaction(STORE,'readwrite');
      tx.objectStore(STORE).delete(key);
      tx.oncomplete = ()=>resolve({key, deleted:true});
      tx.onerror = ()=>reject(tx.error);
    });
  }
  return { get, set, delete: del };
})();
// Thin wrapper matching the {get(key,shared), set(key,value,shared), delete(key,shared)} shape
// used throughout this file; the "shared" flag is irrelevant for a local single-user file.
// Falls back to an in-memory store (no persistence across reloads) if IndexedDB itself
// is unavailable, so the dashboard still works for the current session either way.
const MemoryFallback = new Map();
let usingMemoryFallback = false;
const LocalStorageBackend = {
  get: async (key)=>{
    if(usingMemoryFallback) return MemoryFallback.has(key) ? {key, value:MemoryFallback.get(key)} : null;
    try{ return await LocalKV.get(key); }
    catch(e){ usingMemoryFallback = true; console.warn('IndexedDB tidak tersedia, memakai penyimpanan sementara (tidak persisten):', e); return MemoryFallback.has(key) ? {key, value:MemoryFallback.get(key)} : null; }
  },
  set: async (key, value)=>{
    if(usingMemoryFallback){ MemoryFallback.set(key,value); return {key,value}; }
    try{ return await LocalKV.set(key, value); }
    catch(e){ usingMemoryFallback = true; console.warn('IndexedDB tidak tersedia, memakai penyimpanan sementara (tidak persisten):', e); MemoryFallback.set(key,value); return {key,value}; }
  },
  delete: async (key)=>{
    if(usingMemoryFallback){ MemoryFallback.delete(key); return {key, deleted:true}; }
    try{ return await LocalKV.delete(key); }
    catch(e){ usingMemoryFallback = true; MemoryFallback.delete(key); return {key, deleted:true}; }
  }
};

// Server backend (api.php) — used when this dashboard is hosted alongside api.php,
// so every visitor reads/writes the SAME shared data instead of each browser's own
// local storage. Falls back to LocalStorageBackend automatically if api.php isn't reachable.
const API_URL = 'api.php';
let serverModeAvailable = false;
let serverModeChecked = false;
async function detectServerMode(){
  try{
    const res = await fetch(API_URL + '?key=__ping__', {method:'GET'});
    serverModeAvailable = res.ok;
  }catch(e){ serverModeAvailable = false; }
  serverModeChecked = true;
}
function updateStorageModeIndicator(){
  const dot = document.getElementById('storageModeDot');
  const label = document.getElementById('storageModeLabel');
  const pill = document.getElementById('storageModePill');
  if(!dot || !label) return;
  if(serverModeAvailable){
    dot.classList.add('on');
    label.textContent = 'Server (data dibagikan)';
    pill.title = 'Data tersimpan di server melalui api.php — semua orang yang membuka dashboard ini melihat data yang sama.';
  } else {
    dot.classList.remove('on');
    label.textContent = 'Lokal (browser ini saja)';
    pill.title = 'api.php tidak ditemukan/tidak bisa diakses — data hanya tersimpan di browser ini dan tidak dibagikan ke orang lain.';
  }
}
const ServerStorageBackend = {
  get: async (key)=>{
    const res = await fetch(API_URL + '?key=' + encodeURIComponent(key));
    if(!res.ok) throw new Error('Server API error: HTTP '+res.status);
    const json = await res.json();
    return (json.value!==null && json.value!==undefined) ? {key, value:json.value} : null;
  },
  set: async (key, value)=>{
    const res = await fetch(API_URL, {method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({action:'set', key, value})});
    if(!res.ok) throw new Error('Server API error: HTTP '+res.status);
    return {key, value};
  },
  delete: async (key)=>{
    const res = await fetch(API_URL, {method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({action:'delete', key})});
    if(!res.ok) throw new Error('Server API error: HTTP '+res.status);
    return {key, deleted:true};
  }
};

const storage = {
  get: async (key)=>{
    if(serverModeAvailable){ try{ return await ServerStorageBackend.get(key); }catch(e){ console.warn('Server API gagal, pakai penyimpanan lokal:', e); } }
    return LocalStorageBackend.get(key);
  },
  set: async (key, value)=>{
    if(serverModeAvailable){ try{ return await ServerStorageBackend.set(key, value); }catch(e){ console.warn('Server API gagal, pakai penyimpanan lokal:', e); } }
    return LocalStorageBackend.set(key, value);
  },
  delete: async (key)=>{
    if(serverModeAvailable){ try{ return await ServerStorageBackend.delete(key); }catch(e){ console.warn('Server API gagal, pakai penyimpanan lokal:', e); } }
    return LocalStorageBackend.delete(key);
  }
};

/* =========================================================================
   PERSISTENCE (Sales / Aging / Payment / Stock)
========================================================================= */
async function saveSalesFile(filename, records, meta){
  await storage.set('sales_data:'+filename, JSON.stringify(records));
  const idx = DB.salesFiles.findIndex(f=>f.filename===filename);
  if(idx>=0) DB.salesFiles[idx] = meta; else DB.salesFiles.push(meta);
  await storage.set('sales_files_index', JSON.stringify(DB.salesFiles));
}
async function removeSalesFile(filename){
  if(!confirm('Hapus data Sales dari file "'+filename+'"? Tindakan ini tidak bisa dibatalkan.')) return;
  try{ await storage.delete('sales_data:'+filename); }catch(e){}
  DB.salesFiles = DB.salesFiles.filter(f=>f.filename!==filename);
  await storage.set('sales_files_index', JSON.stringify(DB.salesFiles));
  rebuildSalesFromFiles();
}
async function removeAging(){
  if(!confirm('Hapus seluruh data AR Aging yang sudah diupload? Tindakan ini tidak bisa dibatalkan.')) return;
  try{ await storage.delete('aging_data'); }catch(e){}
  DB.aging = {rows:[], filename:null, uploadedAt:null};
  renderAll(); renderDataStatusList();
}
async function removePayment(){
  if(!confirm('Hapus seluruh data AR Payment yang sudah diupload? Tindakan ini tidak bisa dibatalkan.')) return;
  try{ await storage.delete('payment_data'); }catch(e){}
  DB.payment = {rows:[], filename:null, uploadedAt:null};
  invalidateCabangLookup();
  renderAll(); renderDataStatusList();
}
async function removeStock(){
  if(!confirm('Hapus seluruh data Stock yang sudah diupload? Tindakan ini tidak bisa dibatalkan.')) return;
  try{ await storage.delete('stock_data'); }catch(e){}
  DB.stock = {rows:[], filename:null, uploadedAt:null};
  renderAll(); renderDataStatusList();
}
async function removeTargets(){
  if(!confirm('Hapus seluruh data Target Penjualan yang sudah diupload/diinput? Tindakan ini tidak bisa dibatalkan.')) return;
  try{ await storage.delete('target_data'); }catch(e){}
  DB.targets = {rows:[], filename:null, uploadedAt:null};
  renderAll(); renderDataStatusList();
}
async function removeCustomers(){
  if(!confirm('Hapus seluruh data Database Pelanggan (alamat) yang sudah diupload? Tindakan ini tidak bisa dibatalkan.')) return;
  try{ await storage.delete('customer_data'); }catch(e){}
  DB.customers = {rows:[], filename:null, uploadedAt:null};
  invalidateCustomerAddressCache();
  renderAll(); renderDataStatusList();
}
async function rebuildSalesFromFiles(){
  let all = [];
  for(const f of DB.salesFiles){
    try{
      const res = await storage.get('sales_data:'+f.filename);
      if(res && res.value) all = all.concat(JSON.parse(res.value));
    }catch(e){}
  }
  all.forEach(r=>{ if(r.tanggal) r.tanggal = new Date(r.tanggal); });
  DB.sales = all;
  invalidateCabangLookup();
  invalidateSalesmanCabangCache();
  renderAll();
}
async function saveAging(records, filename){
  DB.aging = {rows:records, filename, uploadedAt:new Date().toISOString()};
  await storage.set('aging_data', JSON.stringify(DB.aging));
}
async function savePayment(records, filename){
  DB.payment = {rows:records, filename, uploadedAt:new Date().toISOString()};
  await storage.set('payment_data', JSON.stringify(DB.payment));
  invalidateCabangLookup();
}
async function saveStock(records, filename){
  DB.stock = {rows:records, filename, uploadedAt:new Date().toISOString()};
  await storage.set('stock_data', JSON.stringify(DB.stock));
}
async function saveTargets(records, filename){
  DB.targets = {rows:records, filename, uploadedAt:new Date().toISOString()};
  await storage.set('target_data', JSON.stringify(DB.targets));
}
async function saveCustomers(records, filename){
  DB.customers = {rows:records, filename, uploadedAt:new Date().toISOString()};
  await storage.set('customer_data', JSON.stringify(DB.customers));
  invalidateCustomerAddressCache();
}

/* =========================================================================
   EMBEDDED DATA (baked-in seed for "Export Dashboard + Data")
========================================================================= */
function loadEmbeddedSeed(){
  const el = document.getElementById('embedded-data-slot');
  if(!el) return null;
  const raw = el.textContent.trim();
  if(!raw || raw === '__EMBEDDED_DATA_PLACEHOLDER__') return null;
  try{
    const json = decodeURIComponent(escape(atob(raw)));
    return JSON.parse(json);
  }catch(e){ console.warn('Gagal membaca data bawaan file:', e); return null; }
}

async function hydrateFromSeed(seed){
  // Write the seed into local storage so it becomes the normal working copy from here on
  // (subsequent deletes/re-uploads in this browser work exactly like unembedded usage).
  if(seed.salesFiles && seed.salesDataByFile){
    for(const f of seed.salesFiles){
      const records = seed.salesDataByFile[f.filename];
      if(records) await storage.set('sales_data:'+f.filename, JSON.stringify(records));
    }
    await storage.set('sales_files_index', JSON.stringify(seed.salesFiles));
  }
  if(seed.aging) await storage.set('aging_data', JSON.stringify(seed.aging));
  if(seed.payment) await storage.set('payment_data', JSON.stringify(seed.payment));
  if(seed.stock) await storage.set('stock_data', JSON.stringify(seed.stock));
  if(seed.targets) await storage.set('target_data', JSON.stringify(seed.targets));
  if(seed.customers) await storage.set('customer_data', JSON.stringify(seed.customers));
}

async function exportDashboardWithData(){
  const seed = { salesFiles: DB.salesFiles, salesDataByFile: {}, aging: DB.aging, payment: DB.payment, stock: DB.stock, targets: DB.targets, customers: DB.customers };
  for(const f of DB.salesFiles){
    try{
      const res = await storage.get('sales_data:'+f.filename);
      if(res && res.value) seed.salesDataByFile[f.filename] = JSON.parse(res.value);
    }catch(e){}
  }
  const anyData = DB.salesFiles.length || DB.aging.filename || DB.payment.filename || DB.stock.filename || DB.targets.filename || DB.customers.filename;
  if(!anyData){ alert('Belum ada data untuk di-export. Upload data dulu.'); return; }
  const json = JSON.stringify(seed);
  const b64 = btoa(unescape(encodeURIComponent(json)));
  let html = ORIGINAL_HTML;
  html = html.replace(
    /(<script[^>]*id="embedded-data-slot"[^>]*>)[\s\S]*?(<\/script>)/,
    '$1' + b64 + '$2'
  );
  html = '<!DOCTYPE html>\n' + html;
  const blob = new Blob([html], {type:'text/html'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'dashboard-dengan-data-'+fmtDate(new Date()).replace(/\//g,'-')+'.html';
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

async function loadAllFromStorage(){
  try{
    const idx = await storage.get('sales_files_index');
    if(idx && idx.value) DB.salesFiles = JSON.parse(idx.value);
  }catch(e){}
  if(DB.salesFiles.length){
    let all = [];
    for(const f of DB.salesFiles){
      try{
        const res = await storage.get('sales_data:'+f.filename);
        if(res && res.value) all = all.concat(JSON.parse(res.value));
      }catch(e){}
    }
    all.forEach(r=>{ if(r.tanggal) r.tanggal = new Date(r.tanggal); });
    DB.sales = all;
  }
  try{
    const a = await storage.get('aging_data');
    if(a && a.value){
      DB.aging = JSON.parse(a.value);
      DB.aging.rows.forEach(r=>{ if(r.tglInvoice) r.tglInvoice=new Date(r.tglInvoice); if(r.tglJatuhTempo) r.tglJatuhTempo=new Date(r.tglJatuhTempo); });
    }
  }catch(e){}
  try{
    const p = await storage.get('payment_data');
    if(p && p.value){
      DB.payment = JSON.parse(p.value);
      DB.payment.rows.forEach(r=>{ if(r.tglInvoice) r.tglInvoice=new Date(r.tglInvoice); if(r.tglPembTerakhir) r.tglPembTerakhir=new Date(r.tglPembTerakhir); });
    }
  }catch(e){}
  try{
    const s = await storage.get('stock_data');
    if(s && s.value){
      DB.stock = JSON.parse(s.value);
      DB.stock.rows.forEach(r=>{ if(r.tglKadaluwarsa) r.tglKadaluwarsa=new Date(r.tglKadaluwarsa); });
    }
  }catch(e){}
  try{
    const t = await storage.get('target_data');
    if(t && t.value) DB.targets = JSON.parse(t.value);
  }catch(e){}
  try{
    const c = await storage.get('customer_data');
    if(c && c.value) DB.customers = JSON.parse(c.value);
  }catch(e){}
}

/* =========================================================================
   UPLOAD HANDLING
========================================================================= */
function logUpload(name, ok, detail){
  const log = document.getElementById('uploadLog');
  const div = document.createElement('div');
  div.className = 'upload-log-item ' + (ok?'ok':'err');
  div.innerHTML = '<span class="dot"></span><span class="name">'+esc(name)+'</span><span class="detail">'+esc(detail)+'</span>';
  log.prepend(div);
}

async function handleFiles(fileList){
  for(const file of Array.from(fileList)){
    try{
      const {header, rows} = await readFileAsRows(file);
      if(!header.length){ logUpload(file.name, false, 'File kosong atau tidak terbaca'); continue; }
      const type = detectFileType(header);
      if(type==='sales'){
        let recs = buildSalesRecords(header, rows, file.name);
        if(!recs.length){ logUpload(file.name, false, 'Terdeteksi sebagai Sales tapi 0 baris valid'); continue; }
        const minYear = 2015, maxYear = TODAY.getFullYear()+1;
        const implausible = recs.filter(r=>r.year<minYear || r.year>maxYear);
        if(implausible.length){
          recs = recs.filter(r=>r.year>=minYear && r.year<=maxYear);
          const badYears = [...new Set(implausible.map(r=>r.year))].sort().join(', ');
          logUpload(file.name, false, implausible.length+' baris dilewati karena tanggalnya tidak masuk akal (tahun: '+badYears+') &mdash; kemungkinan file rusak saat dibuka/disimpan ulang di Excel. Cek kolom TANGGAL di file sumber.');
          if(!recs.length) continue;
        }
        const dates = recs.map(r=>r.tanggal).filter(Boolean);
        const cabangSet = [...new Set(recs.map(r=>r.cabang))].join(', ');
        const meta = {
          filename:file.name, cabang:cabangSet, rowCount:recs.length,
          minDate: new Date(Math.min(...dates)).toISOString(),
          maxDate: new Date(Math.max(...dates)).toISOString(),
          uploadedAt: new Date().toISOString()
        };
        await saveSalesFile(file.name, recs, meta);
        await rebuildSalesFromFiles();
        logUpload(file.name, true, 'Sales &middot; '+recs.length+' baris &middot; '+cabangSet);
      } else if(type==='aging'){
        const recs = buildAgingRecords(header, rows);
        if(!recs.length){ logUpload(file.name, false, 'Terdeteksi sebagai AR Aging tapi 0 baris valid'); continue; }
        await saveAging(recs, file.name);
        logUpload(file.name, true, 'AR Aging &middot; '+recs.length+' invoice (menggantikan data lama)');
      } else if(type==='payment'){
        const recs = buildPaymentRecords(header, rows);
        if(!recs.length){ logUpload(file.name, false, 'Terdeteksi sebagai AR Payment tapi 0 baris valid'); continue; }
        await savePayment(recs, file.name);
        logUpload(file.name, true, 'AR Payment &middot; '+recs.length+' baris (menggantikan data lama)');
      } else if(type==='stock'){
        const recs = buildStockRecords(header, rows);
        if(!recs.length){ logUpload(file.name, false, 'Terdeteksi sebagai Stock tapi 0 baris valid'); continue; }
        await saveStock(recs, file.name);
        logUpload(file.name, true, 'Stock &middot; '+recs.length+' baris (menggantikan data lama)');
      } else if(type==='target'){
        const recs = buildTargetRecords(header, rows);
        if(!recs.length){ logUpload(file.name, false, 'Terdeteksi sebagai Target tapi 0 baris valid'); continue; }
        await saveTargets(recs, file.name);
        logUpload(file.name, true, 'Target Penjualan &middot; '+recs.length+' salesman (menggantikan data lama)');
      } else if(type==='customer'){
        const recs = buildCustomerRecords(header, rows);
        if(!recs.length){ logUpload(file.name, false, 'Terdeteksi sebagai Database Pelanggan tapi 0 baris valid'); continue; }
        await saveCustomers(recs, file.name);
        logUpload(file.name, true, 'Database Pelanggan (alamat) &middot; '+recs.length+' pelanggan (menggantikan data lama)');
      } else {
        logUpload(file.name, false, 'Tipe data tidak dikenali &mdash; cek header kolom file');
        continue;
      }
      renderAll();
      renderDataStatusList();
    }catch(err){
      console.error(err);
      logUpload(file.name, false, 'Gagal diproses: '+err.message);
    }
  }
}

function renderDataStatusList(){
  const el = document.getElementById('dataStatusList');
  let html = '';
  if(usingMemoryFallback){
    html += '<div class="alert-banner warn" style="margin-bottom:12px;"><span class="icon">&#9888;</span><div>Penyimpanan permanen (IndexedDB) tidak tersedia di browser/tampilan ini &mdash; data yang diupload tetap bisa dipakai sekarang, tapi akan hilang kalau halaman ditutup/refresh. Coba buka file ini langsung di browser (bukan preview) untuk penyimpanan permanen.</div></div>';
  }
  if(DB.salesFiles.length){
    html += '<div style="font-weight:600; font-size:12.5px; margin-bottom:6px; color:var(--ink-soft);">Data Sales ('+DB.salesFiles.length+' file)</div>';
    DB.salesFiles.slice().sort((a,b)=>new Date(b.uploadedAt)-new Date(a.uploadedAt)).forEach(f=>{
      html += '<div class="data-status-row"><div><div class="name">'+esc(f.filename)+'</div><div class="meta">'+f.cabang+' &middot; '+f.rowCount+' baris &middot; '+fmtDate(new Date(f.minDate))+' - '+fmtDate(new Date(f.maxDate))+'</div></div><button class="btn-remove" onclick="removeSalesFile(\''+f.filename.replace(/'/g,"\\'")+'\')">Hapus</button></div>';
    });
  }
  if(DB.aging.filename) html += '<div class="data-status-row"><div><div class="name">AR Aging &mdash; '+esc(DB.aging.filename)+'</div><div class="meta">'+DB.aging.rows.length+' invoice &middot; diupload '+fmtDate(new Date(DB.aging.uploadedAt))+'</div></div><button class="btn-remove" onclick="removeAging()">Hapus</button></div>';
  if(DB.payment.filename) html += '<div class="data-status-row"><div><div class="name">AR Payment &mdash; '+esc(DB.payment.filename)+'</div><div class="meta">'+DB.payment.rows.length+' baris &middot; diupload '+fmtDate(new Date(DB.payment.uploadedAt))+'</div></div><button class="btn-remove" onclick="removePayment()">Hapus</button></div>';
  if(DB.stock.filename) html += '<div class="data-status-row"><div><div class="name">Stock &mdash; '+esc(DB.stock.filename)+'</div><div class="meta">'+DB.stock.rows.length+' baris &middot; diupload '+fmtDate(new Date(DB.stock.uploadedAt))+'</div></div><button class="btn-remove" onclick="removeStock()">Hapus</button></div>';
  if(DB.targets.filename) html += '<div class="data-status-row"><div><div class="name">Target Penjualan &mdash; '+esc(DB.targets.filename)+'</div><div class="meta">'+DB.targets.rows.length+' salesman &middot; diupload '+fmtDate(new Date(DB.targets.uploadedAt))+'</div></div><button class="btn-remove" onclick="removeTargets()">Hapus</button></div>';
  if(DB.customers.filename) html += '<div class="data-status-row"><div><div class="name">Database Pelanggan (Alamat) &mdash; '+esc(DB.customers.filename)+'</div><div class="meta">'+DB.customers.rows.length+' pelanggan &middot; diupload '+fmtDate(new Date(DB.customers.uploadedAt))+'</div></div><button class="btn-remove" onclick="removeCustomers()">Hapus</button></div>';
  if(!html) html = '<p style="color:var(--ink-faint); font-size:12.5px;">Belum ada data tersimpan.</p>';
  el.innerHTML = html;
  const anyData = DB.salesFiles.length || DB.aging.filename || DB.payment.filename || DB.stock.filename || DB.targets.filename || DB.customers.filename;
  const resetWrap = document.getElementById('resetAllWrap');
  if(resetWrap) resetWrap.style.display = anyData ? 'block' : 'none';
  const exportWrap = document.getElementById('exportWithDataWrap');
  if(exportWrap) exportWrap.style.display = anyData ? 'block' : 'none';
}

async function removeAllData(){
  if(!confirm('Hapus SEMUA data (Sales, AR Aging, AR Payment, Stock, Target, Database Pelanggan) yang tersimpan di dashboard ini? Tindakan ini tidak bisa dibatalkan.')) return;
  for(const f of DB.salesFiles.slice()){ try{ await storage.delete('sales_data:'+f.filename); }catch(e){} }
  DB.salesFiles = []; DB.sales = [];
  try{ await storage.set('sales_files_index', JSON.stringify([])); }catch(e){}
  try{ await storage.delete('aging_data'); }catch(e){}
  try{ await storage.delete('payment_data'); }catch(e){}
  try{ await storage.delete('stock_data'); }catch(e){}
  try{ await storage.delete('target_data'); }catch(e){}
  try{ await storage.delete('customer_data'); }catch(e){}
  DB.aging = {rows:[], filename:null, uploadedAt:null};
  DB.payment = {rows:[], filename:null, uploadedAt:null};
  DB.stock = {rows:[], filename:null, uploadedAt:null};
  DB.targets = {rows:[], filename:null, uploadedAt:null};
  DB.customers = {rows:[], filename:null, uploadedAt:null};
  invalidateCabangLookup(); invalidateSalesmanCabangCache(); invalidateCustomerAddressCache();
  renderAll();
  renderDataStatusList();
}

/* =========================================================================
   MODAL / DRAWER CONTROLS
========================================================================= */
function openUploadModal(){ document.getElementById('uploadModal').classList.add('open'); renderDataStatusList(); }
function closeUploadModal(){ document.getElementById('uploadModal').classList.remove('open'); }
function openDrawer(title, bodyHtml, wide){
  document.getElementById('drawerTitle').innerHTML = title;
  document.getElementById('drawerBody').innerHTML = bodyHtml;
  document.getElementById('drawerBox').classList.toggle('wide', !!wide);
  document.getElementById('invoiceDrawer').classList.add('open');
}
function closeDrawer(){ document.getElementById('invoiceDrawer').classList.remove('open'); }

document.getElementById('openUploadBtn').addEventListener('click', openUploadModal);
document.getElementById('uploadModal').addEventListener('click', e=>{ if(e.target.id==='uploadModal') closeUploadModal(); });
document.getElementById('invoiceDrawer').addEventListener('click', e=>{ if(e.target.id==='invoiceDrawer') closeDrawer(); });
const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
dropzone.addEventListener('click', ()=>fileInput.click());
fileInput.addEventListener('change', e=>{ handleFiles(e.target.files); fileInput.value=''; });
dropzone.addEventListener('dragover', e=>{ e.preventDefault(); dropzone.classList.add('drag'); });
dropzone.addEventListener('dragleave', ()=> dropzone.classList.remove('drag'));
dropzone.addEventListener('drop', e=>{
  e.preventDefault(); dropzone.classList.remove('drag');
  handleFiles(e.dataTransfer.files);
});

/* =========================================================================
   TABS
========================================================================= */
document.querySelectorAll('.tab-btn').forEach(btn=>{
  btn.addEventListener('click', ()=>{
    document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
    document.querySelectorAll('.view').forEach(v=>v.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById('view-'+btn.dataset.tab).classList.add('active');
  });
});

/* =========================================================================
   FOCUS-PRESERVING REBUILD
   Filter bars are rebuilt via innerHTML on every keystroke (since typing in
   the search box triggers a full re-render). That recreates the DOM nodes,
   including the input being typed in, which steals focus after every
   character. This wraps a rebuild so the previously-focused field (by id)
   regains focus and cursor position immediately after.
========================================================================= */
function rebuildPreservingFocus(container, rebuildFn){
  const active = document.activeElement;
  let restoreId = null, restorePos = null;
  if(active && container && container.contains(active) && (active.tagName==='INPUT' || active.tagName==='SELECT')){
    restoreId = active.id;
    if(active.selectionStart!=null && active.selectionEnd!=null) restorePos = [active.selectionStart, active.selectionEnd];
  }
  rebuildFn();
  if(restoreId){
    const el = document.getElementById(restoreId);
    if(el){
      el.focus();
      if(restorePos && el.setSelectionRange){ try{ el.setSelectionRange(restorePos[0], restorePos[1]); }catch(e){} }
    }
  }
}

/* =========================================================================
   GENERIC TABLE / SORT / PAGINATION HELPER
========================================================================= */
function makeTableController(opts){
  // opts: {getRows, columns:[{key,label,render,sortable}], pageSize, tbodyId, theadId, pagerId, infoId, rowKey}
  let sortKey = opts.defaultSort ? opts.defaultSort.key : null;
  let sortDir = opts.defaultSort ? opts.defaultSort.dir : 'desc';
  let page = 1;
  function render(){
    let rows = opts.getRows();
    if(sortKey){
      rows = rows.slice().sort((a,b)=>{
        let av=a[sortKey], bv=b[sortKey];
        if(av instanceof Date) av = av.getTime(); if(bv instanceof Date) bv = bv.getTime();
        if(typeof av==='string') av = av.toLowerCase(); if(typeof bv==='string') bv = bv.toLowerCase();
        if(av===bv) return 0;
        if(av===null||av===undefined) return 1;
        if(bv===null||bv===undefined) return -1;
        return sortDir==='asc' ? (av>bv?1:-1) : (av<bv?1:-1);
      });
    }
    const total = rows.length;
    const pageSize = opts.pageSize || total || 1;
    const maxPage = Math.max(1, Math.ceil(total/pageSize));
    if(page>maxPage) page = maxPage;
    const start = (page-1)*pageSize;
    const pageRows = opts.pageSize ? rows.slice(start, start+pageSize) : rows;

    const thead = document.getElementById(opts.theadId);
    if(thead){
      thead.innerHTML = '<tr>'+opts.columns.map(c=>{
        let cls='';
        if(c.key===sortKey) cls = sortDir==='asc'?'sort-asc':'sort-desc';
        return '<th class="'+cls+'" data-key="'+c.key+'">'+c.label+'</th>';
      }).join('')+'</tr>';
      thead.querySelectorAll('th').forEach(th=>{
        th.addEventListener('click', ()=>{
          const k = th.dataset.key;
          if(sortKey===k) sortDir = sortDir==='asc'?'desc':'asc'; else { sortKey=k; sortDir='desc'; }
          render();
        });
      });
    }
    const tbody = document.getElementById(opts.tbodyId);
    if(pageRows.length===0){
      tbody.innerHTML = '<tr><td colspan="'+opts.columns.length+'" style="text-align:center; color:var(--ink-faint); padding:24px;">Tidak ada data yang cocok dengan filter</td></tr>';
    } else {
      tbody.innerHTML = pageRows.map(row=>opts.renderRow(row)).join('');
      if(opts.afterRender) opts.afterRender(tbody, pageRows);
    }
    if(opts.infoId){
      const infoEl = document.getElementById(opts.infoId);
      if(infoEl) infoEl.textContent = total===0 ? '0 data' : ('Menampilkan '+(start+1)+'-'+Math.min(start+pageSize,total)+' dari '+total+' data');
    }
    if(opts.pagerId){
      const pagerEl = document.getElementById(opts.pagerId);
      if(pagerEl){
        pagerEl.innerHTML = '<button '+(page<=1?'disabled':'')+' id="'+opts.pagerId+'-prev">&laquo; Prev</button><span style="align-self:center; font-size:12px; padding:0 4px;">'+page+' / '+maxPage+'</span><button '+(page>=maxPage?'disabled':'')+' id="'+opts.pagerId+'-next">Next &raquo;</button>';
        const prevBtn = document.getElementById(opts.pagerId+'-prev');
        const nextBtn = document.getElementById(opts.pagerId+'-next');
        if(prevBtn) prevBtn.addEventListener('click', ()=>{ page--; render(); });
        if(nextBtn) nextBtn.addEventListener('click', ()=>{ page++; render(); });
      }
    }
  }
  return { render, resetPage:()=>{page=1;} };
}

/* =========================================================================
   CHART.JS SHARED CONFIG
========================================================================= */
const HAS_CHART = typeof Chart !== 'undefined';
const HAS_XLSX = typeof XLSX !== 'undefined';

/* =========================================================================
   EXPORT (Excel via SheetJS, PDF via browser print)
========================================================================= */
function exportToExcel(filename, rows, columns){
  if(!HAS_XLSX){ alert('Library Excel (XLSX) gagal dimuat dari CDN — coba refresh halaman lalu ulangi.'); return; }
  if(!rows.length){ alert('Tidak ada data untuk diexport pada filter saat ini.'); return; }
  const data = rows.map(r=>{
    const obj = {};
    columns.forEach(c=>{ obj[c.label] = typeof c.value==='function' ? c.value(r) : r[c.key]; });
    return obj;
  });
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Data');
  XLSX.writeFile(wb, filename);
}

function exportToPDF(title, filterSummaryLines, rows, columns){
  const header = document.getElementById('print-header');
  if(!rows.length){ alert('Tidak ada data untuk diexport pada filter saat ini.'); return; }
  const theadHtml = '<tr>'+columns.map(c=>'<th>'+esc(c.label)+'</th>').join('')+'</tr>';
  const tbodyHtml = rows.map(r=>'<tr>'+columns.map(c=>{
    const v = typeof c.value==='function' ? c.value(r) : r[c.key];
    return '<td>'+esc(v===null||v===undefined?'':v)+'</td>';
  }).join('')+'</tr>').join('');
  header.innerHTML = '<h1>'+esc(title)+'</h1>' +
    '<p>'+filterSummaryLines.map(esc).join(' &middot; ')+'</p>' +
    '<p class="print-generated">Dicetak dari Dashboard Distribusi &middot; '+fmtDate(new Date())+' &middot; '+rows.length+' baris data</p>' +
    '<table class="print-table"><thead>'+theadHtml+'</thead><tbody>'+tbodyHtml+'</tbody></table>';
  window.print();
}

function activeFilterSummary(stateObj, labels){
  // stateObj: filter state object; labels: {key: 'Label'} map for keys to describe when not 'all'/empty
  const parts = [];
  Object.entries(labels).forEach(([key,label])=>{
    const val = stateObj[key];
    if(val && val!=='all' && val!==''){ parts.push(label+': '+val); }
  });
  return parts.length ? parts : ['Tanpa filter (semua data)'];
}
if(HAS_CHART){
  Chart.defaults.font.family = "'IBM Plex Sans', sans-serif";
  Chart.defaults.font.size = 11.5;
  Chart.defaults.color = '#516560';
} else {
  console.error('Chart.js gagal dimuat dari CDN — grafik tidak akan tampil, tapi tabel & fitur lain tetap berjalan.');
}
const CHARTS = {};
function destroyChart(id){ if(CHARTS[id]){ CHARTS[id].destroy(); delete CHARTS[id]; } }
function safeNewChart(canvasId, config){
  if(!HAS_CHART){
    const el = document.getElementById(canvasId);
    if(el && el.parentElement) el.parentElement.innerHTML = '<p style="color:var(--ink-faint); font-size:12.5px; text-align:center; padding:30px 0;">Grafik tidak tersedia (library chart gagal dimuat)</p>';
    return null;
  }
  return new Chart(document.getElementById(canvasId), config);
}
const PALETTE = ['#0F6E64','#D98F4E','#4C8C7A','#CBA43A','#7A9E97','#C94F3F','#3D8FA6','#8A6D1E','#5B7C76','#B5714A'];

/* =========================================================================
   SALES MODULE
========================================================================= */
const salesFilterState = { cabang:'all', tahun:'all', bulan:'all', salesman:'all', principal:'all', outlet:'all', q:'', valueBasis:'dpp' };
function salesVal(r){ return salesFilterState.valueBasis==='netto' ? r.netto : r.dpp; }
function salesValLabel(){ return salesFilterState.valueBasis==='netto' ? 'Netto (termasuk PPN)' : 'DPP (sebelum PPN)'; }

function salesFilteredRows(excludeYear){
  return DB.sales.filter(r=>{
    if(salesFilterState.cabang!=='all' && r.cabang!==salesFilterState.cabang) return false;
    if(!excludeYear && salesFilterState.tahun!=='all' && String(r.year)!==salesFilterState.tahun) return false;
    if(salesFilterState.bulan!=='all' && String(r.month)!==salesFilterState.bulan) return false;
    if(salesFilterState.salesman!=='all' && r.salesman!==salesFilterState.salesman) return false;
    if(salesFilterState.principal!=='all' && r.principal!==salesFilterState.principal) return false;
    if(salesFilterState.outlet!=='all' && r.tipeOutlet!==salesFilterState.outlet) return false;
    if(salesFilterState.q){
      const q = salesFilterState.q.toLowerCase();
      if(!(r.namaProduk.toLowerCase().includes(q) || r.namaProdukGroup.toLowerCase().includes(q) || r.namaPelanggan.toLowerCase().includes(q) || r.invoice.toLowerCase().includes(q))) return false;
    }
    return true;
  });
}

function uniqueVals(arr, key){ return [...new Set(arr.map(r=>r[key]).filter(Boolean))].sort(); }

function salesFilteredExcept(exceptKey){
  return DB.sales.filter(r=>{
    if(exceptKey!=='cabang' && salesFilterState.cabang!=='all' && r.cabang!==salesFilterState.cabang) return false;
    if(exceptKey!=='tahun' && salesFilterState.tahun!=='all' && String(r.year)!==salesFilterState.tahun) return false;
    if(exceptKey!=='bulan' && salesFilterState.bulan!=='all' && String(r.month)!==salesFilterState.bulan) return false;
    if(exceptKey!=='salesman' && salesFilterState.salesman!=='all' && r.salesman!==salesFilterState.salesman) return false;
    if(exceptKey!=='principal' && salesFilterState.principal!=='all' && r.principal!==salesFilterState.principal) return false;
    if(exceptKey!=='outlet' && salesFilterState.outlet!=='all' && r.tipeOutlet!==salesFilterState.outlet) return false;
    return true;
  });
}

function renderSalesFilterBar(){
  const bar = document.getElementById('sales-filterbar');
  // Each dimension's option list is scoped by every OTHER active filter (cascading/faceted filter),
  // so e.g. picking Cabang=JAKARTA narrows Salesman/Principal/etc to only what exists in JAKARTA.
  const cabangs = uniqueVals(salesFilteredExcept('cabang'),'cabang');
  const tahuns = [...new Set(salesFilteredExcept('tahun').map(r=>r.year))].sort((a,b)=>b-a);
  const bulans = [...new Set(salesFilteredExcept('bulan').map(r=>r.month))].sort((a,b)=>a-b);
  const bulanNames = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Ags','Sep','Okt','Nov','Des'];
  const salesmen = uniqueVals(salesFilteredExcept('salesman'),'salesman');
  const principals = uniqueVals(salesFilteredExcept('principal'),'principal');
  const outlets = uniqueVals(salesFilteredExcept('outlet'),'tipeOutlet');

  // If a previously chosen value no longer exists given the other active filters, reset it to "Semua"
  // so the combination never silently produces zero rows.
  if(salesFilterState.cabang!=='all' && !cabangs.includes(salesFilterState.cabang)) salesFilterState.cabang='all';
  if(salesFilterState.tahun!=='all' && !tahuns.map(String).includes(salesFilterState.tahun)) salesFilterState.tahun='all';
  if(salesFilterState.bulan!=='all' && !bulans.map(String).includes(salesFilterState.bulan)) salesFilterState.bulan='all';
  if(salesFilterState.salesman!=='all' && !salesmen.includes(salesFilterState.salesman)) salesFilterState.salesman='all';
  if(salesFilterState.principal!=='all' && !principals.includes(salesFilterState.principal)) salesFilterState.principal='all';
  if(salesFilterState.outlet!=='all' && !outlets.includes(salesFilterState.outlet)) salesFilterState.outlet='all';

  function opts(list, cur, labelFn){ return '<option value="all">Semua</option>' + list.map(v=>'<option value="'+esc(v)+'" '+(cur===String(v)?'selected':'')+'>'+esc(labelFn?labelFn(v):v)+'</option>').join(''); }

  bar.innerHTML = `
    <div class="filter-field"><label>Basis Nilai</label><select id="f-sales-basis"><option value="dpp" ${salesFilterState.valueBasis==='dpp'?'selected':''}>DPP (sebelum PPN)</option><option value="netto" ${salesFilterState.valueBasis==='netto'?'selected':''}>Netto (termasuk PPN)</option></select></div>
    <div class="filter-field"><label>Cabang</label><select id="f-sales-cabang">${opts(cabangs, salesFilterState.cabang)}</select></div>
    <div class="filter-field"><label>Tahun</label><select id="f-sales-tahun">${opts(tahuns, salesFilterState.tahun)}</select></div>
    <div class="filter-field"><label>Bulan</label><select id="f-sales-bulan">${opts(bulans, salesFilterState.bulan, m=>bulanNames[m-1])}</select></div>
    <div class="filter-field"><label>Salesman</label><select id="f-sales-salesman">${opts(salesmen, salesFilterState.salesman)}</select></div>
    <div class="filter-field"><label>Principal</label><select id="f-sales-principal">${opts(principals, salesFilterState.principal)}</select></div>
    <div class="filter-field"><label>Tipe Outlet</label><select id="f-sales-outlet">${opts(outlets, salesFilterState.outlet)}</select></div>
    <div class="filter-field"><label>Cari</label><input type="search" id="f-sales-q" placeholder="Produk / pelanggan / invoice" value="${esc(salesFilterState.q)}"></div>
    <button class="btn-reset" id="f-sales-reset">Reset Filter</button>
  `;
  document.getElementById('f-sales-basis').addEventListener('change', e=>{ salesFilterState.valueBasis=e.target.value; renderSales(); });
  const ids = {cabang:'f-sales-cabang', tahun:'f-sales-tahun', bulan:'f-sales-bulan', salesman:'f-sales-salesman', principal:'f-sales-principal', outlet:'f-sales-outlet'};
  Object.entries(ids).forEach(([k,id])=>{
    document.getElementById(id).addEventListener('change', e=>{ salesFilterState[k]=e.target.value; renderSales(); });
  });
  document.getElementById('f-sales-q').addEventListener('input', e=>{ salesFilterState.q = e.target.value; renderSales(); });
  document.getElementById('f-sales-reset').addEventListener('click', ()=>{
    Object.assign(salesFilterState, {cabang:'all',tahun:'all',bulan:'all',salesman:'all',principal:'all',outlet:'all',q:'',valueBasis:'dpp'});
    renderSales();
  });
}

let salesTableCtrl = null;
function renderSalesKPI(rows){
  const totalNetto = rows.reduce((s,r)=>s+r.netto,0);
  const totalDpp = rows.reduce((s,r)=>s+r.dpp,0);
  const totalVal = salesFilterState.valueBasis==='netto' ? totalNetto : totalDpp;
  const totalQty = rows.reduce((s,r)=>s+r.qty,0);
  const invoices = new Set(rows.map(r=>r.invoice)).size;
  const pelanggan = new Set(rows.map(r=>r.namaPelanggan)).size;

  // YoY comparison for currently filtered scope (compares same set of months across years if a single year not selected, else compares selected year vs previous)
  let yoyHtml = '';
  const years = [...new Set(rows.map(r=>r.year))];
  if(salesFilterState.tahun!=='all'){
    const y = parseInt(salesFilterState.tahun,10);
    const prevRows = salesFilteredRows(true).filter(r=>r.year===y-1 && (salesFilterState.bulan==='all' || r.month===parseInt(salesFilterState.bulan,10)));
    if(prevRows.length){
      const prevVal = prevRows.reduce((s,r)=>s+salesVal(r),0);
      const pct = prevVal ? ((totalVal - prevVal)/prevVal*100) : null;
      yoyHtml = pct!==null ? (pct>=0? '+' : '') + pct.toFixed(1)+'% vs '+(y-1) : '';
    }
  }

  const kpi = document.getElementById('sales-kpi');
  kpi.innerHTML = `
    <div class="kpi-card good"><div class="label">Nilai Penjualan (${salesValLabel()})</div><div class="value">${fmtIDR(totalVal)}</div><div class="sub ${yoyHtml.startsWith('+')?'up':yoyHtml.startsWith('-')?'down':''}">${yoyHtml || 'Sesuai basis nilai terpilih'}</div></div>
    <div class="kpi-card"><div class="label">DPP (sebelum PPN)</div><div class="value">${fmtIDR(totalDpp)}</div><div class="sub">Dasar Pengenaan Pajak</div></div>
    <div class="kpi-card"><div class="label">Netto (termasuk PPN)</div><div class="value">${fmtIDR(totalNetto)}</div><div class="sub">Jumlah Netto Faktur</div></div>
    <div class="kpi-card"><div class="label">Total Qty Terjual</div><div class="value">${fmtNumber(totalQty)}</div><div class="sub">${invoices} invoice &middot; ${pelanggan} pelanggan</div></div>
  `;
}

function renderSalesTrendChart(rows){
  const years = [...new Set(rows.map(r=>r.year))].sort();
  const bulanNames = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Ags','Sep','Okt','Nov','Des'];
  const byYearMonth = {};
  rows.forEach(r=>{
    byYearMonth[r.year] = byYearMonth[r.year] || Array(12).fill(0);
    byYearMonth[r.year][r.month-1] += salesVal(r);
  });
  const datasets = years.map((y,i)=>({
    label: String(y),
    data: byYearMonth[y],
    borderColor: PALETTE[i%PALETTE.length],
    backgroundColor: PALETTE[i%PALETTE.length]+'22',
    tension:0.3, fill:false, borderWidth:2.5, pointRadius:3
  }));
  destroyChart('trend');
  CHARTS['trend'] = safeNewChart('chart-sales-trend', {
    type:'line',
    data:{ labels:bulanNames, datasets },
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{legend:{position:'top', labels:{boxWidth:12, usePointStyle:true}}, tooltip:{callbacks:{label:c=>c.dataset.label+': '+fmtIDR(c.parsed.y)}}},
      scales:{ y:{ticks:{callback:v=>(v/1e6).toFixed(0)+'jt'}}, x:{grid:{display:false}} }
    }
  });
}

function renderSalesProdukChart(rows){
  const map = {};
  rows.forEach(r=>{ map[r.namaProdukGroup] = (map[r.namaProdukGroup]||0) + salesVal(r); });
  const top = Object.entries(map).sort((a,b)=>b[1]-a[1]).slice(0,10);
  const tagEl = document.getElementById('produk-chart-tag'); if(tagEl) tagEl.textContent = 'berdasarkan '+salesValLabel();
  destroyChart('produk');
  CHARTS['produk'] = safeNewChart('chart-top-produk', {
    type:'bar',
    data:{ labels: top.map(t=>t[0].length>28?t[0].slice(0,28)+'…':t[0]), datasets:[{ data: top.map(t=>t[1]), backgroundColor: PALETTE[0], borderRadius:5 }] },
    options:{ indexAxis:'y', responsive:true, maintainAspectRatio:false, plugins:{legend:{display:false}, tooltip:{callbacks:{label:c=>fmtIDR(c.parsed.x)}}}, scales:{x:{ticks:{callback:v=>(v/1e6).toFixed(0)+'jt'}}} }
  });
}
function renderSalesPrincipalChart(rows){
  const map = {};
  rows.forEach(r=>{ map[r.principal] = (map[r.principal]||0) + salesVal(r); });
  const top = Object.entries(map).sort((a,b)=>b[1]-a[1]).slice(0,8);
  const tagEl2 = document.getElementById('principal-chart-tag'); if(tagEl2) tagEl2.textContent = 'berdasarkan '+salesValLabel();
  destroyChart('principal');
  CHARTS['principal'] = safeNewChart('chart-principal', {
    type:'doughnut',
    data:{ labels: top.map(t=>t[0]), datasets:[{ data: top.map(t=>t[1]), backgroundColor: PALETTE, borderWidth:2, borderColor:'#fff' }] },
    options:{ responsive:true, maintainAspectRatio:false, plugins:{legend:{position:'right', labels:{boxWidth:11, font:{size:11}}}, tooltip:{callbacks:{label:c=>c.label+': '+fmtIDR(c.parsed)}}} }
  });
}

function renderSalesPerformerTable(rows){
  const map = {};
  rows.forEach(r=>{
    if(!map[r.salesman]) map[r.salesman] = {salesman:r.salesman, netto:0, dpp:0, qty:0, invoices:new Set(), pelanggan:new Set()};
    const m = map[r.salesman]; m.dpp+=r.dpp; m.netto+=r.netto; m.qty+=r.qty; m.invoices.add(r.invoice); m.pelanggan.add(r.namaPelanggan);
  });
  const sortKey = salesFilterState.valueBasis==='netto' ? 'netto' : 'dpp';
  const list = Object.values(map).map(m=>({...m, invoiceCount:m.invoices.size, pelangganCount:m.pelanggan.size})).sort((a,b)=>b[sortKey]-a[sortKey]);
  const thead = document.querySelector('#table-sales-performer thead');
  const tbody = document.querySelector('#table-sales-performer tbody');
  thead.innerHTML = '<tr><th>Salesman</th><th>DPP</th><th>Netto (+PPN)</th><th>Qty</th><th>Invoice</th><th>Pelanggan</th></tr>';
  tbody.innerHTML = list.length ? list.map(m=>`<tr><td class="text">${esc(m.salesman)}</td><td>${fmtIDR(m.dpp)}</td><td>${fmtIDR(m.netto)}</td><td>${fmtNumber(m.qty)}</td><td>${m.invoiceCount}</td><td>${m.pelangganCount}</td></tr>`).join('') :
    '<tr><td colspan="6" style="text-align:center; color:var(--ink-faint); padding:20px;">Tidak ada data</td></tr>';
}

function renderSalesDetailTable(rows){
  if(!salesTableCtrl){
    salesTableCtrl = makeTableController({
      getRows: ()=>salesFilteredRows(),
      columns:[
        {key:'tanggal',label:'Tanggal'},{key:'invoice',label:'Invoice'},{key:'namaPelanggan',label:'Pelanggan'},
        {key:'namaProduk',label:'Produk'},{key:'salesman',label:'Sales'},{key:'qty',label:'Qty'},
        {key:'dpp',label:'DPP'},{key:'netto',label:'Netto'}
      ],
      renderRow: r=>`<tr><td>${fmtDate(r.tanggal)}</td><td>${esc(r.invoice)}</td><td class="text customer-link" data-customer="${esc(r.namaPelanggan)}">${esc(r.namaPelanggan)}</td><td class="text" title="${esc(r.namaProduk)}">${esc(r.namaProdukGroup)}</td><td class="text">${esc(r.salesman)}</td><td>${fmtNumber(r.qty)}</td><td>${fmtIDR(r.dpp)}</td><td>${fmtIDR(r.netto)}</td></tr>`,
      afterRender: (tbody)=> bindCustomerLinks(tbody),
      pageSize:25, tbodyId:'table-sales-detail-tbody', theadId:'table-sales-detail-thead', pagerId:'sales-pager', infoId:'sales-pager-info',
      defaultSort:{key:'tanggal', dir:'desc'}
    });
    document.querySelector('#table-sales-detail').innerHTML = '<thead id="table-sales-detail-thead"></thead><tbody id="table-sales-detail-tbody"></tbody>';
  }
  salesTableCtrl.resetPage();
  salesTableCtrl.render();
  document.getElementById('sales-table-count').textContent = salesFilteredRows().length + ' baris transaksi';
}

/* ---- Rincian Produk per Salesman ---- */
const salesmanDetailState = { salesman: null };

function salesmanScopeRows(){
  // same as salesFilteredRows() but ignores the Salesman filter itself,
  // so the "sold vs not sold" universe reflects the other active filters only
  return DB.sales.filter(r=>{
    if(salesFilterState.cabang!=='all' && r.cabang!==salesFilterState.cabang) return false;
    if(salesFilterState.tahun!=='all' && String(r.year)!==salesFilterState.tahun) return false;
    if(salesFilterState.bulan!=='all' && String(r.month)!==salesFilterState.bulan) return false;
    if(salesFilterState.principal!=='all' && r.principal!==salesFilterState.principal) return false;
    if(salesFilterState.outlet!=='all' && r.tipeOutlet!==salesFilterState.outlet) return false;
    return true;
  });
}

function renderSalesmanDetail(){
  const scopeRows = salesmanScopeRows();
  const salesmen = uniqueVals(scopeRows,'salesman');
  const sel = document.getElementById('f-salesman-detail');
  if(!salesmanDetailState.salesman || !salesmen.includes(salesmanDetailState.salesman)){
    salesmanDetailState.salesman = salesFilterState.salesman!=='all' && salesmen.includes(salesFilterState.salesman) ? salesFilterState.salesman : (salesmen[0]||null);
  }
  sel.innerHTML = salesmen.map(s=>`<option value="${esc(s)}" ${s===salesmanDetailState.salesman?'selected':''}>${esc(s)}</option>`).join('');
  sel.onchange = ()=>{ salesmanDetailState.salesman = sel.value; renderSalesmanDetail(); };

  document.getElementById('salesman-detail-scope-note').textContent = 'Cakupan mengikuti filter Cabang/Tahun/Bulan/Principal/Tipe Outlet di atas (filter Salesman tidak berlaku di sini).';

  const soldTbody = document.querySelector('#table-salesman-sold tbody');
  const soldThead = document.querySelector('#table-salesman-sold thead');
  const notSoldTbody = document.querySelector('#table-salesman-notsold tbody');
  const notSoldThead = document.querySelector('#table-salesman-notsold thead');

  if(!salesmanDetailState.salesman){
    soldThead.innerHTML=''; notSoldThead.innerHTML='';
    soldTbody.innerHTML = '<tr><td style="text-align:center; color:var(--ink-faint); padding:16px;">Tidak ada data</td></tr>';
    notSoldTbody.innerHTML = '';
    document.getElementById('salesman-sold-count').textContent = '';
    document.getElementById('salesman-notsold-count').textContent = '';
    return;
  }

  // aggregate all products in scope (grouped by canonical product name) and which salesmen sold each
  const allProducts = {}; // namaProdukGroup -> {namaProdukGroup, principal}
  const soldByThis = {}; // namaProdukGroup -> {namaProdukGroup, principal, qty, value}
  const soldByOthersValue = {}; // namaProdukGroup -> value sold by OTHER salesmen (opportunity size)
  scopeRows.forEach(r=>{
    if(!allProducts[r.namaProdukGroup]) allProducts[r.namaProdukGroup] = {namaProdukGroup:r.namaProdukGroup, principal:r.principal};
    if(r.salesman===salesmanDetailState.salesman){
      if(!soldByThis[r.namaProdukGroup]) soldByThis[r.namaProdukGroup] = {namaProdukGroup:r.namaProdukGroup, principal:r.principal, qty:0, value:0};
      soldByThis[r.namaProdukGroup].qty += r.qty; soldByThis[r.namaProdukGroup].value += salesVal(r);
    } else {
      soldByOthersValue[r.namaProdukGroup] = (soldByOthersValue[r.namaProdukGroup]||0) + salesVal(r);
    }
  });
  const soldList = Object.values(soldByThis).sort((a,b)=>b.value-a.value);
  const notSoldList = Object.values(allProducts).filter(p=>!soldByThis[p.namaProdukGroup]).map(p=>({...p, othersValue: soldByOthersValue[p.namaProdukGroup]||0})).sort((a,b)=>b.othersValue-a.othersValue);

  document.getElementById('salesman-sold-count').textContent = soldList.length+' produk';
  document.getElementById('salesman-notsold-count').textContent = notSoldList.length+' produk';

  soldThead.innerHTML = '<tr><th>Produk</th><th>Qty</th><th>Nilai</th></tr>';
  soldTbody.innerHTML = soldList.length ? soldList.map(p=>`<tr><td class="text">${esc(p.namaProdukGroup)}</td><td>${fmtNumber(p.qty)}</td><td>${fmtIDR(p.value)}</td></tr>`).join('') :
    '<tr><td colspan="3" style="text-align:center; color:var(--ink-faint); padding:16px;">Salesman ini belum menjual produk apapun pada cakupan filter saat ini</td></tr>';

  notSoldThead.innerHTML = '<tr><th>Produk</th><th>Nilai Terjual (Sales Lain)</th></tr>';
  notSoldTbody.innerHTML = notSoldList.length ? notSoldList.map(p=>`<tr><td class="text">${esc(p.namaProdukGroup)}</td><td>${p.othersValue>0?fmtIDR(p.othersValue):'<span style="color:var(--ink-faint);">belum ada yang jual</span>'}</td></tr>`).join('') :
    '<tr><td colspan="2" style="text-align:center; color:var(--ink-faint); padding:16px;">Semua produk pada cakupan ini sudah dijual oleh salesman ini</td></tr>';
}

function renderSales(){
  const hasData = DB.sales.length>0;
  document.getElementById('sales-empty').style.display = hasData? 'none':'block';
  document.getElementById('sales-content').style.display = hasData? 'block':'none';
  document.getElementById('statusDotSales').classList.toggle('on', hasData);
  if(!hasData) return;
  rebuildPreservingFocus(document.getElementById('sales-filterbar'), renderSalesFilterBar);
  const rows = salesFilteredRows();
  renderSalesKPI(rows);
  renderSalesTrendChart(salesFilteredRows(true).filter(r=>salesFilterState.bulan==='all'||r.month===parseInt(salesFilterState.bulan,10)));
  renderSalesProdukChart(rows);
  renderSalesPrincipalChart(rows);
  renderSalesPerformerTable(rows);
  renderSalesmanDetail();
  renderSalesDetailTable(rows);
  const dates = DB.sales.map(r=>r.tanggal);
  document.getElementById('sales-period-hint').textContent = 'Data tersedia: ' + fmtDate(new Date(Math.min(...dates))) + ' s/d ' + fmtDate(new Date(Math.max(...dates)));
}

/* =========================================================================
   AR AGING MODULE
========================================================================= */
const agingFilterState = { cabang:'all', bucket:'all', q:'' };
let agingTableCtrl = null;

// AR Aging invoice numbers encode the branch as a 3-letter code right after
// "INV" (e.g. INVJKT... -> JAKARTA). This is exact, unlike name-matching.
const CABANG_CODE_MAP = {
  ACH:'ACEH', BDG:'BANDUNG', BJM:'BANJARMASIN', BKS:'BEKASI', BLI:'BALI', CMS:'CIAMIS',
  CRB:'CIREBON', CRS:'CIRACAS', JBI:'JAMBI', JKT:'JAKARTA', KND:'KENDARI', LPG:'LAMPUNG',
  MDN:'MEDAN', MDO:'MANADO', MKS:'MAKASSAR', MLG:'MALANG', PKU:'PEKANBARU', PLB:'PALEMBANG',
  PLK:'PALANGKARAYA', PLU:'PALU', PTK:'PONTIANAK', PWT:'PURWOKERTO', SBY:'SURABAYA', SLO:'SOLO',
  SMD:'SAMARINDA', SMG:'SEMARANG', SMI:'SUKABUMI', SRG:'SERANG'
};
function cabangFromInvoiceCode(invoice){
  if(!invoice) return null;
  const m = String(invoice).toUpperCase().match(/INV([A-Z]{3})/);
  return m ? (CABANG_CODE_MAP[m[1]] || null) : null;
}

// Fallback for cases where the invoice doesn't carry a recognizable branch code:
// match the customer name against Sales / AR Payment cabang records instead.
let _cabangLookupCache = null;
function getCustomerCabangMap(){
  if(_cabangLookupCache) return _cabangLookupCache;
  const map = new Map();
  DB.payment.rows.forEach(r=>{ if(r.namaPelanggan && r.cabang && !map.has(normName(r.namaPelanggan))) map.set(normName(r.namaPelanggan), r.cabang); });
  DB.sales.forEach(r=>{ if(r.namaPelanggan && r.cabang) map.set(normName(r.namaPelanggan), r.cabang); }); // Sales takes priority
  _cabangLookupCache = map;
  return map;
}
function invalidateCabangLookup(){ _cabangLookupCache = null; }
function agingRowCabang(r){
  return cabangFromInvoiceCode(r.invoice) || getCustomerCabangMap().get(normName(r.namaPelanggan)) || null;
}

function agingFilteredRows(){
  return DB.aging.rows.filter(r=>{
    if(agingFilterState.cabang!=='all' && agingRowCabang(r)!==agingFilterState.cabang) return false;
    if(agingFilterState.bucket!=='all' && r.bucket!==agingFilterState.bucket) return false;
    if(agingFilterState.q){
      const q = agingFilterState.q.toLowerCase();
      if(!(r.namaPelanggan.toLowerCase().includes(q) || r.invoice.toLowerCase().includes(q))) return false;
    }
    return true;
  });
}

const BUCKET_ORDER = ['Belum Jatuh Tempo','1-14 Hari','15-30 Hari','31-60 Hari','61-90 Hari','> 90 Hari','Saldo Kredit','Tanpa Tgl Jatuh Tempo'];
const BUCKET_RISK = {'Belum Jatuh Tempo':0,'1-14 Hari':1,'15-30 Hari':2,'31-60 Hari':3,'61-90 Hari':4,'> 90 Hari':4,'Saldo Kredit':0,'Tanpa Tgl Jatuh Tempo':0};
function bucketBadgeClass(b){ const lv = BUCKET_RISK[b]; return 'b'+(lv===undefined?0:lv); }

function agingRowsExcept(exceptKey){
  return DB.aging.rows.filter(r=>{
    if(exceptKey!=='cabang' && agingFilterState.cabang!=='all' && agingRowCabang(r)!==agingFilterState.cabang) return false;
    if(exceptKey!=='bucket' && agingFilterState.bucket!=='all' && r.bucket!==agingFilterState.bucket) return false;
    return true;
  });
}

function renderAgingFilterBar(){
  const bar = document.getElementById('aging-filterbar');
  const buckets = BUCKET_ORDER.filter(b=>agingRowsExcept('bucket').some(r=>r.bucket===b));
  const cabangs = [...new Set(agingRowsExcept('cabang').map(r=>agingRowCabang(r)).filter(Boolean))].sort();
  if(agingFilterState.cabang!=='all' && !cabangs.includes(agingFilterState.cabang)) agingFilterState.cabang='all';
  if(agingFilterState.bucket!=='all' && !buckets.includes(agingFilterState.bucket)) agingFilterState.bucket='all';
  bar.innerHTML = `
    <div class="filter-field"><label>Cabang</label><select id="f-aging-cabang"><option value="all">Semua</option>${cabangs.map(c=>'<option value="'+esc(c)+'" '+(agingFilterState.cabang===c?'selected':'')+'>'+esc(c)+'</option>').join('')}</select></div>
    <div class="filter-field"><label>Status Umur</label><select id="f-aging-bucket"><option value="all">Semua</option>${buckets.map(b=>'<option value="'+esc(b)+'" '+(agingFilterState.bucket===b?'selected':'')+'>'+esc(b)+'</option>').join('')}</select></div>
    <div class="filter-field"><label>Cari</label><input type="search" id="f-aging-q" placeholder="Pelanggan / invoice" value="${esc(agingFilterState.q)}"></div>
    <button class="btn-reset" id="f-aging-reset">Reset Filter</button>
  `;
  document.getElementById('f-aging-cabang').addEventListener('change', e=>{ agingFilterState.cabang=e.target.value; renderAging(); });
  document.getElementById('f-aging-bucket').addEventListener('change', e=>{ agingFilterState.bucket=e.target.value; renderAging(); });
  document.getElementById('f-aging-q').addEventListener('input', e=>{ agingFilterState.q=e.target.value; renderAging(); });
  document.getElementById('f-aging-reset').addEventListener('click', ()=>{ Object.assign(agingFilterState,{cabang:'all',bucket:'all',q:''}); renderAging(); });
  const unrecognized = DB.aging.rows.filter(r=>!agingRowCabang(r)).length;
  if(unrecognized>0){
    const note = document.createElement('p');
    note.className = 'note'; note.style.width='100%';
    note.textContent = unrecognized+' dari '+DB.aging.rows.length+' invoice tidak dikenali kode cabangnya (format nomor invoice tidak cocok / kode belum terdaftar) dan tidak akan muncul saat filter Cabang dipilih.';
    bar.appendChild(note);
  }
}

function renderAgingAlert(){
  const rows = DB.aging.rows;
  const urgent = rows.filter(r=>BUCKET_RISK[r.bucket]>=3);
  const slot = document.getElementById('aging-alert-slot');
  if(urgent.length===0){ slot.innerHTML=''; return; }
  const totalUrgent = urgent.reduce((s,r)=>s+r.totalSaldo,0);
  slot.innerHTML = `<div class="alert-banner"><span class="icon">&#9888;</span><div><strong>${urgent.length} invoice</strong> sudah menunggak lebih dari 30 hari, total ${fmtIDR(totalUrgent)} &mdash; segera tindak lanjut penagihan.</div></div>`;
}

function renderAgingKPI(rows){
  const piutang = rows.filter(r=>r.totalSaldo>0);
  const kredit = rows.filter(r=>r.totalSaldo<0);
  const totalSaldo = piutang.reduce((s,r)=>s+r.totalSaldo,0);
  const totalOverdue = piutang.filter(r=>BUCKET_RISK[r.bucket]>=1 && r.bucket!=='Belum Jatuh Tempo').reduce((s,r)=>s+r.totalSaldo,0);
  const over60 = piutang.filter(r=>BUCKET_RISK[r.bucket]>=3).reduce((s,r)=>s+r.totalSaldo,0);
  const totalKredit = Math.abs(kredit.reduce((s,r)=>s+r.totalSaldo,0));
  document.getElementById('aging-kpi').innerHTML = `
    <div class="kpi-card"><div class="label">Total Piutang</div><div class="value">${fmtIDR(totalSaldo)}</div><div class="sub">${piutang.length} invoice</div></div>
    <div class="kpi-card warn"><div class="label">Sudah Jatuh Tempo</div><div class="value">${fmtIDR(totalOverdue)}</div><div class="sub">${piutang.filter(r=>BUCKET_RISK[r.bucket]>=1 && r.bucket!=='Belum Jatuh Tempo').length} invoice, &gt; 14 hari</div></div>
    <div class="kpi-card danger"><div class="label">&gt; 60 Hari (Kritis)</div><div class="value">${fmtIDR(over60)}</div><div class="sub">${piutang.filter(r=>BUCKET_RISK[r.bucket]>=3).length} invoice</div></div>
    <div class="kpi-card"><div class="label">Saldo Kredit</div><div class="value">${fmtIDR(totalKredit)}</div><div class="sub">${kredit.length} invoice lebih bayar</div></div>
  `;
}

function renderAgingBucketChart(rows){
  const order = ['Belum Jatuh Tempo','1-14 Hari','15-30 Hari','31-60 Hari','61-90 Hari','> 90 Hari'];
  const colors = ['#4C8C7A','#7A9E97','#CBA43A','#E08A3C','#C94F3F','#7A2E2E'];
  const data = order.map(b=> rows.filter(r=>r.bucket===b).reduce((s,r)=>s+r.totalSaldo,0));
  destroyChart('agingBucket');
  CHARTS['agingBucket'] = safeNewChart('chart-aging-bucket', {
    type:'bar',
    data:{ labels:order, datasets:[{ data, backgroundColor:colors, borderRadius:5 }] },
    options:{ responsive:true, maintainAspectRatio:false, plugins:{legend:{display:false}, tooltip:{callbacks:{label:c=>fmtIDR(c.parsed.y)}}}, scales:{y:{ticks:{callback:v=>(v/1e6).toFixed(0)+'jt'}}} }
  });
}

function renderTopDebtorTable(rows){
  const map = {};
  rows.filter(r=>r.totalSaldo>0).forEach(r=>{
    map[r.namaPelanggan] = map[r.namaPelanggan] || {namaPelanggan:r.namaPelanggan, total:0, maxRisk:0, count:0};
    map[r.namaPelanggan].total += r.totalSaldo; map[r.namaPelanggan].count++;
    map[r.namaPelanggan].maxRisk = Math.max(map[r.namaPelanggan].maxRisk, BUCKET_RISK[r.bucket]||0);
  });
  const top = Object.values(map).sort((a,b)=>b.total-a.total).slice(0,10);
  const thead = document.querySelector('#table-top-debtor thead');
  const tbody = document.querySelector('#table-top-debtor tbody');
  thead.innerHTML = '<tr><th>Pelanggan</th><th>Piutang</th><th>Invoice</th><th>Risiko</th></tr>';
  tbody.innerHTML = top.length ? top.map(m=>`<tr><td class="text customer-link" data-customer="${esc(m.namaPelanggan)}">${esc(m.namaPelanggan)}</td><td>${fmtIDR(m.total)}</td><td>${m.count}</td><td><span class="badge b${m.maxRisk}">${['Aman','1-14 hr','15-30 hr','31-60 hr','&gt;60 hr'][m.maxRisk]}</span></td></tr>`).join('') :
    '<tr><td colspan="4" style="text-align:center; color:var(--ink-faint); padding:20px;">Tidak ada data</td></tr>';
  bindCustomerLinks(tbody);
}

function invoiceDetailHtml(invoiceKey){
  const lines = DB.sales.filter(r=>r.invoiceKey===invoiceKey);
  if(!lines.length){
    return '<p style="color:var(--ink-soft); font-size:13px;">Detail produk belum tersedia &mdash; data Sales untuk bulan invoice ini belum diupload.</p>';
  }
  const total = lines.reduce((s,r)=>s+r.dpp,0);
  return `<div class="table-wrap"><table><thead><tr><th>Produk</th><th>Principal</th><th>Qty</th><th>DPP</th></tr></thead><tbody>
    ${lines.map(l=>`<tr><td class="text" title="${esc(l.namaProduk)}">${esc(l.namaProdukGroup)}</td><td class="text">${esc(l.principal)}</td><td>${fmtNumber(l.qty)}</td><td>${fmtIDR(l.dpp)}</td></tr>`).join('')}
  </tbody></table></div><p style="margin-top:12px; font-weight:600;">Total DPP: ${fmtIDR(total)}</p>`;
}

function renderAgingDetailTable(){
  if(!agingTableCtrl){
    agingTableCtrl = makeTableController({
      getRows: ()=>agingFilteredRows(),
      columns:[
        {key:'invoice',label:'Invoice'},{key:'namaPelanggan',label:'Pelanggan'},{key:'tglInvoice',label:'Tgl Invoice'},
        {key:'tglJatuhTempo',label:'Jatuh Tempo'},{key:'agingDays',label:'Umur (hari)'},{key:'bucket',label:'Status'},{key:'totalSaldo',label:'Saldo'}
      ],
      renderRow: r=>`<tr>
        <td class="link-cell" data-invoice="${esc(r.invoiceKey)}">${esc(r.invoice)}</td>
        <td class="text customer-link" data-customer="${esc(r.namaPelanggan)}">${esc(r.namaPelanggan)}</td>
        <td>${fmtDate(r.tglInvoice)}</td>
        <td>${fmtDate(r.tglJatuhTempo)}</td>
        <td>${r.agingDays===null?'-':r.agingDays}</td>
        <td><span class="badge ${bucketBadgeClass(r.bucket)}">${esc(r.bucket)}</span></td>
        <td>${fmtIDR(r.totalSaldo)}</td>
      </tr>`,
      afterRender: (tbody)=>{
        tbody.querySelectorAll('.link-cell').forEach(el=>{
          el.addEventListener('click', ()=>{
            const inv = el.dataset.invoice;
            const row = DB.aging.rows.find(r=>r.invoiceKey===inv);
            openDrawer('Detail Invoice '+esc(row.invoice), invoiceDetailHtml(inv));
          });
        });
        bindCustomerLinks(tbody);
      },
      pageSize:20, tbodyId:'table-aging-detail-tbody', theadId:'table-aging-detail-thead', pagerId:'aging-pager', infoId:'aging-pager-info',
      defaultSort:{key:'totalSaldo', dir:'desc'}
    });
    document.querySelector('#table-aging-detail').innerHTML = '<thead id="table-aging-detail-thead"></thead><tbody id="table-aging-detail-tbody"></tbody>';
  }
  agingTableCtrl.resetPage();
  agingTableCtrl.render();
}

function renderAging(){
  const hasData = DB.aging.rows.length>0;
  document.getElementById('aging-empty').style.display = hasData? 'none':'block';
  document.getElementById('aging-content').style.display = hasData? 'block':'none';
  document.getElementById('statusDotAging').classList.toggle('on', hasData);
  if(!hasData) return;
  rebuildPreservingFocus(document.getElementById('aging-filterbar'), renderAgingFilterBar);
  renderAgingAlert();
  const rows = agingFilteredRows();
  renderAgingKPI(rows);
  renderAgingBucketChart(rows);
  renderTopDebtorTable(rows);
  renderAgingDetailTable();
  document.getElementById('aging-asof-hint').textContent = 'Dihitung ulang otomatis dari tanggal jatuh tempo, per hari ini ('+fmtDate(TODAY)+')';
}

/* =========================================================================
   AR PAYMENT MODULE
========================================================================= */
const paymentFilterState = { sales:'all', status:'all', q:'' };
let paymentTableCtrl = null;

function paymentFilteredRows(){
  return DB.payment.rows.filter(r=>{
    if(paymentFilterState.sales!=='all' && r.sales!==paymentFilterState.sales) return false;
    if(paymentFilterState.status==='lunas' && !r.lunas) return false;
    if(paymentFilterState.status==='sisa' && r.lunas) return false;
    if(paymentFilterState.q){
      const q = paymentFilterState.q.toLowerCase();
      if(!(r.namaPelanggan.toLowerCase().includes(q) || r.invoice.toLowerCase().includes(q))) return false;
    }
    return true;
  });
}

function paymentRowsExcept(exceptKey){
  return DB.payment.rows.filter(r=>{
    if(exceptKey!=='sales' && paymentFilterState.sales!=='all' && r.sales!==paymentFilterState.sales) return false;
    if(exceptKey!=='status' && paymentFilterState.status==='lunas' && !r.lunas) return false;
    if(exceptKey!=='status' && paymentFilterState.status==='sisa' && r.lunas) return false;
    return true;
  });
}

function renderPaymentFilterBar(){
  const bar = document.getElementById('payment-filterbar');
  const salesList = uniqueVals(paymentRowsExcept('sales'),'sales');
  const statusRows = paymentRowsExcept('status');
  const hasLunas = statusRows.some(r=>r.lunas), hasSisa = statusRows.some(r=>!r.lunas);
  if(paymentFilterState.sales!=='all' && !salesList.includes(paymentFilterState.sales)) paymentFilterState.sales='all';
  if(paymentFilterState.status==='lunas' && !hasLunas) paymentFilterState.status='all';
  if(paymentFilterState.status==='sisa' && !hasSisa) paymentFilterState.status='all';
  bar.innerHTML = `
    <div class="filter-field"><label>Sales</label><select id="f-pay-sales"><option value="all">Semua</option>${salesList.map(s=>'<option value="'+esc(s)+'" '+(paymentFilterState.sales===s?'selected':'')+'>'+esc(s)+'</option>').join('')}</select></div>
    <div class="filter-field"><label>Status</label><select id="f-pay-status"><option value="all">Semua</option>${hasLunas?'<option value="lunas" '+(paymentFilterState.status==='lunas'?'selected':'')+'>Lunas</option>':''}${hasSisa?'<option value="sisa" '+(paymentFilterState.status==='sisa'?'selected':'')+'>Masih Ada Sisa</option>':''}</select></div>
    <div class="filter-field"><label>Cari</label><input type="search" id="f-pay-q" placeholder="Pelanggan / invoice" value="${esc(paymentFilterState.q)}"></div>
    <button class="btn-reset" id="f-pay-reset">Reset Filter</button>
  `;
  document.getElementById('f-pay-sales').addEventListener('change', e=>{ paymentFilterState.sales=e.target.value; renderPayment(); });
  document.getElementById('f-pay-status').addEventListener('change', e=>{ paymentFilterState.status=e.target.value; renderPayment(); });
  document.getElementById('f-pay-q').addEventListener('input', e=>{ paymentFilterState.q=e.target.value; renderPayment(); });
  document.getElementById('f-pay-reset').addEventListener('click', ()=>{ Object.assign(paymentFilterState,{sales:'all',status:'all',q:''}); renderPayment(); });
}

function renderPaymentKPI(rows){
  const totalDiterima = rows.reduce((s,r)=>s+r.pembTunai+r.pembDeposit+r.pembCek,0);
  const totalTunai = rows.reduce((s,r)=>s+r.pembTunai,0);
  const totalDeposit = rows.reduce((s,r)=>s+r.pembDeposit,0);
  const belumLunas = rows.filter(r=>!r.lunas);
  document.getElementById('payment-kpi').innerHTML = `
    <div class="kpi-card good"><div class="label">Total Diterima</div><div class="value">${fmtIDR(totalDiterima)}</div><div class="sub">${rows.length} invoice</div></div>
    <div class="kpi-card"><div class="label">Tunai</div><div class="value">${fmtIDR(totalTunai)}</div></div>
    <div class="kpi-card"><div class="label">Deposit/Transfer</div><div class="value">${fmtIDR(totalDeposit)}</div></div>
    <div class="kpi-card warn"><div class="label">Masih Ada Sisa</div><div class="value">${belumLunas.length}</div><div class="sub">${fmtIDR(belumLunas.reduce((s,r)=>s+r.sisaPemb,0))}</div></div>
  `;
}

function renderPaymentMethodChart(rows){
  const tunai = rows.reduce((s,r)=>s+r.pembTunai,0);
  const deposit = rows.reduce((s,r)=>s+r.pembDeposit,0);
  const cek = rows.reduce((s,r)=>s+r.pembCek,0);
  destroyChart('paymethod');
  CHARTS['paymethod'] = safeNewChart('chart-payment-method', {
    type:'doughnut',
    data:{ labels:['Tunai','Deposit/Transfer','Cek'], datasets:[{ data:[tunai,deposit,cek], backgroundColor:[PALETTE[0],PALETTE[1],PALETTE[2]], borderWidth:2, borderColor:'#fff' }] },
    options:{ responsive:true, maintainAspectRatio:false, plugins:{legend:{position:'bottom'}, tooltip:{callbacks:{label:c=>c.label+': '+fmtIDR(c.parsed)}}} }
  });
}
function renderPaymentSalesChart(rows){
  const map = {};
  rows.forEach(r=>{ map[r.sales] = (map[r.sales]||0) + r.pembTunai+r.pembDeposit+r.pembCek; });
  const top = Object.entries(map).sort((a,b)=>b[1]-a[1]).slice(0,10);
  destroyChart('paysales');
  CHARTS['paysales'] = safeNewChart('chart-payment-sales', {
    type:'bar',
    data:{ labels: top.map(t=>t[0]), datasets:[{ data: top.map(t=>t[1]), backgroundColor: PALETTE[0], borderRadius:5 }] },
    options:{ indexAxis:'y', responsive:true, maintainAspectRatio:false, plugins:{legend:{display:false}, tooltip:{callbacks:{label:c=>fmtIDR(c.parsed.x)}}}, scales:{x:{ticks:{callback:v=>(v/1e6).toFixed(0)+'jt'}}} }
  });
}

function renderPaymentTable(){
  if(!paymentTableCtrl){
    paymentTableCtrl = makeTableController({
      getRows: ()=>paymentFilteredRows(),
      columns:[
        {key:'invoice',label:'Invoice'},{key:'namaPelanggan',label:'Pelanggan'},{key:'sales',label:'Sales'},
        {key:'nilaiInvoice',label:'Nilai Invoice'},{key:'pembTunai',label:'Tunai'},{key:'pembDeposit',label:'Deposit'},
        {key:'tglPembTerakhir',label:'Tgl Bayar Terakhir'},{key:'sisaPemb',label:'Sisa'},{key:'lunas',label:'Status'}
      ],
      renderRow: r=>`<tr><td>${esc(r.invoice)}</td><td class="text customer-link" data-customer="${esc(r.namaPelanggan)}">${esc(r.namaPelanggan)}</td><td class="text">${esc(r.sales)}</td><td>${fmtIDR(r.nilaiInvoice)}</td><td>${fmtIDR(r.pembTunai)}</td><td>${fmtIDR(r.pembDeposit)}</td><td>${fmtDate(r.tglPembTerakhir)}</td><td>${fmtIDR(r.sisaPemb)}</td><td>${r.lunas?'<span class="badge b0">Lunas</span>':'<span class="badge b2">Sisa</span>'}</td></tr>`,
      afterRender: (tbody)=> bindCustomerLinks(tbody),
      pageSize:25, tbodyId:'table-payment-detail-tbody', theadId:'table-payment-detail-thead', pagerId:'payment-pager', infoId:'payment-pager-info',
      defaultSort:{key:'tglPembTerakhir', dir:'desc'}
    });
    document.querySelector('#table-payment-detail').innerHTML = '<thead id="table-payment-detail-thead"></thead><tbody id="table-payment-detail-tbody"></tbody>';
  }
  paymentTableCtrl.resetPage();
  paymentTableCtrl.render();
  document.getElementById('payment-table-count').textContent = paymentFilteredRows().length + ' baris pembayaran';
}

function renderPayment(){
  const hasData = DB.payment.rows.length>0;
  document.getElementById('payment-empty').style.display = hasData? 'none':'block';
  document.getElementById('payment-content').style.display = hasData? 'block':'none';
  document.getElementById('statusDotPayment').classList.toggle('on', hasData);
  if(!hasData) return;
  rebuildPreservingFocus(document.getElementById('payment-filterbar'), renderPaymentFilterBar);
  const rows = paymentFilteredRows();
  renderPaymentKPI(rows);
  renderPaymentMethodChart(rows);
  renderPaymentSalesChart(rows);
  renderPaymentTable();
  document.getElementById('payment-asof-hint').textContent = 'Diupload '+ (DB.payment.uploadedAt? fmtDate(new Date(DB.payment.uploadedAt)) : '-');
}

/* =========================================================================
   STOCK MODULE
========================================================================= */
const stockFilterState = { lokasi:'all', status:'all', q:'' };
let stockTableCtrl = null;

function computeSalesVelocity(){
  // qty terjual per produk dalam 90 hari terakhir data sales, dan sepanjang periode data
  const map = {};
  if(!DB.sales.length) return map;
  const maxDate = new Date(Math.max(...DB.sales.map(r=>r.tanggal.getTime())));
  const minDate = new Date(Math.min(...DB.sales.map(r=>r.tanggal.getTime())));
  const windowStart = new Date(Math.max(minDate.getTime(), (()=>{ const d=new Date(maxDate); d.setDate(d.getDate()-90); return d; })().getTime()));
  DB.sales.forEach(r=>{
    if(!map[r.idProdukKey]) map[r.idProdukKey] = {idProduk:r.idProdukKey, namaProduk:r.namaProduk, principal:r.principal, totalQty:0, totalValue:0, qty90:0, lastSale:null};
    const m = map[r.idProdukKey];
    m.totalQty += r.qty; m.totalValue += r.dpp;
    if(r.tanggal >= windowStart) m.qty90 += r.qty;
    if(!m.lastSale || r.tanggal > m.lastSale) m.lastSale = r.tanggal;
  });
  // windowDays reflects the actual span of data available (up to 90 days) so velocity isn't
  // under-estimated when fewer than 3 months of sales data have been uploaded yet.
  const windowDays = Math.max(1, daysBetween(maxDate, windowStart) + 1);
  Object.values(map).forEach(m=>{ m.dailyVelocity = m.qty90 / windowDays; });
  return map;
}

function stockAggregatedByProduct(){
  const agg = {};
  DB.stock.rows.forEach(r=>{
    if(!agg[r.idProdukKey]) agg[r.idProdukKey] = {idProduk:r.idProdukKey, namaProduk:r.namaProduk, satuan:r.satuan, stokAkhir:0, locations:new Set()};
    agg[r.idProdukKey].stokAkhir += r.stokAkhir;
    if(r.stokAkhir>0) agg[r.idProdukKey].locations.add(r.locationId);
  });
  return agg;
}

function stockFilteredRows(){
  return DB.stock.rows.filter(r=>{
    if(stockFilterState.lokasi!=='all' && r.locationId!==stockFilterState.lokasi) return false;
    if(stockFilterState.status==='ada' && r.stokAkhir<=0) return false;
    if(stockFilterState.status==='habis' && r.stokAkhir>0) return false;
    if(stockFilterState.q){
      const q = stockFilterState.q.toLowerCase();
      if(!(r.namaProduk.toLowerCase().includes(q) || r.idProduk.toLowerCase().includes(q))) return false;
    }
    return true;
  });
}

function stockRowsExcept(exceptKey){
  return DB.stock.rows.filter(r=>{
    if(exceptKey!=='lokasi' && stockFilterState.lokasi!=='all' && r.locationId!==stockFilterState.lokasi) return false;
    if(exceptKey!=='status' && stockFilterState.status==='ada' && r.stokAkhir<=0) return false;
    if(exceptKey!=='status' && stockFilterState.status==='habis' && r.stokAkhir>0) return false;
    return true;
  });
}

function renderStockFilterBar(){
  const bar = document.getElementById('stock-filterbar');
  const lokasis = uniqueVals(stockRowsExcept('lokasi'),'locationId');
  const statusRows = stockRowsExcept('status');
  const hasAda = statusRows.some(r=>r.stokAkhir>0), hasHabis = statusRows.some(r=>r.stokAkhir<=0);
  if(stockFilterState.lokasi!=='all' && !lokasis.includes(stockFilterState.lokasi)) stockFilterState.lokasi='all';
  if(stockFilterState.status==='ada' && !hasAda) stockFilterState.status='all';
  if(stockFilterState.status==='habis' && !hasHabis) stockFilterState.status='all';
  bar.innerHTML = `
    <div class="filter-field"><label>Lokasi/Gudang</label><select id="f-stock-lokasi"><option value="all">Semua</option>${lokasis.map(l=>'<option value="'+esc(l)+'" '+(stockFilterState.lokasi===l?'selected':'')+'>'+esc(l)+'</option>').join('')}</select></div>
    <div class="filter-field"><label>Status</label><select id="f-stock-status"><option value="all">Semua</option>${hasAda?'<option value="ada" '+(stockFilterState.status==='ada'?'selected':'')+'>Ada Stock</option>':''}${hasHabis?'<option value="habis" '+(stockFilterState.status==='habis'?'selected':'')+'>Stock Habis</option>':''}</select></div>
    <div class="filter-field"><label>Cari Produk</label><input type="search" id="f-stock-q" placeholder="Nama / kode produk" value="${esc(stockFilterState.q)}"></div>
    <button class="btn-reset" id="f-stock-reset">Reset Filter</button>
  `;
  document.getElementById('f-stock-lokasi').addEventListener('change', e=>{ stockFilterState.lokasi=e.target.value; renderStock(); });
  document.getElementById('f-stock-status').addEventListener('change', e=>{ stockFilterState.status=e.target.value; renderStock(); });
  document.getElementById('f-stock-q').addEventListener('input', e=>{ stockFilterState.q=e.target.value; renderStock(); });
  document.getElementById('f-stock-reset').addEventListener('click', ()=>{ Object.assign(stockFilterState,{lokasi:'all',status:'all',q:''}); renderStock(); });
}

function renderStockAlert(reorderList, expiredCount){
  const slot = document.getElementById('stock-alert-slot');
  let html = '';
  if(reorderList.length){
    html += `<div class="alert-banner"><span class="icon">&#128680;</span><div><strong>${reorderList.length} produk</strong> diperkirakan stock habis dalam &le; 14 hari &mdash; segera ajukan PO ke principal.</div></div>`;
  }
  if(expiredCount>0){
    html += `<div class="alert-banner warn"><span class="icon">&#9888;</span><div><strong>${expiredCount} lot</strong> sudah melewati tanggal kadaluwarsa namun masih tercatat ada stock.</div></div>`;
  }
  slot.innerHTML = html;
}

function renderStockKPI(agg, velocity){
  const totalSKU = Object.keys(agg).length;
  const totalUnit = Object.values(agg).reduce((s,a)=>s+a.stokAkhir,0);
  const totalValue = Object.values(agg).reduce((s,a)=>{
    const v = velocity[a.idProduk];
    const avgPrice = v && v.totalQty>0 ? v.totalValue/v.totalQty : 0;
    return s + a.stokAkhir*avgPrice;
  },0);
  const zeroStock = Object.values(agg).filter(a=>a.stokAkhir<=0).length;
  document.getElementById('stock-kpi').innerHTML = `
    <div class="kpi-card"><div class="label">Total SKU</div><div class="value">${totalSKU}</div><div class="sub">${zeroStock} SKU stock habis</div></div>
    <div class="kpi-card"><div class="label">Total Unit Stock</div><div class="value">${fmtNumber(totalUnit)}</div></div>
    <div class="kpi-card good"><div class="label">Estimasi Nilai Stock</div><div class="value">${fmtIDR(totalValue)}</div><div class="sub">berdasar harga rata-rata penjualan</div></div>
  `;
}

function renderReorderTable(agg, velocity){
  const list = Object.values(agg).filter(a=>a.stokAkhir>0).map(a=>{
    const v = velocity[a.idProduk];
    const dv = v ? v.dailyVelocity : 0;
    const daysLeft = dv>0 ? a.stokAkhir/dv : null;
    return {...a, dailyVelocity:dv, daysLeft};
  }).filter(a=>a.daysLeft!==null && a.daysLeft<=30).sort((a,b)=>a.daysLeft-b.daysLeft);
  const thead = document.querySelector('#table-reorder thead');
  const tbody = document.querySelector('#table-reorder tbody');
  thead.innerHTML = '<tr><th>Produk</th><th>Stock Saat Ini</th><th>Kecepatan Jual/Hari</th><th>Estimasi Habis</th><th>Urgensi</th></tr>';
  tbody.innerHTML = list.length ? list.map(a=>{
    const lvl = a.daysLeft<=7?4:a.daysLeft<=14?3:2;
    const pct = Math.min(100, (a.daysLeft/30)*100);
    return `<tr><td class="text">${esc(a.namaProduk)}</td><td>${fmtNumber(a.stokAkhir)}</td><td>${a.dailyVelocity.toFixed(2)}</td>
      <td><div class="risk-wrap"><span class="risk-label">${a.daysLeft.toFixed(0)} hari</span><div class="risk-bar" style="opacity:1;"><div class="marker" style="left:${pct}%;"></div></div></div></td>
      <td><span class="badge b${lvl}">${lvl===4?'Sangat Mendesak':lvl===3?'Mendesak':'Perhatikan'}</span></td></tr>`;
  }).join('') : '<tr><td colspan="5" style="text-align:center; color:var(--ink-faint); padding:20px;">Tidak ada produk dengan risiko stock habis dalam 30 hari, atau data Sales belum diupload untuk menghitung kecepatan jual.</td></tr>';
  return list.filter(a=>a.daysLeft<=14);
}

function renderSlowMovingTable(agg, velocity){
  const withStock = Object.values(agg).filter(a=>a.stokAkhir>0).map(a=>{
    const v = velocity[a.idProduk];
    return {...a, totalQty: v?v.totalQty:0, lastSale: v?v.lastSale:null};
  });
  if(!withStock.length || !DB.sales.length){
    document.querySelector('#table-slowmoving thead').innerHTML='';
    document.querySelector('#table-slowmoving tbody').innerHTML = '<tr><td style="text-align:center; color:var(--ink-faint); padding:20px;">Butuh data Sales untuk analisis ini</td></tr>';
    return;
  }
  const sorted = withStock.slice().sort((a,b)=>a.totalQty-b.totalQty);
  const cutoff = Math.max(1, Math.ceil(sorted.length*0.2));
  const slow = sorted.slice(0, cutoff);
  document.querySelector('#table-slowmoving thead').innerHTML = '<tr><th>Produk</th><th>Total Qty Terjual (Semua Periode Data)</th><th>Stock</th><th>Terakhir Terjual</th></tr>';
  document.querySelector('#table-slowmoving tbody').innerHTML = slow.map(a=>`<tr><td class="text">${esc(a.namaProduk)}</td><td>${fmtNumber(a.totalQty)}</td><td>${fmtNumber(a.stokAkhir)}</td><td>${a.lastSale?fmtDate(a.lastSale):'Tidak pernah terjual'}</td></tr>`).join('');
}

function renderFastMovingTable(agg, velocity){
  const list = Object.values(velocity).sort((a,b)=>b.totalQty-a.totalQty).slice(0,15).map(v=>{
    const stock = agg[v.idProduk] ? agg[v.idProduk].stokAkhir : 0;
    return {...v, stock};
  });
  const tagEl = document.getElementById('fastmoving-tag');
  if(tagEl && DB.sales.length){
    const minD = new Date(Math.min(...DB.sales.map(r=>r.tanggal.getTime())));
    const maxD = new Date(Math.max(...DB.sales.map(r=>r.tanggal.getTime())));
    tagEl.textContent = 'total qty terjual ' + fmtDate(minD) + ' – ' + fmtDate(maxD) + ' (bukan rata-rata)';
  }
  document.querySelector('#table-fastmoving thead').innerHTML = '<tr><th>Produk</th><th>Total Qty Terjual (Semua Periode Data)</th><th>Stock Sisa</th></tr>';
  document.querySelector('#table-fastmoving tbody').innerHTML = list.length ? list.map(v=>`<tr><td class="text">${esc(v.namaProduk)}</td><td>${fmtNumber(v.totalQty)}</td><td>${v.stock<=0?'<span class="badge b3">Habis</span>':fmtNumber(v.stock)}</td></tr>`).join('') :
    '<tr><td colspan="3" style="text-align:center; color:var(--ink-faint); padding:20px;">Butuh data Sales</td></tr>';
}

function renderExpiryTable(){
  const rows = DB.stock.rows.filter(r=>r.stokAkhir>0 && r.tglKadaluwarsa);
  const withDays = rows.map(r=>({...r, daysToExpiry: daysBetween(r.tglKadaluwarsa, TODAY)})).filter(r=>r.daysToExpiry<=90).sort((a,b)=>a.daysToExpiry-b.daysToExpiry);
  document.querySelector('#table-expiry thead').innerHTML = '<tr><th>Produk</th><th>Lot</th><th>Lokasi</th><th>Stock</th><th>Kadaluwarsa</th><th>Status</th></tr>';
  document.querySelector('#table-expiry tbody').innerHTML = withDays.length ? withDays.map(r=>{
    const expired = r.daysToExpiry<0;
    return `<tr><td class="text">${esc(r.namaProduk)}</td><td>${esc(r.idLot)}</td><td>${esc(r.locationId)}</td><td>${fmtNumber(r.stokAkhir)}</td><td>${fmtDate(r.tglKadaluwarsa)}</td><td>${expired?'<span class="badge b4">Sudah Expired</span>':'<span class="badge b'+(r.daysToExpiry<=30?3:2)+'">'+r.daysToExpiry+' hari lagi</span>'}</td></tr>`;
  }).join('') : '<tr><td colspan="6" style="text-align:center; color:var(--ink-faint); padding:20px;">Tidak ada produk mendekati/sudah kadaluwarsa dengan stock &gt; 0</td></tr>';
  return withDays.filter(r=>r.daysToExpiry<0).length;
}

function renderStockDetailTable(){
  if(!stockTableCtrl){
    stockTableCtrl = makeTableController({
      getRows: ()=>stockFilteredRows(),
      columns:[
        {key:'namaProduk',label:'Produk'},{key:'locationId',label:'Lokasi'},{key:'idLot',label:'Lot'},
        {key:'tglKadaluwarsa',label:'Kadaluwarsa'},{key:'satuan',label:'Satuan'},{key:'stokAkhir',label:'Stock'}
      ],
      renderRow: r=>`<tr><td class="text">${esc(r.namaProduk)}</td><td>${esc(r.locationId)}</td><td>${esc(r.idLot)}</td><td>${fmtDate(r.tglKadaluwarsa)}</td><td class="text">${esc(r.satuan)}</td><td>${fmtNumber(r.stokAkhir)}</td></tr>`,
      pageSize:25, tbodyId:'table-stock-detail-tbody', theadId:'table-stock-detail-thead', pagerId:'stock-pager', infoId:'stock-pager-info',
      defaultSort:{key:'stokAkhir', dir:'desc'}
    });
    document.querySelector('#table-stock-detail').innerHTML = '<thead id="table-stock-detail-thead"></thead><tbody id="table-stock-detail-tbody"></tbody>';
  }
  stockTableCtrl.resetPage();
  stockTableCtrl.render();
  document.getElementById('stock-table-count').textContent = stockFilteredRows().length + ' baris lot';
}

function renderStock(){
  const hasData = DB.stock.rows.length>0;
  document.getElementById('stock-empty').style.display = hasData? 'none':'block';
  document.getElementById('stock-content').style.display = hasData? 'block':'none';
  document.getElementById('statusDotStock').classList.toggle('on', hasData);
  if(!hasData) return;
  rebuildPreservingFocus(document.getElementById('stock-filterbar'), renderStockFilterBar);
  const agg = stockAggregatedByProduct();
  const velocity = computeSalesVelocity();
  renderStockKPI(agg, velocity);
  const reorderUrgent = renderReorderTable(agg, velocity);
  renderSlowMovingTable(agg, velocity);
  renderFastMovingTable(agg, velocity);
  const expiredCount = renderExpiryTable();
  renderStockAlert(reorderUrgent, expiredCount);
  renderStockDetailTable();
  document.getElementById('stock-asof-hint').textContent = 'Diupload '+ (DB.stock.uploadedAt? fmtDate(new Date(DB.stock.uploadedAt)) : '-');
}

/* =========================================================================
   CUSTOMER 360
========================================================================= */
function normName(s){ return (s||'').trim().toLowerCase(); }

function getAllCustomerNames(){
  const set = new Map(); // normalized -> display name
  DB.sales.forEach(r=>{ if(r.namaPelanggan) set.set(normName(r.namaPelanggan), r.namaPelanggan); });
  DB.aging.rows.forEach(r=>{ if(r.namaPelanggan) set.set(normName(r.namaPelanggan), r.namaPelanggan); });
  DB.payment.rows.forEach(r=>{ if(r.namaPelanggan) set.set(normName(r.namaPelanggan), r.namaPelanggan); });
  return set;
}

function openCustomer360(namaPelanggan){
  const key = normName(namaPelanggan);
  const salesRows = DB.sales.filter(r=>normName(r.namaPelanggan)===key);
  const agingRows = DB.aging.rows.filter(r=>normName(r.namaPelanggan)===key);
  const paymentRows = DB.payment.rows.filter(r=>normName(r.namaPelanggan)===key);
  const displayName = (salesRows[0]||agingRows[0]||paymentRows[0]||{}).namaPelanggan || namaPelanggan;

  const totalBeli = salesRows.reduce((s,r)=>s+r.dpp,0);
  const totalQty = salesRows.reduce((s,r)=>s+r.qty,0);
  const invoiceCount = new Set(salesRows.map(r=>r.invoice)).size;
  const lastBeli = salesRows.length ? new Date(Math.max(...salesRows.map(r=>r.tanggal.getTime()))) : null;
  const tipeOutlet = salesRows[0] ? salesRows[0].tipeOutlet : null;
  const cabang = salesRows[0] ? salesRows[0].cabang : (agingRows[0]?'':'');

  const totalPiutang = agingRows.filter(r=>r.totalSaldo>0).reduce((s,r)=>s+r.totalSaldo,0);
  const maxRisk = agingRows.reduce((m,r)=>Math.max(m, BUCKET_RISK[r.bucket]||0), 0);

  const totalDibayar = paymentRows.reduce((s,r)=>s+r.pembTunai+r.pembDeposit+r.pembCek,0);
  const lastBayar = paymentRows.length ? new Date(Math.max(...paymentRows.filter(r=>r.tglPembTerakhir).map(r=>r.tglPembTerakhir.getTime()))) : null;

  // top produk dibeli pelanggan ini
  const prodMap = {};
  salesRows.forEach(r=>{ prodMap[r.namaProdukGroup] = (prodMap[r.namaProdukGroup]||0) + r.dpp; });
  const topProduk = Object.entries(prodMap).sort((a,b)=>b[1]-a[1]).slice(0,5);

  const html = `
    <div class="c360-head">
      <div>
        <h3>${esc(displayName)}</h3>
        <div class="meta">${cabang?esc(cabang)+' &middot; ':''}${tipeOutlet?'Tipe: '+esc(tipeOutlet):'Data pelanggan gabungan'}</div>
      </div>
      ${maxRisk>=3?'<span class="badge b'+maxRisk+'">Piutang Berisiko</span>':''}
    </div>
    <div class="kpi-row" style="margin-bottom:22px;">
      <div class="kpi-card"><div class="label">Total Pembelian</div><div class="value" style="font-size:18px;">${fmtIDR(totalBeli)}</div><div class="sub">${invoiceCount} invoice ${lastBeli?'&middot; terakhir '+fmtDate(lastBeli):''}</div></div>
      <div class="kpi-card ${totalPiutang>0?'warn':''}"><div class="label">Piutang Saat Ini</div><div class="value" style="font-size:18px;">${fmtIDR(totalPiutang)}</div><div class="sub">${agingRows.filter(r=>r.totalSaldo>0).length} invoice belum lunas</div></div>
      <div class="kpi-card good"><div class="label">Total Dibayar</div><div class="value" style="font-size:18px;">${fmtIDR(totalDibayar)}</div><div class="sub">${lastBayar?'terakhir '+fmtDate(lastBayar):'belum ada data bayar'}</div></div>
    </div>

    <div class="c360-section">
      <h4>Produk yang Sering Dibeli</h4>
      ${topProduk.length ? '<div class="table-wrap"><table><thead><tr><th>Produk</th><th>Nilai</th></tr></thead><tbody>'+
        topProduk.map(p=>`<tr><td class="text">${esc(p[0])}</td><td>${fmtIDR(p[1])}</td></tr>`).join('')+'</tbody></table></div>'
        : '<p style="color:var(--ink-faint); font-size:12.5px;">Belum ada data pembelian untuk pelanggan ini.</p>'}
    </div>

    <div class="c360-section">
      <h4>Invoice Belum Lunas</h4>
      ${agingRows.filter(r=>r.totalSaldo>0).length ? '<div class="table-wrap"><table><thead><tr><th>Invoice</th><th>Jatuh Tempo</th><th>Status</th><th>Saldo</th></tr></thead><tbody>'+
        agingRows.filter(r=>r.totalSaldo>0).sort((a,b)=>b.totalSaldo-a.totalSaldo).map(r=>`<tr><td>${esc(r.invoice)}</td><td>${fmtDate(r.tglJatuhTempo)}</td><td><span class="badge ${bucketBadgeClass(r.bucket)}">${esc(r.bucket)}</span></td><td>${fmtIDR(r.totalSaldo)}</td></tr>`).join('')+'</tbody></table></div>'
        : '<p style="color:var(--ink-faint); font-size:12.5px;">Tidak ada piutang outstanding untuk pelanggan ini.</p>'}
    </div>

    <div class="c360-section">
      <h4>Riwayat Pembayaran Terakhir</h4>
      ${paymentRows.length ? '<div class="table-wrap"><table><thead><tr><th>Invoice</th><th>Nilai</th><th>Tgl Bayar Terakhir</th><th>Status</th></tr></thead><tbody>'+
        paymentRows.slice().sort((a,b)=>(b.tglPembTerakhir||0)-(a.tglPembTerakhir||0)).slice(0,8).map(r=>`<tr><td>${esc(r.invoice)}</td><td>${fmtIDR(r.nilaiInvoice)}</td><td>${fmtDate(r.tglPembTerakhir)}</td><td>${r.lunas?'<span class="badge b0">Lunas</span>':'<span class="badge b2">Sisa</span>'}</td></tr>`).join('')+'</tbody></table></div>'
        : '<p style="color:var(--ink-faint); font-size:12.5px;">Belum ada data pembayaran untuk pelanggan ini.</p>'}
    </div>

    <div class="c360-section">
      <h4>Transaksi Penjualan Terakhir</h4>
      ${salesRows.length ? '<div class="table-wrap"><table><thead><tr><th>Tanggal</th><th>Invoice</th><th>Produk</th><th>Qty</th><th>DPP</th></tr></thead><tbody>'+
        salesRows.slice().sort((a,b)=>b.tanggal-a.tanggal).slice(0,10).map(r=>`<tr><td>${fmtDate(r.tanggal)}</td><td>${esc(r.invoice)}</td><td class="text" title="${esc(r.namaProduk)}">${esc(r.namaProdukGroup)}</td><td>${fmtNumber(r.qty)}</td><td>${fmtIDR(r.dpp)}</td></tr>`).join('')+'</tbody></table></div>'
        : '<p style="color:var(--ink-faint); font-size:12.5px;">Belum ada data penjualan untuk pelanggan ini.</p>'}
    </div>
  `;
  openDrawer('Profil Pelanggan', html, true);
}

function bindCustomerLinks(container){
  container.querySelectorAll('.customer-link').forEach(el=>{
    el.addEventListener('click', ()=> openCustomer360(el.dataset.customer));
  });
}

function setupCustomer360Search(){
  const input = document.getElementById('customer360-search');
  const results = document.getElementById('customer360-results');
  if(!input) return;
  input.addEventListener('input', ()=>{
    const q = input.value.trim().toLowerCase();
    if(q.length<2){ results.style.display='none'; results.innerHTML=''; return; }
    const names = [...getAllCustomerNames().values()].filter(n=>n.toLowerCase().includes(q)).slice(0,15);
    if(!names.length){ results.innerHTML = '<div class="cust-result-item">Tidak ada pelanggan cocok</div>'; results.style.display='block'; return; }
    results.innerHTML = names.map(n=>`<div class="cust-result-item" data-name="${esc(n)}">${esc(n)}</div>`).join('');
    results.style.display = 'block';
    results.querySelectorAll('.cust-result-item[data-name]').forEach(el=>{
      el.addEventListener('click', ()=>{
        openCustomer360(el.dataset.name);
        results.style.display='none'; input.value='';
      });
    });
  });
  document.addEventListener('click', e=>{
    if(!results.contains(e.target) && e.target!==input) results.style.display='none';
  });
}

/* =========================================================================
   HOME / EXECUTIVE SUMMARY
========================================================================= */
function sameYMD(a,b){ return a && b && a.getFullYear()===b.getFullYear() && a.getMonth()===b.getMonth() && a.getDate()===b.getDate(); }

function renderHomeDailyWeekly(){
  let html = '';
  let hints = [];

  // --- Sales: last recorded day + last 7 days vs prior 7 days (relative to the most recent Sales data available) ---
  if(DB.sales.length){
    const lastTx = new Date(Math.max(...DB.sales.map(r=>r.tanggal.getTime())));
    const dayRows = DB.sales.filter(r=>sameYMD(r.tanggal, lastTx));
    const w1start = new Date(lastTx); w1start.setDate(w1start.getDate()-6);
    const w2start = new Date(lastTx); w2start.setDate(w2start.getDate()-13);
    const w2end = new Date(lastTx); w2end.setDate(w2end.getDate()-7);
    const last7 = DB.sales.filter(r=>r.tanggal>=w1start && r.tanggal<=lastTx);
    const prev7 = DB.sales.filter(r=>r.tanggal>=w2start && r.tanggal<=w2end);
    const omzetDay = dayRows.reduce((s,r)=>s+r.dpp,0);
    const omzet7 = last7.reduce((s,r)=>s+r.dpp,0);
    const omzetPrev7 = prev7.reduce((s,r)=>s+r.dpp,0);
    const wow = omzetPrev7 ? ((omzet7-omzetPrev7)/omzetPrev7*100) : null;
    const wowHtml = wow===null ? 'tidak ada data pembanding' : (wow>=0?'+':'')+wow.toFixed(1)+'% vs 7 hari sebelumnya';
    html += `<div class="kpi-card"><div class="label">Omzet Hari Terakhir Tercatat</div><div class="value">${fmtIDR(omzetDay)}</div><div class="sub">${fmtDate(lastTx)} &middot; ${dayRows.length} invoice</div></div>`;
    html += `<div class="kpi-card ${wow!==null && wow>=0?'good':''}"><div class="label">Omzet 7 Hari Terakhir</div><div class="value">${fmtIDR(omzet7)}</div><div class="sub ${wow!==null?(wow>=0?'up':'down'):''}">${wowHtml}</div></div>`;
    hints.push('data Sales per '+fmtDate(lastTx));
  }

  // --- AR Aging: due today + due within next 7 days (based on the real current date, since Aging is a live pull) ---
  if(DB.aging.rows.length){
    const dueToday = DB.aging.rows.filter(r=>r.totalSaldo>0 && r.tglJatuhTempo && sameYMD(r.tglJatuhTempo, TODAY));
    const in7 = new Date(TODAY); in7.setDate(in7.getDate()+7);
    const dueSoon = DB.aging.rows.filter(r=>r.totalSaldo>0 && r.tglJatuhTempo && r.tglJatuhTempo>TODAY && r.tglJatuhTempo<=in7);
    const dueTodayVal = dueToday.reduce((s,r)=>s+r.totalSaldo,0);
    const dueSoonVal = dueSoon.reduce((s,r)=>s+r.totalSaldo,0);
    html += `<div class="kpi-card ${dueToday.length?'danger':''}"><div class="label">Jatuh Tempo Hari Ini</div><div class="value">${fmtIDR(dueTodayVal)}</div><div class="sub">${dueToday.length} invoice &mdash; segera tagih</div></div>`;
    html += `<div class="kpi-card ${dueSoon.length?'warn':''}"><div class="label">Jatuh Tempo 7 Hari ke Depan</div><div class="value">${fmtIDR(dueSoonVal)}</div><div class="sub">${dueSoon.length} invoice &mdash; siapkan penagihan</div></div>`;
    hints.push('piutang per hari ini '+fmtDate(TODAY));
  }

  // --- AR Payment: last 7 days of payments received (relative to the most recent payment data available) ---
  if(DB.payment.rows.length){
    const payDates = DB.payment.rows.filter(r=>r.tglPembTerakhir).map(r=>r.tglPembTerakhir.getTime());
    if(payDates.length){
      const lastPay = new Date(Math.max(...payDates));
      const pw1start = new Date(lastPay); pw1start.setDate(pw1start.getDate()-6);
      const last7Pay = DB.payment.rows.filter(r=>r.tglPembTerakhir && r.tglPembTerakhir>=pw1start && r.tglPembTerakhir<=lastPay);
      const totalPay7 = last7Pay.reduce((s,r)=>s+r.pembTunai+r.pembDeposit+r.pembCek,0);
      html += `<div class="kpi-card good"><div class="label">Pembayaran Diterima (7 Hari Terakhir)</div><div class="value">${fmtIDR(totalPay7)}</div><div class="sub">s/d ${fmtDate(lastPay)} &middot; ${last7Pay.length} invoice</div></div>`;
      hints.push('pembayaran per '+fmtDate(lastPay));
    }
  }

  document.getElementById('home-daily-kpi').innerHTML = html || '<p style="color:var(--ink-faint); font-size:12.5px;">Belum cukup data untuk ringkasan harian.</p>';
  document.getElementById('home-daily-hint').textContent = hints.join(' · ');
}

function renderHome(){
  const hasAny = DB.sales.length || DB.aging.rows.length || DB.payment.rows.length || DB.stock.rows.length;
  document.getElementById('home-empty').style.display = hasAny ? 'none' : 'block';
  if(!hasAny){
    document.getElementById('home-kpi').innerHTML = '';
    document.getElementById('home-daily-kpi').innerHTML = '';
    document.getElementById('home-daily-hint').textContent = '';
    document.getElementById('home-attention').innerHTML = '';
    document.getElementById('home-alert-slot').innerHTML = '';
    destroyChart('homeSales'); destroyChart('homeAging');
    return;
  }
  renderHomeDailyWeekly();

  // ---- KPI row ----
  let kpiHtml = '';
  if(DB.sales.length){
    const maxDate = new Date(Math.max(...DB.sales.map(r=>r.tanggal.getTime())));
    const curMonthRows = DB.sales.filter(r=>r.year===maxDate.getFullYear() && r.month===maxDate.getMonth()+1);
    const omzetBulanIni = curMonthRows.reduce((s,r)=>s+r.dpp,0);
    kpiHtml += `<div class="kpi-card good"><div class="label">Omzet Bulan Berjalan</div><div class="value">${fmtIDR(omzetBulanIni)}</div><div class="sub">${['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Ags','Sep','Okt','Nov','Des'][maxDate.getMonth()]} ${maxDate.getFullYear()}</div></div>`;
  }
  if(DB.aging.rows.length){
    const piutang = DB.aging.rows.filter(r=>r.totalSaldo>0);
    const totalPiutang = piutang.reduce((s,r)=>s+r.totalSaldo,0);
    const kritis = piutang.filter(r=>BUCKET_RISK[r.bucket]>=3).reduce((s,r)=>s+r.totalSaldo,0);
    kpiHtml += `<div class="kpi-card"><div class="label">Total Piutang</div><div class="value">${fmtIDR(totalPiutang)}</div><div class="sub">${piutang.length} invoice</div></div>`;
    kpiHtml += `<div class="kpi-card danger"><div class="label">Piutang Kritis (&gt;60 hr)</div><div class="value">${fmtIDR(kritis)}</div><div class="sub">${piutang.filter(r=>BUCKET_RISK[r.bucket]>=3).length} invoice</div></div>`;
  }
  if(DB.payment.rows.length){
    const totalDiterima = DB.payment.rows.reduce((s,r)=>s+r.pembTunai+r.pembDeposit+r.pembCek,0);
    kpiHtml += `<div class="kpi-card"><div class="label">Pembayaran Diterima</div><div class="value">${fmtIDR(totalDiterima)}</div><div class="sub">data upload terakhir</div></div>`;
  }
  if(DB.stock.rows.length){
    const agg = stockAggregatedByProduct();
    const velocity = computeSalesVelocity();
    const reorderCount = Object.values(agg).filter(a=>{
      if(a.stokAkhir<=0) return false;
      const v = velocity[a.idProduk]; const dv = v?v.dailyVelocity:0;
      return dv>0 && a.stokAkhir/dv<=14;
    }).length;
    kpiHtml += `<div class="kpi-card ${reorderCount>0?'warn':''}"><div class="label">SKU Stock Kritis</div><div class="value">${reorderCount}</div><div class="sub">estimasi habis &le; 14 hari</div></div>`;
  }
  document.getElementById('home-kpi').innerHTML = kpiHtml;
  document.getElementById('home-asof-hint').textContent = 'Per hari ini, ' + fmtDate(TODAY);

  // ---- Sales trend chart (6 bulan terakhir, gabungan cabang) ----
  destroyChart('homeSales');
  if(DB.sales.length){
    const maxDate = new Date(Math.max(...DB.sales.map(r=>r.tanggal.getTime())));
    const months = [];
    for(let i=5;i>=0;i--){ const d = new Date(maxDate.getFullYear(), maxDate.getMonth()-i, 1); months.push({y:d.getFullYear(), m:d.getMonth()+1, label:['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Ags','Sep','Okt','Nov','Des'][d.getMonth()]+' '+d.getFullYear()}); }
    const data = months.map(mo => DB.sales.filter(r=>r.year===mo.y && r.month===mo.m).reduce((s,r)=>s+r.dpp,0));
    CHARTS['homeSales'] = safeNewChart('chart-home-sales', {
      type:'bar',
      data:{ labels: months.map(m=>m.label), datasets:[{ data, backgroundColor: PALETTE[0], borderRadius:6 }] },
      options:{ responsive:true, maintainAspectRatio:false, plugins:{legend:{display:false}, tooltip:{callbacks:{label:c=>fmtIDR(c.parsed.y)}}}, scales:{y:{ticks:{callback:v=>(v/1e6).toFixed(0)+'jt'}}} }
    });
  } else {
    const el = document.getElementById('chart-home-sales'); if(el&&el.parentElement) el.parentElement.innerHTML='<p style="color:var(--ink-faint); font-size:12.5px; text-align:center; padding:30px 0;">Belum ada data Sales</p>';
  }

  // ---- Aging composition chart ----
  destroyChart('homeAging');
  if(DB.aging.rows.length){
    const order = ['Belum Jatuh Tempo','1-14 Hari','15-30 Hari','31-60 Hari','61-90 Hari','> 90 Hari'];
    const colors = ['#4C8C7A','#7A9E97','#CBA43A','#E08A3C','#C94F3F','#7A2E2E'];
    const data = order.map(b=> DB.aging.rows.filter(r=>r.bucket===b).reduce((s,r)=>s+r.totalSaldo,0));
    CHARTS['homeAging'] = safeNewChart('chart-home-aging', {
      type:'doughnut',
      data:{ labels:order, datasets:[{ data, backgroundColor:colors, borderWidth:2, borderColor:'#fff' }] },
      options:{ responsive:true, maintainAspectRatio:false, plugins:{legend:{position:'right', labels:{boxWidth:11, font:{size:10.5}}}, tooltip:{callbacks:{label:c=>c.label+': '+fmtIDR(c.parsed)}}} }
    });
  } else {
    const el = document.getElementById('chart-home-aging'); if(el&&el.parentElement) el.parentElement.innerHTML='<p style="color:var(--ink-faint); font-size:12.5px; text-align:center; padding:30px 0;">Belum ada data AR Aging</p>';
  }

  // ---- Attention cards ----
  let attHtml = '';
  // Piutang urgent
  const urgentAging = DB.aging.rows.filter(r=>BUCKET_RISK[r.bucket]>=3).sort((a,b)=>b.totalSaldo-a.totalSaldo).slice(0,5);
  attHtml += `<div class="attention-card"><h4>&#9203; Piutang Paling Kritis</h4><ul>${
    urgentAging.length ? urgentAging.map(r=>`<li><span class="text customer-link" data-customer="${esc(r.namaPelanggan)}">${esc(r.namaPelanggan)}</span><span>${fmtIDR(r.totalSaldo)}</span></li>`).join('')
    : '<li class="empty">Tidak ada piutang &gt;60 hari</li>'
  }</ul></div>`;
  // Stock reorder
  let reorderList = [];
  if(DB.stock.rows.length){
    const agg = stockAggregatedByProduct(); const velocity = computeSalesVelocity();
    reorderList = Object.values(agg).filter(a=>a.stokAkhir>0).map(a=>{
      const v = velocity[a.idProduk]; const dv = v?v.dailyVelocity:0;
      return {...a, daysLeft: dv>0 ? a.stokAkhir/dv : null};
    }).filter(a=>a.daysLeft!==null && a.daysLeft<=14).sort((a,b)=>a.daysLeft-b.daysLeft).slice(0,5);
  }
  attHtml += `<div class="attention-card"><h4>&#128680; Stock Paling Mendesak</h4><ul>${
    reorderList.length ? reorderList.map(a=>`<li><span class="text">${esc(a.namaProduk)}</span><span>${a.daysLeft.toFixed(0)} hari</span></li>`).join('')
    : '<li class="empty">Tidak ada produk berisiko stock habis</li>'
  }</ul></div>`;
  // Payment sisa
  const sisaPayment = DB.payment.rows.filter(r=>!r.lunas).sort((a,b)=>b.sisaPemb-a.sisaPemb).slice(0,5);
  attHtml += `<div class="attention-card"><h4>&#128179; Pembayaran Belum Lunas</h4><ul>${
    sisaPayment.length ? sisaPayment.map(r=>`<li><span class="text customer-link" data-customer="${esc(r.namaPelanggan)}">${esc(r.namaPelanggan)}</span><span>${fmtIDR(r.sisaPemb)}</span></li>`).join('')
    : '<li class="empty">Tidak ada sisa pembayaran</li>'
  }</ul></div>`;
  document.getElementById('home-attention').innerHTML = attHtml;
  bindCustomerLinks(document.getElementById('home-attention'));

  // top-level alert banner combining the most urgent items
  const alertSlot = document.getElementById('home-alert-slot');
  const urgentCount = urgentAging.length, reorderCount2 = reorderList.length;
  if(urgentCount || reorderCount2){
    alertSlot.innerHTML = `<div class="alert-banner"><span class="icon">&#9888;</span><div>${urgentCount?('<strong>'+urgentCount+' piutang</strong> kritis &gt;60 hari'):''}${urgentCount&&reorderCount2?' &middot; ':''}${reorderCount2?('<strong>'+reorderCount2+' produk</strong> berisiko stock habis dalam 14 hari'):''} &mdash; lihat detail di kartu perhatian di bawah.</div></div>`;
  } else {
    alertSlot.innerHTML = '';
  }
}

/* =========================================================================
   CUSTOMER RETENTION
========================================================================= */
const retentionFilterState = { thresholdDays:'60', cabang:'all', tipeOutlet:'all', wilayah:'all' };
let retentionTableCtrl = null;
let retentionInactiveList = [];

function retentionCustomerStats(){
  // last purchase date + total historical value + last known salesman/cabang/tipeOutlet per
  // customer, computed across ALL uploaded sales data regardless of current sales-tab filters
  const map = {};
  DB.sales.forEach(r=>{
    const key = normName(r.namaPelanggan);
    if(!map[key]) map[key] = {namaPelanggan:r.namaPelanggan, idPelanggan:r.idPelanggan, cabang:r.cabang, tipeOutlet:r.tipeOutlet, lastPurchase:r.tanggal, lastSalesman:r.salesman, totalValue:0, totalQty:0, invoices:new Set()};
    const m = map[key];
    m.totalValue += r.dpp; m.totalQty += r.qty; m.invoices.add(r.invoice);
    if(r.tanggal > m.lastPurchase){ m.lastPurchase = r.tanggal; m.lastSalesman = r.salesman; m.cabang = r.cabang; m.tipeOutlet = r.tipeOutlet; m.idPelanggan = r.idPelanggan; }
  });
  const list = Object.values(map);
  list.forEach(m=>{
    const addr = getCustomerAddress(m.idPelanggan, m.namaPelanggan);
    m.alamat = addr ? addr.alamat : '';
    m.wilayah = addr ? addr.wilayah : '(Tidak diketahui)';
    m.kontakPerson = addr ? addr.kontakPerson : '';
    m.telp = addr ? addr.telp : '';
  });
  return list;
}

function renderRetentionFilterBar(){
  const bar = document.getElementById('retention-filterbar');
  const hasAddr = DB.customers.rows.length > 0;
  const allStats = retentionCustomerStats();
  const statsExceptCabang = retentionFilterState.tipeOutlet==='all' ? allStats : allStats.filter(s=>s.tipeOutlet===retentionFilterState.tipeOutlet);
  const statsExceptOutlet = retentionFilterState.cabang==='all' ? allStats : allStats.filter(s=>s.cabang===retentionFilterState.cabang);
  const statsExceptWilayah = allStats.filter(s=> (retentionFilterState.cabang==='all'||s.cabang===retentionFilterState.cabang) && (retentionFilterState.tipeOutlet==='all'||s.tipeOutlet===retentionFilterState.tipeOutlet));
  const cabangs = [...new Set(statsExceptCabang.map(s=>s.cabang).filter(Boolean))].sort();
  const outlets = [...new Set(statsExceptOutlet.map(s=>s.tipeOutlet).filter(Boolean))].sort();
  const wilayahs = [...new Set(statsExceptWilayah.map(s=>s.wilayah).filter(Boolean))].sort();
  if(retentionFilterState.cabang!=='all' && !cabangs.includes(retentionFilterState.cabang)) retentionFilterState.cabang='all';
  if(retentionFilterState.tipeOutlet!=='all' && !outlets.includes(retentionFilterState.tipeOutlet)) retentionFilterState.tipeOutlet='all';
  if(retentionFilterState.wilayah!=='all' && !wilayahs.includes(retentionFilterState.wilayah)) retentionFilterState.wilayah='all';
  bar.innerHTML = `
    <div class="filter-field"><label>Cabang</label><select id="f-ret-cabang"><option value="all">Semua</option>${cabangs.map(c=>'<option value="'+esc(c)+'" '+(retentionFilterState.cabang===c?'selected':'')+'>'+esc(c)+'</option>').join('')}</select></div>
    <div class="filter-field"><label>Tipe Outlet</label><select id="f-ret-outlet"><option value="all">Semua</option>${outlets.map(o=>'<option value="'+esc(o)+'" '+(retentionFilterState.tipeOutlet===o?'selected':'')+'>'+esc(o)+'</option>').join('')}</select></div>
    ${hasAddr ? '<div class="filter-field"><label>Wilayah</label><select id="f-ret-wilayah"><option value="all">Semua</option>'+wilayahs.map(w=>'<option value="'+esc(w)+'" '+(retentionFilterState.wilayah===w?'selected':'')+'>'+esc(w)+'</option>').join('')+'</select></div>' : ''}
    <div class="filter-field"><label>Dianggap Tidak Aktif Jika &ge;</label><select id="f-ret-threshold">
      <option value="30" ${retentionFilterState.thresholdDays==='30'?'selected':''}>30 hari</option>
      <option value="60" ${retentionFilterState.thresholdDays==='60'?'selected':''}>60 hari</option>
      <option value="90" ${retentionFilterState.thresholdDays==='90'?'selected':''}>90 hari</option>
      <option value="120" ${retentionFilterState.thresholdDays==='120'?'selected':''}>120 hari</option>
    </select></div>
    <button class="btn-reset" id="f-ret-reset">Reset Filter</button>
  `;
  document.getElementById('f-ret-cabang').addEventListener('change', e=>{ retentionFilterState.cabang=e.target.value; renderRetention(); });
  document.getElementById('f-ret-outlet').addEventListener('change', e=>{ retentionFilterState.tipeOutlet=e.target.value; renderRetention(); });
  if(hasAddr) document.getElementById('f-ret-wilayah').addEventListener('change', e=>{ retentionFilterState.wilayah=e.target.value; renderRetention(); });
  document.getElementById('f-ret-threshold').addEventListener('change', e=>{ retentionFilterState.thresholdDays=e.target.value; renderRetention(); });
  document.getElementById('f-ret-reset').addEventListener('click', ()=>{ Object.assign(retentionFilterState,{thresholdDays:'60',cabang:'all',tipeOutlet:'all',wilayah:'all'}); renderRetention(); });
}

function renderRetention(){
  const hasData = DB.sales.length>0;
  document.getElementById('retention-empty').style.display = hasData? 'none':'block';
  document.getElementById('retention-content').style.display = hasData? 'block':'none';
  if(!hasData) return;
  rebuildPreservingFocus(document.getElementById('retention-filterbar'), renderRetentionFilterBar);

  const maxDate = new Date(Math.max(...DB.sales.map(r=>r.tanggal.getTime())));
  const threshold = parseInt(retentionFilterState.thresholdDays,10);
  let stats = retentionCustomerStats();
  if(retentionFilterState.cabang!=='all') stats = stats.filter(s=>s.cabang===retentionFilterState.cabang);
  if(retentionFilterState.tipeOutlet!=='all') stats = stats.filter(s=>s.tipeOutlet===retentionFilterState.tipeOutlet);
  if(retentionFilterState.wilayah!=='all') stats = stats.filter(s=>s.wilayah===retentionFilterState.wilayah);
  stats.forEach(s=>{ s.daysSince = daysBetween(maxDate, s.lastPurchase); s.invoiceCount = s.invoices.size; });

  const aktif = stats.filter(s=>s.daysSince < threshold);
  const tidakAktif = stats.filter(s=>s.daysSince >= threshold).sort((a,b)=>b.totalValue-a.totalValue);
  retentionInactiveList = tidakAktif;
  const avgInactive = tidakAktif.length ? tidakAktif.reduce((s,r)=>s+r.daysSince,0)/tidakAktif.length : 0;
  const valueAtRisk = tidakAktif.reduce((s,r)=>s+r.totalValue,0);

  document.getElementById('retention-kpi').innerHTML = `
    <div class="kpi-card"><div class="label">Total Pelanggan Pernah Beli</div><div class="value">${stats.length}</div></div>
    <div class="kpi-card good"><div class="label">Aktif (&lt; ${threshold} hari)</div><div class="value">${aktif.length}</div><div class="sub">${stats.length?((aktif.length/stats.length*100).toFixed(0)+'% dari total'):''}</div></div>
    <div class="kpi-card danger"><div class="label">Tidak Aktif (&ge; ${threshold} hari)</div><div class="value">${tidakAktif.length}</div><div class="sub">rata-rata ${avgInactive.toFixed(0)} hari tidak beli</div></div>
    <div class="kpi-card warn"><div class="label">Nilai Historis Berisiko</div><div class="value">${fmtIDR(valueAtRisk)}</div><div class="sub">total pembelian lama pelanggan tidak aktif</div></div>
  `;
  document.getElementById('retention-asof-hint').textContent = 'Dihitung dari tanggal transaksi terakhir per pelanggan, per data termutakhir ('+fmtDate(maxDate)+')';

  // distribution chart
  const buckets = [
    {label:'Aktif (<'+threshold+' hr)', min:-Infinity, max:threshold-1},
    {label:threshold+'-'+(threshold+29)+' hr', min:threshold, max:threshold+29},
    {label:(threshold+30)+'-'+(threshold+89)+' hr', min:threshold+30, max:threshold+89},
    {label:'>'+(threshold+89)+' hr', min:threshold+90, max:Infinity}
  ];
  const distData = buckets.map(b=> stats.filter(s=>s.daysSince>=b.min && s.daysSince<=b.max).length);
  destroyChart('retentionDist');
  CHARTS['retentionDist'] = safeNewChart('chart-retention-dist', {
    type:'bar',
    data:{ labels: buckets.map(b=>b.label), datasets:[{ data:distData, backgroundColor:[PALETTE[2],'#CBA43A','#E08A3C','#C94F3F'], borderRadius:6 }] },
    options:{ responsive:true, maintainAspectRatio:false, plugins:{legend:{display:false}}, scales:{y:{ticks:{precision:0}}} }
  });

  // per-cabang summary
  const cabMap = {};
  stats.forEach(s=>{
    if(!cabMap[s.cabang]) cabMap[s.cabang] = {cabang:s.cabang, total:0, aktif:0, tidakAktif:0};
    cabMap[s.cabang].total++;
    if(s.daysSince<threshold) cabMap[s.cabang].aktif++; else cabMap[s.cabang].tidakAktif++;
  });
  const cabList = Object.values(cabMap).sort((a,b)=>b.tidakAktif-a.tidakAktif);
  document.querySelector('#table-retention-cabang thead').innerHTML = '<tr><th>Cabang</th><th>Total</th><th>Aktif</th><th>Tidak Aktif</th></tr>';
  document.querySelector('#table-retention-cabang tbody').innerHTML = cabList.length ? cabList.map(c=>`<tr><td class="text">${esc(c.cabang)}</td><td>${c.total}</td><td>${c.aktif}</td><td>${c.tidakAktif}</td></tr>`).join('') :
    '<tr><td colspan="4" style="text-align:center; color:var(--ink-faint); padding:16px;">Tidak ada data</td></tr>';

  // per-tipe-outlet summary — sorted by highest inactive count first, so the outlet
  // segment with the most non-ordering customers stands out immediately
  const outletMap = {};
  stats.forEach(s=>{
    const key = s.tipeOutlet || '(kosong)';
    if(!outletMap[key]) outletMap[key] = {tipeOutlet:key, total:0, aktif:0, tidakAktif:0};
    outletMap[key].total++;
    if(s.daysSince<threshold) outletMap[key].aktif++; else outletMap[key].tidakAktif++;
  });
  const outletList = Object.values(outletMap).sort((a,b)=>b.tidakAktif-a.tidakAktif);
  document.querySelector('#table-retention-outlet thead').innerHTML = '<tr><th>Tipe Outlet</th><th>Total</th><th>Aktif</th><th>Tidak Aktif</th></tr>';
  document.querySelector('#table-retention-outlet tbody').innerHTML = outletList.length ? outletList.map(o=>`<tr><td class="text">${esc(o.tipeOutlet)}</td><td>${o.total}</td><td>${o.aktif}</td><td>${o.tidakAktif>0?'<strong>'+o.tidakAktif+'</strong>':'0'}</td></tr>`).join('') :
    '<tr><td colspan="4" style="text-align:center; color:var(--ink-faint); padding:16px;">Tidak ada data</td></tr>';

  // per-wilayah summary (from the uploaded customer address database) — only shown once
  // address data has been uploaded, sorted by highest inactive count first
  const wilayahWrap = document.getElementById('retention-wilayah-wrap');
  const hasAddr = DB.customers.rows.length > 0;
  wilayahWrap.style.display = hasAddr ? 'block' : 'none';
  if(hasAddr){
    const wilMap = {};
    stats.forEach(s=>{
      const key = s.wilayah || '(Tidak diketahui)';
      if(!wilMap[key]) wilMap[key] = {wilayah:key, total:0, aktif:0, tidakAktif:0};
      wilMap[key].total++;
      if(s.daysSince<threshold) wilMap[key].aktif++; else wilMap[key].tidakAktif++;
    });
    const wilList = Object.values(wilMap).sort((a,b)=>b.tidakAktif-a.tidakAktif).slice(0,15);
    document.querySelector('#table-retention-wilayah thead').innerHTML = '<tr><th>Wilayah</th><th>Total</th><th>Aktif</th><th>Tidak Aktif</th></tr>';
    document.querySelector('#table-retention-wilayah tbody').innerHTML = wilList.length ? wilList.map(w=>`<tr><td class="text">${esc(w.wilayah)}</td><td>${w.total}</td><td>${w.aktif}</td><td>${w.tidakAktif>0?'<strong>'+w.tidakAktif+'</strong>':'0'}</td></tr>`).join('') :
      '<tr><td colspan="4" style="text-align:center; color:var(--ink-faint); padding:16px;">Tidak ada data</td></tr>';
  }

  // detail table of inactive customers
  if(!retentionTableCtrl){
    retentionTableCtrl = makeTableController({
      getRows: ()=> retentionInactiveList,
      columns:[
        {key:'namaPelanggan',label:'Pelanggan'},{key:'cabang',label:'Cabang'},{key:'tipeOutlet',label:'Tipe Outlet'},
        {key:'wilayah',label:'Wilayah'},{key:'alamat',label:'Alamat'},{key:'kontakPerson',label:'Kontak'},{key:'telp',label:'Telp'},
        {key:'lastPurchase',label:'Terakhir Beli'},
        {key:'daysSince',label:'Hari Tidak Beli'},{key:'totalValue',label:'Total Pembelian Historis'},{key:'lastSalesman',label:'Sales Terakhir'}
      ],
      renderRow: r=>`<tr><td class="text customer-link" data-customer="${esc(r.namaPelanggan)}">${esc(r.namaPelanggan)}</td><td class="text">${esc(r.cabang)}</td><td class="text">${esc(r.tipeOutlet)}</td><td class="text">${esc(r.wilayah)}</td><td class="text" title="${esc(r.alamat)}" style="max-width:220px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">${esc(r.alamat && r.alamat.length>45 ? r.alamat.slice(0,45)+'…' : (r.alamat||'-'))}</td><td class="text">${esc(r.kontakPerson||'-')}</td><td class="text">${esc(r.telp||'-')}</td><td>${fmtDate(r.lastPurchase)}</td><td>${r.daysSince}</td><td>${fmtIDR(r.totalValue)}</td><td class="text">${esc(r.lastSalesman)}</td></tr>`,
      afterRender: (tbody)=> bindCustomerLinks(tbody),
      pageSize:20, tbodyId:'table-retention-detail-tbody', theadId:'table-retention-detail-thead', pagerId:'retention-pager', infoId:'retention-pager-info',
      defaultSort:{key:'totalValue', dir:'desc'}
    });
    document.querySelector('#table-retention-detail').innerHTML = '<thead id="table-retention-detail-thead"></thead><tbody id="table-retention-detail-tbody"></tbody>';
  }
  retentionTableCtrl.resetPage();
  retentionTableCtrl.render();
  document.getElementById('retention-table-count').textContent = tidakAktif.length + ' pelanggan tidak aktif';
}

const BULAN_NAMES = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Ags','Sep','Okt','Nov','Des'];

/* =========================================================================
   EVALUASI TIM (Perbandingan MoM / YoY, drilldown per Salesman -> Produk)
========================================================================= */
const evaluasiFilterState = { mode:'mom', anchorYear:null, anchorMonth:null, cabang:'all', valueBasis:'dpp' };
function evalVal(r){ return evaluasiFilterState.valueBasis==='netto' ? r.netto : r.dpp; }
function evalValLabel(){ return evaluasiFilterState.valueBasis==='netto' ? 'Netto (+PPN)' : 'DPP'; }

function evaluasiAnchorOptions(){
  const set = new Map();
  DB.sales.forEach(r=>{ set.set(r.year+'-'+r.month, {year:r.year, month:r.month}); });
  return [...set.values()].sort((a,b)=> b.year-a.year || b.month-a.month);
}
function evaluasiComparePeriod(anchorYear, anchorMonth, mode){
  if(mode==='yoy') return {year:anchorYear-1, month:anchorMonth};
  let y=anchorYear, m=anchorMonth-1;
  if(m<1){ m=12; y-=1; }
  return {year:y, month:m};
}
function evaluasiRowsFor(year, month, cabang){
  return DB.sales.filter(r=> r.year===year && r.month===month && (cabang==='all'||r.cabang===cabang));
}

function renderEvaluasiFilterBar(){
  const bar = document.getElementById('evaluasi-filterbar');
  const anchors = evaluasiAnchorOptions();
  if(!evaluasiFilterState.anchorYear && anchors.length){
    evaluasiFilterState.anchorYear = anchors[0].year;
    evaluasiFilterState.anchorMonth = anchors[0].month;
  }
  const cabangs = uniqueVals(DB.sales,'cabang');
  if(evaluasiFilterState.cabang!=='all' && !cabangs.includes(evaluasiFilterState.cabang)) evaluasiFilterState.cabang='all';
  const anchorVal = evaluasiFilterState.anchorYear+'-'+evaluasiFilterState.anchorMonth;
  bar.innerHTML = `
    <div class="filter-field"><label>Basis Nilai</label><select id="f-eval-basis">
      <option value="dpp" ${evaluasiFilterState.valueBasis==='dpp'?'selected':''}>DPP (sebelum PPN)</option>
      <option value="netto" ${evaluasiFilterState.valueBasis==='netto'?'selected':''}>Netto (termasuk PPN)</option>
    </select></div>
    <div class="filter-field"><label>Tipe Perbandingan</label><select id="f-eval-mode">
      <option value="mom" ${evaluasiFilterState.mode==='mom'?'selected':''}>MoM (bulan vs bulan lalu)</option>
      <option value="yoy" ${evaluasiFilterState.mode==='yoy'?'selected':''}>YoY (bulan vs tahun lalu)</option>
    </select></div>
    <div class="filter-field"><label>Periode Acuan</label><select id="f-eval-anchor">${anchors.map(a=>{
      const v=a.year+'-'+a.month;
      return '<option value="'+v+'" '+(anchorVal===v?'selected':'')+'>'+BULAN_NAMES[a.month-1]+' '+a.year+'</option>';
    }).join('')}</select></div>
    <div class="filter-field"><label>Cabang</label><select id="f-eval-cabang"><option value="all">Semua</option>${cabangs.map(c=>'<option value="'+esc(c)+'" '+(evaluasiFilterState.cabang===c?'selected':'')+'>'+esc(c)+'</option>').join('')}</select></div>
    <button class="btn-reset" id="f-eval-reset">Reset Filter</button>
  `;
  document.getElementById('f-eval-basis').addEventListener('change', e=>{ evaluasiFilterState.valueBasis=e.target.value; renderEvaluasi(); });
  document.getElementById('f-eval-mode').addEventListener('change', e=>{ evaluasiFilterState.mode=e.target.value; renderEvaluasi(); });
  document.getElementById('f-eval-anchor').addEventListener('change', e=>{
    const [y,m] = e.target.value.split('-').map(Number);
    evaluasiFilterState.anchorYear=y; evaluasiFilterState.anchorMonth=m; renderEvaluasi();
  });
  document.getElementById('f-eval-cabang').addEventListener('change', e=>{ evaluasiFilterState.cabang=e.target.value; renderEvaluasi(); });
  document.getElementById('f-eval-reset').addEventListener('click', ()=>{
    evaluasiFilterState.mode='mom'; evaluasiFilterState.cabang='all'; evaluasiFilterState.valueBasis='dpp';
    if(anchors.length){ evaluasiFilterState.anchorYear=anchors[0].year; evaluasiFilterState.anchorMonth=anchors[0].month; }
    renderEvaluasi();
  });
}

function evaluasiPctBadge(deltaPct){
  if(deltaPct===null) return '<span class="badge neutral">Baru</span>';
  const lvl = deltaPct>=0 ? 'b0' : 'b3';
  return '<span class="badge '+lvl+'">'+(deltaPct>=0?'+':'')+deltaPct.toFixed(1)+'%</span>';
}

function renderEvaluasiKPI(curRows, cmpRows){
  const curTotal = curRows.reduce((s,r)=>s+evalVal(r),0);
  const cmpTotal = cmpRows.reduce((s,r)=>s+evalVal(r),0);
  const delta = curTotal-cmpTotal;
  const deltaPct = cmpTotal ? (delta/cmpTotal*100) : null;
  const salesmen = new Set([...curRows.map(r=>r.salesman), ...cmpRows.map(r=>r.salesman)]);
  let naik=0, turun=0;
  salesmen.forEach(sm=>{
    const c = curRows.filter(r=>r.salesman===sm).reduce((s,r)=>s+evalVal(r),0);
    const p = cmpRows.filter(r=>r.salesman===sm).reduce((s,r)=>s+evalVal(r),0);
    if(p>0){ if(c>p) naik++; else if(c<p) turun++; }
  });
  document.getElementById('evaluasi-kpi').innerHTML = `
    <div class="kpi-card"><div class="label">Omzet Periode Ini (${evalValLabel()})</div><div class="value">${fmtIDR(curTotal)}</div></div>
    <div class="kpi-card"><div class="label">Omzet Periode Pembanding</div><div class="value">${fmtIDR(cmpTotal)}</div><div class="sub">${cmpRows.length===0?'belum ada data periode ini':''}</div></div>
    <div class="kpi-card ${deltaPct===null?'':deltaPct>=0?'good':'danger'}"><div class="label">Perubahan</div><div class="value">${fmtIDR(delta)}</div><div class="sub">${deltaPct===null?'tidak ada pembanding':(deltaPct>=0?'+':'')+deltaPct.toFixed(1)+'%'}</div></div>
    <div class="kpi-card"><div class="label">Salesman Naik / Turun</div><div class="value">${naik} / ${turun}</div><div class="sub">dari ${salesmen.size} salesman</div></div>
  `;
}

function renderEvaluasiProductDrilldown(salesman, curRows, cmpRows, container){
  const cur = curRows.filter(r=>r.salesman===salesman);
  const cmp = cmpRows.filter(r=>r.salesman===salesman);
  const map = {};
  cur.forEach(r=>{ if(!map[r.namaProdukGroup]) map[r.namaProdukGroup]={namaProdukGroup:r.namaProdukGroup, cur:0, cmp:0}; map[r.namaProdukGroup].cur+=evalVal(r); });
  cmp.forEach(r=>{ if(!map[r.namaProdukGroup]) map[r.namaProdukGroup]={namaProdukGroup:r.namaProdukGroup, cur:0, cmp:0}; map[r.namaProdukGroup].cmp+=evalVal(r); });
  const list = Object.values(map).map(m=>({...m, delta:m.cur-m.cmp})).sort((a,b)=>b.delta-a.delta);
  if(!list.length){ container.innerHTML='<p style="color:var(--ink-faint); font-size:12.5px; padding:6px 0;">Tidak ada data produk untuk salesman ini pada kedua periode.</p>'; return; }
  container.innerHTML = '<table><thead><tr><th>Produk</th><th>Periode Ini</th><th>Periode Pembanding</th><th>Selisih</th></tr></thead><tbody>'+
    list.map(m=>`<tr><td class="text">${esc(m.namaProdukGroup)}</td><td>${fmtIDR(m.cur)}</td><td>${fmtIDR(m.cmp)}</td><td style="color:${m.delta>=0?'var(--risk-0)':'var(--risk-3)'}; font-weight:600;">${m.delta>=0?'+':''}${fmtIDR(m.delta)}</td></tr>`).join('')+
    '</tbody></table>';
}

function renderEvaluasiSalesmanTable(curRows, cmpRows){
  const map = {};
  curRows.forEach(r=>{ if(!map[r.salesman]) map[r.salesman]={salesman:r.salesman, cur:0, cmp:0}; map[r.salesman].cur+=evalVal(r); });
  cmpRows.forEach(r=>{ if(!map[r.salesman]) map[r.salesman]={salesman:r.salesman, cur:0, cmp:0}; map[r.salesman].cmp+=evalVal(r); });
  const list = Object.values(map).map(m=>({...m, delta:m.cur-m.cmp, deltaPct: m.cmp? (m.cur-m.cmp)/m.cmp*100 : null}))
    .sort((a,b)=>{ const ap=a.deltaPct===null?Infinity:a.deltaPct, bp=b.deltaPct===null?Infinity:b.deltaPct; return ap-bp; });
  const thead = document.querySelector('#table-evaluasi-salesman thead');
  const tbody = document.querySelector('#table-evaluasi-salesman tbody');
  thead.innerHTML = '<tr><th>Salesman</th><th>Periode Ini</th><th>Periode Pembanding</th><th>Selisih</th><th>Perubahan</th></tr>';
  if(!list.length){
    tbody.innerHTML = '<tr><td colspan="5" style="text-align:center; color:var(--ink-faint); padding:20px;">Tidak ada data</td></tr>';
    return;
  }
  tbody.innerHTML = list.map(m=>`
    <tr class="expand-row" data-salesman="${esc(m.salesman)}">
      <td class="text">${esc(m.salesman)}</td><td>${fmtIDR(m.cur)}</td><td>${fmtIDR(m.cmp)}</td><td>${fmtIDR(m.delta)}</td><td>${evaluasiPctBadge(m.deltaPct)}</td>
    </tr>
    <tr class="sub-panel" data-subfor="${esc(m.salesman)}" style="display:none;"><td colspan="5"><div class="sub-panel-inner"></div></td></tr>
  `).join('');
  tbody.querySelectorAll('tr.expand-row').forEach(tr=>{
    tr.addEventListener('click', ()=>{
      const sm = tr.dataset.salesman;
      const subRow = [...tbody.querySelectorAll('tr.sub-panel')].find(x=>x.dataset.subfor===sm);
      const isOpen = tr.classList.contains('open');
      tbody.querySelectorAll('tr.expand-row.open').forEach(o=>{
        o.classList.remove('open');
        const s = [...tbody.querySelectorAll('tr.sub-panel')].find(x=>x.dataset.subfor===o.dataset.salesman);
        if(s) s.style.display='none';
      });
      if(!isOpen){
        tr.classList.add('open');
        subRow.style.display='table-row';
        renderEvaluasiProductDrilldown(sm, curRows, cmpRows, subRow.querySelector('.sub-panel-inner'));
      }
    });
  });
}

function renderEvaluasiTopMovers(curRows, cmpRows){
  const map = {};
  curRows.forEach(r=>{ if(!map[r.namaProdukGroup]) map[r.namaProdukGroup]={namaProdukGroup:r.namaProdukGroup, cur:0, cmp:0}; map[r.namaProdukGroup].cur+=evalVal(r); });
  cmpRows.forEach(r=>{ if(!map[r.namaProdukGroup]) map[r.namaProdukGroup]={namaProdukGroup:r.namaProdukGroup, cur:0, cmp:0}; map[r.namaProdukGroup].cmp+=evalVal(r); });
  const list = Object.values(map).map(m=>({...m, delta:m.cur-m.cmp}));
  const naik = list.filter(m=>m.cmp>0 && m.delta>0).sort((a,b)=>b.delta-a.delta).slice(0,8);
  const turun = list.filter(m=>m.cmp>0 && m.delta<0).sort((a,b)=>a.delta-b.delta).slice(0,8);
  document.querySelector('#table-evaluasi-naik thead').innerHTML = '<tr><th>Produk</th><th>Selisih</th></tr>';
  document.querySelector('#table-evaluasi-naik tbody').innerHTML = naik.length ? naik.map(m=>`<tr><td class="text">${esc(m.namaProdukGroup)}</td><td style="color:var(--risk-0); font-weight:600;">+${fmtIDR(m.delta)}</td></tr>`).join('') :
    '<tr><td style="text-align:center; color:var(--ink-faint); padding:16px;">Tidak ada data pembanding</td></tr>';
  document.querySelector('#table-evaluasi-turun thead').innerHTML = '<tr><th>Produk</th><th>Selisih</th></tr>';
  document.querySelector('#table-evaluasi-turun tbody').innerHTML = turun.length ? turun.map(m=>`<tr><td class="text">${esc(m.namaProdukGroup)}</td><td style="color:var(--risk-3); font-weight:600;">${fmtIDR(m.delta)}</td></tr>`).join('') :
    '<tr><td style="text-align:center; color:var(--ink-faint); padding:16px;">Tidak ada data pembanding</td></tr>';
}

/* ---- Target Penjualan & Pencapaian ---- */
let _salesmanCabangCache = null;
function getSalesmanCabangMap(){
  if(_salesmanCabangCache) return _salesmanCabangCache;
  const map = new Map();
  DB.sales.forEach(r=>{ if(r.salesman && r.cabang && !map.has(r.salesman)) map.set(r.salesman, r.cabang); });
  _salesmanCabangCache = map;
  return map;
}
function invalidateSalesmanCabangCache(){ _salesmanCabangCache = null; }

/* ---- Database Pelanggan (alamat) lookup ---- */
let _customerAddressCacheById = null, _customerAddressCacheByName = null;
function buildCustomerAddressCaches(){
  const byId = new Map(), byName = new Map();
  DB.customers.rows.forEach(c=>{
    if(c.idPelangganKey) byId.set(c.idPelangganKey, c);
    if(c.namaPelanggan && !byName.has(normName(c.namaPelanggan))) byName.set(normName(c.namaPelanggan), c);
  });
  _customerAddressCacheById = byId; _customerAddressCacheByName = byName;
}
function invalidateCustomerAddressCache(){ _customerAddressCacheById = null; _customerAddressCacheByName = null; }
function getCustomerAddress(idPelanggan, namaPelanggan){
  if(!DB.customers.rows.length) return null;
  if(!_customerAddressCacheById) buildCustomerAddressCaches();
  if(idPelanggan){
    const hit = _customerAddressCacheById.get(normProdCode(idPelanggan));
    if(hit) return hit;
  }
  return _customerAddressCacheByName.get(normName(namaPelanggan)) || null;
}

function getTargetMap(){
  const map = new Map();
  DB.targets.rows.forEach(r=>{ map.set(normName(r.salesman), r.targetBulanan); });
  return map;
}

function renderTargetSection(curRows, ytdRows){
  const targetMap = getTargetMap();
  const cabangMap = getSalesmanCabangMap();
  const cabangFilter = evaluasiFilterState.cabang;

  const salesmenSet = new Set();
  DB.targets.rows.forEach(r=>{
    const c = cabangMap.get(r.salesman);
    if(cabangFilter==='all' || c===cabangFilter || !c) salesmenSet.add(r.salesman);
  });
  curRows.forEach(r=>salesmenSet.add(r.salesman));

  const list = [...salesmenSet].map(sm=>{
    const targetBulanan = targetMap.get(normName(sm)) || 0;
    const actualBulan = curRows.filter(r=>r.salesman===sm).reduce((s,r)=>s+evalVal(r),0);
    const actualYtd = ytdRows.filter(r=>r.salesman===sm).reduce((s,r)=>s+evalVal(r),0);
    const targetTahunan = targetBulanan*12;
    const pctBulan = targetBulanan ? actualBulan/targetBulanan*100 : null;
    const pctTahun = targetTahunan ? actualYtd/targetTahunan*100 : null;
    return {salesman:sm, targetBulanan, actualBulan, pctBulan, targetTahunan, actualYtd, pctTahun};
  }).sort((a,b)=>{ const ap=a.pctBulan===null?Infinity:a.pctBulan, bp=b.pctBulan===null?Infinity:b.pctBulan; return ap-bp; });

  const totalTargetBulan = list.reduce((s,m)=>s+m.targetBulanan,0);
  const totalActualBulan = list.reduce((s,m)=>s+m.actualBulan,0);
  const totalTargetTahun = list.reduce((s,m)=>s+m.targetTahunan,0);
  const totalActualYtd = list.reduce((s,m)=>s+m.actualYtd,0);
  const pctBulanTotal = totalTargetBulan ? totalActualBulan/totalTargetBulan*100 : null;
  const pctTahunTotal = totalTargetTahun ? totalActualYtd/totalTargetTahun*100 : null;

  document.getElementById('target-scope-hint').textContent =
    (cabangFilter==='all'?'Semua cabang':cabangFilter) + ' &middot; ' + BULAN_NAMES[evaluasiFilterState.anchorMonth-1]+' '+evaluasiFilterState.anchorYear;

  document.getElementById('target-kpi').innerHTML = `
    <div class="kpi-card"><div class="label">Target Bulanan</div><div class="value">${fmtIDR(totalTargetBulan)}</div></div>
    <div class="kpi-card ${pctBulanTotal===null?'':pctBulanTotal>=100?'good':'warn'}"><div class="label">Actual Bulan Ini</div><div class="value">${fmtIDR(totalActualBulan)}</div><div class="sub">${pctBulanTotal===null?'belum ada target':pctBulanTotal.toFixed(0)+'% dari target'}</div></div>
    <div class="kpi-card"><div class="label">Target Tahunan</div><div class="value">${fmtIDR(totalTargetTahun)}</div><div class="sub">target bulanan &times; 12</div></div>
    <div class="kpi-card ${pctTahunTotal===null?'':pctTahunTotal>=100?'good':'warn'}"><div class="label">Actual Tahun Berjalan</div><div class="value">${fmtIDR(totalActualYtd)}</div><div class="sub">${pctTahunTotal===null?'belum ada target':pctTahunTotal.toFixed(0)+'% dari target tahunan'}</div></div>
  `;

  const thead = document.querySelector('#table-target-salesman thead');
  const tbody = document.querySelector('#table-target-salesman tbody');
  thead.innerHTML = '<tr><th>Salesman</th><th>Target Bulanan</th><th>Actual Bulan Ini</th><th>Pencapaian</th><th>Target Tahunan</th><th>Actual YTD</th><th>Pencapaian Tahunan</th></tr>';
  tbody.innerHTML = list.length ? list.map(m=>{
    const badgeBulan = m.pctBulan===null ? '<span class="badge neutral">Tanpa target</span>' : `<span class="badge ${m.pctBulan>=100?'b0':m.pctBulan>=80?'b1':'b3'}">${m.pctBulan.toFixed(0)}%</span>`;
    const badgeTahun = m.pctTahun===null ? '<span class="badge neutral">Tanpa target</span>' : `<span class="badge ${m.pctTahun>=100?'b0':m.pctTahun>=80?'b1':'b3'}">${m.pctTahun.toFixed(0)}%</span>`;
    return `<tr><td class="text">${esc(m.salesman)}</td><td>${fmtIDR(m.targetBulanan)}</td><td>${fmtIDR(m.actualBulan)}</td><td>${badgeBulan}</td><td>${fmtIDR(m.targetTahunan)}</td><td>${fmtIDR(m.actualYtd)}</td><td>${badgeTahun}</td></tr>`;
  }).join('') : '<tr><td colspan="7" style="text-align:center; color:var(--ink-faint); padding:20px;">Belum ada data</td></tr>';
}

/* ---- Tren Pencapaian vs Target per Salesman (line chart) ---- */
const targetTrendState = { salesman: null };

function renderTargetTrendChart(){
  const sel = document.getElementById('f-target-trend-salesman');
  const cabangMap = getSalesmanCabangMap();
  const cabangFilter = evaluasiFilterState.cabang;
  let salesmen = uniqueVals(DB.sales,'salesman');
  if(cabangFilter!=='all') salesmen = salesmen.filter(sm=> cabangMap.get(sm)===cabangFilter);

  if(!targetTrendState.salesman || !salesmen.includes(targetTrendState.salesman)){
    targetTrendState.salesman = salesmen[0] || null;
  }
  sel.innerHTML = salesmen.map(sm=>`<option value="${esc(sm)}" ${sm===targetTrendState.salesman?'selected':''}>${esc(sm)}</option>`).join('');
  sel.onchange = ()=>{ targetTrendState.salesman = sel.value; renderTargetTrendChart(); };

  destroyChart('targetTrend');
  if(!targetTrendState.salesman){
    const el = document.getElementById('chart-target-trend');
    if(el && el.parentElement) el.parentElement.innerHTML = '<p style="color:var(--ink-faint); font-size:12.5px; text-align:center; padding:30px 0;">Upload data Sales dulu untuk melihat tren ini</p>';
    document.getElementById('target-trend-hint').textContent = '';
    return;
  }

  const sm = targetTrendState.salesman;
  const targetBulanan = getTargetMap().get(normName(sm)) || null;

  // Every year-month combo present in the data for this salesman (chronological order)
  const monthsSet = new Map();
  DB.sales.filter(r=>r.salesman===sm && (cabangFilter==='all'||r.cabang===cabangFilter)).forEach(r=>{
    monthsSet.set(r.year+'-'+r.month, {year:r.year, month:r.month});
  });
  const months = [...monthsSet.values()].sort((a,b)=> a.year-b.year || a.month-b.month);
  const labels = months.map(m=>BULAN_NAMES[m.month-1]+' '+m.year);
  const actuals = months.map(m=>{
    const rows = DB.sales.filter(r=>r.salesman===sm && r.year===m.year && r.month===m.month && (cabangFilter==='all'||r.cabang===cabangFilter));
    return rows.reduce((s,r)=>s+evalVal(r),0);
  });
  const targets = months.map(()=> targetBulanan);

  const datasets = [
    { label:'Pencapaian Aktual ('+evalValLabel()+')', data:actuals, borderColor:PALETTE[0], backgroundColor:PALETTE[0]+'22', tension:0.3, fill:false, borderWidth:2.5, pointRadius:3 }
  ];
  if(targetBulanan){
    datasets.push({ label:'Target Bulanan', data:targets, borderColor:'#C94F3F', borderDash:[6,4], borderWidth:2, pointRadius:0, tension:0, fill:false });
  }

  CHARTS['targetTrend'] = safeNewChart('chart-target-trend', {
    type:'line',
    data:{ labels, datasets },
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{legend:{position:'top', labels:{boxWidth:12, usePointStyle:true}}, tooltip:{callbacks:{label:c=>c.dataset.label+': '+fmtIDR(c.parsed.y)}}},
      scales:{ y:{ticks:{callback:v=>(v/1e6).toFixed(0)+'jt'}}, x:{grid:{display:false}} }
    }
  });
  document.getElementById('target-trend-hint').textContent = targetBulanan ? '' : sm+' belum punya target bulanan yang diset (isi di "Kelola Target" di bawah)';
}

function renderTargetManageTable(){
  const targetMap = getTargetMap();
  const cabangMap = getSalesmanCabangMap();
  const cabangFilter = evaluasiFilterState.cabang;
  let salesmen = [...new Set([...uniqueVals(DB.sales,'salesman'), ...DB.targets.rows.map(r=>r.salesman)])].sort();
  if(cabangFilter!=='all'){
    // Only keep salesmen we can confirm belong to the selected cabang (from Sales data).
    // With single-cabang data every salesman maps to that one cabang, so this is a no-op then.
    salesmen = salesmen.filter(sm=> cabangMap.get(sm)===cabangFilter);
  }
  const thead = document.querySelector('#table-target-manage thead');
  const tbody = document.querySelector('#table-target-manage tbody');
  thead.innerHTML = '<tr><th>Salesman</th><th>Target Bulanan (Rp)</th></tr>';
  tbody.innerHTML = salesmen.length ? salesmen.map(sm=>{
    const val = targetMap.get(normName(sm)) || '';
    return `<tr><td class="text">${esc(sm)}</td><td><input type="number" min="0" step="100000" data-salesman="${esc(sm)}" value="${val}" style="width:160px; border:1px solid var(--line); border-radius:6px; padding:5px 8px; font-family:var(--font-mono);"></td></tr>`;
  }).join('') : (cabangFilter!=='all'
    ? '<tr><td colspan="2" style="text-align:center; color:var(--ink-faint); padding:16px;">Tidak ada salesman yang terkonfirmasi dari cabang '+esc(cabangFilter)+' di data Sales</td></tr>'
    : '<tr><td colspan="2" style="text-align:center; color:var(--ink-faint); padding:16px;">Upload data Sales dulu untuk melihat daftar salesman</td></tr>');
}

function renderEvaluasi(){
  const hasData = DB.sales.length>0;
  document.getElementById('evaluasi-empty').style.display = hasData? 'none':'block';
  document.getElementById('evaluasi-content').style.display = hasData? 'block':'none';
  if(!hasData) return;
  rebuildPreservingFocus(document.getElementById('evaluasi-filterbar'), renderEvaluasiFilterBar);
  if(!evaluasiFilterState.anchorYear) return;
  const cmp = evaluasiComparePeriod(evaluasiFilterState.anchorYear, evaluasiFilterState.anchorMonth, evaluasiFilterState.mode);
  const curRows = evaluasiRowsFor(evaluasiFilterState.anchorYear, evaluasiFilterState.anchorMonth, evaluasiFilterState.cabang);
  const cmpRows = evaluasiRowsFor(cmp.year, cmp.month, evaluasiFilterState.cabang);
  renderEvaluasiKPI(curRows, cmpRows);
  renderEvaluasiSalesmanTable(curRows, cmpRows);
  renderEvaluasiTopMovers(curRows, cmpRows);
  const ytdRows = DB.sales.filter(r=> r.year===evaluasiFilterState.anchorYear && r.month<=evaluasiFilterState.anchorMonth && (evaluasiFilterState.cabang==='all'||r.cabang===evaluasiFilterState.cabang));
  renderTargetSection(curRows, ytdRows);
  renderTargetTrendChart();
  renderTargetManageTable();
  document.getElementById('evaluasi-period-hint').textContent =
    BULAN_NAMES[evaluasiFilterState.anchorMonth-1]+' '+evaluasiFilterState.anchorYear+' vs '+BULAN_NAMES[cmp.month-1]+' '+cmp.year;
}

/* =========================================================================
   ANALISIS SILANG (CROSS-TAB / PIVOT)
========================================================================= */
const crosstabFilterState = { dim1:'salesman', dim2:'tipeOutlet', metric:'dpp', tahun:'all', bulan:'all', cabang:'all' };

const CROSSTAB_DIMS = {
  salesman:    { label:'Salesman',     get:r=>r.salesman },
  tipeOutlet:  { label:'Tipe Outlet',  get:r=>r.tipeOutlet },
  produk:      { label:'Produk',       get:r=>r.namaProdukGroup },
  cabang:      { label:'Cabang',       get:r=>r.cabang },
  principal:   { label:'Principal',    get:r=>r.principal }
};
const CROSSTAB_MAX_ROWS = 15, CROSSTAB_MAX_COLS = 10;

function crosstabScopeRows(){
  return DB.sales.filter(r=>{
    if(crosstabFilterState.tahun!=='all' && String(r.year)!==crosstabFilterState.tahun) return false;
    if(crosstabFilterState.bulan!=='all' && String(r.month)!==crosstabFilterState.bulan) return false;
    if(crosstabFilterState.cabang!=='all' && r.cabang!==crosstabFilterState.cabang) return false;
    return true;
  });
}
function crosstabVal(r){
  return crosstabFilterState.metric==='qty' ? r.qty : crosstabFilterState.metric==='netto' ? r.netto : r.dpp;
}
function crosstabMetricLabel(){
  return crosstabFilterState.metric==='qty' ? 'Qty' : crosstabFilterState.metric==='netto' ? 'Netto (+PPN)' : 'DPP';
}

function renderCrosstabFilterBar(){
  const bar = document.getElementById('crosstab-filterbar');
  const tahuns = [...new Set(DB.sales.map(r=>r.year))].sort((a,b)=>b-a);
  const bulans = [...new Set(DB.sales.map(r=>r.month))].sort((a,b)=>a-b);
  const cabangs = uniqueVals(DB.sales,'cabang');
  const dimOpts = (excludeKey)=> Object.entries(CROSSTAB_DIMS).filter(([k])=>k!==excludeKey).map(([k,d])=>'<option value="'+k+'">'+d.label+'</option>').join('');
  bar.innerHTML = `
    <div class="filter-field"><label>Baris</label><select id="f-ct-dim1">${dimOpts(crosstabFilterState.dim2)}</select></div>
    <div class="filter-field"><label>Kolom</label><select id="f-ct-dim2">${dimOpts(crosstabFilterState.dim1)}</select></div>
    <div class="filter-field"><label>Metrik</label><select id="f-ct-metric">
      <option value="dpp" ${crosstabFilterState.metric==='dpp'?'selected':''}>Nilai (DPP)</option>
      <option value="netto" ${crosstabFilterState.metric==='netto'?'selected':''}>Nilai (Netto +PPN)</option>
      <option value="qty" ${crosstabFilterState.metric==='qty'?'selected':''}>Qty</option>
    </select></div>
    <div class="filter-field"><label>Tahun</label><select id="f-ct-tahun"><option value="all">Semua</option>${tahuns.map(y=>'<option value="'+y+'" '+(crosstabFilterState.tahun===String(y)?'selected':'')+'>'+y+'</option>').join('')}</select></div>
    <div class="filter-field"><label>Bulan</label><select id="f-ct-bulan"><option value="all">Semua</option>${bulans.map(m=>'<option value="'+m+'" '+(crosstabFilterState.bulan===String(m)?'selected':'')+'>'+BULAN_NAMES[m-1]+'</option>').join('')}</select></div>
    <div class="filter-field"><label>Cabang</label><select id="f-ct-cabang"><option value="all">Semua</option>${cabangs.map(c=>'<option value="'+esc(c)+'" '+(crosstabFilterState.cabang===c?'selected':'')+'>'+esc(c)+'</option>').join('')}</select></div>
    <button class="btn-reset" id="f-ct-reset">Reset Filter</button>
  `;
  document.getElementById('f-ct-dim1').value = crosstabFilterState.dim1;
  document.getElementById('f-ct-dim2').value = crosstabFilterState.dim2;
  document.getElementById('f-ct-dim1').addEventListener('change', e=>{ crosstabFilterState.dim1=e.target.value; renderCrosstab(); });
  document.getElementById('f-ct-dim2').addEventListener('change', e=>{ crosstabFilterState.dim2=e.target.value; renderCrosstab(); });
  document.getElementById('f-ct-metric').addEventListener('change', e=>{ crosstabFilterState.metric=e.target.value; renderCrosstab(); });
  document.getElementById('f-ct-tahun').addEventListener('change', e=>{ crosstabFilterState.tahun=e.target.value; renderCrosstab(); });
  document.getElementById('f-ct-bulan').addEventListener('change', e=>{ crosstabFilterState.bulan=e.target.value; renderCrosstab(); });
  document.getElementById('f-ct-cabang').addEventListener('change', e=>{ crosstabFilterState.cabang=e.target.value; renderCrosstab(); });
  document.getElementById('f-ct-reset').addEventListener('click', ()=>{
    Object.assign(crosstabFilterState, {dim1:'salesman', dim2:'tipeOutlet', metric:'dpp', tahun:'all', bulan:'all', cabang:'all'});
    renderCrosstab();
  });
}

function crosstabTopValues(rows, dimKey, maxCount){
  const getFn = CROSSTAB_DIMS[dimKey].get;
  const totals = {};
  rows.forEach(r=>{
    const v = getFn(r) || '(kosong)';
    totals[v] = (totals[v]||0) + crosstabVal(r);
  });
  const sorted = Object.entries(totals).sort((a,b)=>b[1]-a[1]);
  if(sorted.length<=maxCount) return { values: sorted.map(s=>s[0]), overflow:false };
  return { values: sorted.slice(0,maxCount).map(s=>s[0]), overflow:true };
}

function crosstabCellRows(rows, dim1Key, dim1Val, dim1Overflow, dim1Values, dim2Key, dim2Val, dim2Overflow, dim2Values){
  const get1 = CROSSTAB_DIMS[dim1Key].get, get2 = CROSSTAB_DIMS[dim2Key].get;
  return rows.filter(r=>{
    const v1 = get1(r) || '(kosong)';
    const v2 = get2(r) || '(kosong)';
    const match1 = dim1Val==='__LAINNYA__' ? !dim1Values.includes(v1) : v1===dim1Val;
    const match2 = dim2Val==='__LAINNYA__' ? !dim2Values.includes(v2) : v2===dim2Val;
    return match1 && match2;
  });
}

function heatClass(v, max){
  if(!v || v<=0) return 'heat-0';
  const t = v/max;
  if(t<0.15) return 'heat-1';
  if(t<0.35) return 'heat-2';
  if(t<0.6) return 'heat-3';
  if(t<0.85) return 'heat-4';
  return 'heat-5';
}
function fmtCellVal(v){
  if(crosstabFilterState.metric==='qty') return fmtNumber(v);
  return v>=1e6 ? (v/1e6).toFixed(1)+'jt' : fmtNumber(v);
}

function renderCrosstab(){
  const hasData = DB.sales.length>0;
  document.getElementById('crosstab-empty').style.display = hasData? 'none':'block';
  document.getElementById('crosstab-content').style.display = hasData? 'block':'none';
  if(!hasData) return;
  renderCrosstabFilterBar();

  const rows = crosstabScopeRows();
  const { dim1, dim2 } = crosstabFilterState;
  const r1 = crosstabTopValues(rows, dim1, CROSSTAB_MAX_ROWS);
  const r2 = crosstabTopValues(rows, dim2, CROSSTAB_MAX_COLS);
  const rowValues = r1.values.concat(r1.overflow ? ['__LAINNYA__'] : []);
  const colValues = r2.values.concat(r2.overflow ? ['__LAINNYA__'] : []);

  const get1 = CROSSTAB_DIMS[dim1].get, get2 = CROSSTAB_DIMS[dim2].get;
  const matrix = {}; // rowVal -> colVal -> total
  const rowTotals = {}, colTotals = {}; let grandTotal = 0;
  rows.forEach(r=>{
    const v1raw = get1(r) || '(kosong)';
    const v2raw = get2(r) || '(kosong)';
    const v1 = r1.overflow && !r1.values.includes(v1raw) ? '__LAINNYA__' : v1raw;
    const v2 = r2.overflow && !r2.values.includes(v2raw) ? '__LAINNYA__' : v2raw;
    if(!rowValues.includes(v1) || !colValues.includes(v2)) return;
    matrix[v1] = matrix[v1] || {};
    matrix[v1][v2] = (matrix[v1][v2]||0) + crosstabVal(r);
    rowTotals[v1] = (rowTotals[v1]||0) + crosstabVal(r);
    colTotals[v2] = (colTotals[v2]||0) + crosstabVal(r);
    grandTotal += crosstabVal(r);
  });
  const maxCell = Math.max(1, ...rowValues.flatMap(rv=>colValues.map(cv=>(matrix[rv]&&matrix[rv][cv])||0)));

  const label = v => v==='__LAINNYA__' ? 'Lainnya' : v;
  let html = '<thead><tr><th>'+CROSSTAB_DIMS[dim1].label+' / '+CROSSTAB_DIMS[dim2].label+'</th>';
  colValues.forEach(cv=> html += '<th>'+esc(label(cv))+'</th>');
  html += '<th>Total</th></tr></thead><tbody>';
  rowValues.forEach(rv=>{
    html += '<tr><th>'+esc(label(rv))+'</th>';
    colValues.forEach(cv=>{
      const v = (matrix[rv] && matrix[rv][cv]) || 0;
      html += '<td class="cell '+heatClass(v,maxCell)+'" data-rv="'+esc(rv)+'" data-cv="'+esc(cv)+'">'+(v>0?fmtCellVal(v):'&ndash;')+'</td>';
    });
    html += '<td class="total-cell">'+fmtCellVal(rowTotals[rv]||0)+'</td></tr>';
  });
  html += '<tr class="total-row"><th>Total</th>';
  colValues.forEach(cv=> html += '<td class="total-cell">'+fmtCellVal(colTotals[cv]||0)+'</td>');
  html += '<td class="total-cell">'+fmtCellVal(grandTotal)+'</td></tr>';
  html += '</tbody>';
  document.getElementById('table-crosstab').innerHTML = html;

  document.querySelectorAll('#table-crosstab td.cell').forEach(td=>{
    td.addEventListener('click', ()=>{
      const rv = td.dataset.rv, cv = td.dataset.cv;
      const detail = crosstabCellRows(rows, dim1, rv, r1.overflow, r1.values, dim2, cv, r2.overflow, r2.values);
      const total = detail.reduce((s,r)=>s+crosstabVal(r),0);
      const bodyHtml = detail.length ? '<div class="table-wrap"><table><thead><tr><th>Tanggal</th><th>Invoice</th><th>Pelanggan</th><th>Produk</th><th>Salesman</th><th>'+crosstabMetricLabel()+'</th></tr></thead><tbody>'+
        detail.slice().sort((a,b)=>b.tanggal-a.tanggal).slice(0,200).map(r=>`<tr><td>${fmtDate(r.tanggal)}</td><td>${esc(r.invoice)}</td><td class="text">${esc(r.namaPelanggan)}</td><td class="text">${esc(r.namaProdukGroup)}</td><td class="text">${esc(r.salesman)}</td><td>${fmtCellVal(crosstabVal(r))}</td></tr>`).join('')+
        '</tbody></table></div><p style="margin-top:12px; font-weight:600;">Total '+crosstabMetricLabel()+': '+fmtCellVal(total)+' ('+detail.length+' baris'+(detail.length>200?', menampilkan 200 pertama':'')+')</p>'
        : '<p style="color:var(--ink-soft); font-size:13px;">Tidak ada transaksi untuk kombinasi ini.</p>';
      openDrawer('Detail: '+esc(label(rv))+' &times; '+esc(label(cv)), bodyHtml, true);
    });
  });

  document.getElementById('crosstab-hint').textContent = 'Basis: '+crosstabMetricLabel()+' · klik sel untuk lihat detail transaksinya';
  renderCustProdTable();
}

/* ---- Riwayat Produk per Pelanggan ---- */
const custProdFilterState = { salesman:'all', cabang:'all', tahun:'all', bulan:'all', q:'' };
let custProdTableCtrl = null;

function custProdRowsExcept(exceptKey){
  return DB.sales.filter(r=>{
    if(exceptKey!=='salesman' && custProdFilterState.salesman!=='all' && r.salesman!==custProdFilterState.salesman) return false;
    if(exceptKey!=='cabang' && custProdFilterState.cabang!=='all' && r.cabang!==custProdFilterState.cabang) return false;
    if(exceptKey!=='tahun' && custProdFilterState.tahun!=='all' && String(r.year)!==custProdFilterState.tahun) return false;
    if(exceptKey!=='bulan' && custProdFilterState.bulan!=='all' && String(r.month)!==custProdFilterState.bulan) return false;
    return true;
  });
}
function custProdFilteredRows(){ return custProdRowsExcept(null); }

function renderCustProdFilterBar(){
  const bar = document.getElementById('custprod-filterbar');
  const salesmen = uniqueVals(custProdRowsExcept('salesman'),'salesman');
  const cabangs = uniqueVals(custProdRowsExcept('cabang'),'cabang');
  const tahuns = [...new Set(custProdRowsExcept('tahun').map(r=>r.year))].sort((a,b)=>b-a);
  const bulans = [...new Set(custProdRowsExcept('bulan').map(r=>r.month))].sort((a,b)=>a-b);
  if(custProdFilterState.salesman!=='all' && !salesmen.includes(custProdFilterState.salesman)) custProdFilterState.salesman='all';
  if(custProdFilterState.cabang!=='all' && !cabangs.includes(custProdFilterState.cabang)) custProdFilterState.cabang='all';
  if(custProdFilterState.tahun!=='all' && !tahuns.map(String).includes(custProdFilterState.tahun)) custProdFilterState.tahun='all';
  if(custProdFilterState.bulan!=='all' && !bulans.map(String).includes(custProdFilterState.bulan)) custProdFilterState.bulan='all';
  bar.innerHTML = `
    <div class="filter-field"><label>Salesman</label><select id="f-cp-salesman"><option value="all">Semua</option>${salesmen.map(s=>'<option value="'+esc(s)+'" '+(custProdFilterState.salesman===s?'selected':'')+'>'+esc(s)+'</option>').join('')}</select></div>
    <div class="filter-field"><label>Cabang</label><select id="f-cp-cabang"><option value="all">Semua</option>${cabangs.map(c=>'<option value="'+esc(c)+'" '+(custProdFilterState.cabang===c?'selected':'')+'>'+esc(c)+'</option>').join('')}</select></div>
    <div class="filter-field"><label>Tahun</label><select id="f-cp-tahun"><option value="all">Semua</option>${tahuns.map(y=>'<option value="'+y+'" '+(custProdFilterState.tahun===String(y)?'selected':'')+'>'+y+'</option>').join('')}</select></div>
    <div class="filter-field"><label>Bulan</label><select id="f-cp-bulan"><option value="all">Semua</option>${bulans.map(m=>'<option value="'+m+'" '+(custProdFilterState.bulan===String(m)?'selected':'')+'>'+BULAN_NAMES[m-1]+'</option>').join('')}</select></div>
    <div class="filter-field"><label>Cari Pelanggan</label><input type="search" id="f-cp-q" placeholder="Nama pelanggan" value="${esc(custProdFilterState.q)}"></div>
    <button class="btn-reset" id="f-cp-reset">Reset Filter</button>
  `;
  document.getElementById('f-cp-salesman').addEventListener('change', e=>{ custProdFilterState.salesman=e.target.value; renderCustProdTable(); });
  document.getElementById('f-cp-cabang').addEventListener('change', e=>{ custProdFilterState.cabang=e.target.value; renderCustProdTable(); });
  document.getElementById('f-cp-tahun').addEventListener('change', e=>{ custProdFilterState.tahun=e.target.value; renderCustProdTable(); });
  document.getElementById('f-cp-bulan').addEventListener('change', e=>{ custProdFilterState.bulan=e.target.value; renderCustProdTable(); });
  document.getElementById('f-cp-q').addEventListener('input', e=>{ custProdFilterState.q=e.target.value; renderCustProdTable(); });
  document.getElementById('f-cp-reset').addEventListener('click', ()=>{
    Object.assign(custProdFilterState, {salesman:'all', cabang:'all', tahun:'all', bulan:'all', q:''});
    renderCustProdTable();
  });
}

function custProdAggregate(){
  const rows = custProdFilteredRows();
  const map = {};
  rows.forEach(r=>{
    const key = normName(r.namaPelanggan);
    if(!map[key]) map[key] = { namaPelanggan:r.namaPelanggan, salesmen:new Set(), products:{}, totalValue:0, totalQty:0, lastPurchase:r.tanggal };
    const c = map[key];
    c.salesmen.add(r.salesman);
    if(!c.products[r.namaProdukGroup]) c.products[r.namaProdukGroup] = { namaProdukGroup:r.namaProdukGroup, qty:0, value:0 };
    c.products[r.namaProdukGroup].qty += r.qty;
    c.products[r.namaProdukGroup].value += r.dpp;
    c.totalValue += r.dpp; c.totalQty += r.qty;
    if(r.tanggal > c.lastPurchase) c.lastPurchase = r.tanggal;
  });
  let list = Object.values(map).map(c=>{
    const prodList = Object.values(c.products).sort((a,b)=>b.value-a.value);
    return { ...c, salesmenLabel: [...c.salesmen].join(', '), prodList, prodCount: prodList.length };
  });
  if(custProdFilterState.q){
    const q = custProdFilterState.q.toLowerCase();
    list = list.filter(c=>c.namaPelanggan.toLowerCase().includes(q));
  }
  return list;
}

function custProdDetailHtml(c){
  return '<div class="table-wrap"><table><thead><tr><th>Produk</th><th>Qty</th><th>Nilai (DPP)</th></tr></thead><tbody>'+
    c.prodList.map(p=>`<tr><td class="text">${esc(p.namaProdukGroup)}</td><td>${fmtNumber(p.qty)}</td><td>${fmtIDR(p.value)}</td></tr>`).join('')+
    '</tbody></table></div><p style="margin-top:12px; font-weight:600;">Total: '+fmtIDR(c.totalValue)+' &middot; '+c.prodCount+' produk berbeda</p>';
}

function renderCustProdTable(){
  const hasData = DB.sales.length>0;
  if(!hasData) return;
  renderCustProdFilterBar();
  if(!custProdTableCtrl){
    custProdTableCtrl = makeTableController({
      getRows: ()=> custProdAggregate(),
      columns:[
        {key:'namaPelanggan', label:'Pelanggan'}, {key:'salesmenLabel', label:'Sales'},
        {key:'prodCount', label:'Jumlah Produk Berbeda'}, {key:'prodPreview', label:'Produk yang Pernah Dibeli'},
        {key:'totalQty', label:'Total Qty'}, {key:'totalValue', label:'Total Pembelian'}, {key:'lastPurchase', label:'Terakhir Beli'}
      ],
      renderRow: c=>{
        const preview = c.prodList.slice(0,3).map(p=>p.namaProdukGroup).join(', ');
        const more = c.prodCount>3 ? ' <span class="link-cell" data-cust="'+esc(c.namaPelanggan)+'">+'+(c.prodCount-3)+' lainnya</span>' : '';
        return `<tr><td class="text customer-link" data-customer="${esc(c.namaPelanggan)}">${esc(c.namaPelanggan)}</td><td class="text">${esc(c.salesmenLabel)}</td><td>${c.prodCount}</td><td class="text" style="white-space:normal; min-width:260px;">${esc(preview)}${more}</td><td>${fmtNumber(c.totalQty)}</td><td>${fmtIDR(c.totalValue)}</td><td>${fmtDate(c.lastPurchase)}</td></tr>`;
      },
      afterRender: (tbody, pageRows)=>{
        bindCustomerLinks(tbody);
        tbody.querySelectorAll('.link-cell[data-cust]').forEach(el=>{
          el.addEventListener('click', (ev)=>{
            ev.stopPropagation();
            const c = pageRows.find(x=>x.namaPelanggan===el.dataset.cust);
            if(c) openDrawer('Produk yang Pernah Dibeli: '+esc(c.namaPelanggan), custProdDetailHtml(c));
          });
        });
      },
      pageSize:20, tbodyId:'table-custprod-tbody', theadId:'table-custprod-thead', pagerId:'custprod-pager', infoId:'custprod-pager-info',
      defaultSort:{key:'totalValue', dir:'desc'}
    });
    document.querySelector('#table-custprod').innerHTML = '<thead id="table-custprod-thead"></thead><tbody id="table-custprod-tbody"></tbody>';
  }
  custProdTableCtrl.resetPage();
  custProdTableCtrl.render();
  document.getElementById('custprod-count').textContent = custProdAggregate().length + ' pelanggan';
}

/* =========================================================================
   MASTER RENDER
========================================================================= */
function renderAll(){
  renderHome();
  renderSales();
  renderAging();
  renderPayment();
  renderStock();
  renderRetention();
  renderEvaluasi();
  renderCrosstab();
}

/* =========================================================================
   INIT
========================================================================= */
function setupExportButtons(){
  const dateTag = ()=> fmtDate(new Date()).replace(/\//g,'-');

  // ---- Sales ----
  const salesColumns = [
    {label:'Tanggal', value:r=>fmtDate(r.tanggal)},
    {label:'Invoice', key:'invoice'},
    {label:'Cabang', key:'cabang'},
    {label:'Pelanggan', key:'namaPelanggan'},
    {label:'Tipe Outlet', key:'tipeOutlet'},
    {label:'Produk', value:r=>r.namaProdukGroup},
    {label:'Nama Produk Asli', key:'namaProduk'},
    {label:'Principal', key:'principal'},
    {label:'Salesman', key:'salesman'},
    {label:'Qty', key:'qty'},
    {label:'DPP', key:'dpp'},
    {label:'Netto (+PPN)', key:'netto'}
  ];
  document.getElementById('btn-export-sales-xlsx').addEventListener('click', ()=>{
    exportToExcel('Sales_Export_'+dateTag()+'.xlsx', salesFilteredRows(), salesColumns);
  });
  document.getElementById('btn-export-sales-pdf').addEventListener('click', ()=>{
    const s = salesFilterState;
    exportToPDF('Laporan Data Sales', activeFilterSummary(s, {cabang:'Cabang', tahun:'Tahun', bulan:'Bulan', salesman:'Salesman', principal:'Principal', outlet:'Tipe Outlet', q:'Cari'}), salesFilteredRows(), salesColumns);
  });

  // ---- AR Aging ----
  const agingColumns = [
    {label:'Invoice', key:'invoice'},
    {label:'Pelanggan', key:'namaPelanggan'},
    {label:'Cabang', value:r=>agingRowCabang(r) || '-'},
    {label:'Tgl Invoice', value:r=>fmtDate(r.tglInvoice)},
    {label:'Tgl Jatuh Tempo', value:r=>fmtDate(r.tglJatuhTempo)},
    {label:'Umur (hari)', key:'agingDays'},
    {label:'Status', key:'bucket'},
    {label:'Saldo', key:'totalSaldo'}
  ];
  document.getElementById('btn-export-aging-xlsx').addEventListener('click', ()=>{
    exportToExcel('AR_Aging_Export_'+dateTag()+'.xlsx', agingFilteredRows(), agingColumns);
  });
  document.getElementById('btn-export-aging-pdf').addEventListener('click', ()=>{
    const s = agingFilterState;
    exportToPDF('Laporan AR Aging', activeFilterSummary(s, {cabang:'Cabang', bucket:'Status Umur', q:'Cari'}), agingFilteredRows(), agingColumns);
  });

  // ---- AR Payment ----
  const paymentColumns = [
    {label:'Invoice', key:'invoice'},
    {label:'Pelanggan', key:'namaPelanggan'},
    {label:'Sales', key:'sales'},
    {label:'Nilai Invoice', key:'nilaiInvoice'},
    {label:'Tunai', key:'pembTunai'},
    {label:'Deposit', key:'pembDeposit'},
    {label:'Cek', key:'pembCek'},
    {label:'Tgl Bayar Terakhir', value:r=>fmtDate(r.tglPembTerakhir)},
    {label:'Sisa', key:'sisaPemb'},
    {label:'Status', value:r=>r.lunas?'Lunas':'Sisa'}
  ];
  document.getElementById('btn-export-payment-xlsx').addEventListener('click', ()=>{
    exportToExcel('AR_Payment_Export_'+dateTag()+'.xlsx', paymentFilteredRows(), paymentColumns);
  });
  document.getElementById('btn-export-payment-pdf').addEventListener('click', ()=>{
    const s = paymentFilterState;
    exportToPDF('Laporan AR Payment', activeFilterSummary(s, {sales:'Sales', status:'Status', q:'Cari'}), paymentFilteredRows(), paymentColumns);
  });

  // ---- Stock ----
  const stockColumns = [
    {label:'Produk', key:'namaProduk'},
    {label:'Lokasi', key:'locationId'},
    {label:'Lot', key:'idLot'},
    {label:'Tgl Kadaluwarsa', value:r=>fmtDate(r.tglKadaluwarsa)},
    {label:'Satuan', key:'satuan'},
    {label:'Stock Akhir', key:'stokAkhir'}
  ];
  document.getElementById('btn-export-stock-xlsx').addEventListener('click', ()=>{
    exportToExcel('Stock_Export_'+dateTag()+'.xlsx', stockFilteredRows(), stockColumns);
  });
  document.getElementById('btn-export-stock-pdf').addEventListener('click', ()=>{
    const s = stockFilterState;
    exportToPDF('Laporan Stock Produk', activeFilterSummary(s, {lokasi:'Lokasi/Gudang', status:'Status', q:'Cari Produk'}), stockFilteredRows(), stockColumns);
  });

  // ---- Retention ----
  const retentionColumns = [
    {label:'Pelanggan', key:'namaPelanggan'},
    {label:'Cabang', key:'cabang'},
    {label:'Tipe Outlet', key:'tipeOutlet'},
    {label:'Wilayah', key:'wilayah'},
    {label:'Alamat', key:'alamat'},
    {label:'Kontak', key:'kontakPerson'},
    {label:'Telp', key:'telp'},
    {label:'Terakhir Beli', value:r=>fmtDate(r.lastPurchase)},
    {label:'Hari Tidak Beli', key:'daysSince'},
    {label:'Total Pembelian Historis', key:'totalValue'},
    {label:'Sales Terakhir', key:'lastSalesman'}
  ];
  function retentionInactiveRows(){
    const maxDate = DB.sales.length ? new Date(Math.max(...DB.sales.map(r=>r.tanggal.getTime()))) : TODAY;
    const threshold = parseInt(retentionFilterState.thresholdDays,10);
    let stats = retentionCustomerStats();
    if(retentionFilterState.cabang!=='all') stats = stats.filter(s=>s.cabang===retentionFilterState.cabang);
    if(retentionFilterState.tipeOutlet!=='all') stats = stats.filter(s=>s.tipeOutlet===retentionFilterState.tipeOutlet);
    if(retentionFilterState.wilayah!=='all') stats = stats.filter(s=>s.wilayah===retentionFilterState.wilayah);
    stats.forEach(s=>{ s.daysSince = daysBetween(maxDate, s.lastPurchase); });
    return stats.filter(s=>s.daysSince>=threshold).sort((a,b)=>b.totalValue-a.totalValue);
  }
  document.getElementById('btn-export-retention-xlsx').addEventListener('click', ()=>{
    exportToExcel('Retensi_Pelanggan_Export_'+dateTag()+'.xlsx', retentionInactiveRows(), retentionColumns);
  });
  document.getElementById('btn-export-retention-pdf').addEventListener('click', ()=>{
    const s = retentionFilterState;
    exportToPDF('Laporan Retensi Pelanggan', activeFilterSummary(s, {cabang:'Cabang', tipeOutlet:'Tipe Outlet', wilayah:'Wilayah', thresholdDays:'Ambang Tidak Aktif (hari)'}), retentionInactiveRows(), retentionColumns);
  });
}

// PWA: register service worker (only works when served over http(s), e.g. via api.php
// hosting — silently does nothing when this file is just opened directly/locally).
if('serviceWorker' in navigator && (location.protocol==='http:' || location.protocol==='https:')){
  window.addEventListener('load', ()=>{
    navigator.serviceWorker.register('./sw.js').catch(()=>{ /* not hosted, ignore */ });
  });
}

(async function init(){
  await detectServerMode();
  updateStorageModeIndicator();
  setupCustomer360Search();
  setupExportButtons();
  document.getElementById('btn-save-targets').addEventListener('click', async ()=>{
    const inputs = document.querySelectorAll('#table-target-manage input[data-salesman]');
    const records = [];
    inputs.forEach(inp=>{
      const val = parseFloat(inp.value);
      if(!isNaN(val) && val>0) records.push({salesman: inp.dataset.salesman, targetBulanan: val});
    });
    await saveTargets(records, 'Input Manual');
    renderEvaluasi();
  });
  document.getElementById('targetManageToggle').addEventListener('click', ()=>{
    const panel = document.getElementById('targetManagePanel');
    const toggle = document.getElementById('targetManageToggle');
    const isOpen = panel.style.display !== 'none';
    panel.style.display = isOpen ? 'none' : 'block';
    toggle.classList.toggle('open', !isOpen);
  });
  document.getElementById('btnResetAll').addEventListener('click', removeAllData);
  document.getElementById('btnExportWithData').addEventListener('click', exportDashboardWithData);
  await loadAllFromStorage();
  let anyData = DB.sales.length || DB.aging.rows.length || DB.payment.rows.length || DB.stock.rows.length || DB.targets.rows.length || DB.customers.rows.length;
  if(!anyData){
    const seed = loadEmbeddedSeed();
    if(seed){
      await hydrateFromSeed(seed);
      await loadAllFromStorage();
      anyData = DB.sales.length || DB.aging.rows.length || DB.payment.rows.length || DB.stock.rows.length || DB.targets.rows.length || DB.customers.rows.length;
    }
  }
  renderAll();
  renderDataStatusList();
  if(!anyData) openUploadModal();
})();
</script>
</body>
</html>
