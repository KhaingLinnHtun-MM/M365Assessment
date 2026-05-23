#============================================================================== 
# M365 Assessment Dashboard Generator (ENTERPRISE v4)
# Full enterprise UI: Side navigation + filter drawer + slide detail panel + charts + CSV/XLSX exports
# Self-contained HTML. Uses local SheetJS (xlsx.full.min.js) for XLSX export.
#
# Usage:
#   Unblock-File .\M365_Assessment_Dashboard_Generator_ENTERPRISE_v4.ps1
#   .\M365_Assessment_Dashboard_Generator_ENTERPRISE_v4.ps1 -OutputRoot "C:\Users\<you>\Desktop\M365_Assessment_Consultant_YYYYMMDD_HHMMSS"
#==============================================================================

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$OutputRoot,

    # Embedded data limits (0 = all)
    [int]$MaxFindings = 3000,
    [int]$MaxRoadmap  = 6000,
    [int]$MaxInventory = 8000,

    # SheetJS download URL (authoritative CDN)
    [string]$SheetJsUrl = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js'
)

function Write-Ok($m){Write-Host "[OK]   $m" -ForegroundColor Green}
function Write-Warn($m){Write-Host "[WARN] $m" -ForegroundColor Yellow}
function Write-Err($m){Write-Host "[ERR]  $m" -ForegroundColor Red}

if (-not (Test-Path $OutputRoot)) {
    Write-Err "OutputRoot not found: $OutputRoot"
    exit 1
}

function Load-CsvIfExists {
    param([string]$Path)
    if (Test-Path $Path) { return Import-Csv $Path }
    return @()
}

# Ensure SheetJS library exists locally (for XLSX export)
$sheetJsLocal = Join-Path $OutputRoot 'xlsx.full.min.js'
if (-not (Test-Path $sheetJsLocal)) {
    Write-Warn "SheetJS library not found locally. Downloading to: $sheetJsLocal"
    try {
        Invoke-WebRequest -Uri $SheetJsUrl -UseBasicParsing -OutFile $sheetJsLocal -ErrorAction Stop
        Write-Ok "Downloaded SheetJS: $SheetJsUrl"
    } catch {
        Write-Warn "Failed to download SheetJS. XLSX export will be disabled in dashboard. Error: $($_)"
    }
} else {
    Write-Ok "Found local SheetJS: $sheetJsLocal"
}

# Primary outputs
$scorecard = Load-CsvIfExists (Join-Path $OutputRoot '00_Scorecard.csv')
$findings  = Load-CsvIfExists (Join-Path $OutputRoot '41_Findings_RiskRegister.csv')
$roadmap   = Load-CsvIfExists (Join-Path $OutputRoot '40_Roadmap_30_60_90.csv')

# Optional inventory outputs
$licenses  = Load-CsvIfExists (Join-Path $OutputRoot '05_License_Inventory.csv')
$mailbox   = Load-CsvIfExists (Join-Path $OutputRoot '02_Mailbox_Usage.csv')
$onedrive  = Load-CsvIfExists (Join-Path $OutputRoot '03_OneDrive_Usage.csv')
$groups    = Load-CsvIfExists (Join-Path $OutputRoot '09_All_Groups.csv')
$ssRecs    = Load-CsvIfExists (Join-Path $OutputRoot '30_Top_SecureScore_Recommendations.csv')

# Clamp row counts
if ($MaxFindings -gt 0 -and $findings.Count -gt $MaxFindings) { $findings = $findings | Select-Object -First $MaxFindings }
if ($MaxRoadmap  -gt 0 -and $roadmap.Count  -gt $MaxRoadmap)  { $roadmap  = $roadmap  | Select-Object -First $MaxRoadmap }
if ($MaxInventory -gt 0) {
    if ($licenses.Count -gt $MaxInventory) { $licenses = $licenses | Select-Object -First $MaxInventory }
    if ($mailbox.Count  -gt $MaxInventory) { $mailbox  = $mailbox  | Select-Object -First $MaxInventory }
    if ($onedrive.Count -gt $MaxInventory) { $onedrive = $onedrive | Select-Object -First $MaxInventory }
    if ($groups.Count   -gt $MaxInventory) { $groups   = $groups   | Select-Object -First $MaxInventory }
    if ($ssRecs.Count   -gt $MaxInventory) { $ssRecs   = $ssRecs   | Select-Object -First $MaxInventory }
}

# Convert datasets to JSON (compressed)
$scoreJson    = ($scorecard | ConvertTo-Json -Depth 6 -Compress)
$findingsJson = ($findings  | ConvertTo-Json -Depth 6 -Compress)
$roadmapJson  = ($roadmap   | ConvertTo-Json -Depth 6 -Compress)
$licensesJson = ($licenses  | ConvertTo-Json -Depth 6 -Compress)
$mailboxJson  = ($mailbox   | ConvertTo-Json -Depth 6 -Compress)
$onedriveJson = ($onedrive  | ConvertTo-Json -Depth 6 -Compress)
$groupsJson   = ($groups    | ConvertTo-Json -Depth 6 -Compress)
$recsJson     = ($ssRecs    | ConvertTo-Json -Depth 6 -Compress)

$template = @'
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1"/>
  <title>M365 Assessment Dashboard (Enterprise)</title>
  <style>
    /* ===== Fluent-style Tokens ===== */
    :root{
      --bg:#f6f7fb;
      --card:#ffffff;
      --border:rgba(0,0,0,.08);
      --text:#1f1f1f;
      --muted:rgba(0,0,0,.65);
      --brand:#0f6cbd;
      --brandWeak:rgba(15,108,189,.12);
      --critical:#d13438;
      --high:#ca5010;
      --medium:#8a6d3b;
      --low:#107c10;
      --r:14px;
      --shadow:0 1px 2px rgba(0,0,0,.06);
      --shadow2:0 12px 30px rgba(0,0,0,.12);
      --sidebar:#0b1220;
      --sidebar2:#0f172a;
      --sidebarText:rgba(255,255,255,.88);
      --sidebarMuted:rgba(255,255,255,.65);
      --focus:0 0 0 3px rgba(15,108,189,.25);
      --font: Segoe UI, system-ui, -apple-system, Arial, sans-serif;
    }

    [data-theme="dark"]{
      --bg:#0b0f17;
      --card:#121826;
      --border:rgba(255,255,255,.08);
      --text:#f3f4f6;
      --muted:rgba(255,255,255,.7);
      --brand:#60a5fa;
      --brandWeak:rgba(96,165,250,.18);
      --sidebar:#0b0f17;
      --sidebar2:#111827;
      --sidebarText:rgba(255,255,255,.9);
      --sidebarMuted:rgba(255,255,255,.7);
      --shadow:0 1px 2px rgba(0,0,0,.30);
      --shadow2:0 18px 46px rgba(0,0,0,.45);
      --focus:0 0 0 3px rgba(96,165,250,.25);
    }

    *{box-sizing:border-box;}
    body{margin:0;font-family:var(--font);background:
      radial-gradient(900px 500px at 15% 5%, rgba(15,108,189,.14), transparent 60%),
      radial-gradient(900px 500px at 85% 10%, rgba(120,84,255,.10), transparent 55%),
      var(--bg);
      color:var(--text);
    }

    /* ===== App Shell ===== */
    .app{display:flex;height:100vh;}

    /* Sidebar */
    .sidebar{width:260px;background:linear-gradient(180deg,var(--sidebar),var(--sidebar2));color:var(--sidebarText);display:flex;flex-direction:column;padding:16px;gap:12px;}
    .brand{display:flex;align-items:center;gap:10px;padding:10px 10px 6px 10px;}
    .brandDot{width:10px;height:10px;border-radius:50%;background:var(--brand);box-shadow:0 0 0 3px rgba(255,255,255,.08);} 
    .brandTitle{font-weight:800;font-size:14px;letter-spacing:.2px;}
    .brandSub{font-size:11px;color:var(--sidebarMuted);margin-top:2px;}

    .nav{display:flex;flex-direction:column;gap:6px;margin-top:8px;}
    .navBtn{display:flex;align-items:center;gap:10px;padding:10px 12px;border-radius:10px;cursor:pointer;font-size:13px;color:var(--sidebarText);border:1px solid transparent;}
    .navBtn:hover{background:rgba(255,255,255,.06);} 
    .navBtn.active{background:rgba(15,108,189,.20);border-color:rgba(15,108,189,.35);} 
    .navIcon{width:18px;text-align:center;opacity:.95}

    .sideFooter{margin-top:auto;display:flex;flex-direction:column;gap:8px;padding-top:10px;border-top:1px solid rgba(255,255,255,.08);} 
    .sideMeta{font-size:11px;color:var(--sidebarMuted);line-height:1.4}

    /* Main */
    .main{flex:1;display:flex;flex-direction:column;min-width:0;}
    .topbar{height:56px;display:flex;align-items:center;gap:12px;padding:0 16px;border-bottom:1px solid var(--border);
      background:rgba(255,255,255,.72);backdrop-filter: blur(10px);
      position:sticky;top:0;z-index:20;
    }
    [data-theme="dark"] .topbar{background:rgba(18,24,38,.70);} 

    .topTitle{font-weight:800;font-size:14px;color:var(--text);} 
    .topSub{font-size:12px;color:var(--muted);} 

    .topGrow{flex:1;}
    .btn{padding:8px 12px;border-radius:10px;border:1px solid var(--border);background:var(--card);cursor:pointer;font-weight:700;font-size:12px;color:var(--text);}
    .btn:hover{box-shadow:var(--shadow);} 
    .btn.primary{background:var(--brand);color:#fff;border-color:transparent;}
    .btn:focus{outline:none;box-shadow:var(--focus);} 

    .searchWrap{display:flex;align-items:center;gap:8px;background:var(--card);border:1px solid var(--border);border-radius:999px;padding:6px 10px;min-width:260px;}
    .searchWrap input{border:none;outline:none;width:100%;font-size:13px;background:transparent;color:var(--text);} 
    .searchIcon{opacity:.6}

    .content{padding:16px;overflow:auto;min-height:0;}

    /* Cards + grids */
    .grid{display:grid;grid-template-columns:repeat(12,1fr);gap:12px;}
    .card{background:var(--card);border:1px solid var(--border);border-radius:var(--r);padding:14px;box-shadow:var(--shadow);} 
    .card:hover{box-shadow:var(--shadow2);transform:translateY(-1px);transition:.15s ease;} 

    .span12{grid-column:span 12;} .span6{grid-column:span 6;} .span4{grid-column:span 4;} .span3{grid-column:span 3;} .span8{grid-column:span 8;}

    .kpiLabel{font-size:12px;color:var(--muted);} 
    .kpiValue{font-size:30px;font-weight:900;color:var(--brand);margin-top:6px;} 
    .kpiHint{font-size:11px;color:var(--muted);} 

    .badge{display:inline-flex;align-items:center;gap:6px;padding:4px 10px;border-radius:999px;font-weight:800;font-size:12px;border:1px solid var(--border);} 
    .bCritical{background:#fde7e9;color:var(--critical);border-color:#f7c3c8;} 
    .bHigh{background:#fff4e5;color:var(--high);border-color:#ffd1a8;} 
    .bMedium{background:#fff9db;color:var(--medium);border-color:#ffe9a5;} 
    .bLow{background:#e8f5e9;color:var(--low);border-color:#bfe7c1;} 
    .bInfo{background:rgba(15,108,189,.08);color:var(--brand);border-color:rgba(15,108,189,.20);} 

    /* Tables */
    .tablewrap{border:1px solid var(--border);border-radius:12px;overflow:auto;max-height:64vh;}
    table{width:100%;border-collapse:collapse;font-size:12px;}
    thead th{position:sticky;top:0;background:rgba(243,244,247,.95);backdrop-filter: blur(8px);text-align:left;font-weight:800;border-bottom:1px solid var(--border);padding:10px;white-space:nowrap;}
    [data-theme="dark"] thead th{background:rgba(17,24,39,.92);} 
    tbody td{border-bottom:1px solid var(--border);padding:10px;vertical-align:top;}
    tbody tr:nth-child(even) td{background:rgba(0,0,0,.02);} 
    [data-theme="dark"] tbody tr:nth-child(even) td{background:rgba(255,255,255,.03);} 
    tbody tr:hover td{background:var(--brandWeak);} 

    .rowClickable{cursor:pointer;}

    /* Pages */
    .page{display:none;}
    .page.active{display:block;}

    /* Charts */
    .chartWrap{display:flex;flex-direction:column;gap:8px;}
    canvas{width:100%;height:260px;border-radius:12px;}

    /* Filter Drawer */
    .drawer{position:fixed;top:0;right:-420px;width:420px;height:100vh;background:var(--card);border-left:1px solid var(--border);
      box-shadow:var(--shadow2);z-index:60;transition:right .18s ease;display:flex;flex-direction:column;}
    .drawer.open{right:0;}
    .drawerHeader{padding:14px 14px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;}
    .drawerTitle{font-weight:900;font-size:13px;}
    .drawerBody{padding:14px;overflow:auto;display:flex;flex-direction:column;gap:12px;}
    .field label{display:block;font-size:12px;color:var(--muted);margin-bottom:6px;}
    .field input,.field select{width:100%;padding:9px 10px;border-radius:10px;border:1px solid var(--border);font-size:13px;background:transparent;color:var(--text);} 
    .drawerFooter{padding:14px;border-top:1px solid var(--border);display:flex;gap:8px;}

    /* Detail Panel */
    .panel{position:fixed;top:0;right:-520px;width:520px;height:100vh;background:var(--card);border-left:1px solid var(--border);
      box-shadow:var(--shadow2);z-index:70;transition:right .18s ease;display:flex;flex-direction:column;}
    .panel.open{right:0;}
    .panelHeader{padding:14px;border-bottom:1px solid var(--border);display:flex;align-items:flex-start;gap:10px;justify-content:space-between;}
    .panelTitle{font-weight:950;font-size:14px;line-height:1.2;}
    .panelBody{padding:14px;overflow:auto;display:flex;flex-direction:column;gap:10px;}
    .panelSection{border:1px solid var(--border);border-radius:12px;padding:12px;background:rgba(0,0,0,.01);} 
    [data-theme="dark"] .panelSection{background:rgba(255,255,255,.03);} 
    .kv{display:grid;grid-template-columns:110px 1fr;gap:8px;font-size:12px;}
    .kv b{color:var(--muted);font-weight:800;}

    /* Overlay */
    .overlay{position:fixed;inset:0;background:rgba(0,0,0,.28);backdrop-filter: blur(1px);z-index:55;display:none;}
    .overlay.show{display:block;}

    /* Responsive */
    @media (max-width:1100px){
      .sidebar{display:none;}
      .span6,.span4,.span3,.span8{grid-column:span 12;}
      .searchWrap{min-width:160px;}
      .drawer{width:100%;right:-100%;}
      .panel{width:100%;right:-100%;}
    }
  </style>

  <script src="xlsx.full.min.js"></script>
</head>
<body>
<div class="app" id="appRoot">

  <aside class="sidebar">
    <div class="brand">
      <div class="brandDot"></div>
      <div>
        <div class="brandTitle">M365 Assessment</div>
        <div class="brandSub">Enterprise Dashboard</div>
      </div>
    </div>

    <div class="nav">
      <div class="navBtn active" data-page="pageDashboard" onclick="goPage('pageDashboard', this)"><span class="navIcon">📊</span>Dashboard</div>
      <div class="navBtn" data-page="pageFindings" onclick="goPage('pageFindings', this)"><span class="navIcon">⚠️</span>Findings</div>
      <div class="navBtn" data-page="pageRoadmap" onclick="goPage('pageRoadmap', this)"><span class="navIcon">🗺️</span>Roadmap</div>
      <div class="navBtn" data-page="pageInventory" onclick="goPage('pageInventory', this)"><span class="navIcon">📦</span>Inventory</div>
      <div class="navBtn" data-page="pageScorecard" onclick="goPage('pageScorecard', this)"><span class="navIcon">📈</span>Scorecard</div>
      <div class="navBtn" data-page="pageFiles" onclick="goPage('pageFiles', this)"><span class="navIcon">📁</span>Files</div>
    </div>

    <div class="sideFooter">
      <div class="sideMeta">Tip: Click a finding row to open the detail panel.<br/>Use Filters to narrow results and export XLSX.</div>
      <button class="btn" onclick="toggleTheme()">🌙 Theme</button>
    </div>
  </aside>

  <main class="main">
    <div class="topbar">
      <div>
        <div class="topTitle">M365 Assessment Dashboard</div>
        <div class="topSub">Output: <b>__OUTPUT_ROOT__</b> &nbsp; • &nbsp; Generated: <b>__GENERATED__</b></div>
      </div>
      <div class="topGrow"></div>

      <div class="searchWrap" title="Search Findings">
        <span class="searchIcon">🔎</span>
        <input id="globalSearch" placeholder="Search findings..." oninput="quickSearch()" />
      </div>

      <button class="btn" onclick="openDrawer()">🎛️ Filters</button>
      <button class="btn primary" onclick="exportAllWorkbook()">⬇ Export ALL XLSX</button>
    </div>

    <div class="content">

      <!-- Dashboard -->
      <div class="page active" id="pageDashboard">
        <div class="grid">
          <div class="card span3"><div class="kpiLabel">Secure Score</div><div class="kpiValue" id="kpiSecure">N/A</div><div class="kpiHint">Target ≥ 70%</div></div>
          <div class="card span3"><div class="kpiLabel">Global Admin</div><div class="kpiValue" id="kpiGA">N/A</div><div class="kpiHint">Recommended 2–4</div></div>
          <div class="card span3"><div class="kpiLabel">MFA Registered</div><div class="kpiValue" id="kpiMFA">N/A</div><div class="kpiHint">Members ≥ 90%</div></div>
          <div class="card span3"><div class="kpiLabel">Critical + High</div><div class="kpiValue" id="kpiHigh">0</div><div class="kpiHint">Immediate attention</div></div>

          <div class="card span6">
            <div class="chartWrap">
              <div style="display:flex;justify-content:space-between;align-items:center;">
                <b>Findings by Severity</b>
                <span class="muted">Click a slice to filter</span>
              </div>
              <canvas id="sevPie" width="900" height="260"></canvas>
            </div>
          </div>

          <div class="card span6">
            <div class="chartWrap">
              <div style="display:flex;justify-content:space-between;align-items:center;">
                <b>Roadmap Workload (30/60/90)</b>
                <span class="muted">Items per timeline</span>
              </div>
              <canvas id="roadBar" width="900" height="260"></canvas>
            </div>
          </div>

          <div class="card span12">
            <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;">
              <div>
                <b>Top Risks</b>
                <div class="muted">Top Critical/High items (auto)</div>
              </div>
              <div style="display:flex;gap:8px;">
                <button class="btn" onclick="goPage('pageFindings', document.querySelector('[data-page=pageFindings]'))">Open Findings</button>
                <button class="btn" onclick="openDrawer()">Refine Filters</button>
              </div>
            </div>
            <div style="margin-top:10px;" id="topRisks"></div>
          </div>
        </div>
      </div>

      <!-- Findings -->
      <div class="page" id="pageFindings">
        <div class="grid">
          <div class="card span12">
            <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;">
              <div>
                <b>Findings</b>
                <div class="muted" id="findingsMeta">0 items</div>
              </div>
              <div style="display:flex;gap:8px;flex-wrap:wrap;">
                <button class="btn" onclick="exportFilteredCsv('findings')">Export CSV</button>
                <button class="btn" onclick="exportFilteredXlsx('findings')">Export XLSX</button>
              </div>
            </div>
            <div class="tablewrap" style="margin-top:12px;"><table id="findingsTable"></table></div>
            <div class="muted" style="margin-top:10px;">Tip: Click a row to open detail panel. Use Filters for advanced filtering.</div>
          </div>
        </div>
      </div>

      <!-- Roadmap -->
      <div class="page" id="pageRoadmap">
        <div class="grid">
          <div class="card span12">
            <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;">
              <div>
                <b>30/60/90 Roadmap</b>
                <div class="muted" id="roadmapMeta">0 items</div>
              </div>
              <div style="display:flex;gap:8px;flex-wrap:wrap;">
                <button class="btn" onclick="exportFilteredCsv('roadmap')">Export CSV</button>
                <button class="btn" onclick="exportFilteredXlsx('roadmap')">Export XLSX</button>
              </div>
            </div>
            <div class="tablewrap" style="margin-top:12px;"><table id="roadmapTable"></table></div>
          </div>
        </div>
      </div>

      <!-- Inventory -->
      <div class="page" id="pageInventory">
        <div class="grid">
          <div class="card span12">
            <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;">
              <div>
                <b>Inventory</b>
                <div class="muted" id="invMeta">Choose a view</div>
              </div>
              <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center;">
                <select id="invView" onchange="renderInventory()" style="padding:8px 10px;border-radius:10px;border:1px solid var(--border);background:var(--card);color:var(--text);">
                  <option value="licenses">Licenses</option>
                  <option value="mailbox">Mailbox Usage</option>
                  <option value="onedrive">OneDrive Usage</option>
                  <option value="groups">Groups</option>
                  <option value="recs">Top Secure Score Recommendations</option>
                </select>
                <button class="btn" onclick="exportFilteredCsv('inventory')">Export CSV</button>
                <button class="btn" onclick="exportFilteredXlsx('inventory')">Export XLSX</button>
              </div>
            </div>
            <div class="tablewrap" style="margin-top:12px;"><table id="invTable"></table></div>
          </div>
        </div>
      </div>

      <!-- Scorecard -->
      <div class="page" id="pageScorecard">
        <div class="grid">
          <div class="card span12">
            <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;">
              <div>
                <b>Scorecard</b>
                <div class="muted">Key metrics from assessment</div>
              </div>
              <div style="display:flex;gap:8px;flex-wrap:wrap;">
                <button class="btn" onclick="exportScorecardCsv()">Export CSV</button>
                <button class="btn" onclick="exportScorecardXlsx()">Export XLSX</button>
              </div>
            </div>
            <div class="tablewrap" style="margin-top:12px;"><table id="scoreTable"></table></div>
          </div>
        </div>
      </div>

      <!-- Files -->
      <div class="page" id="pageFiles">
        <div class="grid">
          <div class="card span12">
            <b>Files</b>
            <div class="muted" style="margin-top:6px;">Quick links (opens local files)</div>
            <ul id="fileList" style="margin-top:12px;"></ul>
          </div>
        </div>
      </div>

    </div>
  </main>

</div>

<div class="overlay" id="overlay" onclick="closeOverlays()"></div>

<!-- Filter Drawer -->
<div class="drawer" id="drawer">
  <div class="drawerHeader">
    <div class="drawerTitle">Filters (Findings & Roadmap)</div>
    <button class="btn" onclick="closeDrawer()">✕</button>
  </div>
  <div class="drawerBody">
    <div class="field"><label>Findings Search</label><input id="f_q" placeholder="Finding / recommendation / evidence"/></div>
    <div class="field"><label>Severity</label>
      <select id="f_sev">
        <option value="">All</option>
        <option>Critical</option><option>High</option><option>Medium</option><option>Low</option><option>Info</option>
      </select>
    </div>
    <div class="field"><label>Category</label><select id="f_cat"><option value="">All</option></select></div>
    <div class="field"><label>Effort</label>
      <select id="f_eff"><option value="">All</option><option>Quick</option><option>Planned</option><option>Project</option></select>
    </div>

    <hr style="border:none;border-top:1px solid var(--border);width:100%;"/>

    <div class="field"><label>Roadmap Timeline</label>
      <select id="r_tl"><option value="">All</option><option>30 Days</option><option>60 Days</option><option>90 Days</option></select>
    </div>
    <div class="field"><label>Roadmap Search</label><input id="r_q" placeholder="Work item / recommendation"/></div>
    <div class="field"><label>Roadmap Severity</label>
      <select id="r_sev"><option value="">All</option><option>Critical</option><option>High</option><option>Medium</option><option>Low</option><option>Info</option></select>
    </div>
  </div>
  <div class="drawerFooter">
    <button class="btn primary" onclick="applyDrawerFilters()">Apply</button>
    <button class="btn" onclick="resetDrawerFilters()">Reset</button>
  </div>
</div>

<!-- Detail Panel -->
<div class="panel" id="panel">
  <div class="panelHeader">
    <div style="min-width:0;">
      <div class="panelTitle" id="p_title">Detail</div>
      <div class="muted" id="p_sub">—</div>
    </div>
    <button class="btn" onclick="closePanel()">✕</button>
  </div>
  <div class="panelBody" id="p_body"></div>
</div>

<script>
// Embedded data
const SCORECARD = __SCORECARD_JSON__;
const FINDINGS_ALL  = __FINDINGS_JSON__;
const ROADMAP_ALL   = __ROADMAP_JSON__;
const LICENSES  = __LICENSES_JSON__;
const MAILBOX   = __MAILBOX_JSON__;
const ONEDRIVE  = __ONEDRIVE_JSON__;
const GROUPS    = __GROUPS_JSON__;
const RECS      = __RECS_JSON__;

let FINDINGS = [...FINDINGS_ALL];
let ROADMAP  = [...ROADMAP_ALL];
let INVENTORY = [];
let INVENTORY_VIEW = 'licenses';

let filteredFindings = [];
let filteredRoadmap  = [];

function hasXlsx(){ return (typeof XLSX !== 'undefined' && XLSX.utils && XLSX.writeFile); }

function toggleTheme(){
  const root=document.documentElement;
  const cur=root.getAttribute('data-theme');
  if(cur==='dark') root.removeAttribute('data-theme');
  else root.setAttribute('data-theme','dark');
}

function goPage(id, el){
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  document.querySelectorAll('.navBtn').forEach(b=>b.classList.remove('active'));
  if(el) el.classList.add('active');
}

function openDrawer(){
  document.getElementById('drawer').classList.add('open');
  document.getElementById('overlay').classList.add('show');
}
function closeDrawer(){
  document.getElementById('drawer').classList.remove('open');
  document.getElementById('overlay').classList.remove('show');
}
function openPanel(){
  document.getElementById('panel').classList.add('open');
  document.getElementById('overlay').classList.add('show');
}
function closePanel(){
  document.getElementById('panel').classList.remove('open');
  document.getElementById('overlay').classList.remove('show');
}
function closeOverlays(){
  closeDrawer();
  closePanel();
}

function esc(s){ if(s===null||s===undefined) return ''; return String(s).replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;'); }

function badge(sev){
  const s = sev || 'Info';
  if(s==='Critical') return `<span class="badge bCritical">● Critical</span>`;
  if(s==='High') return `<span class="badge bHigh">● High</span>`;
  if(s==='Medium') return `<span class="badge bMedium">● Medium</span>`;
  if(s==='Low') return `<span class="badge bLow">● Low</span>`;
  return `<span class="badge bInfo">● Info</span>`;
}

function setKpis(){
  const getVal=(m)=>{ const r=SCORECARD.find(x=>x.Metric===m); return r? r.Value : ''; };
  const ss = getVal('Secure Score %');
  const ga = getVal('Global Admin Assignments');
  const mfa = getVal('MFA Registered % (Members)');

  const ch = FINDINGS_ALL.filter(x=>['Critical','High'].includes(x.Severity)).length;

  const elSS=document.getElementById('kpiSecure');
  elSS.textContent = ss? (ss+'%') : 'N/A';
  if(ss && Number(ss)<50) elSS.style.color = 'var(--critical)';

  document.getElementById('kpiGA').textContent = ga || 'N/A';
  document.getElementById('kpiMFA').textContent = mfa? (mfa+'%') : 'N/A';
  const elH=document.getElementById('kpiHigh');
  elH.textContent = String(ch);
  if(ch>0) elH.style.color = 'var(--high)';
}

function renderTable(elId, rows, cols, colFormat){
  const el=document.getElementById(elId);
  const thead = '<thead><tr>' + cols.map(c=>`<th>${esc(c)}</th>`).join('') + '</tr></thead>';
  const tbody = rows.map((r)=>{
    const tds = cols.map(c=>{
      const v = r[c];
      const f = colFormat && colFormat[c];
      return `<td>${f? f(v,r) : esc(v)}</td>`;
    }).join('');
    return `<tr class="rowClickable" data-row="${encodeURIComponent(JSON.stringify(r))}">${tds}</tr>`;
  }).join('');
  el.innerHTML = thead + '<tbody>' + tbody + '</tbody>';
}

function attachRowClicks(tableId){
  const table=document.getElementById(tableId);
  table.querySelectorAll('tbody tr').forEach(tr=>{
    tr.addEventListener('click', ()=>{
      const obj = JSON.parse(decodeURIComponent(tr.getAttribute('data-row')));
      showDetail(obj);
    });
  });
}

function showDetail(obj){
  const title = obj.Finding || obj.WorkItem || obj.Metric || 'Detail';
  document.getElementById('p_title').textContent = title;
  const sev = obj.Severity || obj.Timeline || '';
  document.getElementById('p_sub').innerHTML = sev ? (badge(obj.Severity) + ' &nbsp; ' + esc(obj.Category||obj.Timeline||'')) : esc(obj.Category||'');

  const body=document.getElementById('p_body');
  body.innerHTML='';

  // Build sections
  const sec1=document.createElement('div');
  sec1.className='panelSection';
  sec1.innerHTML = `<div class="kv">
      <b>Category</b><div>${esc(obj.Category||'—')}</div>
      <b>Severity</b><div>${obj.Severity? badge(obj.Severity): esc('—')}</div>
      <b>Effort</b><div>${esc(obj.Effort||'—')}</div>
      <b>Timeline</b><div>${esc(obj.Timeline||'—')}</div>
    </div>`;

  const sec2=document.createElement('div');
  sec2.className='panelSection';
  sec2.innerHTML = `<b>Evidence</b><div class="muted" style="margin-top:6px;white-space:pre-wrap;">${esc(obj.Evidence||'—')}</div>`;

  const sec3=document.createElement('div');
  sec3.className='panelSection';
  sec3.innerHTML = `<b>Recommendation</b><div class="muted" style="margin-top:6px;white-space:pre-wrap;">${esc(obj.Recommendation||'—')}</div>`;

  body.appendChild(sec1);
  body.appendChild(sec2);
  body.appendChild(sec3);

  // Quick actions
  const sec4=document.createElement('div');
  sec4.className='panelSection';
  sec4.innerHTML = `<div style="display:flex;gap:8px;flex-wrap:wrap;">
      <button class="btn" onclick="copyText('${encodeURIComponent(title)}')">Copy Finding</button>
      <button class="btn" onclick="copyText('${encodeURIComponent(obj.Recommendation||'')}')">Copy Recommendation</button>
      <button class="btn" onclick="exportOneRowXlsx()">Export this row (XLSX)</button>
    </div>`;
  body.appendChild(sec4);

  window.__lastDetailRow = obj;
  openPanel();
}

function copyText(encoded){
  const t = decodeURIComponent(encoded);
  navigator.clipboard.writeText(t).then(()=>{},()=>{});
}

function exportOneRowXlsx(){
  if(!hasXlsx()){ alert('XLSX library not available (xlsx.full.min.js missing).'); return; }
  const row = window.__lastDetailRow ? [window.__lastDetailRow] : [];
  if(row.length===0){ alert('No row selected.'); return; }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(row), 'Detail');
  XLSX.writeFile(wb, 'Selected_Detail.xlsx');
}

// Filters
function buildCategoryOptions(){
  const cats = Array.from(new Set(FINDINGS_ALL.map(x=>x.Category).filter(Boolean))).sort();
  const sel = document.getElementById('f_cat');
  cats.forEach(c=>{ const opt=document.createElement('option'); opt.value=c; opt.textContent=c; sel.appendChild(opt); });
}

function applyDrawerFilters(){
  const q = document.getElementById('f_q').value.trim().toLowerCase();
  const sev = document.getElementById('f_sev').value;
  const cat = document.getElementById('f_cat').value;
  const eff = document.getElementById('f_eff').value;

  let f = FINDINGS_ALL;
  if(sev) f = f.filter(x=>(x.Severity||'Info')===sev);
  if(cat) f = f.filter(x=>x.Category===cat);
  if(eff) f = f.filter(x=>(x.Effort||'')===eff);
  if(q) f = f.filter(x=>`${x.Category||''} ${x.Finding||''} ${x.Severity||''} ${x.Evidence||''} ${x.Recommendation||''} ${x.Effort||''}`.toLowerCase().includes(q));

  FINDINGS = f;

  const tl = document.getElementById('r_tl').value;
  const rq = document.getElementById('r_q').value.trim().toLowerCase();
  const rsev = document.getElementById('r_sev').value;

  let r = ROADMAP_ALL;
  if(tl) r = r.filter(x=>x.Timeline===tl);
  if(rsev) r = r.filter(x=>(x.Severity||'Info')===rsev);
  if(rq) r = r.filter(x=>`${x.Timeline||''} ${x.Category||''} ${x.WorkItem||''} ${x.Recommendation||''} ${x.Evidence||''} ${x.Severity||''}`.toLowerCase().includes(rq));

  ROADMAP = r;

  refreshAllViews();
  closeDrawer();
}

function resetDrawerFilters(){
  document.getElementById('f_q').value='';
  document.getElementById('f_sev').value='';
  document.getElementById('f_cat').value='';
  document.getElementById('f_eff').value='';
  document.getElementById('r_tl').value='';
  document.getElementById('r_q').value='';
  document.getElementById('r_sev').value='';
  FINDINGS = [...FINDINGS_ALL];
  ROADMAP  = [...ROADMAP_ALL];
  refreshAllViews();
}

function quickSearch(){
  // lightweight search applied to findings only
  const q = document.getElementById('globalSearch').value.trim().toLowerCase();
  if(!q){
    FINDINGS = [...FINDINGS_ALL];
  } else {
    FINDINGS = FINDINGS_ALL.filter(x=>`${x.Category||''} ${x.Finding||''} ${x.Severity||''} ${x.Evidence||''} ${x.Recommendation||''}`.toLowerCase().includes(q));
  }
  refreshFindings();
}

// Rendering
function refreshFindings(){
  filteredFindings = FINDINGS;
  document.getElementById('findingsMeta').textContent = `${filteredFindings.length} items (showing up to 600)`;
  const rows = filteredFindings.slice(0,600);
  renderTable('findingsTable', rows, ['Category','Finding','Severity','Effort'], {Severity:(v)=>badge(v)});
  attachRowClicks('findingsTable');
  renderTopRisks();
  buildCharts();
}

function refreshRoadmap(){
  filteredRoadmap = ROADMAP;
  document.getElementById('roadmapMeta').textContent = `${filteredRoadmap.length} items (showing up to 800)`;
  const rows = filteredRoadmap.slice(0,800);
  renderTable('roadmapTable', rows, ['Timeline','Category','WorkItem','Severity'], {Severity:(v)=>badge(v)});
  attachRowClicks('roadmapTable');
  buildCharts();
}

function renderScorecard(){
  renderTable('scoreTable', SCORECARD, ['Metric','Value']);
}

function renderInventory(){
  const view = document.getElementById('invView').value;
  INVENTORY_VIEW = view;
  let rows = [];
  if(view==='licenses') rows = LICENSES;
  if(view==='mailbox') rows = MAILBOX;
  if(view==='onedrive') rows = ONEDRIVE;
  if(view==='groups') rows = GROUPS;
  if(view==='recs') rows = RECS;
  INVENTORY = rows;
  document.getElementById('invMeta').textContent = `${rows.length} rows (showing up to 800)`;
  const show = rows.slice(0,800);
  const cols = show[0] ? Object.keys(show[0]) : [];
  renderTable('invTable', show, cols);
  attachRowClicks('invTable');
}

function renderFiles(){
  const files = [
    '00_Dashboard_ENTERPRISE.html','xlsx.full.min.js','00_Scorecard.csv','41_Findings_RiskRegister.csv','40_Roadmap_30_60_90.csv','00_Executive_Summary.html',
    '01_All_Users.csv','02_Mailbox_Usage.csv','03_OneDrive_Usage.csv','04_Mailbox_Under_Threshold.csv','05_License_Inventory.csv',
    '06a_Security_Score_Summary.csv','06b_Security_Score_Controls.csv','12b_Security_Control_Profiles.csv','30_Top_SecureScore_Recommendations.csv',
    '20_Role_Assignments.csv','21_ConditionalAccess_Policies.json','22_MFA_Registration.csv','08_Shared_Mailboxes.csv'
  ];
  const ul = document.getElementById('fileList');
  ul.innerHTML='';
  const root = "__OUTPUT_ROOT__";
  files.forEach(f=>{
    const li=document.createElement('li');
    li.innerHTML = `<a href="file:///${root.replaceAll('\\\\','/')}/${f}">${esc(f)}</a>`;
    ul.appendChild(li);
  });
}

function renderTopRisks(){
  const top = filteredFindings.filter(x=>['Critical','High'].includes(x.Severity)).slice(0,8);
  if(top.length===0){
    document.getElementById('topRisks').innerHTML = `<div class="muted">No Critical/High findings in current filter.</div>`;
    return;
  }
  const html = top.map((x,i)=>{
    return `<div style="display:flex;align-items:flex-start;gap:10px;padding:10px;border-bottom:1px solid var(--border);">
      <div style="width:28px;font-weight:900;color:var(--muted);">${i+1}</div>
      <div style="flex:1;min-width:0;">
        <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">${badge(x.Severity)} <b>${esc(x.Category||'')}</b></div>
        <div style="margin-top:6px;">${esc(x.Finding||'')}</div>
        <div class="muted" style="margin-top:6px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${esc(x.Recommendation||'')}</div>
      </div>
      <button class="btn" onclick="showDetail(${encodeURIComponent(JSON.stringify(x))}.startsWith('')? JSON.parse(decodeURIComponent('${encodeURIComponent(JSON.stringify(x))}')) : null )">View</button>
    </div>`;
  }).join('');

  // the onclick above is messy in HTML string; we will attach click via delegation instead.
  document.getElementById('topRisks').innerHTML = '';
  const wrapper=document.createElement('div');
  wrapper.style.marginTop='10px';
  top.forEach((x,i)=>{
    const row=document.createElement('div');
    row.style.display='flex'; row.style.alignItems='flex-start'; row.style.gap='10px';
    row.style.padding='10px'; row.style.borderBottom='1px solid var(--border)';
    row.innerHTML = `<div style="width:28px;font-weight:900;color:var(--muted);">${i+1}</div>
      <div style="flex:1;min-width:0;">
        <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">${badge(x.Severity)} <b>${esc(x.Category||'')}</b></div>
        <div style="margin-top:6px;">${esc(x.Finding||'')}</div>
        <div class="muted" style="margin-top:6px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${esc(x.Recommendation||'')}</div>
      </div>`;
    const btn=document.createElement('button');
    btn.className='btn';
    btn.textContent='View';
    btn.onclick=()=>showDetail(x);
    row.appendChild(btn);
    wrapper.appendChild(row);
  });
  document.getElementById('topRisks').appendChild(wrapper);
}

function refreshAllViews(){
  refreshFindings();
  refreshRoadmap();
  renderInventory();
}

// Exports
function toCsv(rows){
  if(!rows||rows.length===0) return '';
  const cols=Object.keys(rows[0]);
  const escCsv=(v)=>{ if(v===null||v===undefined) return ''; const s=String(v).replaceAll('"','""'); return (/[",\n]/.test(s))? `"${s}"` : s; };
  const head=cols.join(',');
  const body=rows.map(r=>cols.map(c=>escCsv(r[c])).join(',')).join('\n');
  return head+'\n'+body;
}

function downloadBlob(data, filename, mime){
  const blob=new Blob([data], {type:mime});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');
  a.href=url; a.download=filename;
  document.body.appendChild(a);
  a.click(); a.remove();
  URL.revokeObjectURL(url);
}

function exportFilteredCsv(which){
  if(which==='findings') return downloadBlob(toCsv(filteredFindings), 'Findings_Filtered.csv', 'text/csv;charset=utf-8;');
  if(which==='roadmap') return downloadBlob(toCsv(filteredRoadmap), 'Roadmap_Filtered.csv', 'text/csv;charset=utf-8;');
  if(which==='inventory') return downloadBlob(toCsv(INVENTORY), `Inventory_${INVENTORY_VIEW}.csv`, 'text/csv;charset=utf-8;');
}

function exportFilteredXlsx(which){
  if(!hasXlsx()){ alert('XLSX library not available (xlsx.full.min.js missing).'); return; }
  let rows=[], name='Sheet', file='Export.xlsx';
  if(which==='findings'){ rows=filteredFindings; name='Findings'; file='Findings_Filtered.xlsx'; }
  if(which==='roadmap'){ rows=filteredRoadmap; name='Roadmap'; file='Roadmap_Filtered.xlsx'; }
  if(which==='inventory'){ rows=INVENTORY; name='Inventory_'+INVENTORY_VIEW; file=`Inventory_${INVENTORY_VIEW}.xlsx`; }
  if(!rows||rows.length===0){ alert('No data to export.'); return; }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), name);
  XLSX.writeFile(wb, file);
}

function exportScorecardCsv(){ downloadBlob(toCsv(SCORECARD), 'Scorecard.csv', 'text/csv;charset=utf-8;'); }
function exportScorecardXlsx(){
  if(!hasXlsx()){ alert('XLSX library not available (xlsx.full.min.js missing).'); return; }
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(SCORECARD), 'Scorecard');
  XLSX.writeFile(wb, 'Scorecard.xlsx');
}

function exportAllWorkbook(){
  if(!hasXlsx()){ alert('XLSX library not available (xlsx.full.min.js missing).'); return; }
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(SCORECARD), 'Scorecard');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(filteredFindings.length?filteredFindings:FINDINGS_ALL), 'Findings');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(filteredRoadmap.length?filteredRoadmap:ROADMAP_ALL), 'Roadmap');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(INVENTORY || []), 'Inventory_'+INVENTORY_VIEW);
  XLSX.writeFile(wb, 'M365_Assessment_Export_All.xlsx');
}

// Charts (pure canvas)
function drawPie(canvasId, data){
  const c=document.getElementById(canvasId); const ctx=c.getContext('2d'); const w=c.width, h=c.height; ctx.clearRect(0,0,w,h);
  const total=data.reduce((a,b)=>a+b.value,0)||1;
  const cx=w*0.32, cy=h*0.56, r=Math.min(w,h)*0.40;
  let start=-Math.PI/2;
  data.forEach(d=>{
    const angle=(d.value/total)*Math.PI*2; const end=start+angle;
    ctx.beginPath(); ctx.moveTo(cx,cy); ctx.arc(cx,cy,r,start,end); ctx.closePath(); ctx.fillStyle=d.color; ctx.fill();
    start=end;
  });
  // Legend
  const lx=w*0.66, ly=h*0.18; ctx.font='14px Segoe UI';
  data.forEach((d,i)=>{ ctx.fillStyle=d.color; ctx.fillRect(lx,ly+i*26,14,14); ctx.fillStyle=getComputedStyle(document.documentElement).getPropertyValue('--text'); ctx.fillText(`${d.label}: ${d.value}`, lx+20, ly+12+i*26); });
  // Click filter
  c.onclick=(ev)=>{
    const rect=c.getBoundingClientRect();
    const x=(ev.clientX-rect.left)*(c.width/rect.width);
    const y=(ev.clientY-rect.top)*(c.height/rect.height);
    const dx=x-cx, dy=y-cy;
    const dist=Math.sqrt(dx*dx+dy*dy);
    if(dist>r) return;
    let ang=Math.atan2(dy,dx);
    let norm=ang - (-Math.PI/2);
    if(norm<0) norm+=Math.PI*2;
    let acc=0;
    for(const d of data){
      const a=(d.value/total)*Math.PI*2;
      if(norm>=acc && norm<=acc+a){
        document.getElementById('f_sev').value=d.label;
        applyDrawerFilters();
        goPage('pageFindings', document.querySelector('[data-page=pageFindings]'));
        return;
      }
      acc+=a;
    }
  };
}

function drawBar(canvasId, data){
  const c=document.getElementById(canvasId); const ctx=c.getContext('2d'); const w=c.width, h=c.height; ctx.clearRect(0,0,w,h);
  const max=Math.max(...data.map(d=>d.value),1);
  const left=60, top=20, right=20, bottom=40;
  const slot=(w-left-right)/data.length; const bw=slot*0.6;
  ctx.strokeStyle='rgba(0,0,0,.2)';
  ctx.beginPath(); ctx.moveTo(left,top); ctx.lineTo(left,h-bottom); ctx.lineTo(w-right,h-bottom); ctx.stroke();
  ctx.font='14px Segoe UI';
  data.forEach((d,i)=>{
    const x=left+i*slot+slot*0.2;
    const barH=(h-top-bottom)*(d.value/max);
    const y=h-bottom-barH;
    ctx.fillStyle=getComputedStyle(document.documentElement).getPropertyValue('--brand');
    ctx.fillRect(x,y,bw,barH);
    ctx.fillStyle=getComputedStyle(document.documentElement).getPropertyValue('--text');
    ctx.fillText(String(d.value), x, y-6);
    ctx.fillStyle=getComputedStyle(document.documentElement).getPropertyValue('--muted');
    ctx.fillText(d.label, x, h-bottom+22);
  });
}

function buildCharts(){
  const counts={Critical:0,High:0,Medium:0,Low:0,Info:0};
  filteredFindings.forEach(f=>{ const s=f.Severity||'Info'; if(counts[s]===undefined) counts.Info++; else counts[s]++; });
  const pieData=[
    {label:'Critical',value:counts.Critical,color:getComputedStyle(document.documentElement).getPropertyValue('--critical').trim()||'#d13438'},
    {label:'High',value:counts.High,color:getComputedStyle(document.documentElement).getPropertyValue('--high').trim()||'#ca5010'},
    {label:'Medium',value:counts.Medium,color:getComputedStyle(document.documentElement).getPropertyValue('--medium').trim()||'#8a6d3b'},
    {label:'Low',value:counts.Low,color:getComputedStyle(document.documentElement).getPropertyValue('--low').trim()||'#107c10'},
    {label:'Info',value:counts.Info,color:getComputedStyle(document.documentElement).getPropertyValue('--brand').trim()||'#0f6cbd'}
  ];
  drawPie('sevPie', pieData);

  const rc={'30':0,'60':0,'90':0};
  filteredRoadmap.forEach(r=>{ if(r.Timeline==='30 Days') rc['30']++; if(r.Timeline==='60 Days') rc['60']++; if(r.Timeline==='90 Days') rc['90']++; });
  drawBar('roadBar', [{label:'30',value:rc['30']},{label:'60',value:rc['60']},{label:'90',value:rc['90']}]);
}

// Init
setKpis();
buildCategoryOptions();
renderScorecard();
renderInventory();
renderFiles();
refreshAllViews();
</script>
</body>
</html>
'@

$html = $template
$html = $html.Replace('__OUTPUT_ROOT__', ($OutputRoot -replace '\\','/'))
$html = $html.Replace('__GENERATED__', (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
$html = $html.Replace('__SCORECARD_JSON__', ($scoreJson  ?? '[]'))
$html = $html.Replace('__FINDINGS_JSON__',  ($findingsJson ?? '[]'))
$html = $html.Replace('__ROADMAP_JSON__',   ($roadmapJson ?? '[]'))
$html = $html.Replace('__LICENSES_JSON__',  ($licensesJson ?? '[]'))
$html = $html.Replace('__MAILBOX_JSON__',   ($mailboxJson ?? '[]'))
$html = $html.Replace('__ONEDRIVE_JSON__',  ($onedriveJson ?? '[]'))
$html = $html.Replace('__GROUPS_JSON__',    ($groupsJson ?? '[]'))
$html = $html.Replace('__RECS_JSON__',      ($recsJson ?? '[]'))

$dashPath = Join-Path $OutputRoot '00_Dashboard_ENTERPRISE.html'
$html | Out-File -FilePath $dashPath -Encoding UTF8

Write-Ok "Dashboard generated: $dashPath"
Invoke-Item $dashPath
