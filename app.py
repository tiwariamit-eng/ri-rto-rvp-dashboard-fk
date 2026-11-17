#!/usr/bin/env python3
"""
RTO / RVP RI% Dashboard — Streamlit launcher, styled like the RI Loss dashboard.
Data source: Google Drive folder (all .xlsx / .csv files combined, like the
original local ri_dashboard_streamlit.py script).
"""
import io
import json
import os
import tempfile
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload



# ==============================================================================
# CONFIG (from the uploaded script)
# ==============================================================================
DRIVE_FOLDER_ID = "1rLayEdey465A4nMMBOrmO0ZYp0_441Tt"

BU_EXCLUSIONS = [
    'CoreElectronics', 'EmergingElectronics', 'GiftCard', 'Refurbished',
    'ShopsyTest', 'NDC', 'Shopsy', 'TestNDC'
]

ZONE_MAPPING = {
    'ahm_khe_wh_nl_01nl': 'West', 'ban_ven_wh_nl_01nl': 'South', 'bhi_pad_wh_nl_01nl': 'West',
    'bhi_vas_wh_nl_01nl': 'West', 'cbe_oth_01': 'South', 'chn_tvr_04': 'South',
    'coc_mfc_wh_nl_01nl': 'North', 'cut_tan_wh_nl_01nl': 'East', 'dha_kot_wh_nl_01nl': 'South',
    'frk_bts': 'North', 'gur_san_wh_nl_01nl': 'North', 'guw_kam_wh_nl_02nl': 'East',
    'hyderabad_medchal_01': 'South', 'ind_teh_wh_nl_01nl': 'West', 'jai_san_wh_nl_01nl': 'North',
    'jal_bin_wh_nl_01nl': 'East', 'luc_has_wh_nl_02nl': 'North', 'lud_kha_wh_nl_01nl': 'North',
    'malur_bts': 'South', 'nad_har_wh_kl_nl_01nl': 'East', 'nag_wad_wh_nl_02nl': 'West',
    'pat_haj_wh_nl_01nl': 'East', 'rai_dha_wh_nl_02nl': 'East', 'ran_mfc_wh_nl_01nl': 'East',
    'ulub_bts': 'East', 'vij_ved_wh_nl_01nl': 'South'
}

MONTH_ORDER = ["January", "February", "March", "April", "May", "June",
               "July", "August", "September", "October", "November", "December"]

ZONE_COLORS = {"North": "#3B82F6", "South": "#10B981", "East": "#F59E0B",
               "West": "#EF4444", "Unknown": "#94A3B8"}


# ==============================================================================
# GOOGLE DRIVE FETCH (same pattern as the resealing/refinishing/RI Loss dashboards)
# ==============================================================================
def get_drive_service():
    creds = Credentials(
        token=None,
        refresh_token=st.secrets["GOOGLE_REFRESH_TOKEN"],
        client_id=st.secrets["GOOGLE_CLIENT_ID"],
        client_secret=st.secrets["GOOGLE_CLIENT_SECRET"],
        token_uri="https://oauth2.googleapis.com/token",
    )
    return build("drive", "v3", credentials=creds)


def list_data_files(service, folder_id):
    query = f"'{folder_id}' in parents and trashed = false"
    results = service.files().list(
        q=query, fields="files(id, name, modifiedTime, mimeType)",
        orderBy="modifiedTime desc", pageSize=200
    ).execute()
    files = results.get("files", [])
    return [f for f in files if f["name"].lower().endswith((".csv", ".xlsx", ".xls"))]


def download_file_to_disk(service, file_id, suffix):
    request = service.files().get_media(fileId=file_id)
    fd, path = tempfile.mkstemp(suffix=suffix)
    with os.fdopen(fd, "wb") as f:
        downloader = MediaIoBaseDownload(f, request, chunksize=10 * 1024 * 1024)
        done = False
        while not done:
            _, done = downloader.next_chunk()
    return path


# Memory: only these columns are read, and each file is pre-aggregated to the
# grain the dashboard needs (sums are additive, so results are unchanged).
NEEDED_COLS = ['month', 'return_warehouse', 'product_categorization_dim_business_unit',
               'Alpha_flag', 'return_type', 'selling_price', 'quantity', 'w_sp', 'w_qty']
READ_DTYPES = {'month': 'category', 'return_warehouse': 'category',
               'product_categorization_dim_business_unit': 'category',
               'Alpha_flag': 'category', 'return_type': 'category'}
GROUP_COLS = ['month', 'return_warehouse', 'product_categorization_dim_business_unit',
              'Alpha_flag', 'return_type']
SUM_COLS = ['selling_price', 'quantity', 'w_sp', 'w_qty']


def _aggregate(raw):
    # No cleaning here (no strip / NaN->str): prepare_data does it exactly as before.
    raw = raw.copy()
    for c in GROUP_COLS:  # category -> object so files concat cleanly
        raw[c] = raw[c].astype(object)
    raw['rows'] = 1
    return raw.groupby(GROUP_COLS, as_index=False, dropna=False, observed=True)[SUM_COLS + ['rows']].sum()


@st.cache_data(show_spinner="Fetching RI files from Google Drive...", ttl=600)
def fetch_and_combine_from_drive(folder_id):
    service = get_drive_service()
    files = list_data_files(service, folder_id)
    if not files:
        raise FileNotFoundError(f"No .csv/.xlsx files found in Drive folder {folder_id}")

    frames = []
    names = []
    for f in files:
        suffix = os.path.splitext(f["name"])[1].lower()
        path = download_file_to_disk(service, f["id"], suffix)
        try:
            if suffix == ".csv":
                raw = pd.read_csv(path, usecols=NEEDED_COLS, dtype=READ_DTYPES)
            else:
                raw = pd.read_excel(path, usecols=NEEDED_COLS)
            frames.append(_aggregate(raw))  # shrink per file; never hold all raw rows
            del raw
            names.append(f["name"])
        finally:
            os.remove(path)

    df = pd.concat(frames, ignore_index=True)
    return df, names


# ==============================================================================
# PREPROCESS + RI% CALCULATION (logic from the uploaded script)
# ==============================================================================
@st.cache_data(show_spinner="Calculating RI% metrics...")
def prepare_data(df):
    df = df.copy()
    df['month'] = df['month'].astype(str).str.strip().str.title()
    df['zone'] = df['return_warehouse'].map(ZONE_MAPPING).fillna('Unknown')

    if 'product_categorization_dim_business_unit' in df.columns:
        df['product_categorization_dim_business_unit'] = (
            df['product_categorization_dim_business_unit'].astype(str).str.strip()
        )

    exclusions_lower = [bu.lower().strip() for bu in BU_EXCLUSIONS]
    df = df[~df['product_categorization_dim_business_unit'].str.lower().isin(exclusions_lower)].copy()

    # Drop Unknown zone and Unknown business unit rows
    df = df[(df['zone'] != 'Unknown') &
            (df['product_categorization_dim_business_unit'].str.lower() != 'unknown')].copy()

    df['cleaned_flag'] = pd.to_numeric(df['Alpha_flag'], errors='coerce').fillna(-1).astype(int).astype(str)
    df['cleaned_return_type'] = df['return_type'].astype(str).str.lower().str.strip()
    df['cleaned_bu'] = df['product_categorization_dim_business_unit'].str.lower()

    return df




def _granular(dfx, value_col, weight_col):
    """Compact per-(warehouse, zone, bu, month) sums — client-side JS derives
    RI% after applying whatever Month/BU/Warehouse filters are active."""
    if dfx.empty:
        return []
    g = dfx.groupby(
        ['return_warehouse', 'zone', 'product_categorization_dim_business_unit', 'month'],
        observed=True
    ).agg(
        w=(weight_col, 'sum'),
        v=(value_col, 'sum'),
    ).reset_index()
    g.columns = ['wh', 'zone', 'bu', 'month', 'w', 'v']
    return [
        {"wh": str(r.wh), "zone": str(r.zone), "bu": str(r.bu), "month": str(r.month),
         "w": round(float(r.w), 4), "v": round(float(r.v), 4)}
        for r in g.itertuples(index=False)
    ]


@st.cache_data(show_spinner="Building dashboard payload...")
def build_payload(df):
    df_rto = df[df['cleaned_return_type'] == 'courier_return'].copy()
    df_rvp = df[
        (df['cleaned_return_type'] == 'customer_return') &
        (df['cleaned_flag'] == '1')
    ].copy()

    months_present = sorted(
        set(df['month'].unique()),
        key=lambda x: MONTH_ORDER.index(x) if x in MONTH_ORDER else 999
    )
    zones_present = sorted(set(df['zone'].unique()))
    bus_present = sorted(set(df['cleaned_bu'].str.title().unique()) | set(df['product_categorization_dim_business_unit'].unique()))
    warehouses_present = sorted(set(df['return_warehouse'].unique()))

    payload = {
        "meta": {
            "months": months_present,
            "zones": zones_present,
            "warehouses": warehouses_present,
            "row_count": int(df['rows'].sum()),
        },
        "rto": _granular(df_rto, 'selling_price', 'w_sp'),
        "rvp": _granular(df_rvp, 'quantity', 'w_qty'),
        "rto_bus": sorted(set(df_rto['product_categorization_dim_business_unit'].unique())),
        "rvp_bus": sorted(set(df_rvp['product_categorization_dim_business_unit'].unique())),
    }
    return payload


# ==============================================================================
# HTML/JS DASHBOARD TEMPLATE
# ==============================================================================
TEMPLATE = r""" <meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>RTO / RVP RI% Dashboard</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600&display=swap" rel="stylesheet">

<style>
:root{
  --bg:#f6f7f9; --panel:#ffffff; --ink:#141b26; --muted:#6a7482; --faint:#98a1ad;
  --line:#e7eaef; --line2:#eef1f5; --accent:#2f5bd4;
  --north:#3b6fd4; --south:#0e9b8a; --east:#e0952a; --west:#8a4fd0; --unknown:#94a0af;
  --up:#d64550; --down:#1f9d6b;
  --shadow:0 1px 2px rgba(20,27,38,.04),0 2px 8px rgba(20,27,38,.05);
}
*{box-sizing:border-box}
html, body{height:100%; overflow:hidden;}
body{margin:0;background:var(--bg);color:var(--ink);
  font-family:Inter,-apple-system,Segoe UI,Roboto,sans-serif;font-size:13px;line-height:1.45;
  display:flex; flex-direction:column;}
h1,h2,h3{font-family:"Space Grotesk",Inter,sans-serif;margin:0;letter-spacing:-.01em}
.num{font-family:"Space Grotesk",Inter,sans-serif;font-variant-numeric:tabular-nums}

.scroll-container{flex:1; overflow-y:auto; padding:0 20px 64px;}
.wrap{max-width:1480px;margin:0 auto}
.split{display:grid;grid-template-columns:1.5fr 1fr;gap:14px}
@media(max-width:1000px){.split{grid-template-columns:1fr}}

.sticky-top-panel{position:sticky; top:0; z-index:100; background:var(--bg);
  border-bottom:1px solid var(--line); padding:0 20px;}

header{padding:18px 0 10px;}
.eyebrow{font-size:11px;font-weight:600;letter-spacing:.14em;text-transform:uppercase;color:var(--accent)}
header h1{font-size:24px;font-weight:700;margin-top:3px}
.sub{color:var(--muted);margin-top:4px;font-size:12.5px}

.filters{padding:10px 0 14px;}
.filters .row{display:flex;gap:14px;flex-wrap:wrap;align-items:flex-end}
.fg{display:flex;flex-direction:column;gap:5px;min-width:150px}
.fg label{font-size:10.5px;font-weight:600;letter-spacing:.08em;text-transform:uppercase;color:var(--muted)}
.chips{display:flex;gap:5px;flex-wrap:wrap;max-width:380px}
.chip{border:1px solid var(--line);background:var(--panel);border-radius:999px;
  padding:4px 10px;font-size:11.5px;cursor:pointer;user-select:none;transition:.12s;font-weight:500}
.chip:hover{border-color:var(--accent)}
.chip.on{background:var(--accent);border-color:var(--accent);color:#fff}
select{font-family:inherit;font-size:12.5px;padding:6px 9px;border:1px solid var(--line);
  border-radius:8px;background:var(--panel);color:var(--ink);min-width:170px}
.btn{border:1px solid var(--line);background:var(--panel);border-radius:8px;padding:7px 13px;
  font-size:12px;font-weight:600;cursor:pointer}
.btn:hover{border-color:var(--accent);color:var(--accent)}
.scope{font-size:12px;color:var(--muted);margin-left:auto}
.scope b{color:var(--ink);font-weight:600}

.tabs{display:flex;gap:8px;margin:14px 0 4px}
.tab-btn{padding:7px 15px;border-radius:8px;border:1px solid var(--line);background:var(--panel);
  cursor:pointer;font-size:12.5px;font-weight:600;color:var(--muted)}
.tab-btn.active{background:var(--accent);color:#fff;border-color:var(--accent)}

.kpis{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin:16px 0 26px}
.kpi{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:14px 15px;box-shadow:var(--shadow)}
.kpi .k{font-size:10.5px;font-weight:600;letter-spacing:.06em;text-transform:uppercase;color:var(--muted)}
.kpi .v{font-size:23px;font-weight:700;margin-top:6px;font-family:"Space Grotesk";font-variant-numeric:tabular-nums}
.kpi .v.good{color:var(--down)} .kpi .v.bad{color:var(--up)} .kpi .v.warn{color:#a5820a}
td.pct-good{background:#e6f7ef;color:#0f6e56;font-weight:600}
td.pct-warn{background:#fff7de;color:#8a6d1a;font-weight:600}
td.pct-bad{background:#fdeceb;color:#a3312d;font-weight:600}
.kpi .d{font-size:11.5px;margin-top:3px;color:var(--faint)}

section{margin-bottom:24px}
.shead{display:flex;align-items:baseline;gap:10px;margin-bottom:12px}
.shead .n{font-family:"Space Grotesk";font-weight:700;color:var(--accent);font-size:13px}
.shead h2{font-size:16px;font-weight:600}
.shead .h{font-size:12px;color:var(--muted);margin-left:6px}
.card{background:var(--panel);border:1px solid var(--line);border-radius:12px;box-shadow:var(--shadow);overflow:hidden}
.tscroll{max-height:440px;overflow:auto}

table{width:100%;border-collapse:separate;border-spacing:0;font-size:12.5px}
thead th{position:sticky;top:0;z-index:2;background:#f2f4f8;color:var(--muted);
  font-weight:600;text-align:right;padding:9px 12px;font-size:11px;letter-spacing:.03em;
  text-transform:uppercase;border-bottom:1px solid var(--line);white-space:nowrap}
thead th:first-child{text-align:left}
tbody td{padding:8px 12px;text-align:right;border-bottom:1px solid var(--line2);white-space:nowrap;
  font-variant-numeric:tabular-nums}
tbody td:first-child{text-align:left;font-weight:500}
tbody tr:hover{background:#f8fafc; cursor:pointer;}
.zband{border-left:3px solid var(--unknown)}
.zband.North{border-left-color:var(--north)} .zband.South{border-left-color:var(--south)}
.zband.East{border-left-color:var(--east)} .zband.West{border-left-color:var(--west)}
.zband.selected{background:#f0f4ff}
.muted{color:var(--muted)}

@media(max-width:1000px){.kpis{grid-template-columns:repeat(1,1fr)}}
@media(prefers-reduced-motion:reduce){*{transition:none!important}}
</style>

<div class="sticky-top-panel">
  <div class="wrap">
    <header>
      <div class="eyebrow">Reverse Chain · Returns Centre</div>
      <h1>RTO / RVP RI% Dashboard</h1>
      <div class="sub" id="subline"></div>
    </header>

    <div class="filters">
      <div class="row">
        <div class="fg"><label>Month <span style="font-weight:400;text-transform:none;opacity:.6">· pick one or more, none = all</span></label><div class="chips" id="f-month"></div></div>
        <div class="fg"><label>Business Unit <span style="font-weight:400;text-transform:none;opacity:.6">· pick one or more, none = all</span></label><div class="chips" id="f-bu"></div></div>
        <div class="fg"><label>Warehouse / Site</label><select id="f-wh"></select></div>
        <button class="btn" id="reset">Reset filters</button>
        <div class="scope" id="scope"></div>
      </div>
    </div>

    <div class="tabs">
      <button class="tab-btn active" data-metric="rto" id="tabRTO">RTO RI %</button>
      <button class="tab-btn" data-metric="rvp" id="tabRVP">RVP RI %</button>
    </div>
  </div>
</div>

<div class="scroll-container">
  <div class="wrap">
    <div class="kpis" id="kpiRow"></div>

    <section>
      <div class="shead"><span class="n">1</span><h2>Zone-wise, month on month</h2>
        <span class="h">RI% weighted by loss ÷ value, respects Month, Business Unit and Warehouse filters</span></div>
      <div class="card"><div class="tscroll"><table id="zoneTable"><thead></thead><tbody></tbody></table></div></div>
    </section>

    <div class="split">
      <section>
        <div class="shead"><span class="n">2</span><h2>Site-wise, month on month</h2>
          <span class="h">Click a row to filter the whole dashboard to that site (Overall to reset)</span></div>
        <div class="card"><div class="tscroll"><table id="siteTable"><thead></thead><tbody></tbody></table></div></div>
      </section>

      <section>
        <div class="shead"><span class="n">3</span><h2 id="buPanelTitle">Business unit breakdown</h2></div>
        <div class="card"><div class="tscroll"><table id="buTable"><thead></thead><tbody></tbody></table></div></div>
      </section>
    </div>
  </div>
</div>

<script>
const DATA = /*__RI_DATA__*/;
const ZONE_COLORS = {North:"North",South:"South",East:"East",West:"West"};
const state = { month: new Set(), bu: new Set(), wh: "All" };
let metric = "rto";
let selectedSite = "All";

// Sites hidden from the Site-wise table + Warehouse filter (data still counted
// in Zone-wise totals and KPI headers, since those aren't warehouse-scoped views)
const EXCLUDED_SITES = new Set(["bil","binola","rclite_noida01","rclite_noida02"]);
const visibleWarehouses = DATA.meta.warehouses.filter(w => !EXCLUDED_SITES.has(w.toLowerCase()));

// Business unit hidden from the KPI header, BU filter chips, and the BU-wise
// summary panel — its rows stay in Zone-wise / Site-wise totals (data is kept,
// just not surfaced in these three places)
const EXCLUDED_BU = "wrong procurement";
function visibleBUs(list){ return list.filter(b => b.toLowerCase() !== EXCLUDED_BU); }
function excludeHiddenBU(records){ return records.filter(r => r.bu.toLowerCase() !== EXCLUDED_BU); }

function fmtPct(v){ return (v===undefined||v===null||isNaN(v)) ? "—" : v.toFixed(2)+"%"; }
function ri(records){ let w=0,v=0; for(const r of records){w+=r.w; v+=r.v;} return v>0 ? (w/v)*100 : undefined; }

// RTO: >98.5% good, 98%-98.5% warn, <98% bad
// RVP: >97% good, 95%-97% warn, <95% bad
function pctClass(m, v){
  if(v===undefined || v===null || isNaN(v)) return "";
  if(m === "rto"){
    if(v > 98.5) return "good";
    if(v >= 98) return "warn";
    return "bad";
  }
  if(v > 97) return "good";
  if(v >= 95) return "warn";
  return "bad";
}
function pctCell(m, v){
  const cls = pctClass(m, v);
  return `<td class="${cls ? 'pct-'+cls : ''}">${fmtPct(v)}</td>`;
}

function applyFilters(records, {bu, month, wh}){
  return records.filter(r =>
    (bu.size===0 || bu.has(r.bu)) &&
    (month.size===0 || month.has(r.month)) &&
    (wh==="All" || r.wh===wh)
  );
}

function buildChips(){
  const mWrap = document.getElementById("f-month");
  mWrap.innerHTML = DATA.meta.months.map(m =>
    `<div class="chip ${state.month.has(m)?'on':''}" data-m="${m}">${m}</div>`).join("");
  mWrap.querySelectorAll(".chip").forEach(c => c.onclick = () => {
    const m = c.dataset.m;
    state.month.has(m) ? state.month.delete(m) : state.month.add(m);
    render();
  });

  const bus = visibleBUs(metric === "rto" ? DATA.rto_bus : DATA.rvp_bus);
  const buWrap = document.getElementById("f-bu");
  buWrap.innerHTML = bus.map(b =>
    `<div class="chip ${state.bu.has(b)?'on':''}" data-b="${b}">${b}</div>`).join("");
  buWrap.querySelectorAll(".chip").forEach(c => c.onclick = () => {
    const b = c.dataset.b;
    state.bu.has(b) ? state.bu.delete(b) : state.bu.add(b);
    render();
  });

  const whSel = document.getElementById("f-wh");
  const prevWh = whSel.value || "All";
  whSel.innerHTML = '<option value="All">All warehouses</option>' +
    visibleWarehouses.map(w => `<option value="${w}">${w}</option>`).join("");
  whSel.value = visibleWarehouses.includes(prevWh) ? prevWh : "All";
}

function switchMetric(m){
  metric = m;
  document.getElementById("tabRTO").classList.toggle("active", m==="rto");
  document.getElementById("tabRVP").classList.toggle("active", m==="rvp");
  state.bu.clear();
  buildChips();
  render();
}

function renderKPIs(){
  const months = state.month.size ? DATA.meta.months.filter(m=>state.month.has(m)) : DATA.meta.months;
  const whSel = document.getElementById("f-wh").value;
  const lastMonth = months[months.length-1];
  const prevMonth = months.length>1 ? months[months.length-2] : null;

  const rtoAll = excludeHiddenBU(applyFilters(DATA.rto, {bu:state.bu, month:new Set(), wh:whSel}));
  const rvpAll = excludeHiddenBU(applyFilters(DATA.rvp, {bu:state.bu, month:new Set(), wh:whSel}));
  const rtoVal = ri(rtoAll.filter(r=>r.month===lastMonth));
  const rvpVal = ri(rvpAll.filter(r=>r.month===lastMonth));
  const rtoPrev = prevMonth ? ri(rtoAll.filter(r=>r.month===prevMonth)) : undefined;
  const rvpPrev = prevMonth ? ri(rvpAll.filter(r=>r.month===prevMonth)) : undefined;

  function delta(cur, prev){
    if(cur===undefined || prev===undefined || !prev) return "";
    const d = cur - prev;
    const up = d > 0;
    return `<span style="color:${up?'var(--up)':'var(--down)'}">${up?'▲':'▼'} ${Math.abs(d).toFixed(2)} pts vs ${prevMonth}</span>`;
  }

  document.getElementById("kpiRow").innerHTML = `
    <div class="kpi">
      <div class="k">RTO RI % (${lastMonth||"-"})</div>
      <div class="v ${pctClass("rto", rtoVal)}">${fmtPct(rtoVal)}</div>
      <div class="d">${delta(rtoVal, rtoPrev)}</div>
    </div>
    <div class="kpi">
      <div class="k">RVP RI % (${lastMonth||"-"})</div>
      <div class="v ${pctClass("rvp", rvpVal)}">${fmtPct(rvpVal)}</div>
      <div class="d">${delta(rvpVal, rvpPrev)}</div>
    </div>
    <div class="kpi">
      <div class="k">Rows processed</div>
      <div class="v">${DATA.meta.row_count.toLocaleString()}</div>
      <div class="d">${DATA.meta.months.length} months · ${DATA.meta.zones.length} zones</div>
    </div>
  `;
}

function render(){
  const wh = document.getElementById("f-wh").value;
  selectedSite = wh;  // top Warehouse filter drives every section
  const months = state.month.size ? DATA.meta.months.filter(m=>state.month.has(m)) : DATA.meta.months;
  const allRecords = DATA[metric] || [];

  renderKPIs();

  // Scope now reflects Month + BU only — Warehouse selection no longer hides rows,
  // it only steers which site the right-hand panel focuses on.
  const scopeRecords = applyFilters(allRecords, {bu:state.bu, month:state.month, wh:wh});
  document.getElementById("subline").innerText =
    `${DATA.meta.row_count.toLocaleString()} rows · ${DATA.meta.months.length} months (${DATA.meta.months.join(", ")}) · ${DATA.meta.zones.length} zones`;
  document.getElementById("scope").innerHTML =
    scopeRecords.length===0
      ? `<b style="color:var(--up)">In scope: 0 combinations</b> — no data matches this filter selection.`
      : `In scope: <b>${scopeRecords.length.toLocaleString()}</b> combinations · <b>${months.length}</b> month${months.length>1?'s':''}`;


  // --- Section 1: Zone-wise, month on month (Month + BU filters only) ---
  const zoneFiltered = applyFilters(allRecords, {bu:state.bu, month:new Set(), wh:wh});
  const zoneThead = document.querySelector("#zoneTable thead");
  const zoneTbody = document.querySelector("#zoneTable tbody");
  zoneThead.innerHTML = "<tr><th>Zone</th>" + months.map(m=>`<th>${m}</th>`).join("") + (months.length>1?"<th>MoM</th>":"") + "</tr>";
  const zoneRows = DATA.meta.zones.map(z => {
    const vals = months.map(m => ri(zoneFiltered.filter(r=>r.zone===z && r.month===m)));
    const cells = vals.map(v => pctCell(metric, v)).join("");
    let mom = "";
    if(months.length>1){
      const a = vals[vals.length-2], b = vals[vals.length-1];
      if(a!==undefined && b!==undefined){
        const d = b-a; const up = d>0;
        mom = `<td style="color:${up?'var(--up)':'var(--down)'}">${up?'▲':'▼'} ${Math.abs(d).toFixed(2)}</td>`;
      } else mom = "<td class='muted'>—</td>";
    }
    return `<tr class="zband ${z}"><td>${z}</td>${cells}${mom}</tr>`;
  }).join("");
  const zoneOverallVals = months.map(m => ri(zoneFiltered.filter(r=>r.month===m)));
  const zoneOverallCells = zoneOverallVals.map(v => pctCell(metric, v)).join("");
  let zoneOverallMom = "";
  if(months.length>1){
    const a = zoneOverallVals[zoneOverallVals.length-2], b = zoneOverallVals[zoneOverallVals.length-1];
    if(a!==undefined && b!==undefined){
      const d = b-a; const up = d>0;
      zoneOverallMom = `<td style="color:${up?'var(--up)':'var(--down)'}">${up?'▲':'▼'} ${Math.abs(d).toFixed(2)}</td>`;
    } else zoneOverallMom = "<td class='muted'>—</td>";
  }
  const zoneOverallRow = `<tr style="background:#f2f4f8;font-weight:600;"><td>Overall</td>${zoneOverallCells}${zoneOverallMom}</tr>`;
  zoneTbody.innerHTML = zoneOverallRow + zoneRows;

  // --- Section 2: Site-wise, month on month (Month + BU filters only — always shows ALL sites) ---
  const siteFiltered = applyFilters(allRecords, {bu:state.bu, month:new Set(), wh:wh});
  const siteThead = document.querySelector("#siteTable thead");
  const siteTbody = document.querySelector("#siteTable tbody");
  siteThead.innerHTML = "<tr><th>Warehouse</th>" + months.map(m=>`<th>${m}</th>`).join("") + "</tr>";
  const siteRows = visibleWarehouses.filter(s => wh==="All" || s===wh).map(s => {
    const zoneOf = (siteFiltered.find(r=>r.wh===s) || {}).zone || "";
    const cells = months.map(m => pctCell(metric, ri(siteFiltered.filter(r=>r.wh===s && r.month===m)))).join("");
    const sel = s === selectedSite ? "selected" : "";
    return `<tr class="zband ${zoneOf} ${sel}" onclick="focusSite('${s}')"><td>${s}</td>${cells}</tr>`;
  }).join("");
  siteTbody.innerHTML = `<tr style="background:#f2f4f8;font-weight:600;cursor:pointer;" onclick="focusSite('All')"><td>Overall (all sites)</td>${months.map(m=>pctCell(metric, ri(siteFiltered.filter(r=>r.month===m)))).join("")}</tr>` + siteRows;

  // --- Section 3: BU breakdown for the currently focused site ---
  const siteLabel = selectedSite === "All" ? "all sites" : selectedSite;
  document.getElementById("buPanelTitle").innerText = `Business unit breakdown — ${siteLabel}`;
  const buRecords = applyFilters(allRecords, {bu:state.bu, month:new Set(), wh:selectedSite});
  const bus = visibleBUs(metric === "rto" ? DATA.rto_bus : DATA.rvp_bus).filter(b => state.bu.size===0 || state.bu.has(b));
  const buThead = document.querySelector("#buTable thead");
  const buTbody = document.querySelector("#buTable tbody");
  buThead.innerHTML = "<tr><th>Business unit</th>" + months.map(m=>`<th>${m}</th>`).join("") + "</tr>";
  buTbody.innerHTML = bus.map(b => {
    const cells = months.map(m => pctCell(metric, ri(buRecords.filter(r=>r.bu===b && r.month===m)))).join("");
    return `<tr><td>${b}</td>${cells}</tr>`;
  }).join("") || "<tr><td colspan='99' class='muted' style='text-align:center'>No data</td></tr>";
}

function focusSite(s){
  selectedSite = s;
  const whSel = document.getElementById("f-wh");
  if ([...whSel.options].some(o=>o.value===s)) whSel.value = s;
  render();
}

document.getElementById("f-wh").addEventListener("change", render);
document.getElementById("reset").addEventListener("click", () => {
  state.month.clear(); state.bu.clear(); state.wh = "All"; selectedSite = "All";
  document.getElementById("f-wh").value = "All";
  buildChips();
  render();
});
document.getElementById("tabRTO").addEventListener("click", () => switchMetric("rto"));
document.getElementById("tabRVP").addEventListener("click", () => switchMetric("rvp"));

buildChips();
render();
</script>
"""


def main():
    st.set_page_config(page_title="RTO / RVP RI% Dashboard", layout="wide")
    st.markdown(
        "<style>.block-container{padding:0rem 0.5rem;max-width:100%}"
        "header[data-testid='stHeader']{height:0; display:none;}"
        "[data-testid='stAppViewBlockContainer']{overflow:hidden;}"
        "iframe{border:none !important; height:95vh !important;}</style>",
        unsafe_allow_html=True)

    height = st.sidebar.slider("Viewport height (px)", 600, 2000, 920, 40)

    if st.sidebar.button("🔄 Refresh data from Drive"):
        fetch_and_combine_from_drive.clear()

    try:
        raw_df, file_names = fetch_and_combine_from_drive(DRIVE_FOLDER_ID)
        st.sidebar.caption(f"Loaded {len(file_names)} file(s):")
        for n in file_names:
            st.sidebar.caption(f"• {n}")
    except Exception as e:
        st.error(f"Could not fetch data from Google Drive: {e}")
        st.stop()

    try:
        df = prepare_data(raw_df)
        payload = build_payload(df)
    except Exception as e:
        st.error(f"Could not process the data: {e}")
        st.stop()

    data_str = json.dumps(payload, separators=(",", ":"))
    print(f"payload JSON size: {len(data_str)/1e6:.2f} MB")
    html = TEMPLATE.replace("/*__RI_DATA__*/", data_str)
    components.html(html, height=height, scrolling=False)


if __name__ == "__main__":
    main()
