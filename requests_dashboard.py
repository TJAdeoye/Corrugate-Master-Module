"""
requests_dashboard.py - HK1 corrugate request history + delivery ETA dashboard.

Reads HK1_Corrugate_Request_File.xlsx, where each worksheet is one order date
("HK 9-8-26"), and writes a self-contained HTML dashboard showing every order
date, the SKUs requested on it, and the delivery ETA assigned to it.

    python requests_dashboard.py --file "HK1_Corrugate_Request_File.xlsx"
    python requests_dashboard.py --folder "<folder>"      # newest matching file

Notebook:
    import requests_dashboard as rq
    result = rq.run(file=r'...\\HK1_Corrugate_Request_File.xlsx')

Sheets carry: SKU · Burn/Day · On Hand · Days Cover · Status · Request
and, where it has been filled in, ETA.
"""

from __future__ import annotations
import argparse
import json
import re
from datetime import datetime, date
from pathlib import Path

import openpyxl

try:
    import yaml
except ImportError:
    yaml = None

DEFAULT_CFG = {
    "paths": {"source_folder": ".", "file": None,
              "output_html": "HK1_Request_ETA_Dashboard.html", "archive_folder": None},
    "source": {"glob": "*.xls*"},
    "display": {"title": "Order & ETA Tracker", "title_emphasis": "ETA",
                "eyebrow": "Corrugate requests", "show_parse_warnings": False},
    "sites": {},
}


def deep_merge(base: dict, over: dict) -> dict:
    out = dict(base)
    for k, v in (over or {}).items():
        out[k] = deep_merge(base[k], v) if isinstance(v, dict) and isinstance(base.get(k), dict) else v
    return out


def load_config(path: Path | None) -> dict:
    if path is None:
        return DEFAULT_CFG
    if yaml is None:
        raise SystemExit("PyYAML is required to read a config file:  pip install pyyaml")
    with open(path, "r", encoding="utf-8") as fh:
        return deep_merge(DEFAULT_CFG, yaml.safe_load(fh) or {})

SHEET_DATE = re.compile(r"(\d{1,2})\s*[-/]\s*(\d{1,2})\s*[-/]\s*(\d{2,4})")

# Sheet-name prefix -> site. Matched on the letters before the date,
# uppercased with spaces and punctuation stripped.
SITE_PREFIX = {
    "HK": "HK1", "HK1": "HK1", "HODGKINS": "HK1", "HODGKINS1": "HK1",
    "BP": "BP1", "BP1": "BP1", "BEDFORD": "BP1", "BEDFORDPARK": "BP1", "BEDFORDPARK1": "BP1",
}
SITE_NAMES = {"HK1": "Hodgkins 1", "BP1": "Bedford Park 1"}


def sheet_to_site(name: str, prefixes: dict = None):
    """Read the site from the sheet name: 'BP 9-8-26' -> ('BP1', False).

    Returns (site, is_guess). An unrecognised prefix comes back as the raw
    prefix with is_guess True, so it is reported rather than silently binned
    into the wrong site.
    """
    m = SHEET_DATE.search(name)
    head = name[: m.start()] if m else name
    key = re.sub(r"[^A-Z0-9]", "", head.upper())
    table = prefixes if prefixes is not None else SITE_PREFIX
    if key in table:
        return table[key], False
    return (key or "?"), True

# Header label -> internal field. Matched case-insensitively, spaces stripped.
HEADER_MAP = {
    "sku": "sku",
    "burn/day": "burn", "burnday": "burn", "burnrate": "burn",
    "onhand": "onHand",
    "dayscover": "daysCover", "daysofcover": "daysCover",
    "status": "status",
    "request": "request", "requested": "request", "qty": "request",
    "eta": "eta", "etadelivery": "eta", "deliveryeta": "eta",
}


def norm_hdr(s) -> str:
    return re.sub(r"[^a-z0-9/]", "", str(s or "").lower())


def sheet_to_date(name: str):
    m = SHEET_DATE.search(name)
    if not m:
        return None
    mo, dy, yr = (int(g) for g in m.groups())
    yr += 2000 if yr < 100 else 0
    try:
        return date(yr, mo, dy)
    except ValueError:
        return None


def cell_date(v):
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    if v is None or not str(v).strip():
        return None
    m = SHEET_DATE.search(str(v))
    if m:
        mo, dy, yr = (int(g) for g in m.groups())
        yr += 2000 if yr < 100 else 0
        try:
            return date(yr, mo, dy).isoformat()
        except ValueError:
            return None
    return None


def num(v):
    if isinstance(v, (int, float)):
        return float(v)
    if v is None:
        return None
    m = re.search(r"-?\d[\d,]*(?:\.\d+)?", str(v))
    return float(m.group(0).replace(",", "")) if m else None


def parse_workbook(path: Path, cfg: dict = None):
    cfg = cfg or DEFAULT_CFG
    extra = {re.sub(r"[^A-Z0-9]", "", str(k).upper()): v
             for k, v in (cfg.get("sites") or {}).items()}
    prefixes = {**SITE_PREFIX, **extra}

    wb = openpyxl.load_workbook(path, data_only=True)
    orders, warnings = [], []

    for name in wb.sheetnames:
        ws = wb[name]
        order_dt = sheet_to_date(name)
        if order_dt is None:
            warnings.append(f"Sheet “{name}” has no readable date in its name - skipped.")
            continue

        # map header row
        cols = {}
        for c in range(1, ws.max_column + 1):
            key = HEADER_MAP.get(norm_hdr(ws.cell(1, c).value))
            if key and key not in cols:
                cols[key] = c
        if "sku" not in cols:
            warnings.append(f"Sheet “{name}” has no SKU column - skipped.")
            continue

        site, guessed = sheet_to_site(name, prefixes)
        if guessed:
            warnings.append(
                f"Sheet “{name.strip()}” - could not tell which DC from the name "
                f"(expected a HK or BP prefix). Shown as “{site}”; rename the tab "
                f"or add the prefix to SITE_PREFIX.")

        lines = []
        for r in range(2, ws.max_row + 1):
            sku = ws.cell(r, cols["sku"]).value
            if sku is None or not str(sku).strip():
                continue
            get = lambda k: ws.cell(r, cols[k]).value if k in cols else None
            lines.append({
                "sku": str(sku).strip(),
                "burn": num(get("burn")),
                "onHand": num(get("onHand")),
                "daysCover": num(get("daysCover")),
                "status": (str(get("status")).strip() if get("status") else None),
                "request": num(get("request")),
                "eta": cell_date(get("eta")),
            })

        if not lines:
            warnings.append(f"Sheet “{name.strip()}” ({order_dt}) has a header but no line items.")

        etas = sorted({l["eta"] for l in lines if l["eta"]})
        orders.append({
            "sheet": name.strip(),
            "site": site,
            "siteName": SITE_NAMES.get(site, site),
            "orderDate": order_dt.isoformat(),
            "lines": lines,
            "skuCount": len(lines),
            "units": sum(l["request"] or 0 for l in lines),
            "etas": etas,
            "hasEta": bool(etas),
            "etaColumn": "eta" in cols,
        })

    orders.sort(key=lambda o: (o["orderDate"], o["site"]), reverse=True)
    return orders, warnings


HTML = """<!DOCTYPE html>
<html lang="en"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>HK1 Corrugate Requests · Order &amp; ETA Tracker</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;700&family=Fraunces:ital,wght@0,400;0,600;0,800;1,400&display=swap" rel="stylesheet">
<style>
:root{--bg:#0f1114;--bg-panel:#16191e;--bg-elev:#1d2128;--border:#272c35;--border-bright:#3a414e;
--text:#e8e9ec;--text-dim:#8a8f99;--text-faint:#5a5f6a;--amber:#f4a52a;--red:#ef4444;--green:#10b981;--blue:#60a5fa;}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'JetBrains Mono',Consolas,monospace;background:var(--bg);color:var(--text);font-size:13px;line-height:1.5;
background-image:radial-gradient(circle at 15% 0%,rgba(244,165,42,.035),transparent 42%),radial-gradient(circle at 85% 100%,rgba(96,165,250,.03),transparent 42%);}
.container{max-width:1320px;margin:0 auto;padding:28px 32px 90px}
header{border-bottom:1px solid var(--border);padding-bottom:20px;margin-bottom:22px;display:grid;grid-template-columns:1fr auto;gap:24px;align-items:end}
.eyebrow{font-size:10px;letter-spacing:.2em;color:var(--text-faint);text-transform:uppercase;margin-bottom:8px;display:flex;align-items:center;gap:10px}
.eyebrow::before{content:'';width:7px;height:7px;border-radius:50%;background:var(--green);box-shadow:0 0 8px var(--green)}
h1{font-family:'Fraunces',serif;font-weight:800;font-size:36px;letter-spacing:-.02em;line-height:1}
h1 em{font-style:italic;font-weight:400;color:var(--amber)}
.header-meta{text-align:right;font-size:11px;color:var(--text-dim);line-height:1.9}
.header-meta .label{color:var(--text-faint);text-transform:uppercase;letter-spacing:.1em}
.header-meta .value{color:var(--text);font-weight:500}
.tabs{display:flex;gap:2px;margin-bottom:22px;flex-wrap:wrap}
.tab{background:transparent;border:1px solid var(--border);color:var(--text-dim);font-family:inherit;font-size:11px;
letter-spacing:.12em;text-transform:uppercase;padding:10px 18px;cursor:pointer;transition:all .15s}
.tab:hover{color:var(--text);border-color:var(--border-bright)}
.tab.active{background:var(--text);color:var(--bg);border-color:var(--text);font-weight:700}
.kpi-row{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:26px}
.kpi{background:var(--bg-panel);border:1px solid var(--border);padding:18px 20px;position:relative;overflow:hidden}
.kpi::before{content:'';position:absolute;top:0;left:0;width:3px;height:100%;background:var(--border-bright)}
.kpi.warn::before{background:var(--amber)}
.kpi-label{font-size:10px;color:var(--text-faint);text-transform:uppercase;letter-spacing:.12em;margin-bottom:8px}
.kpi-value{font-family:'Fraunces',serif;font-size:32px;font-weight:600;line-height:1}
.kpi-value.warn{color:var(--amber)}
.kpi-value small{font-size:16px;color:var(--text-faint)}
.kpi-sub{font-size:11px;color:var(--text-dim);margin-top:6px}
.group-head{display:flex;align-items:baseline;gap:14px;margin-bottom:14px}
.group-head h3{font-family:'Fraunces',serif;font-size:18px;font-weight:600}
.group-head .tagline{font-size:10px;color:var(--text-faint);letter-spacing:.12em;text-transform:uppercase}
.group-head .filler{margin-left:auto;font-size:11px;color:var(--text-faint)}
.tile-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(248px,1fr));gap:14px}
.tile{background:var(--bg-panel);border:1px solid var(--border);padding:18px 18px 14px;cursor:pointer;text-align:left;
font-family:inherit;color:inherit;position:relative;transition:border-color .15s,transform .15s,background .15s}
.tile:hover{border-color:var(--border-bright);background:var(--bg-elev);transform:translateY(-2px)}
.tile::before{content:'';position:absolute;top:0;left:0;width:3px;height:100%;background:var(--border-bright)}
.tile.hk1::before{background:var(--amber)}
.tile.bp1::before{background:var(--blue)}
.tile-top{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:12px;gap:10px}
.tile-date{font-family:'Fraunces',serif;font-size:21px;font-weight:600;letter-spacing:-.01em;line-height:1.1}
.tile-sheet{font-size:10px;color:var(--text-faint);text-transform:uppercase;letter-spacing:.1em;margin-top:3px}
.tile-stats{display:flex;gap:20px;padding:12px 0 10px;border-top:1px dashed var(--border)}
.tile-stat-label{font-size:9px;color:var(--text-faint);text-transform:uppercase;letter-spacing:.1em}
.tile-stat-val{font-size:17px;font-weight:500;margin-top:2px}
.tile-foot{border-top:1px solid var(--border);padding-top:10px;font-size:10px;color:var(--text-faint);display:flex;justify-content:space-between;gap:8px}
.tile-foot b{color:var(--green);font-weight:500}
.tile-foot .none{color:var(--text-faint)}
.detail-bar{display:flex;gap:14px;align-items:center;margin-bottom:20px;flex-wrap:wrap}
.btn{font-family:inherit;font-size:11px;letter-spacing:.12em;text-transform:uppercase;padding:11px 22px;cursor:pointer;
border:1px solid var(--border-bright);background:transparent;color:var(--text-dim);transition:all .15s}
.btn:hover{color:var(--text);border-color:var(--text-dim)}
.panel{background:var(--bg-panel);border:1px solid var(--border);margin-bottom:20px}
.panel-head{display:flex;justify-content:space-between;align-items:center;padding:14px 20px;border-bottom:1px solid var(--border);gap:14px;flex-wrap:wrap}
.panel-head h4{font-family:'Fraunces',serif;font-size:15px;font-weight:600}
.panel-head .stamp{font-size:11px;color:var(--text-faint)}
.panel-head .stamp b{color:var(--text-dim);font-weight:500}
.filter-btns{display:flex;gap:4px;flex-wrap:wrap}
.filter-btn{background:transparent;border:1px solid var(--border);color:var(--text-dim);font-family:inherit;font-size:10px;
letter-spacing:.1em;text-transform:uppercase;padding:5px 10px;cursor:pointer;transition:all .15s}
.filter-btn:hover{color:var(--text);border-color:var(--border-bright)}
.filter-btn.active{background:var(--text);color:var(--bg);border-color:var(--text);font-weight:700}
table{width:100%;border-collapse:collapse;font-size:12px}
thead tr{background:var(--bg-elev)}
th{text-align:left;padding:10px 14px;font-size:10px;font-weight:500;color:var(--text-faint);text-transform:uppercase;
letter-spacing:.1em;border-bottom:1px solid var(--border);cursor:pointer;user-select:none;white-space:nowrap}
th:hover{color:var(--text)}
th.num,td.num{text-align:right}
th .arrow{opacity:0;margin-left:4px;font-size:9px}
th.sorted .arrow{opacity:1;color:var(--amber)}
td{padding:9px 14px;border-bottom:1px solid var(--border)}
tbody tr:hover{background:rgba(255,255,255,.015)}
tbody tr:last-child td{border-bottom:none}
td.sku{color:var(--text);font-weight:500}
td.faint{color:var(--text-faint)}
td.eta{color:var(--green)}
td.crit{color:var(--red);font-weight:500}
.pill{display:inline-block;padding:2px 8px;font-size:10px;font-weight:500;letter-spacing:.08em;text-transform:uppercase;border:1px solid}
.pill.crit{color:var(--red);border-color:var(--red)}
.pill.hk1{color:var(--amber);border-color:var(--amber)}
.pill.bp1{color:var(--blue);border-color:var(--blue)}
.pill.unknown{color:var(--text-faint);border-color:var(--border)}
.sitebar{display:flex;gap:4px;align-items:center;margin-bottom:20px;flex-wrap:wrap}
.sitebar .lbl{font-size:10px;color:var(--text-faint);text-transform:uppercase;letter-spacing:.12em;margin-right:8px}
.pill.eta{color:var(--green);border-color:var(--green)}
.pill.noeta{color:var(--text-faint);border-color:var(--border)}
.bar{height:6px;background:var(--bg-elev);position:relative;min-width:60px}
.bar span{display:block;height:100%;background:var(--amber)}
.empty{padding:44px 20px;text-align:center;color:var(--text-faint);font-size:12px;line-height:2}
.empty b{display:block;color:var(--text-dim);font-family:'Fraunces',serif;font-size:16px;font-weight:600}
.warnlist{margin-top:26px;padding:16px 20px;border-left:2px solid var(--red);background:rgba(239,68,68,.05);
font-size:11px;color:var(--text-dim);line-height:1.9}
.warnlist b{color:var(--red);text-transform:uppercase;letter-spacing:.12em;font-size:10px}
.warnlist ul{margin:6px 0 0 16px}
.hidden{display:none!important}
@media(max-width:1000px){.kpi-row{grid-template-columns:repeat(2,1fr)}}
@media(max-width:640px){.container{padding:18px 16px 60px}header{grid-template-columns:1fr}.header-meta{text-align:left}
h1{font-size:27px}.kpi-row{grid-template-columns:1fr}.tile-grid{grid-template-columns:1fr}th,td{padding:8px 10px}}
</style></head><body><div class="container">

<header>
  <div>
    <div class="eyebrow">__EYEBROW__</div>
    <h1>__TITLE_HTML__</h1>
  </div>
  <div class="header-meta">
    <div><span class="label">Source</span> &nbsp; <span class="value">__SRCFILE__</span></div>
    <div><span class="label">Order dates</span> &nbsp; <span class="value">__NORDERS__</span></div>
    <div><span class="label">Generated</span> &nbsp; <span class="value">__GENERATED__</span></div>
  </div>
</header>

<div class="tabs">
  <button class="tab active" data-view="orders">Order dates</button>
  <button class="tab" data-view="lines">All requests</button>
  <button class="tab" data-view="skus">By SKU</button>
</div>

<div class="sitebar" id="sitebar">
  <span class="lbl">Distribution center</span>
  <button class="filter-btn active" data-site="all">Both</button>
  <button class="filter-btn" data-site="HK1">Hodgkins 1</button>
  <button class="filter-btn" data-site="BP1">Bedford Park 1</button>
</div>

<div class="kpi-row" id="kpis"></div>

<!-- ORDERS -->
<section id="view-orders">
  <div id="orders-grid">
    <div class="group-head">
      <h3>Order dates</h3><span class="tagline">Newest first</span>
      <span class="filler" id="orders-n"></span>
    </div>
    <div class="tile-grid" id="tiles"></div>
  </div>

  <div id="orders-detail" class="hidden">
    <div class="detail-bar"><button class="btn" id="btn-back">← All order dates</button></div>
    <div class="panel">
      <div class="panel-head"><h4 id="d-title">Order</h4><span class="stamp" id="d-stamp"></span></div>
      <div id="d-body"></div>
    </div>
  </div>
</section>

<!-- ALL LINES -->
<section id="view-lines" class="hidden">
  <div class="panel">
    <div class="panel-head">
      <h4>Every requested line</h4>
      <div class="filter-btns">
        <button class="filter-btn active" data-f="all">All</button>
        <button class="filter-btn" data-f="eta">Has ETA</button>
        <button class="filter-btn" data-f="noeta">No ETA</button>
        <button class="filter-btn" data-f="zero">Was at zero</button>
      </div>
    </div>
    <table>
      <thead><tr>
        <th class="sorted" data-sort="orderDate">Ordered <span class="arrow">▼</span></th>
        <th data-sort="site">DC <span class="arrow">▼</span></th>
        <th data-sort="sku">SKU <span class="arrow">▼</span></th>
        <th class="num" data-sort="request">Requested <span class="arrow">▼</span></th>
        <th class="num" data-sort="onHand">On hand <span class="arrow">▼</span></th>
        <th class="num" data-sort="burn">Burn / day <span class="arrow">▼</span></th>
        <th class="num" data-sort="daysCover">Days cover <span class="arrow">▼</span></th>
        <th data-sort="status">Status <span class="arrow">▼</span></th>
        <th data-sort="eta">Delivery ETA <span class="arrow">▼</span></th>
      </tr></thead>
      <tbody id="lines-body"></tbody>
    </table>
  </div>
</section>

<!-- BY SKU -->
<section id="view-skus" class="hidden">
  <div class="panel">
    <div class="panel-head">
      <h4>Request history by SKU</h4>
      <span class="stamp">Which boxes keep coming back</span>
    </div>
    <table>
      <thead><tr>
        <th data-s="sku">SKU <span class="arrow">▼</span></th>
        <th class="num sorted" data-s="times">Times ordered <span class="arrow">▼</span></th>
        <th class="num" data-s="units">Total units <span class="arrow">▼</span></th>
        <th class="num" data-s="hk1">HK1 units <span class="arrow">▼</span></th>
        <th class="num" data-s="bp1">BP1 units <span class="arrow">▼</span></th>
        <th class="num" data-s="avg">Avg per order <span class="arrow">▼</span></th>
        <th data-s="last">Last ordered <span class="arrow">▼</span></th>
        <th style="width:170px;">Share of all requests</th>
      </tr></thead>
      <tbody id="skus-body"></tbody>
    </table>
  </div>
</section>

__WARNINGS__

</div>
<script>
const ORDERS = __ORDERS__;
const ALL_LINES = [];
ORDERS.forEach(o => o.lines.forEach(l => ALL_LINES.push(
  {...l, orderDate:o.orderDate, sheet:o.sheet, site:o.site, siteName:o.siteName})));

let SITE = 'all';
const siteClass = s => (s === 'HK1' || s === 'BP1') ? s.toLowerCase() : 'unknown';
const inSite = x => SITE === 'all' || x.site === SITE;
const orders = () => ORDERS.filter(inSite);
const lines  = () => ALL_LINES.filter(inSite);

const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const fmt = n => (n === null || n === undefined) ? '-' : Number(n).toLocaleString('en-US');
const fmt1 = n => (n === null || n === undefined) ? '-' : Number(n).toFixed(1);
const pretty = iso => { if(!iso) return '-'; const [y,m,d]=iso.split('-').map(Number);
  return new Date(y,m-1,d).toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'}); };
const prettyShort = iso => { if(!iso) return '-'; const [y,m,d]=iso.split('-').map(Number);
  return new Date(y,m-1,d).toLocaleDateString('en-US',{month:'short',day:'numeric'}); };

/* ---- KPIs ---- */
function renderKpis(){
  const O = orders(), L = lines();
  const withEta = O.filter(o => o.hasEta);
  const totalUnits = O.reduce((s,o) => s + o.units, 0);
  const uniqueSkus = new Set(L.map(l => l.sku)).size;
  const hk = ORDERS.filter(o => o.site === 'HK1').length;
  const bp = ORDERS.filter(o => o.site === 'BP1').length;
  const span = O.length
    ? `${pretty(O[O.length-1].orderDate)} → ${pretty(O[0].orderDate)}`
    : 'no orders for this filter';

  document.getElementById('kpis').innerHTML = `
    <div class="kpi">
      <div class="kpi-label">Order dates</div>
      <div class="kpi-value">${O.length}</div>
      <div class="kpi-sub">${span}</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Lines requested</div>
      <div class="kpi-value">${L.length}</div>
      <div class="kpi-sub">across ${uniqueSkus} distinct SKUs</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Total units</div>
      <div class="kpi-value">${fmt(totalUnits)}</div>
      <div class="kpi-sub">${SITE === 'all' ? `HK1 ${hk} orders · BP1 ${bp} orders` : 'bundles requested'}</div>
    </div>
    <div class="kpi ${withEta.length < O.length ? 'warn' : ''}">
      <div class="kpi-label">Orders with an ETA</div>
      <div class="kpi-value ${withEta.length < O.length ? 'warn' : ''}">${withEta.length}<small> / ${O.length}</small></div>
      <div class="kpi-sub">${O.length - withEta.length} have no delivery date recorded</div>
    </div>`;
}

/* ---- tabs ---- */
document.querySelectorAll('.tab').forEach(t => t.addEventListener('click', () => {
  document.querySelectorAll('.tab').forEach(x => x.classList.remove('active'));
  t.classList.add('active');
  ['orders','lines','skus'].forEach(v => document.getElementById('view-'+v).classList.toggle('hidden', v !== t.dataset.view));
}));

/* ---- order tiles ---- */
function renderTiles(){
  const O = orders();
  document.getElementById('orders-n').textContent =
    O.length + (O.length === 1 ? ' order date' : ' order dates');

  if (!O.length){
    document.getElementById('tiles').innerHTML =
      `<div class="empty" style="grid-column:1/-1;"><b>No orders for this center yet</b>
       Sheets named with a BP prefix will appear here automatically.</div>`;
    return;
  }

  document.getElementById('tiles').innerHTML = O.map(o => `
    <button class="tile ${siteClass(o.site)}" data-sheet="${esc(o.sheet)}">
      <div class="tile-top">
        <div>
          <div class="tile-date">${esc(pretty(o.orderDate))}</div>
          <div class="tile-sheet">${esc(o.sheet)}</div>
        </div>
        <span class="pill ${siteClass(o.site)}">${esc(o.site)}</span>
      </div>
      <div class="tile-stats">
        <div><div class="tile-stat-label">SKUs</div><div class="tile-stat-val">${o.skuCount}</div></div>
        <div><div class="tile-stat-label">Units</div><div class="tile-stat-val">${fmt(o.units)}</div></div>
      </div>
      <div class="tile-foot">
        <span>Delivery ETA</span>
        ${o.hasEta ? `<b>${o.etas.map(prettyShort).join(', ')}</b>` : `<span class="none">not assigned</span>`}
      </div>
    </button>`).join('');

  document.querySelectorAll('.tile').forEach(el => el.addEventListener('click', () => {
    showOrder(ORDERS.findIndex(o => o.sheet === el.dataset.sheet));
  }));
}

function showOrder(i){
  const o = ORDERS[i];
  document.getElementById('orders-grid').classList.add('hidden');
  document.getElementById('orders-detail').classList.remove('hidden');
  document.getElementById('d-title').textContent =
    o.siteName + ' · ordered ' + pretty(o.orderDate);
  document.getElementById('d-stamp').innerHTML =
    `<span class="pill ${siteClass(o.site)}">${esc(o.site)}</span> &nbsp; ` +
    `${o.skuCount} SKUs &nbsp;·&nbsp; ${fmt(o.units)} units &nbsp;·&nbsp; ETA ` +
    (o.hasEta ? `<b>${o.etas.map(pretty).join(', ')}</b>` : '<span style="color:var(--text-faint)">not assigned</span>');

  document.getElementById('d-body').innerHTML = `
    <table><thead><tr>
      <th>SKU</th><th class="num">Requested</th><th class="num">On hand</th>
      <th class="num">Burn / day</th><th class="num">Days cover</th><th>Status</th><th>Delivery ETA</th>
    </tr></thead><tbody>
    ${o.lines.map(l => `<tr>
      <td class="sku">${esc(l.sku)}</td>
      <td class="num">${fmt(l.request)}</td>
      <td class="num ${l.onHand === 0 ? 'crit' : ''}">${fmt(l.onHand)}</td>
      <td class="num">${fmt(l.burn)}</td>
      <td class="num ${l.daysCover !== null && l.daysCover <= 1 ? 'crit' : ''}">${fmt1(l.daysCover)}</td>
      <td>${l.status ? `<span class="pill crit">${esc(l.status)}</span>` : '<span class="faint">-</span>'}</td>
      <td class="${l.eta ? 'eta' : 'faint'}">${l.eta ? esc(pretty(l.eta)) : 'not assigned'}</td>
    </tr>`).join('')}
    </tbody></table>`;
  window.scrollTo({top:0, behavior:'smooth'});
}
document.getElementById('btn-back').addEventListener('click', () => {
  document.getElementById('orders-detail').classList.add('hidden');
  document.getElementById('orders-grid').classList.remove('hidden');
});

/* ---- all lines ---- */
let lSort = 'orderDate', lDir = -1, lFilter = 'all';
function renderLines(){
  let rows = lines();
  if (lFilter === 'eta')   rows = rows.filter(l => l.eta);
  if (lFilter === 'noeta') rows = rows.filter(l => !l.eta);
  if (lFilter === 'zero')  rows = rows.filter(l => l.onHand === 0);

  rows.sort((a,b) => {
    let av = a[lSort], bv = b[lSort];
    if (av === null || av === undefined) av = lDir === 1 ? Infinity : -Infinity;
    if (bv === null || bv === undefined) bv = lDir === 1 ? Infinity : -Infinity;
    if (av < bv) return -lDir;
    if (av > bv) return lDir;
    return 0;
  });

  document.getElementById('lines-body').innerHTML = rows.map(l => `<tr>
    <td>${esc(pretty(l.orderDate))}</td>
    <td><span class="pill ${siteClass(l.site)}">${esc(l.site)}</span></td>
    <td class="sku">${esc(l.sku)}</td>
    <td class="num">${fmt(l.request)}</td>
    <td class="num ${l.onHand === 0 ? 'crit' : ''}">${fmt(l.onHand)}</td>
    <td class="num">${fmt(l.burn)}</td>
    <td class="num ${l.daysCover !== null && l.daysCover <= 1 ? 'crit' : ''}">${fmt1(l.daysCover)}</td>
    <td>${l.status ? `<span class="pill crit">${esc(l.status)}</span>` : '<span class="faint">-</span>'}</td>
    <td class="${l.eta ? 'eta' : 'faint'}">${l.eta ? esc(pretty(l.eta)) : 'not assigned'}</td>
  </tr>`).join('');

  document.querySelectorAll('#view-lines th').forEach(th => {
    th.classList.toggle('sorted', th.dataset.sort === lSort);
    const a = th.querySelector('.arrow'); if (a) a.textContent = lDir === 1 ? '▲' : '▼';
  });
}
document.querySelectorAll('#view-lines th').forEach(th => th.addEventListener('click', () => {
  if (lSort === th.dataset.sort) lDir *= -1; else { lSort = th.dataset.sort; lDir = -1; }
  renderLines();
}));
document.querySelectorAll('#view-lines .filter-btn').forEach(b => b.addEventListener('click', () => {
  document.querySelectorAll('#view-lines .filter-btn').forEach(x => x.classList.remove('active'));
  b.classList.add('active'); lFilter = b.dataset.f; renderLines();
}));
renderLines();

/* ---- by SKU ---- */
let sSort = 'times', sDir = -1;

function renderSkus(){
  const bySku = {};
  lines().forEach(l => {
    const s = bySku[l.sku] || (bySku[l.sku] = {sku:l.sku, times:0, units:0, hk1:0, bp1:0, last:null});
    s.times++; s.units += (l.request || 0);
    if (l.site === 'HK1') s.hk1 += (l.request || 0);
    if (l.site === 'BP1') s.bp1 += (l.request || 0);
    if (!s.last || l.orderDate > s.last) s.last = l.orderDate;
  });
  const rowsAll = Object.values(bySku).map(s => ({...s, avg: s.units / s.times}));
  const maxUnits = Math.max(...rowsAll.map(s => s.units), 1);

  if (!rowsAll.length){
    document.getElementById('skus-body').innerHTML =
      '<tr><td colspan="7"><div class="empty">Nothing requested for this center yet.</div></td></tr>';
    return;
  }

  const rows = rowsAll.sort((a,b) => {
    const av = a[sSort], bv = b[sSort];
    if (av < bv) return -sDir; if (av > bv) return sDir; return 0;
  });

  document.getElementById('skus-body').innerHTML = rows.map(s => `<tr>
    <td class="sku">${esc(s.sku)}</td>
    <td class="num">${s.times}</td>
    <td class="num">${fmt(s.units)}</td>
    <td class="num" style="color:var(--amber)">${s.hk1 ? fmt(s.hk1) : '<span class="faint">-</span>'}</td>
    <td class="num" style="color:var(--blue)">${s.bp1 ? fmt(s.bp1) : '<span class="faint">-</span>'}</td>
    <td class="num">${s.avg.toFixed(1)}</td>
    <td class="faint">${esc(pretty(s.last))}</td>
    <td><div class="bar"><span style="width:${(s.units/maxUnits*100).toFixed(1)}%"></span></div></td>
  </tr>`).join('');

  document.querySelectorAll('#view-skus th').forEach(th => {
    th.classList.toggle('sorted', th.dataset.s === sSort);
    const a = th.querySelector('.arrow'); if (a) a.textContent = sDir === 1 ? '▲' : '▼';
  });
}
document.querySelectorAll('#view-skus th').forEach(th => th.addEventListener('click', () => {
  if (!th.dataset.s) return;
  if (sSort === th.dataset.s) sDir *= -1; else { sSort = th.dataset.s; sDir = -1; }
  renderSkus();
}));

/* ---- site filter drives everything ---- */
document.querySelectorAll('#sitebar .filter-btn').forEach(b => b.addEventListener('click', () => {
  document.querySelectorAll('#sitebar .filter-btn').forEach(x => x.classList.remove('active'));
  b.classList.add('active');
  SITE = b.dataset.site;
  document.getElementById('orders-detail').classList.add('hidden');
  document.getElementById('orders-grid').classList.remove('hidden');
  renderKpis(); renderTiles(); renderLines(); renderSkus();
}));

renderKpis();
renderTiles();
renderSkus();
</script></body></html>
"""


def build(orders, warnings, src: Path, cfg: dict = None):
    cfg = cfg or DEFAULT_CFG
    disp = cfg["display"]

    warn_html = ""
    if warnings and disp.get("show_parse_warnings", True):
        lis = "".join(f"<li>{w}</li>" for w in warnings)
        warn_html = f'<div class="warnlist"><b>Parse warnings - {len(warnings)}</b><ul>{lis}</ul></div>'

    title = disp["title"]
    emph = disp.get("title_emphasis")
    title_html = title.replace(emph, f"<em>{emph}</em>", 1) if emph and emph in title else title

    return (HTML
            .replace("__ORDERS__", json.dumps(orders))
            .replace("__TITLE_HTML__", title_html)
            .replace("__EYEBROW__", disp.get("eyebrow", "Corrugate requests"))
            .replace("__SRCFILE__", src.name)
            .replace("__NORDERS__", str(len(orders)))
            .replace("__GENERATED__", datetime.now().strftime("%b %d, %Y %I:%M %p").replace(" 0", " "))
            .replace("__WARNINGS__", warn_html))


def run(config=None, file=None, folder=None, out=None, verbose=True):
    """Build the dashboard. Safe to call from a notebook.

        import requests_dashboard as rq
        result = rq.run(config=PROJECT / 'Configuration' / 'requests_config.yaml')

    Relative paths in the config resolve against the config file's own folder.
    Raises FileNotFoundError rather than SystemExit, so a bad path does not try
    to shut the kernel down.
    """
    def say(*a):
        if verbose:
            print(*a)

    cfg = load_config(Path(config) if config else None)
    base = Path(config).expanduser().resolve().parent if config else Path.cwd()

    def resolve(p):
        p = Path(p).expanduser()
        return p if p.is_absolute() else base / p

    # 1 · locate the workbook
    chosen = file or cfg["paths"].get("file")
    if chosen:
        src = resolve(chosen)
        if not src.is_file():
            raise FileNotFoundError(f"Workbook not found:\n  {src}")
    else:
        fold = resolve(folder or cfg["paths"]["source_folder"])
        if not fold.is_dir():
            raise FileNotFoundError(f"Source folder not found:\n  {fold}")
        cands = [p for p in fold.glob(cfg["source"]["glob"]) if not p.name.startswith("~$")]
        if not cands:
            raise FileNotFoundError(f"No workbooks in {fold}")
        src = max(cands, key=lambda p: p.stat().st_mtime)
        if len(cands) > 1:
            say(f"  {len(cands)} workbooks present, using the most recently modified")

    say(f"Reading {src.name}")

    # 2 · parse
    orders, warnings = parse_workbook(src, cfg)
    n_lines = sum(len(o["lines"]) for o in orders)
    by_site = {}
    for o in orders:
        by_site[o["site"]] = by_site.get(o["site"], 0) + 1
    say(f"  {len(orders)} order dates · {n_lines} lines · "
        + " · ".join(f"{k} {v}" for k, v in sorted(by_site.items())))
    say(f"  {sum(1 for o in orders if o['hasEta'])} of {len(orders)} have a delivery ETA")
    for w in warnings:
        say(f"  ! {w}")

    # 3 · write
    html = build(orders, warnings, src, cfg)
    out_path = resolve(out or cfg["paths"]["output_html"])
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(html, encoding="utf-8")
    say(f"Wrote {out_path}")

    archived = None
    arch = cfg["paths"].get("archive_folder")
    if arch:
        ap_dir = resolve(arch)
        ap_dir.mkdir(parents=True, exist_ok=True)
        archived = ap_dir / f"Request_ETA_Dashboard_{date.today():%Y-%m-%d}.html"
        archived.write_text(html, encoding="utf-8")
        say(f"Archived {archived}")

    return {"orders": orders, "warnings": warnings, "source": src,
            "output": out_path, "archived": archived, "config": cfg}


def main(argv=None):
    ap = argparse.ArgumentParser(description="HK1/BP1 corrugate request + ETA dashboard.")
    ap.add_argument("--config", help="Path to requests_config.yaml")
    ap.add_argument("--file", help="Override the workbook path")
    ap.add_argument("--folder", help="Override the source folder")
    ap.add_argument("--out", help="Override the output HTML path")
    # parse_known_args so a notebook kernel's own -f argument is ignored
    args, _ = ap.parse_known_args(argv)
    try:
        run(config=args.config, file=args.file, folder=args.folder, out=args.out)
    except FileNotFoundError as e:
        raise SystemExit(str(e))


if __name__ == "__main__":
    main()
