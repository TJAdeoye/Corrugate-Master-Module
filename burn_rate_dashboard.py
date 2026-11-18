"""
burn_rate_dashboard.py - HK1 / BP1 corrugate burn rate, 21-day rolling.

Burn is consumption, not the raw drop in count:

    consumed(prev -> cur) = count(prev) - count(cur) + received(prev < d <= cur)
    burn per day          = sum(consumed) / sum(elapsed days)   over the window

Receipts come from the delivery file's *Actual Delivery Date* and *Delivered
QTY* - never the ETA, so a promised-but-not-arrived load is not counted as
stock on hand.

    python burn_rate_dashboard.py --config "Configuration\\burn_config.yaml"

Notebook:
    import burn_rate_dashboard as burn
    result = burn.run(config=PROJECT / 'Configuration' / 'burn_config.yaml')
"""

from __future__ import annotations
import argparse
import json
import re
from datetime import datetime, date, timedelta
from pathlib import Path

import openpyxl

try:
    import yaml
except ImportError:
    yaml = None

DEFAULT_CFG = {
    "paths": {"tracker_folder": ".", "tracker_file": None,
              "delivery_folder": ".", "delivery_file": None,
              "output_html": "Burn_Rate_Dashboard.html", "archive_folder": None},
    "window": {"days": 21, "as_of": None},
    "thresholds": {"critical_days": 2, "watch_days": 5, "min_intervals": 2},
    "display": {"title": "Corrugate Burn Rate", "title_emphasis": "Burn",
                "eyebrow": "HK1 & BP1 · 21-day rolling", "show_parse_warnings": False,
                "default_method": "drops"},
    "sites": {},
    "archived_skus": [],
}

SITE_PREFIX = {
    "HK": "HK1", "HK1": "HK1", "HODGKINS": "HK1", "HODGKINS1": "HK1",
    "BP": "BP1", "BP1": "BP1", "BEDFORD": "BP1", "BEDFORDPARK": "BP1", "BEDFORDPARK1": "BP1",
}
SITE_NAMES = {"HK1": "Hodgkins 1", "BP1": "Bedford Park 1"}
DATE_RE = re.compile(r"(\d{1,2})\s*[-/]\s*(\d{1,2})\s*[-/]\s*(\d{2,4})")


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


def win(p) -> str:
    return str(p).replace("\\", "/")


def norm_sku(s) -> str:
    """Normalise a SKU label so tracker and delivery spellings line up."""
    return re.sub(r"[^a-z0-9]", "", str(s or "").lower())


def site_from_text(text: str, extra: dict = None):
    """Read a site from a SharePoint-style '[\"Hodgkins 1\"]' value or a tab prefix."""
    table = {**SITE_PREFIX, **(extra or {})}
    m = DATE_RE.search(text)
    head = text[: m.start()] if m else text
    key = re.sub(r"[^A-Z0-9]", "", head.upper())
    if key in table:
        return table[key], False
    # substring fallback for the tracker's JSON-ish Location column
    up = re.sub(r"[^A-Z0-9]", "", text.upper())
    for k, v in table.items():
        if k in up:
            return v, False
    return (key or "?"), True


def parse_sheet_date(name: str):
    m = DATE_RE.search(name)
    if not m:
        return None
    mo, dy, yr = (int(g) for g in m.groups())
    yr += 2000 if yr < 100 else 0
    try:
        return date(yr, mo, dy)
    except ValueError:
        return None


# --------------------------------------------------------------------------
# Tracker: one row per count. Columns A-D are metadata, E onward are SKUs.
# --------------------------------------------------------------------------
def parse_tracker(path: Path, cfg: dict):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    extra = {re.sub(r"[^A-Z0-9]", "", str(k).upper()): v for k, v in (cfg.get("sites") or {}).items()}

    sku_cols = {}
    for c in range(5, ws.max_column + 1):
        label = ws.cell(1, c).value
        if label and str(label).strip():
            sku_cols[c] = str(label).strip()

    archived = {norm_sku(s) for s in (cfg.get("archived_skus") or [])}
    counts, warnings, unknown_sites = {}, [], set()

    for r in range(2, ws.max_row + 1):
        d = ws.cell(r, 2).value
        if not isinstance(d, datetime):
            continue
        site, guessed = site_from_text(str(ws.cell(r, 3).value or ""), extra)
        if guessed:
            unknown_sites.add(site)
            continue

        day = d.date()
        # Later rows for the same site+day supersede earlier ones, so the
        # last count recorded that day is the one used as end-of-day on hand.
        key = (site, day)
        rec = counts.setdefault(key, {"site": site, "date": day, "vals": {},
                                      "stamps": [], "rows": 0})
        rec["rows"] += 1
        rec["stamps"].append(d)
        for c, label in sku_cols.items():
            if norm_sku(label) in archived:
                continue
            v = ws.cell(r, c).value
            if isinstance(v, (int, float)):
                rec["vals"][label] = float(v)

    if unknown_sites:
        warnings.append(
            "Tracker rows with an unrecognised Location were skipped: "
            + ", ".join(sorted(unknown_sites)))

    multi = [k for k, v in counts.items() if v["rows"] > 1]
    return counts, sku_cols, warnings, multi


# --------------------------------------------------------------------------
# Deliveries: one tab per site+date. Only *actual* receipts count.
# --------------------------------------------------------------------------
def parse_deliveries(path: Path | None, cfg: dict):
    receipts, warnings = {}, []      # (site, date) -> {sku: qty}
    pending, mismatches = [], []
    if path is None or not path.is_file():
        return receipts, pending, mismatches, ["No delivery file found, receipts treated as zero."]

    wb = openpyxl.load_workbook(path, data_only=True)
    extra = {re.sub(r"[^A-Z0-9]", "", str(k).upper()): v for k, v in (cfg.get("sites") or {}).items()}

    for name in wb.sheetnames:
        ws = wb[name]
        site, guessed = site_from_text(name, extra)
        tab_date = parse_sheet_date(name)
        if guessed:
            warnings.append(f"Delivery tab “{name}” - could not tell which DC from the name; skipped.")
            continue

        hdr = {}
        for c in range(1, ws.max_column + 1):
            h = re.sub(r"[^a-z]", "", str(ws.cell(1, c).value or "").lower())
            hdr[h] = c
        c_sku = hdr.get("sku")
        c_qty = hdr.get("deliveredqty")
        c_act = hdr.get("actualdeliverydate")
        c_eta = hdr.get("eta")
        if not (c_sku and c_qty):
            warnings.append(f"Delivery tab “{name}” has no SKU / Delivered QTY columns; skipped.")
            continue

        for r in range(2, ws.max_row + 1):
            sku = ws.cell(r, c_sku).value
            if not sku or not str(sku).strip():
                continue
            qty = ws.cell(r, c_qty).value
            act = ws.cell(r, c_act).value if c_act else None

            if not isinstance(qty, (int, float)) or qty == 0:
                eta = ws.cell(r, c_eta).value if c_eta else None
                pending.append({
                    "site": site, "sku": str(sku).strip(),
                    "eta": (eta.date().isoformat() if isinstance(eta, datetime) else None),
                    "tab": name.strip(),
                })
                continue

            eta_val = ws.cell(r, c_eta).value if c_eta else None
            when = act.date() if isinstance(act, datetime) else tab_date

            # delivered on a different day than promised
            if isinstance(act, datetime) and isinstance(eta_val, datetime):
                a_d, e_d = act.date(), eta_val.date()
                if a_d != e_d:
                    mismatches.append({
                        "site": site, "sku": str(sku).strip(),
                        "eta": e_d.isoformat(), "actual": a_d.isoformat(),
                        "variance": (a_d - e_d).days, "qty": float(qty),
                        "tab": name.strip(),
                    })
            if when is None:
                warnings.append(
                    f"Delivery of {sku} on tab “{name}” has a quantity but no delivery date; not counted.")
                continue

            slot = receipts.setdefault((site, when), {})
            slot[str(sku).strip()] = slot.get(str(sku).strip(), 0) + float(qty)

    mismatches.sort(key=lambda m: (-abs(m["variance"]), m["sku"]))
    return receipts, pending, mismatches, warnings


# --------------------------------------------------------------------------
# Burn computation
# --------------------------------------------------------------------------
def compute(counts, receipts, sku_cols, cfg, as_of: date):
    window = int(cfg["window"]["days"])
    th = cfg["thresholds"]
    start = as_of - timedelta(days=window)

    # receipts keyed for quick interval lookup
    rec_by_site = {}
    for (site, d), skus in receipts.items():
        rec_by_site.setdefault(site, {})[d] = skus

    results = {}
    for site in sorted({c["site"] for c in counts.values()}):
        days = sorted(d for (s, d) in counts if s == site and start <= d <= as_of)
        if not days:
            continue

        latest_day = days[-1]
        latest = counts[(site, latest_day)]["vals"]

        per_sku = {}
        for prev_d, cur_d in zip(days, days[1:]):
            elapsed = (cur_d - prev_d).days or 1
            a = counts[(site, prev_d)]["vals"]
            b = counts[(site, cur_d)]["vals"]
            site_rec = rec_by_site.get(site, {})

            for sku in sku_cols.values():
                if sku not in a or sku not in b:
                    continue
                got = 0.0
                for rd, skus in site_rec.items():
                    if prev_d < rd <= cur_d:
                        for rsku, q in skus.items():
                            if norm_sku(rsku) == norm_sku(sku):
                                got += q
                consumed = a[sku] - b[sku] + got
                drop = max(a[sku] - b[sku], 0)    # legacy method: ignore restocks
                s = per_sku.setdefault(sku, {"consumed": 0.0, "drops": 0.0, "days": 0,
                                             "intervals": 0, "received": 0.0,
                                             "negatives": 0, "covered": 0})
                s["consumed"] += consumed
                s["drops"] += drop
                s["days"] += elapsed
                s["intervals"] += 1
                s["received"] += got
                if got:
                    s["covered"] += 1
                if consumed < 0:
                    s["negatives"] += 1

        rows = []
        for sku in sku_cols.values():
            s = per_sku.get(sku)
            on_hand = latest.get(sku)
            if s is None and on_hand is None:
                continue
            net = (s["consumed"] / s["days"]) if (s and s["days"]) else None
            drops = (s["drops"] / s["days"]) if (s and s["days"]) else None

            def cover(rate):
                if rate and rate > 0 and on_hand is not None:
                    return round(on_hand / rate, 1)
                return None

            rows.append({
                "sku": sku,
                "onHand": on_hand,
                "burnNet": round(net, 2) if net is not None else None,
                "burnDrops": round(drops, 2) if drops is not None else None,
                "coverNet": cover(net),
                "coverDrops": cover(drops),
                "received": round(s["received"], 1) if s else 0.0,
                "intervals": s["intervals"] if s else 0,
                "negatives": s["negatives"] if s else 0,
                "thin": bool(s and s["intervals"] < int(th.get("min_intervals", 2))),
            })

        rows.sort(key=lambda r: (r["burnDrops"] is None, -(r["burnDrops"] or 0)))
        results[site] = {
            "site": site,
            "siteName": SITE_NAMES.get(site, site),
            "countDays": len(days),
            "firstDay": days[0].isoformat(),
            "lastDay": latest_day.isoformat(),
            "rows": rows,
        }
    return results


HTML = """<!DOCTYPE html>
<html lang="en"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Corrugate Burn Rate</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;700&family=Fraunces:ital,wght@0,400;0,600;0,800;1,400&display=swap" rel="stylesheet">
<style>
:root{--bg:#0f1114;--bg-panel:#16191e;--bg-elev:#1d2128;--border:#272c35;--border-bright:#3a414e;
--text:#e8e9ec;--text-dim:#8a8f99;--text-faint:#5a5f6a;--amber:#f4a52a;--red:#ef4444;--green:#10b981;--blue:#60a5fa;}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'JetBrains Mono',Consolas,monospace;background:var(--bg);color:var(--text);font-size:13px;line-height:1.5;
background-image:radial-gradient(circle at 15% 0%,rgba(244,165,42,.035),transparent 42%),radial-gradient(circle at 85% 100%,rgba(96,165,250,.03),transparent 42%);}
.container{max-width:1400px;margin:0 auto;padding:28px 32px 90px}
header{border-bottom:1px solid var(--border);padding-bottom:20px;margin-bottom:22px;display:grid;grid-template-columns:1fr auto;gap:24px;align-items:end}
.eyebrow{font-size:10px;letter-spacing:.2em;color:var(--text-faint);text-transform:uppercase;margin-bottom:8px;display:flex;align-items:center;gap:10px}
.eyebrow::before{content:'';width:7px;height:7px;border-radius:50%;background:var(--green);box-shadow:0 0 8px var(--green)}
h1{font-family:'Fraunces',serif;font-weight:800;font-size:36px;letter-spacing:-.02em;line-height:1}
h1 em{font-style:italic;font-weight:400;color:var(--amber)}
.header-meta{text-align:right;font-size:11px;color:var(--text-dim);line-height:1.9}
.header-meta .label{color:var(--text-faint);text-transform:uppercase;letter-spacing:.1em}
.header-meta .value{color:var(--text);font-weight:500}
.banner{display:flex;gap:14px;align-items:flex-start;border-left:2px solid var(--red);background:rgba(239,68,68,.06);
padding:14px 18px;margin-bottom:20px;font-size:11px;color:var(--text-dim);line-height:1.8}
.banner b{color:var(--red);text-transform:uppercase;letter-spacing:.12em;font-size:10px;white-space:nowrap}
.banner.amber{border-color:var(--amber);background:rgba(244,165,42,.05)}
.banner.amber b{color:var(--amber)}
.controls{display:flex;gap:26px;align-items:center;flex-wrap:wrap;margin-bottom:22px;
padding:14px 18px;background:var(--bg-panel);border:1px solid var(--border)}
.cgroup{display:flex;gap:4px;align-items:center;flex-wrap:wrap}
.cgroup .lbl{font-size:10px;color:var(--text-faint);text-transform:uppercase;letter-spacing:.12em;margin-right:8px}
.filter-btn{background:transparent;border:1px solid var(--border);color:var(--text-dim);font-family:inherit;font-size:10px;
letter-spacing:.1em;text-transform:uppercase;padding:6px 12px;cursor:pointer;transition:all .15s}
.filter-btn:hover{color:var(--text);border-color:var(--border-bright)}
.filter-btn.active{background:var(--text);color:var(--bg);border-color:var(--text);font-weight:700}
.kpi-row{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:26px}
.kpi{background:var(--bg-panel);border:1px solid var(--border);padding:18px 20px;position:relative;overflow:hidden}
.kpi::before{content:'';position:absolute;top:0;left:0;width:3px;height:100%;background:var(--border-bright)}
.kpi.alert::before{background:var(--red)}.kpi.warn::before{background:var(--amber)}
.kpi-label{font-size:10px;color:var(--text-faint);text-transform:uppercase;letter-spacing:.12em;margin-bottom:8px}
.kpi-value{font-family:'Fraunces',serif;font-size:32px;font-weight:600;line-height:1}
.kpi-value.alert{color:var(--red)}.kpi-value.warn{color:var(--amber)}
.kpi-value small{font-size:16px;color:var(--text-faint)}
.kpi-sub{font-size:11px;color:var(--text-dim);margin-top:6px}
.site{margin-bottom:40px}
.site-head{display:grid;grid-template-columns:auto 1fr auto;gap:20px;align-items:center;margin-bottom:16px;
padding-bottom:14px;border-bottom:2px solid var(--border-bright)}
.site-name{font-family:'Fraunces',serif;font-size:26px;font-weight:600}
.site-tag{font-size:10px;border:1px solid var(--border-bright);color:var(--text-faint);padding:4px 9px;letter-spacing:.1em}
.site-tag.hk1{color:var(--amber);border-color:var(--amber)}
.site-tag.bp1{color:var(--blue);border-color:var(--blue)}
.pill.hk1{color:var(--amber);border-color:var(--amber)}
.pill.bp1{color:var(--blue);border-color:var(--blue)}
.site-win{font-size:11px;color:var(--text-dim);text-align:right}
.site-win b{color:var(--text)}
.panel{background:var(--bg-panel);border:1px solid var(--border)}
table{width:100%;border-collapse:collapse;font-size:12px}
thead tr{background:var(--bg-elev)}
th{text-align:left;padding:10px 13px;font-size:10px;font-weight:500;color:var(--text-faint);text-transform:uppercase;
letter-spacing:.1em;border-bottom:1px solid var(--border);cursor:pointer;user-select:none;white-space:nowrap}
th:hover{color:var(--text)}
th.num,td.num{text-align:right}
th .arrow{opacity:0;margin-left:4px;font-size:9px}
th.sorted .arrow{opacity:1;color:var(--amber)}
td{padding:9px 13px;border-bottom:1px solid var(--border)}
tbody tr:hover{background:rgba(255,255,255,.015)}
tbody tr:last-child td{border-bottom:none}
td.sku{color:var(--text);font-weight:500}
td.crit{color:var(--red);font-weight:500}td.warn{color:var(--amber);font-weight:500}
td.good{color:var(--green)}td.faint{color:var(--text-faint)}
td.recv{color:var(--blue)}
.pill{display:inline-block;padding:2px 8px;font-size:10px;font-weight:500;letter-spacing:.08em;text-transform:uppercase;border:1px solid}
.pill.crit{color:var(--red);border-color:var(--red)}
.pill.warn{color:var(--amber);border-color:var(--amber)}
.pill.good{color:var(--green);border-color:var(--green)}
.pill.none{color:var(--text-faint);border-color:var(--border)}
.flag{color:var(--red);font-size:11px;cursor:help}
.warnlist{margin-top:26px;padding:16px 20px;border-left:2px solid var(--amber);background:rgba(244,165,42,.05);
font-size:11px;color:var(--text-dim);line-height:1.9}
.warnlist b{color:var(--amber);text-transform:uppercase;letter-spacing:.12em;font-size:10px}
.warnlist ul{margin:6px 0 0 16px}
@media(max-width:1000px){.kpi-row{grid-template-columns:repeat(2,1fr)}}
@media(max-width:640px){.container{padding:18px 16px 60px}header{grid-template-columns:1fr}.header-meta{text-align:left}
h1{font-size:27px}.kpi-row{grid-template-columns:1fr}.site-head{grid-template-columns:1fr}th,td{padding:8px 9px}}
</style></head><body><div class="container">

<header>
  <div>
    <div class="eyebrow">__EYEBROW__</div>
    <h1>__TITLE_HTML__</h1>
  </div>
  <div class="header-meta">
    <div><span class="label">Window</span> &nbsp; <span class="value">__WINDOW__ days to __ASOF__</span></div>
    <div><span class="label">Receipts logged</span> &nbsp; <span class="value">__RECDAYS__</span></div>
    <div><span class="label">Generated</span> &nbsp; <span class="value">__GENERATED__</span></div>
  </div>
</header>

__BANNER__

<div class="controls">
  <div class="cgroup">
    <span class="lbl">Method</span>
    <button class="filter-btn __NETACT__" data-m="net">Receipt-adjusted</button>
    <button class="filter-btn __DROPACT__" data-m="drops">Count drops only</button>
  </div>
  <div class="cgroup">
    <span class="lbl">Show</span>
    <button class="filter-btn active" data-f="all">All</button>
    <button class="filter-btn" data-f="active">Burning</button>
    <button class="filter-btn" data-f="crit">Critical</button>
    <button class="filter-btn" data-f="flagged">Flagged</button>
  </div>
</div>

<div class="kpi-row" id="kpis"></div>
<div id="sites"></div>
__WARNINGS__

</div>
<script>
const SITES = __SITES__;
const CRIT = __CRIT__, WATCH = __WATCH__, ETA_MISS = __ETAMISS__;
let METHOD = '__DEFMETHOD__', FILTER = 'all';

const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const fmt = n => (n === null || n === undefined) ? '-' : Number(n).toLocaleString('en-US');
const f2 = n => (n === null || n === undefined) ? '-' : Number(n).toFixed(2);
const f1 = n => (n === null || n === undefined) ? '-' : Number(n).toFixed(1);
const burnOf = r => METHOD === 'net' ? r.burnNet : r.burnDrops;
const coverOf = r => METHOD === 'net' ? r.coverNet : r.coverDrops;

function tier(r){
  const b = burnOf(r), c = coverOf(r);
  if (b === null || b <= 0) return 'none';
  if (r.onHand === 0) return 'crit';
  if (c === null) return 'none';
  if (c <= CRIT) return 'crit';
  if (c <= WATCH) return 'warn';
  return 'good';
}
const label = t => ({crit:'Critical', warn:'Watch', good:'Healthy', none:'No burn'})[t];

function visible(rows){
  if (FILTER === 'active')  return rows.filter(r => (burnOf(r) || 0) > 0);
  if (FILTER === 'crit')    return rows.filter(r => tier(r) === 'crit');
  if (FILTER === 'flagged') return rows.filter(r => r.negatives > 0 || r.thin);
  return rows;
}

function renderKpis(){
  const all = SITES.flatMap(s => s.rows);
  const crit = all.filter(r => tier(r) === 'crit').length;
  const warn = all.filter(r => tier(r) === 'warn').length;
  const flagged = all.filter(r => r.negatives > 0).length;
  const recv = all.reduce((s,r) => s + (r.received || 0), 0);
  document.getElementById('kpis').innerHTML = `
    <div class="kpi ${crit ? 'alert' : ''}">
      <div class="kpi-label">Critical SKUs</div>
      <div class="kpi-value ${crit ? 'alert' : ''}">${crit}</div>
      <div class="kpi-sub">zero on hand or &le; ${CRIT} days cover</div>
    </div>
    <div class="kpi ${warn > 2 ? 'warn' : ''}">
      <div class="kpi-label">Watch SKUs</div>
      <div class="kpi-value ${warn > 2 ? 'warn' : ''}">${warn}</div>
      <div class="kpi-sub">${CRIT}-${WATCH} days of cover</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Units received</div>
      <div class="kpi-value">${fmt(recv)}</div>
      <div class="kpi-sub">logged in the window</div>
    </div>
    <div class="kpi ${ETA_MISS ? 'alert' : ''}">
      <div class="kpi-label">ETA misses</div>
      <div class="kpi-value ${ETA_MISS ? 'alert' : ''}">${ETA_MISS}</div>
      <div class="kpi-sub">delivered off the promised date</div>
    </div>
    <div class="kpi ${flagged ? 'warn' : ''}" style="display:none">
      <div class="kpi-label">Flagged lines</div>
      <div class="kpi-value ${flagged ? 'warn' : ''}">${flagged}</div>
      <div class="kpi-sub">count rose with no receipt logged</div>
    </div>`;
}

function renderSites(){
  document.getElementById('sites').innerHTML = SITES.map(s => {
    const rows = visible(s.rows);
    return `<section class="site">
      <div class="site-head">
        <span class="site-tag ${s.site.toLowerCase()}">${esc(s.site)}</span>
        <div class="site-name">${esc(s.siteName)}</div>
        <div class="site-win"><b>${s.countDays}</b> count days &nbsp;·&nbsp; ${esc(s.firstDay)} → <b>${esc(s.lastDay)}</b></div>
      </div>
      <div class="panel"><table>
        <thead><tr>
          <th data-s="sku">SKU <span class="arrow">▼</span></th>
          <th class="num" data-s="onHand">On hand <span class="arrow">▼</span></th>
          <th class="num sorted" data-s="burn">Burn / day <span class="arrow">▼</span></th>
          <th class="num" data-s="cover">Days cover <span class="arrow">▼</span></th>
          <th class="num" data-s="received">Received <span class="arrow">▼</span></th>
          <th class="num" data-s="intervals">Intervals <span class="arrow">▼</span></th>
          <th data-s="status">Status <span class="arrow">▼</span></th>
        </tr></thead>
        <tbody>${rows.length ? rows.map(r => {
          const t = tier(r);
          return `<tr>
            <td class="sku">${esc(r.sku)}${r.negatives ? ` <span class="flag" title="${r.negatives} of ${r.intervals} intervals show a stock increase with no receipt logged">&#9650;</span>` : ''}</td>
            <td class="num ${r.onHand === 0 ? 'crit' : ''}">${fmt(r.onHand)}</td>
            <td class="num">${f2(burnOf(r))}</td>
            <td class="num ${t}">${f1(coverOf(r))}</td>
            <td class="num ${r.received ? 'recv' : 'faint'}">${r.received ? fmt(r.received) : '-'}</td>
            <td class="num faint">${r.intervals}</td>
            <td><span class="pill ${t}">${label(t)}</span></td>
          </tr>`;
        }).join('') : '<tr><td colspan="7" style="padding:26px;text-align:center;color:var(--text-faint)">Nothing matches this filter.</td></tr>'}</tbody>
      </table></div>
    </section>`;
  }).join('');

  document.querySelectorAll('#sites th').forEach(th => th.addEventListener('click', () => {
    const k = th.dataset.s;
    SITES.forEach(s => s.rows.sort((a,b) => {
      let av, bv;
      if (k === 'burn'){ av = burnOf(a); bv = burnOf(b); }
      else if (k === 'cover'){ av = coverOf(a); bv = coverOf(b); }
      else if (k === 'status'){ av = tier(a); bv = tier(b); }
      else { av = a[k]; bv = b[k]; }
      if (av === null || av === undefined) av = -Infinity;
      if (bv === null || bv === undefined) bv = -Infinity;
      if (av < bv) return 1; if (av > bv) return -1; return 0;
    }));
    renderSites();
  }));
}

function renderAll(){ renderKpis(); renderSites(); }

document.querySelectorAll('[data-m]').forEach(b => b.addEventListener('click', () => {
  document.querySelectorAll('[data-m]').forEach(x => x.classList.remove('active'));
  b.classList.add('active'); METHOD = b.dataset.m; renderAll();
}));
document.querySelectorAll('[data-f]').forEach(b => b.addEventListener('click', () => {
  document.querySelectorAll('[data-f]').forEach(x => x.classList.remove('active'));
  b.classList.add('active'); FILTER = b.dataset.f; renderAll();
}));
renderAll();
</script></body></html>
"""


def build(results, warnings, cfg, as_of, rec_days, total_days, pending, mismatches=None):
    mismatches = mismatches or []
    disp, th = cfg["display"], cfg["thresholds"]
    sites = [results[k] for k in sorted(results)]

    method = str(disp.get("default_method", "drops")).lower()
    if method not in ("net", "drops"):
        method = "drops"

    banner = ""

    if pending:
        names = ", ".join(f"{p['sku']} ({p['site']})" for p in pending[:6])
        more = f" and {len(pending) - 6} more" if len(pending) > 6 else ""
        banner += (
            f'<div class="banner amber"><b>{len(pending)} awaiting delivery</b>'
            f'<span>On the delivery file with an ETA but no delivered quantity: {names}{more}. '
            f'Not counted as received.</span></div>')

    if mismatches:
        late = sum(1 for m in mismatches if m["variance"] > 0)
        rows = "".join(
            f'<tr><td><span class="pill {m["site"].lower()}">{m["site"]}</span></td>'
            f'<td class="sku">{m["sku"]}</td>'
            f'<td class="num">{m["qty"]:g}</td>'
            f'<td class="faint">{m["eta"]}</td>'
            f'<td>{m["actual"]}</td>'
            f'<td class="num {"crit" if m["variance"] > 0 else "good"}">'
            f'{"+" if m["variance"] > 0 else ""}{m["variance"]}d</td></tr>'
            for m in mismatches)
        banner += (
            '<section class="site"><div class="site-head">'
            '<span class="site-tag" style="color:var(--red);border-color:var(--red)">ETA MISS</span>'
            f'<div class="site-name">Delivered off the promised date</div>'
            f'<div class="site-win"><b>{len(mismatches)}</b> deliveries '
            f'&nbsp;-&nbsp; {late} late, {len(mismatches) - late} early</div></div>'
            '<div class="panel"><table><thead><tr>'
            '<th>DC</th><th>SKU</th><th class="num">Qty</th>'
            '<th>ETA</th><th>Actual</th><th class="num">Variance</th>'
            '</tr></thead><tbody>' + rows + '</tbody></table></div></section>')

    warn_html = ""
    if warnings and disp.get("show_parse_warnings", True):
        lis = "".join(f"<li>{w}</li>" for w in warnings)
        warn_html = f'<div class="warnlist"><b>Notes - {len(warnings)}</b><ul>{lis}</ul></div>'

    title = disp["title"]
    emph = disp.get("title_emphasis")
    title_html = title.replace(emph, f"<em>{emph}</em>", 1) if emph and emph in title else title

    return (HTML
            .replace("__SITES__", json.dumps(sites))
            .replace("__CRIT__", str(th["critical_days"]))
            .replace("__ETAMISS__", str(len(mismatches or [])))
            .replace("__WATCH__", str(th["watch_days"]))
            .replace("__TITLE_HTML__", title_html)
            .replace("__EYEBROW__", str(disp.get("eyebrow", "")))
            .replace("__WINDOW__", str(cfg["window"]["days"]))
            .replace("__ASOF__", as_of.strftime("%b %d, %Y").replace(" 0", " "))
            .replace("__RECDAYS__", f"{rec_days} of {total_days} days")
            .replace("__GENERATED__", datetime.now().strftime("%b %d, %Y %I:%M %p").replace(" 0", " "))
            .replace("__DEFMETHOD__", method)
            .replace("__NETACT__", "active" if method == "net" else "")
            .replace("__DROPACT__", "active" if method == "drops" else "")
            .replace("__BANNER__", banner)
            .replace("__WARNINGS__", warn_html))


def _pick(folder: Path, explicit, label, base):
    if explicit:
        p = Path(win(explicit)).expanduser()
        p = p if p.is_absolute() else base / p
        if not p.is_file():
            raise FileNotFoundError(f"{label} not found:\n  {p}")
        return p
    f = Path(win(folder)).expanduser()
    f = f if f.is_absolute() else base / f
    if not f.is_dir():
        raise FileNotFoundError(f"{label} folder not found:\n  {f}")
    cands = [p for p in f.glob("*.xls*") if not p.name.startswith("~$")]
    if not cands:
        raise FileNotFoundError(f"No workbooks in {f}")
    return max(cands, key=lambda p: p.stat().st_mtime)


def run(config=None, out=None, as_of=None, verbose=True):
    """Build the burn-rate dashboard. Safe to call from a notebook."""
    def say(*a):
        if verbose:
            print(*a)

    cfg = load_config(Path(config) if config else None)
    base = Path(config).expanduser().resolve().parent if config else Path.cwd()

    tracker = _pick(cfg["paths"]["tracker_folder"], cfg["paths"].get("tracker_file"), "Tracker", base)
    delivery = None
    try:
        delivery = _pick(cfg["paths"]["delivery_folder"], cfg["paths"].get("delivery_file"), "Delivery file", base)
    except FileNotFoundError as e:
        say(f"  ! {e}")

    say(f"Tracker  {tracker.name}")
    say(f"Delivery {delivery.name if delivery else '(none)'}")

    counts, sku_cols, w1, multi = parse_tracker(tracker, cfg)
    receipts, pending, mismatches, w2 = parse_deliveries(delivery, cfg)
    warnings = list(w1) + list(w2)

    if as_of is None:
        as_of = cfg["window"].get("as_of")
        as_of = (datetime.strptime(as_of, "%Y-%m-%d").date()
                 if isinstance(as_of, str) else (as_of or date.today()))

    window = int(cfg["window"]["days"])
    start = as_of - timedelta(days=window)
    recent_multi = [k for k in multi if start <= k[1] <= as_of]
    if recent_multi:
        warnings.append(
            f"{len(recent_multi)} site/day(s) in the window have more than one count row "
            f"(different shifts, or a repeated entry). The last row recorded for that day is used "
            f"as end-of-day on hand: "
            + ", ".join(f"{s} {d}" for s, d in sorted(recent_multi, key=lambda x: x[1])[:6])
            + ("…" if len(recent_multi) > 6 else ""))

    results = compute(counts, receipts, sku_cols, cfg, as_of)

    rec_days = len({d for (s, d) in receipts if start <= d <= as_of})
    total_days = window + 1
    for site, r in results.items():
        flagged = sum(1 for x in r["rows"] if x["negatives"])
        say(f"  {site}: {r['countDays']} count days · {len(r['rows'])} SKUs · {flagged} flagged")
    say(f"  receipts logged on {rec_days} of {total_days} days in the window")
    if mismatches:
        say(f"  {len(mismatches)} delivery(s) did not land on the promised ETA")
        for m in mismatches[:5]:
            sign = "+" if m["variance"] > 0 else ""
            say(f"      {m['site']} {m['sku']}: ETA {m['eta']} -> {m['actual']} ({sign}{m['variance']}d)")
    for w in warnings:
        say(f"  ! {w}")

    html_out = build(results, warnings, cfg, as_of, rec_days, total_days, pending, mismatches)
    out_path = Path(win(out or cfg["paths"]["output_html"])).expanduser()
    if not out_path.is_absolute():
        out_path = base / out_path
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(html_out, encoding="utf-8")
    say(f"Wrote {out_path}")

    archived = None
    arch = cfg["paths"].get("archive_folder")
    if arch:
        ad = Path(win(arch)).expanduser()
        if not ad.is_absolute():
            ad = base / ad
        ad.mkdir(parents=True, exist_ok=True)
        archived = ad / f"Burn_Rate_{as_of:%Y-%m-%d}.html"
        archived.write_text(html_out, encoding="utf-8")
        say(f"Archived {archived}")

    return {"results": results, "warnings": warnings, "pending": pending,
            "mismatches": mismatches,
            "receipts": receipts, "counts": counts, "asOf": as_of,
            "recDays": rec_days, "output": out_path, "archived": archived, "config": cfg}


def main(argv=None):
    ap = argparse.ArgumentParser(description="HK1/BP1 corrugate burn rate dashboard.")
    ap.add_argument("--config", help="Path to burn_config.yaml")
    ap.add_argument("--out", help="Override the output HTML path")
    ap.add_argument("--as-of", help="Anchor the window to this date (YYYY-MM-DD)")
    args, _ = ap.parse_known_args(argv)
    ao = datetime.strptime(args.as_of, "%Y-%m-%d").date() if args.as_of else None
    try:
        run(config=args.config, out=args.out, as_of=ao)
    except FileNotFoundError as e:
        raise SystemExit(str(e))


if __name__ == "__main__":
    main()
