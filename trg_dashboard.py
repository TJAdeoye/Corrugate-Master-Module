"""
trg_dashboard.py - TRG supplier on-hand dashboard.

Scans a folder for TRG inventory reports, picks the most recent one,
parses it, and writes a self-contained HTML dashboard.

    python trg_dashboard.py --folder "<path to TRG reports>" --out trg_dashboard.html

The TRG workbook is a hand-maintained sheet, so the parser is deliberately
defensive: anything it cannot read confidently is surfaced on the dashboard
as a parse warning rather than guessed at or silently dropped.

Layout it expects (one SKU per 3-row block, starting row 6):
    A: box name           B: TRG order due    C: qty on hand (+ in-process)
    D: last released      E: re-order qty     F: burn rate
    A+1: units per pallet
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
    "paths": {"source_folder": ".", "output_html": "trg_dashboard.html", "archive_folder": None},
    "source": {"glob": "*.xls*", "prefer_filename_date": True, "sheet": None,
               "first_data_row": 6,
               "columns": {"name": 1, "due": 2, "qty": 3, "released": 4, "reorder": 5, "burn": 6}},
    "thresholds": {"critical_days": 2, "watch_days": 5, "stale_after_days": 1},
    "display": {"title": "TRG On-Hand Inventory",
                "subtitle": "Boxes held at TRG, burn rate, and days of cover",
                "title_emphasis": "On-Hand",
                "show_freshness_banner": False, "show_missing_sku_banner": False,
                "show_parse_warnings": True},
    "sku_map": {}, "not_stocked_at_trg": [],
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

# --------------------------------------------------------------------------
# TRG box name -> our canonical nomenclature.
# Keys are normalized (uppercase, non-alphanumerics stripped).
# --------------------------------------------------------------------------
TRG_TO_NOM = {
    "12X12X8PLAINA":        "APLAIN",
    "12X12X8PRINTEDA":      "APRINTED",
    "14X14X8PLAINC":        "CPLAIN",
    "14X14X8PRINTEDC":      "CPRINTED",
    "14X14X10PLEDW":        "EPLAIN",
    "14X14X10PREDW":        "EPRINTED",
    "14X14X14PLBDW":        "BPLAIN",
    "14X14X14PRBDW":        "BPRINTED",
    "16X16X10PLFDW":        "FPLAIN",
    "16X16X10PRFDW":        "FPRINTED",
    "16X16X12PRINTED":      "DPRINT",
    "16X16X16PLGDW":        "GPLAIN",
    "16X16X16PRGDW":        "GPRINTED",
    "14X14X18CAZOUT":       "OPRINTED",
    "PALLETCHIMNEY":        "CHIMNEY",
    "PALLETCHIMENYINSERT":  "INSERTS",   # sic - TRG spells it "Chimeny"
    "PALLETCHIMNEYINSERT":  "INSERTS",
    "PALLETSHEET":          "SHEETS",
    "PALLETSHEETS":         "SHEETS",
    "POWERCLEARWIPERCARTON": "WIPER",
    "161620FAZ":            "JPLAIN",    # 16x16x20 - confirmed as JPLAIN
    "171722":               "HPRINT",
}

NOM_META = {
    "OPRINTED": ("OP",  "Amazon Tall Box - 14x14x18"),
    "JPLAIN":   ("JP",  "Amazon Tall Box - 16x16x20"),
    "DPRINT":   ("DP",  "16x16x12 Printed"),
    "APLAIN":   ("AP",  "12x12x8 Plain"),
    "APRINTED": ("A",   "12x12x8 Printed"),
    "BPLAIN":   ("B",   "14x14x14 Plain"),
    "BPRINTED": ("BP",  "14x14x14 Printed"),
    "CPLAIN":   ("C",   "14x14x8 Plain"),
    "CPRINTED": ("CP",  "14x14x8 Printed"),
    "EPLAIN":   ("E",   "14x14x10 Plain"),
    "EPRINTED": ("EP",  "14x14x10 Printed"),
    "FPLAIN":   ("F",   "16x16x10 Plain"),
    "FPRINTED": ("FP",  "16x16x10 Printed"),
    "GPLAIN":   ("G",   "16x16x16 Plain"),
    "GPRINTED": ("GP",  "16x16x16 Printed"),
    "CHIMNEY":  ("CHM", "Pallet Chimney"),
    "SHEETS":   ("PS",  "Pallet Sheets"),
    "HPRINT":   ("HP",  "17x17x22 Printed"),
    "INSERTS":  ("IN",  "Chimney Inserts"),
    "WIPER":    ("WP",  "Wiper Boxes"),
    "ACAL":     ("ACL", "12x12x8 Callahan"),
    "ECAL":     ("ECL", "14x14x10 Callahan"),
    "FCAL":     ("FCL", "16x16x10 Callahan"),
    "GCAL":     ("GCL", "16x16x16 Callahan"),
    "BCAL":     ("BCL", "14x14x14 Callahan"),
}


def norm_key(s) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(s).upper())


def parse_k(token: str):
    """'10k' -> 10000, '5000' -> 5000, else None."""
    t = token.strip().lower().replace(",", "")
    m = re.fullmatch(r"(\d+(?:\.\d+)?)\s*k", t)
    if m:
        return int(float(m.group(1)) * 1000)
    if re.fullmatch(r"\d+", t):
        return int(t)
    return None


def parse_qty(raw):
    """Parse column C.

    Handles: 2460 · '790' · '13550 - 10k I/P' · '0: 5000 I/P' · '263 - 850 I/P'
    Returns (on_hand, in_process, note) - any of which may be None.
    """
    if raw is None:
        return None, None, None
    if isinstance(raw, (int, float)):
        return int(raw), None, None

    s = str(raw).strip()
    if not s:
        return None, None, None

    in_process = None
    ip = re.search(r"(\d+(?:\.\d+)?\s*k|\d[\d,]*)\s*(?:I\s*/\s*P)", s, re.I)
    if ip:
        in_process = parse_k(ip.group(1))
        s = s[: ip.start()]

    s = re.split(r"[-:]", s)[0]
    on_hand = parse_k(s)
    note = None if on_hand is not None else str(raw).strip()
    return on_hand, in_process, note


def parse_burn(raw):
    """Parse column F.

    '900' -> (900, None) · '250 (-250)' -> (250, '-250')
    '(+6000 BP)' -> (None, '+6000 BP')   <- annotation only, no rate given
    """
    if raw is None:
        return None, None
    if isinstance(raw, (int, float)):
        return float(raw), None

    s = str(raw).strip()
    if not s:
        return None, None

    annot = None
    paren = re.search(r"\(([^)]*)\)", s)
    if paren:
        annot = paren.group(1).strip()
        s = s[: paren.start()] + s[paren.end():]

    m = re.search(r"\d[\d,]*(?:\.\d+)?", s)
    rate = float(m.group(0).replace(",", "")) if m else None
    return rate, annot


def parse_pallet(raw):
    """'600/Unit' -> 600"""
    if raw is None:
        return None
    m = re.search(r"(\d[\d,]*)\s*/\s*unit", str(raw), re.I)
    return int(m.group(1).replace(",", "")) if m else None


def parse_due(raw, report_dt: date):
    """'9/8' -> ISO date, inferring the year from the report date."""
    if raw is None:
        return None
    if isinstance(raw, datetime):
        return raw.date().isoformat()
    s = str(raw).strip()
    m = re.fullmatch(r"(\d{1,2})\s*/\s*(\d{1,2})(?:\s*/\s*(\d{2,4}))?", s)
    if not m:
        return None
    mo, dy = int(m.group(1)), int(m.group(2))
    if m.group(3):
        yr = int(m.group(3))
        yr += 2000 if yr < 100 else 0
    else:
        yr = report_dt.year
        # a due month far ahead of the report month means it rolled over
        if mo < report_dt.month - 6:
            yr += 1
    try:
        return date(yr, mo, dy).isoformat()
    except ValueError:
        return None


DATE_IN_NAME = re.compile(r"(\d{1,2})[-_.](\d{1,2})[-_.](\d{2,4})")


def report_date_from(path: Path):
    """Prefer the date in the filename; fall back to file mtime."""
    m = DATE_IN_NAME.search(path.stem)
    if m:
        mo, dy, yr = (int(g) for g in m.groups())
        yr += 2000 if yr < 100 else 0
        try:
            return date(yr, mo, dy), "filename"
        except ValueError:
            pass
    return date.fromtimestamp(path.stat().st_mtime), "file timestamp"


def find_latest(folder: Path):
    files = [p for p in folder.glob("*.xls*") if not p.name.startswith("~$")]
    if not files:
        raise SystemExit(f"No .xlsx files found in {folder}")
    dated = [(p, *report_date_from(p)) for p in files]
    dated.sort(key=lambda t: (t[1], t[0].stat().st_mtime), reverse=True)
    return dated[0], [d[0].name for d in dated[1:6]]


def parse_report(path: Path, report_dt: date, cfg: dict = None):
    cfg = cfg or DEFAULT_CFG
    src = cfg["source"]
    col = src["columns"]

    # config sku_map is layered over the built-in map
    sku_map = dict(TRG_TO_NOM)
    sku_map.update({norm_key(k): v for k, v in (cfg.get("sku_map") or {}).items()})

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[src["sheet"]] if src.get("sheet") else wb[wb.sheetnames[0]]

    items, warnings = [], []
    row = src["first_data_row"]
    while row <= ws.max_row:
        name = ws.cell(row, col["name"]).value
        if name is None or not str(name).strip():
            row += 1
            continue
        if re.search(r"/\s*unit", str(name), re.I):   # pallet-size line
            row += 1
            continue

        on_hand, in_process, qty_note = parse_qty(ws.cell(row, col["qty"]).value)
        burn, burn_note = parse_burn(ws.cell(row, col["burn"]).value)
        pallet = parse_pallet(ws.cell(row + 1, col["name"]).value)

        released = ws.cell(row, col["released"]).value
        released = released.date().isoformat() if isinstance(released, datetime) else None

        reorder = ws.cell(row, col["reorder"]).value
        reorder = int(reorder) if isinstance(reorder, (int, float)) else None

        raw_name = str(name).strip()
        nom = sku_map.get(norm_key(raw_name))
        tag, disp = NOM_META.get(nom, ("-", raw_name)) if nom else ("-", raw_name)

        if nom is None:
            warnings.append(f"Unmapped TRG line “{raw_name}” (row {row}) - shown, but not matched to a SKU. "
                            f"Add it to sku_map in the config to fix.")
        if qty_note:
            warnings.append(f"{disp}: could not read on-hand from “{qty_note}”.")

        doh = round(on_hand / burn, 1) if (on_hand is not None and burn) else None

        items.append({
            "nom": nom, "tag": tag, "name": disp, "trgName": raw_name,
            "onHand": on_hand, "inProcess": in_process,
            "burn": burn, "burnNote": burn_note,
            "daysCover": doh, "pallet": pallet,
            "due": parse_due(ws.cell(row, col["due"]).value, report_dt),
            "released": released, "reorder": reorder,
        })
        row += 1

    seen = {i["nom"] for i in items}
    expected = set(cfg.get("not_stocked_at_trg") or [])
    missing = [n for n in NOM_META if n not in seen and n not in expected]
    return items, warnings, missing


def find_latest_cfg(folder: Path, cfg: dict):
    pattern = cfg["source"].get("glob", "*.xls*")
    files = [p for p in folder.glob(pattern) if not p.name.startswith("~$")]
    if not files:
        raise SystemExit(f"No files matching {pattern} in {folder}")
    dated = [(p, *report_date_from(p)) for p in files]
    dated.sort(key=lambda t: (t[1], t[0].stat().st_mtime), reverse=True)
    return dated[0], [d[0].name for d in dated[1:6]]


# --------------------------------------------------------------------------
# HTML
# --------------------------------------------------------------------------
HTML = """<!DOCTYPE html>
<html lang="en"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>TRG Supplier Inventory · Daily Report</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;700&family=Fraunces:ital,wght@0,400;0,600;0,800;1,400&display=swap" rel="stylesheet">
<style>
:root{--bg:#0f1114;--bg-panel:#16191e;--bg-elev:#1d2128;--border:#272c35;--border-bright:#3a414e;
--text:#e8e9ec;--text-dim:#8a8f99;--text-faint:#5a5f6a;--amber:#f4a52a;--red:#ef4444;--green:#10b981;--blue:#60a5fa;}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'JetBrains Mono',Consolas,monospace;background:var(--bg);color:var(--text);font-size:13px;line-height:1.5;
background-image:radial-gradient(circle at 15% 0%,rgba(244,165,42,.035),transparent 42%),radial-gradient(circle at 85% 100%,rgba(96,165,250,.03),transparent 42%);}
.container{max-width:1380px;margin:0 auto;padding:28px 32px 80px}
header{border-bottom:1px solid var(--border);padding-bottom:20px;margin-bottom:22px;display:grid;grid-template-columns:1fr auto;gap:24px;align-items:end}
.eyebrow{font-size:10px;letter-spacing:.2em;color:var(--text-faint);text-transform:uppercase;margin-bottom:8px;display:flex;align-items:center;gap:10px}
.eyebrow::before{content:'';width:7px;height:7px;border-radius:50%;background:var(--green);box-shadow:0 0 8px var(--green)}
.eyebrow.stale::before{background:var(--amber);box-shadow:0 0 8px var(--amber)}
h1{font-family:'Fraunces',serif;font-weight:800;font-size:36px;letter-spacing:-.02em;line-height:1}
h1 em{font-style:italic;font-weight:400;color:var(--amber)}
.subtitle{margin-top:8px;color:var(--text-dim);font-size:12px}
.header-meta{text-align:right;font-size:11px;color:var(--text-dim);line-height:1.9}
.header-meta .label{color:var(--text-faint);text-transform:uppercase;letter-spacing:.1em}
.header-meta .value{color:var(--text);font-weight:500}
.banner{display:flex;gap:12px;align-items:flex-start;border-left:2px solid var(--amber);background:rgba(244,165,42,.05);
padding:13px 18px;margin-bottom:20px;font-size:11px;color:var(--text-dim);line-height:1.8}
.banner b{color:var(--amber);text-transform:uppercase;letter-spacing:.12em;font-size:10px;white-space:nowrap}
.banner.blue{border-color:var(--blue);background:rgba(96,165,250,.05)}
.banner.blue b{color:var(--blue)}
.kpi-row{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:26px}
.kpi{background:var(--bg-panel);border:1px solid var(--border);padding:18px 20px;position:relative;overflow:hidden}
.kpi::before{content:'';position:absolute;top:0;left:0;width:3px;height:100%;background:var(--border-bright)}
.kpi.alert::before{background:var(--red)}.kpi.warn::before{background:var(--amber)}
.kpi-label{font-size:10px;color:var(--text-faint);text-transform:uppercase;letter-spacing:.12em;margin-bottom:8px}
.kpi-value{font-family:'Fraunces',serif;font-size:32px;font-weight:600;line-height:1}
.kpi-value.alert{color:var(--red)}.kpi-value.warn{color:var(--amber)}
.kpi-value small{font-size:16px;color:var(--text-faint)}
.kpi-sub{font-size:11px;color:var(--text-dim);margin-top:6px}
.section-head{display:flex;align-items:baseline;gap:14px;margin-bottom:14px}
.section-head h3{font-family:'Fraunces',serif;font-size:18px;font-weight:600}
.section-head .tag{font-size:10px;color:var(--amber);letter-spacing:.15em;text-transform:uppercase;border:1px solid var(--amber);padding:2px 8px}
.section-head .meta{font-size:11px;color:var(--text-faint);margin-left:auto}
.cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(232px,1fr));gap:12px;margin-bottom:30px}
.card{background:var(--bg-panel);border:1px solid var(--border);padding:16px;position:relative}
.card::before{content:'';position:absolute;top:0;left:0;width:3px;height:100%;background:var(--red)}
.card.warn::before{background:var(--amber)}
.card .chip{font-size:11px;font-weight:700;letter-spacing:.08em;color:var(--amber);margin-bottom:4px}
.card .sku{font-size:13px;font-weight:500;margin-bottom:12px;padding-right:8px}
.card .stat{display:flex;justify-content:space-between;align-items:baseline;padding:5px 0;border-top:1px dashed var(--border)}
.card .stat:first-of-type{border-top:none}
.card .stat-label{font-size:10px;color:var(--text-faint);text-transform:uppercase;letter-spacing:.08em}
.card .stat-val{font-size:15px;font-weight:500}
.card .stat-val.crit{color:var(--red)}.card .stat-val.warn{color:var(--amber)}
.card .ip{font-size:10px;color:var(--blue);margin-top:9px;letter-spacing:.05em}
.table-wrap{background:var(--bg-panel);border:1px solid var(--border)}
.table-head{display:flex;justify-content:space-between;align-items:center;padding:14px 20px;border-bottom:1px solid var(--border);gap:12px;flex-wrap:wrap}
.table-head h4{font-family:'Fraunces',serif;font-size:15px;font-weight:600}
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
td.tagcell{color:var(--amber);font-weight:700;letter-spacing:.06em}
td.name{color:var(--text)}
td.crit{color:var(--red);font-weight:500}td.warn{color:var(--amber);font-weight:500}td.good{color:var(--green)}
td.faint{color:var(--text-faint)}
.ip-badge{color:var(--blue);font-size:11px}
.pill{display:inline-block;padding:2px 8px;font-size:10px;font-weight:500;letter-spacing:.08em;text-transform:uppercase;border:1px solid}
.pill.crit{color:var(--red);border-color:var(--red)}
.pill.warn{color:var(--amber);border-color:var(--amber)}
.pill.good{color:var(--green);border-color:var(--green)}
.pill.none{color:var(--text-faint);border-color:var(--border)}
.warnlist{margin-top:26px;padding:16px 20px;border-left:2px solid var(--red);background:rgba(239,68,68,.05);font-size:11px;color:var(--text-dim);line-height:1.9}
.warnlist b{color:var(--red);text-transform:uppercase;letter-spacing:.12em;font-size:10px}
.warnlist ul{margin:6px 0 0 16px}
@media(max-width:1000px){.kpi-row{grid-template-columns:repeat(2,1fr)}}
@media(max-width:640px){.container{padding:18px 16px 60px}header{grid-template-columns:1fr}.header-meta{text-align:left}
h1{font-size:27px}.kpi-row{grid-template-columns:1fr}.cards{grid-template-columns:1fr}th,td{padding:8px 10px}}
</style></head><body><div class="container">

<header>
  <div>
    <div class="eyebrow __STALECLS__">TRG · Supplier floor stock</div>
    <h1>__TITLE_HTML__</h1>
    <div class="subtitle">__SUBTITLE__</div>
  </div>
  <div class="header-meta">
    <div><span class="label">Report date</span> &nbsp; <span class="value">__REPORTDATE__</span></div>
    <div><span class="label">Source file</span> &nbsp; <span class="value">__SRCFILE__</span></div>
    <div><span class="label">Generated</span> &nbsp; <span class="value">__GENERATED__</span></div>
  </div>
</header>

__FRESHNESS__

<div class="kpi-row" id="kpis"></div>

<div class="section-head">
  <h3>Needs attention</h3><span class="tag">Action</span>
  <span class="meta">Zero on hand, or under __WATCH_DAYS__ days of cover</span>
</div>
<div class="cards" id="cards"></div>

<div class="table-wrap">
  <div class="table-head">
    <h4>All TRG lines</h4>
    <div class="filter-btns">
      <button class="filter-btn active" data-f="all">All</button>
      <button class="filter-btn" data-f="crit">Critical</button>
      <button class="filter-btn" data-f="warn">Watch</button>
      <button class="filter-btn" data-f="ip">In process</button>
      <button class="filter-btn" data-f="nodata">No burn rate</button>
    </div>
  </div>
  <table>
    <thead><tr>
      <th data-sort="tag">Tag <span class="arrow">▼</span></th>
      <th data-sort="name">Box <span class="arrow">▼</span></th>
      <th class="num" data-sort="onHand">On hand <span class="arrow">▼</span></th>
      <th class="num" data-sort="inProcess">In process <span class="arrow">▼</span></th>
      <th class="num" data-sort="burn">Burn / day <span class="arrow">▼</span></th>
      <th class="num sorted" data-sort="daysCover">Days cover <span class="arrow">▼</span></th>
      <th class="num" data-sort="reorder">Re-order <span class="arrow">▼</span></th>
      <th data-sort="due">Due <span class="arrow">▼</span></th>
      <th data-sort="released">Last released <span class="arrow">▼</span></th>
      <th data-sort="status">Status <span class="arrow">▼</span></th>
    </tr></thead>
    <tbody id="tbody"></tbody>
  </table>
</div>

__WARNINGS__

</div>
<script>
const ITEMS = __ITEMS__;
const REPORT_DATE = "__REPORTDATE_ISO__";

const fmt = n => (n === null || n === undefined) ? '-' : Number(n).toLocaleString('en-US');
const fmt1 = n => (n === null || n === undefined) ? '-' : Number(n).toFixed(1);

const CRIT_DAYS = __CRIT_DAYS__, WATCH_DAYS = __WATCH_DAYS__;
function tier(i){
  if (i.onHand === 0) return 'crit';
  if (i.daysCover === null) return 'none';
  if (i.daysCover <= CRIT_DAYS) return 'crit';
  if (i.daysCover <= WATCH_DAYS) return 'warn';
  return 'good';
}
const tierLabel = t => ({crit:'Critical', warn:'Watch', good:'Healthy', none:'No rate'})[t];

/* ---- KPIs ---- */
const zero    = ITEMS.filter(i => i.onHand === 0);
const crit    = ITEMS.filter(i => tier(i) === 'crit');
const warn    = ITEMS.filter(i => tier(i) === 'warn');
const inproc  = ITEMS.filter(i => i.inProcess);
const ipTotal = inproc.reduce((s,i) => s + i.inProcess, 0);
const ohTotal = ITEMS.reduce((s,i) => s + (i.onHand || 0), 0);

document.getElementById('kpis').innerHTML = `
  <div class="kpi ${crit.length ? 'alert' : ''}">
    <div class="kpi-label">Critical lines</div>
    <div class="kpi-value ${crit.length ? 'alert' : ''}">${crit.length}<small> / ${ITEMS.length}</small></div>
    <div class="kpi-sub">${zero.length} at zero on hand</div>
  </div>
  <div class="kpi ${warn.length > 2 ? 'warn' : ''}">
    <div class="kpi-label">Watch lines</div>
    <div class="kpi-value ${warn.length > 2 ? 'warn' : ''}">${warn.length}</div>
    <div class="kpi-sub">__CRIT_DAYS__-__WATCH_DAYS__ days of cover</div>
  </div>
  <div class="kpi">
    <div class="kpi-label">Total on hand</div>
    <div class="kpi-value">${fmt(ohTotal)}</div>
    <div class="kpi-sub">units across ${ITEMS.length} lines</div>
  </div>
  <div class="kpi">
    <div class="kpi-label">In process</div>
    <div class="kpi-value">${fmt(ipTotal)}</div>
    <div class="kpi-sub">units on ${inproc.length} lines</div>
  </div>`;

/* ---- attention cards ---- */
const attention = ITEMS
  .filter(i => tier(i) === 'crit' || tier(i) === 'warn')
  .sort((a,b) => {
    const av = a.daysCover === null ? -1 : a.daysCover;
    const bv = b.daysCover === null ? -1 : b.daysCover;
    return av - bv;
  });

document.getElementById('cards').innerHTML = attention.length ? attention.map(i => {
  const t = tier(i);
  return `<div class="card ${t === 'warn' ? 'warn' : ''}">
    <div class="chip">${i.tag}</div>
    <div class="sku">${i.name}</div>
    <div class="stat"><span class="stat-label">On hand</span><span class="stat-val ${i.onHand === 0 ? 'crit' : ''}">${fmt(i.onHand)}</span></div>
    <div class="stat"><span class="stat-label">Burn / day</span><span class="stat-val">${fmt(i.burn)}</span></div>
    <div class="stat"><span class="stat-label">Days cover</span><span class="stat-val ${t}">${fmt1(i.daysCover)}</span></div>
    ${i.inProcess ? `<div class="ip">▲ ${fmt(i.inProcess)} in process${i.due ? ` · due ${i.due}` : ''}</div>` : ''}
  </div>`;
}).join('') : '<div style="color:var(--text-faint);padding:20px;">Nothing under ${WATCH_DAYS} days of cover.</div>';

/* ---- table ---- */
let sortKey = 'daysCover', sortDir = 1, filter = 'all';
const tbody = document.getElementById('tbody');

function render(){
  let rows = ITEMS.slice();
  if (filter === 'crit')   rows = rows.filter(i => tier(i) === 'crit');
  if (filter === 'warn')   rows = rows.filter(i => tier(i) === 'warn');
  if (filter === 'ip')     rows = rows.filter(i => i.inProcess);
  if (filter === 'nodata') rows = rows.filter(i => i.burn === null);

  rows.sort((a,b) => {
    let av = a[sortKey], bv = b[sortKey];
    if (sortKey === 'status'){ av = tier(a); bv = tier(b); }
    if (av === null || av === undefined) av = sortDir === 1 ?  Infinity : -Infinity;
    if (bv === null || bv === undefined) bv = sortDir === 1 ?  Infinity : -Infinity;
    if (av < bv) return -sortDir;
    if (av > bv) return sortDir;
    return 0;
  });

  tbody.innerHTML = rows.map(i => {
    const t = tier(i);
    return `<tr>
      <td class="tagcell">${i.tag}</td>
      <td class="name">${i.name}</td>
      <td class="num ${i.onHand === 0 ? 'crit' : ''}">${fmt(i.onHand)}</td>
      <td class="num ip-badge">${i.inProcess ? fmt(i.inProcess) : '<span class="faint">-</span>'}</td>
      <td class="num ${i.burn === null ? 'faint' : ''}">${i.burn === null ? (i.burnNote ? i.burnNote : '-') : fmt(i.burn)}</td>
      <td class="num ${t}">${fmt1(i.daysCover)}</td>
      <td class="num ${i.reorder ? 'warn' : 'faint'}">${i.reorder ? fmt(i.reorder) : '-'}</td>
      <td class="faint">${i.due || '-'}</td>
      <td class="faint">${i.released || '-'}</td>
      <td><span class="pill ${t}">${tierLabel(t)}</span></td>
    </tr>`;
  }).join('');

  document.querySelectorAll('th').forEach(th => {
    th.classList.toggle('sorted', th.dataset.sort === sortKey);
    const a = th.querySelector('.arrow');
    if (a) a.textContent = sortDir === 1 ? '▲' : '▼';
  });
}

document.querySelectorAll('th').forEach(th => th.addEventListener('click', () => {
  const k = th.dataset.sort;
  if (sortKey === k) sortDir *= -1;
  else { sortKey = k; sortDir = (k === 'tag' || k === 'name' || k === 'status') ? 1 : 1; }
  render();
}));
document.querySelectorAll('.filter-btn').forEach(b => b.addEventListener('click', () => {
  document.querySelectorAll('.filter-btn').forEach(x => x.classList.remove('active'));
  b.classList.add('active');
  filter = b.dataset.f;
  render();
}));
render();
</script></body></html>
"""


def build(items, warnings, missing, src: Path, report_dt: date, date_src: str, others, cfg: dict = None):
    cfg = cfg or DEFAULT_CFG
    disp = cfg["display"]
    th = cfg["thresholds"]

    age = (date.today() - report_dt).days
    stale_cls = "" if age <= th["stale_after_days"] else "stale"

    freshness = ""
    if disp.get("show_freshness_banner") and age > th["stale_after_days"]:
        freshness += (
            f'<div class="banner"><b>Report is {age} days old</b>'
            f'<span>Newest file in the folder is dated {report_dt:%b %d, %Y} (from {date_src}). '
            f'The numbers below are as of the report date, not today.</span></div>'
        )
    if disp.get("show_missing_sku_banner") and missing:
        names = ", ".join(NOM_META[n][1] for n in missing)
        freshness += (
            f'<div class="banner blue"><b>Not on TRG\'s sheet</b>'
            f'<span>{len(missing)} of our SKUs do not appear in this report: {names}.</span></div>'
        )

    warn_html = ""
    if warnings and disp.get("show_parse_warnings", True):
        lis = "".join(f"<li>{w}</li>" for w in warnings)
        warn_html = f'<div class="warnlist"><b>Parse warnings - {len(warnings)}</b><ul>{lis}</ul></div>'

    title = disp["title"]
    emph = disp.get("title_emphasis")
    title_html = title.replace(emph, f"<em>{emph}</em>", 1) if emph and emph in title else title

    html = (HTML
            .replace("__ITEMS__", json.dumps(items))
            .replace("__CRIT_DAYS__", str(th["critical_days"]))
            .replace("__WATCH_DAYS__", str(th["watch_days"]))
            .replace("__TITLE_HTML__", title_html)
            .replace("__SUBTITLE__", disp["subtitle"])
            .replace("__REPORTDATE_ISO__", report_dt.isoformat())
            .replace("__REPORTDATE__", report_dt.strftime("%b %d, %Y").replace(" 0", " "))
            .replace("__SRCFILE__", src.name)
            .replace("__GENERATED__", datetime.now().strftime("%b %d, %Y %I:%M %p").replace(" 0", " "))
            .replace("__FRESHNESS__", freshness)
            .replace("__WARNINGS__", warn_html)
            .replace("__STALECLS__", stale_cls))
    return html


def run(config=None, folder=None, out=None, verbose=True):
    """Generate the dashboard. Safe to call from a Jupyter notebook.

        import trg_dashboard as trg
        result = trg.run(config=r'Configuration\\trg_config.yaml')
        result['items']      # parsed rows, as a list of dicts
        result['warnings']   # anything the parser could not read
        result['output']     # where the HTML was written

    Raises FileNotFoundError rather than SystemExit, so a bad path does not
    try to shut the kernel down.
    """
    def say(*a):
        if verbose:
            print(*a)

    cfg = load_config(Path(config) if config else None)
    base = Path(config).expanduser().resolve().parent if config else Path.cwd()

    src_folder = Path(folder or cfg["paths"]["source_folder"]).expanduser()
    if not src_folder.is_absolute():
        src_folder = base / src_folder
    if not src_folder.is_dir():
        raise FileNotFoundError(f"Source folder not found:\n  {src_folder}")

    (src, report_dt, date_src), others = find_latest_cfg(src_folder, cfg)
    say(f"Using {src.name}  (report date {report_dt}, from {date_src})")
    if others:
        say(f"  older files skipped: {', '.join(others)}")

    items, warnings, missing = parse_report(src, report_dt, cfg)
    say(f"Parsed {len(items)} lines · {len(warnings)} warning(s) · {len(missing)} unexpected SKU gap(s)")
    for w in warnings:
        say(f"  ! {w}")

    html = build(items, warnings, missing, src, report_dt, date_src, others, cfg)

    out_path = Path(out) if out else Path(cfg["paths"]["output_html"])
    out_path = out_path.expanduser()
    if not out_path.is_absolute():
        out_path = base / out_path
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(html, encoding="utf-8")
    say(f"Wrote {out_path}")

    archived = None
    arch = cfg["paths"].get("archive_folder")
    if arch:
        ap_dir = Path(arch).expanduser()
        if not ap_dir.is_absolute():
            ap_dir = base / ap_dir
        ap_dir.mkdir(parents=True, exist_ok=True)
        archived = ap_dir / f"TRG_Dashboard_{report_dt:%Y-%m-%d}.html"
        archived.write_text(html, encoding="utf-8")
        say(f"Archived {archived}")

    return {"items": items, "warnings": warnings, "missing": missing,
            "source": src, "reportDate": report_dt,
            "output": out_path, "archived": archived, "config": cfg}


def main(argv=None):
    ap = argparse.ArgumentParser(description="Generate the TRG supplier inventory dashboard.")
    ap.add_argument("--config", help="Path to trg_config.yaml")
    ap.add_argument("--folder", help="Override the source folder from the config")
    ap.add_argument("--out", help="Override the output HTML path")

    # parse_known_args so a notebook kernel's own arguments (-f kernel.json)
    # are ignored instead of being read as --folder.
    args, _ = ap.parse_known_args(argv)

    try:
        run(config=args.config, folder=args.folder, out=args.out)
    except FileNotFoundError as e:
        raise SystemExit(str(e))


if __name__ == "__main__":
    main()
