"""
master_hub.py - the single landing page for the Corrugate Master Module.

Builds one page of tiles, each linking to a dashboard and reporting when that
dashboard was last generated, so a stale tile shows up before you present it.

    python master_hub.py --config "Configuration\\hub_config.yaml"

Notebook:
    import master_hub as hub
    result = hub.run(config=PROJECT / 'Configuration' / 'hub_config.yaml')

Tiles are defined entirely in the config - add, reorder, or retire one there
without touching this file.
"""

from __future__ import annotations
import argparse
import html
import os
from datetime import datetime, date
from pathlib import Path

try:
    import yaml
except ImportError:
    yaml = None

DEFAULT_CFG = {
    "paths": {"module_root": ".", "output_html": "Corrugate Master Module.html"},
    "display": {"title": "Corrugate Master Module", "title_emphasis": "Master",
                "eyebrow": "Power Stop", "stale_after_days": 1},
    "tiles": [],
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


def win(p: str) -> str:
    """Treat a config path as Windows-style regardless of host OS."""
    return str(p).replace("\\", "/")


def describe_age(ts: float | None, stale_days: int):
    """Return (text, state) for a file's modified time."""
    if ts is None:
        return "not generated yet", "missing"
    dt = datetime.fromtimestamp(ts)
    age = (date.today() - dt.date()).days
    if age == 0:
        text = "updated today, " + dt.strftime("%I:%M %p").lstrip("0")
    elif age == 1:
        text = "updated yesterday"
    else:
        text = f"updated {age} days ago"
    return text, ("fresh" if age <= stale_days else "stale")


HTML = """<!DOCTYPE html>
<html lang="en"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>__TITLE_TEXT__</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;700&family=Fraunces:ital,wght@0,400;0,600;0,800;1,400&display=swap" rel="stylesheet">
<style>
:root{--bg:#0f1114;--bg-panel:#16191e;--bg-elev:#1d2128;--border:#272c35;--border-bright:#3a414e;
--text:#e8e9ec;--text-dim:#8a8f99;--text-faint:#5a5f6a;--amber:#f4a52a;--red:#ef4444;--green:#10b981;--blue:#60a5fa;}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'JetBrains Mono',Consolas,monospace;background:var(--bg);color:var(--text);font-size:13px;line-height:1.5;
min-height:100vh;
background-image:radial-gradient(circle at 15% 0%,rgba(244,165,42,.04),transparent 44%),
radial-gradient(circle at 85% 100%,rgba(96,165,250,.035),transparent 44%);}
.container{max-width:1180px;margin:0 auto;padding:44px 32px 90px}

header{border-bottom:1px solid var(--border);padding-bottom:26px;margin-bottom:30px;
display:grid;grid-template-columns:1fr auto;gap:24px;align-items:end}
.eyebrow{font-size:10px;letter-spacing:.2em;color:var(--text-faint);text-transform:uppercase;
margin-bottom:10px;display:flex;align-items:center;gap:10px}
.eyebrow::before{content:'';width:7px;height:7px;border-radius:50%;background:var(--green);box-shadow:0 0 8px var(--green)}
h1{font-family:'Fraunces',serif;font-weight:800;font-size:46px;letter-spacing:-.025em;line-height:1}
h1 em{font-style:italic;font-weight:400;color:var(--amber)}
.header-meta{text-align:right;font-size:11px;color:var(--text-dim);line-height:1.9}
.header-meta .label{color:var(--text-faint);text-transform:uppercase;letter-spacing:.1em}
.header-meta .value{color:var(--text);font-weight:500}

.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(316px,1fr));gap:16px}
.tile{display:block;text-decoration:none;color:inherit;background:var(--bg-panel);
border:1px solid var(--border);padding:24px 24px 18px;position:relative;overflow:hidden;
transition:border-color .16s,transform .16s,background .16s}
.tile::before{content:'';position:absolute;top:0;left:0;width:3px;height:100%;background:var(--border-bright)}
.tile.amber::before{background:var(--amber)}
.tile.blue::before{background:var(--blue)}
.tile.green::before{background:var(--green)}
.tile.red::before{background:var(--red)}
a.tile:hover{border-color:var(--border-bright);background:var(--bg-elev);transform:translateY(-2px)}
.tile.pending{opacity:.55;cursor:default}
.tile-head{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;margin-bottom:10px}
.tile-name{font-family:'Fraunces',serif;font-size:21px;font-weight:600;letter-spacing:-.01em;line-height:1.2}
.tile-blurb{color:var(--text-dim);font-size:12px;line-height:1.7;min-height:40px}
.tile-src{font-size:10px;color:var(--text-faint);letter-spacing:.04em;margin-top:14px;
padding-top:12px;border-top:1px dashed var(--border);word-break:break-word}
.tile-foot{display:flex;justify-content:space-between;align-items:center;gap:10px;
margin-top:12px;padding-top:12px;border-top:1px solid var(--border);font-size:10px;
text-transform:uppercase;letter-spacing:.1em}
.stamp.fresh{color:var(--green)}
.stamp.stale{color:var(--amber)}
.stamp.missing{color:var(--text-faint)}
.arrow{color:var(--text-faint);font-size:14px}
a.tile:hover .arrow{color:var(--amber)}
.pill{display:inline-block;padding:2px 9px;font-size:10px;font-weight:500;letter-spacing:.09em;
text-transform:uppercase;border:1px solid;white-space:nowrap}
.pill.pending{color:var(--text-faint);border-color:var(--border)}
.note{margin-top:14px;padding:10px 14px;border-left:2px solid var(--border-bright);
background:rgba(255,255,255,.02);font-size:11px;color:var(--text-dim);line-height:1.7}


@media(max-width:640px){.container{padding:26px 16px 60px}header{grid-template-columns:1fr}
.header-meta{text-align:left}h1{font-size:32px}.grid{grid-template-columns:1fr}}
</style></head><body><div class="container">

<header>
  <div>
    <div class="eyebrow">__EYEBROW__</div>
    <h1>__TITLE_HTML__</h1>
  </div>
  <div class="header-meta">
    <div><span class="label">Dashboards</span> &nbsp; <span class="value">__NLIVE__ live__NPENDING__</span></div>
    <div><span class="label">Page built</span> &nbsp; <span class="value">__GENERATED__</span></div>
  </div>
</header>

<div class="grid">
__TILES__
</div>

</div></body></html>
"""


def build(cfg: dict, out_path: Path):
    disp = cfg["display"]
    root = Path(win(cfg["paths"]["module_root"]))
    stale_days = int(disp.get("stale_after_days", 1))

    cards, n_live, n_pending = [], 0, 0

    for t in cfg.get("tiles", []):
        name = html.escape(str(t.get("name", "Untitled")))
        blurb = html.escape(str(t.get("blurb", "")))
        source = html.escape(str(t.get("source", "")))
        accent = str(t.get("accent", "plain")).lower()
        pending = str(t.get("status", "live")).lower() == "pending"
        href_raw = win(t.get("href", ""))

        target = (root / href_raw) if href_raw else None
        exists = bool(target and target.is_file())
        ts = target.stat().st_mtime if exists else None
        stamp_txt, stamp_state = describe_age(ts, stale_days)

        # relative link where possible, so the folder can move or sync
        link = ""
        if href_raw:
            try:
                link = os.path.relpath(target, out_path.parent).replace("\\", "/")
            except ValueError:
                link = target.as_uri()

        if pending:
            n_pending += 1
            note = t.get("note")
            cards.append(f"""  <div class="tile {accent} pending">
    <div class="tile-head">
      <div class="tile-name">{name}</div>
      <span class="pill pending">Pending</span>
    </div>
    <div class="tile-blurb">{blurb}</div>
    {f'<div class="note">{html.escape(str(note))}</div>' if note else ''}
    <div class="tile-src">{source}</div>
  </div>""")
            continue

        n_live += 1
        cards.append(f"""  <a class="tile {accent}" href="{html.escape(link)}">
    <div class="tile-head">
      <div class="tile-name">{name}</div>
      <span class="arrow">&rarr;</span>
    </div>
    <div class="tile-blurb">{blurb}</div>
    <div class="tile-src">{source}</div>
    <div class="tile-foot">
      <span class="stamp {stamp_state}">{html.escape(stamp_txt)}</span>
    </div>
  </a>""")

    title = disp["title"]
    emph = disp.get("title_emphasis")
    title_html = title.replace(emph, f"<em>{emph}</em>", 1) if emph and emph in title else title

    return (HTML
            .replace("__TILES__", "\n".join(cards))
            .replace("__TITLE_HTML__", title_html)
            .replace("__TITLE_TEXT__", html.escape(title))
            .replace("__EYEBROW__", html.escape(str(disp.get("eyebrow", ""))))
            .replace("__NLIVE__", str(n_live))
            .replace("__NPENDING__", f" · {n_pending} pending" if n_pending else "")
            .replace("__GENERATED__",
                     datetime.now().strftime("%b %d, %Y %I:%M %p").replace(" 0", " ")))


def run(config=None, out=None, verbose=True):
    """Build the hub page. Safe to call from a notebook."""
    def say(*a):
        if verbose:
            print(*a)

    cfg = load_config(Path(config) if config else None)
    base = Path(config).expanduser().resolve().parent if config else Path.cwd()

    def resolve(p):
        p = Path(win(p)).expanduser()
        return p if p.is_absolute() else base / p

    root = resolve(cfg["paths"]["module_root"])
    if not root.is_dir():
        raise FileNotFoundError(f"Module root not found:\n  {root}")

    out_path = resolve(out or cfg["paths"]["output_html"])
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # report each tile's target so a broken link is obvious at build time
    missing = []
    for t in cfg.get("tiles", []):
        if str(t.get("status", "live")).lower() == "pending":
            say(f"  pending  {t.get('name')}")
            continue
        target = root / win(t.get("href", ""))
        if target.is_file():
            say(f"  ok       {t.get('name')}")
        else:
            missing.append(t.get("name"))
            say(f"  MISSING  {t.get('name')}  ->  {target}")

    cfg["paths"]["module_root"] = str(root)
    out_path.write_text(build(cfg, out_path), encoding="utf-8")
    say(f"Wrote {out_path}")
    if missing:
        say(f"\n  {len(missing)} tile(s) point at a file that does not exist yet:")
        say("  " + ", ".join(str(m) for m in missing))
        say("  They still render, stamped 'not generated yet'.")

    return {"output": out_path, "missing": missing, "config": cfg}


def main(argv=None):
    ap = argparse.ArgumentParser(description="Build the Corrugate Master Module hub page.")
    ap.add_argument("--config", help="Path to hub_config.yaml")
    ap.add_argument("--out", help="Override the output HTML path")
    args, _ = ap.parse_known_args(argv)
    try:
        run(config=args.config, out=args.out)
    except FileNotFoundError as e:
        raise SystemExit(str(e))


if __name__ == "__main__":
    main()
