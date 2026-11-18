"""
build_all.py - rebuild every dashboard in the Corrugate Master Module, then the hub.

    python build_all.py

Run it from the Corrugate Master Module folder. Each dashboard is built in turn;
if one fails the others still run, and the failure is reported at the end rather
than stopping the batch. The hub is always built last, so its "last generated"
stamps reflect the run that just happened.

Notebook:
    import build_all
    build_all.main()
"""

from __future__ import annotations
import sys
import traceback
from pathlib import Path

HERE = Path(__file__).resolve().parent
CONFIG_DIR = HERE / "Configuration"
sys.path.insert(0, str(HERE))

# (label, module name, config file) - the hub must stay last.
JOBS = [
    ("Burn rate",     "burn_rate_dashboard", "burn_config.yaml"),
    ("TRG on-hand",   "trg_dashboard",       "trg_config.yaml"),
    ("Order & ETA",   "requests_dashboard",  "requests_config.yaml"),
    ("Hub page",      "master_hub",          "hub_config.yaml"),
]


def main():
    print(f"Corrugate Master Module - rebuilding from {HERE}\n")
    failed = []

    for label, mod_name, cfg_name in JOBS:
        cfg = CONFIG_DIR / cfg_name
        print(f"--- {label} " + "-" * (58 - len(label)))

        if not cfg.is_file():
            print(f"  SKIPPED - config not found: {cfg}\n")
            failed.append((label, "config missing"))
            continue

        try:
            mod = __import__(mod_name)
            if not hasattr(mod, "run"):
                raise AttributeError(
                    f"{mod_name}.py has no run() - it is an old copy. "
                    f"Loaded from {getattr(mod, '__file__', '?')}")
            mod.run(config=cfg)
        except Exception as e:
            print(f"  FAILED - {type(e).__name__}: {e}")
            if "--trace" in sys.argv:
                traceback.print_exc()
            failed.append((label, f"{type(e).__name__}: {e}"))
        print()

    print("=" * 64)
    if failed:
        print(f"{len(JOBS) - len(failed)} of {len(JOBS)} built. Problems:")
        for label, why in failed:
            print(f"  - {label}: {why}")
        print("\nRe-run with --trace for full tracebacks.")
        return 1

    print(f"All {len(JOBS)} built. Open 'Corrugate Master Module.html'.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
