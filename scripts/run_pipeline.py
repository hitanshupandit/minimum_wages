"""
run_pipeline.py
===============
One-command pipeline: fetch -> process -> validate.
Runs all steps in sequence and prints a final summary.

Usage
-----
    python run_pipeline.py                   # full run
    python run_pipeline.py --skip-ucb        # skip UCB Excel download
    python run_pipeline.py --state WA        # also build WA deep-dive
    python run_pipeline.py --validate-only   # re-run validation on existing data

Steps
-----
    1. fetch_data.py     -> ./data/  (raw CSVs)
    2. process_data.py   -> ./data/processed/  (clean CSVs + JSONs)
    3. Validate outputs and print summary table
"""

import os
import sys
import subprocess
import argparse
import pandas as pd
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))


def run(script, extra_args=None):
    cmd = [sys.executable, os.path.join(ROOT, script)] + (extra_args or [])
    print(f"\n{'='*60}")
    print(f"  Running: {' '.join(cmd)}")
    print(f"{'='*60}")
    result = subprocess.run(cmd, cwd=ROOT)
    if result.returncode != 0:
        print(f"\nERROR: {script} exited with code {result.returncode}")
        sys.exit(result.returncode)


def validate():
    print(f"\n{'='*60}")
    print("  Validation")
    print(f"{'='*60}")

    EXPECTED_RAW = [
        "../data/federal_minwage.csv",
        "../data/state_minwage_long.csv",
        "../data/state_snapshot.csv",
    ]
    EXPECTED_PROCESSED = [
        "../data/processed/federal_clean.csv",
        "../data/processed/state_long.csv",
        "../data/processed/state_latest.csv",
        "../data/processed/national_overview.csv",
        "../data/processed/federal_history.json",
        "../data/processed/all_states_above_federal.json",
        "../data/processed/states_summary.json",
        "../data/processed/ca_detail.json",
    ]
    EXPECTED_DASHBOARDS = [
        "../dashboards/dashboard_with_map.html",
        "../dashboards/ca_minwage_map.html",
    ]

    all_ok = True
    rows = []

    for rel in EXPECTED_RAW + EXPECTED_PROCESSED + EXPECTED_DASHBOARDS:
        path = os.path.join(ROOT, rel)
        exists = os.path.exists(path)
        size   = os.path.getsize(path) // 1024 if exists else 0
        extra  = ""
        if exists and rel.endswith(".csv"):
            try:
                df    = pd.read_csv(path)
                extra = f"{len(df):,} rows x {len(df.columns)} cols"
            except Exception:
                extra = "parse error"
        elif exists and rel.endswith(".html"):
            extra = f"{size} KB"
        elif exists and rel.endswith(".json"):
            import json
            try:
                with open(path) as f:
                    obj = json.load(f)
                if isinstance(obj, list):
                    extra = f"{len(obj)} entries"
                elif isinstance(obj, dict):
                    extra = f"{len(obj)} keys"
            except Exception:
                extra = "parse error"
        status = "OK" if exists else "MISSING"
        if not exists:
            all_ok = False
        rows.append((status, rel, extra))

    col1 = max(len(r[0]) for r in rows)
    col2 = max(len(r[1]) for r in rows)
    print(f"\n  {'STATUS':<8}  {'FILE':<{col2}}  DETAIL")
    print(f"  {'-'*8}  {'-'*col2}  {'------'}")
    for status, path, detail in rows:
        marker = "OK  " if status == "OK" else "MISS"
        print(f"  {marker}      {path:<{col2}}  {detail}")

    print(f"\n  {'All files present.' if all_ok else 'Some files missing — re-run pipeline.'}")

    # Quick data quality checks
    print("\n  Data quality checks:")
    snap_path = os.path.join(ROOT, "../data/state_snapshot.csv")
    if os.path.exists(snap_path):
        snap = pd.read_csv(snap_path)
        above = (snap["minwage"] > 7.25).sum()
        highest = snap.iloc[0]
        print(f"    States above $7.25 federal: {above} / {len(snap)}")
        print(f"    Highest state: {highest['state']} ${highest['minwage']:.2f}")

    ucb_path = os.path.join(ROOT, "../data/ucb_local_wages.csv")
    if os.path.exists(ucb_path):
        ucb = pd.read_csv(ucb_path)
        print(f"    UCB localities: {len(ucb)} "
              f"({(ucb['level']=='City').sum()} cities, {(ucb['level']=='County').sum()} counties)")
        print(f"    Highest local wage: ${ucb['current_wage'].max():.2f} "
              f"({ucb.loc[ucb['current_wage'].idxmax(), 'locality']})")

    return all_ok


def main():
    parser = argparse.ArgumentParser(description="Run full minimum wage data pipeline")
    parser.add_argument("--skip-ucb",       action="store_true", help="Skip UCB download")
    parser.add_argument("--state",          default=None, help="Extra state deep-dive (e.g. WA)")
    parser.add_argument("--validate-only",  action="store_true", help="Only validate existing files")
    args = parser.parse_args()

    print(f"""
{'='*60}
  U.S. Minimum Wage — Full Data Pipeline
  {datetime.now().strftime('%Y-%m-%d %H:%M')}
{'='*60}""")

    if not args.validate_only:
        fetch_args   = ["--skip-ucb"] if args.skip_ucb else []
        process_args = ["--state", args.state.upper()] if args.state else []

        run("fetch_data.py",   fetch_args)
        run("process_data.py", process_args)

    ok = validate()

    if ok:
        print(f"""
{'='*60}
  Pipeline complete!

  Serve locally:
    python -m http.server 8000
    open http://localhost:8000/dashboards/dashboard_with_map.html

  For GitHub Pages:
    git add .
    git commit -m "Update minimum wage data"
    git push
{'='*60}
""")
    else:
        sys.exit(1)


if __name__ == "__main__":
    main()
