"""
process_data.py
===============
Reads raw CSVs from fetch_data.py, cleans and reshapes them,
and writes analysis-ready files to ./data/processed/.

Usage
-----
    python process_data.py
    python process_data.py --state WA     # also build deep-dive JSON for WA

Outputs (./data/processed/)
---------------------------
    federal_clean.csv          Federal wage, year by year
    state_long.csv             Long format: year, state, minwage, state_name
    state_latest.csv           Latest wage per state (for map / bar chart)
    local_clean.csv            UCB localities with derived fields
    national_overview.csv      Federal floor + state median/max each year
    federal_history.json       Step-function points for Chart.js
    all_states_above_federal.json   History for every state above $7.25
    states_summary.json        Per-state metadata (raises, premium, etc.)
    ca_detail.json             California deep-dive (YoY, cumulative %)
    [state]_detail.json        Same for any --state argument
"""

import os
import json
import argparse
import pandas as pd
import numpy as np

ROOT = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.join(ROOT, "..")
DATA_DIR = os.path.join(ROOT, "..", "data")
OUT_DIR  = os.path.join(DATA_DIR, "processed")
os.makedirs(OUT_DIR, exist_ok=True)

FEDERAL_FLOOR = 7.25

STATE_NAMES = {
    "AL":"Alabama","AK":"Alaska","AZ":"Arizona","AR":"Arkansas","CA":"California",
    "CO":"Colorado","CT":"Connecticut","DE":"Delaware","DC":"Washington D.C.",
    "FL":"Florida","GA":"Georgia","HI":"Hawaii","ID":"Idaho","IL":"Illinois",
    "IN":"Indiana","IA":"Iowa","KS":"Kansas","KY":"Kentucky","LA":"Louisiana",
    "ME":"Maine","MD":"Maryland","MA":"Massachusetts","MI":"Michigan","MN":"Minnesota",
    "MS":"Mississippi","MO":"Missouri","MT":"Montana","NE":"Nebraska","NV":"Nevada",
    "NH":"New Hampshire","NJ":"New Jersey","NM":"New Mexico","NY":"New York",
    "NC":"North Carolina","ND":"North Dakota","OH":"Ohio","OK":"Oklahoma","OR":"Oregon",
    "PA":"Pennsylvania","RI":"Rhode Island","SC":"South Carolina","SD":"South Dakota",
    "TN":"Tennessee","TX":"Texas","UT":"Utah","VT":"Vermont","VA":"Virginia",
    "WA":"Washington","WV":"West Virginia","WI":"Wisconsin","WY":"Wyoming",
}


# ── Loaders ───────────────────────────────────────────────────────────────────
def _require(filename):
    path = os.path.join(DATA_DIR, filename)
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"'{path}' not found. Run fetch_data.py first."
        )
    return path


def load_federal():
    df = pd.read_csv(_require("federal_minwage.csv"))
    df.columns = ["year", "federal_minwage"]
    df["year"]            = df["year"].astype(int)
    df["federal_minwage"] = pd.to_numeric(df["federal_minwage"], errors="coerce")
    return df.dropna().sort_values("year").reset_index(drop=True)


def load_state_long():
    df = pd.read_csv(_require("state_minwage_long.csv"))
    df["year"]    = df["year"].astype(int)
    df["minwage"] = pd.to_numeric(df["minwage"], errors="coerce")
    df["state"]   = df["state"].str.upper()
    return df.dropna(subset=["minwage"]).sort_values(["state","year"]).reset_index(drop=True)


def load_local():
    path = os.path.join(DATA_DIR, "ucb_local_wages.csv")
    if not os.path.exists(path):
        print("  ucb_local_wages.csv not found — skipping local processing")
        return pd.DataFrame()
    return pd.read_csv(path)


# ── 1. Federal CSV ────────────────────────────────────────────────────────────
def process_federal(fed):
    path = os.path.join(OUT_DIR, "federal_clean.csv")
    fed.to_csv(path, index=False)
    print(f"  Federal: {len(fed)} rows -> {path}")
    return fed


# ── 2. State CSVs ─────────────────────────────────────────────────────────────
def process_state(long):
    # Latest snapshot
    latest_year = long["year"].max()
    latest = (long[long["year"] == latest_year]
              .copy()
              .sort_values("minwage", ascending=False)
              .reset_index(drop=True))

    path_long   = os.path.join(OUT_DIR, "state_long.csv")
    path_latest = os.path.join(OUT_DIR, "state_latest.csv")
    long.to_csv(path_long, index=False)
    latest.to_csv(path_latest, index=False)
    above = (latest["minwage"] > FEDERAL_FLOOR).sum()
    print(f"  State long: {len(long):,} rows -> {path_long}")
    print(f"  State latest ({latest_year}): {len(latest)} states, {above} above federal")
    return long, latest


# ── 3. Local CSV ──────────────────────────────────────────────────────────────
def process_local(local):
    if local.empty:
        return local
    local = local.copy()
    local["state_name"]   = local["state"].map(STATE_NAMES).fillna(local["state"])
    local["pct_increase"] = ((local["wage_increase"] / local["initial_wage"]) * 100).round(1)
    path = os.path.join(OUT_DIR, "local_clean.csv")
    local.to_csv(path, index=False)
    print(f"  Local: {len(local)} localities -> {path}")
    return local


# ── 4. National overview ──────────────────────────────────────────────────────
def build_national_overview(fed, long):
    if fed.empty or long.empty:
        return pd.DataFrame()
    summary = (
        long.groupby("year")["minwage"]
        .agg(state_median="median", state_max="max", state_min="min", state_mean="mean")
        .reset_index()
    )
    merged = pd.merge(summary, fed, on="year", how="outer").sort_values("year")
    merged["federal_minwage"] = merged["federal_minwage"].ffill()
    merged = merged.dropna(subset=["state_median"]).round(2)
    path = os.path.join(OUT_DIR, "national_overview.csv")
    merged.to_csv(path, index=False)
    print(f"  National overview: {len(merged)} rows -> {path}")
    return merged


# ── 5. Federal history JSON (step-function for Chart.js) ──────────────────────
def build_federal_json(fed):
    records = fed.sort_values("year").to_dict("records")
    points  = []
    for i, row in enumerate(records):
        points.append({"year": row["year"], "wage": round(float(row["federal_minwage"]), 2)})
        if i + 1 < len(records):
            points.append({"year": records[i+1]["year"] - 0.01,
                           "wage": round(float(row["federal_minwage"]), 2)})
    obj = {
        "label":       "Federal",
        "description": "U.S. federal minimum wage ($/hr). Unchanged since July 2009.",
        "source":      "Vaghul & Zipperer (2022) v1.4.0",
        "points":      points,
    }
    path = os.path.join(OUT_DIR, "federal_history.json")
    with open(path, "w") as f:
        json.dump(obj, f, indent=2)
    print(f"  Federal JSON: {len(points)} step points -> {path}")
    return obj


# ── 6. All states above federal JSON ─────────────────────────────────────────
def build_states_above_json(long):
    result = {}
    for state, grp in long.groupby("state"):
        grp = grp.sort_values("year")
        if grp["minwage"].max() <= FEDERAL_FLOOR:
            continue
        name        = STATE_NAMES.get(state, state)
        latest      = float(grp["minwage"].iloc[-1])
        first_above = int(grp[grp["minwage"] > FEDERAL_FLOOR]["year"].min())
        total_inc   = round(latest - float(grp["minwage"].iloc[0]), 2)
        rows        = grp.to_dict("records")
        points      = []
        for i, row in enumerate(rows):
            points.append({"year": row["year"], "wage": round(float(row["minwage"]), 2)})
            if i + 1 < len(rows) and rows[i+1]["year"] - row["year"] > 1:
                points.append({"year": rows[i+1]["year"] - 0.01,
                               "wage": round(float(row["minwage"]), 2)})
        result[state] = {
            "state":            state,
            "state_name":       name,
            "latest_wage":      latest,
            "first_above_year": first_above,
            "total_increase":   total_inc,
            "n_years_data":     len(grp),
            "points":           points,
        }
    path = os.path.join(OUT_DIR, "all_states_above_federal.json")
    with open(path, "w") as f:
        json.dump(result, f, indent=2)
    print(f"  States above federal: {len(result)} states -> {path}")
    return result


# ── 7. States summary JSON ────────────────────────────────────────────────────
def build_states_summary_json(long):
    latest_year = long["year"].max()
    rows = []
    for state, grp in long.groupby("state"):
        grp     = grp.sort_values("year")
        lat_row = grp[grp["year"] == latest_year]["minwage"].values
        latest  = float(lat_row[0]) if len(lat_row) else float(grp["minwage"].iloc[-1])
        first   = float(grp["minwage"].iloc[0])
        above   = grp[grp["minwage"] > FEDERAL_FLOOR]
        rows.append({
            "state":            state,
            "state_name":       STATE_NAMES.get(state, state),
            "latest_wage":      round(latest, 2),
            "first_wage":       round(first, 2),
            "total_increase":   round(latest - first, 2),
            "above_federal":    latest > FEDERAL_FLOOR,
            "premium":          round(max(0.0, latest - FEDERAL_FLOOR), 2),
            "first_above_year": int(above["year"].min()) if len(above) else None,
            "n_increases":      int((grp["minwage"].diff() > 0).sum()),
        })
    rows.sort(key=lambda x: -x["latest_wage"])
    path = os.path.join(OUT_DIR, "states_summary.json")
    with open(path, "w") as f:
        json.dump(rows, f, indent=2)
    print(f"  States summary: {len(rows)} states -> {path}")
    return rows


# ── 8. State deep-dive JSON ───────────────────────────────────────────────────
def build_state_detail_json(long, state="CA"):
    grp = long[long["state"] == state].sort_values("year").copy()
    if grp.empty:
        print(f"  WARNING: No data for '{state}' — skipping detail JSON")
        return {}
    grp["yoy_change"]    = grp["minwage"].diff().round(2)
    grp["cum_pct"]       = ((grp["minwage"] / grp["minwage"].iloc[0] - 1) * 100).round(1)
    grp["above_federal"] = (grp["minwage"] - FEDERAL_FLOOR).clip(lower=0).round(2)
    step_years = grp[grp["yoy_change"] > 0]["year"].tolist()
    detail = {
        "state":          state,
        "state_name":     STATE_NAMES.get(state, state),
        "first_year":     int(grp["year"].iloc[0]),
        "latest_year":    int(grp["year"].iloc[-1]),
        "first_wage":     round(float(grp["minwage"].iloc[0]), 2),
        "latest_wage":    round(float(grp["minwage"].iloc[-1]), 2),
        "total_increase": round(float(grp["minwage"].iloc[-1] - grp["minwage"].iloc[0]), 2),
        "n_increases":    len(step_years),
        "step_years":     step_years,
        "source":         "Vaghul & Zipperer (2022) v1.4.0",
        "annual_series":  [
            {
                "year":          int(r["year"]),
                "wage":          round(float(r["minwage"]), 2),
                "yoy_change":    round(float(r.get("yoy_change") or 0), 2),
                "cum_pct":       round(float(r.get("cum_pct") or 0), 1),
                "above_federal": round(float(r.get("above_federal") or 0), 2),
                "is_step":       int(r["year"]) in step_years,
            }
            for r in grp.fillna(0).to_dict("records")
        ],
    }
    fname = f"{state.lower()}_detail.json"
    path  = os.path.join(OUT_DIR, fname)
    with open(path, "w") as f:
        json.dump(detail, f, indent=2)
    print(f"  {STATE_NAMES.get(state,state)} detail: "
          f"{len(grp)} years, {len(step_years)} increases -> {path}")
    return detail


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="Process minimum wage data")
    parser.add_argument("--state", default=None,
                        help="Also build a deep-dive JSON for this state (e.g. WA)")
    args = parser.parse_args()

    print("=" * 60)
    print("  U.S. Minimum Wage Data Processor")
    print("=" * 60)

    print("\nLoading raw files...")
    fed  = load_federal()
    long = load_state_long()
    local= load_local()
    print(f"  Federal: {len(fed)} rows | State: {len(long):,} rows | "
          f"Local: {len(local)} localities")

    print("\n[1/8] Federal CSV")
    process_federal(fed)

    print("\n[2/8] State CSVs")
    long, latest = process_state(long)

    print("\n[3/8] Local CSV")
    process_local(local)

    print("\n[4/8] National overview CSV")
    build_national_overview(fed, long)

    print("\n[5/8] Federal history JSON")
    build_federal_json(fed)

    print("\n[6/8] All states above federal JSON")
    build_states_above_json(long)

    print("\n[7/8] States summary JSON")
    build_states_summary_json(long)

    print("\n[8/8] State deep-dive JSONs")
    build_state_detail_json(long, "CA")
    if args.state and args.state.upper() != "CA":
        build_state_detail_json(long, args.state.upper())

    print(f"""
Done. Files saved to ./data/processed/
  CSVs:  federal_clean, state_long, state_latest, local_clean, national_overview
  JSONs: federal_history, all_states_above_federal, states_summary, ca_detail

  Tip: serve with  python -m http.server 8000
       then open   http://localhost:8000/dashboard_with_map.html
""")


if __name__ == "__main__":
    main()
