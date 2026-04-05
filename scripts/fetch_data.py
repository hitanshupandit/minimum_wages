"""
fetch_data.py
=============
Downloads raw minimum wage data from public sources and saves to ./data/.

Sources
-------
1. Vaghul & Zipperer (2022) — Federal + state historical wages, CC BY 4.0
   https://github.com/benzipperer/historicalminwage

2. UC Berkeley Labor Center — City & county ordinance inventory, Dec 2025
   https://laborcenter.berkeley.edu/minimum-wage-living-wage-resources/

Usage
-----
    python fetch_data.py                  # fetch all
    python fetch_data.py --skip-ucb       # skip UCB Excel (use embedded data)

Outputs (./data/)
-----------------
    federal_minwage.csv       Federal floor by year (1938-present)
    state_minwage_wide.csv    State wages x year, wide format
    state_minwage_long.csv    State wages, long/tidy format
    state_snapshot.csv        Latest wage per state
    ucb_local_wages.csv       Parsed UCB city/county ordinances
"""

import os
import re
import sys
import argparse
import requests
import pandas as pd
from io import StringIO
from datetime import datetime

# Paths
ROOT = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.join(ROOT, "..")
DATA_DIR = os.path.join(ROOT, "..", "data")
os.makedirs(DATA_DIR, exist_ok=True)

# Sources
ZIPPERER_BASE = "https://raw.githubusercontent.com/benzipperer/historicalminwage/master/release"
FEDERAL_URL   = f"{ZIPPERER_BASE}/minwage_federal.csv"
STATE_URL     = f"{ZIPPERER_BASE}/minwage_state_annual.csv"

UCB_WEB_URL = (
    "https://laborcenter.berkeley.edu/wp-content/uploads/2025/12/"
    "UCB-Labor-Center-City-and-County-Minimum-Wage-Inventory.xlsx"
)
UCB_LOCAL_NAMES = [
    "UCB-Labor-Center-City-and-County-Minimum-Wage-Inventory.xlsx",
    "ucb_inventory.xlsx",
]

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


def _fetch_text(url, label):
    print(f"  Downloading {label}...")
    try:
        r = requests.get(url, timeout=30)
        r.raise_for_status()
        print(f"    OK  {len(r.content) // 1024} KB")
        return r.text
    except requests.RequestException as e:
        print(f"    FAILED: {e}")
        return None


def _fetch_bytes(url, label):
    print(f"  Downloading {label}...")
    try:
        r = requests.get(url, timeout=60)
        r.raise_for_status()
        print(f"    OK  {len(r.content) // 1024} KB")
        return r.content
    except requests.RequestException as e:
        print(f"    FAILED: {e}")
        return None


def _save(df, filename):
    path = os.path.join(DATA_DIR, filename)
    df.to_csv(path, index=False)
    print(f"    Saved -> {path}  ({len(df):,} rows)")
    return path


def fetch_federal():
    text = _fetch_text(FEDERAL_URL, "federal minimum wage (Zipperer)")
    if text is None:
        return pd.DataFrame()
    df = pd.read_csv(StringIO(text))
    df.columns = [c.strip().lower() for c in df.columns]
    year_col = next(c for c in df.columns if "year" in c)
    wage_col = next(c for c in df.columns if c != year_col)
    df = df.rename(columns={year_col: "year", wage_col: "federal_minwage"})
    df["year"]            = df["year"].astype(int)
    df["federal_minwage"] = pd.to_numeric(df["federal_minwage"], errors="coerce")
    df = df[["year", "federal_minwage"]].dropna().sort_values("year")
    _save(df, "federal_minwage.csv")
    return df


def fetch_state():
    text = _fetch_text(STATE_URL, "state annual wages (Zipperer)")
    if text is None:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    wide = pd.read_csv(StringIO(text))
    wide.columns = [c.strip().lower() for c in wide.columns]
    year_col   = next(c for c in wide.columns if "year" in c)
    state_cols = [c for c in wide.columns if c != year_col]
    wide = wide.rename(columns={year_col: "year"})
    wide["year"] = wide["year"].astype(int)
    _save(wide, "state_minwage_wide.csv")

    long = wide.melt(id_vars="year", value_vars=state_cols,
                     var_name="state", value_name="minwage")
    long["state"]      = long["state"].str.upper()
    long["state_name"] = long["state"].map(STATE_NAMES).fillna(long["state"])
    long["minwage"]    = pd.to_numeric(long["minwage"], errors="coerce")
    long = long.dropna(subset=["minwage"]).sort_values(["state","year"]).reset_index(drop=True)
    _save(long, "state_minwage_long.csv")

    snapshot = (
        long.loc[long.groupby("state")["year"].idxmax()]
        .sort_values("minwage", ascending=False).reset_index(drop=True)
    )
    _save(snapshot, "state_snapshot.csv")
    above = (snapshot["minwage"] > 7.25).sum()
    print(f"    Latest year: {long['year'].max()} | {above} states above federal $7.25")
    return wide, long, snapshot


def _find_local_ucb():
    for name in UCB_LOCAL_NAMES:  # search project root
        path = os.path.join(ROOT, name)
        if os.path.exists(path):
            return path
    cached = os.path.join(DATA_DIR, "ucb_raw.xlsx")
    if os.path.exists(cached):
        return cached
    return None


def _parse_ucb_excel(path):
    raw = pd.read_excel(path, sheet_name=0, header=None)
    records = []
    for i, row in raw.iterrows():
        if i < 9:
            continue
        if i > 200:
            break
        vals   = list(row)
        first  = vals[0] if pd.notna(vals[0]) else None
        second = vals[1] if len(vals) > 1 and pd.notna(vals[1]) else None
        if first is None or second is None:
            continue
        if not isinstance(first, (int, float)) or isinstance(first, bool):
            continue
        if not isinstance(second, str):
            continue
        name_raw   = str(second).strip()
        state_m    = re.search(r",\s*([A-Z]{2})\b", name_raw)
        state      = state_m.group(1) if state_m else "DC"
        name_clean = re.sub(r"\s*\([\d;, a-z\s]+\)", "", name_raw)
        name_clean = re.sub(r"\d+$", "", name_clean).split("\n")[0].strip().rstrip("*").strip()
        wages = [round(float(v), 2) for v in vals[2:]
                 if isinstance(v, (int, float)) and not isinstance(v, bool) and 5 < v < 35]
        dates = [v.strftime("%Y-%m-%d") for v in vals[2:] if isinstance(v, datetime)]
        if not wages:
            continue
        records.append({
            "locality":            name_clean,
            "state":               state,
            "level":               "County" if "county" in name_clean.lower() else "City",
            "initial_wage":        wages[0],
            "current_wage":        max(wages),
            "wage_increase":       round(max(wages) - wages[0], 2),
            "n_steps":             len(wages),
            "first_increase_date": dates[0] if dates else None,
            "last_scheduled_date": dates[-1] if dates else None,
            "wage_history":        "|".join(map(str, wages)),
        })
    return pd.DataFrame(records)


def fetch_ucb_local(skip=False):
    if skip:
        print("  Skipping UCB download (--skip-ucb)")
        return pd.DataFrame()
    local_path = _find_local_ucb()
    if local_path:
        print(f"  Using local UCB file: {local_path}")
        path = local_path
    else:
        data = _fetch_bytes(UCB_WEB_URL, "UCB local wage inventory (Excel)")
        if data is None:
            print("  WARNING: UCB Excel unavailable. Dashboards will use embedded data.")
            return pd.DataFrame()
        path = os.path.join(DATA_DIR, "ucb_raw.xlsx")
        with open(path, "wb") as f:
            f.write(data)
    df = _parse_ucb_excel(path)
    if df.empty:
        print("  WARNING: UCB Excel parsed but returned no rows.")
        return df
    _save(df, "ucb_local_wages.csv")
    print(f"    {(df['level']=='City').sum()} cities, {(df['level']=='County').sum()} counties")
    return df


def main():
    parser = argparse.ArgumentParser(description="Fetch U.S. minimum wage data")
    parser.add_argument("--skip-ucb", action="store_true",
                        help="Skip UCB Excel download")
    args = parser.parse_args()

    print("=" * 60)
    print("  U.S. Minimum Wage Data Fetcher")
    print("  Zipperer (2022) + UC Berkeley Labor Center")
    print("=" * 60)

    print("\n[1/3] Federal minimum wage")
    fed = fetch_federal()
    if not fed.empty:
        print(f"    {fed['year'].min()}-{fed['year'].max()} ({len(fed)} rows)")

    print("\n[2/3] State minimum wages")
    wide, long, snap = fetch_state()
    if not long.empty:
        print(f"    {long['state'].nunique()} states, {long['year'].min()}-{long['year'].max()}")

    print("\n[3/3] UCB local wages")
    fetch_ucb_local(skip=args.skip_ucb)

    print(f"""
Done. Files saved to ./data/
  Run next: python process_data.py
""")


if __name__ == "__main__":
    main()
