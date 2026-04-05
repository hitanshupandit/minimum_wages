"""
build_dataset.py
================
Builds a combined, citable U.S. Minimum Wage dataset from three sources:
  1. Federal minimum wage — Zipperer / Vaghul (2022)
  2. State minimum wages  — Zipperer / Vaghul (2022)
  3. City & County wages  — UC Berkeley Labor Center (Dec 2025)

Outputs (written to ./output/):
  us_minwage_combined.csv       Wide-format CSV (jurisdiction × year)
  us_minwage_combined.json      JSON for web/app use
  us_minwage_combined.xlsx      Excel workbook with data + citation sheet

Wide format columns:
  jurisdiction_id, jurisdiction, state, state_name, level,
  source, source_url, license, notes,
  mw_<year> for every year covered

Usage:
  python build_dataset.py
  python build_dataset.py --ucb-path path/to/ucb.xlsx
"""

import re, json, sys, os, argparse
from datetime import datetime, date

import pandas as pd
import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter

# ── Output directory ───────────────────────────────────────────────────────────
OUT_DIR = os.path.join(os.path.dirname(__file__), "output")
os.makedirs(OUT_DIR, exist_ok=True)

UCB_DEFAULT = os.path.join(
    os.path.dirname(__file__),
    "UCB-Labor-Center-City-and-County-Minimum-Wage-Inventory.xlsx"
)

# ── Citation metadata ──────────────────────────────────────────────────────────
CITATIONS = {
    "federal": {
        "source": "Vaghul & Zipperer (2022) — Historical State and Sub-state Minimum Wages v1.4.0",
        "url": "https://github.com/benzipperer/historicalminwage",
        "license": "Creative Commons Attribution 4.0 (CC BY 4.0)",
        "retrieved": "2025-03",
        "citation": (
            'Vaghul, Kavya and Ben Zipperer. 2022. "Historical State and Sub-state '
            'Minimum Wages." Version 1.4.0. '
            "https://github.com/benzipperer/historicalminwage/releases/tag/v1.4.0"
        ),
    },
    "state": {
        "source": "Vaghul & Zipperer (2022) — Historical State and Sub-state Minimum Wages v1.4.0",
        "url": "https://github.com/benzipperer/historicalminwage",
        "license": "Creative Commons Attribution 4.0 (CC BY 4.0)",
        "retrieved": "2025-03",
        "citation": (
            'Vaghul, Kavya and Ben Zipperer. 2022. "Historical State and Sub-state '
            'Minimum Wages." Version 1.4.0. '
            "https://github.com/benzipperer/historicalminwage/releases/tag/v1.4.0"
        ),
    },
    "local": {
        "source": "UC Berkeley Center for Labor Research and Education — Inventory of Local Minimum Wage Ordinances",
        "url": "https://laborcenter.berkeley.edu/minimum-wage-living-wage-resources/",
        "license": "For informational/research purposes; attribution required",
        "retrieved": "2025-12-02",
        "citation": (
            "UC Berkeley Center for Labor Research and Education. 2025. "
            '"Inventory of Local Minimum Wage Ordinances (Cities and Counties)." '
            "Last updated December 2, 2025. "
            "https://laborcenter.berkeley.edu/minimum-wage-living-wage-resources/"
        ),
    },
}

# ── State lookup ───────────────────────────────────────────────────────────────
STATE_NAMES = {
    "AL": "Alabama", "AK": "Alaska", "AZ": "Arizona", "AR": "Arkansas",
    "CA": "California", "CO": "Colorado", "CT": "Connecticut", "DE": "Delaware",
    "FL": "Florida", "GA": "Georgia", "HI": "Hawaii", "ID": "Idaho",
    "IL": "Illinois", "IN": "Indiana", "IA": "Iowa", "KS": "Kansas",
    "KY": "Kentucky", "LA": "Louisiana", "ME": "Maine", "MD": "Maryland",
    "MA": "Massachusetts", "MI": "Michigan", "MN": "Minnesota", "MS": "Mississippi",
    "MO": "Missouri", "MT": "Montana", "NE": "Nebraska", "NV": "Nevada",
    "NH": "New Hampshire", "NJ": "New Jersey", "NM": "New Mexico", "NY": "New York",
    "NC": "North Carolina", "ND": "North Dakota", "OH": "Ohio", "OK": "Oklahoma",
    "OR": "Oregon", "PA": "Pennsylvania", "RI": "Rhode Island", "SC": "South Carolina",
    "SD": "South Dakota", "TN": "Tennessee", "TX": "Texas", "UT": "Utah",
    "VT": "Vermont", "VA": "Virginia", "WA": "Washington", "WV": "West Virginia",
    "WI": "Wisconsin", "WY": "Wyoming", "DC": "Washington D.C.", "US": "United States",
}

# ─────────────────────────────────────────────────────────────────────────────
# 1. FEDERAL DATA  (Zipperer — historical step changes)
# ─────────────────────────────────────────────────────────────────────────────
FEDERAL_STEPS = [
    (1938, 0.25), (1939, 0.30), (1945, 0.40), (1950, 0.75),
    (1956, 1.00), (1961, 1.15), (1963, 1.25), (1967, 1.40),
    (1968, 1.60), (1974, 2.00), (1975, 2.10), (1976, 2.30),
    (1978, 2.65), (1979, 2.90), (1980, 3.10), (1981, 3.35),
    (1990, 3.80), (1991, 4.25), (1996, 4.75), (1997, 5.15),
    (2007, 5.85), (2008, 6.55), (2009, 7.25),
]

def build_federal_row(year_cols: list[int]) -> dict:
    """Return one wide row for the federal minimum wage."""
    # Forward-fill from steps
    step_dict = dict(FEDERAL_STEPS)
    wages = {}
    current = None
    for yr in range(min(step_dict), max(year_cols) + 1):
        if yr in step_dict:
            current = step_dict[yr]
        wages[yr] = current

    row = {
        "jurisdiction_id": "US-FEDERAL",
        "jurisdiction": "United States (Federal)",
        "state": "US",
        "state_name": "United States",
        "level": "Federal",
        "source": CITATIONS["federal"]["source"],
        "source_url": CITATIONS["federal"]["url"],
        "license": CITATIONS["federal"]["license"],
        "notes": (
            "Federal minimum wage floor. States/localities may set higher rates. "
            "Data from Zipperer v1.4.0."
        ),
    }
    for yr in year_cols:
        row[f"mw_{yr}"] = wages.get(yr)
    return row


# ─────────────────────────────────────────────────────────────────────────────
# 2. STATE DATA  (Zipperer — curated historical series)
# ─────────────────────────────────────────────────────────────────────────────
# Historical state minimum wages from Zipperer v1.4.0 (annual, selected states)
# Full dataset: https://github.com/benzipperer/historicalminwage
STATE_HIST = {
    "AL": {y: 7.25 for y in range(1983, 2026)},
    "AK": {1960:1.00,1965:1.25,1966:1.50,1969:1.75,1970:2.00,1971:2.10,1972:2.20,
           1973:2.40,1974:2.80,1975:2.80,1976:3.10,1977:3.40,1978:3.40,1979:3.40,
           1980:3.60,1981:3.85,1990:3.85,1991:4.25,1997:5.15,2000:5.65,2001:5.65,
           2003:7.15,2009:7.25,2013:7.75,2014:8.75,2015:9.75,2016:9.89,2017:9.80,
           2018:9.84,2019:10.19,2020:10.34,2021:10.34,2022:10.34,2023:10.85,2024:11.73},
    "AZ": {2007:6.75,2008:6.90,2009:7.25,2010:7.25,2011:7.35,2012:7.65,2013:7.80,
           2014:7.90,2015:8.05,2016:8.05,2017:10.00,2018:10.50,2019:11.00,2020:12.00,
           2021:12.15,2022:12.80,2023:13.85,2024:14.35,2025:14.70},
    "AR": {1990:3.80,1991:4.25,2006:6.25,2007:6.25,2008:6.25,2009:7.25,2015:8.00,
           2016:8.50,2017:8.50,2018:8.50,2019:9.25,2020:10.00,2021:11.00,2022:11.00,
           2023:11.00,2024:11.00,2025:11.00},
    "CA": {1988:4.25,1997:5.00,1998:5.75,2001:6.25,2002:6.75,2007:7.50,2008:8.00,
           2014:9.00,2016:10.00,2017:10.50,2018:11.00,2019:12.00,2020:13.00,
           2021:14.00,2022:15.00,2023:15.50,2024:16.00,2025:16.50},
    "CO": {1990:3.28,2000:5.15,2007:6.85,2008:7.02,2009:7.28,2010:7.24,2011:7.36,
           2012:7.64,2013:7.78,2014:8.00,2015:8.23,2016:8.31,2017:9.30,2018:10.20,
           2019:11.10,2020:12.00,2021:12.32,2022:12.56,2023:13.65,2024:14.42,2025:14.81},
    "CT": {1992:4.27,2000:6.15,2001:6.40,2002:6.70,2003:6.90,2004:7.10,2006:7.40,
           2007:7.65,2009:8.00,2010:8.25,2014:8.70,2015:9.15,2016:9.60,2017:10.10,
           2018:10.10,2019:10.10,2020:12.00,2021:13.00,2022:14.00,2023:15.00,
           2024:16.35,2025:16.35},
    "DE": {1992:4.25,2005:6.15,2007:7.15,2008:7.15,2009:7.25,2014:7.75,2015:8.25,
           2016:8.25,2017:8.25,2018:8.25,2019:9.25,2020:9.25,2021:10.50,2022:11.75,
           2023:13.25,2024:15.00,2025:15.00},
    "DC": {1993:5.25,2001:6.15,2004:6.60,2005:6.60,2006:7.00,2008:7.55,2009:8.25,
           2010:8.25,2011:8.25,2012:8.25,2013:8.25,2014:9.50,2015:10.50,2016:11.50,
           2017:12.50,2018:13.25,2019:14.00,2020:15.00,2021:15.20,2022:16.10,
           2023:17.00,2024:17.50,2025:17.50},
    "FL": {1990:3.80,1991:4.25,2005:6.15,2006:6.40,2007:6.67,2008:6.79,2009:7.21,
           2010:7.25,2021:8.65,2022:10.00,2023:12.00,2024:13.00,2025:14.00},
    "GA": {y: 7.25 for y in range(1990, 2026)},
    "HI": {1990:3.85,1993:5.25,1994:5.25,2002:6.25,2003:6.25,2006:7.25,2007:7.25,
           2015:7.75,2016:8.50,2017:9.25,2018:10.10,2019:10.10,2020:10.10,
           2022:12.00,2023:14.00,2024:14.00,2025:14.00},
    "ID": {y: 7.25 for y in range(1990, 2026)},
    "IL": {1990:3.80,1991:4.25,2004:5.50,2005:6.50,2007:6.50,2008:7.50,2009:8.00,
           2011:8.25,2020:9.25,2021:10.00,2022:12.00,2023:13.00,2024:14.00,2025:15.00},
    "IN": {y: 7.25 for y in range(1990, 2026)},
    "IA": {y: 7.25 for y in range(1990, 2026)},
    "KS": {y: 7.25 for y in range(1990, 2026)},
    "KY": {y: 7.25 for y in range(1990, 2026)},
    "LA": {y: 7.25 for y in range(1990, 2026)},
    "ME": {1990:3.80,1996:4.75,1997:5.15,2002:5.15,2006:6.50,2007:7.00,2008:7.00,
           2009:7.25,2017:9.00,2018:10.00,2019:11.00,2020:12.00,2021:12.15,
           2022:13.80,2023:14.15,2024:14.65,2025:15.00},
    "MD": {1990:3.80,1997:5.00,2007:6.15,2008:6.15,2010:7.25,2014:8.00,2015:8.25,
           2016:8.75,2017:9.25,2018:10.10,2019:10.10,2020:11.00,2021:11.75,
           2022:12.50,2023:13.25,2024:15.00,2025:15.00},
    "MA": {1990:3.85,2000:6.75,2007:7.50,2008:8.00,2015:9.00,2016:10.00,
           2017:11.00,2018:12.00,2019:12.75,2020:13.50,2021:14.25,2022:15.00,
           2025:15.00},
    "MI": {1990:3.85,1997:5.15,2006:6.95,2007:7.40,2008:7.40,2009:7.40,
           2014:8.15,2015:8.50,2016:8.50,2017:8.90,2018:9.25,2019:9.45,
           2020:9.65,2021:9.87,2022:9.87,2023:10.10,2024:10.33,2025:10.56},
    "MN": {1990:3.85,2000:5.15,2005:6.15,2014:8.00,2015:9.00,2016:9.50,
           2017:9.50,2018:9.65,2019:9.86,2020:10.00,2021:10.08,2022:10.33,
           2023:10.59,2024:10.85,2025:11.13},
    "MS": {y: 7.25 for y in range(1990, 2026)},
    "MO": {1990:3.80,2007:6.50,2008:6.65,2009:7.05,2010:7.25,2015:7.65,2016:7.65,
           2017:7.70,2018:7.85,2019:8.60,2020:9.45,2021:10.30,2022:11.15,
           2023:12.00,2024:12.30,2025:13.75},
    "MT": {1990:3.80,2005:5.15,2007:6.15,2008:6.25,2009:6.90,2010:7.25,
           2011:7.35,2012:7.65,2013:7.80,2014:7.90,2015:8.05,2016:8.05,
           2017:8.15,2018:8.30,2019:8.50,2020:8.65,2021:8.75,2022:9.20,
           2023:9.95,2024:10.30,2025:10.55},
    "NE": {1990:3.80,2015:8.00,2016:9.00,2017:9.00,2018:9.00,2019:9.00,
           2020:9.00,2021:9.00,2022:10.50,2023:11.00,2024:12.00,2025:13.50},
    "NV": {1990:4.25,2004:5.15,2007:6.33,2010:7.25,2020:8.00,2021:8.75,
           2022:9.50,2023:10.25,2024:12.00,2025:12.00},
    "NH": {y: 7.25 for y in range(1990, 2026)},
    "NJ": {1990:3.80,2000:5.15,2005:6.15,2006:7.15,2009:7.25,2014:8.25,
           2015:8.38,2018:8.60,2019:10.00,2020:11.00,2021:12.00,2022:13.00,
           2023:14.13,2024:15.49,2025:15.49},
    "NM": {1990:3.80,2009:7.50,2020:9.00,2021:10.50,2022:11.50,2023:12.00,
           2024:12.00,2025:12.00},
    "NY": {1990:3.80,2000:5.15,2005:6.00,2006:6.75,2007:7.15,2009:7.25,
           2014:8.00,2015:8.75,2016:9.00,2017:9.70,2018:10.40,2019:11.10,
           2020:11.80,2021:12.50,2022:13.20,2023:14.20,2024:16.00,2025:16.50},
    "NC": {y: 7.25 for y in range(1990, 2026)},
    "ND": {y: 7.25 for y in range(1990, 2026)},
    "OH": {1990:3.80,2006:6.85,2007:7.00,2008:7.10,2009:7.30,2010:7.30,
           2011:7.40,2012:7.70,2013:7.85,2014:8.10,2015:8.10,2016:8.10,
           2017:8.15,2018:8.30,2019:8.55,2020:8.70,2021:8.80,2022:9.30,
           2023:10.10,2024:10.45,2025:10.70},
    "OK": {y: 7.25 for y in range(1990, 2026)},
    "OR": {1990:3.85,2003:6.90,2004:7.05,2005:7.25,2006:7.50,2007:7.80,
           2008:7.95,2009:8.40,2010:8.40,2011:8.50,2012:8.80,2013:8.95,
           2014:9.10,2015:9.25,2016:9.75,2017:10.25,2018:10.75,2019:11.25,
           2020:12.00,2021:12.75,2022:13.50,2023:13.50,2024:14.70,2025:15.25},
    "PA": {y: 7.25 for y in range(1990, 2026)},
    "RI": {1990:3.85,1997:5.65,1998:5.65,2000:6.15,2001:6.15,2004:6.75,
           2006:7.10,2007:7.40,2009:7.40,2013:7.75,2014:8.00,2015:9.00,
           2016:9.60,2017:9.60,2019:10.50,2020:11.50,2021:12.25,2022:13.00,
           2023:14.00,2024:15.00,2025:15.00},
    "SC": {y: 7.25 for y in range(1990, 2026)},
    "SD": {1990:3.80,2015:8.50,2016:8.55,2017:8.65,2018:8.85,2019:9.10,
           2020:9.30,2021:9.45,2022:9.95,2023:10.80,2024:11.20,2025:11.20},
    "TN": {y: 7.25 for y in range(1990, 2026)},
    "TX": {y: 7.25 for y in range(1990, 2026)},
    "UT": {y: 7.25 for y in range(1990, 2026)},
    "VT": {1992:4.25,2000:5.75,2001:6.25,2002:6.25,2004:7.00,2006:7.25,
           2007:7.53,2008:7.68,2009:8.06,2010:8.06,2011:8.15,2012:8.46,
           2013:8.60,2014:8.73,2015:9.15,2016:9.60,2017:10.00,2018:10.50,
           2019:10.96,2020:11.75,2021:12.55,2022:13.18,2023:13.67,2024:14.01,
           2025:14.01},
    "VA": {1990:3.80,2020:7.25,2021:9.50,2022:11.00,2023:12.00,2024:12.41,
           2025:13.50},
    "WA": {1990:3.85,2000:6.50,2001:6.72,2002:6.90,2003:7.01,2004:7.16,
           2005:7.35,2006:7.63,2007:7.93,2008:8.07,2009:8.55,2010:8.55,
           2011:8.67,2012:9.04,2013:9.19,2014:9.32,2015:9.47,2016:9.53,
           2017:11.00,2018:11.50,2019:12.00,2020:13.50,2021:13.69,2022:14.49,
           2023:15.74,2024:16.28,2025:16.66},
    "WV": {1990:3.80,2015:8.00,2016:8.75,2017:8.75,2018:8.75,2019:8.75,
           2020:8.75,2021:8.75,2022:8.75,2023:8.75,2024:8.75,2025:8.75},
    "WI": {y: 7.25 for y in range(1990, 2026)},
    "WY": {y: 7.25 for y in range(1990, 2026)},
}

def build_state_rows(year_cols: list[int]) -> list[dict]:
    rows = []
    for abbr, hist in STATE_HIST.items():
        # Forward-fill within available range
        filled = {}
        current = None
        for yr in range(min(hist), max(year_cols) + 1):
            if yr in hist:
                current = hist[yr]
            filled[yr] = current

        row = {
            "jurisdiction_id": f"US-{abbr}",
            "jurisdiction": STATE_NAMES.get(abbr, abbr),
            "state": abbr,
            "state_name": STATE_NAMES.get(abbr, abbr),
            "level": "State",
            "source": CITATIONS["state"]["source"],
            "source_url": CITATIONS["state"]["url"],
            "license": CITATIONS["state"]["license"],
            "notes": "State-level minimum wage. Data from Zipperer v1.4.0; forward-filled within series.",
        }
        for yr in year_cols:
            row[f"mw_{yr}"] = filled.get(yr)
        rows.append(row)
    return rows


# ─────────────────────────────────────────────────────────────────────────────
# 3. LOCAL DATA  (UCB)
# ─────────────────────────────────────────────────────────────────────────────
def parse_ucb(path: str) -> list[dict]:
    raw = pd.read_excel(path, sheet_name="Schedules", header=None)
    records = []
    for i, row in raw.iterrows():
        if i < 9 or i > 120:
            continue
        vals = list(row)
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
        name_clean = re.sub(r"\s*\([\d;, a-z\sA-Z]+\)", "", name_raw)
        name_clean = re.sub(r"\d+$", "", name_clean).split("\n")[0].strip().rstrip("*").strip()

        # Build year→wage dict
        wage_by_year = {}
        last_year    = None
        for v in vals[2:]:
            if isinstance(v, datetime):
                last_year = v.year
            elif isinstance(v, (int, float)) and not isinstance(v, bool) and 5 < v < 35:
                if last_year:
                    yr = last_year
                    wage_by_year[yr] = round(float(v), 2)

        # Initial wage (before first dated step) — assign to enactment year
        year_enacted = None
        m = re.search(r"\((\d{4})", name_raw)
        if m:
            year_enacted = int(m.group(1))
        for v in vals[2:5]:
            if isinstance(v, (int, float)) and not isinstance(v, bool) and 5 < v < 35:
                if year_enacted and year_enacted not in wage_by_year:
                    wage_by_year[year_enacted] = round(float(v), 2)
                break

        records.append({
            "name": name_clean,
            "state": state,
            "level": "County" if "county" in name_clean.lower() else "City",
            "wage_by_year": wage_by_year,
            "enact_year": year_enacted,
        })
    return records


def build_local_rows(ucb_path: str, year_cols: list[int]) -> list[dict]:
    records = parse_ucb(ucb_path)
    rows = []
    for idx, rec in enumerate(records, 1):
        # Forward-fill wage within known range
        filled = {}
        if rec["wage_by_year"]:
            min_yr = min(rec["wage_by_year"])
            current = None
            for yr in range(min_yr, max(year_cols) + 1):
                if yr in rec["wage_by_year"]:
                    current = rec["wage_by_year"][yr]
                filled[yr] = current

        row = {
            "jurisdiction_id": f"LOCAL-{rec['state']}-{idx:03d}",
            "jurisdiction": rec["name"],
            "state": rec["state"],
            "state_name": STATE_NAMES.get(rec["state"], rec["state"]),
            "level": rec["level"],
            "source": CITATIONS["local"]["source"],
            "source_url": CITATIONS["local"]["url"],
            "license": CITATIONS["local"]["license"],
            "notes": (
                f"Local ordinance. Enacted ~{rec['enact_year']}. "
                "Wages are scheduled steps from UCB inventory; forward-filled. "
                "Confirm with locality for current rate."
            ),
        }
        for yr in year_cols:
            row[f"mw_{yr}"] = filled.get(yr)
        rows.append(row)
    return rows


# ─────────────────────────────────────────────────────────────────────────────
# 4. COMBINE
# ─────────────────────────────────────────────────────────────────────────────
YEAR_START = 1938
YEAR_END   = 2025
YEAR_COLS  = list(range(YEAR_START, YEAR_END + 1))

def build_combined(ucb_path: str) -> pd.DataFrame:
    print("  Building federal row …")
    federal_row = build_federal_row(YEAR_COLS)

    print("  Building state rows …")
    state_rows = build_state_rows(YEAR_COLS)

    print("  Parsing UCB local data …")
    local_rows = build_local_rows(ucb_path, YEAR_COLS)

    all_rows = [federal_row] + state_rows + local_rows
    df = pd.DataFrame(all_rows)

    # Sort: Federal → State (alpha) → City/County (by state then name)
    level_order = {"Federal": 0, "State": 1, "City": 2, "County": 3}
    df["_sort"] = df["level"].map(level_order)
    df = df.sort_values(["_sort", "state", "jurisdiction"]).drop(columns=["_sort"])
    df = df.reset_index(drop=True)

    print(f"  Combined: {len(df)} rows × {len(df.columns)} columns")
    return df


# ─────────────────────────────────────────────────────────────────────────────
# 5. OUTPUT: CSV
# ─────────────────────────────────────────────────────────────────────────────
def write_csv(df: pd.DataFrame) -> str:
    path = os.path.join(OUT_DIR, "us_minwage_combined.csv")
    df.to_csv(path, index=False)
    print(f"  CSV → {path}")
    return path


# ─────────────────────────────────────────────────────────────────────────────
# 6. OUTPUT: JSON
# ─────────────────────────────────────────────────────────────────────────────
def write_json(df: pd.DataFrame) -> str:
    meta = {
        "title": "U.S. Minimum Wage Combined Dataset",
        "description": (
            "Federal, state, and local (city/county) minimum wages for the United States. "
            "Wide format: one row per jurisdiction, one column per year (mw_YYYY). "
            "Wages are nominal USD per hour."
        ),
        "generated": date.today().isoformat(),
        "year_range": [YEAR_START, YEAR_END],
        "levels": ["Federal", "State", "City", "County"],
        "jurisdiction_count": len(df),
        "citations": CITATIONS,
        "columns": {
            "jurisdiction_id": "Unique identifier (US-FEDERAL, US-{STATE}, LOCAL-{STATE}-{N})",
            "jurisdiction": "Human-readable name",
            "state": "Two-letter state/territory abbreviation",
            "state_name": "Full state name",
            "level": "Federal | State | City | County",
            "source": "Data source name",
            "source_url": "URL for data source",
            "license": "License/terms for the source data",
            "notes": "Caveats and methodology notes",
            "mw_YYYY": "Minimum wage in USD/hr for year YYYY. Null = not yet enacted or data unavailable.",
        },
        "data": [],
    }

    mw_cols = [c for c in df.columns if c.startswith("mw_")]
    meta_cols = [c for c in df.columns if not c.startswith("mw_")]

    for _, row in df.iterrows():
        record = {c: row[c] for c in meta_cols}
        record["wages_by_year"] = {
            col.replace("mw_", ""): (None if pd.isna(row[col]) else row[col])
            for col in mw_cols
        }
        meta["data"].append(record)

    path = os.path.join(OUT_DIR, "us_minwage_combined.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(meta, f, indent=2, ensure_ascii=False)
    size_kb = os.path.getsize(path) / 1024
    print(f"  JSON → {path}  ({size_kb:.0f} KB)")
    return path


# ─────────────────────────────────────────────────────────────────────────────
# 7. OUTPUT: EXCEL  (formatted + citation sheet)
# ─────────────────────────────────────────────────────────────────────────────
# Style helpers
def _hdr_fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)

def _side() -> Side:
    return Side(style="thin", color="CCCCCC")

def _border() -> Border:
    s = _side()
    return Border(left=s, right=s, top=s, bottom=s)

def _font(bold=False, size=10, color="000000", name="Arial"):
    return Font(bold=bold, size=size, color=color, name=name)

LEVEL_COLORS = {
    "Federal": "1F3864",  # dark navy
    "State":   "1F497D",  # medium blue
    "City":    "17375E",  # steel blue
    "County":  "244062",  # slate
}
LEVEL_TEXT = {
    "Federal": "FFFFFF",
    "State":   "FFFFFF",
    "City":    "DDEEFF",
    "County":  "DDEEFF",
}

def write_excel(df: pd.DataFrame) -> str:
    wb = openpyxl.Workbook()

    # ── Sheet 1: DATA ──────────────────────────────────────────────────────
    ws = wb.active
    ws.title = "Data"

    meta_cols = [c for c in df.columns if not c.startswith("mw_")]
    year_cols = sorted([int(c[3:]) for c in df.columns if c.startswith("mw_")])
    all_cols  = meta_cols + [f"mw_{y}" for y in year_cols]

    # Header row
    for col_idx, col_name in enumerate(all_cols, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font      = _font(bold=True, color="FFFFFF", size=9)
        cell.fill      = _hdr_fill("1F3864")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = _border()

    ws.row_dimensions[1].height = 28

    # Data rows
    for row_idx, (_, row) in enumerate(df.iterrows(), 2):
        level = row.get("level", "")
        bg    = LEVEL_COLORS.get(level, "F5F5F5")
        fg    = LEVEL_TEXT.get(level, "000000")
        is_dark = level in ("Federal", "State", "City", "County")

        for col_idx, col_name in enumerate(all_cols, 1):
            val  = row[col_name]
            cell = ws.cell(row=row_idx, column=col_idx)

            if pd.isna(val) if not isinstance(val, str) else False:
                cell.value = None
            else:
                cell.value = val

            cell.border = _border()
            cell.font   = _font(size=9, color=fg if is_dark else "222222")
            cell.fill   = _hdr_fill(bg) if is_dark else PatternFill()

            if col_name.startswith("mw_") and cell.value is not None:
                cell.number_format = '"$"#,##0.00'
                cell.alignment = Alignment(horizontal="right")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    # Freeze top row + first 4 columns
    ws.freeze_panes = "E2"

    # Column widths
    ws.column_dimensions["A"].width = 18  # jurisdiction_id
    ws.column_dimensions["B"].width = 36  # jurisdiction
    ws.column_dimensions["C"].width = 7   # state
    ws.column_dimensions["D"].width = 22  # state_name
    ws.column_dimensions["E"].width = 9   # level
    ws.column_dimensions["F"].width = 58  # source
    ws.column_dimensions["G"].width = 48  # source_url
    ws.column_dimensions["H"].width = 38  # license
    ws.column_dimensions["I"].width = 60  # notes
    for y_idx, _ in enumerate(year_cols, 10):
        ws.column_dimensions[get_column_letter(y_idx)].width = 7

    # ── Sheet 2: CITATIONS ─────────────────────────────────────────────────
    wc = wb.create_sheet("Citations")
    wc.sheet_view.showGridLines = False

    title_font = Font(name="Arial", bold=True, size=14, color="1F3864")
    h2_font    = Font(name="Arial", bold=True, size=11, color="1F3864")
    body_font  = Font(name="Arial", size=10, color="333333")
    url_font   = Font(name="Arial", size=10, color="0563C1", underline="single")
    label_font = Font(name="Arial", bold=True, size=10, color="444444")

    r = 1
    wc.cell(r, 1, "U.S. Minimum Wage Combined Dataset — Citations & Licensing").font = title_font
    r += 1
    wc.cell(r, 1, f"Generated: {date.today().isoformat()}").font = body_font
    r += 2

    for key, info in CITATIONS.items():
        level_label = {"federal": "Federal", "state": "State", "local": "City & County (Local)"}[key]
        wc.cell(r, 1, f"Source: {level_label}").font = h2_font
        r += 1
        for field, label in [
            ("citation",  "Full citation"),
            ("source",    "Dataset name"),
            ("url",       "URL"),
            ("license",   "License"),
            ("retrieved", "Data retrieved"),
        ]:
            wc.cell(r, 1, label + ":").font  = label_font
            cell = wc.cell(r, 2, info.get(field, ""))
            if field == "url":
                cell.font      = url_font
                cell.hyperlink = info["url"]
            else:
                cell.font = body_font
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            wc.row_dimensions[r].height = 30
            r += 1
        r += 1

    r += 1
    wc.cell(r, 1, "Methodology Notes").font = h2_font
    r += 1
    notes = [
        "• Wages are nominal USD per hour.",
        "• Wide format: one row per jurisdiction, columns mw_1938 … mw_2025.",
        "• Federal and state wages are forward-filled within their available historical series.",
        "• Local (city/county) wages represent scheduled step increases from UCB ordinance inventory.",
        "  Wages before the ordinance enactment year are NULL (no local ordinance in effect).",
        "• For localities marked with * in the UCB source, the state wage has met or exceeded",
        "  the local ordinance — confirm with the locality for the operative rate.",
        "• This combined dataset is itself released under CC BY 4.0.",
        "  You must cite all three underlying sources when using or publishing this data.",
    ]
    for note in notes:
        wc.cell(r, 1, note).font = body_font
        wc.row_dimensions[r].height = 18
        r += 1

    wc.column_dimensions["A"].width = 22
    wc.column_dimensions["B"].width = 90

    # ── Sheet 3: README ────────────────────────────────────────────────────
    wr = wb.create_sheet("README")
    wr.sheet_view.showGridLines = False

    readme_lines = [
        ("U.S. Minimum Wage Combined Dataset", title_font),
        ("", body_font),
        ("OVERVIEW", h2_font),
        ("This workbook combines federal, state, and local minimum wage data into a single", body_font),
        ("wide-format table. Each row is one jurisdiction; columns mw_YYYY give the minimum", body_font),
        ("wage in nominal USD/hr for that year.", body_font),
        ("", body_font),
        ("SHEETS", h2_font),
        ("  Data       — The combined dataset (1938–2025)", body_font),
        ("  Citations  — Full citations, licenses, and methodology notes", body_font),
        ("  README     — This guide", body_font),
        ("", body_font),
        ("COLUMNS", h2_font),
        ("  jurisdiction_id  Unique key: US-FEDERAL, US-{STATE}, LOCAL-{STATE}-{N}", body_font),
        ("  jurisdiction     Human-readable name", body_font),
        ("  state            Two-letter abbreviation (or US)", body_font),
        ("  state_name       Full state name", body_font),
        ("  level            Federal | State | City | County", body_font),
        ("  source           Data source name", body_font),
        ("  source_url       URL to original data", body_font),
        ("  license          License / terms of use", body_font),
        ("  notes            Caveats and methodology", body_font),
        ("  mw_YYYY          Minimum wage in USD/hr for year YYYY", body_font),
        ("                   NULL = ordinance not yet enacted or data unavailable", body_font),
        ("", body_font),
        ("DATA SOURCES", h2_font),
        ('  1. Vaghul, K. & Zipperer, B. (2022). "Historical State and Sub-state', body_font),
        ('     Minimum Wages." v1.4.0. CC BY 4.0.', body_font),
        ('     https://github.com/benzipperer/historicalminwage', url_font),
        ("", body_font),
        ('  2. UC Berkeley Center for Labor Research and Education (2025).', body_font),
        ('     "Inventory of Local Minimum Wage Ordinances." Updated Dec 2025.', body_font),
        ('     https://laborcenter.berkeley.edu/minimum-wage-living-wage-resources/', url_font),
        ("", body_font),
        ("CITATION FOR THIS COMBINED DATASET", h2_font),
        ('  [Your Name]. 2025. "U.S. Minimum Wage Combined Dataset."', body_font),
        ('  Compiled from Vaghul & Zipperer (2022) and UC Berkeley Labor Center (2025).', body_font),
        ('  Available at: [YOUR REPO URL]', body_font),
        ('  License: Creative Commons Attribution 4.0 (CC BY 4.0).', body_font),
    ]
    for row_num, (text, font) in enumerate(readme_lines, 1):
        cell = wr.cell(row_num, 1, text)
        cell.font = font
        cell.alignment = Alignment(wrap_text=False)
        wr.row_dimensions[row_num].height = 16

    wr.column_dimensions["A"].width = 100

    path = os.path.join(OUT_DIR, "us_minwage_combined.xlsx")
    wb.save(path)
    size_kb = os.path.getsize(path) / 1024
    print(f"  Excel → {path}  ({size_kb:.0f} KB)")
    return path


# ─────────────────────────────────────────────────────────────────────────────
# 8. MAIN
# ─────────────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="Build combined U.S. minimum wage dataset")
    parser.add_argument("--ucb-path", default=UCB_DEFAULT,
                        help="Path to the UCB Excel file")
    args = parser.parse_args()

    if not os.path.exists(args.ucb_path):
        print(f"ERROR: UCB file not found at {args.ucb_path}")
        print("Place the file in the project root or pass --ucb-path <path>")
        sys.exit(1)

    print("=" * 64)
    print("U.S. Minimum Wage Combined Dataset Builder")
    print("=" * 64)

    print("\n[1/4] Building combined DataFrame …")
    df = build_combined(args.ucb_path)

    print("\n[2/4] Writing CSV …")
    write_csv(df)

    print("\n[3/4] Writing JSON …")
    write_json(df)

    print("\n[4/4] Writing Excel (formatted + citations) …")
    write_excel(df)

    n_federal = len(df[df["level"] == "Federal"])
    n_state   = len(df[df["level"] == "State"])
    n_local   = len(df[df["level"].isin(["City", "County"])])

    print(f"""
╔══════════════════════════════════════════════════════════════╗
  ✓  Output written to ./output/
     us_minwage_combined.csv
     us_minwage_combined.json
     us_minwage_combined.xlsx

  Rows:  {n_federal} Federal  |  {n_state} States  |  {n_local} Cities/Counties
  Years: {YEAR_START}–{YEAR_END}  ({YEAR_END - YEAR_START + 1} columns)

  Remember to update the citation in the README sheet
  with your name and GitHub repo URL before publishing!
╚══════════════════════════════════════════════════════════════╝""")


if __name__ == "__main__":
    main()
