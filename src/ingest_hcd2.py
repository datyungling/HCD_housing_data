"""
HCD APR Data Ingester – Table A and Table A2
---------------------------------------------
Pulls Table A and Table A2 from the HCD APR open data portal into a local
DuckDB database. On each run, checks whether the data has changed before
re-ingesting — avoids unnecessary writes if nothing is new.

Usage:
    python ingest_hcd.py              # ingest both tables
    python ingest_hcd.py --force      # force re-ingest even if unchanged

Schedule (optional):
    Add to cron:  0 6 * * 1 /usr/bin/python3 /path/to/ingest_hcd.py
    Or run manually before any analysis session.

Output:
    data/hcd.duckdb        # DuckDB database with tables hcd_table_a, hcd_table_a2
    data/ingest_log.json   # Log of every run with row counts and change status
"""

import json
import hashlib
import argparse
import logging
from datetime import datetime
from pathlib import Path

import requests
import pandas as pd
import duckdb

# ── Config ────────────────────────────────────────────────────────────────────

# Anchored to the script's location — works regardless of which directory
# your terminal is in when you run the script
DATA_DIR = Path(__file__).parent / "data"
DB_PATH  = DATA_DIR / "hcd.duckdb"
LOG_PATH = DATA_DIR / "ingest_log.json"

BASE_URL  = "https://data.ca.gov/api/3/action/datastore_search"
PAGE_SIZE = 1000

TABLES = {
    "hcd_table_a": {
        "resource_id": "c78b769d-cc02-4050-91ef-79ded665b5a8",
        "description": "APR Table A – Permit activity by jurisdiction and affordability level",
    },
    "hcd_table_a2": {
        "resource_id": "fe505d9b-8c36-42ba-ba30-08bc4f34e022",
        "description": "APR Table A2 – Entitlements, building permits, and COs by affordability level",
    },
}

# Columns to cast to numeric for each table.
# Any value that can't convert becomes NULL (via pd.to_numeric errors="coerce").
NUMERIC_COLS = {
    "hcd_table_a": [
        "vlow_income_dr", "vlow_income_ndr",
        "low_income_dr",  "low_income_ndr",
        "mod_income_dr",  "mod_income_ndr",
        "above_mod_income",
        "tot_proposed_units",
        "tot_approved_units",
        "tot_disapproved_units",
        "density_bonus_received",
        "density_bonus_approved",
    ],
    "hcd_table_a2": [
        # Entitlement units
        "vlow_income_dr",    "vlow_income_ndr",
        "low_income_dr",     "low_income_ndr",
        "mod_income_dr",     "mod_income_ndr",
        "above_mod_income",
        "no_entitlements",
        # Building permit units
        "bp_vlow_income_dr", "bp_vlow_income_ndr",
        "bp_low_income_dr",  "bp_low_income_ndr",
        "bp_mod_income_dr",  "bp_mod_income_ndr",
        "bp_above_mod_income",
        "no_building_permits",
        # Certificate of occupancy units
        "co_vlow_income_dr", "co_vlow_income_ndr",
        "co_low_income_dr",  "co_low_income_ndr",
        "co_mod_income_dr",  "co_mod_income_ndr",
        "co_above_mod_income",
        "no_other_forms_of_readiness",
        # Other numeric fields
        "extr_low_income_units",
        "infill_units",
        "no_fa_dr",
        "term_aff_dr",
        "dem_des_units",
        "dem_or_des_units",
        "density_bonus_total",
        "density_bonus_number_other_incentives",
    ],
}

# ── Logging ───────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)

# ── Helpers ───────────────────────────────────────────────────────────────────

def fetch_all_records(resource_id: str) -> list[dict]:
    """Page through the CKAN API and return all records."""
    records = []
    offset  = 0

    while True:
        params   = {"resource_id": resource_id, "limit": PAGE_SIZE, "offset": offset}
        response = requests.get(BASE_URL, params=params, timeout=30)
        response.raise_for_status()

        result = response.json().get("result", {})
        batch  = result.get("records", [])

        if not batch:
            break

        records.extend(batch)
        offset += PAGE_SIZE

        total = result.get("total", "?")
        log.info(f"  Fetched {len(records):,} / {total:,} rows ...")

    return records


def dataframe_hash(df: pd.DataFrame) -> str:
    """Compute a stable hash of a DataFrame to detect changes."""
    return hashlib.md5(
        pd.util.hash_pandas_object(df, index=False).values.tobytes()
    ).hexdigest()


def load_log() -> dict:
    if LOG_PATH.exists():
        with open(LOG_PATH) as f:
            return json.load(f)
    return {}


def save_log(log_data: dict):
    with open(LOG_PATH, "w") as f:
        json.dump(log_data, f, indent=2)


def get_stored_hash(log_data: dict, table_name: str) -> str | None:
    return log_data.get(table_name, {}).get("last_hash")


def clean_dataframe(df: pd.DataFrame, table_name: str) -> pd.DataFrame:
    """
    Standardize the raw API output:
    - Drop internal CKAN metadata column (_id)
    - Lowercase and strip column names
    - Strip whitespace from all string columns
    - Cast known numeric columns to float (NaN for unparseable values)
    """
    # Normalize column names
    df.columns = [c.lower().strip() for c in df.columns]
    df = df.drop(columns=["_id"], errors="ignore")

    # Strip whitespace from all string columns
    str_cols = df.select_dtypes(include="object").columns
    df[str_cols] = df[str_cols].apply(lambda col: col.str.strip())

    # Replace empty strings with NaN so numeric casting works cleanly
    df.replace("", pd.NA, inplace=True)

    # Cast numeric columns — errors="coerce" turns bad values into NaN
    for col in NUMERIC_COLS.get(table_name, []):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
            log.debug(f"  Cast {col} to numeric")

    return df


# ── Core ingestion ────────────────────────────────────────────────────────────

def ingest_table(
    con: duckdb.DuckDBPyConnection,
    table_name: str,
    resource_id: str,
    log_data: dict,
    force: bool = False,
) -> dict:
    """
    Fetch one HCD table, check if it changed, and write to DuckDB if needed.
    Returns a result dict for the log.
    """
    log.info(f"── {table_name} (resource: {resource_id})")

    raw      = fetch_all_records(resource_id)
    df       = clean_dataframe(pd.DataFrame(raw), table_name)
    new_hash = dataframe_hash(df)
    old_hash = get_stored_hash(log_data, table_name)

    changed = (new_hash != old_hash)

    if not changed and not force:
        log.info(f"  No change detected — skipping write.")
        return {
            "status":     "unchanged",
            "rows":       len(df),
            "last_hash":  new_hash,
            "checked_at": datetime.utcnow().isoformat(),
        }

    reason = "forced" if force else "data changed"
    log.info(f"  Writing {len(df):,} rows to DuckDB ({reason}) ...")

    con.execute(f"DROP TABLE IF EXISTS {table_name}")
    con.execute(f"CREATE TABLE {table_name} AS SELECT * FROM df")

    # Log how many values were coerced to NULL during numeric casting
    null_report = {}
    for col in NUMERIC_COLS.get(table_name, []):
        if col in df.columns:
            n_nulls = int(df[col].isna().sum())
            if n_nulls > 0:
                null_report[col] = n_nulls

    if null_report:
        log.warning(f"  Columns with NULL values after cast (may indicate dirty source data):")
        for col, count in null_report.items():
            log.warning(f"    {col}: {count:,} NULLs")

    log.info(f"  Done. Table '{table_name}' updated.")
    return {
        "status":      "updated",
        "rows":        len(df),
        "last_hash":   new_hash,
        "null_report": null_report,
        "updated_at":  datetime.utcnow().isoformat(),
    }


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Ingest HCD APR Table A and A2 into DuckDB.")
    parser.add_argument(
        "--force", action="store_true",
        help="Force re-ingest even if data has not changed"
    )
    args = parser.parse_args()

    DATA_DIR.mkdir(exist_ok=True)

    log_data = load_log()
    con      = duckdb.connect(str(DB_PATH))

    log.info("=" * 55)
    log.info("HCD APR Ingestion Run")
    log.info(f"DB:    {DB_PATH}")
    log.info(f"Force: {args.force}")
    log.info("=" * 55)

    for table_name, meta in TABLES.items():
        result = ingest_table(
            con         = con,
            table_name  = table_name,
            resource_id = meta["resource_id"],
            log_data    = log_data,
            force       = args.force,
        )
        log_data[table_name] = result

    save_log(log_data)
    con.close()

    # ── Summary ──
    log.info("=" * 55)
    for table_name, result in log_data.items():
        if table_name not in TABLES:
            continue
        status = result.get("status", "unknown")
        rows   = result.get("rows", 0)
        nulls  = result.get("null_report", {})
        log.info(f"  {table_name:20s}  {status:10s}  {rows:,} rows")
        if nulls:
            log.info(f"    └─ {len(nulls)} columns had NULL values after cast")
    log.info("=" * 55)
    log.info(f"Log saved to: {LOG_PATH}")


if __name__ == "__main__":
    main()