# app/scheduling_prep.py
# -----------------------------------------------------------------------------
# Prepares availability data for duty scheduling (in-port only).
# It DOES NOT assign watches; it only collects/filters/organizes pools.
# Robust to column name differences; reads logs from flat ./logs/*.csv files.
# -----------------------------------------------------------------------------

from __future__ import annotations
import os
import pandas as pd
from .constants import RANKS
from .i18n_display_mapping import I18N

# Seniority: lower index = more senior
SENIORITY_ORDER = {rank: i for i, rank in enumerate(RANKS)}

def _seniority_key(rank: str) -> int:
    # THIS IS THE FIX: Removed the unnecessary and incorrect conversion.
    # The code now directly uses the English rank as the key.
    key = str(rank).strip()
    return SENIORITY_ORDER.get(key, 10_000)

# Valid in-port watch types (fixed)
WATCH_TYPES = ["ΑΦ", "ΥΦ", "ΥΦΜ", "ΒΥΦΜ", "ΒΥΦ"]

def _load_csv(path: str) -> pd.DataFrame:
    if os.path.exists(path):
        # Ensure this line has the encoding specified
        return pd.read_csv(path, dtype=str, encoding='utf-8-sig').fillna("")
    return pd.DataFrame()

def _pick_first_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    for c in candidates:
        if c in df.columns:
            return c
    return None

def day_availability(date_str: str) -> dict:
    """
    Prepare availability pools for a specific date (YYYY-MM-DD).
    Returns a dict with ship_status, exclusions, and pools per watch type.
    """
    y, m, _ = date_str.split("-")

    # 1) Ship status (flat files under ./logs)
    status_df = _load_csv(f"logs/ship_status_{y}_{m}.csv")
    if status_df.empty or "date" not in status_df.columns or "status" not in status_df.columns:
        return {"error": "A (correct) ship status file for the month was not found."}
    status_df = status_df.set_index("date")
    if date_str not in status_df.index:
        return {"error": f"There is no ship status entry for {date_str}."}
    ship_status = str(status_df.loc[date_str, "status"]).strip()
    
    if ship_status.lower() not in {"in port", "in-port", "port"}:
        return {
            "date": date_str,
            "ship_status": ship_status,
            "message": "The day is not 'in port' (it is either 'at sea' or not defined)."
        }

    # 2) Personnel registry (harmonize columns)
    people = _load_csv("data/personnel.csv")
    if people.empty:
        return {"error": "The personnel registry is empty."}
    reg_col = _pick_first_col(people, ["registry_number", "registry_id", "service_number"])
    if not reg_col:
        return {"error": "The registry number column (registry_number) is missing."}
    name_col = _pick_first_col(people, ["name", "fullname"])
    if not name_col:
         return {"error": "The full name column (name/fullname) is missing."}

    # 3) Monthly logs (leave / cannot / prefer)
    leave_df  = _load_csv(f"logs/daily_leave_{y}_{m}.csv")
    cannot_df = _load_csv(f"logs/daily_cannot_{y}_{m}.csv")
    prefer_df = _load_csv(f"logs/daily_prefer_{y}_{m}.csv")

    def _ids_on_day(df: pd.DataFrame) -> set[str]:
        if df.empty or "date" not in df.columns:
            return set()
        rc = _pick_first_col(df, ["registry_id", "registry_number", "service_number"])
        if not rc:
            return set()
        return set(df.loc[df["date"] == date_str, rc].astype(str).tolist())

    leave_ids  = _ids_on_day(leave_df)
    cannot_ids = _ids_on_day(cannot_df)
    prefer_ids = _ids_on_day(prefer_df)

    # 4) Build pools by eligibility (rank rules)
    pools = {wt: [] for wt in WATCH_TYPES}
    officer_ranks = {
        "Commander","Commander (M)","Lieutenant Commander","Lieutenant Commander (M)",
        "Lieutenant","Lieutenant (M)","Lieutenant (E)",
        "Ensign","Ensign (M)","Ensign (E)"
    }
    
    for _, row in people.iterrows():
        reg = str(row[reg_col]).strip()
        if not reg or reg in leave_ids or reg in cannot_ids:
            continue

        rk  = str(row["rank"]).strip()
        sp  = str(row["specialty"]).strip()
        nm  = str(row[name_col]).strip()

        if rk in officer_ranks:
            eligible = ["ΑΦ"]
        elif rk == "Warrant Officer":
            eligible = ["ΑΦ","ΥΦ","ΥΦΜ","ΒΥΦΜ","ΒΥΦ"]
        else:
            eligible = ["ΥΦ","ΥΦΜ","ΒΥΦΜ","ΒΥΦ"]

        for wt in eligible:
            pools[wt].append({
                "registry_id": reg,
                "name": nm,
                "rank": rk,
                "specialty": sp,
                "preferred": (reg in prefer_ids)
            })

    # Sort pools by seniority then name
    for w in pools:
        pools[w].sort(key=lambda p: (_seniority_key(p["rank"]), str(p["name"])))

    return {
        "date": date_str,
        "ship_status": ship_status,
        "leave": sorted(leave_ids),
        "cannot": sorted(cannot_ids),
        "prefer": sorted(prefer_ids),
        "pools": pools
    }