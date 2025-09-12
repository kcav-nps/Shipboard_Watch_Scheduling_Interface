# app/calendar_service.py
# -----------------------------------------------------------------------------
# Calendar & logs writer utilities:
# - Ship status (in-port / at-sea)
# - Holidays
# - Daily leaves (date ranges expanded to daily rows on disk)
# - Daily cannot / prefer (multi-day inputs)
# - Month setup helpers (patterns)
#
# Robust column harmonization: each CSV is loaded with canonical columns,
# missing columns are created, extras are ignored on write, and write order
# is fixed to avoid "mismatched columns" errors.
# -----------------------------------------------------------------------------

from __future__ import annotations
from pathlib import Path
from datetime import datetime, timedelta
import calendar
import pandas as pd

LOGS_DIR = Path("logs")
LOGS_DIR.mkdir(exist_ok=True)

# --------------------------- Generic helpers --------------------------------

def _canon_df(path: Path, columns: list[str]) -> pd.DataFrame:
    """Load CSV and ensure exactly these columns/order."""
    if path.exists():
        df = pd.read_csv(path, dtype=str).fillna("")
        df = df[[c for c in df.columns if c in columns]]
    else:
        df = pd.DataFrame(columns=columns)
    for c in columns:
        if c not in df.columns:
            df[c] = ""
    df = df[columns]
    return df

def _save_df(path: Path, df: pd.DataFrame, columns: list[str]) -> None:
    """Write CSV with exact columns/order (utf-8-sig)."""
    out = df.copy()
    for c in columns:
        if c not in out.columns:
            out[c] = ""
    out = out[columns].fillna("")
    path.parent.mkdir(parents=True, exist_ok=True)
    out.to_csv(path, index=False, encoding="utf-8-sig")

def _parse_days_list(days_text: str) -> list[int]:
    """Parse '1, 2, 5-7' → [1,2,5,6,7]."""
    if not days_text:
        return []
    out = []
    parts = [p.strip() for p in days_text.split(",")]
    for p in parts:
        if "-" in p:
            a, b = p.split("-", 1)
            try:
                a, b = int(a.strip()), int(b.strip())
                out.extend(list(range(min(a,b), max(a,b)+1)))
            except ValueError:
                continue
        else:
            try:
                out.append(int(p))
            except ValueError:
                continue
    return sorted(set(out))

def _daterange(start_iso: str, end_iso: str):
    """Yield YYYY-MM-DD from start to end inclusive."""
    s = datetime.fromisoformat(start_iso).date()
    e = datetime.fromisoformat(end_iso).date()
    cur = s
    while cur <= e:
        yield cur.isoformat()
        cur = cur + timedelta(days=1)

# ---------------------------- Ship status -----------------------------------

def clear_ship_status(year: int, month: int) -> str:
    path = LOGS_DIR / f"ship_status_{year:04d}_{month:02d}.csv"
    if path.exists():
        path.unlink()
    return "✅ Cleared ship status for the month."

def set_ship_status_bulk(year: int, month: int, days: str, status: str) -> str:
    """Set ship status for given days of a month ('in port' or 'at sea')."""
    path = LOGS_DIR / f"ship_status_{year:04d}_{month:02d}.csv"
    cols = ["date", "status"]
    df = _canon_df(path, cols)

    valid = {"in port", "at sea"}
    if status not in valid:
        return "❌ Error: Invalid ship status."

    dlist = _parse_days_list(days)
    if not dlist:
        return "❌ Error: No days provided."

    # index map
    existing = {r["date"]: i for i, r in enumerate(df.to_dict("records"))}
    for d in dlist:
        try:
            day_iso = datetime(year, month, d).date().isoformat()
        except ValueError:
            continue
        if day_iso in existing:
            idx = existing[day_iso]
            df.at[idx, "status"] = status
        else:
            df.loc[len(df)] = [day_iso, status]
            existing[day_iso] = len(df) - 1

    _save_df(path, df, cols)
    return "✅ Ship status updated for selected days."

# Quick month patterns
def set_month_weekdays_in_port(year: int, month: int) -> str:
    """Mon-Fri = 'in port', Sat/Sun = 'at sea'."""
    clear_ship_status(year, month)
    _, last = calendar.monthrange(year, month)
    in_port = ",".join(str(d) for d in range(1, last+1)
                       if datetime(year, month, d).weekday() < 5)
    at_sea  = ",".join(str(d) for d in range(1, last+1)
                       if datetime(year, month, d).weekday() >= 5)
    msg1 = set_ship_status_bulk(year, month, in_port, "in port")
    msg2 = set_ship_status_bulk(year, month, at_sea, "at sea")
    return msg1 + "\n" + msg2

def set_month_all_in_port(year: int, month: int) -> str:
    clear_ship_status(year, month)
    _, last = calendar.monthrange(year, month)
    days = ",".join(str(d) for d in range(1, last+1))
    return set_ship_status_bulk(year, month, days, "in port")

def set_month_all_at_sea(year: int, month: int) -> str:
    clear_ship_status(year, month)
    _, last = calendar.monthrange(year, month)
    days = ",".join(str(d) for d in range(1, last+1))
    return set_ship_status_bulk(year, month, days, "at sea")

# ------------------------------- Holidays -----------------------------------

def add_holiday(year: int, month: int, day: int, description: str) -> str:
    """Add a holiday; idempotent for the same date/description."""
    path = LOGS_DIR / f"holidays_{year:04d}_{month:02d}.csv"
    cols = ["date", "description"]
    df = _canon_df(path, cols)
    try:
        day_iso = datetime(year, month, day).date().isoformat()
    except ValueError:
        return "❌ Error: Invalid holiday date."
    if ((df["date"] == day_iso) & (df["description"] == description)).any():
        return "ℹ️ Holiday already recorded for this date."
    df.loc[len(df)] = [day_iso, description]
    _save_df(path, df, cols)
    return "✅ Holiday recorded."

# -------------------------- Leaves / Cannot / Prefer ------------------------

def add_leave(registry_number: str, leave_type: str, start_iso: str, end_iso: str, comments: str = "") -> str:
    """Add leave for a person; expands range into daily rows."""
    year, month = int(start_iso[:4]), int(start_iso[5:7])
    path = LOGS_DIR / f"daily_leave_{year:04d}_{month:02d}.csv"
    cols = ["date", "registry_id", "leave_type", "comments"]
    df = _canon_df(path, cols)
    try:
        s = datetime.fromisoformat(start_iso).date()
        e = datetime.fromisoformat(end_iso).date()
        if e < s:
            s, e = e, s
    except Exception:
        return "❌ Error: Invalid leave date range."

    existing = set(tuple(r) for r in df[cols].itertuples(index=False, name=None))
    for day_iso in _daterange(s.isoformat(), e.isoformat()):
        row = (day_iso, registry_number, leave_type, comments)
        if row not in existing:
            df.loc[len(df)] = list(row)
    _save_df(path, df, cols)
    return "✅ The leave was registered and the monthly files were updated."

def add_unavailable(registry_number: str, year: int, month: int, days: str, remarks: str = "") -> str:
    path = LOGS_DIR / f"daily_cannot_{year:04d}_{month:02d}.csv"
    cols = ["date", "registry_id", "comments"]
    df = _canon_df(path, cols)
    dlist = _parse_days_list(days)
    if not dlist:
        return "❌ Error: No days provided."
    existing = set(tuple(r) for r in df[cols].itertuples(index=False, name=None))
    for d in dlist:
        try:
            day_iso = datetime(year, month, d).date().isoformat()
        except ValueError:
            continue
        row = (day_iso, registry_number, remarks)
        if row not in existing:
            df.loc[len(df)] = list(row)
    _save_df(path, df, cols)
    return "✅ The unavailabilities for the selected days were registered."

def add_preference(registry_number: str, year: int, month: int, days: str, remarks: str = "") -> str:
    path = LOGS_DIR / f"daily_prefer_{year:04d}_{month:02d}.csv"
    cols = ["date", "registry_id", "comments"]
    df = _canon_df(path, cols)
    dlist = _parse_days_list(days)
    if not dlist:
        return "❌ Error: No days provided."
    existing = set(tuple(r) for r in df[cols].itertuples(index=False, name=None))
    for d in dlist:
        try:
            day_iso = datetime(year, month, d).date().isoformat()
        except ValueError:
            continue
        row = (day_iso, registry_number, remarks)
        if row not in existing:
            df.loc[len(df)] = list(row)
    _save_df(path, df, cols)
    return "✅ The preferences for the selected days were registered."
