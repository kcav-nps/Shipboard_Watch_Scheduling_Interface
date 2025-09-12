# app/scheduler_rules.py
# -----------------------------------------------------------------------------
# Helper rules & utilities shared by schedulers.
# Comments in English; runtime messages/headers in English.
# -----------------------------------------------------------------------------

from __future__ import annotations
from .i18n_display_mapping import I18N
from .constants import RANKS, SPECIALTIES, DUTIES
from datetime import datetime, timedelta
from pathlib import Path
import pandas as pd

LOGS_DIR = Path("logs")

# --------------------------- Rank constraints -------------------------------

# Max allowed duties per month by rank (global across all watch types)
MAX_PER_MONTH = {
    # Officers
    "Commander": 4,
    "Commander (M)": 4,
    "Lieutenant Commander": 4,
    "Lieutenant Commander (M)": 4,
    "Lieutenant": 4,
    "Lieutenant (M)": 4,
    "Lieutenant (E)": 4,
    "Ensign": 4,
    "Ensign (M)": 4,
    "Ensign (E)": 4,
    # Warrant & NCOs
    "Warrant Officer": 4,
    "Chief Petty Officer": 4,
    "Senior Petty Officer": 4,
    "Petty Officer": 5,
    "Seaman": 6,
    "Sailor": 15,
}

# Duties with special behavior
DUTY_NEVER = {"Captain"}                 # Never scheduled in any watch
DUTY_WEEKDAY_ONLY = {"Executive Officer", "DPO"}  # For AF: Mon-Fri only

# --------------------------- Calendar helpers -------------------------------

def _load_csv(path: Path) -> pd.DataFrame:
    if path.exists():
        return pd.read_csv(path, dtype=str).fillna("")
    return pd.DataFrame()

def is_holiday(date_iso: str) -> bool:
    y, m, _ = date_iso.split("-")
    hol = _load_csv(LOGS_DIR / f"holidays_{int(y):04d}_{int(m):02d}.csv")
    if hol.empty or "date" not in hol.columns:
        return False
    return date_iso in set(hol["date"].astype(str))

def weekday_name_gr(date_iso: str) -> str:
    # Return uppercase English weekday name
    d = datetime.fromisoformat(date_iso).date()
    names = ["MONDAY","TUESDAY","WEDNESDAY","THURSDAY","FRIDAY","SATURDAY","SUNDAY"]
    return names[d.weekday()]

def is_weekend(date_iso: str) -> bool:
    d = datetime.fromisoformat(date_iso).date()
    return d.weekday() >= 5  # 5=Saturday, 6=Sunday

def is_weekday(date_iso: str) -> bool:
    return not is_weekend(date_iso)

def two_day_gap_ok(prev_assigned_dates: set[str], date_iso: str) -> bool:
    """
    Enforce at least two empty days between duties:
    If assigned on X, next allowed is ≥ X+3 (difference in days ≥ 3).
    """
    if not prev_assigned_dates:
        return True
    d = datetime.fromisoformat(date_iso).date()
    for s in prev_assigned_dates:
        dd = datetime.fromisoformat(s).date()
        if abs((d - dd).days) < 3:
            return False
    return True

# Bilingual helper
try:
    from .i18n_display_mapping import I18N
    from .constants import RANKS
    _I18N_RULES = I18N(RANKS, [], [], ["", "AF","YF","YFM","BYFM","BYF"])
    def rank_key(v: str) -> str:
        return _I18N_RULES.to_storage('rank', str(v))
except Exception:
    def rank_key(v: str) -> str:
        return str(v)
