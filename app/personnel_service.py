# personnel_service.py
# Handles personnel registry operations: add/update/list and simple validations.
# All user-facing messages are in English; comments are in English.

from __future__ import annotations
from .i18n_display_mapping import I18N
from .constants import RANKS, SPECIALTIES, DUTIES
from pathlib import Path
from typing import Dict, List, Tuple
import pandas as pd

from .store import DATA_DIR, save_to_csv, load_csv
from .constants import RANKS, SPECIALTIES, DUTIES

# --- Seniority order: lower index = more senior (RANKS is already high→low) ---
SENIORITY_ORDER = {rank: i for i, rank in enumerate(RANKS)}

def _seniority_key(rank: str) -> int:
    """Return integer priority for rank; smaller means more senior."""
    # THIS IS THE FIX: Directly use the English rank as the key.
    return SENIORITY_ORDER.get(str(rank).strip(), 10_000)

PERSONNEL_FILE = "personnel.csv"

# ----- Helpers ---------------------------------------------------------------

def _is_officer(rank: str) -> bool:
    """
    Officers are from Commander down to Ensign (including 'M'/'E' variants).
    """
    officer_set = {
        "Commander","Commander (M)",
        "Lieutenant Commander","Lieutenant Commander (M)",
        "Lieutenant","Lieutenant (M)","Lieutenant (E)",
        "Ensign","Ensign (M)","Ensign (E)"
    }
    return rank in officer_set

def _validate_person_payload(p: Dict) -> Tuple[bool, str]:
    """
    Validate minimal fields and dropdown membership.
    """
    required = ["rank", "name", "specialty"]
    for r in required:
        if not str(p.get(r, "")).strip():
            return False, f"The field «{r}» is required."

    # THIS IS THE FIX: Validate directly against the English lists.
    if str(p['rank']) not in RANKS:
        return False, "Invalid rank."
    if str(p['specialty']) not in SPECIALTIES:
        return False, "Invalid specialty."
    duty = str(p.get('duty','')).strip()
    if duty and duty not in DUTIES:
        return False, f"Invalid duty: {duty}"
    for k in ("primary_shift", "alt_shift"):
        v = str(p.get(k, '')).strip()
        if v and v not in {"ΑΦ","ΥΦΜ","ΒΥΦΜ","ΥΦ","ΒΥΦ"}:
            return False, f"Invalid in-port watch in field «{k}»."
    return True, "OK"

def _personnel_df() -> pd.DataFrame:
    """
    Load current personnel registry as a DataFrame with stable columns.
    """
    df = load_csv(PERSONNEL_FILE)
    if df.empty:
        df = pd.DataFrame(columns=[
            "name","rank","specialty","duty",
            "primary_shift","alt_shift","at_sea_shift",
            "height","weight","registry_number","address","phone",
            "marital_status","children","pye_expiration","notes"
        ])
    for c in df.columns: # Ensure all expected columns exist
        if c not in df.columns:
            df[c] = ""
    return df

def _save_personnel_df(df: pd.DataFrame) -> None:
    save_to_csv(df.to_dict(orient="records"), PERSONNEL_FILE)

# ----- Public API ------------------------------------------------------------

def add_or_update_person(payload: Dict) -> str:
    """
    Upsert a person by 'registry_number'.
    """
    ok, msg = _validate_person_payload(payload)
    if not ok:
        return f"❌ Error: {msg}"

    df = _personnel_df()

    # Build clean row
    row = {
        "name": str(payload["name"]).strip(),
        "rank": str(payload.get('rank','')).strip(),
        "specialty": str(payload.get('specialty','')).strip(),
        "duty": str(payload.get('duty','')).strip(),
        "primary_shift": str(payload.get("primary_shift","")).strip(),
        "alt_shift": str(payload.get("alt_shift","")).strip(),
        "at_sea_shift": str(payload.get('at_sea_shift','')).strip(),
        "height": str(payload.get("height","")).strip(),
        "weight": str(payload.get("weight","")).strip(),
        "registry_number": str(payload.get("registry_number","")).strip(),
        "address": str(payload.get("address","")).strip(),
        "phone": str(payload.get("phone","")).strip(),
        "marital_status": str(payload.get("marital_status","")).strip(),
        "children": str(payload.get("children","")).strip(),
        "pye_expiration": str(payload.get("pye_expiration","")).strip(),
        "notes": str(payload.get("notes","")).strip(),
    }
    
    if _is_officer(row["rank"]):
        row["primary_shift"], row["alt_shift"] = "", ""

    mask = (df["registry_number"].astype(str) == row["registry_number"])
    if mask.any():
        df.loc[mask, list(row.keys())] = list(row.values())
        msg = "✅ Personnel update completed."
    else:
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        msg = "✅ Personnel registration completed."
    
    _save_personnel_df(df)
    return msg

def list_personnel() -> pd.DataFrame:
    """
    Return the whole personnel table sorted by seniority (rank) and then by name.
    """
    df = _personnel_df()
    if df.empty:
        return df
    df["__k"] = df["rank"].apply(_seniority_key)
    df = df.sort_values(["__k", "name"], ascending=[True, True]).drop(columns="__k")
    return df
