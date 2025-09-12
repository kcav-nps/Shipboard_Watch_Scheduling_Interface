
# migrate_personnel_to_english.py
"""
One-time migration: convert Greek domain values in data/personnel.csv to English,
using i18n_maps/*.csv. Makes a timestamped backup first.
"""

import os, sys, datetime
import pandas as pd

# Assume this script is placed under the project root or the app/ folder.
HERE = os.path.dirname(os.path.abspath(__file__))
APP_DIR = HERE if os.path.basename(HERE) == "app" else os.path.join(HERE, "app")
DATA_DIR = os.path.join(APP_DIR)
MAP_DIR  = os.path.join(APP_DIR, "i18n_maps")
CSV_PATH = os.path.join(APP_DIR, "data", "personnel.csv")

def read_map(name):
    path = os.path.join(MAP_DIR, f"{name}.csv")
    if not os.path.exists(path):
        raise SystemExit(f"Missing mapping file: {path}")
    df = pd.read_csv(path, dtype=str).fillna("")
    m = {}
    for _, r in df.iterrows():
        el = r.get("el","").strip()
        en = r.get("en","").strip()
        if el:
            m[el] = en or el
    return m

def main():
    if not os.path.exists(CSV_PATH):
        raise SystemExit(f"Not found: {CSV_PATH}")

    df = pd.read_csv(CSV_PATH, dtype=str).fillna("")

    rank_map = read_map("ranks")
    spec_map = read_map("specialties")
    duty_map = read_map("duties")
    watch_map = read_map("watch_codes")

    def map_duty_field(s):
        s = (s or "").strip()
        if not s: return ""
        parts = [p.strip() for p in s.replace("|",";").split(";") if p.strip()]
        return "; ".join([duty_map.get(p, p) for p in parts])

    # backup
    ts = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    backup = os.path.join(APP_DIR, "data", f"personnel_BACKUP_EL_{ts}.csv")
    df.to_csv(backup, index=False, encoding="utf-8-sig")

    # migrate columns
    if "rank" in df.columns: df["rank"] = df["rank"].map(lambda v: rank_map.get(v, v))
    if "specialty" in df.columns: df["specialty"] = df["specialty"].map(lambda v: spec_map.get(v, v))
    if "duty" in df.columns: df["duty"] = df["duty"].map(map_duty_field)
    for k in ("primary_shift","alt_shift","at_sea_shift"):
        if k in df.columns: df[k] = df[k].map(lambda v: watch_map.get(v, v))

    out = os.path.join(APP_DIR, "data", "personnel.csv")
    df.to_csv(out, index=False, encoding="utf-8-sig")
    print("Migration complete.")
    print("Backup:", backup)
    print("Updated:", out)

if __name__ == "__main__":
    main()
