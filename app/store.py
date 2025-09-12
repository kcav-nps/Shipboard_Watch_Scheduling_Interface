# store.py
# Centralized data loading and saving utilities.

import pandas as pd
from pathlib import Path

# Assume the script runs from the project root where the 'data' folder is.
DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)

def save_to_csv(records: list[dict], name: str) -> Path:
    """Save a list of dictionaries to a CSV in the data/ directory."""
    path = DATA_DIR / name
    df = pd.DataFrame(records)
    df.to_csv(path, index=False, encoding="utf-8-sig")
    return path

def load_csv(name: str) -> pd.DataFrame:
    """Load a CSV from the data/ directory."""
    path = DATA_DIR / name
    if not path.exists():
        return pd.DataFrame()
    
    # THIS IS THE FIX: Added the required encoding to correctly read the CSV.
    return pd.read_csv(path, dtype=str, encoding="utf-8-sig").fillna("")