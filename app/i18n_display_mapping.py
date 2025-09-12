
# app/i18n_display_mapping.py
# Lightweight i18n layer that maps Greek <-> English for domain values.
# Reads mapping CSVs from ./i18n_maps/*.csv. If a mapping is missing, it falls back to identity.

import os
import pandas as pd

DATA_DIR = os.path.dirname(os.path.abspath(__file__))

def _read_map(filename, base_values):
    path = os.path.join(DATA_DIR, "i18n_maps", filename)
    # start with identity mapping so unmapped values just pass through
    m = {v: v for v in base_values}
    if os.path.exists(path):
        df = pd.read_csv(path, dtype=str).fillna("")
        for _, r in df.iterrows():
            el = r.get("el", "")
            en = r.get("en", "")
            if el:
                m[el] = en or el
    return m

def _invert_map(m):
    # When multiple Greek -> same English, prefer first; collisions won't crash.
    inv = {}
    for k, v in m.items():
        inv.setdefault(v, k)
    return inv

class I18N:
    def __init__(self, ranks, specialties, duties, watches):
        self.rank_map = _read_map("ranks.csv", ranks)
        self.spec_map = _read_map("specialties.csv", specialties)
        self.duty_map = _read_map("duties.csv", duties)
        self.watch_map = _read_map("watch_codes.csv", watches)
        self._rank_inv = _invert_map(self.rank_map)
        self._spec_inv = _invert_map(self.spec_map)
        self._duty_inv = _invert_map(self.duty_map)
        self._watch_inv = _invert_map(self.watch_map)

    def to_display(self, cat, val):
        if cat == "rank": return self.rank_map.get(val, val)
        if cat == "specialty": return self.spec_map.get(val, val)
        if cat == "duty": return self.duty_map.get(val, val)
        if cat == "watch": return self.watch_map.get(val, val)
        return val

    def to_storage(self, cat, val):
        if cat == "rank": return self._rank_inv.get(val, val)
        if cat == "specialty": return self._spec_inv.get(val, val)
        if cat == "duty": return self._duty_inv.get(val, val)
        if cat == "watch": return self._watch_inv.get(val, val)
        return val

    def seq_display(self, cat, seq):
        return [self.to_display(cat, x) for x in seq]

    def duties_display_to_storage_field(self, s):
        """Convert a '; '-separated display string to storage Greek 'duty' field."""
        s = (s or "").strip()
        if not s: return ""
        parts = [p.strip() for p in s.replace("|",";").split(";") if p.strip()]
        el_parts = [self.to_storage("duty", p) for p in parts]
        return "; ".join(el_parts)

    def duties_storage_to_display_list(self, s):
        s = (s or "").strip()
        if not s: return []
        parts = [p.strip() for p in s.replace("|",";").split(";") if p.strip()]
        return [self.to_display("duty", p) for p in parts]
