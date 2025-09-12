# app/seed_generate.py
# -----------------------------------------------------------------------------
# Generate a full 45-person crew seed into data/personnel.csv
# Names are realistic anglicized placeholders; all fields filled.
# Officers get AF (Captain has no shift). Warrant/NCOs per your matrix:
# - AF: Officers + Warrant (where applicable)
# - YFM: Chief Petty Officers (Archikelefstes) primarily
# - YF : Senior Petty Officers (Epikelefstes) primarily
# - BYFM/BYF: Petty Officers/Seamen etc.
# Only PRIMARY shift is used by the scheduler; ALT is informational.
# -----------------------------------------------------------------------------

from __future__ import annotations
from pathlib import Path
import pandas as pd

DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)

def row(idx, reg, name, rank, spec, duty, primary, alt):
    """Build one person row with generic 'general' fields filled."""
    address = "Salamis Naval Base"
    phone = f"2106{800000+idx:06d}"
    return {
        "registry_number": reg,
        "name": name,
        "rank": rank,
        "specialty": spec,
        "duty": duty,
        "primary_shift": primary,
        "alt_shift": alt,
        "at_sea_shift": "",
        "height": str(170 + (idx % 15)),   # 170-184
        "weight": str(70 + (idx % 20)),    # 70-89
        "address": address,
        "phone": phone,
        "marital_status": "Married" if idx % 3 == 0 else "Single",
        "children": str(idx % 3),
        "pye_expiration": f"2027-{(idx%12)+1:02d}-{"28" if (idx%12)+1==2 else "30"}",
        "notes": "—",
    }

def main():
    people = []
    i = 0

    # ---------------------- Officers ----------------------
    # 1 Commander (Captain) – no shift
    i+=1; people.append(row(i,"PN-1001","Andreas Konstantinou","Commander","FW","Captain","",""))

    # 1 Lieutenant Commander (Combat) – Executive Officer – AF
    i+=1; people.append(row(i,"PN-1002","Michael Panagiotou","Lieutenant Commander","SEA","Executive Officer","AF",""))

    # 1 Lieutenant Commander (M) – DPO – AF
    i+=1; people.append(row(i,"PN-1003","Spyridon Alexiou","Lieutenant Commander (M)","ENG","DPO","AF",""))

    # 2 Lieutenants: Weapons Director, Operations & EW Director – AF
    i+=1; people.append(row(i,"PN-1004","Dimitrios Lymperopoulos","Lieutenant","FW","Weapons Director","AF",""))
    i+=1; people.append(row(i,"PN-1005","Emmanouil Tsaknakis","Lieutenant","EW/SN","Operations/EW Director","AF",""))

    # 1 Ensign (Combat) – NK Director – AF
    i+=1; people.append(row(i,"PN-1006","Isidoros Gerakaris","Ensign","NK","NK Director","AF",""))

    # 1 Ensign (M) – Second Engineer – AF
    i+=1; people.append(row(i,"PN-1007","Christos Maragkos","Ensign (M)","ENG","Second Engineer","AF",""))

    # 2 Ensigns (E) – AF
    i+=1; people.append(row(i,"PN-1008","Periklis Antonakos","Ensign (E)","EW/RE","Warfare EW Officer","AF",""))
    i+=1; people.append(row(i,"PN-1009","Athanasios Nikiforou","Ensign (E)","FW","FW Officer","AF",""))

    # ---------------------- Warrant (3) ----------------------
    i+=1; people.append(row(i,"PN-2001","Ioannis Dimitriou","Warrant Officer","ENG","Engine Accountant","AF","YFM"))
    i+=1; people.append(row(i,"PN-2002","Konstantinos Raptis","Warrant Officer","ARM","Armaments Officer","AF","YFM"))
    i+=1; people.append(row(i,"PN-2003","Panagiotis Kyriazis","Warrant Officer","EW/DB","EW/DB Accountant","AF","YFM"))

    # ---------------------- Chief Petty Officers (8) – YFM/BYFM ----------------------
    i+=1; people.append(row(i,"PN-3001","Nikolaos Antoniou","Chief Petty Officer","ENG","Engine Accountant","YFM","BYFM"))
    i+=1; people.append(row(i,"PN-3002","Stavros Iliopoulos","Chief Petty Officer","ELEC","ELEC Accountant","YFM","BYFM"))
    i+=1; people.append(row(i,"PN-3003","Vasileios Reppas","Chief Petty Officer","RE","RE Accountant","YFM","BYFM"))
    i+=1; people.append(row(i,"PN-3004","Theodoros Sifakis","Chief Petty Officer","SIG","Signalman","YFM","BYFM"))
    i+=1; people.append(row(i,"PN-3005","Achilleas Armenis","Chief Petty Officer","ARM","Armaments Accountant","YFM","BYFM"))
    i+=1; people.append(row(i,"PN-3006","Klearchos Pierrakos","Chief Petty Officer","FW","FW Accountant","YFM","BYFM"))
    i+=1; people.append(row(i,"PN-3007","Sotirios Deligiannis","Chief Petty Officer","EW/DB","EW/DB Accountant","YFM","BYFM"))
    i+=1; people.append(row(i,"PN-3008","Michalis Diamantis","Chief Petty Officer","ADMIN","General Administrator","YFM","BYFM"))

    # ---------------------- Senior Petty Officers (10) – YF/BYF ------------------------
    i+=1; people.append(row(i,"PN-3101","Panagiotis Tzanetakos","Senior Petty Officer","ENG","Assistant Engine Accountant","YF","BYF"))
    i+=1; people.append(row(i,"PN-3102","Alexandros Lagos","Senior Petty Officer","ENG","Engine Technician","YF","BYF"))
    i+=1; people.append(row(i,"PN-3103","Giorgos Katsantonis","Senior Petty Officer","ENG","Engine Technician","YF","BYF"))
    i+=1; people.append(row(i,"PN-3104","Efthymios Manthos","Senior Petty Officer","ELEC","Assistant ELEC Accountant","YF","BYF"))
    i+=1; people.append(row(i,"PN-3105","Iakovos Triantafyllou","Senior Petty Officer","ELEC","ELEC Technician","YF","BYF"))
    i+=1; people.append(row(i,"PN-3106","Christos Armenis","Senior Petty Officer","ARM","Assistant Armaments","YF","BYF"))
    i+=1; people.append(row(i,"PN-3107","Evangelos Pyrimachos","Senior Petty Officer","FW","Assistant FW","YF","BYF"))
    i+=1; people.append(row(i,"PN-3108","Leonidas Volidas","Senior Petty Officer","FW","Gunnery Chief","YF","BYF"))
    i+=1; people.append(row(i,"PN-3109","Spyridon Markopoulos","Senior Petty Officer","EW/DB","Assistant EW/DB Accountant","YF","BYF"))
    i+=1; people.append(row(i,"PN-3110","Ilias Porfyris","Senior Petty Officer","EW/DB","DB Operator","YF","BYF"))

    # ---------------------- Petty Officers (10) – BYFM/BYF -------------------------
    i+=1; people.append(row(i,"PN-3201","Konstantinos Zervas","Petty Officer","ELEC","ELEC Technician","BYFM","BYF"))
    i+=1; people.append(row(i,"PN-3202","Theofanis Mitrou","Petty Officer","ENG","Engine Technician","BYFM","BYF"))
    i+=1; people.append(row(i,"PN-3203","Stamatis Aslanidis","Petty Officer","EW/AS","Radio Accountant","BYFM","BYF"))
    i+=1; people.append(row(i,"PN-3204","Nikolaos Retsas","Petty Officer","EW/RE","EW/RE Accountant","BYFM","BYF"))
    i+=1; people.append(row(i,"PN-3205","Argyrios Tiliakos","Petty Officer","TEL","TEL Accountant","BYFM","BYF"))
    i+=1; people.append(row(i,"PN-3206","Dimitrios Simainon","Petty Officer","SIG","Assistant Signalman","BYFM","BYF"))
    i+=1; people.append(row(i,"PN-3207","Ioannis Simitikos","Petty Officer","EW/SN","Assistant SN","BYFM","BYF"))
    i+=1; people.append(row(i,"PN-3208","Stefanos Efschimatistos","Petty Officer","COOK","Noise Hygiene","BYFM","BYF"))
    i+=1; people.append(row(i,"PN-3209","Angelos Thalassios","Petty Officer","SEA","Assistant SEA","BYFM","BYF"))
    i+=1; people.append(row(i,"PN-3210","Marios Doryforos","Petty Officer","EW/DB","DB Operator","BYFM","BYF"))

    # ---------------------- Seamen (6) – BYFM/BYF -----------------------------
    i+=1; people.append(row(i,"PN-3301","Anastasios Ventouris","Seaman","ENG","Assistant Engines","BYFM","YF"))
    i+=1; people.append(row(i,"PN-3302","Petros Karras","Seaman","SIG","Assistant Signalman","BYFM","YF"))
    i+=1; people.append(row(i,"PN-3303","Georgios Diacheiristis","Seaman","ADMIN","Assistant Administrator","BYFM","YF"))
    i+=1; people.append(row(i,"PN-3304","Kimon Paschalis","Seaman","EW/DB","Assistant EW/DB","BYFM","YF"))
    i+=1; people.append(row(i,"PN-3305","Efthymios Mavridis","Seaman","EW/SN","Assistant EW/SN","BYFM","YF"))
    i+=1; people.append(row(i,"PN-3306","Klearchos Triantafyllou","Seaman","COOL","Assistant COOL","BYFM","YF"))

    # ---------------------- Sailors (5) – BYF -------------------------------
    i+=1; people.append(row(i,"PN-3401","Apostolos Galanos","Sailor","ENG","NK/Engine","BYF","YF"))
    i+=1; people.append(row(i,"PN-3402","Thomas Zafiris","Sailor","ELEC","Electrician","BYF","YF"))
    i+=1; people.append(row(i,"PN-3403","Athanasios Sideris","Sailor","FW","Assistant FW","BYF","YF"))
    i+=1; people.append(row(i,"PN-3404","Nikolas Kapetanos","Sailor","SIG","NK/Markings","BYF","YF"))
    i+=1; people.append(row(i,"PN-3405","Filippos Diacheiristakis","Sailor","ADMIN","Assistant Supply","BYF","YF"))

    # Assemble DataFrame
    df = pd.DataFrame(people)
    cols = ["registry_number","name","rank","specialty","duty","primary_shift","alt_shift","at_sea_shift",
            "height","weight","address","phone","marital_status","children","pye_expiration","notes"]
    for c in cols:
        if c not in df.columns: df[c] = ""
    df = df[cols]

    out = DATA_DIR / "personnel.csv"
    df.to_csv(out, index=False, encoding="utf-8-sig")
    print(f"✅ Personnel seed created: {out} ({len(df)} records)")

if __name__ == "__main__":
    main()