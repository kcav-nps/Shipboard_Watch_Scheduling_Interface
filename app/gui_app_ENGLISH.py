# app/gui_app.py
# -----------------------------------------------------------------------------
# Tkinter GUI – OA3801
# Tab 1: PersonnelManager  – Personnel Management (empty on start; load/replace/save)
# Tab 2: LeaveManager      – ONLY Leave Management (ranges), with filters/sorting,
#                            context menu (macOS-friendly), delete/edit,
#                            and Excel export (one sheet per month: MONTH_YEAR).
# Tab 3: ShiftsManager      – Shifts (in port): ship status, holidays, unavailability/preferences,
#                            calculation and 3 Excel exports.
#
# Run:  PYTHONPATH=. python app/gui_app.py
# -----------------------------------------------------------------------------

import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from calendar import monthrange

from app.export_service import export_personnel_excel
from app.calendar_service import add_leave
from app.calendar_service import (
    add_unavailable, add_preference, set_ship_status_bulk, add_holiday,
    set_month_all_in_port
)

from app.scheduler_in_port import make_month_schedule_all, export_month_schedule_all

# -----------------------------------------------------------------------------
# Paths from project root
# -----------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.normpath(os.path.join(BASE_DIR, ".."))
DATA_DIR = os.path.join(ROOT_DIR, "data")
LOGS_DIR = os.path.join(ROOT_DIR, "logs")
PERSONNEL_CSV = os.path.join(DATA_DIR, "personnel.csv")

# Columns persisted in data/personnel.csv
CSV_COLS = [
    "registry_number","name","rank","specialty","duty",
    "primary_shift","alt_shift","at_sea_shift",
    "height","weight","address","phone",
    "marital_status","children","pye_expiration","notes"
]

# Domain choices
RANKS = [
    "Commander","Commander (M)","Lieutenant Commander","Lieutenant Commander (M)",
    "Lieutenant","Lieutenant (M)","Lieutenant (E)",
    "Ensign","Ensign (M)","Ensign (E)",
    "Warrant Officer","Chief Petty Officer","Senior Petty Officer",
    "Petty Officer","Seaman","Sailor"
]

SPECIALTIES = [
    "Combat","Engineer",
    "FW","EW/DB","EW/RE","EW/AS","EW/SN","ENG","ELEC","COOL","ARM",
    "ADMIN","SEA","COOK","RE","SIG","TEL","NK"
]

DUTIES = [
    "Captain","Executive Officer","DPO","Operations Director","EW Director",
    "NK Director","Weapons Director","Second Engineer","FW Officer",
    "Warfare EW Officer","SN Officer","Armaments Officer",
    "Engine Room Officer","General Administrator","Signalman",
    "Assistant Signalman","Engine Accountant","Assistant Engine Accountant",
    "ELEC Accountant","Assistant ELEC Accountant","EW/DB Accountant","Assistant EW/DB Accountant",
    "EW/RE Accountant","EW/AS Accountant","EW/SN Accountant","CPM Accountant",
    "TEL Accountant","COOL Accountant","Gunnery Chief","Cook",
    "Officer's Steward","Assistant Administrator","NK/Engine",
    "NK/Markings","Engine Technician","ELEC Technician","Noise Hygiene"
]

WATCH_CHOICES = ["", "AF", "YF", "YFM", "BYFM", "BYF"]
YEARS  = [str(y) for y in range(2000, 2061)]
MONTHS = [f"{m:02d}" for m in range(1, 13)]
MINIMAL_COLS_GR = ["Rank","Specialty","First Name","Last Name"]

LEAVE_TYPES = [
    "Regular","AMD","Child Rearing Leave",
    "Verbal Leave","Parental Leave","Marriage Leave","Maternity Leave"
]

# ------------------------------- IO helpers ----------------------------------
def ensure_personnel_csv():
    """Ensure data folder and an empty personnel.csv exist."""
    os.makedirs(DATA_DIR, exist_ok=True)
    if not os.path.exists(PERSONNEL_CSV):
        pd.DataFrame(columns=CSV_COLS).to_csv(PERSONNEL_CSV, index=False, encoding="utf-8-sig")

def load_personnel_df() -> pd.DataFrame:
    """Load personnel from disk and guarantee all columns exist."""
    ensure_personnel_csv()
    df = pd.read_csv(PERSONNEL_CSV, dtype=str, encoding='utf-8-sig').fillna("")
    for c in CSV_COLS:
        if c not in df.columns:
            df[c] = ""
    return df[CSV_COLS]

def save_personnel_df(df: pd.DataFrame):
    """Persist personnel to disk (UTF-8 with BOM for Excel)."""
    os.makedirs(DATA_DIR, exist_ok=True)
    out = df.copy()[CSV_COLS].fillna("")
    out.to_csv(PERSONNEL_CSV, index=False, encoding="utf-8-sig")

def import_minimal_excel_replace(path: str) -> pd.DataFrame:
    """
    Import a minimal Excel with columns (English): Rank, Specialty, First Name, Last Name.
    Build a fresh CSV with generated registry_number PN-IM-XXXX.
    """
    df = pd.read_excel(path, dtype=str).fillna("")
    missing = [c for c in MINIMAL_COLS_GR if c not in df.columns]
    if missing:
        raise ValueError("Missing columns: " + ", ".join(missing))
    out_rows = []
    for i, r in df.iterrows():
        rank = r["Rank"].strip()
        spec = r["Specialty"].strip()
        first = r["First Name"].strip()
        last  = r["Last Name"].strip()
        name  = f"{first} {last}".strip()
        reg   = f"PN-IM-{i+1:04d}"
        out_rows.append({
            "registry_number": reg, "name": name, "rank": rank, "specialty": spec,
            "duty": "", "primary_shift": "", "alt_shift": "", "at_sea_shift": "",
            "height": "", "weight": "", "address": "", "phone": "",
            "marital_status": "", "children": "", "pye_expiration": "", "notes": ""
        })
    return pd.DataFrame(out_rows, columns=CSV_COLS)

# ============================ TAB 1: PERSONNEL ===============================
class PersonnelManager(ttk.Frame):
    """
    Personnel tab.
    - Starts with an EMPTY DataFrame (in-memory).
    - You can import a "simple" Excel (ONLY in memory — DOES NOT write to CSV).
    - Edit/New entry and Save -> updates ONLY the specific record in the CSV.
    - Delete -> updates the CSV.
    - On each save/delete, it notifies the other tabs via callback.
    """
    def __init__(self, master, on_personnel_changed=None):
        super().__init__(master, padding=12)
        self.on_personnel_changed = on_personnel_changed

        # in-memory dataframe (starts empty)
        self.df = pd.DataFrame(columns=CSV_COLS)

        ensure_personnel_csv()
        self._build_ui()
        self._refresh_table()
        self._log("The table starts empty. Load Excel (IN MEMORY ONLY) or manually enter + Save to write to CSV.")

    # ------------------------------- UI --------------------------------------
    def _build_ui(self):
        # Vertical PanedWindow: top (table+form) / bottom (log)
        paned = ttk.PanedWindow(self, orient="vertical")
        paned.pack(fill="both", expand=True)

        top = ttk.Frame(paned); paned.add(top, weight=4)
        bottom = ttk.LabelFrame(paned, text="Logs"); paned.add(bottom, weight=1)

        # Toolbar
        toolbar = ttk.Frame(top)
        toolbar.pack(fill="x", pady=(0,8))

        ttk.Button(
            toolbar,
            text="Load from Excel",
            command=self._on_import_minimal_replace
        ).pack(side="left")

        ttk.Button(
            toolbar,
            text="Export Personnel (Excel)",
            command=self._on_export_excel
        ).pack(side="left", padx=12)

        # Split (left table / right form)
        main = ttk.Frame(top); main.pack(fill="both", expand=True)
        left = ttk.Frame(main); left.pack(side="left", fill="both", expand=True, padx=(0,8))
        right = ttk.LabelFrame(main, text="Personnel Details"); right.pack(side="left", fill="y")

        # right form: 3 columns (labels / fields / inline save button)
        right.grid_columnconfigure(0, minsize=140)
        right.grid_columnconfigure(1, weight=1)
        right.grid_columnconfigure(2, minsize=160)

        cols = ("idx","rank","specialty","name","registry_number")
        heads = ["S/N","Rank","Specialty","Full Name","Registry Number"]
        widths = [60, 150, 150, 240, 150]
        self.tree = ttk.Treeview(left, columns=cols, show="headings", height=18)
        for c, h, w in zip(cols, heads, widths):
            self.tree.heading(c, text=h)
            self.tree.column(c, width=w, anchor="w")

        yscroll = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(left, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        left.grid_rowconfigure(0, weight=1); left.grid_columnconfigure(0, weight=1)
        self.tree.bind("<<TreeviewSelect>>", self._on_select_row)

        tbl_btns = ttk.Frame(left); tbl_btns.grid(row=2, column=0, sticky="w", pady=(6,0))
        ttk.Button(tbl_btns, text="Refresh (read CSV)", command=self._on_reload).pack(side="left")
        ttk.Button(tbl_btns, text="Delete", command=self._on_delete).pack(side="left", padx=6)

        # ---- Form
        self.inputs = {}
        def add_row(r, label, key, kind="entry", values=None, width=28):
            ttk.Label(right, text=label + ":").grid(row=r, column=0, sticky="e", padx=6, pady=4)
            if kind == "entry":
                var = tk.StringVar()
                ent = ttk.Entry(right, textvariable=var, width=width)
                ent.grid(row=r, column=1, sticky="w", padx=6, pady=4)
                self.inputs[key] = var
            elif kind == "combo":
                var = tk.StringVar()
                cb = ttk.Combobox(right, textvariable=var, values=values or [], width=width-2, state="readonly")
                cb.grid(row=r, column=1, sticky="w", padx=6, pady=4)
                self.inputs[key] = var

        r = 0
        add_row(r,"Registry No.","registry_number"); r += 1
        add_row(r,"Full Name","name"); r += 1
        add_row(r,"Rank","rank",kind="combo",values=RANKS); r += 1
        add_row(r,"Specialty","specialty",kind="combo",values=SPECIALTIES); r += 1

        ttk.Label(right,text="Duty (select):").grid(row=r,column=0,sticky="e",padx=6,pady=(12,4))
        self.var_duty_choice = tk.StringVar()
        ttk.Combobox(
            right, textvariable=self.var_duty_choice, values=DUTIES, width=26, state="readonly"
        ).grid(row=r,column=1,sticky="w",padx=6,pady=(12,4))
        r += 1

        duty_btns = ttk.Frame(right); duty_btns.grid(row=r,column=0,columnspan=2,sticky="w",padx=6,pady=(0,4))
        ttk.Button(duty_btns,text="Add",command=self._on_duty_add).pack(side="left")
        ttk.Button(duty_btns,text="Remove",command=self._on_duty_remove).pack(side="left",padx=8)
        r += 1

        ttk.Label(right,text="Duties (selected):").grid(row=r,column=0,sticky="ne",padx=6,pady=4)
        self.list_duties = tk.Listbox(right,height=5,selectmode="extended",exportselection=False,width=28)
        self.list_duties.grid(row=r,column=1,sticky="w",padx=6,pady=4)
        r += 1

        # --- PRIMARY SHIFT + inline "Save now"
        row_primary = r
        add_row(row_primary, "Primary Shift", "primary_shift", kind="combo", values=WATCH_CHOICES)
        ttk.Button(
            right, text="Save now", command=self._on_save
        ).grid(row=row_primary, column=2, padx=6, pady=4, sticky="w")
        r += 1

        add_row(r,"Alternate","alt_shift",kind="combo",values=WATCH_CHOICES); r += 1
        add_row(r,"At Sea Shift","at_sea_shift",kind="combo",values=[""]+WATCH_CHOICES[1:]); r += 1
        add_row(r,"Height (cm)","height"); r += 1
        add_row(r,"Weight (kg)","weight"); r += 1
        add_row(r,"Address","address"); r += 1
        add_row(r,"Phone","phone"); r += 1
        add_row(r,"Marital Status","marital_status",kind="combo",values=["Married","Single"]); r += 1
        add_row(r,"Children","children"); r += 1

        # PYE date dropdowns
        ttk.Label(right,text="PYE (Year/Month/Day):").grid(row=r,column=0,sticky="e",padx=6,pady=4)
        frm_pye = ttk.Frame(right); frm_pye.grid(row=r,column=1,sticky="w",padx=6,pady=4)
        self.var_pye_year  = tk.StringVar(value="")
        self.var_pye_month = tk.StringVar(value="")
        self.var_pye_day   = tk.StringVar(value="")
        self.cb_pye_year  = ttk.Combobox(frm_pye,textvariable=self.var_pye_year, values=YEARS,  width=6, state="readonly")
        self.cb_pye_month = ttk.Combobox(frm_pye,textvariable=self.var_pye_month,values=MONTHS,width=4, state="readonly")
        self.cb_pye_day   = ttk.Combobox(frm_pye,textvariable=self.var_pye_day,  values=[f"{d:02d}" for d in range(1,32)], width=4, state="readonly")
        self.cb_pye_year.pack(side="left"); ttk.Label(frm_pye,text="-").pack(side="left",padx=3)
        self.cb_pye_month.pack(side="left"); ttk.Label(frm_pye,text="-").pack(side="left",padx=3)
        self.cb_pye_day.pack(side="left")
        r += 1

        def on_pye_change(*_):
            y = self.var_pye_year.get(); m = self.var_pye_month.get()
            if y and m:
                try:
                    last = monthrange(int(y), int(m))[1]
                    self.cb_pye_day["values"] = [f"{dd:02d}" for dd in range(1,last+1)]
                    if self.var_pye_day.get() and int(self.var_pye_day.get())>last:
                        self.var_pye_day.set(f"{last:02d}")
                except Exception:
                    pass
        self.cb_pye_year.bind("<<ComboboxSelected>>", on_pye_change)
        self.cb_pye_month.bind("<<ComboboxSelected>>", on_pye_change)

        add_row(r,"Other information","notes",width=28); r += 1

        # Save button at the bottom
        frm_btns = ttk.Frame(right)
        frm_btns.grid(row=r, column=0, columnspan=3, pady=(12, 6), sticky="w")
        ttk.Button(frm_btns, text="Save now", command=self._on_save).pack(side="left", padx=8)

        # Log (always visible)
        self.txt_log = tk.Text(bottom, height=8, wrap="word")
        self.txt_log.pack(fill="both", expand=True, padx=6, pady=6)

    # ---------------------------- helpers ------------------------------------
    def _log(self, msg: str):
        self.txt_log.insert("end", msg + "\n"); self.txt_log.see("end")

    def _refresh_table(self):
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        if not self.df.empty:
            view = self.df.sort_values("registry_number")
            for idx, (_, r) in enumerate(view.iterrows(), start=1):
                self.tree.insert(
                    "", "end", iid=r["registry_number"],
                    values=(idx, r["rank"], r["specialty"], r["name"], r["registry_number"])
                )

    def _duties_from_field(self, s: str):
        s = (s or "").strip()
        return [] if not s else [p.strip() for p in s.replace("|",";").split(";") if p.strip()]

    def _duties_to_field(self) -> str:
        return "; ".join(self.list_duties.get(0,"end"))

    def _collect_pye_iso(self) -> str:
        y,m,d = self.var_pye_year.get(), self.var_pye_month.get(), self.var_pye_day.get()
        return f"{y}-{m}-{d}" if (y and m and d) else ""

    def _load_pye_into_dropdowns(self, iso: str):
        iso = (iso or "").strip()
        if len(iso)==10 and iso[4]=="-" and iso[7]=="-":
            y,m,d = iso.split("-")
            if y in YEARS: self.var_pye_year.set(y)
            if m in MONTHS: self.var_pye_month.set(m)
            try:
                last = monthrange(int(y), int(m))[1]
                self.cb_pye_day["values"] = [f"{dd:02d}" for dd in range(1,last+1)]
            except Exception:
                pass
            if len(d)==2: self.var_pye_day.set(d)
        else:
            self.var_pye_year.set(""); self.var_pye_month.set(""); self.var_pye_day.set("")

    # ---------------------------- actions ------------------------------------
    def _on_import_minimal_replace(self):
        """
        Load simple Excel -> LOADS ONLY IN MEMORY (DOES NOT write to CSV).
        The CSV will only be updated when you 'Save' a record.
        """
        path = filedialog.askopenfilename(
            title="Load personnel Excel (in memory)",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not path:
            return
        try:
            new_df = import_minimal_excel_replace(path)
        except Exception as e:
            messagebox.showerror("Import Error", f"{e}")
            return

        # load into memory, DO NOT write to CSV
        self.df = new_df.copy()
        self._refresh_table()
        self._log(f"✅ Loaded (IN MEMORY ONLY) from: {os.path.basename(path)}")
        messagebox.showinfo(
            "Done",
            "The personnel has been loaded into memory.\n"
            "The personnel.csv file has NOT been changed.\n"
            "Make changes/Save to write individual records to the CSV."
        )

    def _on_duty_add(self):
        val = self.var_duty_choice.get().strip()
        if not val:
            return
        existing = set(self.list_duties.get(0,"end"))
        if val in existing:
            self._log(f"The duty «{val}» already exists.")
            return
        self.list_duties.insert("end", val)
        self._log(f"Added duty: {val}")

    def _on_duty_remove(self):
        sel = list(self.list_duties.curselection())
        if not sel:
            self._log("Select duty(s) to remove.")
            return
        for ix in reversed(sel):
            self._log(f"Removed duty: {self.list_duties.get(ix)}")
            self.list_duties.delete(ix)

    def _on_select_row(self, _=None):
        sel = self.tree.selection()
        if not sel:
            return
        rid = sel[0]
        row = self.df[self.df["registry_number"]==rid]
        if row.empty:
            return
        rec = row.iloc[0].to_dict()
        for k in CSV_COLS:
            if k in self.inputs and k not in ("duty","pye_expiration"):
                self.inputs[k].set(str(rec.get(k,"")))
        self.list_duties.delete(0,"end")
        for d in self._duties_from_field(rec.get("duty","")):
            self.list_duties.insert("end", d)
        self._load_pye_into_dropdowns(rec.get("pye_expiration",""))
        self._log(f"Loading: {rid}")

    def _on_new(self):
        for k in self.inputs:
            self.inputs[k].set("")
        self.list_duties.delete(0,"end")
        self.var_pye_year.set(""); self.var_pye_month.set(""); self.var_pye_day.set("")
        self._log("New record: fill in the fields and press «Save».")

    def _notify_change(self):
        if callable(self.on_personnel_changed):
            try:
                self.on_personnel_changed()
            except Exception:
                pass

    def _on_save(self):
        # collect values from form
        rec = {k: (self.inputs[k].get().strip() if k in self.inputs else "") for k in CSV_COLS}
        rec["duty"] = self._duties_to_field()
        rec["pye_expiration"] = self._collect_pye_iso()

        rid = rec.get("registry_number","").strip()

        # basic validations
        if not rid:
            messagebox.showerror("Error","Registry Number is required.")
            return
        if rec["rank"] and rec["rank"] not in RANKS:
            messagebox.showerror("Error","Invalid Rank.")
            return

        # ---- Safeguard for Primary/Alternate Shift ----
        if not rec["primary_shift"] and not rec["alt_shift"]:
            proceed = messagebox.askyesno(
                "Confirmation",
                (
                    f"The person '{rec.get('name','')}' (Registry No. {rid}) has neither a Primary nor an Alternate shift.\n"
                    "If you continue, the scheduler might not produce correct shifts.\n\n"
                    "Do you want to continue saving?"
                )
            )
            if not proceed:
                return

        # write ONLY this record to personnel.csv
        disk_df = load_personnel_df()
        idx_list = disk_df.index[disk_df["registry_number"]==rid].tolist()
        if idx_list:
            disk_df.loc[idx_list[0], CSV_COLS] = [rec[c] for c in CSV_COLS]
            self._log(f"Updated (on disk): {rid}")
        else:
            disk_df.loc[len(disk_df)] = [rec[c] for c in CSV_COLS]
            self._log(f"Added (on disk): {rid}")

        save_personnel_df(disk_df)

        # refresh in-memory table
        self.df = disk_df.copy()
        self._refresh_table()
        self._notify_change()
        messagebox.showinfo("Done","Saved to personnel.csv")

    def _on_delete(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Warning","Select a record to delete.")
            return
        rid = sel[0]
        if messagebox.askyesno("Confirmation", f"Delete {rid}?"):
            disk_df = load_personnel_df()
            disk_df = disk_df[disk_df["registry_number"]!=rid].reset_index(drop=True)
            save_personnel_df(disk_df)
            self.df = disk_df.copy()
            self._refresh_table()
            self._notify_change()
            self._log(f"Deleted: {rid}")

    def _on_reload(self):
        # load from CSV -> replaces the in-memory view (e.g., after manual changes)
        self.df = load_personnel_df()
        self._refresh_table()
        self._log("Loading from CSV.")
        self._notify_change()

    def _on_export_excel(self):
        out_path = os.path.join(DATA_DIR,"Personnel.xlsx")
        os.makedirs(DATA_DIR, exist_ok=True)
        out = export_personnel_excel(out_path)
        self._log(f"✅ Personnel export: {out}")
        messagebox.showinfo("Success", f"Personnel exported to:\n{out}")

# ======================== TAB 2: LEAVE MANAGEMENT ONLY =======================
class LeaveManager(ttk.Frame):
    """Tab 2: Leave Management with view, filters, sorting,
    delete/edit, and export (one sheet per month)."""

    def __init__(self, master):
        super().__init__(master, padding=12)
        os.makedirs(LOGS_DIR, exist_ok=True)
        self.people = []
        self.sort_asc = True
        self.filter_year = ""   # "" = all
        self.filter_month = ""  # "" = all
        self._edit_buffer = None
        self._build_ui()
        self.reload_people()

    # ---------- public API (called by Tab1 when personnel changes)
    def reload_people(self):
        df = load_personnel_df()
        people = []
        for _, r in df.iterrows():
            rid = str(r["registry_number"]).strip()
            nm  = str(r["name"]).strip()
            rk  = str(r["rank"]).strip()
            sp  = str(r["specialty"]).strip()
            people.append({"rid": rid, "label": f"{rid} | {rk} ({sp}) | {nm}"})
        self.people = people
        self.cb_person["values"] = [p["label"] for p in people]
        if people and not self.var_person.get():
            self.cb_person.current(0)
        self._log(f"Personnel available: {len(people)}")
        self._refresh_table_for_current_person()

    # ---------- UI ----------
    def _build_ui(self):
        # Header: person + buttons
        hdr = ttk.Frame(self); hdr.pack(fill="x", pady=(0,8))
        ttk.Label(hdr, text="Person:").pack(side="left")
        self.var_person = tk.StringVar()
        self.cb_person = ttk.Combobox(hdr, textvariable=self.var_person, values=[], width=60, state="readonly")
        self.cb_person.pack(side="left", padx=(4,10))
        self.cb_person.bind("<<ComboboxSelected>>", lambda e: self._refresh_table_for_current_person())
        ttk.Button(hdr, text="Refresh Personnel", command=self.reload_people).pack(side="left", padx=(0,10))
        ttk.Button(hdr, text="Export Excel (Leaves)", command=self._on_export_leaves_excel).pack(side="left")

        # Filters + sorting
        filters = ttk.Frame(self); filters.pack(fill="x", pady=(0,6))
        ttk.Label(filters, text="Filter: Month").pack(side="left")
        self.var_f_m = tk.StringVar(value="All")
        self.var_f_y = tk.StringVar(value="All")
        months_ui = ["All"] + [f"{m:02d}" for m in range(1,13)]
        years_ui  = ["All"] + YEARS
        ttk.Combobox(filters, textvariable=self.var_f_m, values=months_ui, width=6, state="readonly").pack(side="left", padx=(4,12))
        ttk.Label(filters, text="Year").pack(side="left")
        ttk.Combobox(filters, textvariable=self.var_f_y, values=years_ui, width=8, state="readonly").pack(side="left", padx=(4,12))
        ttk.Button(filters, text="Apply Filters", command=self._apply_filters).pack(side="left", padx=(0,8))
        ttk.Button(filters, text="Sort (A↕Z)", command=self._toggle_sort).pack(side="left")
        ttk.Button(filters, text="Delete Selected", command=self._on_delete_selected).pack(side="left", padx=(12,0))

        # Entry form (Day-Month-Year)
        grp_leave = ttk.LabelFrame(self, text="Enter Leave (range From–To)")
        grp_leave.pack(fill="x", pady=(6,6))

        row1 = ttk.Frame(grp_leave); row1.pack(fill="x", padx=6, pady=4)
        ttk.Label(row1, text="Type:").pack(side="left")
        self.LEAVE_TYPES = [
            "Regular","AMD","Child Rearing Leave",
            "Verbal Leave","Parental Leave","Marriage Leave","Maternity Leave"
        ]
        self.var_leave_type = tk.StringVar(value=self.LEAVE_TYPES[0])
        ttk.Combobox(row1, textvariable=self.var_leave_type, values=self.LEAVE_TYPES,
                         width=28, state="readonly").pack(side="left", padx=(4,18))

        DAYS_31   = [f"{d:02d}" for d in range(1, 32)]
        MONTHS_12 = [f"{m:02d}" for m in range(1, 13)]

        def date_picker(parent, label, vday, vmon, vyear):
            fr = ttk.Frame(parent); fr.pack(side="left", padx=(0,20))
            ttk.Label(fr, text=label).pack(side="left")
            cb_d = ttk.Combobox(fr, textvariable=vday,  values=DAYS_31,  width=4, state="readonly");  cb_d.pack(side="left", padx=(4,2))
            cb_m = ttk.Combobox(fr, textvariable=vmon,  values=MONTHS_12, width=4, state="readonly"); cb_m.pack(side="left", padx=(2,2))
            cb_y = ttk.Combobox(fr, textvariable=vyear, values=YEARS,     width=6, state="readonly"); cb_y.pack(side="left", padx=(2,0))
            def _sync_days(*_):
                y, m = vyear.get(), vmon.get()
                if y and m:
                    try:
                        last = monthrange(int(y), int(m))[1]
                        cb_d["values"] = [f"{dd:02d}" for dd in range(1, last+1)]
                        if vday.get():
                            try:
                                if int(vday.get()) > last:
                                    vday.set(f"{last:02d}")
                            except Exception:
                                pass
                    except Exception:
                        cb_d["values"] = DAYS_31
            cb_m.bind("<<ComboboxSelected>>", _sync_days)
            cb_y.bind("<<ComboboxSelected>>", _sync_days)

        # From / To pickers
        self.var_s_d = tk.StringVar(value=""); self.var_s_m = tk.StringVar(value=""); self.var_s_y = tk.StringVar(value="")
        self.var_e_d = tk.StringVar(value=""); self.var_e_m = tk.StringVar(value=""); self.var_e_y = tk.StringVar(value="")
        date_picker(row1, "From:", self.var_s_d, self.var_s_m, self.var_s_y)
        date_picker(row1, "To:",   self.var_e_d, self.var_e_m, self.var_e_y)

        row2 = ttk.Frame(grp_leave); row2.pack(fill="x", padx=6, pady=4)
        ttk.Label(row2, text="Comment:").pack(side="left")
        self.var_leave_note = tk.StringVar(value="")
        ttk.Entry(row2, textvariable=self.var_leave_note, width=80).pack(side="left", padx=(4,8))
        ttk.Button(row2, text="Submit Leave", command=self._on_add_or_replace_leave).pack(side="left")

        # Table (one row per range)
        grp_tbl = ttk.LabelFrame(self, text="Leaves of selected person (grouped into ranges)")
        grp_tbl.pack(fill="both", expand=True, pady=(6,0))

        cols   = ("start","end","type","note")
        heads  = ["From","To","Type","Comment"]
        widths = [110, 110, 260, 580]
        self.tree = ttk.Treeview(grp_tbl, columns=cols, show="headings", height=12)
        for c, h, w in zip(cols, heads, widths):
            self.tree.heading(c, text=h)
            self.tree.column(c, width=w, anchor="w")
        yscroll = ttk.Scrollbar(grp_tbl, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(grp_tbl, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        grp_tbl.grid_rowconfigure(0, weight=1); grp_tbl.grid_columnconfigure(0, weight=1)

        # Context menu (macOS-friendly)
        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="Load for editing", command=self._on_load_for_edit)
        self.menu.add_command(label="Delete", command=self._on_delete_selected)
        # Bind variants for macOS and others
        self.tree.bind("<Button-3>", self._open_ctx_menu)         # right-click (Win/Linux)
        self.tree.bind("<ButtonRelease-3>", self._open_ctx_menu)
        self.tree.bind("<Button-2>", self._open_ctx_menu)         # two-finger/right on mac
        self.tree.bind("<ButtonRelease-2>", self._open_ctx_menu)
        self.tree.bind("<Control-Button-1>", self._open_ctx_menu) # Ctrl+Click (mac)

        # Log
        grp_log = ttk.LabelFrame(self, text="Logs")
        grp_log.pack(fill="both", expand=False, pady=(8,0))
        self.txt_log = tk.Text(grp_log, height=8, wrap="word")
        self.txt_log.pack(fill="both", expand=True, padx=6, pady=6)

    # ---------- Data helpers ----------
    def _log(self, msg: str):
        self.txt_log.insert("end", msg + "\n"); self.txt_log.see("end")

    def _selected_registry(self) -> str:
        label = self.var_person.get().strip()
        for p in self.people:
            if p["label"] == label:
                return p["rid"]
        return ""

    def _open_ctx_menu(self, event):
        """Open context menu selecting the row under cursor (macOS-safe)."""
        iid = self.tree.identify_row(event.y)
        if iid:
            self.tree.selection_set(iid)
            self.tree.focus(iid)
        try:
            self.menu.tk_popup(event.x_root, event.y_root)
        finally:
            try:
                self.menu.grab_release()
            except Exception:
                pass
        return "break"

    # ---------- IO: read leaves (supports daily + ranged formats) ----------
    def _read_all_leaves_rows(self):
        rows = []
        if not os.path.exists(LOGS_DIR):
            return rows
        for filename in os.listdir(LOGS_DIR):
            if not (filename.startswith("daily_leave_") and filename.endswith(".csv")):
                continue
            full = os.path.join(LOGS_DIR, filename)
            try:
                df = pd.read_csv(full, dtype=str, encoding='utf-8-sig').fillna("")
            except Exception:
                continue
            if "start_date" in df.columns:
                # New format: start/end range
                if "registry_number" not in df.columns and "registry_id" in df.columns:
                    df["registry_number"] = df["registry_id"]
                for _, r in df.iterrows():
                    rows.append({
                        "start_date": r.get("start_date",""),
                        "end_date":   r.get("end_date","") or r.get("start_date",""),
                        "registry_number": r.get("registry_number",""),
                        "leave_type": r.get("leave_type",""),
                        "comments":   r.get("comments","")
                    })
            elif "date" in df.columns:
                # Old daily format: collapse to ranges
                if "registry_number" not in df.columns and "registry_id" in df.columns:
                    df["registry_number"] = df["registry_id"]
                daily = []
                for _, r in df.iterrows():
                    daily.append({
                        "date": r.get("date",""),
                        "registry_number": r.get("registry_number",""),
                        "leave_type": r.get("leave_type",""),
                        "comments": r.get("comments","")
                    })
                rows.extend(self._collapse_daily_to_ranges(daily))
        return rows

    def _collapse_daily_to_ranges(self, daily_rows):
        """Group consecutive daily leaves per person/type/comment into ranges."""
        from datetime import datetime, timedelta
        out = []
        keyfunc = lambda r: (r["registry_number"], r["leave_type"], r.get("comments",""))
        daily_rows = [r for r in daily_rows if r.get("date")]
        daily_rows.sort(key=lambda r: (r["registry_number"], r["leave_type"], r.get("comments",""), r["date"]))
        i = 0
        while i < len(daily_rows):
            r = daily_rows[i]
            rid, ltype, comm = keyfunc(r)
            start = end = r["date"]
            j = i + 1
            while j < len(daily_rows):
                r2 = daily_rows[j]
                if keyfunc(r2) != (rid, ltype, comm):
                    break
                try:
                    d1 = datetime.strptime(end, "%Y-%m-%d").date()
                    d2 = datetime.strptime(r2["date"], "%Y-%m-%d").date()
                    if (d1 + timedelta(days=1)) == d2:
                        end = r2["date"]; j += 1; continue
                except Exception:
                    pass
                break
            out.append({
                "start_date": start, "end_date": end,
                "registry_number": rid, "leave_type": ltype, "comments": comm
            })
            i = j
        return out

    def _refresh_table_for_current_person(self):
        """Apply filters/sort and display one row per (start–end) range."""
        for iid in self.tree.get_children():
            self.tree.delete(iid)

        rid = self._selected_registry()
        if not rid:
            return

        rows = [r for r in self._read_all_leaves_rows() if r["registry_number"] == rid]

        # filters
        m = self.filter_month
        y = self.filter_year
        if m or y:
            def keep(r):
                src = r.get("start_date","") or r.get("end_date","")
                if len(src) < 7: return False
                ry, rm = src[:4], src[5:7]
                if y and ry != y: return False
                if m and rm != m: return False
                return True
            rows = [r for r in rows if keep(r)]

        # sorting
        rows.sort(key=lambda r: r.get("start_date",""), reverse=not self.sort_asc)

        # show
        for r in rows:
            self.tree.insert("", "end", values=(r.get("start_date",""), r.get("end_date",""),
                                                 r.get("leave_type",""), r.get("comments","")))
        self._log(f"Displaying leaves for {rid}: {len(rows)} records.")

    # ---------- validations ----------
    def _dates_ok(self, start, end) -> bool:
        from datetime import datetime
        try:
            s = datetime.strptime(start, "%Y-%m-%d").date()
            e = datetime.strptime(end, "%Y-%m-%d").date()
            return s <= e
        except Exception:
            return False

    def _overlaps_existing(self, rid, start, end) -> bool:
        from datetime import datetime
        s_new = datetime.strptime(start, "%Y-%m-%d").date()
        e_new = datetime.strptime(end, "%Y-%m-%d").date()
        rows = [r for r in self._read_all_leaves_rows() if r["registry_number"] == rid]
        if self._edit_buffer:
            rows = [r for r in rows if not (
                r["start_date"] == self._edit_buffer["start_date"] and
                r["end_date"]   == self._edit_buffer["end_date"]   and
                r["leave_type"] == self._edit_buffer["leave_type"] and
                (r.get("comments","") == self._edit_buffer.get("comments",""))
            )]
        for r in rows:
            try:
                s = datetime.strptime(r["start_date"], "%Y-%m-%d").date()
                e = datetime.strptime(r["end_date"], "%Y-%m-%d").date()
                if s <= e_new and e >= s_new:
                    return True
            except Exception:
                continue
        return False

    # ---------- actions ----------
    def _on_add_or_replace_leave(self):
        rid = self._selected_registry()
        if not rid:
            messagebox.showerror("Error", "Select a person."); return

        lt = self.var_leave_type.get().strip()
        sd, sm, sy = self.var_s_d.get(), self.var_s_m.get(), self.var_s_y.get()
        ed, em, ey = self.var_e_d.get(), self.var_e_m.get(), self.var_e_y.get()
        if not (sd and sm and sy and ed and em and ey):
            messagebox.showerror("Error", "Please fill in the complete From–To dates."); return

        start = f"{sy}-{sm}-{sd}"
        end   = f"{ey}-{em}-{ed}"
        note  = self.var_leave_note.get().strip()

        if not self._dates_ok(start, end):
            messagebox.showerror("Error", "The 'From' date must be ≤ 'To' date."); return
        if self._overlaps_existing(rid, start, end):
            messagebox.showerror("Error", "The date range overlaps with an existing leave."); return

        # if editing, delete old row first
        if self._edit_buffer:
            self._delete_row_exact(self._edit_buffer)

        msg = add_leave(rid, lt, start, end, note or "")
        self._log(msg)
        messagebox.showinfo("Done", msg)
        self._edit_buffer = None
        self._refresh_table_for_current_person()

    def _on_load_for_edit(self):
        sel = self.tree.selection()
        if not sel: return
        vals = self.tree.item(sel[0], "values")
        start, end, ltype, note = vals
        self._edit_buffer = {
            "registry_number": self._selected_registry(),
            "start_date": start, "end_date": end,
            "leave_type": ltype, "comments": note
        }
        self.var_leave_type.set(ltype)
        self.var_s_y.set(start[0:4]); self.var_s_m.set(start[5:7]); self.var_s_d.set(start[8:10])
        self.var_e_y.set(end[0:4]);   self.var_e_m.set(end[5:7]);   self.var_e_d.set(end[8:10])
        self.var_leave_note.set(note or "")
        self._log("Loaded for editing. Make changes and press «Submit Leave» to replace.")

    def _on_delete_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Warning", "Please select a leave from the table first.")
            return

        start, end, ltype, note = self.tree.item(sel[0], "values")
        rid = self._selected_registry()
        if not rid:
            messagebox.showerror("Error", "Could not find the selected person.")
            return

        if not messagebox.askyesno("Confirmation",
                                       f"Permanently delete leave:\n"
                                       f"{start} – {end}\nType: {ltype}\nComment: {note or '—'}"):
            return

        rec = {
            "registry_number": rid,
            "start_date": start,
            "end_date": end,
            "leave_type": ltype,
            "comments": note or ""
        }

        ok, affected_files = self._delete_row_exact(rec)

        if ok:
            self._log("✅ The leave was deleted.")
            if affected_files:
                for p in affected_files:
                    self._log(f"Updated file: {p}")
            self._refresh_table_for_current_person()
            messagebox.showinfo("Done", "The leave was deleted.")
            # (Optional) Re-export the total Excel after each deletion:
            # self._on_export_leaves_excel()
        else:
            self._log("❌ The leave was not found in the files for deletion.")
            messagebox.showwarning("Warning", "No record found to delete.")

    def _delete_row_exact(self, rec) -> tuple[bool, list]:
        """
        Deletes the leave from the CSVs in logs/.
        Supports:
          - NEW format (range): columns start_date/end_date
          - OLD format (daily): column date (deletes all days within the range)
        Returns: (success: bool, affected_files: list[str])
        """
        from datetime import datetime, timedelta

        def daterange(d1, d2):
            cur = d1
            while cur <= d2:
                yield cur
                cur += timedelta(days=1)

        rid = rec["registry_number"]
        start = rec["start_date"]
        end = rec["end_date"]
        ltype = rec["leave_type"]
        note = rec.get("comments", "")

        try:
            d_start = datetime.strptime(start, "%Y-%m-%d").date()
            d_end   = datetime.strptime(end,   "%Y-%m-%d").date()
        except Exception:
            return False, []

        success = False
        changed_paths = []

        if not os.path.exists(LOGS_DIR):
            return False, []

        # Iterate through all daily_leave_*.csv files
        for filename in os.listdir(LOGS_DIR):
            if not (filename.startswith("daily_leave_") and filename.endswith(".csv")):
                continue
            full = os.path.join(LOGS_DIR, filename)
            try:
                df = pd.read_csv(full, dtype=str).fillna("")
            except Exception:
                continue

            # Normalize registry column
            if "registry_number" not in df.columns and "registry_id" in df.columns:
                df["registry_number"] = df["registry_id"]

            orig_len = len(df)

            if "start_date" in df.columns:
                # -------- NEW format (ranges) --------
                # Remove ONLY the exact record (full match)
                mask = (
                    (df["registry_number"] == rid) &
                    (df["start_date"] == start) &
                    (df["end_date"]   == end) &
                    (df["leave_type"] == ltype) &
                    (df.get("comments", "") == note)
                )
                if mask.any():
                    df = df[~mask].reset_index(drop=True)

            elif "date" in df.columns:
                # -------- OLD format (daily rows) --------
                # Remove all days of the range that match type/comment
                to_remove_dates = {d.isoformat() for d in daterange(d_start, d_end)}
                mask = (
                    (df["registry_number"] == rid) &
                    (df["leave_type"] == ltype) &
                    (df.get("comments", "") == note) &
                    (df["date"].isin(to_remove_dates))
                )
                if mask.any():
                    df = df[~mask].reset_index(drop=True)

            # If changed, write back (or delete if empty - optional)
            if len(df) != orig_len:
                success = True
                changed_paths.append(full)
                if df.empty:
                    # Optionally: delete the file that became empty
                    try:
                        os.remove(full)
                    except Exception:
                        pass
                else:
                    df.to_csv(full, index=False, encoding="utf-8-sig")

        return success, changed_paths

    def _apply_filters(self):
        self.filter_month = "" if self.var_f_m.get() == "All" else self.var_f_m.get()
        self.filter_year  = "" if self.var_f_y.get() == "All" else self.var_f_y.get()
        self._refresh_table_for_current_person()

    def _toggle_sort(self):
        self.sort_asc = not self.sort_asc
        self._refresh_table_for_current_person()

    def _on_export_leaves_excel(self):
        """Create data/Personnel Leaves.xlsx with a sheet per month (MONTH_YEAR)."""
        rows = self._read_all_leaves_rows()
        if not rows:
            messagebox.showwarning("Warning", "No leaves found in logs/."); return

        # Join with personnel for display
        pers = load_personnel_df().copy()
        pers = pers[["registry_number","name","rank","specialty"]]
        df = pd.DataFrame(rows)
        df = df.merge(pers, on="registry_number", how="left")

        MONTH_EN = {
            "01":"JANUARY","02":"FEBRUARY","03":"MARCH","04":"APRIL",
            "05":"MAY","06":"JUNE","07":"JULY","08":"AUGUST",
            "09":"SEPTEMBER","10":"OCTOBER","11":"NOVEMBER","12":"DECEMBER"
        }

        out_path = os.path.join(DATA_DIR, "Personnel Leaves.xlsx")
        os.makedirs(DATA_DIR, exist_ok=True)

        def ym(s):
            s = (s or "")
            return (s[0:4], s[5:7]) if len(s) >= 7 else ("", "")
        year_mon = df["start_date"].apply(ym)
        df["_year"] = year_mon.apply(lambda t: t[0])
        df["_mon"]  = year_mon.apply(lambda t: t[1])
        missing = df["_year"] == ""
        if missing.any():
            ym2 = df.loc[missing, "end_date"].apply(ym)
            df.loc[missing, "_year"] = ym2.apply(lambda t: t[0])
            df.loc[missing, "_mon"]  = ym2.apply(lambda t: t[1])

        with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
            for (y, m), g in df.groupby(["_year","_mon"]):
                if not y or not m:
                    continue
                sheet = f"{MONTH_EN.get(m,m)}_{y}"
                out = g.copy()
                out = out[[
                    "registry_number","rank","specialty","name",
                    "leave_type","start_date","end_date","comments"
                ]]
                out.columns = ["Registry No.","Rank","Specialty","Full Name",
                               "Type","From","To","Comment"]
                out.to_excel(writer, index=False, sheet_name=sheet)

        self._log(f"✅ Export: {out_path}")
        messagebox.showinfo("Success", f"Leaves were exported to:\n{out_path}")

# =============================== TAB 3: SHIFTS ===============================
class ShiftsManager(ttk.Frame):
    """
    Tab 3: Shifts (in port)
    - Select Year/Month
    - Initialization: all days 'in port'
    - Define 'at sea' days
    - Declare Holidays
    - Declare Unavailabilities / Preferences (with Submit/Delete buttons)
    - Calculate shifts + Preview
    - Export 3 Excels
    """

    def __init__(self, master):
        super().__init__(master, padding=12)
        os.makedirs(DATA_DIR, exist_ok=True)
        os.makedirs(LOGS_DIR, exist_ok=True)

        # ---- selections (Year/Month)
        today = pd.Timestamp.today()
        self.var_year  = tk.StringVar(value=str(today.year))
        self.var_month = tk.StringVar(value=f"{today.month:02d}")

        # ---- preview cache/result
        self.people = []      # [{'rid':..., 'label':...}, ...]
        self.result = None      # scheduler result cache for preview/exports

        # ---- state for Unavailabilities / Preferences
        self.var_person_label = tk.StringVar(value="")
        self.var_is_pref      = tk.BooleanVar(value=False)  # False=Unavailabilities, True=Preferences
        self.var_days_unav    = tk.StringVar(value="")
        self.var_unav_note    = tk.StringVar(value="")

        # ---- state for ship status & holidays
        self.var_days_ship    = tk.StringVar(value="")
        self.var_holiday_days = tk.StringVar(value="")
        self.var_holiday_title= tk.StringVar(value="Holiday")

        self._build_ui()
        self.reload_people()

    def _is_officer(self, rank: str) -> bool:
        """Return True if rank is considered an officer."""
        officer_ranks = {
            "Commander","Commander (M)","Lieutenant Commander","Lieutenant Commander (M)",
            "Lieutenant","Lieutenant (M)","Lieutenant (E)",
            "Ensign","Ensign (M)","Ensign (E)"
        }
        return str(rank).strip() in officer_ranks

    def _display_name(self, rank: str, specialty: str, name: str) -> str:
        """Formats a name for display in the preview table."""
        rank = str(rank).strip()
        specialty = str(specialty).strip()
        name = str(name).strip()
        if self._is_officer(rank):
            return f"{rank} | {name} HN"
        spec = f" ({specialty})" if specialty else ""
        return f"{rank}{spec} | {name}"
    
    # ---------------- UI ----------------
    def _build_ui(self):
        # Header: Year / Month
        hdr = ttk.LabelFrame(self, text="Shifts Month")
        hdr.pack(fill="x", pady=(0,8))
        ttk.Label(hdr, text="Year:").pack(side="left")
        ttk.Combobox(hdr, textvariable=self.var_year, values=YEARS, width=8, state="readonly").pack(side="left", padx=(4,12))
        ttk.Label(hdr, text="Month:").pack(side="left")
        ttk.Combobox(hdr, textvariable=self.var_month, values=[f"{m:02d}" for m in range(1,13)], width=6, state="readonly").pack(side="left", padx=(4,12))

        ttk.Button(hdr, text="Initialize (All In Port)", command=self._on_init_month_ormo).pack(side="left", padx=(0,8))
        ttk.Button(hdr, text="Calculate Shifts", command=self._on_compute).pack(side="left", padx=(0,8))
        ttk.Button(hdr, text="Export All (3 Excel files)", command=self._on_export_all_excels).pack(side="left")

        # Ship status (EN PLO)
        frm_ship = ttk.LabelFrame(self, text="«At Sea» days for the selected month")
        frm_ship.pack(fill="x", pady=(6,6))
        ttk.Label(frm_ship, text="Days (e.g. 3, 5-7, 12):").pack(side="left", padx=(6,4))
        ttk.Entry(frm_ship, textvariable=self.var_days_ship, width=30).pack(side="left", padx=(0,8))
        ttk.Button(frm_ship, text="Set AT SEA", command=self._on_ship_en_plo).pack(side="left")

        # Holidays
        frm_hol = ttk.LabelFrame(self, text="Holidays for the selected month")
        frm_hol.pack(fill="x", pady=(6,6))
        ttk.Label(frm_hol, text="Days:").pack(side="left", padx=(6,4))
        ttk.Entry(frm_hol, textvariable=self.var_holiday_days, width=18).pack(side="left", padx=(0,12))
        ttk.Label(frm_hol, text="Title:").pack(side="left")
        ttk.Entry(frm_hol, textvariable=self.var_holiday_title, width=32).pack(side="left", padx=(4,12))
        ttk.Button(frm_hol, text="Submit Holidays", command=self._on_add_holiday).pack(side="left")

        # ---- Unavailabilities / Preferences ----
    
        # ---- Unavailabilities / Preferences ----
        frm_unav = ttk.LabelFrame(self, text="Unavailabilities")
        frm_unav.pack(fill="x", pady=(6,6))

        # Use grid so the button column (right) is always visible
        frm_unav.grid_columnconfigure(0, weight=1)  # left column grows
        frm_unav.grid_columnconfigure(1, weight=0)

        left_unav = ttk.Frame(frm_unav)
        left_unav.grid(row=0, column=0, sticky="w", padx=6, pady=4)

        right_unav = ttk.Frame(frm_unav)
        right_unav.grid(row=0, column=1, sticky="e", padx=6, pady=4)

        # Left part (options)
        ttk.Label(left_unav, text="Person:").pack(side="left", padx=(0,4))
        self.cb_person = ttk.Combobox(
            left_unav,
            textvariable=self.var_person_label,
            values=[], width=60, state="readonly"
        )
        self.cb_person.pack(side="left", padx=(0,8))

        ttk.Checkbutton(
            left_unav,
            text="Preferences (instead of Unavailabilities)",
            variable=self.var_is_pref
        ).pack(side="left", padx=(0,12))

        ttk.Label(left_unav, text="Days (e.g., 4,10-12):").pack(side="left")
        ttk.Entry(left_unav, textvariable=self.var_days_unav, width=18).pack(side="left", padx=(4,12))

        ttk.Label(left_unav, text="Comment:").pack(side="left")
        ttk.Entry(left_unav, textvariable=self.var_unav_note, width=36).pack(side="left", padx=(4,12))

# Right part (buttons)
        ttk.Button(right_unav, text="Submit", command=self._on_add_unav_or_pref)\
            .pack(side="top", fill="x", pady=(0,4))
        ttk.Button(right_unav, text="Delete", command=self._on_delete_unav_or_pref)\
            .pack(side="top", fill="x")      

        # Preview table
        grp_tbl = ttk.LabelFrame(self, text="Preview (after calculation)")
        grp_tbl.pack(fill="both", expand=True, pady=(6,0))

        cols   = ("date","weekday","AF","YF","YFM","BYFM","BYF")
        heads  = ["Date","Day","AF","YF","YFM","BYFM","BYF"]
        widths = [120, 150, 120, 120, 120, 120, 120]
        self.tree = ttk.Treeview(grp_tbl, columns=cols, show="headings", height=12)
        for c, h, w in zip(cols, heads, widths):
            self.tree.heading(c, text=h)
            self.tree.column(c, width=w, anchor="w")
        yscroll = ttk.Scrollbar(grp_tbl, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(grp_tbl, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        grp_tbl.grid_rowconfigure(0, weight=1); grp_tbl.grid_columnconfigure(0, weight=1)

        # Log
        grp_log = ttk.LabelFrame(self, text="Logs")
        grp_log.pack(fill="both", expand=False, pady=(8,0))
        self.txt_log = tk.Text(grp_log, height=8, wrap="word")
        self.txt_log.pack(fill="both", expand=True, padx=6, pady=6)

    # ---------------- helpers ----------------
    def _log(self, msg: str):
        self.txt_log.insert("end", msg + "\n"); self.txt_log.see("end")

    def _sel_ym(self):
        """Return (year:int, month:int) from UI vars."""
        try:
            y = int(self.var_year.get()); m = int(self.var_month.get())
            return y, m
        except Exception:
            # safe default: current month
            t = pd.Timestamp.today()
            return int(t.year), int(t.month)

    def reload_people(self):
        """Populate the personnel list for the combobox."""
        df = load_personnel_df()
        people = []
        for _, r in df.iterrows():
            rid = str(r["registry_number"]).strip()
            nm  = str(r["name"]).strip()
            rk  = str(r["rank"]).strip()
            sp  = str(r["specialty"]).strip()
            label = f"{rid} | {rk} ({sp}) | {nm}"
            people.append({"rid": rid, "label": label})
        self.people = people
        labels = [p["label"] for p in people]
        self.cb_person["values"] = labels
        if labels and not self.var_person_label.get():
            self.cb_person.current(0)
        self._log(f"Available personnel: {len(people)}")

    # ---------------- actions: ship/holidays ----------------
    def _on_init_month_ormo(self):
        """Set the entire month to IN PORT."""
        y, m = self._sel_ym()
        try:
            msg = set_month_all_in_port(y, m)
            self._log(f"✅ Initialization: {msg}")
            messagebox.showinfo("Done", msg)
        except Exception as e:
            self._log(f"❌ Initialization Error: {e}")
            messagebox.showerror("Initialization Failed", f"{e}")

    def _on_ship_en_plo(self):
        """Set selected days to AT SEA."""
        y, m = self._sel_ym()
        days_str = self.var_days_ship.get().strip()
        if not days_str:
            messagebox.showwarning("Warning", "Enter days (e.g. 3, 5-7).")
            return
        try:
            msg = set_ship_status_bulk(y, m, days_str, "at sea")
            self._log(f"✅ {msg}")
            messagebox.showinfo("Done", "Ship status updated for the selected days.")
        except Exception as e:
            messagebox.showerror("Failed to set 'at sea'", f"{e}")

    def _on_add_holiday(self):
        """Submit holidays for the month."""
        y, m = self._sel_ym()
        days_str = self.var_holiday_days.get().strip()
        title    = (self.var_holiday_title.get() or "Holiday").strip()
        if not days_str:
            messagebox.showwarning("Warning", "Enter holiday days.")
            return
        try:
            msg = add_holiday(y, m, days_str, title)  # positional
            self._log(f"✅ {msg}")
            messagebox.showinfo("Done", "Holidays have been submitted.")
        except Exception as e:
            messagebox.showerror("Failed to submit holidays", f"{e}")

    # ---------------- actions: Unav / Pref ----------------
    def _selected_registry(self) -> str:
        """Find the registry_number from the selected label."""
        label = self.var_person_label.get().strip()
        for p in self.people:
            if p["label"] == label:
                return p["rid"]
        return ""

    def _on_add_unav_or_pref(self):
        """Submit unavailabilities or preferences for the selected month."""
        rid = self._selected_registry()
        if not rid:
            messagebox.showerror("Error", "Choose a person."); return

        days_str = self.var_days_unav.get().strip()
        note = self.var_unav_note.get().strip()
        if not days_str:
            messagebox.showwarning("Warning", "Enter days (e.g. 4,10-12).")
            return

        y, m = self._sel_ym()
        try:
            if self.var_is_pref.get():
                msg = add_preference(rid, y, m, days_str, note or "")
            else:
                msg = add_unavailable(rid, y, m, days_str, note or "")
            self._log("✅ " + msg)
            messagebox.showinfo("Done", "Submission complete.")
        except Exception as e:
            messagebox.showerror("Submission Error", f"{e}")

    def _on_delete_unav_or_pref(self):
        """
        Delete records that match person/days/comment.
        If you have explicit delete* functions in calendar_service,
        use them here. Alternatively, convention: pass
        days_str with a '-' prefix for deletion.
        """
        rid = self._selected_registry()
        if not rid:
            messagebox.showerror("Error", "Choose a person."); return

        days_str = self.var_days_unav.get().strip()
        note = self.var_unav_note.get().strip()
        if not days_str:
            messagebox.showwarning("Warning", "Enter days to delete.")
            return

        y, m = self._sel_ym()
        try:
            if self.var_is_pref.get():
                msg = add_preference(rid, y, m, f"-{days_str}", note or "")
            else:
                msg = add_unavailable(rid, y, m, f"-{days_str}", note or "")
            self._log("🗑️ " + msg)
            messagebox.showinfo("Done", "Corresponding records removed (if they existed).")
        except Exception as e:
            messagebox.showerror("Deletion Error", f"{e}")

    # ---------------- compute / preview / export ----------------
    def _on_compute(self):
        """Run the scheduler for the selected month and show a preview."""
        y, m = self._sel_ym()
        try:
            self.result = make_month_schedule_all(y, m)
            self._refresh_preview_from_result()
            self._log(f"✅ Shift calculation for {y}-{m:02d} completed.")
            messagebox.showinfo("Done", f"Shift calculation for {y}-{m:02d}.")
        except Exception as e:
            messagebox.showerror("Calculation Error", f"{e}")

    def _refresh_preview_from_result(self):
        """Safe preview from self.result."""
        # FIX 4a: Add these two helper methods right inside this function for clarity and access
        def _is_officer(rank: str) -> bool:
            officer_ranks = {
                "Commander","Commander (M)","Lieutenant Commander","Lieutenant Commander (M)",
                "Lieutenant","Lieutenant (M)","Lieutenant (E)",
                "Ensign","Ensign (M)","Ensign (E)"
            }
            return str(rank).strip() in officer_ranks

        def _display_name(rank: str, specialty: str, name: str) -> str:
            rank = str(rank).strip()
            specialty = str(specialty).strip()
            name = str(name).strip()
            if _is_officer(rank):
                return f"{rank} | {name} HN"
            spec = f" ({specialty})" if specialty else ""
            return f"{rank}{spec} | {name}"

        for iid in self.tree.get_children():
            self.tree.delete(iid)
        if not self.result:
            return

        # FIX 4b: The rest of the function logic is corrected here
        assignments = self.result.get("by_watch", {})
        y, m = self._sel_ym()
        _, last_day = monthrange(y, m)
        
        WEEKDAY_EN = ["MONDAY","TUESDAY","WEDNESDAY","THURSDAY","FRIDAY","SATURDAY","SUNDAY"]
        
        from app.scheduling_prep import WATCH_TYPES as SCHEDULER_WATCH_TYPES
        
        for d in range(1, last_day + 1):
            date_iso = pd.Timestamp(year=y, month=m, day=d).date().isoformat()
            wname = WEEKDAY_EN[pd.Timestamp(date_iso).weekday()]
            
            values = [date_iso, wname]
            
            # The treeview columns are ("AF","YF","YFM","BYFM","BYF") after date/weekday
            # SCHEDULER_WATCH_TYPES is ["ΑΦ", "ΥΦ", "ΥΦΜ", "ΒΥΦΜ", "ΒΥΦ"]
            for w in SCHEDULER_WATCH_TYPES:
                entry = assignments.get(w, {}).get(date_iso)
                
                name = ""
                if entry == "SEA":
                    name = "At Sea"
                elif isinstance(entry, dict):
                    name = _display_name(entry.get("rank", ""), entry.get("specialty", ""), entry.get("name", ""))
                
                values.append(name)
                
            self.tree.insert("", "end", values=tuple(values))

    def _on_export_all_excels(self):
        """Export 3 Excel files for the selected month, using the result cache (if available)."""
        y, m = self._sel_ym()
        try:
            out_paths = export_month_schedule_all(y, m, self.result)
            self._log(f"✅ Export complete: {out_paths}")
            messagebox.showinfo("Success", "Export of 3 Excel files is complete.")
        except Exception as e:
            messagebox.showerror("Export Failed", f"{e}")
# =============================== APP (Notebook) ===============================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("OA3801 – Personnel / Leave / Shift Management")
        self.geometry("1200x780")

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True)

        # Tab 1
        self.tab1 = PersonnelManager(nb, on_personnel_changed=self._on_people_changed)
        nb.add(self.tab1, text="Personnel")

        # Tab 2
        self.tab2 = LeaveManager(nb)
        nb.add(self.tab2, text="Leave")

        # Tab 3
        self.tab3 = ShiftsManager(nb)
        nb.add(self.tab3, text="Shifts")

    def _on_people_changed(self):
        # When Personnel changes, we update the other tabs
        try:
            self.tab2.reload_people()
        except Exception:
            pass
        try:
            self.tab3.reload_people()
        except Exception:
            pass


def main():
    ensure_personnel_csv()  # make sure the folder/file exists
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()