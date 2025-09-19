# Shipboard_Watch_Scheduling_Interface

OA3801 – Personnel / Leave / Watchbill Management

Overview
--------
This application is a Tkinter-based GUI tool for managing naval personnel, leave, and in-port watchbills. It was developed for the OA3801 course to automate several manual scheduling and reporting tasks.

The program has three main tabs:

1. Personnel
   - Manage all personnel records (name, rank, specialty, duties).
   - Assign Primary, Alternative, and At-Sea watches.
   - Import personnel data from a simple Excel file (this will replace the current list).
   - Save changes to a persistent CSV file (data/personnel.csv).
   - Export personnel data to Excel (Personnel.xlsx).

2. Leave Management
   - Assign leave periods (From–To) with type and comments.
   - Supports filtering by year/month and sorting.
   - Right-click context menu for editing or deleting leave entries.
   - Export all leave data to a single Excel file with one sheet per month (Personnel_Leave.xlsx).

3. Shifts / Watchbill (In-Port)
   - Select Year/Month for scheduling.
   - Initialize all days as "in port".
   - Set specific days as "at sea" or "holidays".
   - Add Unavailabilities (days a person cannot stand watch) and Preferences (preferred days).
   - Run the scheduler for automatic assignment of AF, YF, YFM, BYFM, BYF watches.
   - Preview results in the GUI.
   - Export three Excel files:
     - Daily Calendar (Calendar_YYYY-MM.xlsx)
     - Monthly Summary (Monthly_Summary_YYYY-MM.xlsx)
     - Personnel Overview (Personnel.xlsx)

Excel exports feature:
- Grey shading for weekends/holidays.
- Clearly marked “at sea” days.
- Per-person watch statistics (total watches, holidays, weekends).
- A final "Statistics_Total" sheet with overall totals.

Project Structure
-----------------
```
OA3801_CM_Project/
│
├── app/
│   ├── gui_app_ENGLISH.py              # Main Tkinter GUI
│   ├── scheduler_in_port.py    # In-port watch scheduler + Excel export
│   ├── calendar_service.py     # Leave/unavailability/holiday handling
│   ├── export_service.py       # Export helpers
│   ├── scheduler_rules.py      # Constraints and business rules
│   └── scheduling_prep.py      # Daily availability utilities
│
├── data/
│   └── personnel.csv           # Persistent personnel database (UTF-8 with BOM)
│
├── logs/                       # Leave, ship status, unavailability logs
│
└── readme.txt                  # This file
```

Requirements
------------
- Python 3.10+
- Required packages (install with pip if missing):
    pip install pandas openpyxl xlsxwriter
- Compatible with macOS, Linux, and Windows.
- Recommended editor: VS Code.

Running the Application
----------------------
1. Open the project folder in VS Code.
2. Make sure your terminal is set to the project root (OA3801_CM_Project).
3. Run the GUI with:

    PYTHONPATH=. python app/gui_app_ENGLISH.py

4. The main window will open with three tabs: Personnel, Leave, Watchbills.

Workflow
--------
Personnel Tab
- Add or edit personnel records.
- Assign Primary/Alternative/At-Sea watches.
- Save changes to update data/personnel.csv.
- Import from minimal Excel (optional) – this replaces all records (confirmation required).
- Export personnel to data/Personnel.xlsx.

Leave Tab
- Select a person and add a leave period.
- Supports multiple leave types (normal, parental, marriage, etc.).
- Filter by month/year or sort ascending/descending.
- Right-click an entry to edit or delete it.
- Export all leave data to data/Personnel_Leave.xlsx with one sheet per month.

Watchbill Tab
- Select Year/Month.
- Initialize all days as in port.
- Set "at sea" days and holidays.
- Add Unavailabilities or Preferences per person.
- Run Compute Shifts – preview table will be filled.
- Export three Excel files:
    - Calendar_YYYY-MM.xlsx
    - Monthly_Summary_YYYY-MM.xlsx
    - Personnel.xlsx

Constraints & Business Rules
---------------------------
- The Captain is never scheduled.
- Maximum watches per rank/month are enforced (see scheduler_rules.py).
- At least 2 days gap between duties for the same person.
- Max 2 weekend watches per person/month.
- Max 1 holiday watch per person/month.
- AF watch: assigned from youngest to oldest; Executive Officer / Deputy Commanding Officer only Monday–Friday.
- Other watches: fair distribution by total count, seniority, and name.

Troubleshooting
---------------
- Empty Excel exports:
    Be sure to Compute Shifts (button in Tab 3) before exporting.
- Data lost in personnel.csv:
    This happens if you re-import Excel. Only use Save when editing individual records to preserve shifts.
- Restarting quickly in VS Code:
    Press Ctrl+Shift+P → "Python: Restart Terminal" and re-run the command.

Translation & Data Mapping
--------------------------
This application was originally developed with a Greek user interface and data structure. It has been translated to use English for all user-facing elements and most data. However, there is a critical distinction between display text and internal data keys.

Translated Data: All user-facing text in the GUI, as well as the data values in constants.py and personnel.csv (e.g., rank, duty, name), have been translated to English.

Untranslated Keys: For the scheduling logic to function correctly, the values in the primary_shift and alt_shift columns of the data/personnel.csv file must remain in their original Greek acronym format (e.g., ΑΦ, ΥΦ, ΥΦΜ). These values are used as internal keys to match personnel to the correct watch types. Changing them to English (AF, YF, etc.) in the CSV file will break the scheduler.

The i18n_display_mapping.py file is designed to manage the mapping between these internal keys and their display values, though in the current English version, this mapping is mostly 1-to-1.

Contributors
------------

LT Nikolaos Thanos  
LT Kirsten Cavanah  
LT Alexis Harrell-Parada  
LT Ryan Gallagher

Developed for OA3801 – Computational Methods
Naval Postgraduate School, 2025
