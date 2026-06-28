# Gelatik Automation — SOS Dashboard Generator

**Version:** 1.0.0

Generates per-division SOS (Share of Shelf) summary workbooks from weekly Report Product CSV exports.

---

## Folder Structure

```
SOS Dashboard Automation/
├── Automation.exe          ← Double-click to run (Windows)
├── main.py                 ← Entry point (development)
├── tes.py                  ← Engine (do not edit)
├── config.json             ← User preferences (auto-managed)
├── VERSION                 ← Version string
├── requirements.txt        ← Python dependencies
├── requirements_sheets.txt ← Optional Google Sheets deps
├── gui/                    ← GUI package
├── assets/                 ← Logo and icon
├── logs/
│   └── automation.log      ← Full process log
├── Nutrition/              ← Example cluster folder
│   ├── Report Product - Jan 25 - Nutrition.csv
│   └── Summary SOS_Nutrition.xlsx   ← Generated output
└── Indulgence/
    ├── Report Product - Jan 25 - Indulgence.csv
    └── Summary SOS_Indulgence.xlsx
```

---

## How to Run

### Windows (production)
Double-click `Automation.exe`.

### Development (macOS / Windows with Python)
```bash
pip install -r requirements.txt
python3 main.py
```

---

## How to Use

1. **Open the application** → splash screen appears for 2 seconds.
2. **Browse Folder** → select the root automation folder (e.g. `SOS Dashboard Automation/`).
   - The app auto-detects sub-folders containing Report Product CSV files.
3. **Generate Summary** → runs the full pipeline in the background.
   - Progress bar and log stream update in real time.
   - The window stays responsive throughout.
4. **Completion dialog** → shows per-cluster result and elapsed time.
5. **Buka Folder Output** → opens the folder in Explorer / Finder.

---

## CSV File Format

CSV files must be named:

```
Report Product - <Month> <YY> - <Division>.csv
```

Example:
```
Report Product - Jan 25 - Nutrition.csv
Report Product - Feb 25 - Indulgence.csv
```

Required columns: `Visit Date`, `Product Code`, `Facing`

Optional columns: `Month`, `Region`, `Channel`, `Account`, `Category Channel`, `Brand`, `Store Code`, `Week`, `Area`

---

## How to Edit TARGET

The TARGET workbook is embedded in `Summary SOS_<Cluster>.xlsx` on the **TARGETS** sheet.

1. Open `Summary SOS_<Cluster>.xlsx`.
2. Go to the **TARGETS** sheet.
3. Edit the `Target` column values (numbers, e.g. `70` = 70%).
4. Save and close the file.
5. Click **Generate Summary** again — targets are read from the saved workbook.

TARGET columns:

| Column | Description |
|--------|-------------|
| Division | Source division name (e.g. `Nutrition`) |
| Dimension | `REGION`, `CHANNEL`, `ACCOUNT GELATIK`, `CATEGORY CHANNEL`, `CHANNEL-ACCOUNT` |
| Name | Dimension value (e.g. `WEST`, `MT`) |
| Target | Target SOS% (number, e.g. `65` for 65%) |

---

## How to Add New CSV

1. Place the new CSV file in the cluster folder (e.g. `Nutrition/`).
2. Ensure the filename follows the naming convention above.
3. Click **Generate Summary**.

---

## Output

For each detected cluster, the app generates:

```
Summary SOS_<ClusterName>.xlsx
```

Sheets per workbook:
- One sheet per Source Division (e.g. `Nutrition`, `Indulgence`)
- `TARGETS` sheet — editable targets

Each division sheet contains:
- Division Summary table
- SOS Monthly table (by region, W1–W5 breakdown)
- Channel & Account table
- Category by Division table
- Bar charts for each category

---

## Logging

All process details are written to:

```
logs/automation.log
```

Log is rotated at 5 MB (3 backup files kept).

If a generate fails, check this file for the full error traceback.

---

## Troubleshooting

| Symptom | Solution |
|---------|----------|
| "File Excel masih terbuka" | Close all open Excel files before generating |
| "Tidak ada CSV ditemukan" | Ensure CSVs are named `Report Product*.csv` |
| Progress bar stops | Check `logs/automation.log` for error detail |
| Output file is 0 KB | Deleted corrupt file — re-generate |
| Clusters not detected | Ensure cluster sub-folders contain `.csv` files |

---

## Packaging (Windows Executable)

On a Windows machine with Python installed:

```
pip install pyinstaller
build_windows.bat
```

Output: `dist/Automation.exe`

---

## FAQ

**Q: Can I have multiple cluster folders?**
A: Yes. Browse to the parent folder — the app detects all sub-folders with CSVs automatically.

**Q: Will re-generating overwrite my TARGET edits?**
A: No. Targets are read from the existing workbook (if present) and merged with new data values. User edits are never overwritten.

**Q: Where is config stored?**
A: `config.json` next to the exe / `main.py`. It saves window size and last used folder.

**Q: Can I use this with Google Sheets?**
A: The engine supports Google Sheets sync via the optional `requirements_sheets.txt` dependencies. The GUI currently generates Excel only.
