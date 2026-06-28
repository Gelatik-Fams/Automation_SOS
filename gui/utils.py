"""Pure-Python utilities — no GUI/tkinter dependency."""
import os


def user_friendly_error(exc: Exception) -> str:
    msg = str(exc).lower()
    if 'permission' in msg or 'access denied' in msg or 'winerror 32' in msg:
        return 'File Excel masih terbuka. Tutup dahulu sebelum generate.'
    if 'tidak ada csv' in msg or ('not found' in msg and 'csv' in msg):
        return 'Tidak ada file CSV ditemukan.'
    if 'targets' in msg and ('tidak' in msg or 'not found' in msg):
        return 'TARGET workbook tidak ditemukan atau rusak.'
    if 'corrupt' in msg or ('invalid' in msg and 'workbook' in msg):
        return 'Workbook output corrupt — hapus file lama lalu coba lagi.'
    return f'Error: {exc}'


def validate_workbook(path: str) -> str | None:
    """Return None if workbook is valid, else a short error string."""
    try:
        from openpyxl import load_workbook as _lw
        wb = _lw(path, read_only=True)
        sheets = wb.sheetnames
        wb.close()
        if 'TARGETS' not in sheets:
            return 'Sheet TARGETS tidak ada'
        if len(sheets) < 2:
            return 'Workbook tidak lengkap (tidak ada sheet divisi)'
        return None
    except Exception as e:
        return f'Workbook tidak dapat dibuka: {e}'


def fmt_elapsed(seconds: float) -> str:
    m, s = divmod(int(seconds), 60)
    return f'{m:02d}:{s:02d}'


def _has_csv_anywhere(path: str) -> bool:
    """True if path contains any .csv file in its entire subtree."""
    for _, _, files in os.walk(path):
        if any(f.endswith('.csv') for f in files):
            return True
    return False


def detect_clusters(folder: str) -> list[str]:
    """
    Detect clusters in the browsed folder.

    Rules:
    - If the root folder itself contains .csv files → it IS the single cluster.
    - Otherwise, each immediate subfolder that contains .csv files anywhere
      in its subtree is treated as one cluster. The engine's recursive CSV
      discovery (discover_report_product_files) will pick up all division
      CSVs within that subtree and combine them into one workbook.

    Example:
        Indulgence/         ← user browses here → one cluster "Indulgence"
            Noodle/
                Report Product - Jan 25 - Noodle.csv
            Nici/
                Report Product - Jan 25 - Nici.csv
        → Summary SOS_Indulgence.xlsx  (sheets: Noodle, Nici, TARGETS)

        Automation/         ← user browses here → two clusters
            Indulgence/     ← has CSVs in subtree → cluster 1
            Nutrition/      ← has CSVs in subtree → cluster 2
    """
    if not os.path.isdir(folder):
        return []

    # Root folder has CSVs directly → treat it as the sole cluster
    if any(f.endswith('.csv') for f in os.listdir(folder)):
        return [os.path.basename(folder)]

    # Otherwise each immediate subdir that has CSVs anywhere inside = one cluster
    found = []
    for entry in sorted(os.scandir(folder), key=lambda e: e.name):
        if entry.is_dir() and _has_csv_anywhere(entry.path):
            found.append(entry.name)
    return found
