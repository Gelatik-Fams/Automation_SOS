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


def detect_clusters(folder: str) -> list[str]:
    if not os.path.isdir(folder):
        return []
    found = []
    if any(f.endswith('.csv') for f in os.listdir(folder)):
        found.append(os.path.basename(folder))
    for entry in sorted(os.scandir(folder), key=lambda e: e.name):
        if entry.is_dir() and any(f.endswith('.csv') for f in os.listdir(entry.path)):
            found.append(entry.name)
    return found
