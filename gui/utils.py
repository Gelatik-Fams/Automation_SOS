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
    for _, _, files in os.walk(path):
        if any(f.endswith('.csv') for f in files):
            return True
    return False


def _has_csv_at_root(path: str) -> bool:
    try:
        return any(f.endswith('.csv') for f in os.listdir(path))
    except OSError:
        return False


def detect_cluster_dirs(folder: str) -> list[tuple[str, str]]:
    """
    Return (dir_path, cluster_name) pairs for the browsed folder.

    Detection rules — three cases:

    1. Browsed folder has CSVs directly at root → single cluster (the folder itself).

    2. Immediate subdirs have NO CSVs at their own root (only deeper down) →
       those subdirs are cluster-level folders. Each becomes its own cluster.

       Automation/          ← browse here
           Indulgence/      ← no CSV at root, but has CSV in subtree → cluster
           Nutrition/       ← same → cluster

    3. Immediate subdirs DO have CSVs at their own root → those are division
       folders, not cluster folders. Treat the browsed folder itself as the
       single cluster; the engine's recursive discovery will collect all CSVs.

       Indulgence/          ← browse here → single cluster "Indulgence"
           Nici/            ← has CSVs at root → this is a division folder
           IFM/             ← same
           Noodle/          ← same
           SIMP/            ← same
       → Summary SOS_Indulgence.xlsx (sheets: Nici, IFM, Noodle, SIMP)
    """
    if not os.path.isdir(folder):
        return []

    # Case 1: CSVs directly in browsed folder
    if _has_csv_at_root(folder):
        return [(folder, os.path.basename(folder))]

    cluster_dirs: list[tuple[str, str]] = []
    division_dirs: list[tuple[str, str]] = []

    for entry in sorted(os.scandir(folder), key=lambda e: e.name):
        if not entry.is_dir() or not _has_csv_anywhere(entry.path):
            continue
        if _has_csv_at_root(entry.path):
            division_dirs.append((entry.path, entry.name))
        else:
            cluster_dirs.append((entry.path, entry.name))

    # Case 2: subfolders are cluster-level containers
    if cluster_dirs:
        return cluster_dirs

    # Case 3: subfolders are division-level folders → browsed folder is the cluster
    if division_dirs:
        return [(folder, os.path.basename(folder))]

    return []


def detect_clusters(folder: str) -> list[str]:
    return [name for _, name in detect_cluster_dirs(folder)]
