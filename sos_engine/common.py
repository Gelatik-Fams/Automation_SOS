"""Shared constants and helpers for the SOS engine."""

import os


# ── Warna conditional formatting ──────────────────────────────
GREEN_BG = {'red': 0.714, 'green': 0.843, 'blue': 0.659}
RED_BG   = {'red': 0.918, 'green': 0.600, 'blue': 0.600}

# ── Metrik detail untuk STORE DETAIL ──────────────────────────
METRIC_LABELS = ['Indofood', 'Kompetitor', 'Total Facing', 'SOS%', 'Store Count']
N_METRICS     = len(METRIC_LABELS)

WEEK_ORDER = ['W1', 'W2', 'W3', 'W4', 'W5']
UNKNOWN_SOURCE_DIVISION = 'UNKNOWN'
METADATA_COLS_FOR_DEDUP = {'_source_file', 'Source Division', 'CSV Row'}
TARGETS_FILENAME = 'TARGETS.xlsx'
DEFAULT_TARGET = 65.0
TARGET_DIM_ORDER = ['REGION', 'CHANNEL', 'ACCOUNT GELATIK', 'CATEGORY CHANNEL', 'CHANNEL-ACCOUNT']
EXCEL_CATEGORY_CHART_ROW_STEP = 22


def col_letter(n):
    """0-indexed → huruf kolom (0=A, 25=Z, 26=AA, ...)."""
    result = ''
    n += 1
    while n > 0:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result


def get_dashboard_division_options(df):
    if 'Source Division' not in df.columns:
        return ['ALL']
    divisions = sorted(str(v) for v in df['Source Division'].dropna().unique())
    return ['ALL'] + divisions


def filter_dashboard_by_division(df, selected_division):
    if selected_division == 'ALL' or 'Source Division' not in df.columns:
        return df
    return df[df['Source Division'].astype(str) == selected_division].copy()


def sort_key_period(label):
    """Sort 'Jan 25', 'Feb 25' dst. secara kronologis."""
    MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
              'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    parts = str(label).split()
    if len(parts) == 2:
        m = MONTHS.index(parts[0]) if parts[0] in MONTHS else 99
        y = int(parts[1]) if parts[1].isdigit() else 9999
        return (y, m)
    return (9999, 99)


def get_cluster_name(base_dir='.'):
    name = os.path.basename(os.path.abspath(base_dir))
    return name or 'Cluster'


def get_summary_output_path(output_dir='.', cluster_name=None, extension='.xlsx'):
    cluster = cluster_name or get_cluster_name(output_dir)
    return os.path.join(output_dir, f'Summary SOS_{cluster}{extension}')


def get_store_detail_output_path(output_dir='.', cluster_name=None, extension='.xlsx'):
    cluster = cluster_name or get_cluster_name(output_dir)
    return os.path.join(output_dir, f'Store Detail_{cluster}{extension}')


def get_source_divisions(df):
    if 'Source Division' not in df.columns:
        return ['ALL']
    return sorted(str(v) for v in df['Source Division'].dropna().unique())
