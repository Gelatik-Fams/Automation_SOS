"""CSV discovery, ingestion, cleaning, and duplicate validation."""

import glob
import os
import pandas as pd

from sos_engine.calculations import (
    is_competitor,
)
from sos_engine.common import (
    METADATA_COLS_FOR_DEDUP,
    UNKNOWN_SOURCE_DIVISION,
    sort_key_period,
)


def extract_source_division_from_filename(filename):
    """Ambil divisi dari segmen terakhir filename: Report Product - Mei 25 - Oil & Fat.csv."""
    stem = os.path.splitext(os.path.basename(str(filename)))[0]
    parts = [p.strip() for p in stem.split(' - ')]
    if len(parts) >= 3 and parts[-1]:
        return parts[-1]
    return UNKNOWN_SOURCE_DIVISION


def extract_source_division_from_raw_data(df, file_label='CSV'):
    """Ambil Source Division dari kolom raw Divisi; satu file harus berisi satu divisi."""
    if 'Divisi' not in df.columns:
        raise ValueError(f'CSV tidak valid: kolom Divisi tidak ditemukan di {file_label}')

    divisions = (
        df['Divisi']
        .dropna()
        .astype(str)
        .str.strip()
    )
    divisions = sorted(v for v in divisions.unique() if v)

    if not divisions:
        raise ValueError(f'CSV tidak valid: kolom Divisi kosong di {file_label}')
    if len(divisions) > 1:
        raise ValueError(
            f'CSV tidak valid: lebih dari satu Divisi ditemukan di {file_label}: '
            + ', '.join(divisions)
        )
    return divisions[0]


def discover_report_product_files(base_dir='.'):
    """Temukan semua Report Product CSV di root dan subfolder division."""
    pattern = os.path.join(base_dir, '**', 'Report Product*.csv')
    files = glob.glob(pattern, recursive=True)
    return sorted(
        files,
        key=lambda p: os.path.relpath(p, base_dir).replace(os.sep, '/').lower()
    )


def display_csv_path(path, base_dir='.'):
    return os.path.relpath(path, base_dir).replace(os.sep, '/')


def count_files_by_source_division(files):
    counts = {}
    for f in files:
        division = extract_source_division_from_filename(f)
        counts[division] = counts.get(division, 0) + 1
    return dict(sorted(counts.items()))


def count_physical_text_lines(path):
    """Hitung baris fisik text untuk kebutuhan log user-facing."""
    with open(path, 'rb') as f:
        return sum(1 for _ in f)


def validasi_data(df):
    n = len(df)
    cols_to_check = [c for c in df.columns if c not in METADATA_COLS_FOR_DEDUP]
    df_clean_no_src = df[cols_to_check].drop_duplicates()
    df_clean = df.loc[df_clean_no_src.index]
    df_removed = df[~df.index.isin(df_clean.index)].copy()
    df = df_clean
    if len(df) < n:
        print(f'[INFO] {n - len(df)} baris duplikat dihapus.')
    df_i      = df[~is_competitor(df['Produsen'])]
    baris_nan = df_i[df_i['Facing'].isna()]
    if not baris_nan.empty:
        print(f'[WARNING] {len(baris_nan)} baris Indofood tidak punya Facing.')
    return df, df_removed


def baca_semua_csv():
    base_dir = '.'
    files = discover_report_product_files(base_dir)
    if not files:
        print('[WARNING] Tidak ada file CSV ditemukan.')
        return None

    print(f'[INFO] Membaca {len(files)} file:')
    dfs = []
    division_file_counts = {}
    total_physical_lines = 0
    for f in files:
        display_path = display_csv_path(f, base_dir)
        print(f'-> {display_path}')
        total_physical_lines += count_physical_text_lines(f)
        try:
            df_temp = pd.read_csv(f, low_memory=False)
        except pd.errors.ParserError:
            df_temp = pd.read_csv(f, sep=';', low_memory=False)
        source_division = extract_source_division_from_raw_data(df_temp, display_path)
        df_temp['_source_file'] = os.path.splitext(os.path.basename(f))[0]
        df_temp['Source Division'] = source_division
        df_temp['CSV Row'] = df_temp.index + 2
        division_file_counts[source_division] = division_file_counts.get(source_division, 0) + 1
        print(f'   Division: {source_division}')
        print(f'   Rows: {len(df_temp):,}')
        dfs.append(df_temp)

    print('[INFO] File Count by Source Division')
    for division, count in sorted(division_file_counts.items()):
        print(f'{division:<12} {count} files')

    combined = pd.concat(dfs, ignore_index=True)

    mask_valid = combined['Visit Date'].notna()
    if (~mask_valid).sum() > 0:
        print(f'[INFO] Dihapus {(~mask_valid).sum()} baris rusak (tanpa Visit Date).')
        combined = combined[mask_valid].copy()

    combined['Visit Date'] = pd.to_datetime(combined['Visit Date'], format='mixed', dayfirst=False)

    # Period = "Jan 24", "Feb 24", "Jan 25", dst.
    # Hybrid rule: pakai Month column HANYA jika lebih maju dari bulan Visit Date
    # (advance reporting — mis. Jan 24 file punya baris Month="February" → Feb 24).
    # Jika Month sama atau lebih mundur dari Visit Date → pakai Visit Date
    # (late submission). Ini membuat Jan/Feb 25 tetap cocok dengan pivot manual.
    MONTH_NUM = {
        'January':1,'February':2,'March':3,'April':4,'May':5,'June':6,
        'July':7,'August':8,'September':9,'October':10,'November':11,'December':12,
    }
    MONTH_ABB = {
        1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'May',6:'Jun',
        7:'Jul',8:'Aug',9:'Sep',10:'Oct',11:'Nov',12:'Dec',
    }
    visit_m = combined['Visit Date'].dt.month
    year_str = combined['Visit Date'].dt.strftime('%y')
    if 'Month' in combined.columns:
        col_m   = combined['Month'].map(MONTH_NUM)
        use_col = col_m > visit_m        # Month column lebih maju → pakai Month
        eff_m   = col_m.where(use_col, visit_m).fillna(visit_m).astype(int)
    else:
        eff_m = visit_m
    combined['Period'] = eff_m.map(MONTH_ABB) + ' ' + year_str

    # Week — pastikan valid untuk semua baris
    if 'Week' not in combined.columns:
        combined['Week'] = combined['Visit Date'].dt.day.apply(
            lambda d: f'W{min((d - 1) // 7 + 1, 5)}'
        )
    else:
        combined['Week'] = combined['Week'].astype(str).str.strip().str.upper()
        mask_invalid = ~combined['Week'].str.match(r'^W[1-5]$', na=True)
        combined.loc[mask_invalid, 'Week'] = (
            combined.loc[mask_invalid, 'Visit Date'].dt.day
            .apply(lambda d: f'W{min((d - 1) // 7 + 1, 5)}')
        )

    # Pastikan kolom Account ada
    if 'Account' not in combined.columns:
        combined['Account'] = combined.get('Subchannel', 'UNKNOWN')

    combined.attrs['physical_csv_lines'] = total_physical_lines
    print(f'[INFO] {total_physical_lines:,} baris fisik CSV terdeteksi.')
    summary = combined.groupby('Period').size().reset_index(name='rows')
    summary = summary.sort_values('Period', key=lambda s: s.map(sort_key_period))
    print(summary.to_string(index=False))
    return combined
