import pandas as pd
import time
import os
import glob
try:
    import gspread
except ImportError:
    gspread = None
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.data_source import AxDataSource, StrRef
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import quote_sheetname
from watchdog.observers.polling import PollingObserver
from watchdog.events import FileSystemEventHandler

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
_last_excel_write_time = 0.0


# ─────────────────────────── HELPER ────────────────────────────

def col_letter(n):
    """0-indexed → huruf kolom (0=A, 25=Z, 26=AA, ...)."""
    result = ''
    n += 1
    while n > 0:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result


def is_competitor(series):
    return series.astype(str).str.contains('COMPETITOR', case=False, na=False)


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


def get_dashboard_division_options(df):
    if 'Source Division' not in df.columns:
        return ['ALL']
    divisions = sorted(str(v) for v in df['Source Division'].dropna().unique())
    return ['ALL'] + divisions


def baca_dashboard_division_selector(ws, options):
    try:
        selected = str(ws.acell('B1').value or '').strip()
    except Exception:
        selected = ''
    return selected if selected in options else 'ALL'


def filter_dashboard_by_division(df, selected_division):
    if selected_division == 'ALL' or 'Source Division' not in df.columns:
        return df
    return df[df['Source Division'].astype(str) == selected_division].copy()


def buat_division_summary_rows(df):
    rows = [['DIVISION SUMMARY'], ['Division', 'SOS%']]
    if 'Source Division' not in df.columns:
        df_i = df[~is_competitor(df['Produsen'])]
        df_k = df[ is_competitor(df['Produsen'])]
        fi = df_i['Facing'].sum()
        fk = df_k['Facing'].sum()
        total = fi + fk
        rows.append(['ALL', round(fi / total * 100, 1) if total > 0 else 0])
        return rows

    summary = calc_sos(df, ['Source Division']).sort_values('Source Division')
    for _, row in summary.iterrows():
        rows.append([str(row['Source Division']), float(row['SOS%'])])
    return rows


def clear_dashboard_content(ws):
    if hasattr(ws, 'batch_clear'):
        api_retry(ws.batch_clear, ['A2:ZZZ'])
    else:
        api_retry(ws.clear)


def terapkan_dropdown_division(ws, options):
    try:
        requests = [{
            'setDataValidation': {
                'range': {
                    'sheetId': ws.id,
                    'startRowIndex': 0,
                    'endRowIndex': 1,
                    'startColumnIndex': 1,
                    'endColumnIndex': 2,
                },
                'rule': {
                    'condition': {
                        'type': 'ONE_OF_LIST',
                        'values': [{'userEnteredValue': opt} for opt in options],
                    },
                    'strict': True,
                    'showCustomUi': True,
                },
            }
        }]
        ws.spreadsheet.batch_update({'requests': requests})
    except Exception as e:
        print(f'Division dropdown error (diabaikan): {e}')


def extract_parent_brand(brand_name, is_competitor=False):
    """Ekstrak nama brand induk dari nama brand detail.
    Competitor brands tetap pakai nama lengkap.
    Contoh Indofood: 'INDOMIE SOTO MIE-SM' -> 'INDOMIE'
    """
    brand = str(brand_name).strip().upper()
    if is_competitor:
        return brand
    KNOWN_PARENTS = ['SARIMI GELAS', 'CAP 3 AYAM', 'POP MIE', 'IND MIE']
    for parent in KNOWN_PARENTS:
        if brand.startswith(parent):
            return 'INDOMIE' if parent == 'IND MIE' else parent
    return brand.split()[0] if brand else brand


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


def calc_sos(df, groupby_cols):
    """Hitung SOS% untuk sembarang groupby. Return df dengan fi, fk, total, SOS%."""
    df_i = df[~is_competitor(df['Produsen'])]
    df_k = df[ is_competitor(df['Produsen'])]

    fi = df_i.groupby(groupby_cols)['Facing'].sum().reset_index().rename(columns={'Facing': 'fi'})
    fk = df_k.groupby(groupby_cols)['Facing'].sum().reset_index().rename(columns={'Facing': 'fk'})

    merged = fi.merge(fk, on=groupby_cols, how='outer').fillna(0)
    merged['total'] = merged['fi'] + merged['fk']
    merged['SOS%']  = (
        merged['fi'] / merged['total'].replace(0, float('nan')) * 100
    ).round(1).fillna(0)
    return merged


def hitung_sos(df, groupby_cols):
    """Untuk STORE DETAIL — kolom facing_indofood/kompetitor/SOS_% per Period."""
    g = groupby_cols + ['Period']
    df_i = df[~is_competitor(df['Produsen'])]
    df_k = df[ is_competitor(df['Produsen'])]

    fi = df_i.groupby(g)['Facing'].sum().reset_index().rename(columns={'Facing': 'facing_indofood'})
    fk = df_k.groupby(g)['Facing'].sum().reset_index().rename(columns={'Facing': 'facing_kompetitor'})

    merged = fi.merge(fk, on=g, how='outer').fillna(0)
    merged['total_facing'] = merged['facing_indofood'] + merged['facing_kompetitor']
    merged['SOS_%'] = (
        merged['facing_indofood'] / merged['total_facing'].replace(0, float('nan')) * 100
    ).round(1).fillna(0)
    return merged


def get_target(targets, dim, nama):
    """Cari target SOS% — fallback ke DEFAULT lalu 65."""
    dim_up  = dim.upper()
    nama_up = str(nama).upper()
    return targets.get((dim_up, nama_up),
           targets.get((dim_up, 'DEFAULT'), 65.0))


# ─────────────────────────── FORMATTING ────────────────────────

def hapus_semua_chart(spreadsheet, ws_id):
    """Hapus semua chart/embedded object dari sheet sebelum nulis data baru."""
    try:
        meta = spreadsheet.fetch_sheet_metadata()
        for sheet in meta.get('sheets', []):
            if sheet['properties']['sheetId'] == ws_id:
                charts = sheet.get('charts', [])
                if charts:
                    requests = [
                        {'deleteEmbeddedObject': {'objectId': c['chartId']}}
                        for c in charts
                    ]
                    spreadsheet.batch_update({'requests': requests})
                    print(f'[INFO] {len(charts)} chart dihapus dari sheet.')
                break
    except Exception as e:
        print(f'Hapus chart error (diabaikan): {e}')


def hapus_semua_formatting(ws, total_rows, total_cols):
    try:
        from gspread_formatting import format_cell_range, CellFormat
        last_col = col_letter(total_cols - 1)
        ws.spreadsheet.batch_update({'requests': [{'updateCells': {'range': {'sheetId': ws.id}, 'fields': 'userEnteredFormat'}}]})
    except Exception:
        pass


def terapkan_filter(ws, header_row):
    try:
        ws.spreadsheet.batch_update({'requests': [{
            'setBasicFilter': {'filter': {'range': {
                'sheetId': ws.id, 'startRowIndex': header_row - 1, 'startColumnIndex': 0,
            }}}
        }]})
    except Exception as e:
        print(f'Filter error (diabaikan): {e}')


def api_retry(fn, *args, **kwargs):
    """Panggil fn(*args, **kwargs) dengan retry otomatis kalau kena 429 rate limit."""
    for wait in [0, 30, 60, 90]:
        if wait:
            print(f'[WARNING] Rate limit (429), tunggu {wait}s...')
            time.sleep(wait)
        try:
            return fn(*args, **kwargs)
        except Exception as e:
            if gspread is None or not isinstance(e, gspread.exceptions.APIError):
                raise
            if '429' not in str(e):
                raise
    raise RuntimeError('Gagal setelah 4x retry rate limit.')


def hapus_conditional_format(spreadsheet, ws_id):
    """Hapus semua conditional format rule sekaligus (satu batch)."""
    try:
        meta = spreadsheet.fetch_sheet_metadata()
        for sheet in meta.get('sheets', []):
            if sheet['properties']['sheetId'] == ws_id:
                n = len(sheet.get('conditionalFormats', []))
                if n == 0:
                    return
                requests = [
                    {'deleteConditionalFormatRule': {'sheetId': ws_id, 'index': 0}}
                    for _ in range(n)
                ]
                api_retry(spreadsheet.batch_update, {'requests': requests})
                return
    except Exception:
        pass


def terapkan_conditional_format(spreadsheet, ws_id, data_start_row, data_end_row,
                                 sos_col_indices, target_col_idx):
    """Green jika SOS% >= TARGET, merah jika di bawah. sos_col_indices: list 0-indexed."""
    tl = col_letter(target_col_idx)
    r  = data_start_row

    requests = []
    for col_idx in sos_col_indices:
        sl = col_letter(col_idx)
        rng = {
            'sheetId'         : ws_id,
            'startRowIndex'   : data_start_row - 1,
            'endRowIndex'     : data_end_row,
            'startColumnIndex': col_idx,
            'endColumnIndex'  : col_idx + 1,
        }
        requests += [
            {'addConditionalFormatRule': {'rule': {
                'ranges': [rng],
                'booleanRule': {
                    'condition': {'type': 'CUSTOM_FORMULA',
                                  'values': [{'userEnteredValue': f'={sl}{r}>=${tl}{r}'}]},
                    'format': {'backgroundColor': GREEN_BG}
                }
            }, 'index': 0}},
            {'addConditionalFormatRule': {'rule': {
                'ranges': [rng],
                'booleanRule': {
                    'condition': {'type': 'CUSTOM_FORMULA',
                                  'values': [{'userEnteredValue': f'=ISNUMBER({sl}{r})*({sl}{r}<${tl}{r})'}]},
                    'format': {'backgroundColor': RED_BG}
                }
            }, 'index': 1}},
        ]
    try:
        if requests:
            spreadsheet.batch_update({'requests': requests})
    except Exception as e:
        print(f'Conditional format error: {e}')


# ─────────────────────────── TARGET ─────────────────────────────

def default_target_rows(divisions=None):
    divisions = list(divisions or ['ALL'])
    rows = [['Division', 'Dimension', 'Name', 'Target']]
    for division in divisions:
        rows.extend([[division, dim, 'DEFAULT', DEFAULT_TARGET] for dim in TARGET_DIM_ORDER])
    return rows


def default_targets_dict():
    return {
        (dim, 'DEFAULT'): DEFAULT_TARGET
        for dim in TARGET_DIM_ORDER
    }


def buat_targets_excel_default(path=TARGETS_FILENAME):
    wb = Workbook()
    ws = wb.active
    ws.title = 'TARGETS'

    for row in default_target_rows():
        ws.append(row)

    header_fill = PatternFill('solid', fgColor='1F4E78')
    header_font = Font(bold=True, color='FFFFFF')
    thin = Side(style='thin', color='B7B7B7')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=4):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical='center')

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')

    for cell in ws['D'][1:]:
        cell.number_format = '0.0'

    ws.column_dimensions['A'].width = 24
    ws.column_dimensions['B'].width = 24
    ws.column_dimensions['C'].width = 34
    ws.column_dimensions['D'].width = 14
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:D{ws.max_row}'

    wb.save(path)


def _parse_target_value(value, row_number):
    if value is None or str(value).strip() == '':
        raise ValueError(f'TARGETS row {row_number}: nilai target kosong.')
    try:
        if isinstance(value, str):
            value = value.replace('%', '').replace(',', '.').strip()
        return float(value)
    except ValueError as exc:
        raise ValueError(f'TARGETS row {row_number}: nilai target tidak valid ({value}).') from exc


def _target_rows_to_dict(target_rows):
    if not target_rows:
        raise ValueError('TARGETS tidak valid: tidak ada data target.')

    header = [str(v or '').strip().upper() for v in target_rows[0]]
    if header[:4] == ['DIVISION', 'DIMENSION', 'NAME', 'TARGET']:
        dim_idx, name_idx, target_idx = 1, 2, 3
    elif header[:3] == ['DIMENSION', 'NAME', 'TARGET (%)']:
        dim_idx, name_idx, target_idx = 0, 1, 2
    else:
        raise ValueError('TARGETS tidak valid: header harus Division | Dimension | Name | Target.')

    targets = {}
    for row_number, row in enumerate(target_rows[1:], start=2):
        values = list(row) + [''] * 4
        dim = str(values[dim_idx] or '').strip().upper()
        nama = str(values[name_idx] or '').strip().upper()
        target_value = values[target_idx]

        if not dim and not nama and (target_value is None or str(target_value).strip() == ''):
            continue
        if not dim or not nama:
            raise ValueError(f'TARGETS row {row_number}: Dimension dan Name wajib diisi.')

        targets[(dim, nama)] = _parse_target_value(target_value, row_number)

    if not targets:
        raise ValueError('TARGETS tidak valid: tidak ada data target.')

    return targets


def _build_division_targets(target_rows, division):
    """Build a (dim, name) → value dict scoped to one division.

    Merges ALL-division rows (base) with division-specific rows (override),
    so per-division values take precedence while global defaults still apply.
    """
    if not target_rows:
        return default_targets_dict()
    header = target_rows[0]
    all_data = [r for r in target_rows[1:] if str(r[0] or '').strip().upper() == 'ALL']
    div_data = [r for r in target_rows[1:] if r[0] == division]
    merged = [header] + all_data + div_data
    if len(merged) == 1:
        merged = target_rows
    return _target_rows_to_dict(merged)


def _enrich_targets_with_df_values(df, target_rows):
    """Append-only: tambah baris TARGETS untuk setiap nilai dimensi di df yang belum ada.

    Nilai awal tiap baris baru = DEFAULT target untuk division+dim tersebut.
    Baris yang sudah ada tidak pernah diubah.
    """
    DIM_COL_MAP = [
        ('REGION',           'Region'),
        ('CHANNEL',          'Channel'),
        ('ACCOUNT GELATIK',  'Account'),
        ('CATEGORY CHANNEL', 'Category Channel'),
    ]

    base = target_rows if target_rows else [['Division', 'Dimension', 'Name', 'Target']]

    existing = set()
    for row in base[1:]:
        if len(row) >= 3 and row[0] is not None:
            existing.add((str(row[0]).strip(), str(row[1]).strip().upper(), str(row[2]).strip().upper()))

    def _seed(division, dim):
        for scope in (division, 'ALL'):
            for row in base[1:]:
                if (len(row) >= 4
                        and str(row[0]).strip() == scope
                        and str(row[1]).strip().upper() == dim
                        and str(row[2]).strip().upper() == 'DEFAULT'):
                    try:
                        return float(row[3])
                    except (ValueError, TypeError):
                        pass
        return DEFAULT_TARGET

    new_rows = []
    for division in get_source_divisions(df):
        df_div = filter_dashboard_by_division(df, division)

        for dim, col in DIM_COL_MAP:
            if col not in df_div.columns:
                continue
            seed = _seed(division, dim)
            for val in sorted(str(v).strip().upper() for v in df_div[col].dropna().unique() if str(v).strip()):
                if (division, dim, val) not in existing:
                    new_rows.append([division, dim, val, seed])
                    existing.add((division, dim, val))

        if 'Channel' in df_div.columns and 'Account' in df_div.columns:
            pairs = (df_div[['Channel', 'Account']]
                     .drop_duplicates()
                     .dropna()
                     .sort_values(['Channel', 'Account']))
            seed_ca = _seed(division, 'CHANNEL-ACCOUNT')
            for _, r in pairs.iterrows():
                ch, acc = str(r['Channel']).strip(), str(r['Account']).strip()
                if not ch or not acc:
                    continue
                val = f'{ch} - {acc}'.upper()
                if (division, 'CHANNEL-ACCOUNT', val) not in existing:
                    new_rows.append([division, 'CHANNEL-ACCOUNT', val, seed_ca])
                    existing.add((division, 'CHANNEL-ACCOUNT', val))

        # Add GRAND TOTAL entry for each dimension so users can edit it in TARGETS
        all_dims = [dim for dim, _ in DIM_COL_MAP] + ['CHANNEL-ACCOUNT']
        for dim in all_dims:
            gt_key = (division, dim, 'GRAND TOTAL')
            if gt_key not in existing:
                seed = _seed(division, dim)
                new_rows.append([division, dim, 'GRAND TOTAL', seed])
                existing.add(gt_key)

    if not new_rows:
        return target_rows
    return base + new_rows


def read_targets_from_workbook(path):
    wb = None
    try:
        wb = load_workbook(path, data_only=True)
    except Exception as exc:
        raise ValueError(f'Workbook summary ada tetapi tidak bisa dibaca: {exc}') from exc

    try:
        if 'TARGETS' not in wb.sheetnames:
            raise ValueError('Workbook summary tidak valid: sheet TARGETS tidak ditemukan.')

        ws = wb['TARGETS']
        max_col = max(ws.max_column, 4)
        raw_rows = [
            [ws.cell(row=row_number, column=col).value for col in range(1, max_col + 1)]
            for row_number in range(1, ws.max_row + 1)
        ]

        header = [str(v or '').strip().upper() for v in raw_rows[0]]
        if header[:4] == ['DIVISION', 'DIMENSION', 'NAME', 'TARGET']:
            target_rows = [row[:4] for row in raw_rows]
        elif header[:3] == ['DIMENSION', 'NAME', 'TARGET (%)']:
            target_rows = [['Division', 'Dimension', 'Name', 'Target']]
            for row in raw_rows[1:]:
                values = list(row) + [''] * 3
                target_rows.append(['ALL', values[0], values[1], values[2]])
        else:
            raise ValueError('TARGETS tidak valid: header harus Division | Dimension | Name | Target.')
    finally:
        wb.close()

    return _target_rows_to_dict(target_rows), target_rows


def load_or_create_summary_targets(summary_path):
    if not os.path.exists(summary_path):
        return default_targets_dict(), None
    return read_targets_from_workbook(summary_path)


def baca_target_dari_excel(path=TARGETS_FILENAME):
    targets, _target_rows = read_targets_from_workbook(path)
    return targets


def load_or_create_local_targets(path=TARGETS_FILENAME):
    if not os.path.exists(path):
        buat_targets_excel_default(path)
        return None
    return baca_target_dari_excel(path)


def baca_target_dari_ws_targets(ws_targets):
    """
    Baca target dari sheet TARGETS.
    Format sheet: baris header (Dimension | Name | Target), lalu data.
    Contoh: REGION | BANDUNG | 70
    """
    try:
        rows = ws_targets.get_all_values()
    except Exception:
        return {}

    targets = {}
    for row in rows:
        if len(row) < 3:
            continue
        dim, nama, tgt_str = row[0].strip(), row[1].strip(), row[2].strip()
        if not dim or not nama or not tgt_str:
            continue
        if dim.upper() in ('DIMENSION', 'DIM'):
            continue  # skip header
        try:
            targets[(dim.upper(), nama.upper())] = float(
                tgt_str.replace('%', '').replace(',', '.')
            )
        except ValueError:
            pass
    return targets


def baca_target_dari_dashboard(ws, ws_targets=None):
    """
    Baca target — prioritas: sheet TARGETS (jika ada), lalu kolom B di DASHBOARD.
    """
    if isinstance(ws_targets, dict):
        return dict(ws_targets)

    # Prioritas 1: sheet TARGETS terpisah
    if ws_targets is not None:
        targets = baca_target_dari_ws_targets(ws_targets)
        if targets:
            return targets

    # Prioritas 2: kolom B di DASHBOARD (backward compat)
    try:
        all_vals = ws.get_all_values()
    except Exception:
        return {}

    targets    = {}
    current_dim = None
    skip_count  = 0

    for row in all_vals:
        if not row or not row[0].strip():
            current_dim = None; skip_count = 0
            continue

        cell_a = row[0].strip()

        if cell_a.startswith('SOS% BY '):
            current_dim = cell_a.replace('SOS% BY ', '').strip()
            current_dim = current_dim.replace(' × ', '-').replace('×', '-')
            skip_count  = 2
            continue

        if skip_count > 0:
            skip_count -= 1
            continue

        if current_dim and len(row) >= 2:
            nama    = cell_a.upper()
            tgt_str = row[1].strip()
            if nama and tgt_str:
                try:
                    targets[(current_dim, nama)] = float(
                        tgt_str.replace('%', '').replace(',', '.')
                    )
                except ValueError:
                    pass

    return targets


def inisialisasi_ws_targets(ws_targets, targets_existing, df=None):
    """
    Sinkronkan sheet TARGETS: pertahankan nilai yang sudah ada,
    tambahkan entri yang belum ada (misal CHANNEL-ACCOUNT baru).
    """
    DEFAULT_TARGET = 65

    try:
        existing_rows = ws_targets.get_all_values()
    except Exception:
        existing_rows = []

    # Baca entri yang sudah ada di sheet (beserta nilai targetnya)
    existing_keys = {}  # (DIM, NAMA) -> baris ke-n (0-indexed)
    for i, row in enumerate(existing_rows):
        if len(row) < 2:
            continue
        dim, nama = row[0].strip().upper(), row[1].strip().upper()
        if dim and nama and dim not in ('DIMENSION', 'DIM'):
            existing_keys[(dim, nama)] = i

    # Kumpulkan semua kombinasi dari data nyata
    dim_values = {
        'REGION': [],
        'CHANNEL': [],
        'CATEGORY CHANNEL': [],
        'ACCOUNT GELATIK': [],
        'CHANNEL-ACCOUNT': [],
    }
    if df is not None:
        if 'Region' in df.columns:
            dim_values['REGION'] = sorted(df['Region'].dropna().unique().tolist())
        if 'Channel' in df.columns:
            dim_values['CHANNEL'] = sorted(df['Channel'].dropna().unique().tolist())
        if 'Category Channel' in df.columns:
            dim_values['CATEGORY CHANNEL'] = sorted(df['Category Channel'].dropna().unique().tolist())
        if 'Account' in df.columns:
            dim_values['ACCOUNT GELATIK'] = sorted(df['Account'].dropna().unique().tolist())
        if 'Channel' in df.columns and 'Account' in df.columns:
            pairs = (df[['Channel', 'Account']].dropna()
                     .drop_duplicates()
                     .sort_values(['Channel', 'Account']))
            dim_values['CHANNEL-ACCOUNT'] = [
                f'{r.Channel} - {r.Account}' for r in pairs.itertuples()
            ]

    # Gabungkan: targets_existing dari dashboard + data nyata
    all_needed = {}
    for (dim, nama), val in targets_existing.items():
        all_needed[(dim, nama)] = val
    for dim, names in dim_values.items():
        all_needed.setdefault((dim, 'DEFAULT'), DEFAULT_TARGET)
        for nama in names:
            all_needed.setdefault((dim, nama.upper()), DEFAULT_TARGET)
    if not all_needed:
        for dim in dim_values:
            all_needed[(dim, 'DEFAULT')] = DEFAULT_TARGET

    # Tentukan entri yang perlu ditambahkan (belum ada di sheet)
    new_rows = []
    dim_order = ['REGION', 'CHANNEL', 'ACCOUNT GELATIK', 'CATEGORY CHANNEL', 'CHANNEL-ACCOUNT']
    for dim in dim_order:
        dim_entries = sorted(
            [(nama, val) for (d, nama), val in all_needed.items() if d == dim],
            key=lambda x: (x[0] != 'DEFAULT', x[0])
        )
        for nama, val in dim_entries:
            if (dim, nama) not in existing_keys:
                new_rows.append([dim, nama, val])

    if not existing_rows:
        # Sheet kosong — tulis header + semua entri
        header = [['Dimension', 'Name', 'Target (%)']]
        dim_order_rows = []
        for dim in dim_order:
            dim_entries = sorted(
                [(nama, val) for (d, nama), val in all_needed.items() if d == dim],
                key=lambda x: (x[0] != 'DEFAULT', x[0])
            )
            for nama, val in dim_entries:
                dim_order_rows.append([dim, nama, val])
        try:
            ws_targets.update(range_name='A1', values=header + dim_order_rows)
            print(f'[INFO] Sheet TARGETS diinisialisasi: {len(dim_order_rows)} entri.')
        except Exception as e:
            print(f'[WARNING] Gagal inisialisasi sheet TARGETS: {e}')
    elif new_rows:
        # Sheet sudah ada — append entri yang belum ada
        next_row = len(existing_rows) + 1
        try:
            ws_targets.update(range_name=f'A{next_row}', values=new_rows)
            print(f'[INFO] Sheet TARGETS: ditambahkan {len(new_rows)} entri baru.')
        except Exception as e:
            print(f'[WARNING] Gagal update sheet TARGETS: {e}')


# ─────────────────────────── BANGUN TABEL SOS ──────────────────

def buat_tabel_sos_monthly(df, index_col, dim_label, semua_period, targets):
    """
    Tabel SOS% per bulan dengan kolom Indofood | Kompetitor | Total | SOS% per period.
    [index | TARGET | Jan 25→ | Feb 25→ | ... | AVG | STORE COV. | AKTUAL | %]
                      fi fk tot %   fi fk tot %
    """
    df_i = df[~is_competitor(df['Produsen'])]
    df_k = df[ is_competitor(df['Produsen'])]

    index_cols = index_col if isinstance(index_col, list) else [index_col]
    n_idx = len(index_cols)
    target_col_idx = n_idx

    sos_monthly = calc_sos(df, index_cols + ['Period'])
    ordered_periods = [p for p in semua_period if p in sos_monthly['Period'].unique()]
    n = len(ordered_periods)

    # ── Header ──
    header_labels = ['DIVISION' if c == 'Source Division' else c for c in index_cols]
    header1 = header_labels + ['TARGET']
    header2 = [''] * (n_idx + 1)
    for p in ordered_periods:
        header1 += [p, '', '', '']
        header2 += ['Indofood', 'Kompetitor', 'Total', 'SOS%']

    # ── Precompute fi/fk per (index_col, Period) ──
    group_period_cols = index_cols + ['Period']
    fi_grp = df_i.groupby(group_period_cols)['Facing'].sum()
    fk_grp = df_k.groupby(group_period_cols)['Facing'].sum()

    all_indices = (
        sos_monthly[index_cols]
        .drop_duplicates()
        .sort_values(index_cols)
        .itertuples(index=False, name=None)
    )

    # ── Baris data ──
    data_rows = []
    division_totals = {}
    for idx_tuple in all_indices:
        target_name = idx_tuple[-1]
        target_val  = get_target(targets, dim_label, target_name)
        row         = [str(v) for v in idx_tuple] + [target_val]

        for p in ordered_periods:
            grp_key = idx_tuple + (p,)
            fi  = int(round(fi_grp.get(grp_key, 0)))
            fk  = int(round(fk_grp.get(grp_key, 0)))
            tot = fi + fk
            sos = round(fi / tot * 100, 1) if tot > 0 else 0
            row += [fi, fk, tot, sos]

        data_rows.append(row)
        if index_cols[0] == 'Source Division':
            division_totals.setdefault(idx_tuple[0], []).append(idx_tuple)

    if index_cols[0] == 'Source Division' and len(division_totals) > 1:
        for division in sorted(division_totals):
            sub_row = [f'{division} TOTAL'] + [''] * (n_idx - 1) + [get_target(targets, dim_label, 'GRAND TOTAL')]
            for p in ordered_periods:
                fi = int(round(df_i[(df_i['Source Division'] == division) & (df_i['Period'] == p)]['Facing'].sum()))
                fk = int(round(df_k[(df_k['Source Division'] == division) & (df_k['Period'] == p)]['Facing'].sum()))
                tot = fi + fk
                sos = round(fi / tot * 100, 1) if tot > 0 else 0
                sub_row += [fi, fk, tot, sos]
            data_rows.append(sub_row)

    # ── Grand Total ──
    grand    = ['GRAND TOTAL'] + [''] * (n_idx - 1) + [get_target(targets, dim_label, 'GRAND TOTAL')]
    gt_sos   = []
    for p in ordered_periods:
        fi  = int(round(df_i[df_i['Period'] == p]['Facing'].sum()))
        fk  = int(round(df_k[df_k['Period'] == p]['Facing'].sum()))
        tot = fi + fk
        sos = round(fi / tot * 100, 1) if tot > 0 else 0
        grand += [fi, fk, tot, sos]
        gt_sos.append(sos)

    rows     = [header1, header2] + data_rows + [grand]
    num_cols = len(header1)

    sos_cols = [n_idx + 1 + 4*i + 3 for i in range(n)]

    meta = {
        'num_cols'      : num_cols,
        'sos_col_indices': sos_cols,
        'target_col_idx' : target_col_idx,
    }
    return rows, meta


def buat_tabel_channel_account(df, semua_period, targets):
    """
    Tabel SOS% Channel × Account dengan subtotal per Channel.
    [Channel | Account | TARGET | Jan 25→ | Feb 25→ | ... | AVG | COV | AKTUAL | %]
                                   fi fk tot %  fi fk tot %
    """
    if 'Source Division' in df.columns:
        index_cols = ['Source Division', 'Channel', 'Account']
        subtotal_cols = ['Source Division', 'Channel']
        target_col_idx = len(index_cols)

        sos_monthly = calc_sos(df, index_cols + ['Period'])
        ordered_periods = [p for p in semua_period if p in sos_monthly['Period'].unique()]
        n = len(ordered_periods)

        df_i = df[~is_competitor(df['Produsen'])]
        df_k = df[ is_competitor(df['Produsen'])]

        fi_grp = df_i.groupby(index_cols + ['Period'])['Facing'].sum()
        fk_grp = df_k.groupby(index_cols + ['Period'])['Facing'].sum()

        header1 = ['DIVISION', 'CHANNEL', 'ACCOUNT', 'TARGET']
        header2 = ['', '', '', '']
        for p in ordered_periods:
            header1 += [p, '', '', '']
            header2 += ['Indofood', 'Kompetitor', 'Total', 'SOS%']

        all_indices = (
            sos_monthly[index_cols]
            .drop_duplicates()
            .sort_values(index_cols)
            .itertuples(index=False, name=None)
        )

        group_order = []
        group_rows = {}
        for idx_tuple in all_indices:
            div, ch, acc = idx_tuple
            subtotal_key = (div, ch)
            if subtotal_key not in group_rows:
                group_rows[subtotal_key] = []
                group_order.append(subtotal_key)

            target_val = get_target(targets, 'CHANNEL-ACCOUNT', f'{ch} - {acc}')
            row = [str(div), str(ch), str(acc), target_val]

            for p in ordered_periods:
                grp_key = idx_tuple + (p,)
                fi  = int(round(fi_grp.get(grp_key, 0)))
                fk  = int(round(fk_grp.get(grp_key, 0)))
                tot = fi + fk
                sos = round(fi / tot * 100, 1) if tot > 0 else 0
                row += [fi, fk, tot, sos]

            group_rows[subtotal_key].append(row)

        fi_sub = df_i.groupby(subtotal_cols + ['Period'])['Facing'].sum()
        fk_sub = df_k.groupby(subtotal_cols + ['Period'])['Facing'].sum()

        result_rows = []
        for div, ch in group_order:
            result_rows.extend(group_rows[(div, ch)])

            sub_row = [str(div), f'{ch} TOTAL', '', '']
            for p in ordered_periods:
                sub_key = (div, ch, p)
                fi  = int(round(fi_sub.get(sub_key, 0)))
                fk  = int(round(fk_sub.get(sub_key, 0)))
                tot = fi + fk
                sos = round(fi / tot * 100, 1) if tot > 0 else 0
                sub_row += [fi, fk, tot, sos]

            result_rows.append(sub_row)

        grand = ['GRAND TOTAL', '', '', '']
        for p in ordered_periods:
            fi  = int(round(df_i[df_i['Period'] == p]['Facing'].sum()))
            fk  = int(round(df_k[df_k['Period'] == p]['Facing'].sum()))
            tot = fi + fk
            sos = round(fi / tot * 100, 1) if tot > 0 else 0
            grand += [fi, fk, tot, sos]

        rows = [header1, header2] + result_rows + [grand]
        return rows, {
            'num_cols'       : len(header1),
            'sos_col_indices': [len(index_cols) + 1 + 4*i + 3 for i in range(n)],
            'target_col_idx' : target_col_idx,
        }

    sos_monthly = calc_sos(df, ['Channel', 'Account', 'Period'])
    ordered_periods = [p for p in semua_period if p in sos_monthly['Period'].unique()]
    n = len(ordered_periods)

    df_i = df[~is_competitor(df['Produsen'])]
    df_k = df[ is_competitor(df['Produsen'])]

    # Precompute fi/fk per (Channel, Account, Period)
    fi_grp = df_i.groupby(['Channel', 'Account', 'Period'])['Facing'].sum()
    fk_grp = df_k.groupby(['Channel', 'Account', 'Period'])['Facing'].sum()

    header1 = ['CHANNEL', 'ACCOUNT', 'TARGET']
    header2 = ['', '', '']
    for p in ordered_periods:
        header1 += [p, '', '', '']
        header2 += ['Indofood', 'Kompetitor', 'Total', 'SOS%']

    store_sos = calc_sos(df, ['Channel', 'Account', 'Store Code'])

    try:
        sm_pivot = sos_monthly.pivot_table(
            index=['Channel', 'Account'], columns='Period', values='SOS%',
            aggfunc='mean', fill_value=0)
    except Exception as e:
        print(f'Channel×Account pivot error: {e}')
        num_cols = len(header1)
        return [header1, header2], {
            'num_cols': num_cols,
            'sos_col_indices': [3 + 4*i + 3 for i in range(n)],
            'target_col_idx': 2,
        }

    channels_order = []
    ch_account_map = {}
    for (ch, acc) in list(sm_pivot.index):
        if ch not in ch_account_map:
            ch_account_map[ch] = []
            channels_order.append(ch)

        target_val  = get_target(targets, 'CHANNEL-ACCOUNT', f'{ch} - {acc}')
        row         = [str(ch), str(acc), target_val]

        for p in ordered_periods:
            fi  = int(round(fi_grp.get((ch, acc, p), 0)))
            fk  = int(round(fk_grp.get((ch, acc, p), 0)))
            tot = fi + fk
            sos = round(fi / tot * 100, 1) if tot > 0 else 0
            row += [fi, fk, tot, sos]

        ch_account_map[ch].append(row)

    # ── Susun baris: account rows + subtotal per channel ──
    # Precompute fi/fk per (Channel, Period) untuk subtotal
    fi_ch = df_i.groupby(['Channel', 'Period'])['Facing'].sum()
    fk_ch = df_k.groupby(['Channel', 'Period'])['Facing'].sum()

    result_rows = []
    for ch in channels_order:
        acc_rows = ch_account_map[ch]
        result_rows.extend(acc_rows)

        sub_row  = [f'{ch} TOTAL', '', '']
        for p in ordered_periods:
            fi  = int(round(fi_ch.get((ch, p), 0)))
            fk  = int(round(fk_ch.get((ch, p), 0)))
            tot = fi + fk
            sos = round(fi / tot * 100, 1) if tot > 0 else 0
            sub_row += [fi, fk, tot, sos]

        result_rows.append(sub_row)

    # ── Grand Total ──
    grand  = ['GRAND TOTAL', '', '']
    gt_sos = []
    for p in ordered_periods:
        fi  = int(round(df_i[df_i['Period'] == p]['Facing'].sum()))
        fk  = int(round(df_k[df_k['Period'] == p]['Facing'].sum()))
        tot = fi + fk
        sos = round(fi / tot * 100, 1) if tot > 0 else 0
        grand += [fi, fk, tot, sos]
        gt_sos.append(sos)

    rows     = [header1, header2] + result_rows + [grand]
    num_cols = len(header1)

    sos_cols = [3 + 4*i + 3 for i in range(n)]

    meta = {
        'num_cols'       : num_cols,
        'sos_col_indices': sos_cols,
        'target_col_idx' : 2,
    }
    return rows, meta


def buat_category_divisi_section(df, periods, targets):
    """Bangun tabel SOS% BY CATEGORY BY DIVISI.
    Per brand per period: Facing (brand), Total (kategori), SOS%.
    Indofood & Competitor di baris terpisah.
    Logic Period diadopsi dari progress1.
    """
    is_comp = is_competitor(df['Produsen'])

    df = df.copy()
    if 'Period' not in df.columns:
        df = tambah_period_column(df)
        
    df['Parent Brand'] = df.apply(
        lambda row: extract_parent_brand(row['Brand'], is_comp[row.name]), axis=1
    )

    all_periods = sorted(df['Period'].unique(), key=sort_key_period)

    categories = sorted(str(x) for x in df['Category Channel'].dropna().unique())

    # Header baris 1
    header1 = ['Category by Divisi', 'Brand By Facing'] + all_periods + ['TOTAL']

    # Header baris 2
    header2 = ['', ''] + ['SOS%'] * (len(all_periods) + 1)

    rows = [header1, header2]
    subtotal_row_indices = []
    cat_ranges = []

    for cat in categories:
        df_cat = df[df['Category Channel'] == cat]
        cat_comp = is_comp[df_cat.index]

        # Total facing per period untuk kategori ini
        total_period = df_cat.groupby('Period')['Facing'].sum()

        # Split brands
        brands_indo = sorted(df_cat[~cat_comp]['Parent Brand'].unique())
        brands_comp = sorted(df_cat[cat_comp]['Parent Brand'].unique())

        def make_brand_row(brand_name, df_brand):
            row = [cat, brand_name]
            grand_facing = 0
            grand_total = 0

            for p in all_periods:
                bf = int(round(df_brand[df_brand['Period'] == p]['Facing'].sum()))
                tf = int(round(total_period.get(p, 0)))
                sos = round(bf / tf * 100, 2) if tf > 0 else 0
                row.append(sos)
                grand_facing += bf
                grand_total += tf

            grand_sos = round(grand_facing / grand_total * 100, 2) if grand_total > 0 else 0
            row.append(grand_sos)
            return row

        # Baris per brand Indofood lalu Competitor
        cat_start = len(rows)
        for b in brands_indo:
            rows.append(make_brand_row(b, df_cat[df_cat['Parent Brand'] == b]))
        for b in brands_comp:
            rows.append(make_brand_row(b, df_cat[df_cat['Parent Brand'] == b]))
        cat_end = len(rows)
        
        if cat_end > cat_start:
            cat_ranges.append({'cat': cat, 'start': cat_start, 'end': cat_end})

        # Subtotal rows
        def make_subtotal_row(label, df_sub):
            row = [label, '']
            s_facing = 0
            s_total = 0

            for p in all_periods:
                bf = int(round(df_sub[df_sub['Period'] == p]['Facing'].sum()))
                tf = int(round(total_period.get(p, 0)))
                sos = round(bf / tf * 100, 2) if tf > 0 else 0
                row.append(sos)
                s_facing += bf
                s_total += tf

            s_sos = round(s_facing / s_total * 100, 2) if s_total > 0 else 0
            row.append(s_sos)
            return row

        subtotal_row_indices.append(len(rows))
        rows.append(make_subtotal_row(f'{cat} COMPETITOR Total', df_cat[cat_comp]))
        subtotal_row_indices.append(len(rows))
        rows.append(make_subtotal_row(f'{cat} INDOFOOD Total', df_cat[~cat_comp]))

    # Grand Total
    total_all_period = df.groupby('Period')['Facing'].sum()
    df_indo = df[~is_comp]
    
    grand = ['GRAND TOTAL', '']
    g_facing = 0
    g_total = 0

    for p in all_periods:
        fi = int(round(df_indo[df_indo['Period'] == p]['Facing'].sum()))
        tf = int(round(total_all_period.get(p, 0)))
        sos = round(fi / tf * 100, 2) if tf > 0 else 0
        grand.append(sos)
        g_facing += fi
        g_total += tf

    g_sos = round(g_facing / g_total * 100, 2) if g_total > 0 else 0
    grand.append(g_sos)
    rows.append(grand)

    num_cols = len(header1)
    
    sos_cols = [2 + i for i in range(len(all_periods) + 1)]
    
    meta = {
        'num_cols': num_cols,
        'sos_col_indices': sos_cols,
        'target_col_idx': None,
        'subtotal_rows': subtotal_row_indices,
        'cat_ranges': cat_ranges,
        'periods': all_periods
    }
    
    return rows, meta



def buat_region_divisi_section(df, periods):
    is_comp = is_competitor(df['Produsen'])
    df_i = df[~is_comp]
    
    categories = sorted(str(x) for x in df['Category Channel'].dropna().unique())
    
    ordered_periods = [p for p in periods if p in df['Period'].unique()]
    
    fi_grp = df_i.groupby(['Region', 'Category Channel', 'Period'])['Facing'].sum()
    ft_grp = df.groupby(['Region', 'Category Channel', 'Period'])['Facing'].sum()
    
    fi_nat = df_i.groupby(['Category Channel', 'Period'])['Facing'].sum()
    ft_nat = df.groupby(['Category Channel', 'Period'])['Facing'].sum()
    
    header1 = ['', 'REGION']
    header2 = ['', '']
    for cat in categories:
        header1 += [cat] + [''] * (len(ordered_periods) - 1)
        header2 += ordered_periods
        
    rows = [header1, header2]
    
    regions = sorted([r for r in df['Region'].dropna().unique()])
    
    for r in regions:
        row = ['', str(r).upper()]
        for cat in categories:
            for p in ordered_periods:
                fi = int(round(fi_grp.get((r, cat, p), 0)))
                ft = int(round(ft_grp.get((r, cat, p), 0)))
                sos = round(fi / ft * 100, 2) if ft > 0 else 0
                row.append(f"{sos:.2f}%".replace('.', ','))
        rows.append(row)
        
    gt_row = ['', 'GRAND TOTAL']
    for cat in categories:
        for p in ordered_periods:
            fi = int(round(fi_nat.get((cat, p), 0)))
            ft = int(round(ft_nat.get((cat, p), 0)))
            sos = round(fi / ft * 100, 2) if ft > 0 else 0
            gt_row.append(f"{sos:.2f}%".replace('.', ','))
            
    rows.append(gt_row)
    
    meta = {
        'num_cols': len(header1),
        'sos_col_indices': [],
        'target_col_idx': None
    }
    
    return rows, meta


def buat_account_divisi_section(df, periods):
    is_comp = is_competitor(df['Produsen'])
    df_i = df[~is_comp]
    
    categories = sorted(str(x) for x in df['Category Channel'].dropna().unique())
    
    ordered_periods = [p for p in periods if p in df['Period'].unique()]
    
    fi_grp = df_i.groupby(['Account', 'Category Channel', 'Period'])['Facing'].sum()
    ft_grp = df.groupby(['Account', 'Category Channel', 'Period'])['Facing'].sum()
    
    fi_nat = df_i.groupby(['Category Channel', 'Period'])['Facing'].sum()
    ft_nat = df.groupby(['Category Channel', 'Period'])['Facing'].sum()
    
    header1 = ['', 'ACCOUNT']
    header2 = ['', '']
    for cat in categories:
        header1 += [cat] + [''] * (len(ordered_periods) - 1)
        header2 += ordered_periods
        
    rows = [header1, header2]
    
    accounts = sorted([a for a in df['Account'].dropna().unique()])
    
    for a in accounts:
        row = ['', str(a).upper()]
        for cat in categories:
            for p in ordered_periods:
                fi = int(round(fi_grp.get((a, cat, p), 0)))
                ft = int(round(ft_grp.get((a, cat, p), 0)))
                sos = round(fi / ft * 100, 2) if ft > 0 else 0
                row.append(f"{sos:.2f}%".replace('.', ','))
        rows.append(row)
        
    gt_row = ['', 'GRAND TOTAL']
    for cat in categories:
        for p in ordered_periods:
            fi = int(round(fi_nat.get((cat, p), 0)))
            ft = int(round(ft_nat.get((cat, p), 0)))
            sos = round(fi / ft * 100, 2) if ft > 0 else 0
            gt_row.append(f"{sos:.2f}%".replace('.', ','))
            
    rows.append(gt_row)
    
    meta = {
        'num_cols': len(header1),
        'sos_col_indices': [],
        'target_col_idx': None
    }
    
    return rows, meta


def tambahkan_chart_category_divisi(spreadsheet, ws_id, fmt_section, anchor_row=None):
    requests = []
    
    data_start = fmt_section['data_start']
    cat_ranges = fmt_section.get('cat_ranges', [])
    periods = fmt_section.get('periods', [])
    sos_cols = fmt_section.get('sos_col_indices', [])
    
    if anchor_row is None:
        anchor_row = fmt_section['grand_row'] + 2
    
    for c_idx, cr in enumerate(cat_ranges):
        cat_name = cr['cat']
        start_row = data_start + cr['start'] - 3
        end_row = data_start + cr['end'] - 3
        
        for p_idx, p_name in enumerate(periods):
            sos_col = sos_cols[p_idx]
            
            # anchor column based on p_idx
            anchor_col = p_idx * 8 
            
            chart_req = {
                "addChart": {
                    "chart": {
                        "spec": {
                            "title": f"{cat_name} - {p_name}",
                            "hiddenDimensionStrategy": "SHOW_ALL",
                            "basicChart": {
                                "chartType": "COLUMN",
                                "legendPosition": "NO_LEGEND",
                                "axis": [
                                    {"position": "BOTTOM_AXIS", "title": ""},
                                    {"position": "LEFT_AXIS", "title": "SOS%"}
                                ],
                                "domains": [
                                    {
                                        "domain": {
                                            "sourceRange": {
                                                "sources": [
                                                    {
                                                        "sheetId": ws_id,
                                                        "startRowIndex": start_row,
                                                        "endRowIndex": end_row,
                                                        "startColumnIndex": 1, 
                                                        "endColumnIndex": 2
                                                    }
                                                ]
                                            }
                                        }
                                    }
                                ],
                                "series": [
                                    {
                                        "series": {
                                            "sourceRange": {
                                                "sources": [
                                                    {
                                                        "sheetId": ws_id,
                                                        "startRowIndex": start_row,
                                                        "endRowIndex": end_row,
                                                        "startColumnIndex": sos_col,
                                                        "endColumnIndex": sos_col + 1
                                                    }
                                                ]
                                            }
                                        },
                                        "targetAxis": "LEFT_AXIS",
                                        "dataLabel": {
                                            "type": "DATA",
                                            "textFormat": {
                                                "fontSize": 10,
                                                "bold": True,
                                                "foregroundColorStyle": {
                                                    "rgbColor": {"red": 0.2, "green": 0.2, "blue": 0.2}
                                                }
                                            }
                                        }
                                    }
                                ],
                                "headerCount": 0
                            }
                        },
                        "position": {
                            "overlayPosition": {
                                "anchorCell": {
                                    "sheetId": ws_id,
                                    "rowIndex": anchor_row + c_idx * 24,
                                    "columnIndex": anchor_col
                                },
                                "offsetXPixels": 0,
                                "offsetYPixels": 0,
                                "widthPixels": 750,
                                "heightPixels": 450
                            }
                        }
                    }
                }
            }
            requests.append(chart_req)
            
    if requests:
        try:
            api_retry(spreadsheet.batch_update, {'requests': requests})
            print(f'[INFO] Berhasil menambahkan {len(requests)} barchart CATEGORY BY DIVISI.')
        except Exception as e:
            print(f'Gagal menambahkan chart: {e}')


def tambahkan_chart_division_summary(spreadsheet, ws_id, summary_section, anchor_row):
    data_start = summary_section['data_start']
    data_end = summary_section['data_end']
    if data_end < data_start:
        return

    chart_req = {
        "addChart": {
            "chart": {
                "spec": {
                    "title": "DIVISION SUMMARY",
                    "hiddenDimensionStrategy": "SHOW_ALL",
                    "basicChart": {
                        "chartType": "COLUMN",
                        "legendPosition": "NO_LEGEND",
                        "axis": [
                            {"position": "BOTTOM_AXIS", "title": "Division"},
                            {"position": "LEFT_AXIS", "title": "SOS%"}
                        ],
                        "domains": [
                            {
                                "domain": {
                                    "sourceRange": {
                                        "sources": [
                                            {
                                                "sheetId": ws_id,
                                                "startRowIndex": data_start - 1,
                                                "endRowIndex": data_end,
                                                "startColumnIndex": 0,
                                                "endColumnIndex": 1
                                            }
                                        ]
                                    }
                                }
                            }
                        ],
                        "series": [
                            {
                                "series": {
                                    "sourceRange": {
                                        "sources": [
                                            {
                                                "sheetId": ws_id,
                                                "startRowIndex": data_start - 1,
                                                "endRowIndex": data_end,
                                                "startColumnIndex": 1,
                                                "endColumnIndex": 2
                                            }
                                        ]
                                    }
                                },
                                "targetAxis": "LEFT_AXIS",
                                "dataLabel": {
                                    "type": "DATA",
                                    "textFormat": {
                                        "fontSize": 10,
                                        "bold": True,
                                        "foregroundColorStyle": {
                                            "rgbColor": {"red": 0.2, "green": 0.2, "blue": 0.2}
                                        }
                                    }
                                }
                            }
                        ],
                        "headerCount": 0
                    }
                },
                "position": {
                    "overlayPosition": {
                        "anchorCell": {
                            "sheetId": ws_id,
                            "rowIndex": anchor_row,
                            "columnIndex": 3
                        },
                        "offsetXPixels": 0,
                        "offsetYPixels": 0,
                        "widthPixels": 620,
                        "heightPixels": 360
                    }
                }
            }
        }
    }

    try:
        api_retry(spreadsheet.batch_update, {'requests': [chart_req]})
        print('[INFO] Berhasil menambahkan chart DIVISION SUMMARY.')
    except Exception as e:
        print(f'Gagal menambahkan chart DIVISION SUMMARY: {e}')


def sembunyikan_dashboard_row_ranges(ws, row_ranges):
    requests = []
    for start_row, end_row in row_ranges:
        if end_row < start_row:
            continue
        requests.append({
            'updateDimensionProperties': {
                'range': {
                    'sheetId': ws.id,
                    'dimension': 'ROWS',
                    'startIndex': start_row - 1,
                    'endIndex': end_row,
                },
                'properties': {'hiddenByUser': True},
                'fields': 'hiddenByUser',
            }
        })

    if requests:
        api_retry(ws.spreadsheet.batch_update, {'requests': requests})


# ─────────────────────────── DASHBOARD ─────────────────────────

def build_dashboard_payload(df, targets=None, selected_division='ALL'):
    targets = targets or {}
    division_options = get_dashboard_division_options(df)
    if selected_division not in division_options:
        selected_division = 'ALL'

    df_dashboard = df
    semua_period = sorted(df_dashboard['Period'].unique(), key=sort_key_period)

    single_levels = [
        ('REGION',           'Region',           'REGION'),
        ('CHANNEL',          'Channel',          'CHANNEL'),
        ('ACCOUNT GELATIK',  'Account',          'ACCOUNT GELATIK'),
        ('CATEGORY CHANNEL', 'Category Channel', 'CATEGORY CHANNEL'),
    ]

    all_rows = [['DIVISION', selected_division], ['']]
    division_summary_start = len(all_rows) + 1
    division_summary_rows = buat_division_summary_rows(df)
    all_rows.extend(division_summary_rows)
    division_summary_section = {
        'title_row': division_summary_start,
        'data_start': division_summary_start + 2,
        'data_end': division_summary_start + len(division_summary_rows) - 1,
    }
    all_rows.append([])
    fmt_sections = []
    hidden_row_ranges = []

    for label, col, dim_label in single_levels:
        if col not in df_dashboard.columns or df_dashboard[col].dropna().empty:
            continue

        dashboard_index_col = ['Source Division', col] if 'Source Division' in df_dashboard.columns else col
        table_rows, meta = buat_tabel_sos_monthly(
            df_dashboard, dashboard_index_col, dim_label, semua_period, targets
        )

        title_row = len(all_rows) + 1
        all_rows.append([f'SOS% BY {label}'])
        header1_row = len(all_rows) + 1
        header2_row = len(all_rows) + 2
        data_start = len(all_rows) + 3
        all_rows.extend(table_rows)
        grand_row = len(all_rows)

        fmt_sections.append({
            'label': label,
            'title_row': title_row,
            'header1_row': header1_row,
            'header2_row': header2_row,
            'data_start': data_start,
            'grand_row': grand_row,
            'num_cols': meta['num_cols'],
            'sos_col_indices': meta['sos_col_indices'],
            'target_col_idx': meta['target_col_idx'],
        })
        all_rows.append([])
        all_rows.append([])

    if 'Channel' in df_dashboard.columns and 'Account' in df_dashboard.columns:
        table_rows, meta = buat_tabel_channel_account(df_dashboard, semua_period, targets)

        title_row = len(all_rows) + 1
        all_rows.append(['SOS% BY CHANNEL × ACCOUNT'])
        header1_row = len(all_rows) + 1
        header2_row = len(all_rows) + 2
        data_start = len(all_rows) + 3
        all_rows.extend(table_rows)
        grand_row = len(all_rows)

        fmt_sections.append({
            'label': 'CHANNEL × ACCOUNT',
            'title_row': title_row,
            'header1_row': header1_row,
            'header2_row': header2_row,
            'data_start': data_start,
            'grand_row': grand_row,
            'num_cols': meta['num_cols'],
            'sos_col_indices': meta['sos_col_indices'],
            'target_col_idx': meta['target_col_idx'],
        })
        all_rows.append([])
        all_rows.append([])

    if 'Region' in df_dashboard.columns and 'Category Channel' in df_dashboard.columns:
        table_rows, meta = buat_region_divisi_section(df_dashboard, semua_period)

        title_row = len(all_rows) + 1
        all_rows.append(['', 'SOS% BY REGION x DIVISI'])
        header1_row = len(all_rows) + 1
        header2_row = len(all_rows) + 2
        data_start = len(all_rows) + 3
        all_rows.extend(table_rows)
        grand_row = len(all_rows)

        fmt_sections.append({
            'label': 'REGION x DIVISI',
            'title_row': title_row,
            'header1_row': header1_row,
            'header2_row': header2_row,
            'data_start': data_start,
            'grand_row': grand_row,
            'num_cols': meta['num_cols'],
            'sos_col_indices': meta['sos_col_indices'],
            'target_col_idx': meta['target_col_idx'],
        })
        all_rows.append([])
        all_rows.append([])

    if 'Account' in df_dashboard.columns and 'Category Channel' in df_dashboard.columns:
        table_rows, meta = buat_account_divisi_section(df_dashboard, semua_period)

        title_row = len(all_rows) + 1
        all_rows.append(['', 'SOS% BY ACCOUNT x DIVISI'])
        header1_row = len(all_rows) + 1
        header2_row = len(all_rows) + 2
        data_start = len(all_rows) + 3
        all_rows.extend(table_rows)
        grand_row = len(all_rows)

        fmt_sections.append({
            'label': 'ACCOUNT x DIVISI',
            'title_row': title_row,
            'header1_row': header1_row,
            'header2_row': header2_row,
            'data_start': data_start,
            'grand_row': grand_row,
            'num_cols': meta['num_cols'],
            'sos_col_indices': meta['sos_col_indices'],
            'target_col_idx': meta['target_col_idx'],
        })
        all_rows.append([])
        all_rows.append([])

    if (
        'Source Division' in df_dashboard.columns
        and 'Category Channel' in df_dashboard.columns
        and 'Brand' in df_dashboard.columns
    ):
        for chart_division in [d for d in division_options if d != 'ALL']:
            df_chart = filter_dashboard_by_division(df_dashboard, chart_division)
            if df_chart.empty:
                continue

            table_rows, meta = buat_category_divisi_section(df_chart, semua_period, targets)
            if meta.get('cat_ranges'):
                title_row = len(all_rows) + 1
                all_rows.append([f'CHART SOURCE CATEGORY BY DIVISI - {chart_division}'])
                header1_row = len(all_rows) + 1
                header2_row = len(all_rows) + 2
                data_start = len(all_rows) + 3
                all_rows.extend(table_rows)
                grand_row = len(all_rows)
                hidden_row_ranges.append((title_row, grand_row))

                fmt_sections.append({
                    'label': 'CATEGORY BY DIVISI',
                    'title_row': title_row,
                    'header1_row': header1_row,
                    'header2_row': header2_row,
                    'data_start': data_start,
                    'grand_row': grand_row,
                    'num_cols': meta['num_cols'],
                    'sos_col_indices': meta['sos_col_indices'],
                    'target_col_idx': meta['target_col_idx'],
                    'subtotal_rows': meta['subtotal_rows'],
                    'cat_ranges': meta['cat_ranges'],
                    'periods': meta['periods'],
                    'chart_source': True,
                    'division': chart_division,
                })
                all_rows.append([])

    if not fmt_sections:
        return None

    return {
        'rows': all_rows,
        'fmt_sections': fmt_sections,
        'hidden_row_ranges': hidden_row_ranges,
        'division_options': division_options,
        'selected_division': selected_division,
        'division_summary_section': division_summary_section,
        'total_rows': len(all_rows),
        'total_cols': max(s['num_cols'] for s in fmt_sections),
    }


def _prune_dashboard_row_for_excel(row, keep_indices, sos_source_indices, header1_row, row_number):
    result = []
    for old_idx in keep_indices:
        if row_number == header1_row and old_idx in sos_source_indices and old_idx >= 3:
            result.append(row[old_idx - 3] if old_idx - 3 < len(row) else '')
        else:
            result.append(row[old_idx] if old_idx < len(row) else '')
    return result


def dashboard_payload_sos_only(payload):
    rows = [list(row) for row in payload['rows']]
    fmt_sections = []

    for section in payload['fmt_sections']:
        new_section = dict(section)
        target_col = section.get('target_col_idx')
        sos_cols = section.get('sos_col_indices') or []

        if target_col is not None and sos_cols:
            keep_indices = list(range(target_col + 1)) + sos_cols
            mapping = {old_idx: new_idx for new_idx, old_idx in enumerate(keep_indices)}

            for row_number in range(section['title_row'], section['grand_row'] + 1):
                rows[row_number - 1] = _prune_dashboard_row_for_excel(
                    rows[row_number - 1],
                    keep_indices,
                    set(sos_cols),
                    section['header1_row'],
                    row_number,
                )

            new_section['num_cols'] = len(keep_indices)
            new_section['sos_col_indices'] = [mapping[c] for c in sos_cols]
            new_section['target_col_idx'] = mapping[target_col]

        fmt_sections.append(new_section)

    result = dict(payload)
    result['rows'] = rows
    result['fmt_sections'] = fmt_sections
    result['total_cols'] = max(s['num_cols'] for s in fmt_sections)
    return result


def get_cluster_name(base_dir='.'):
    name = os.path.basename(os.path.abspath(base_dir))
    return name or 'Cluster'


def get_summary_output_path(output_dir='.', cluster_name=None, extension='.xlsx'):
    cluster = cluster_name or get_cluster_name(output_dir)
    return os.path.join(output_dir, f'Summary SOS_{cluster}{extension}')


def get_store_detail_output_path(output_dir='.', cluster_name=None, extension='.xlsx'):
    cluster = cluster_name or get_cluster_name(output_dir)
    return os.path.join(output_dir, f'Store Detail_{cluster}{extension}')


def _excel_color(rgb):
    return rgb.replace('#', '')


def _write_rows(ws, rows):
    for row in rows:
        ws.append(row)


def sanitize_excel_sheet_name(name, used_names=None):
    used_names = used_names or set()
    cleaned = ''.join('_' if ch in r'[]:*?/\\' else ch for ch in str(name)).strip()
    cleaned = cleaned or 'Division'
    base = cleaned[:31]
    candidate = base
    suffix = 1
    while candidate in used_names:
        tail = f'_{suffix}'
        candidate = f'{base[:31 - len(tail)]}{tail}'
        suffix += 1
    used_names.add(candidate)
    return candidate


def get_source_divisions(df):
    if 'Source Division' not in df.columns:
        return ['ALL']
    return sorted(str(v) for v in df['Source Division'].dropna().unique())


def _shift_section_rows(section, row_offset):
    shifted = dict(section)
    for key in ('title_row', 'header1_row', 'header2_row', 'data_start', 'grand_row'):
        if key in shifted:
            shifted[key] = shifted[key] + row_offset
    return shifted


def build_division_excel_payload(df, targets, division, target_rows=None):
    division_targets = _build_division_targets(target_rows, division) if target_rows is not None else targets
    df_division = filter_dashboard_by_division(df, division)
    payload = build_dashboard_payload(df_division, division_targets, selected_division=division)
    if payload is None:
        return None

    payload = dashboard_payload_sos_only(payload)
    rows = [list(row) for row in payload['rows'][2:]]
    fmt_sections = []

    for section in payload['fmt_sections']:
        new_section = _shift_section_rows(section, -2)
        if new_section.get('chart_source'):
            title_row = new_section['title_row']
            rows[title_row - 1] = ['CATEGORY BY DIVISI']
            new_section.pop('chart_source', None)
            new_section.pop('division', None)
        fmt_sections.append(new_section)

    result = dict(payload)
    result['rows'] = rows
    result['fmt_sections'] = fmt_sections
    result['hidden_row_ranges'] = []
    result['division_summary_section'] = {
        key: value - 2
        for key, value in payload['division_summary_section'].items()
    }
    result['total_rows'] = len(rows)
    result['total_cols'] = max(s['num_cols'] for s in fmt_sections)
    return result


def get_excel_conditional_format_ranges(payload):
    ranges = []

    for section in payload.get('fmt_sections', []):
        if section.get('chart_source'):
            continue

        target_col = section.get('target_col_idx')
        sos_cols = section.get('sos_col_indices') or []
        if target_col is None or not sos_cols:
            continue

        excluded_rows = {section['grand_row']}
        excluded_rows.update(
            section['header1_row'] + subtotal_idx
            for subtotal_idx in section.get('subtotal_rows', [])
        )

        row_runs = []
        run_start = None
        for row_number in range(section['data_start'], section['grand_row']):
            if row_number in excluded_rows:
                if run_start is not None:
                    row_runs.append((run_start, row_number - 1))
                    run_start = None
                continue

            if run_start is None:
                run_start = row_number

        if run_start is not None:
            row_runs.append((run_start, section['grand_row'] - 1))

        target_letter = col_letter(target_col)
        subtotal_rows = sorted(excluded_rows - {section['grand_row']})

        for sos_col in sos_cols:
            sos_letter = col_letter(sos_col)
            for start_row, end_row in row_runs:
                ranges.append({
                    'label': section.get('label'),
                    'range': f'{sos_letter}{start_row}:{sos_letter}{end_row}',
                    'start_row': start_row,
                    'end_row': end_row,
                    'sos_col_idx': sos_col,
                    'sos_letter': sos_letter,
                    'target_col_idx': target_col,
                    'target_letter': target_letter,
                    'grand_row': section['grand_row'],
                    'subtotal_rows': subtotal_rows,
                    'chart_source': bool(section.get('chart_source')),
                })

    return ranges


def _format_excel_dashboard(ws, payload):
    fills = {
        'dark': PatternFill('solid', fgColor='215E9E'),
        'med': PatternFill('solid', fgColor='4582B5'),
        'light': PatternFill('solid', fgColor='A3C2E3'),
        'grey': PatternFill('solid', fgColor='D9D9D9'),
        'orange': PatternFill('solid', fgColor='FF9933'),
        'green': PatternFill(fill_type='solid', start_color='FFC4D79B', end_color='FFC4D79B'),
        'red': PatternFill(fill_type='solid', start_color='FFE6B8B7', end_color='FFE6B8B7'),
    }
    white_font = Font(bold=True, color='FFFFFF')
    bold_font = Font(bold=True)
    thin = Side(style='thin', color='B7B7B7')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.freeze_panes = 'A2'
    ws.sheet_view.showGridLines = True

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border
            if isinstance(cell.value, float):
                cell.number_format = '0.0'

    for section in payload['fmt_sections']:
        ec = section['num_cols']
        for col in range(1, ec + 1):
            title_cell = ws.cell(section['title_row'], col)
            title_cell.fill = fills['grey']
            title_cell.font = bold_font

            h1_cell = ws.cell(section['header1_row'], col)
            h1_cell.fill = fills['dark']
            h1_cell.font = white_font

            h2_cell = ws.cell(section['header2_row'], col)
            h2_cell.fill = fills['med']
            h2_cell.font = white_font

            grand_cell = ws.cell(section['grand_row'], col)
            grand_cell.fill = fills['light']
            grand_cell.font = bold_font

        target_col = section.get('target_col_idx')
        if target_col is not None:
            for row in range(section['data_start'], section['grand_row'] + 1):
                cell = ws.cell(row, target_col + 1)
                cell.fill = fills['orange']
                cell.font = bold_font

            for cf_range in get_excel_conditional_format_ranges({'fmt_sections': [section]}):
                sos_letter = cf_range['sos_letter']
                target_letter = cf_range['target_letter']
                start_row = cf_range['start_row']
                ws.conditional_formatting.add(
                    cf_range['range'],
                    FormulaRule(
                        formula=[f'{sos_letter}{start_row}>=${target_letter}{start_row}'],
                        fill=fills['green'],
                    ),
                )
                ws.conditional_formatting.add(
                    cf_range['range'],
                    FormulaRule(
                        formula=[f'{sos_letter}{start_row}<${target_letter}{start_row}'],
                        fill=fills['red'],
                    ),
                )

        for subtotal_idx in section.get('subtotal_rows', []):
            subtotal_row = section['header1_row'] + subtotal_idx
            for col in range(1, ec + 1):
                cell = ws.cell(subtotal_row, col)
                cell.fill = fills['med']
                cell.font = white_font

    for start_row, end_row in payload.get('hidden_row_ranges', []):
        for row_number in range(start_row, end_row + 1):
            ws.row_dimensions[row_number].hidden = True

    max_col_to_format = max(ws.max_column, 150)
    for col in range(1, max_col_to_format + 1):
        ws.column_dimensions[col_letter(col - 1)].width = 15
    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 24


def _add_division_summary_chart(ws, summary_section, anchor_row):
    if summary_section['data_end'] < summary_section['data_start']:
        return

    chart = BarChart()
    chart.type = 'col'
    chart.style = 10
    chart.title = 'DIVISION SUMMARY'
    chart.y_axis.title = 'SOS%'
    chart.x_axis.title = 'Division'
    chart.x_axis.delete = False
    chart.x_axis.tickLblPos = 'low'
    chart.legend = None
    chart.width = 15
    chart.height = 9

    data = Reference(
        ws,
        min_col=2,
        min_row=summary_section['data_start'],
        max_row=summary_section['data_end'],
    )
    categories = Reference(
        ws,
        min_col=1,
        min_row=summary_section['data_start'],
        max_row=summary_section['data_end'],
    )
    chart.dLbls = DataLabelList()
    chart.dLbls.showVal = True
    chart.dLbls.showLegendKey = False
    chart.dLbls.showPercent = False
    chart.dLbls.showCatName = False
    chart.dLbls.showSerName = False
    chart.dLbls.dLblPos = 'outEnd'

    chart.add_data(data, titles_from_data=False)
    chart.set_categories(categories)
    chart.series[0].cat = _excel_text_axis_ref(ws, 1, summary_section['data_start'], summary_section['data_end'])
    ws.add_chart(chart, f'D{anchor_row}')


def _excel_text_axis_ref(ws, col, start_row, end_row):
    sheet_name = quote_sheetname(ws.title)
    col_name = col_letter(col - 1)
    return AxDataSource(strRef=StrRef(f=f'{sheet_name}!${col_name}${start_row}:${col_name}${end_row}'))


def _add_category_by_divisi_charts(ws, fmt_section, anchor_row):
    cat_ranges = fmt_section.get('cat_ranges') or []
    periods = fmt_section.get('periods') or []
    sos_cols = fmt_section.get('sos_col_indices') or []
    if not cat_ranges or not periods or not sos_cols:
        return

    base_col = 4

    # Hitung jumlah item data (brands) maksimum di seluruh kategori untuk lembar ini
    max_num_items = 0
    for category_range in cat_ranges:
        start_row = fmt_section['data_start'] + category_range['start'] - 2
        end_row = fmt_section['data_start'] + category_range['end'] - 3
        if end_row >= start_row:
            num_items = end_row - start_row + 1
            if num_items > max_num_items:
                max_num_items = num_items

    if max_num_items < 1:
        max_num_items = 1

    # Gunakan lebar dan jarak kolom yang seragam untuk semua kategori di sheet ini agar rapi sejajar secara vertikal
    chart_width = 6.0 + (max_num_items * 1.5)
    chart_height = 9.0
    chart_width_cols = int(chart_width / 2.7) + 1

    for category_index, category_range in enumerate(cat_ranges):
        category_name = category_range['cat']
        start_row = fmt_section['data_start'] + category_range['start'] - 2
        end_row = fmt_section['data_start'] + category_range['end'] - 3
        if end_row < start_row:
            continue

        for period_index, period_name in enumerate(periods):
            if period_index >= len(sos_cols):
                continue

            chart = BarChart()
            chart.type = 'col'
            chart.style = 10
            chart.title = f'{category_name} - {period_name}'
            chart.y_axis.title = 'SOS%'
            chart.x_axis.delete = False
            chart.x_axis.tickLblPos = 'low'
            chart.legend = None
            chart.width = chart_width
            chart.height = chart_height

            chart.dLbls = DataLabelList()
            chart.dLbls.showVal = True
            chart.dLbls.showLegendKey = False
            chart.dLbls.showPercent = False
            chart.dLbls.showCatName = False
            chart.dLbls.showSerName = False
            chart.dLbls.dLblPos = 'outEnd'

            data = Reference(
                ws,
                min_col=sos_cols[period_index] + 1,
                min_row=start_row,
                max_row=end_row,
            )
            categories = Reference(
                ws,
                min_col=2,
                min_row=start_row,
                max_row=end_row,
            )
            chart.add_data(data, titles_from_data=False)
            chart.set_categories(categories)
            chart.series[0].cat = _excel_text_axis_ref(ws, 2, start_row, end_row)

            chart_row = anchor_row + category_index * EXCEL_CATEGORY_CHART_ROW_STEP
            chart_col = col_letter(base_col - 1 + period_index * chart_width_cols)
            ws.add_chart(chart, f'{chart_col}{chart_row}')


def _add_dashboard_charts(ws, payload):
    anchor_row = payload['total_rows'] + 3
    _add_division_summary_chart(ws, payload['division_summary_section'], anchor_row)

    category_anchor_row = anchor_row + EXCEL_CATEGORY_CHART_ROW_STEP
    for section in payload['fmt_sections']:
        if section.get('label') == 'CATEGORY BY DIVISI':
            _add_category_by_divisi_charts(ws, section, category_anchor_row)


def _write_targets_sheet(ws, targets, target_rows=None, divisions=None):
    if target_rows is None:
        rows = [['Division', 'Dimension', 'Name', 'Target']]
        divisions = list(divisions or ['ALL'])
        for dim in TARGET_DIM_ORDER:
            entries = sorted(
                [(name, val) for (target_dim, name), val in targets.items() if target_dim == dim],
                key=lambda item: (item[0] != 'DEFAULT', item[0]),
            )
            if not entries:
                entries = [('DEFAULT', DEFAULT_TARGET)]
            for division in divisions:
                for name, val in entries:
                    rows.append([division, dim, name, val])
    else:
        rows = [list(row[:4]) for row in target_rows]

    for row in rows:
        ws.append(row)

    header_fill = PatternFill('solid', fgColor='1F4E78')
    header_font = Font(bold=True, color='FFFFFF')
    thin = Side(style='thin', color='B7B7B7')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=4):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical='center')
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
    for cell in ws['D'][1:]:
        cell.number_format = '0.0'

    ws.column_dimensions['A'].width = 24
    ws.column_dimensions['B'].width = 24
    ws.column_dimensions['C'].width = 34
    ws.column_dimensions['D'].width = 14
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:D{ws.max_row}'


def _write_validation_report_sheet(ws, df_removed):
    """Write a VALIDATION REPORT sheet into the given openpyxl worksheet.

    Layout:
      Section 1 – Summary table:  File Name | Total Rows Removed | Total Facing Lost
      (blank row)
      Section 2 – Detail table:   individual removed rows with metadata columns
    """
    if df_removed is None or df_removed.empty:
        ws.append(['Tidak ada data duplikat yang dihapus.'])
        return

    # ---------- filter valid removed rows ----------
    mask_complete = (
        df_removed['Store Code'].notna() &
        df_removed['Visit Date'].notna() &
        df_removed['Product Code'].notna()
    )
    df_valid_rem = df_removed[mask_complete]

    # ---------- Section 1: Summary per file ----------
    summary_header = ['FILE NAME', 'TOTAL ROWS REMOVED', 'TOTAL FACING LOST']
    ws.append(summary_header)

    files = sorted(df_removed['_source_file'].unique()) if '_source_file' in df_removed.columns else []
    for f in files:
        if '_source_file' in df_valid_rem.columns:
            sub = df_valid_rem[df_valid_rem['_source_file'] == f]
        else:
            sub = df_valid_rem
        rows_removed = len(df_removed[df_removed['_source_file'] == f]) if '_source_file' in df_removed.columns else len(df_removed)
        facing_lost = int(sub['Facing'].fillna(0).sum())
        ws.append([f, rows_removed, facing_lost])

    summary_end_row = 1 + len(files)  # row 1 = header, then 1 row per file

    # ---------- detail header row = summary rows + 2 blank rows + 1 ----------
    detail_header_row = summary_end_row + 3  # 2 blank separator rows then header

    # ---------- Section 2: Detail rows ----------
    detail_cols = ['_source_file', 'CSV Row', 'Region', 'Area', 'Channel', 'Account',
                   'Store Name', 'Store Code', 'Visit Date', 'Product Code', 'Brand', 'Facing']
    avail_cols = [c for c in detail_cols if c in df_valid_rem.columns]

    detail_header = avail_cols.copy()
    if '_source_file' in detail_header:
        detail_header[detail_header.index('_source_file')] = 'FILE NAME'
    if 'Facing' in detail_header:
        detail_header[detail_header.index('Facing')] = 'FACING LOST'
    detail_header = [h.upper() for h in detail_header]

    # Write detail header explicitly at the correct row
    for col_idx, val in enumerate(detail_header, 1):
        ws.cell(row=detail_header_row, column=col_idx, value=val)

    detail_data_df = df_valid_rem[avail_cols].copy()
    for col in detail_data_df.select_dtypes(include=['datetime64[ns]', 'datetime64[ns, UTC]']).columns:
        detail_data_df[col] = detail_data_df[col].dt.strftime('%Y-%m-%d')

    current_row = detail_header_row + 1
    for row_data in detail_data_df.fillna('').itertuples(index=False, name=None):
        for col_idx, val in enumerate(list(row_data), 1):
            ws.cell(row=current_row, column=col_idx, value=val)
        current_row += 1

    last_data_row = current_row - 1 if current_row > detail_header_row + 1 else detail_header_row

    # ---------- Formatting ----------
    header_fill = PatternFill('solid', fgColor='215E9E')
    header_font = Font(bold=True, color='FFFFFF')
    thin = Side(style='thin', color='B7B7B7')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Format summary header (row 1)
    for col_idx in range(1, 4):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # Format detail header row (same fill as summary header)
    for col_idx in range(1, len(detail_header) + 1):
        cell = ws.cell(row=detail_header_row, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # Apply borders and center to summary data rows
    for r_idx in range(1, summary_end_row + 1):
        for col_idx in range(1, 4):
            cell = ws.cell(row=r_idx, column=col_idx)
            cell.border = border
            cell.alignment = Alignment(horizontal='center', vertical='center')

    # Apply borders and center to detail rows (header + data)
    for r_idx in range(detail_header_row, last_data_row + 1):
        for col_idx in range(1, len(detail_header) + 1):
            cell = ws.cell(row=r_idx, column=col_idx)
            cell.border = border
            cell.alignment = Alignment(horizontal='center', vertical='center')

    # Column widths
    col_widths = {
        'A': 40, 'B': 10, 'C': 20, 'D': 20, 'E': 18, 'F': 18,
        'G': 25, 'H': 16, 'I': 14, 'J': 18, 'K': 22, 'L': 14,
    }
    for col_letter_key, w in col_widths.items():
        if col_letter_key in ws.column_dimensions:
            ws.column_dimensions[col_letter_key].width = w
        else:
            ws.column_dimensions[col_letter_key].width = w # Still assign just in case

    # Filter on detail header (FILE NAME only)
    ws.auto_filter.ref = f'A{detail_header_row}:A{last_data_row}'


def export_summary_excel(df, targets, output_dir='.', cluster_name=None, target_rows=None, df_removed=None):
    divisions = get_source_divisions(df)
    if not divisions:
        raise ValueError('Tidak ada data untuk dashboard.')

    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    used_sheet_names = set()
    wrote_dashboard = False
    for division in divisions:
        excel_payload = build_division_excel_payload(df, targets, division, target_rows=target_rows)
        if excel_payload is None:
            continue

        ws_dashboard = wb.create_sheet(sanitize_excel_sheet_name(division, used_sheet_names))
        _write_rows(ws_dashboard, excel_payload['rows'])
        _format_excel_dashboard(ws_dashboard, excel_payload)
        _add_dashboard_charts(ws_dashboard, excel_payload)
        wrote_dashboard = True

    if not wrote_dashboard:
        raise ValueError('Tidak ada data untuk dashboard.')

    # Enrich targets to ensure GRAND TOTAL entries exist for every division × dimension
    target_rows = _enrich_targets_with_df_values(df, target_rows)
    if target_rows is not None:
        targets = _target_rows_to_dict(target_rows)

    ws_targets = wb.create_sheet('TARGETS')
    _write_targets_sheet(ws_targets, targets, target_rows=target_rows, divisions=divisions)

    # VALIDATION REPORT sheet
    ws_validation = wb.create_sheet('VALIDATION REPORT')
    _write_validation_report_sheet(ws_validation, df_removed)

    output_path = get_summary_output_path(output_dir, cluster_name, extension='.xlsx')
    wb.save(output_path)
    global _last_excel_write_time
    _last_excel_write_time = time.time()
    return output_path

def export_store_detail_excel(df, output_dir='.', cluster_name=None):
    divisions = get_source_divisions(df)
    if not divisions:
        raise ValueError('Tidak ada data untuk Store Detail.')

    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    used_sheet_names = set()
    wrote_any = False

    semua_period = sorted(df['Period'].unique(), key=sort_key_period)
    index_cols = ['Region', 'Area', 'Channel', 'Account', 'Store Code', 'Store Name']
    index_cols = [c for c in index_cols if c in df.columns]

    df_i = df[~is_competitor(df['Produsen'])]
    df_k = df[ is_competitor(df['Produsen'])]

    merge_cols = ['Region', 'Store Code', 'Period', 'Source Division']
    merge_cols_present = [c for c in merge_cols if c in df.columns]

    fi_grp = df_i.groupby(merge_cols_present)['Facing'].sum()
    fk_grp = df_k.groupby(merge_cols_present)['Facing'].sum()
    valid_keys = set(df[merge_cols_present].drop_duplicates().itertuples(index=False, name=None))

    for division in divisions:
        df_div = df[df['Source Division'] == division]
        if df_div.empty:
            continue

        sheet_name = sanitize_excel_sheet_name(division, used_sheet_names)
        ws = wb.create_sheet(sheet_name)

        ordered_periods = [p for p in semua_period if p in df_div['Period'].unique()]

        header1 = [c.upper() for c in index_cols] + ordered_periods
        header2 = [''] * len(index_cols) + ['SOS%'] * len(ordered_periods)

        _write_rows(ws, [header1, header2])

        # Urutkan berdasarkan Region dan Store Code agar posisi baris yang pecah bisa saling berdekatan
        sort_cols = [c for c in ['Region', 'Store Code'] if c in index_cols]
        sort_cols += [c for c in index_cols if c not in sort_cols]

        all_stores = (
            df_div[index_cols]
            .drop_duplicates()
            .sort_values(sort_cols)
            .itertuples(index=False, name=None)
        )

        data_rows = []
        for store_tuple in all_stores:
            row = list(store_tuple)
            store_dict = dict(zip(index_cols, store_tuple))
            
            for p in ordered_periods:
                merge_key_list = []
                for c in merge_cols_present:
                    if c == 'Period':
                        merge_key_list.append(p)
                    elif c == 'Source Division':
                        merge_key_list.append(division)
                    else:
                        merge_key_list.append(store_dict[c])
                merge_key = tuple(merge_key_list)
                
                if merge_key in valid_keys:
                    fi = int(round(fi_grp.get(merge_key, 0)))
                    fk = int(round(fk_grp.get(merge_key, 0)))
                    tot = fi + fk
                    if tot > 0:
                        sos = (fi / tot)
                    else:
                        sos = 0
                else:
                    sos = ''
                row.append(sos)
            data_rows.append(row)

        _write_rows(ws, data_rows)

        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 25
        ws.column_dimensions['E'].width = 15
        ws.column_dimensions['F'].width = 30
        for i in range(len(ordered_periods)):
            col_letter = get_column_letter(len(index_cols) + 1 + i)
            ws.column_dimensions[col_letter].width = 12

        header1_fill = PatternFill('solid', fgColor='215E9E')
        header1_font = Font(bold=True, color='FFFFFF')
        
        ws.freeze_panes = 'G3'
        ws.auto_filter.ref = f'A1:{get_column_letter(len(index_cols))}{len(data_rows) + 2}'
        
        header2_fill = PatternFill('solid', fgColor='4582B5')
        header2_font = Font(bold=True, color='FFFFFF')

        for row_cells in ws.iter_rows(min_row=1, max_row=1, min_col=1, max_col=len(header1)):
            for cell in row_cells:
                cell.font = header1_font
                cell.fill = header1_fill
                cell.alignment = Alignment(horizontal='center', vertical='center')
                
        for row_cells in ws.iter_rows(min_row=2, max_row=2, min_col=1, max_col=len(header2)):
            for cell in row_cells:
                cell.font = header2_font
                cell.fill = header2_fill
                cell.alignment = Alignment(horizontal='center', vertical='center')

        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                             top=Side(style='thin'), bottom=Side(style='thin'))
        for r_idx, row_cells in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column), start=1):
            for c_idx, cell in enumerate(row_cells, start=1):
                cell.border = thin_border
                
                # Format SOS values as percentage
                if r_idx > 2 and c_idx > len(index_cols):
                    cell.number_format = '0.00%'

        wrote_any = True

    if not wrote_any:
        raise ValueError('Tidak ada data untuk Store Detail.')

    output_path = get_store_detail_output_path(output_dir, cluster_name, extension='.xlsx')
    wb.save(output_path)
    return output_path


def buat_dashboard(ws, df, ws_targets=None):
    targets = baca_target_dari_dashboard(ws, ws_targets)
    division_options = get_dashboard_division_options(df)
    selected_division = baca_dashboard_division_selector(ws, division_options)
    df_dashboard = df

    hapus_semua_chart(ws.spreadsheet, ws.id)
    clear_dashboard_content(ws)

    semua_period = sorted(df_dashboard['Period'].unique(), key=sort_key_period)

    # Level tunggal: level_label, kolom_di_df, dim_key_untuk_targets
    single_levels = [
        ('REGION',           'Region',           'REGION'),
        ('CHANNEL',          'Channel',          'CHANNEL'),
        ('ACCOUNT GELATIK',  'Account',          'ACCOUNT GELATIK'),
        ('CATEGORY CHANNEL', 'Category Channel', 'CATEGORY CHANNEL'),
    ]

    all_rows     = [['DIVISION', selected_division], ['']]
    division_summary_start = len(all_rows) + 1
    division_summary_rows = buat_division_summary_rows(df)
    all_rows.extend(division_summary_rows)
    division_summary_section = {
        'title_row': division_summary_start,
        'data_start': division_summary_start + 2,
        'data_end': division_summary_start + len(division_summary_rows) - 1,
    }
    all_rows.append([])
    fmt_sections = []
    hidden_row_ranges = []

    # ── Section: single-column levels ──
    for label, col, dim_label in single_levels:
        if col not in df_dashboard.columns or df_dashboard[col].dropna().empty:
            continue

        dashboard_index_col = ['Source Division', col] if 'Source Division' in df_dashboard.columns else col
        table_rows, meta = buat_tabel_sos_monthly(
            df_dashboard, dashboard_index_col, dim_label, semua_period, targets
        )

        title_row   = len(all_rows) + 1
        all_rows.append([f'SOS% BY {label}'])
        header1_row = len(all_rows) + 1
        header2_row = len(all_rows) + 2
        data_start  = len(all_rows) + 3
        all_rows.extend(table_rows)
        grand_row   = len(all_rows)

        fmt_sections.append({
            'label'          : label,
            'title_row'      : title_row,
            'header1_row'    : header1_row,
            'header2_row'    : header2_row,
            'data_start'     : data_start,
            'grand_row'      : grand_row,
            'num_cols'       : meta['num_cols'],
            'sos_col_indices': meta['sos_col_indices'],
            'target_col_idx' : meta['target_col_idx'],
        })
        all_rows.append([]); all_rows.append([])

    # ── Section: CHANNEL × ACCOUNT ──
    if 'Channel' in df_dashboard.columns and 'Account' in df_dashboard.columns:
        table_rows, meta = buat_tabel_channel_account(df_dashboard, semua_period, targets)

        title_row   = len(all_rows) + 1
        all_rows.append(['SOS% BY CHANNEL × ACCOUNT'])
        header1_row = len(all_rows) + 1
        header2_row = len(all_rows) + 2
        data_start  = len(all_rows) + 3
        all_rows.extend(table_rows)
        grand_row   = len(all_rows)

        fmt_sections.append({
            'label'          : 'CHANNEL × ACCOUNT',
            'title_row'      : title_row,
            'header1_row'    : header1_row,
            'header2_row'    : header2_row,
            'data_start'     : data_start,
            'grand_row'      : grand_row,
            'num_cols'       : meta['num_cols'],
            'sos_col_indices': meta['sos_col_indices'],
            'target_col_idx' : meta['target_col_idx'],
        })
        all_rows.append([]); all_rows.append([])

    # ── Section: REGION x DIVISI ──
    if 'Region' in df_dashboard.columns and 'Category Channel' in df_dashboard.columns:
        table_rows, meta = buat_region_divisi_section(df_dashboard, semua_period)

        title_row   = len(all_rows) + 1
        all_rows.append(['', 'SOS% BY REGION x DIVISI'])
        header1_row = len(all_rows) + 1
        header2_row = len(all_rows) + 2
        data_start  = len(all_rows) + 3
        all_rows.extend(table_rows)
        grand_row   = len(all_rows)

        fmt_sections.append({
            'label'          : 'REGION x DIVISI',
            'title_row'      : title_row,
            'header1_row'    : header1_row,
            'header2_row'    : header2_row,
            'data_start'     : data_start,
            'grand_row'      : grand_row,
            'num_cols'       : meta['num_cols'],
            'sos_col_indices': meta['sos_col_indices'],
            'target_col_idx' : meta['target_col_idx'],
        })
        all_rows.append([]); all_rows.append([])

    # ── Section: ACCOUNT x DIVISI ──
    if 'Account' in df_dashboard.columns and 'Category Channel' in df_dashboard.columns:
        table_rows, meta = buat_account_divisi_section(df_dashboard, semua_period)

        title_row   = len(all_rows) + 1
        all_rows.append(['', 'SOS% BY ACCOUNT x DIVISI'])
        header1_row = len(all_rows) + 1
        header2_row = len(all_rows) + 2
        data_start  = len(all_rows) + 3
        all_rows.extend(table_rows)
        grand_row   = len(all_rows)

        fmt_sections.append({
            'label'          : 'ACCOUNT x DIVISI',
            'title_row'      : title_row,
            'header1_row'    : header1_row,
            'header2_row'    : header2_row,
            'data_start'     : data_start,
            'grand_row'      : grand_row,
            'num_cols'       : meta['num_cols'],
            'sos_col_indices': meta['sos_col_indices'],
            'target_col_idx' : meta['target_col_idx'],
        })
        all_rows.append([]); all_rows.append([])

    # Hidden chart source: CATEGORY BY DIVISI for selected division.
    if (
        'Source Division' in df_dashboard.columns
        and 'Category Channel' in df_dashboard.columns
        and 'Brand' in df_dashboard.columns
    ):
        for chart_division in [d for d in division_options if d != 'ALL']:
            df_chart = filter_dashboard_by_division(df_dashboard, chart_division)
            if df_chart.empty:
                continue

            table_rows, meta = buat_category_divisi_section(df_chart, semua_period, targets)
            if meta.get('cat_ranges'):
                title_row = len(all_rows) + 1
                all_rows.append([f'CHART SOURCE CATEGORY BY DIVISI - {chart_division}'])
                header1_row = len(all_rows) + 1
                header2_row = len(all_rows) + 2
                data_start = len(all_rows) + 3
                all_rows.extend(table_rows)
                grand_row = len(all_rows)
                hidden_row_ranges.append((title_row, grand_row))

                fmt_sections.append({
                    'label'          : 'CATEGORY BY DIVISI',
                    'title_row'      : title_row,
                    'header1_row'    : header1_row,
                    'header2_row'    : header2_row,
                    'data_start'     : data_start,
                    'grand_row'      : grand_row,
                    'num_cols'       : meta['num_cols'],
                    'sos_col_indices': meta['sos_col_indices'],
                    'target_col_idx' : meta['target_col_idx'],
                    'subtotal_rows'  : meta['subtotal_rows'],
                    'cat_ranges'     : meta['cat_ranges'],
                    'periods'        : meta['periods'],
                    'chart_source'   : True,
                    'division'       : chart_division,
                })
                all_rows.append([])

    # ── Section: CATEGORY BY DIVISI ──
    if not fmt_sections:
        print('[WARNING] Tidak ada data untuk dashboard.')
        return

    total_rows = len(all_rows)
    total_cols = max(s['num_cols'] for s in fmt_sections)

    hapus_semua_formatting(ws, total_rows + 10, total_cols + 5)
    time.sleep(1)
    api_retry(ws.update, range_name='A1', values=all_rows)
    terapkan_dropdown_division(ws, division_options)
    sembunyikan_dashboard_row_ranges(ws, hidden_row_ranges)
    time.sleep(1)

    # Reset freeze
    api_retry(ws.spreadsheet.batch_update, {'requests': [{'updateSheetProperties': {
        'properties': {'sheetId': ws.id,
                       'gridProperties': {'frozenRowCount': 0, 'frozenColumnCount': 0}},
        'fields': 'gridProperties.frozenRowCount,gridProperties.frozenColumnCount'
    }}]})
    time.sleep(1)

    hapus_conditional_format(ws.spreadsheet, ws.id)
    time.sleep(1)

    safe_anchor_row = total_rows + 2

    if selected_division == 'ALL':
        try:
            tambahkan_chart_division_summary(ws.spreadsheet, ws.id, division_summary_section, safe_anchor_row)
            time.sleep(1)
        except Exception as e:
            print(f'Chart division summary error (diabaikan): {e}')
    else:
        for s in fmt_sections:
            if s['label'] == 'CATEGORY BY DIVISI' and s.get('division') == selected_division:
                try:
                    tambahkan_chart_category_divisi(ws.spreadsheet, ws.id, s, safe_anchor_row)
                    time.sleep(1)
                except Exception as e:
                    print(f'Chart category divisi error (diabaikan): {e}')

    try:
        from gspread_formatting import format_cell_ranges, CellFormat, Color, TextFormat

        BLUE_DARK  = Color(0.13, 0.37, 0.62)
        BLUE_MED   = Color(0.27, 0.51, 0.71)
        BLUE_LIGHT = Color(0.64, 0.76, 0.89)
        GREY       = Color(0.85, 0.85, 0.85)
        ORANGE     = Color(1.0, 0.60, 0.20)
        WHITE      = Color(1, 1, 1)

        cell_fmt = []
        for s in fmt_sections:
            ec = col_letter(s['num_cols'] - 1)
            tr, h1, h2, gr, ds = s['title_row'], s['header1_row'], s['header2_row'], s['grand_row'], s['data_start']

            cell_fmt += [
                (f'A{tr}:{ec}{tr}', CellFormat(backgroundColor=GREY,
                    textFormat=TextFormat(bold=True, fontSize=11))),
                (f'A{h1}:{ec}{h1}', CellFormat(backgroundColor=BLUE_DARK,
                    textFormat=TextFormat(bold=True, foregroundColor=WHITE))),
                (f'A{h2}:{ec}{h2}', CellFormat(backgroundColor=BLUE_MED,
                    textFormat=TextFormat(bold=True, foregroundColor=WHITE))),
                (f'A{gr}:{ec}{gr}', CellFormat(backgroundColor=BLUE_LIGHT,
                    textFormat=TextFormat(bold=True))),
            ]
            if s.get('target_col_idx') is not None:
                tgt_col = col_letter(s['target_col_idx'])
                cell_fmt.append((f'{tgt_col}{ds}:{tgt_col}{gr}', CellFormat(backgroundColor=ORANGE, textFormat=TextFormat(bold=True))))

            if 'subtotal_rows' in s:
                for sr in s['subtotal_rows']:
                    subtotal_row = s['header1_row'] + sr
                    cell_fmt.append(
                        (f'A{subtotal_row}:{ec}{subtotal_row}', CellFormat(backgroundColor=BLUE_MED,
                            textFormat=TextFormat(bold=True, foregroundColor=WHITE)))
                    )

        api_retry(format_cell_ranges, ws, cell_fmt)
        time.sleep(1)

        # Conditional formatting per section (hanya kolom SOS%)
        for s in fmt_sections:
            if s.get('sos_col_indices') and s.get('target_col_idx') is not None:
                terapkan_conditional_format(
                    ws.spreadsheet, ws.id,
                    s['data_start'], s['grand_row'],
                    s['sos_col_indices'],
                    s['target_col_idx'],
                )
                time.sleep(1)

        print('Formatting dashboard berhasil!')

    except ImportError:
        print('Install: pip install gspread-formatting')
    except Exception as e:
        print(f'Formatting error (diabaikan): {e}')


# ─────────────────────────── STORE DETAIL ──────────────────────

def buat_store_detail(ws, df):
    api_retry(ws.clear)

    semua_period = sorted(df['Period'].unique(), key=sort_key_period)
    index_cols   = ['Region', 'Area', 'Store Name']

    sos_data   = hitung_sos(df, index_cols)
    store_data = df.groupby(index_cols + ['Period'])['Store Code'].nunique().reset_index()
    store_data.rename(columns={'Store Code': 'store_count'}, inplace=True)

    ordered_periods = [p for p in semua_period if p in sos_data['Period'].unique()]

    def make_pivot(data, val, agg):
        p = data.pivot_table(index=index_cols, columns='Period', values=val, aggfunc=agg, fill_value=0)
        return p.reindex(columns=ordered_periods, fill_value=0)

    pv_fi  = make_pivot(sos_data,   'facing_indofood',   'sum')
    pv_fk  = make_pivot(sos_data,   'facing_kompetitor', 'sum')
    pv_ft  = make_pivot(sos_data,   'total_facing',      'sum')
    pv_sos = make_pivot(sos_data,   'SOS_%',             'mean')
    pv_sc  = make_pivot(store_data, 'store_count',       'sum')

    n_idx   = len(index_cols)
    header1 = index_cols[:]
    header2 = [''] * n_idx
    for p in ordered_periods:
        header1 += [p] + [''] * (N_METRICS - 1)
        header2 += METRIC_LABELS
    header1 += ['TOTAL'] + [''] * (N_METRICS - 1)
    header2 += METRIC_LABELS

    table_rows = [header1, header2]

    for idx in pv_fi.index:
        row    = list(idx)
        tot_fi = tot_fk = tot_ft = tot_sc = 0

        for p in ordered_periods:
            fi     = int(round(float(pv_fi.loc[idx, p])))
            fk     = int(round(float(pv_fk.loc[idx, p])))
            ft     = int(round(float(pv_ft.loc[idx, p])))
            sos    = round(float(pv_sos.loc[idx, p]), 1)
            sc_val = int(pv_sc.loc[idx, p]) if idx in pv_sc.index else 0
            row   += [fi, fk, ft, sos, sc_val]
            tot_fi += fi; tot_fk += fk; tot_ft += ft; tot_sc += sc_val

        tot_sos = round(tot_fi / tot_ft * 100, 1) if tot_ft > 0 else 0
        row += [tot_fi, tot_fk, tot_ft, tot_sos, tot_sc]
        table_rows.append(row)

    grand  = ['GRAND TOTAL', '', '']
    gt_fi  = gt_fk = gt_ft = gt_sc = 0
    for p in ordered_periods:
        fi     = int(round(pv_fi[p].sum()))
        fk     = int(round(pv_fk[p].sum()))
        ft     = int(round(pv_ft[p].sum()))
        sos    = round(fi / ft * 100, 1) if ft > 0 else 0
        sc_val = int(pv_sc[p].sum()) if p in pv_sc.columns else 0
        grand += [fi, fk, ft, sos, sc_val]
        gt_fi += fi; gt_fk += fk; gt_ft += ft; gt_sc += sc_val
    gt_sos = round(gt_fi / gt_ft * 100, 1) if gt_ft > 0 else 0
    grand += [gt_fi, gt_fk, gt_ft, gt_sos, gt_sc]
    table_rows.append(grand)

    num_cols = len(header1)
    hapus_semua_formatting(ws, len(table_rows) + 10, num_cols + 5)
    ws.update(range_name='A1', values=table_rows)
    ws.freeze(rows=2, cols=3)
    terapkan_filter(ws, 1)

    try:
        from gspread_formatting import format_cell_ranges, CellFormat, Color, TextFormat
        BLUE_DARK = Color(0.13, 0.37, 0.62); BLUE_MED = Color(0.27, 0.51, 0.71)
        BLUE_LIGHT = Color(0.64, 0.76, 0.89); WHITE = Color(1, 1, 1)
        ec = col_letter(num_cols - 1); gr = len(table_rows)
        format_cell_ranges(ws, [
            (f'A1:{ec}1', CellFormat(backgroundColor=BLUE_DARK,
                textFormat=TextFormat(bold=True, foregroundColor=WHITE))),
            (f'A2:{ec}2', CellFormat(backgroundColor=BLUE_MED,
                textFormat=TextFormat(bold=True, foregroundColor=WHITE))),
            (f'A{gr}:{ec}{gr}', CellFormat(backgroundColor=BLUE_LIGHT,
                textFormat=TextFormat(bold=True))),
        ])
    except Exception as e:
        print(f'Formatting store error: {e}')

    print('Sheet STORE DETAIL berhasil diupdate!')


# ─────────────────────────── VALIDASI & BACA CSV ───────────────

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


# ─────────────────────────── VALIDATION REPORT ───────────────────

def buat_validation_report(ws, df_removed):
    api_retry(ws.clear)

    mask_complete = (
        df_removed['Store Code'].notna() &
        df_removed['Visit Date'].notna() &
        df_removed['Product Code'].notna()
    )
    df_valid_rem = df_removed[mask_complete]

    summary_rows = []
    header_sum = ['File Name', 'Total Rows Removed', 'Total Facing Lost']
    summary_rows.append(header_sum)

    files = df_removed['_source_file'].unique() if '_source_file' in df_removed.columns else []
    for f in sorted(files):
        sub = df_valid_rem[df_valid_rem['_source_file'] == f] if '_source_file' in df_valid_rem.columns else df_valid_rem
        rows_removed = len(df_removed[df_removed['_source_file'] == f]) if '_source_file' in df_removed.columns else len(df_removed)
        facing_lost = sub['Facing'].fillna(0).sum()
        summary_rows.append([f, rows_removed, int(facing_lost)])

    summary_rows.append([])

    detail_cols = ['_source_file', 'Region', 'Area', 'Channel', 'Account',
                   'Store Name', 'Store Code', 'Visit Date', 'Product Code', 'Brand', 'Facing']
    avail_cols = [c for c in detail_cols if c in df_valid_rem.columns]

    detail_header = avail_cols.copy()
    if '_source_file' in detail_header:
        detail_header[detail_header.index('_source_file')] = 'File Name'
    if 'Facing' in detail_header:
        detail_header[detail_header.index('Facing')] = 'Facing Lost'
    summary_rows.append(detail_header)

    detail_data_df = df_valid_rem[avail_cols].copy()
    for col in detail_data_df.select_dtypes(include=['datetime64[ns]', 'datetime64[ns, UTC]']).columns:
        detail_data_df[col] = detail_data_df[col].dt.strftime('%Y-%m-%d')

    summary_rows.extend(detail_data_df.fillna('').values.tolist())

    api_retry(ws.update, range_name='A1', values=summary_rows)

    try:
        from gspread_formatting import format_cell_ranges, CellFormat, Color, TextFormat
        BLUE_DARK = Color(0.13, 0.37, 0.62)
        WHITE = Color(1, 1, 1)
        detail_start_row = len(files) + 3
        format_cell_ranges(ws, [
            ('A1:C1', CellFormat(backgroundColor=BLUE_DARK,
                                 textFormat=TextFormat(bold=True, foregroundColor=WHITE))),
            (f'A{detail_start_row}:{col_letter(len(avail_cols))}{detail_start_row}',
             CellFormat(backgroundColor=BLUE_DARK,
                        textFormat=TextFormat(bold=True, foregroundColor=WHITE))),
        ])
    except Exception as e:
        print(f'[WARNING] Formatting validation report: {e}')
        
    terapkan_filter(ws, detail_start_row)


# ─────────────────────────── MAIN PROCESS ──────────────────────

def proses_data():
    summary_path = get_summary_output_path('.', extension='.xlsx')
    try:
        targets_local, target_rows = load_or_create_summary_targets(summary_path)
    except ValueError as e:
        print(f'[ERROR] {e}')
        return

    print('\nMemproses data...')

    df_raw = baca_semua_csv()
    if df_raw is None:
        return

    df, _df_removed = validasi_data(df_raw)

    target_rows = _enrich_targets_with_df_values(df, target_rows)
    if target_rows is not None:
        targets_local = _target_rows_to_dict(target_rows)

    try:
        output_path = export_summary_excel(
            df,
            targets_local,
            output_dir='.',
            target_rows=target_rows,
            df_removed=_df_removed,
        )
        print(f'Summary SOS berhasil dibuat: {output_path}')
        
        output_store_detail = export_store_detail_excel(
            df,
            output_dir='.'
        )
        print(f'Store Detail berhasil dibuat: {output_store_detail}')
    except ValueError as e:
        print(f'[ERROR] {e}')
        return


# ─────────────────────────── WATCHDOG ──────────────────────────

class CSVHandler(FileSystemEventHandler):
    def on_any_event(self, event):
        if event.event_type not in ('modified', 'created'):
            return
        path = event.src_path
        if path.endswith('.csv'):
            print(f'File CSV terdeteksi: {path}')
            time.sleep(2)
            proses_data()
        elif (path.endswith('.xlsx')
              and os.path.basename(path).startswith('Summary SOS_')
              and time.time() - _last_excel_write_time > 5):
            print(f'File TARGETS berubah: {path}')
            time.sleep(2)
            proses_data()


if __name__ == '__main__':
    proses_data()  # proses langsung saat pertama dijalankan

    observer = PollingObserver()
    observer.schedule(CSVHandler(), path='.', recursive=True)
    observer.start()
    print('Watchdog aktif! Menunggu perubahan CSV...')

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
