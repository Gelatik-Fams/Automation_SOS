import pandas as pd
import gspread
import time
import os
import glob
from google.oauth2.service_account import Credentials
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
METADATA_COLS_FOR_DEDUP = {'_source_file', 'Source Division'}


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


def calc_compliance(df, index_col, targets, dim_label):
    """
    Hitung STORE COVERAGE, AKTUAL COMPLIANCE, % COMPLIANCE per nilai index_col.
    Metrik dihitung atas semua period (kumulatif):
      - STORE COVERAGE  = total toko unik yang dikunjungi
      - AKTUAL COMPLIANCE = toko dengan SOS% kumulatif >= target
      - % COMPLIANCE   = aktual / coverage × 100
    Return: dict {str(nilai): (coverage, aktual, pct_str)}
    """
    store_sos = calc_sos(df, [index_col, 'Store Code'])

    result = {}
    for idx in store_sos[index_col].unique():
        data     = store_sos[store_sos[index_col] == idx]
        target   = get_target(targets, dim_label, idx)
        coverage = len(data)
        aktual   = len(data[data['SOS%'] >= target])
        pct      = round(aktual / coverage * 100, 1) if coverage > 0 else 0
        result[str(idx).upper()] = (coverage, aktual, pct)

    return result


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
        except gspread.exceptions.APIError as e:
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

def buat_tabel_sos_monthly(df, index_col, dim_label, semua_period, targets,
                            compliance_map=None):
    """
    Tabel SOS% per bulan dengan kolom Indofood | Kompetitor | Total | SOS% per period.
    [index | TARGET | ←Jan 25→ | ←Feb 25→ | ... | AVG | STORE COV. | AKTUAL | %]
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

    if index_cols[0] == 'Source Division':
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
    [Channel | Account | TARGET | ←Jan 25→ | ←Feb 25→ | ... | AVG | COV | AKTUAL | %]
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

        compliance_map = calc_compliance(df_dashboard, col, targets, dim_label)
        dashboard_index_col = ['Source Division', col] if 'Source Division' in df_dashboard.columns else col
        table_rows, meta = buat_tabel_sos_monthly(
            df_dashboard, dashboard_index_col, dim_label, semua_period, targets, compliance_map
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

        safe_anchor_row = total_rows + 2

        if selected_division == 'ALL':
            tambahkan_chart_division_summary(ws.spreadsheet, ws.id, division_summary_section, safe_anchor_row)
            time.sleep(1)
        else:
            for s in fmt_sections:
                if s['label'] == 'CATEGORY BY DIVISI' and s.get('division') == selected_division:
                    tambahkan_chart_category_divisi(ws.spreadsheet, ws.id, s, safe_anchor_row)
                    time.sleep(1)

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
    print('[INFO] File Count by Source Division')
    for division, count in count_files_by_source_division(files).items():
        print(f'{division:<12} {count} files')
    dfs = []
    for f in files:
        print(f'-> {display_csv_path(f, base_dir)}')
        try:
            df_temp = pd.read_csv(f, low_memory=False)
        except pd.errors.ParserError:
            df_temp = pd.read_csv(f, sep=';', low_memory=False)
        df_temp['_source_file'] = os.path.splitext(os.path.basename(f))[0]
        df_temp['Source Division'] = extract_source_division_from_filename(f)
        dfs.append(df_temp)

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

    print(f'[INFO] Total baris: {len(combined):,}')
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
    print('\nMemproses data...')

    scope  = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds  = Credentials.from_service_account_file(
        'gelatik-automation-6ba4ee1032b4.json', scopes=scope)

    # Retry koneksi Google Sheets sampai 3x kalau jaringan putus
    for attempt in range(1, 4):
        try:
            client = gspread.authorize(creds)
            sheet  = client.open_by_key('1rExHFDTKoAnBabv5PZazE1AdEISDdPJsAcYhG1TTsnk')
            break
        except Exception as e:
            if attempt == 3:
                print(f'[ERROR] Gagal koneksi ke Google Sheets setelah 3x: {e}')
                return
            print(f'[WARNING] Koneksi gagal (attempt {attempt}), retry dalam 5 detik...')
            time.sleep(5)

    titles = [ws.title for ws in sheet.worksheets()]
    for title in ['DASHBOARD', 'STORE DETAIL', 'VALIDATION_REPORT', 'TARGETS']:
        if title not in titles:
            sheet.add_worksheet(title=title, rows=500, cols=5)

    ws_dashboard  = sheet.worksheet('DASHBOARD')
    ws_store      = sheet.worksheet('STORE DETAIL')
    ws_validation = sheet.worksheet('VALIDATION_REPORT')
    ws_targets    = sheet.worksheet('TARGETS')

    df_raw = baca_semua_csv()
    if df_raw is None:
        return

    df, df_removed = validasi_data(df_raw)

    # Inisialisasi sheet TARGETS jika baru dibuat (kosong)
    targets_from_dashboard = baca_target_dari_dashboard(ws_dashboard)
    inisialisasi_ws_targets(ws_targets, targets_from_dashboard, df)

    buat_dashboard(ws_dashboard, df, ws_targets)
    print('DASHBOARD berhasil diupdate!')

    buat_store_detail(ws_store, df)

    buat_validation_report(ws_validation, df_removed)
    print('VALIDATION REPORT berhasil diupdate!')


# ─────────────────────────── WATCHDOG ──────────────────────────

class CSVHandler(FileSystemEventHandler):
    def on_any_event(self, event):
        if event.event_type in ('modified', 'created') and event.src_path.endswith('.csv'):
            print(f'File CSV terdeteksi: {event.src_path}')
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
