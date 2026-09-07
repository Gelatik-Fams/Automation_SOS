"""Google Sheets dashboard, target, and report operations."""

import time

from sos_engine.calculations import (
    hitung_sos,
)
from sos_engine.common import (
    METRIC_LABELS,
    N_METRICS,
    col_letter,
    filter_dashboard_by_division,
    get_dashboard_division_options,
    sort_key_period,
)
from sos_engine.sheets_formatting import (
    api_retry,
    clear_dashboard_content,
    hapus_conditional_format,
    hapus_semua_chart,
    hapus_semua_formatting,
    sembunyikan_dashboard_row_ranges,
    tambahkan_chart_category_divisi,
    tambahkan_chart_division_summary,
    terapkan_conditional_format,
    terapkan_dropdown_division,
    terapkan_filter,
)
from sos_engine.tables import (
    buat_account_divisi_section,
    buat_category_divisi_section,
    buat_division_summary_rows,
    buat_region_divisi_section,
    buat_tabel_channel_account,
    buat_tabel_sos_monthly,
)


def baca_dashboard_division_selector(ws, options):
    try:
        selected = str(ws.acell('B1').value or '').strip()
    except Exception:
        selected = ''
    return selected if selected in options else 'ALL'


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
