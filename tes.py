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
    df_i = df[~is_competitor(df['Product Code'])]
    df_k = df[ is_competitor(df['Product Code'])]

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
    df_i = df[~is_competitor(df['Product Code'])]
    df_k = df[ is_competitor(df['Product Code'])]

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
        spreadsheet.batch_update({'requests': requests})
    except Exception as e:
        print(f'Conditional format error: {e}')


# ─────────────────────────── TARGET DARI DASHBOARD ─────────────

def baca_target_dari_dashboard(ws):
    """
    Baca TARGET (kolom B) dari dashboard sebelum di-clear.
    User bisa edit langsung di kolom TARGET — dipertahankan saat script update.
    """
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


# ─────────────────────────── BANGUN TABEL SOS ──────────────────

def buat_tabel_sos_monthly(df, index_col, dim_label, semua_period, targets,
                            compliance_map=None):
    """
    Tabel SOS% per bulan dengan kolom Indofood | Kompetitor | Total | SOS% per period.
    [index | TARGET | ←Jan 25→ | ←Feb 25→ | ... | AVG | STORE COV. | AKTUAL | %]
                      fi fk tot %   fi fk tot %
    """
    df_i = df[~is_competitor(df['Product Code'])]
    df_k = df[ is_competitor(df['Product Code'])]

    sos_monthly = calc_sos(df, [index_col, 'Period'])
    ordered_periods = [p for p in semua_period if p in sos_monthly['Period'].unique()]
    n = len(ordered_periods)

    # ── Header ──
    header1 = [index_col, 'TARGET']
    header2 = ['', '']
    for p in ordered_periods:
        header1 += [p, '', '', '']
        header2 += ['Indofood', 'Kompetitor', 'Total', 'SOS%']
    header1 += ['AVG', 'STORE COV.', 'AKTUAL', '%']
    header2 += ['', '', '', '']

    # ── Precompute fi/fk per (index_col, Period) ──
    fi_grp = df_i.groupby([index_col, 'Period'])['Facing'].sum()
    fk_grp = df_k.groupby([index_col, 'Period'])['Facing'].sum()

    all_indices = sos_monthly[index_col].unique()

    # ── Baris data ──
    data_rows = []
    for idx in all_indices:
        target_val  = get_target(targets, dim_label, idx)
        row         = [str(idx), target_val]
        monthly_sos = []

        for p in ordered_periods:
            fi  = int(round(fi_grp.get((idx, p), 0)))
            fk  = int(round(fk_grp.get((idx, p), 0)))
            tot = fi + fk
            sos = round(fi / tot * 100, 1) if tot > 0 else 0
            row += [fi, fk, tot, sos]
            monthly_sos.append(sos)

        avg = round(sum(monthly_sos) / len(monthly_sos), 1) if monthly_sos else 0
        row.append(avg)

        cdata = compliance_map.get(str(idx).upper(), (0, 0, 0)) if compliance_map else (0, 0, 0)
        row += list(cdata)
        data_rows.append(row)

    # ── Grand Total ──
    grand    = ['GRAND TOTAL', get_target(targets, dim_label, 'GRAND TOTAL')]
    gt_sos   = []
    for p in ordered_periods:
        fi  = int(round(df_i[df_i['Period'] == p]['Facing'].sum()))
        fk  = int(round(df_k[df_k['Period'] == p]['Facing'].sum()))
        tot = fi + fk
        sos = round(fi / tot * 100, 1) if tot > 0 else 0
        grand += [fi, fk, tot, sos]
        gt_sos.append(sos)

    grand.append(round(sum(gt_sos) / len(gt_sos), 1) if gt_sos else 0)
    if compliance_map:
        gt_cov = sum(v[0] for v in compliance_map.values())
        gt_akt = sum(v[1] for v in compliance_map.values())
        gt_pct = round(gt_akt / gt_cov * 100, 1) if gt_cov > 0 else 0
        grand += [gt_cov, gt_akt, gt_pct]
    else:
        grand += [0, 0, 0]

    rows     = [header1, header2] + data_rows + [grand]
    num_cols = len(header1)

    # Kolom SOS% = kolom ke-5,9,13,... (0-indexed: 2+3, 2+7, ...) + AVG (2+4n)
    sos_cols = [2 + 4*i + 3 for i in range(n)] + [2 + 4*n]

    meta = {
        'num_cols'      : num_cols,
        'sos_col_indices': sos_cols,
        'target_col_idx' : 1,
    }
    return rows, meta


def buat_tabel_channel_account(df, semua_period, targets):
    """
    Tabel SOS% Channel × Account dengan subtotal per Channel.
    [Channel | Account | TARGET | ←Jan 25→ | ←Feb 25→ | ... | AVG | COV | AKTUAL | %]
                                   fi fk tot %  fi fk tot %
    """
    sos_monthly = calc_sos(df, ['Channel', 'Account', 'Period'])
    ordered_periods = [p for p in semua_period if p in sos_monthly['Period'].unique()]
    n = len(ordered_periods)

    df_i = df[~is_competitor(df['Product Code'])]
    df_k = df[ is_competitor(df['Product Code'])]

    # Precompute fi/fk per (Channel, Account, Period)
    fi_grp = df_i.groupby(['Channel', 'Account', 'Period'])['Facing'].sum()
    fk_grp = df_k.groupby(['Channel', 'Account', 'Period'])['Facing'].sum()

    header1 = ['CHANNEL', 'ACCOUNT', 'TARGET']
    header2 = ['', '', '']
    for p in ordered_periods:
        header1 += [p, '', '', '']
        header2 += ['Indofood', 'Kompetitor', 'Total', 'SOS%']
    header1 += ['AVG', 'STORE COV.', 'AKTUAL', '%']
    header2 += ['', '', '', '']

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
            'sos_col_indices': [3 + 4*i + 3 for i in range(n)] + [3 + 4*n],
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
        monthly_sos = []

        for p in ordered_periods:
            fi  = int(round(fi_grp.get((ch, acc, p), 0)))
            fk  = int(round(fk_grp.get((ch, acc, p), 0)))
            tot = fi + fk
            sos = round(fi / tot * 100, 1) if tot > 0 else 0
            row += [fi, fk, tot, sos]
            monthly_sos.append(sos)

        avg = round(sum(monthly_sos) / len(monthly_sos), 1) if monthly_sos else 0
        row.append(avg)

        acc_sos  = store_sos[(store_sos['Channel'] == ch) & (store_sos['Account'] == acc)]
        coverage = len(acc_sos)
        aktual   = len(acc_sos[acc_sos['SOS%'] >= target_val])
        pct      = round(aktual / coverage * 100, 1) if coverage > 0 else 0
        row += [coverage, aktual, pct]
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
        sub_sos  = []
        for p in ordered_periods:
            fi  = int(round(fi_ch.get((ch, p), 0)))
            fk  = int(round(fk_ch.get((ch, p), 0)))
            tot = fi + fk
            sos = round(fi / tot * 100, 1) if tot > 0 else 0
            sub_row += [fi, fk, tot, sos]
            sub_sos.append(sos)

        sub_avg = round(sum(sub_sos) / len(sub_sos), 1) if sub_sos else 0
        sub_row.append(sub_avg)
        sub_cov = sum(r[-3] for r in acc_rows)
        sub_akt = sum(r[-2] for r in acc_rows)
        sub_pct = round(sub_akt / sub_cov * 100, 1) if sub_cov > 0 else 0
        sub_row += [sub_cov, sub_akt, sub_pct]
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

    grand.append(round(sum(gt_sos) / len(gt_sos), 1) if gt_sos else 0)
    store_overall = calc_sos(df, ['Store Code'])
    gt_cov = len(store_overall)
    gt_akt = len(store_overall[store_overall['SOS%'] >= 65])
    gt_pct = round(gt_akt / gt_cov * 100, 1) if gt_cov > 0 else 0
    grand += [gt_cov, gt_akt, gt_pct]

    rows     = [header1, header2] + result_rows + [grand]
    num_cols = len(header1)

    # SOS% cols: index 3 (Channel), 4 (Account), 5 (TARGET) → period data starts at col 3
    # per period: fi(+0), fk(+1), tot(+2), sos(+3) → SOS% at 3+3, 3+7, 3+11, ...
    sos_cols = [3 + 4*i + 3 for i in range(n)] + [3 + 4*n]

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
    is_comp = is_competitor(df['Product Code'])

    df = df.copy()
    if 'Period' not in df.columns:
        df = tambah_period_column(df)
        
    df['Parent Brand'] = df.apply(
        lambda row: extract_parent_brand(row['Brand'], is_comp[row.name]), axis=1
    )

    all_periods = sorted(df['Period'].unique(), key=sort_key_period)

    categories = [c for c in ['BAG NOODLE', 'CUP NOODLE', 'REGULER NOODLE']
                  if c in df['Category Channel'].unique()]

    CD_METRICS = ['Facing', 'Total', 'SOS%']
    n_met = len(CD_METRICS)

    # Header baris 1
    header1 = ['Category by Divisi', 'Brand By Facing', 'TARGET']
    for p in all_periods:
        header1 += [p] + [''] * (n_met - 1)
    header1 += ['TOTAL'] + [''] * (n_met - 1)

    # Header baris 2
    header2 = ['', '', ''] + CD_METRICS * (len(all_periods) + 1)

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
            target = get_target(targets, 'CATEGORY BY DIVISI', brand_name)
            row = [cat, brand_name, target]
            grand_facing = 0
            grand_total = 0

            for p in all_periods:
                bf = int(round(df_brand[df_brand['Period'] == p]['Facing'].sum()))
                tf = int(round(total_period.get(p, 0)))
                sos = round(bf / tf * 100, 2) if tf > 0 else 0
                row += [bf, tf, sos]
                grand_facing += bf
                grand_total += tf

            grand_sos = round(grand_facing / grand_total * 100, 2) if grand_total > 0 else 0
            row += [grand_facing, grand_total, grand_sos]
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
            row = [label, '', '']
            s_facing = 0
            s_total = 0

            for p in all_periods:
                bf = int(round(df_sub[df_sub['Period'] == p]['Facing'].sum()))
                tf = int(round(total_period.get(p, 0)))
                sos = round(bf / tf * 100, 2) if tf > 0 else 0
                row += [bf, tf, sos]
                s_facing += bf
                s_total += tf

            s_sos = round(s_facing / s_total * 100, 2) if s_total > 0 else 0
            row += [s_facing, s_total, s_sos]
            return row

        subtotal_row_indices.append(len(rows))
        rows.append(make_subtotal_row(f'{cat} COMPETITOR Total', df_cat[cat_comp]))
        subtotal_row_indices.append(len(rows))
        rows.append(make_subtotal_row(f'{cat} INDOFOOD Total', df_cat[~cat_comp]))

    # Grand Total
    total_all_period = df.groupby('Period')['Facing'].sum()
    df_indo = df[~is_comp]
    
    target_gt = get_target(targets, 'CATEGORY BY DIVISI', 'DEFAULT')
    grand = ['GRAND TOTAL', '', target_gt]
    g_facing = 0
    g_total = 0

    for p in all_periods:
        fi = int(round(df_indo[df_indo['Period'] == p]['Facing'].sum()))
        tf = int(round(total_all_period.get(p, 0)))
        sos = round(fi / tf * 100, 2) if tf > 0 else 0
        grand += [fi, tf, sos]
        g_facing += fi
        g_total += tf

    g_sos = round(g_facing / g_total * 100, 2) if g_total > 0 else 0
    grand += [g_facing, g_total, g_sos]
    rows.append(grand)

    num_cols = len(header1)
    
    sos_cols = [5 + 3*i for i in range(len(all_periods) + 1)]
    
    meta = {
        'num_cols': num_cols,
        'sos_col_indices': sos_cols,
        'target_col_idx': 2,
        'subtotal_rows': subtotal_row_indices,
        'cat_ranges': cat_ranges,
        'periods': all_periods
    }
    
    return rows, meta


def tambahkan_chart_category_divisi(spreadsheet, ws_id, fmt_section):
    requests = []
    
    data_start = fmt_section['data_start']
    cat_ranges = fmt_section.get('cat_ranges', [])
    periods = fmt_section.get('periods', [])
    sos_cols = fmt_section.get('sos_col_indices', [])
    
    # We want to place charts below the table.
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


# ─────────────────────────── DASHBOARD ─────────────────────────

def buat_dashboard(ws, df):
    targets = baca_target_dari_dashboard(ws)
    hapus_semua_chart(ws.spreadsheet, ws.id)
    api_retry(ws.clear)

    semua_period = sorted(df['Period'].unique(), key=sort_key_period)

    # Level tunggal: level_label, kolom_di_df, dim_key_untuk_targets
    single_levels = [
        ('REGION',           'Region',           'REGION'),
        ('CHANNEL',          'Channel',          'CHANNEL'),
        ('CATEGORY CHANNEL', 'Category Channel', 'CATEGORY CHANNEL'),
    ]

    all_rows     = []
    fmt_sections = []

    # ── Section: single-column levels ──
    for label, col, dim_label in single_levels:
        if col not in df.columns or df[col].dropna().empty:
            continue

        compliance_map = calc_compliance(df, col, targets, dim_label)
        table_rows, meta = buat_tabel_sos_monthly(
            df, col, dim_label, semua_period, targets, compliance_map
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
    if 'Channel' in df.columns and 'Account' in df.columns:
        table_rows, meta = buat_tabel_channel_account(df, semua_period, targets)

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

    # ── Section: CATEGORY BY DIVISI ──
    table_rows, meta = buat_category_divisi_section(df, semua_period, targets)

    title_row   = len(all_rows) + 1
    all_rows.append(['SOS% BY CATEGORY BY DIVISI'])
    header1_row = len(all_rows) + 1
    header2_row = len(all_rows) + 2
    data_start  = len(all_rows) + 3
    all_rows.extend(table_rows)
    grand_row   = len(all_rows)

    subtotal_rows_abs = [header1_row + i for i in meta.get('subtotal_rows', [])]

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
        'subtotal_rows'  : subtotal_rows_abs, 'cat_ranges': meta.get('cat_ranges', []), 'periods': meta.get('periods', []),
    })
    all_rows.append([]); all_rows.append([])

    if not fmt_sections:
        print('[WARNING] Tidak ada data untuk dashboard.')
        return

    total_rows = len(all_rows)
    total_cols = max(s['num_cols'] for s in fmt_sections)

    hapus_semua_formatting(ws, total_rows + 10, total_cols + 5)
    time.sleep(1)
    api_retry(ws.update, range_name='A1', values=all_rows)
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
            tgt_col = col_letter(s['target_col_idx'])

            cell_fmt += [
                (f'A{tr}:{ec}{tr}', CellFormat(backgroundColor=GREY,
                    textFormat=TextFormat(bold=True, fontSize=11))),
                (f'A{h1}:{ec}{h1}', CellFormat(backgroundColor=BLUE_DARK,
                    textFormat=TextFormat(bold=True, foregroundColor=WHITE))),
                (f'A{h2}:{ec}{h2}', CellFormat(backgroundColor=BLUE_MED,
                    textFormat=TextFormat(bold=True, foregroundColor=WHITE))),
                (f'{tgt_col}{ds}:{tgt_col}{gr}', CellFormat(backgroundColor=ORANGE,
                    textFormat=TextFormat(bold=True))),
                (f'A{gr}:{ec}{gr}', CellFormat(backgroundColor=BLUE_LIGHT,
                    textFormat=TextFormat(bold=True))),
            ]
            
            if 'subtotal_rows' in s:
                for sr in s['subtotal_rows']:
                    cell_fmt.append(
                        (f'A{sr}:{ec}{sr}', CellFormat(backgroundColor=BLUE_MED,
                            textFormat=TextFormat(bold=True, foregroundColor=WHITE)))
                    )

        api_retry(format_cell_ranges, ws, cell_fmt)
        time.sleep(1)

        # Conditional formatting per section (hanya kolom SOS%)
        for s in fmt_sections:
            terapkan_conditional_format(
                ws.spreadsheet, ws.id,
                s['data_start'], s['grand_row'],
                s['sos_col_indices'],
                s['target_col_idx'],
            )
            time.sleep(1)

        print('Formatting dashboard berhasil!')

        for s in fmt_sections:
            if s['label'] == 'CATEGORY BY DIVISI':
                tambahkan_chart_category_divisi(ws.spreadsheet, ws.id, s)
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
    df = df.drop_duplicates()
    if len(df) < n:
        print(f'[INFO] {n - len(df)} baris duplikat dihapus.')
    df_i      = df[~is_competitor(df['Product Code'])]
    baris_nan = df_i[df_i['Facing'].isna()]
    if not baris_nan.empty:
        print(f'[WARNING] {len(baris_nan)} baris Indofood tidak punya Facing.')
    return df


def baca_semua_csv():
    files = sorted(glob.glob('Report Product*.csv'))
    if not files:
        print('[WARNING] Tidak ada file CSV ditemukan.')
        return None

    print(f'[INFO] Membaca {len(files)} file:')
    dfs = []
    for f in files:
        print(f'  -> {f}')
        dfs.append(pd.read_csv(f))

    combined = pd.concat(dfs, ignore_index=True)
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
    for title in ['DASHBOARD', 'STORE DETAIL']:
        if title not in titles:
            sheet.add_worksheet(title=title, rows=5000, cols=200)

    ws_dashboard = sheet.worksheet('DASHBOARD')
    ws_store     = sheet.worksheet('STORE DETAIL')

    df_raw = baca_semua_csv()
    if df_raw is None:
        return

    df = validasi_data(df_raw)

    buat_dashboard(ws_dashboard, df)
    print('DASHBOARD berhasil diupdate!')

    buat_store_detail(ws_store, df)


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
    observer.schedule(CSVHandler(), path='.', recursive=False)
    observer.start()
    print('Watchdog aktif! Menunggu perubahan CSV...')

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
