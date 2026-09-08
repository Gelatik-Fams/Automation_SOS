"""Dashboard table builders and their existing aggregations."""

from sos_engine.calculations import (
    calc_sos,
    extract_parent_brand,
    is_competitor,
)
from sos_engine.common import (
    sort_key_period,
)
from sos_engine.targets import (
    get_target,
)


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
