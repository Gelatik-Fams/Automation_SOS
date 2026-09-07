"""Dashboard payload assembly and Excel row/column projection."""

from sos_engine.common import (
    filter_dashboard_by_division,
    get_dashboard_division_options,
    sort_key_period,
)
from sos_engine.tables import (
    buat_account_divisi_section,
    buat_category_divisi_section,
    buat_division_summary_rows,
    buat_region_divisi_section,
    buat_tabel_channel_account,
    buat_tabel_sos_monthly,
)
from sos_engine.targets import (
    _build_division_targets,
)


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
