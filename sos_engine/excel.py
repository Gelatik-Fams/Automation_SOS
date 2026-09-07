"""Excel worksheet writing, formatting, charts, and Store Detail export."""

from openpyxl import Workbook
from openpyxl.chart import BarChart
from openpyxl.chart import Reference
from openpyxl.chart.data_source import AxDataSource
from openpyxl.chart.data_source import StrRef
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment
from openpyxl.styles import Border
from openpyxl.styles import Font
from openpyxl.styles import PatternFill
from openpyxl.styles import Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import quote_sheetname

from sos_engine.calculations import (
    is_competitor,
)
from sos_engine.common import (
    DEFAULT_TARGET,
    EXCEL_CATEGORY_CHART_ROW_STEP,
    TARGET_DIM_ORDER,
    col_letter,
    get_source_divisions,
    get_store_detail_output_path,
    sort_key_period,
)


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
