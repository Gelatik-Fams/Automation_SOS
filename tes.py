import pandas as pd
import gspread
import time
from google.oauth2.service_account import Credentials
from watchdog.observers.polling import PollingObserver
from watchdog.events import FileSystemEventHandler

CSV_FILE = 'Report Product - Januari 2024.csv'

MONTH_ORDER = ['January', 'February', 'March', 'April', 'May', 'June',
               'July', 'August', 'September', 'October', 'November', 'December']

METRIC_LABELS = ['Indofood', 'Kompetitor', 'Total Facing', 'SOS%', 'Store Count']
N_METRICS     = len(METRIC_LABELS)  # 5 kolom per bulan


def col_letter(n):
    result = ''
    n += 1
    while n > 0:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result


def hitung_sos(df, groupby_cols):
    g = groupby_cols + ['Month']

    df_indofood   = df[~df['Product Code'].astype(str).str.contains('COMPETITOR', case=False, na=False)]
    df_kompetitor = df[ df['Product Code'].astype(str).str.contains('COMPETITOR', case=False, na=False)]

    fi = df_indofood.groupby(g)['Facing'].sum().reset_index().rename(columns={'Facing': 'facing_indofood'})
    fk = df_kompetitor.groupby(g)['Facing'].sum().reset_index().rename(columns={'Facing': 'facing_kompetitor'})

    merged = fi.merge(fk, on=g, how='outer').fillna(0)
    merged['total_facing'] = merged['facing_indofood'] + merged['facing_kompetitor']
    merged['SOS_%'] = (
        merged['facing_indofood'] / merged['total_facing'].replace(0, float('nan')) * 100
    ).round(2).fillna(0)

    return merged


def buat_section_table(df, index_col, months):
    sos_data   = hitung_sos(df, [index_col])
    store_data = df.groupby([index_col, 'Month'])['Store Code'].nunique().reset_index()
    store_data.rename(columns={'Store Code': 'store_count'}, inplace=True)

    ordered_months = [m for m in months if m in sos_data['Month'].unique()]

    # Pivot semua metrik
    def pivot(col, agg):
        p = sos_data.pivot_table(index=index_col, columns='Month', values=col, aggfunc=agg, fill_value=0)
        return p.reindex(columns=ordered_months, fill_value=0)

    pv_fi  = pivot('facing_indofood',   'sum')
    pv_fk  = pivot('facing_kompetitor', 'sum')
    pv_ft  = pivot('total_facing',      'sum')
    pv_sos = pivot('SOS_%',             'mean')

    sc = store_data.pivot_table(index=index_col, columns='Month', values='store_count', aggfunc='sum', fill_value=0)
    sc = sc.reindex(columns=ordered_months, fill_value=0)

    # Header baris 1 — nama bulan (tiap bulan span N_METRICS kolom)
    header1 = [index_col]
    for m in ordered_months:
        header1 += [m] + [''] * (N_METRICS - 1)
    header1 += ['TOTAL'] + [''] * (N_METRICS - 1)

    # Header baris 2 — nama metrik per bulan
    header2 = [''] + METRIC_LABELS * (len(ordered_months) + 1)

    rows = [header1, header2]

    # Baris data
    for idx in pv_fi.index:
        row      = [str(idx)]
        tot_fi   = tot_fk = tot_ft = tot_sc = 0

        for m in ordered_months:
            fi  = int(round(float(pv_fi.loc[idx, m])))
            fk  = int(round(float(pv_fk.loc[idx, m])))
            ft  = int(round(float(pv_ft.loc[idx, m])))
            sos = round(float(pv_sos.loc[idx, m]), 2)
            sc_val = int(sc.loc[idx, m]) if idx in sc.index else 0

            row    += [fi, fk, ft, sos, sc_val]
            tot_fi += fi
            tot_fk += fk
            tot_ft += ft
            tot_sc += sc_val

        tot_sos = round(tot_fi / tot_ft * 100, 2) if tot_ft > 0 else 0
        row += [tot_fi, tot_fk, tot_ft, tot_sos, tot_sc]
        rows.append(row)

    # Grand Total
    grand    = ['GRAND TOTAL']
    gt_fi    = gt_fk = gt_ft = gt_sc = 0

    for m in ordered_months:
        fi  = int(round(pv_fi[m].sum()))
        fk  = int(round(pv_fk[m].sum()))
        ft  = int(round(pv_ft[m].sum()))
        sos = round(fi / ft * 100, 2) if ft > 0 else 0
        sc_val = int(sc[m].sum()) if m in sc.columns else 0

        grand  += [fi, fk, ft, sos, sc_val]
        gt_fi  += fi
        gt_fk  += fk
        gt_ft  += ft
        gt_sc  += sc_val

    gt_sos = round(gt_fi / gt_ft * 100, 2) if gt_ft > 0 else 0
    grand += [gt_fi, gt_fk, gt_ft, gt_sos, gt_sc]
    rows.append(grand)

    num_cols = len(header1)
    return rows, num_cols


def hapus_semua_formatting(ws, total_rows, total_cols):
    try:
        from gspread_formatting import format_cell_range, CellFormat
        last_col = col_letter(total_cols - 1)
        format_cell_range(ws, f'A1:{last_col}{total_rows}', CellFormat())
    except Exception:
        pass


def terapkan_filter(ws, header_row):
    try:
        ws.spreadsheet.batch_update({
            'requests': [{
                'setBasicFilter': {
                    'filter': {
                        'range': {
                            'sheetId'        : ws.id,
                            'startRowIndex'  : header_row - 1,
                            'startColumnIndex': 0,
                        }
                    }
                }
            }]
        })
    except Exception as e:
        print(f'Filter error (diabaikan): {e}')


def buat_dashboard(ws, df):
    ws.clear()

    semua_bulan = [m for m in MONTH_ORDER if m in df['Month'].unique()]

    levels = [
        ('REGION',           'Region'),
        ('AREA',             'Area'),
        ('CHANNEL',          'Channel'),
        ('CATEGORY CHANNEL', 'Category Channel'),
    ]

    all_rows       = []
    fmt_info       = []
    ordered_months = semua_bulan  # akan di-filter di buat_section_table

    for label, col in levels:
        table_rows, num_cols = buat_section_table(df, col, semua_bulan)

        title_row_idx   = len(all_rows) + 1
        all_rows.append([f'DASHBOARD BY {label}'])

        header1_row_idx = len(all_rows) + 1
        header2_row_idx = len(all_rows) + 2
        all_rows.extend(table_rows)
        grand_total_idx = len(all_rows)

        fmt_info.append({
            'title_row'      : title_row_idx,
            'header1_row'    : header1_row_idx,
            'header2_row'    : header2_row_idx,
            'grand_total_row': grand_total_idx,
            'num_cols'       : num_cols,
            'label'          : label,
        })

        all_rows.append([])
        all_rows.append([])

    total_rows = len(all_rows)
    total_cols = max(info['num_cols'] for info in fmt_info)

    # Hapus semua formatting lama dulu sebelum tulis baru
    hapus_semua_formatting(ws, total_rows + 10, total_cols + 5)

    ws.update('A1', all_rows)

    try:
        from gspread_formatting import format_cell_ranges, CellFormat, Color, TextFormat

        BLUE_DARK   = Color(0.13, 0.37, 0.62)
        BLUE_MED    = Color(0.27, 0.51, 0.71)
        BLUE_LIGHT  = Color(0.64, 0.76, 0.89)
        GREY        = Color(0.85, 0.85, 0.85)
        WHITE       = Color(1, 1, 1)

        requests = []
        for info in fmt_info:
            ec = col_letter(info['num_cols'] - 1)
            tr = info['title_row']
            h1 = info['header1_row']
            h2 = info['header2_row']
            gr = info['grand_total_row']

            requests += [
                (f'A{tr}:{ec}{tr}', CellFormat(
                    backgroundColor=GREY,
                    textFormat=TextFormat(bold=True, fontSize=11)
                )),
                (f'A{h1}:{ec}{h1}', CellFormat(
                    backgroundColor=BLUE_DARK,
                    textFormat=TextFormat(bold=True, foregroundColor=WHITE)
                )),
                (f'A{h2}:{ec}{h2}', CellFormat(
                    backgroundColor=BLUE_MED,
                    textFormat=TextFormat(bold=True, foregroundColor=WHITE)
                )),
                (f'A{gr}:{ec}{gr}', CellFormat(
                    backgroundColor=BLUE_LIGHT,
                    textFormat=TextFormat(bold=True)
                )),
            ]

        format_cell_ranges(ws, requests)
        print('Formatting berhasil!')

    except ImportError:
        print('Install gspread-formatting: pip install gspread-formatting')
    except Exception as e:
        print(f'Formatting error (diabaikan): {e}')

    tambah_chart(ws.spreadsheet, ws, fmt_info, ordered_months)


def buat_store_detail(ws, df):
    ws.clear()

    semua_bulan  = [m for m in MONTH_ORDER if m in df['Month'].unique()]
    index_cols   = ['Region', 'Area', 'Store Name']

    sos_data     = hitung_sos(df, index_cols)
    store_data   = df.groupby(index_cols + ['Month'])['Store Code'].nunique().reset_index()
    store_data.rename(columns={'Store Code': 'store_count'}, inplace=True)

    ordered_months = [m for m in semua_bulan if m in sos_data['Month'].unique()]

    def make_pivot(data, val, agg):
        p = data.pivot_table(index=index_cols, columns='Month', values=val, aggfunc=agg, fill_value=0)
        return p.reindex(columns=ordered_months, fill_value=0)

    pv_fi  = make_pivot(sos_data,   'facing_indofood',   'sum')
    pv_fk  = make_pivot(sos_data,   'facing_kompetitor', 'sum')
    pv_ft  = make_pivot(sos_data,   'total_facing',      'sum')
    pv_sos = make_pivot(sos_data,   'SOS_%',             'mean')
    pv_sc  = make_pivot(store_data, 'store_count',       'sum')

    n_idx    = len(index_cols)
    header1  = index_cols[:]
    header2  = [''] * n_idx
    for m in ordered_months:
        header1 += [m] + [''] * (N_METRICS - 1)
        header2 += METRIC_LABELS
    header1 += ['TOTAL'] + [''] * (N_METRICS - 1)
    header2 += METRIC_LABELS

    table_rows = [header1, header2]

    for idx in pv_fi.index:
        row = list(idx)
        tot_fi = tot_fk = tot_ft = tot_sc = 0

        for m in ordered_months:
            fi  = int(round(float(pv_fi.loc[idx, m])))
            fk  = int(round(float(pv_fk.loc[idx, m])))
            ft  = int(round(float(pv_ft.loc[idx, m])))
            sos = round(float(pv_sos.loc[idx, m]), 2)
            sc_val = int(pv_sc.loc[idx, m]) if idx in pv_sc.index else 0

            row += [fi, fk, ft, sos, sc_val]
            tot_fi += fi; tot_fk += fk; tot_ft += ft; tot_sc += sc_val

        tot_sos = round(tot_fi / tot_ft * 100, 2) if tot_ft > 0 else 0
        row += [tot_fi, tot_fk, tot_ft, tot_sos, tot_sc]
        table_rows.append(row)

    grand = ['GRAND TOTAL', '', '']
    gt_fi = gt_fk = gt_ft = gt_sc = 0
    for m in ordered_months:
        fi  = int(round(pv_fi[m].sum()))
        fk  = int(round(pv_fk[m].sum()))
        ft  = int(round(pv_ft[m].sum()))
        sos = round(fi / ft * 100, 2) if ft > 0 else 0
        sc_val = int(pv_sc[m].sum()) if m in pv_sc.columns else 0
        grand += [fi, fk, ft, sos, sc_val]
        gt_fi += fi; gt_fk += fk; gt_ft += ft; gt_sc += sc_val
    gt_sos = round(gt_fi / gt_ft * 100, 2) if gt_ft > 0 else 0
    grand += [gt_fi, gt_fk, gt_ft, gt_sos, gt_sc]
    table_rows.append(grand)

    num_cols = len(header1)
    hapus_semua_formatting(ws, len(table_rows) + 10, num_cols + 5)
    ws.update('A1', table_rows)
    ws.freeze(rows=2, cols=3)   # freeze 2 header rows + 3 kolom (Region, Area, Store)
    terapkan_filter(ws, 1)

    try:
        from gspread_formatting import format_cell_ranges, CellFormat, Color, TextFormat

        BLUE_DARK  = Color(0.13, 0.37, 0.62)
        BLUE_MED   = Color(0.27, 0.51, 0.71)
        BLUE_LIGHT = Color(0.64, 0.76, 0.89)
        WHITE      = Color(1, 1, 1)
        ec = col_letter(num_cols - 1)
        gr = len(table_rows)

        format_cell_ranges(ws, [
            (f'A1:{ec}1', CellFormat(
                backgroundColor=BLUE_DARK,
                textFormat=TextFormat(bold=True, foregroundColor=WHITE)
            )),
            (f'A2:{ec}2', CellFormat(
                backgroundColor=BLUE_MED,
                textFormat=TextFormat(bold=True, foregroundColor=WHITE)
            )),
            (f'A{gr}:{ec}{gr}', CellFormat(
                backgroundColor=BLUE_LIGHT,
                textFormat=TextFormat(bold=True)
            )),
        ])
    except Exception as e:
        print(f'Formatting store error (diabaikan): {e}')

    print('Sheet STORE DETAIL berhasil diupdate!')


def tambah_chart(spreadsheet, ws, fmt_info, ordered_months):
    """Tambahkan bar chart SOS% per Region ke DASHBOARD."""

    region_info = next((i for i in fmt_info if i['label'] == 'REGION'), None)
    if not region_info:
        return

    ws_id      = ws.id
    h2_0idx    = region_info['header2_row'] - 1   # 0-indexed: baris metric labels
    data_start = h2_0idx + 1                       # baris data pertama
    data_end   = region_info['grand_total_row'] - 2  # exclude grand total

    # Kolom SOS% untuk tiap bulan: col 0=index, lalu tiap bulan punya N_METRICS kolom
    # SOS% ada di posisi ke-4 dari setiap grup bulan (0-indexed dalam grup)
    series = []
    for i, m in enumerate(ordered_months):
        sos_col = 1 + N_METRICS * i + 3  # Indofood=0, Komp=1, Total=2, SOS%=3
        series.append({
            'series': {
                'sourceRange': {
                    'sources': [{
                        'sheetId'         : ws_id,
                        'startRowIndex'   : data_start,
                        'endRowIndex'     : data_end + 1,
                        'startColumnIndex': sos_col,
                        'endColumnIndex'  : sos_col + 1,
                    }]
                }
            },
            'targetAxis': 'BOTTOM_AXIS',
        })

    try:
        spreadsheet.batch_update({'requests': [{
            'addChart': {
                'chart': {
                    'spec': {
                        'title': 'SOS % by Region',
                        'basicChart': {
                            'chartType'     : 'BAR',
                            'legendPosition': 'BOTTOM_LEGEND',
                            'axis': [
                                {'position': 'BOTTOM_AXIS', 'title': 'SOS %'},
                                {'position': 'LEFT_AXIS',   'title': 'Region'},
                            ],
                            'domains': [{
                                'domain': {
                                    'sourceRange': {
                                        'sources': [{
                                            'sheetId'         : ws_id,
                                            'startRowIndex'   : data_start,
                                            'endRowIndex'     : data_end + 1,
                                            'startColumnIndex': 0,
                                            'endColumnIndex'  : 1,
                                        }]
                                    }
                                }
                            }],
                            'series'     : series,
                            'headerCount': 0,
                        }
                    },
                    'position': {
                        'overlayPosition': {
                            'anchorCell': {
                                'sheetId'     : ws_id,
                                'rowIndex'    : region_info['title_row'] - 1,
                                'columnIndex' : region_info['num_cols'] + 1,
                            },
                            'widthPixels' : 560,
                            'heightPixels': 380,
                        }
                    }
                }
            }
        }]})
        print('Chart berhasil dibuat!')
    except Exception as e:
        print(f'Chart error (diabaikan): {e}')


def validasi_data(df):
    jumlah_sebelum = len(df)
    df = df.drop_duplicates()
    jumlah_hapus = jumlah_sebelum - len(df)
    if jumlah_hapus > 0:
        print(f'[INFO] {jumlah_hapus} baris duplikat dihapus otomatis.')

    df_indofood = df[~df['Product Code'].astype(str).str.contains('COMPETITOR', case=False, na=False)]
    baris_nan   = df_indofood[df_indofood['Facing'].isna()]
    if not baris_nan.empty:
        print(f'[WARNING] {len(baris_nan)} baris Indofood tidak punya data Facing:')
        print(baris_nan[['Visit Date', 'Region', 'Store Name', 'Brand', 'Facing']].head(10).to_string(index=False))

    return df


def proses_data():
    print('Ada perubahan! Memproses data...')

    scope  = ['https://spreadsheets.google.com/feeds',
               'https://www.googleapis.com/auth/drive']
    creds  = Credentials.from_service_account_file('gelatik-automation-6ba4ee1032b4.json', scopes=scope)
    client = gspread.authorize(creds)
    sheet  = client.open_by_key('1rExHFDTKoAnBabv5PZazE1AdEISDdPJsAcYhG1TTsnk')

    titles = [ws.title for ws in sheet.worksheets()]
    for title in ['DASHBOARD', 'STORE DETAIL']:
        if title not in titles:
            sheet.add_worksheet(title=title, rows=5000, cols=100)
    ws_dashboard = sheet.worksheet('DASHBOARD')
    ws_store     = sheet.worksheet('STORE DETAIL')

    df_raw = pd.read_csv(CSV_FILE)
    df_raw['Visit Date'] = pd.to_datetime(df_raw['Visit Date'], dayfirst=False)
    df_raw['Year']  = df_raw['Visit Date'].dt.year

    df = validasi_data(df_raw)

    buat_dashboard(ws_dashboard, df)
    print('DASHBOARD berhasil diupdate!')

    buat_store_detail(ws_store, df)


class CSVHandler(FileSystemEventHandler):
    def on_any_event(self, event):
        if event.event_type in ('modified', 'created') and event.src_path.endswith('.csv'):
            print('File CSV berubah! Memproses...')
            time.sleep(1)
            proses_data()


observer = PollingObserver()
observer.schedule(CSVHandler(), path='.', recursive=False)
observer.start()
print(f'Watchdog aktif! Menunggu perubahan pada {CSV_FILE}...')

try:
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    observer.stop()
observer.join()
