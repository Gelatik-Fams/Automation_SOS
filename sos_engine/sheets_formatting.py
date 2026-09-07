"""Google Sheets API retry, formatting, filters, and charts."""

import time
try:
    import gspread
except ImportError:
    gspread = None

from sos_engine.common import (
    GREEN_BG,
    RED_BG,
    col_letter,
)


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
