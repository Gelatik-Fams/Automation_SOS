import pandas as pd
import gspread
import time
from google.oauth2.service_account import Credentials
from watchdog.observers.polling import PollingObserver
from watchdog.events import FileSystemEventHandler


def proses_data():
    print('Ada perubahan! Memproses data...')

    # credential untuk masuk sheet
    scope  = ['https://spreadsheets.google.com/feeds',
               'https://www.googleapis.com/auth/drive']
    creds  = Credentials.from_service_account_file('gelatik-automation-6ba4ee1032b4.json', scopes=scope)
    client = gspread.authorize(creds)
    sheet  = client.open_by_key('1rExHFDTKoAnBabv5PZazE1AdEISDdPJsAcYhG1TTsnk')

    # setup SUMMARY
    titles = [ws.title for ws in sheet.worksheets()]
    if 'SUMMARY' not in titles:
        sheet.add_worksheet(title='SUMMARY', rows=1000, cols=20)
    ws_summary = sheet.worksheet('SUMMARY')

    # setup SOS SUMMARY
    if 'SOS SUMMARY' not in titles:
        sheet.add_worksheet(title='SOS SUMMARY', rows=1000, cols=20)
    ws_sos = sheet.worksheet('SOS SUMMARY')

    # file lokal
    df = pd.read_excel('GELATIK_RAW_DATA_DUMMY.xlsx', sheet_name='RAW DATA')

    # logika SUMMARY
    existing    = ws_summary.get_all_records()
    df_existing = pd.DataFrame(existing)

    summary_baru = df.groupby(['Year', 'Month', 'Region']).agg(
        total_qty   = ('Sales Qty (Pcs)',      'sum'),
        total_value = ('Sales Value (Rp 000)', 'sum'),
        jumlah_oos  = ('OOS', lambda x: (x == 'Y').sum()),
        jumlah_toko = ('Store Code',           'nunique')
    ).reset_index()
    summary_baru['Kontribusi'] = (
        summary_baru['total_qty'] / summary_baru['total_qty'].sum() * 100
    ).round(2)

    bulan_baru = summary_baru['Month'].unique()
    tahun_baru = summary_baru['Year'].unique()
    if not df_existing.empty:
        df_existing = df_existing[
            ~((df_existing['Month'].isin(bulan_baru)) &
              (df_existing['Year'].isin(tahun_baru)))
        ]

    df_final = pd.concat([df_existing, summary_baru], ignore_index=True)
    ws_summary.clear()
    ws_summary.update([df_final.columns.tolist()] + df_final.values.tolist())
    print('Sheet SUMMARY berhasil diupdate!')

    # logika SOS
    existing_sos    = ws_sos.get_all_records()
    df_existing_sos = pd.DataFrame(existing_sos)

    df_noodle = df[df['Divisi'] == 'NOODLE']
    sos_baru  = df_noodle.groupby(['Year', 'Month', 'Region']).agg(
        total_facing_indomie    = ('Facing Indomie',  'sum'),
        total_facing_kompetitor = ('Facing Mi Sedap', 'sum'),
        total_facing_all        = ('Total Facing',    'sum'),
    ).reset_index()
    sos_baru['SOS_%'] = (
        sos_baru['total_facing_indomie'] /
        sos_baru['total_facing_all'] * 100
    ).round(2)
    sos_baru['Target_60%'] = sos_baru['SOS_%'].apply(
        lambda x: 'Achieved' if x >= 60 else 'Below Target'
    )

    if not df_existing_sos.empty:
        df_existing_sos = df_existing_sos[
            ~((df_existing_sos['Month'].isin(bulan_baru)) &
              (df_existing_sos['Year'].isin(tahun_baru)))
        ]

    df_sos_final = pd.concat([df_existing_sos, sos_baru], ignore_index=True)
    ws_sos.clear()
    ws_sos.update([df_sos_final.columns.tolist()] + df_sos_final.values.tolist())
    print('Sheet SOS SUMMARY berhasil diupdate!')


# watcher untuk excel lokal
class ExcelHandler(FileSystemEventHandler):
    def on_any_event(self, event):
        print(f'Event: {event.event_type} - {event.src_path}')
        if event.event_type == 'deleted' and '.~lock.GELATIK_RAW_DATA_DUMMY' in event.src_path:
            print('File disimpan! Memproses...')
            time.sleep(2)
            proses_data()


observer = PollingObserver()
observer.schedule(ExcelHandler(), path='.', recursive=False)
observer.start()
print('Watchdog aktif! Menunggu perubahan file...')

try:
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    observer.stop()
observer.join()