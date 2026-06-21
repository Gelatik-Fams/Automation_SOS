# Gelatik Supra SOS Automation

Pipeline ringan untuk menghitung Share of Shelf (SOS) Indofood dari raw data
CSV atau XLSX. Implementasi tahap ini tetap berada dalam satu file,
`tes.py`, agar formula mudah dibandingkan dengan Pivot Table manual.

## Instalasi

Gunakan Python 3.10 atau lebih baru:

```powershell
python -m pip install pandas openpyxl
```

## Menjalankan

Input utama berupa CSV:

```powershell
python tes.py data-real-2024.csv
```

XLSX juga didukung. Jika workbook memiliki sheet `RAW DATA`, sheet tersebut
akan dibaca; jika tidak, script membaca sheet pertama.

```powershell
python tes.py GELATIK_RAW_DATA_DUMMY.xlsx
```

Nama output default:

```text
output_sos_summary.xlsx
```

Lokasi output dapat diubah:

```powershell
python tes.py data-real-2024.csv --output hasil/sos-2024.xlsx
```

## Output

Workbook output berisi:

- `RAW_CLEAN`: seluruh kolom sumber, Facing numerik, Year, dan ProducerGroup.
- `SOS_SUMMARY`: hasil per Year, Month, Region, serta Divisi jika tersedia.
- `VALIDATION_REPORT`: masalah schema dan kualitas data yang ditemukan.

Formula:

```text
SOS_% = Facing Indofood / Grand Total Facing * 100
Grand Total Facing = Facing Indofood + Facing Competitor
```

Jika Grand Total Facing bernilai nol, `SOS_%` dikosongkan.

## Mapping Produsen

Nama produsen dinormalisasi dengan menghapus spasi tepi dan mengubahnya
menjadi huruf kapital.

- Nilai dalam `INDOFOOD_PRODUCERS` menjadi `Indofood`.
- Produsen non-kosong lainnya menjadi `Competitor`.
- Produsen kosong menjadi `Unknown` dan dilaporkan sebagai
  `UNKNOWN_PRODUCER`.

Daftar awal di `tes.py`:

```python
INDOFOOD_PRODUCERS = ["INDOFOOD"]
```

## Kolom Wajib

- `Month`
- `Region`
- `Produsen`
- `Facing`
- Salah satu dari `Year` atau `Visit Date`

`Divisi` bersifat opsional. Nama seperti `Producer`, `Producent`, dan
`Division` dinormalisasi ke nama kolom yang dipakai pipeline.

## Scope Tahap Ini

Tahap ini tidak menjalankan Google Sheets upload atau file watcher. Ranking,
dashboard, OOS/OSA, struktur modul terpisah, dan packaging executable juga
belum termasuk.

Spreadsheet prototype lama:
https://docs.google.com/spreadsheets/d/1rExHFDTKoAnBabv5PZazE1AdEISDdPJsAcYhG1TTsnk/edit?usp=sharing
