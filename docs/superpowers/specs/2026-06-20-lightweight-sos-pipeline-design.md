# Lightweight SOS Pipeline Design

## Goal

Refactor `tes.py` into a lightweight, single-file SOS processing pipeline that reads real 2024 CSV data, calculates pivot-style Share of Shelf, and writes a local Excel workbook for formula validation.

## Scope

The implementation remains in `tes.py`. It does not introduce a modular folder structure, dashboard, rankings, configuration files, PyInstaller packaging, or OOS/OSA calculations.

Google Sheets upload and file watching are not part of the default execution path. The script runs once for an explicitly supplied input:

```powershell
python tes.py data-real-2024.csv
```

The default output is:

```text
output_sos_summary.xlsx
```

## Input

CSV is the primary input format. XLSX remains supported as a fallback. XLSX input reads the `RAW DATA` sheet when it exists; otherwise it reads the first sheet.

Relevant real-data columns include:

```text
Visit Date
Month
Region
Produsen
Divisi
Facing
```

Additional input columns are preserved in `RAW_CLEAN`.

## Functions

### `reader(path)`

- Accepts CSV and XLSX paths.
- Raises a clear error for a missing file or unsupported extension.
- Handles common CSV encodings and delimiter detection where practical.
- Returns a pandas DataFrame without applying business calculations.

### `validator(df)`

- Reports missing required columns.
- Reports rows with missing or non-numeric `Facing`.
- Reports rows with null, empty, or whitespace-only `Produsen` as `UNKNOWN_PRODUCER`.
- Reports rows with missing `Region`.
- Returns a validation-report DataFrame.
- Does not classify valid competitor names as issues.

Required source columns are:

```text
Month
Region
Produsen
Facing
```

At least one usable source for `Year` must exist: either `Year` or `Visit Date`.

`Divisi` is optional because the summary includes it only when available.

### `transformer(df)`

- Trims column names and normalizes known spelling/casing aliases to canonical names.
- Preserves all source columns.
- Converts `Facing` to numeric using coercion; invalid and missing values become null.
- Derives `Year` from `Visit Date` when `Year` is absent or blank.
- Normalizes `Produsen` by trimming whitespace and converting to uppercase.
- Creates `ProducerGroup` using a configurable constant:

```python
INDOFOOD_PRODUCERS = ["INDOFOOD"]
```

Mapping rules:

- Empty `Produsen` becomes `ProducerGroup = "Unknown"`.
- A normalized value in `INDOFOOD_PRODUCERS` becomes `"Indofood"`.
- Every other non-empty value becomes `"Competitor"`.

### `calculate_sos(df, group_cols)`

- Validates that grouping and calculation columns exist.
- Groups data using the requested dimensions.
- Sums numeric `Facing` separately for Indofood and Competitor.
- Unknown-producer rows are retained in `RAW_CLEAN` and validation, but excluded from both the Indofood and Competitor totals so invalid ownership is not silently assigned.
- Returns:

```text
<group columns>
Facing Indofood
Facing Competitor
Grand Total Facing
SOS_%
```

Formula:

```text
Grand Total Facing = Facing Indofood + Facing Competitor
SOS_% = Facing Indofood / Grand Total Facing * 100
```

If Grand Total Facing is zero, `SOS_%` is null.

For this phase, grouping columns are:

```text
Year
Month
Region
Divisi  # only when available
```

### `write_excel_output(...)`

Writes `output_sos_summary.xlsx` with exactly these sheets:

1. `RAW_CLEAN`
2. `SOS_SUMMARY`
3. `VALIDATION_REPORT`

The workbook is overwritten on each successful run.

## Execution Flow

1. Parse the input path and optional output path from command-line arguments.
2. Read the source file.
3. Run validation against source values so invalid Facing and empty producer values remain observable.
4. Transform and normalize the source.
5. Stop with a clear error if required columns prevent calculation.
6. Build the phase-one SOS summary.
7. Append zero-grand-total warnings to the validation report.
8. Write the three-sheet workbook.
9. Log row counts, issue counts, grouping columns, and output location.

## Error Handling

- Fatal schema and file errors produce a concise message and a non-zero exit status.
- Row-level quality issues do not stop processing; they are written to `VALIDATION_REPORT`.
- Rows with invalid/missing Facing do not contribute to sums because transformed Facing is null.
- Missing Region is reported and excluded from grouped output.
- Missing both `Year` and a usable `Visit Date` is a fatal schema issue because the required summary cannot be produced.

## Validation Report Schema

The report uses stable columns:

```text
Issue
Severity
Row
Column
Value
Message
```

Schema-level issues use a blank `Row`. Source row numbers refer to spreadsheet-style rows, where the header is row 1 and the first data record is row 2.

## Acceptance Criteria

- `python tes.py data-real-2024.csv` completes successfully for a valid source.
- `output_sos_summary.xlsx` is created.
- `RAW_CLEAN` contains normalized numeric Facing and `ProducerGroup`.
- `SOS_SUMMARY` contains the required grouping and SOS result columns.
- Zero grand totals produce blank SOS values.
- Only actual quality issues appear in `VALIDATION_REPORT`.
- Non-empty, non-Indofood producer values are valid competitors.
- The default path performs no Google Sheets upload and starts no watcher.
