"""Lightweight Share of Shelf (SOS) processing pipeline.

Usage:
    python tes.py data-real-2024.csv
    python tes.py data-real-2024.xlsx --output output_sos_summary.xlsx
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path
from typing import Iterable

import pandas as pd


INDOFOOD_PRODUCERS = ["INDOFOOD"]
REQUIRED_COLUMNS = ["Month", "Region", "Produsen", "Facing"]
VALIDATION_COLUMNS = [
    "Issue",
    "Severity",
    "Row",
    "Column",
    "Value",
    "Message",
]

LOGGER = logging.getLogger("gelatik_sos")

_COLUMN_ALIASES = {
    "visitdate": "Visit Date",
    "date": "Visit Date",
    "year": "Year",
    "month": "Month",
    "region": "Region",
    "area": "Area",
    "channel": "Channel",
    "subchannel": "Subchannel",
    "account": "Account",
    "produsen": "Produsen",
    "producent": "Produsen",
    "producer": "Produsen",
    "producername": "Produsen",
    "divisi": "Divisi",
    "division": "Divisi",
    "categorychannel": "Category Channel",
    "category": "Category Channel",
    "brand": "Brand",
    "storecode": "Store Code",
    "storename": "Store Name",
    "productcode": "Product Code",
    "productname": "Product Name",
    "facing": "Facing",
}


def _column_key(value: object) -> str:
    """Return a comparison key that ignores spaces, punctuation, and casing."""
    return "".join(character for character in str(value).casefold() if character.isalnum())


def _canonicalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Return a copy with trimmed and recognized column names canonicalized."""
    renamed = {}
    claimed_names: set[str] = set()

    for original in df.columns:
        trimmed = str(original).strip()
        canonical = _COLUMN_ALIASES.get(_column_key(trimmed), trimmed)
        if canonical in claimed_names and canonical != trimmed:
            canonical = trimmed
        renamed[original] = canonical
        claimed_names.add(canonical)

    return df.rename(columns=renamed).copy()


def reader(path: str | Path) -> pd.DataFrame:
    """Read a CSV or XLSX source file into a DataFrame."""
    input_path = Path(path)
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")
    if not input_path.is_file():
        raise ValueError(f"Input path is not a file: {input_path}")

    suffix = input_path.suffix.casefold()
    if suffix == ".csv":
        last_error: UnicodeDecodeError | None = None
        for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                return pd.read_csv(
                    input_path,
                    sep=None,
                    engine="python",
                    encoding=encoding,
                )
            except UnicodeDecodeError as error:
                last_error = error
        raise ValueError(
            f"Could not decode CSV file {input_path}: {last_error}"
        ) from last_error

    if suffix in {".xlsx", ".xlsm"}:
        workbook = pd.ExcelFile(input_path)
        raw_sheet = next(
            (
                sheet_name
                for sheet_name in workbook.sheet_names
                if sheet_name.strip().casefold() == "raw data"
            ),
            workbook.sheet_names[0],
        )
        return pd.read_excel(workbook, sheet_name=raw_sheet)

    raise ValueError(
        f"Unsupported input format '{input_path.suffix}'. Use .csv or .xlsx."
    )


def _is_blank(series: pd.Series) -> pd.Series:
    return series.isna() | series.astype("string").str.strip().eq("")


def _issue(
    issue: str,
    severity: str,
    row: int | None,
    column: str,
    value: object,
    message: str,
) -> dict[str, object]:
    if pd.isna(value):
        value = ""
    return {
        "Issue": issue,
        "Severity": severity,
        "Row": row,
        "Column": column,
        "Value": value,
        "Message": message,
    }


def validator(df: pd.DataFrame) -> pd.DataFrame:
    """Validate source data and return row-level and schema-level issues."""
    source = _canonicalize_columns(df)
    issues: list[dict[str, object]] = []

    for column in REQUIRED_COLUMNS:
        if column not in source.columns:
            issues.append(
                _issue(
                    "MISSING_REQUIRED_COLUMN",
                    "ERROR",
                    None,
                    column,
                    "",
                    f"Required column '{column}' is missing.",
                )
            )

    has_year = "Year" in source.columns and (~_is_blank(source["Year"])).any()
    has_visit_date = (
        "Visit Date" in source.columns
        and pd.to_datetime(source["Visit Date"], errors="coerce").notna().any()
    )
    if not has_year and not has_visit_date:
        issues.append(
            _issue(
                "MISSING_YEAR_SOURCE",
                "ERROR",
                None,
                "Year / Visit Date",
                "",
                "A usable Year or Visit Date column is required.",
            )
        )

    if "Facing" in source.columns:
        facing_blank = _is_blank(source["Facing"])
        facing_numeric = pd.to_numeric(source["Facing"], errors="coerce")
        facing_non_numeric = ~facing_blank & facing_numeric.isna()

        for index in source.index[facing_blank]:
            issues.append(
                _issue(
                    "MISSING_FACING",
                    "WARNING",
                    int(index) + 2,
                    "Facing",
                    source.at[index, "Facing"],
                    "Facing is empty and will not contribute to SOS.",
                )
            )
        for index in source.index[facing_non_numeric]:
            issues.append(
                _issue(
                    "NON_NUMERIC_FACING",
                    "WARNING",
                    int(index) + 2,
                    "Facing",
                    source.at[index, "Facing"],
                    "Facing is not numeric and will not contribute to SOS.",
                )
            )

    if "Produsen" in source.columns:
        producer_blank = _is_blank(source["Produsen"])
        for index in source.index[producer_blank]:
            issues.append(
                _issue(
                    "UNKNOWN_PRODUCER",
                    "WARNING",
                    int(index) + 2,
                    "Produsen",
                    source.at[index, "Produsen"],
                    "Produsen is empty; ProducerGroup is set to Unknown.",
                )
            )

    if "Region" in source.columns:
        region_blank = _is_blank(source["Region"])
        for index in source.index[region_blank]:
            issues.append(
                _issue(
                    "MISSING_REGION",
                    "WARNING",
                    int(index) + 2,
                    "Region",
                    source.at[index, "Region"],
                    "Region is empty; the row is excluded from SOS summary.",
                )
            )

    return pd.DataFrame(issues, columns=VALIDATION_COLUMNS)


def transformer(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize source values and derive Year and ProducerGroup."""
    clean = _canonicalize_columns(df)

    if "Facing" in clean.columns:
        clean["Facing"] = pd.to_numeric(clean["Facing"], errors="coerce")

    visit_year = pd.Series(pd.NA, index=clean.index, dtype="Int64")
    if "Visit Date" in clean.columns:
        visit_year = pd.to_datetime(
            clean["Visit Date"],
            errors="coerce",
        ).dt.year.astype("Int64")

    if "Year" in clean.columns:
        existing_year = pd.to_numeric(clean["Year"], errors="coerce").astype("Int64")
        clean["Year"] = existing_year.fillna(visit_year)
    else:
        clean["Year"] = visit_year

    if "Produsen" in clean.columns:
        normalized_producer = clean["Produsen"].astype("string").str.strip().str.upper()
        producer_missing = clean["Produsen"].isna() | normalized_producer.eq("")
        clean["Produsen"] = normalized_producer.mask(producer_missing, pd.NA)

        indofood_names = {
            str(name).strip().upper()
            for name in INDOFOOD_PRODUCERS
            if str(name).strip()
        }
        clean["ProducerGroup"] = "Competitor"
        clean.loc[producer_missing, "ProducerGroup"] = "Unknown"
        clean.loc[
            ~producer_missing & normalized_producer.isin(indofood_names),
            "ProducerGroup",
        ] = "Indofood"

    return clean


def calculate_sos(
    df: pd.DataFrame,
    group_cols: Iterable[str],
) -> pd.DataFrame:
    """Calculate pivot-style SOS for any set of grouping columns."""
    groups = list(group_cols)
    required = groups + ["ProducerGroup", "Facing"]
    missing = [column for column in required if column not in df.columns]
    if missing:
        raise ValueError(
            "Cannot calculate SOS; missing columns: " + ", ".join(missing)
        )

    output_columns = groups + [
        "Facing Indofood",
        "Facing Competitor",
        "Grand Total Facing",
        "SOS_%",
    ]

    eligible = df[
        df["ProducerGroup"].isin(["Indofood", "Competitor"])
    ].copy()
    if eligible.empty:
        return pd.DataFrame(columns=output_columns)

    grouped = (
        eligible.groupby(
            groups + ["ProducerGroup"],
            dropna=False,
            as_index=False,
        )["Facing"]
        .sum(min_count=1)
    )
    pivot = (
        grouped.set_index(groups + ["ProducerGroup"])["Facing"]
        .unstack("ProducerGroup")
        .reset_index()
    )
    pivot.columns.name = None

    for producer_group in ("Indofood", "Competitor"):
        if producer_group not in pivot.columns:
            pivot[producer_group] = 0.0
        pivot[producer_group] = pivot[producer_group].fillna(0)

    pivot = pivot.rename(
        columns={
            "Indofood": "Facing Indofood",
            "Competitor": "Facing Competitor",
        }
    )
    pivot["Grand Total Facing"] = (
        pivot["Facing Indofood"] + pivot["Facing Competitor"]
    )
    pivot["SOS_%"] = (
        pivot["Facing Indofood"]
        .div(pivot["Grand Total Facing"].replace(0, pd.NA))
        .mul(100)
        .round(2)
    )

    return pivot[output_columns].sort_values(groups).reset_index(drop=True)


def _zero_total_issues(
    summary: pd.DataFrame,
    group_cols: Iterable[str],
) -> pd.DataFrame:
    issues = []
    for _, row in summary[summary["Grand Total Facing"].eq(0)].iterrows():
        dimensions = ", ".join(f"{column}={row[column]}" for column in group_cols)
        issues.append(
            _issue(
                "ZERO_GRAND_TOTAL",
                "WARNING",
                None,
                "Grand Total Facing",
                0,
                f"Grand Total Facing is zero for {dimensions}; SOS_% is blank.",
            )
        )
    return pd.DataFrame(issues, columns=VALIDATION_COLUMNS)


def write_excel_output(
    raw_clean: pd.DataFrame,
    sos_summary: pd.DataFrame,
    validation_report: pd.DataFrame,
    output_path: str | Path,
) -> Path:
    """Write the normalized data, SOS summary, and validation report."""
    destination = Path(output_path)
    destination.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(destination, engine="openpyxl") as writer:
        raw_clean.to_excel(writer, sheet_name="RAW_CLEAN", index=False)
        sos_summary.to_excel(writer, sheet_name="SOS_SUMMARY", index=False)
        validation_report.reindex(columns=VALIDATION_COLUMNS).to_excel(
            writer,
            sheet_name="VALIDATION_REPORT",
            index=False,
        )

    return destination


def _fatal_validation_messages(validation_report: pd.DataFrame) -> list[str]:
    if validation_report.empty:
        return []
    fatal = validation_report[validation_report["Severity"].eq("ERROR")]
    return fatal["Message"].astype(str).tolist()


def run_pipeline(input_path: str | Path, output_path: str | Path) -> Path:
    """Run one local SOS processing job."""
    LOGGER.info("Reading input: %s", input_path)
    source = reader(input_path)
    LOGGER.info("Rows read: %s", f"{len(source):,}")

    validation_report = validator(source)
    fatal_messages = _fatal_validation_messages(validation_report)
    if fatal_messages:
        raise ValueError("Validation failed: " + " | ".join(fatal_messages))

    raw_clean = transformer(source)
    group_cols = ["Year", "Month", "Region"]
    if "Divisi" in raw_clean.columns:
        group_cols.append("Divisi")

    LOGGER.info("Calculating SOS by: %s", ", ".join(group_cols))
    summary_source = raw_clean[
        raw_clean["Region"].notna()
        & raw_clean["Region"].astype("string").str.strip().ne("")
    ]
    sos_summary = calculate_sos(summary_source, group_cols)
    validation_report = pd.concat(
        [
            validation_report,
            _zero_total_issues(sos_summary, group_cols),
        ],
        ignore_index=True,
    ).reindex(columns=VALIDATION_COLUMNS)

    destination = write_excel_output(
        raw_clean,
        sos_summary,
        validation_report,
        output_path,
    )
    LOGGER.info("Summary rows: %s", f"{len(sos_summary):,}")
    LOGGER.info("Validation issues: %s", f"{len(validation_report):,}")
    LOGGER.info("Output created: %s", destination.resolve())
    return destination


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Calculate Indofood Share of Shelf from CSV or XLSX raw data."
    )
    parser.add_argument("input_path", help="Path to the input .csv or .xlsx file.")
    parser.add_argument(
        "--output",
        default="output_sos_summary.xlsx",
        help="Output workbook path (default: output_sos_summary.xlsx).",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="%(levelname)s: %(message)s",
    )
    args = build_parser().parse_args(argv)

    try:
        run_pipeline(args.input_path, args.output)
    except Exception as error:
        LOGGER.error("%s", error)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
