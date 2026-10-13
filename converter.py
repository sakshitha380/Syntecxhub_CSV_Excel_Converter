
import argparse
import logging
import re
import sys
from pathlib import Path

import pandas as pd


# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(levelname)s: %(message)s"
)
logger = logging.getLogger(__name__)


def normalize_column_name(name):
    """Normalize column names into lowercase snake_case."""
    name = str(name).strip().lower()
    name = re.sub(r"[^\w\s]", "", name)
    name = re.sub(r"\s+", "_", name)
    return name


def parse_rename(value):
    """Parse column renaming arguments in old:new format."""
    if ":" not in value:
        raise argparse.ArgumentTypeError(
            "Use --rename in the format old:new"
        )

    old, new = value.split(":", 1)
    old, new = old.strip(), new.strip()

    if not old or not new:
        raise argparse.ArgumentTypeError(
            "Both old and new column names are required."
        )

    return old, new


def convert_csv(
    input_path,
    output_path,
    date_columns=None,
    renames=None,
    fill_missing=None
):
    """Read, clean, and export CSV data to an Excel workbook."""
    input_file = Path(input_path)
    output_file = Path(output_path)

    if not input_file.is_file():
        raise FileNotFoundError(
            f"Input file not found: {input_file}"
        )

    if input_file.suffix.lower() != ".csv":
        raise ValueError("Input file must have a .csv extension.")

    if output_file.suffix.lower() != ".xlsx":
        raise ValueError("Output file must have an .xlsx extension.")

    if input_file.resolve() == output_file.resolve():
        raise ValueError("Input and output paths must be different.")

    try:
        df = pd.read_csv(input_file, encoding="utf-8-sig")
    except pd.errors.EmptyDataError as exc:
        raise ValueError("The CSV file is empty.") from exc
    except UnicodeDecodeError as exc:
        raise ValueError(
            "Unable to read the CSV. Save it using UTF-8 encoding."
        ) from exc
    except pd.errors.ParserError as exc:
        raise ValueError(f"Invalid CSV structure: {exc}") from exc

    if len(df.columns) == 0:
        raise ValueError("The CSV file has no columns.")

    # Normalize column names.
    df.columns = [normalize_column_name(c) for c in df.columns]

    if any(not column for column in df.columns):
        raise ValueError("A column name is empty after normalization.")

    if df.columns.duplicated().any():
        raise ValueError(
            "Duplicate column names found after normalization."
        )

    # Trim text values and treat empty strings as missing.
    for column in df.select_dtypes(include=["object", "string"]).columns:
        df[column] = df[column].map(
            lambda value: value.strip()
            if isinstance(value, str)
            else value
        )
        df[column] = df[column].replace("", pd.NA)

    # Rename columns.
    for old, new in (renames or {}).items():
        old_name = normalize_column_name(old)
        new_name = normalize_column_name(new)

        if old_name not in df.columns:
            raise ValueError(
                f"Column '{old}' not found. "
                f"Available columns: {', '.join(df.columns)}"
            )

        if not new_name:
            raise ValueError("New column name cannot be empty.")

        if new_name != old_name and new_name in df.columns:
            raise ValueError(
                f"Column '{new_name}' already exists."
            )

        df = df.rename(columns={old_name: new_name})

    # Parse specified dates or automatically detect date columns.
    if date_columns:
        selected_dates = [
            normalize_column_name(column)
            for column in date_columns
        ]
    else:
        selected_dates = [
            column for column in df.columns
            if "date" in column
        ]

    for column in selected_dates:
        if column not in df.columns:
            raise ValueError(f"Date column '{column}' not found.")

        original_values = df[column].notna()
        parsed_dates = pd.to_datetime(
            df[column],
            errors="coerce",
            format="mixed"
        )

        invalid_count = int(
            (original_values & parsed_dates.isna()).sum()
        )

        if invalid_count:
            logger.warning(
                "%s invalid date value(s) in '%s' were left blank.",
                invalid_count,
                column
            )

        df[column] = parsed_dates

    # Optionally replace missing values.
    if fill_missing is not None:
        df = df.fillna(fill_missing)

    # Create the output directory when necessary.
    output_file.parent.mkdir(parents=True, exist_ok=True)

    try:
        with pd.ExcelWriter(
            output_file,
            engine="openpyxl",
            date_format="YYYY-MM-DD",
            datetime_format="YYYY-MM-DD"
        ) as writer:
            df.to_excel(
                writer,
                index=False,
                sheet_name="Cleaned Data"
            )
    except PermissionError as exc:
        raise PermissionError(
            "Cannot save the file. Close it in Excel if it is open."
        ) from exc
    except (OSError, ValueError, ImportError) as exc:
        raise RuntimeError(f"Excel export failed: {exc}") from exc

    logger.info("Conversion completed successfully.")
    logger.info("Rows: %s | Columns: %s", len(df), len(df.columns))
    logger.info("Output saved to: %s", output_file.resolve())

    return df


def main():
    parser = argparse.ArgumentParser(
        description="Convert and clean CSV files into Excel workbooks."
    )

    parser.add_argument(
        "--input",
        required=True,
        help="Path to the input CSV file."
    )
    parser.add_argument(
        "--output",
        required=True,
        help="Path to the output .xlsx file."
    )
    parser.add_argument(
        "--date-columns",
        nargs="+",
        help="Optional date column names."
    )
    parser.add_argument(
        "--rename",
        action="append",
        type=parse_rename,
        default=[],
        metavar="OLD:NEW",
        help="Rename a column; this option can be repeated."
    )
    parser.add_argument(
        "--fill-missing",
        help="Optional replacement for missing values, such as N/A."
    )

    args = parser.parse_args()

    try:
        convert_csv(
            input_path=args.input,
            output_path=args.output,
            date_columns=args.date_columns,
            renames=dict(args.rename),
            fill_missing=args.fill_missing
        )
    except (
        FileNotFoundError,
        ValueError,
        PermissionError,
        RuntimeError
    ) as exc:
        logger.error("%s", exc)
        sys.exit(1)


if __name__ == "__main__":
    main()
