"""Export a worksheet table from an Excel workbook to CSV.

Example:
    python excel_table_to_csv.py

For another workbook/table, either change the configuration below or use:
    python excel_table_to_csv.py --excel-file path/to/file.xlsx --table 2 \
        --output-dir path/to/output --output-name my_table.csv

Requires:
    pip install openpyxl
"""

from __future__ import annotations

import argparse
import csv
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook


# Change these lines when running the script for a different default export.
BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "../data/raw/AusPlay-By-Sport-Data-Tables-30-April-2026.xlsx"
TABLE_NUMBER = "4.1"
OUTPUT_DIR = BASE_DIR / "../data/finalized"
OUTPUT_NAME = "2.2.2UpdatedOrganisingEntity.csv"

# Set this to a 1-based Excel row number if the table has a header row that
# cannot be detected automatically. Leave as None for automatic detection.
HEADER_ROW: int | None = None


def non_empty_values(values: Iterable[object]) -> list[str]:
    """Return cell values as trimmed strings, excluding empty cells."""
    return [str(value).strip() for value in values if value is not None and str(value).strip()]


def find_header_row(rows: list[tuple[object, ...]]) -> int:
    """Find the first likely header row in a worksheet."""
    candidates: list[tuple[int, int]] = []
    for row_number, row in enumerate(rows[:30], start=1):
        values = non_empty_values(row)
        contains_number = any(isinstance(value, (int, float)) for value in row)
        if len(values) >= 2 and not contains_number:
            candidates.append((len(values), row_number))
    if candidates:
        return max(candidates, key=lambda candidate: (candidate[0], -candidate[1]))[1]
    raise ValueError(
        "Could not detect a header row. Set HEADER_ROW or pass --header-row."
    )


def make_unique_headers(headers: tuple[object, ...]) -> list[str]:
    """Create non-empty, unique CSV column names from Excel headers."""
    result: list[str] = []
    used: dict[str, int] = {}
    for index, value in enumerate(headers, start=1):
        name = str(value).strip() if value is not None else ""
        name = name or ("activity" if index == 1 else f"column_{index}")
        count = used.get(name, 0) + 1
        used[name] = count
        result.append(name if count == 1 else f"{name}_{count}")
    return result


def exclude_column(
    headers: list[str],
    data_rows: list[list[object]],
    exclude: list[int],
) -> tuple[list[str], list[list[object]]]:
    """
    Exclude specified columns (1-based indices) from headers and data rows.
    Example: exclude=[2,4] will drop the 2nd and 4th columns.
    """
    # Convert to 0-based indices
    exclude_set = {i - 1 for i in exclude}
    new_headers = [h for i, h in enumerate(headers) if i not in exclude_set]
    new_data = [
        [v for i, v in enumerate(row) if i not in exclude_set]
        for row in data_rows
    ]
    return new_headers, new_data


def export_table(
    excel_file: Path,
    table_number: str,
    output_file: Path,
    header_row: int | None = None,
) -> int:
    """Export the worksheet named by table_number and return its row count."""
    if not excel_file.is_file():
        raise FileNotFoundError(f"Excel file not found: {excel_file}")

    workbook = load_workbook(excel_file, read_only=True, data_only=True)
    if table_number not in workbook.sheetnames:
        available = ", ".join(workbook.sheetnames)
        raise ValueError(
            f"Table/sheet '{table_number}' was not found. Available sheets: {available}"
        )

    worksheet = workbook[table_number]
    rows = list(worksheet.iter_rows(values_only=True))
    if not rows:
        raise ValueError(f"Table/sheet '{table_number}' is empty.")

    header_index = header_row or find_header_row(rows)
    if not 1 <= header_index <= len(rows):
        raise ValueError(f"Header row must be between 1 and {len(rows)}.")

    raw_headers = rows[header_index - 1]
    last_header = max(
        index for index, value in enumerate(raw_headers) if value is not None and str(value).strip()
    )
    headers = make_unique_headers(raw_headers[: last_header + 1])
    data_rows = []
    for row in rows[header_index:]:
        values = list(row[: last_header + 1])
        if values and values[0] is not None and str(values[0]).strip():
            data_rows.append(
                ["" if value is None else value for value in values]
            )

    # Example usage: exclude columns 2 through 6. Indices are 1-based.
    #headers, data_rows = exclude_column(headers, data_rows, exclude=[2, 3, 4, 5, 6])

    output_file.parent.mkdir(parents=True, exist_ok=True)
    with output_file.open("w", newline="", encoding="utf-8-sig") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(headers)
        writer.writerows(data_rows)

    workbook.close()
    return len(data_rows)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Export an Excel worksheet table to a UTF-8 CSV file."
    )
    parser.add_argument("--excel-file", type=Path, default=EXCEL_FILE)
    parser.add_argument("--table", default=TABLE_NUMBER, help="Worksheet/table number, e.g. 4")
    parser.add_argument("--output-dir", type=Path, default=OUTPUT_DIR)
    parser.add_argument("--output-name", default=OUTPUT_NAME)
    parser.add_argument(
        "--header-row",
        type=int,
        default=HEADER_ROW,
        help="1-based Excel header row; omit to detect it automatically.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    output_file = args.output_dir / args.output_name
    row_count = export_table(
        args.excel_file,
        args.table,
        output_file,
        args.header_row,
    )
    print(f"Exported {row_count} data rows from table {args.table} to {output_file}")


if __name__ == "__main__":
    main()
