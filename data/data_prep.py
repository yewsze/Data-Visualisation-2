"""
Convert every .xlsx in data/raw/ into CSVs under data/clean/.

  data/raw/AusPlay-National-Data-Tables-30-April-2026.xlsx
  -> data/clean/AusPlay-National-Data-Tables-30-April-2026/Table_1.csv
                                                           Table_2.csv ...

Folder name = original workbook name; one CSV per sheet.
Cleaning is light and structure-preserving (header rows are kept, because
each sheet is laid out differently); chart-specific reshaping happens later.
Run from the project root:  python prep/convert_xlsx.py
"""
import re
import warnings
from pathlib import Path

import pandas as pd

warnings.filterwarnings("ignore")

RAW_DIR = Path("data/raw")
CLEAN_DIR = Path("data/clean")


def sheet_to_filename(sheet: str) -> str:
    """'1' -> 'Table_1', 'Table 1.1' -> 'Table_1.1', 'Index' -> 'Index'."""
    s = str(sheet).strip()
    if re.fullmatch(r"\d+", s):
        return f"Table_{s}"
    m = re.fullmatch(r"Table\s+([\d.]+)", s, flags=re.I)
    if m:
        return f"Table_{m.group(1)}"
    return re.sub(r"[^\w.-]+", "_", s)          # e.g. 'Further information'


def clean_frame(df: pd.DataFrame) -> pd.DataFrame:
    # strip stray whitespace in text cells; blank strings -> empty
    df = df.apply(lambda col: col.map(lambda v: v.strip() if isinstance(v, str) else v))
    df = df.replace("", pd.NA)
    # drop rows/columns that are completely empty
    df = df.dropna(how="all").dropna(axis=1, how="all")
    return df.reset_index(drop=True)


def convert_workbook(xlsx: Path) -> None:
    out_dir = CLEAN_DIR / xlsx.stem
    out_dir.mkdir(parents=True, exist_ok=True)

    sheets = pd.read_excel(xlsx, sheet_name=None, header=None)   # dict of all sheets
    for name, df in sheets.items():
        df = clean_frame(df)
        out = out_dir / f"{sheet_to_filename(name)}.csv"
        df.to_csv(out, index=False, header=False, encoding="utf-8")
    print(f"{xlsx.name}: {len(sheets)} sheets -> {out_dir}")


if __name__ == "__main__":
    files = sorted(RAW_DIR.glob("*.xlsx"))
    if not files:
        raise SystemExit(f"No .xlsx files found in {RAW_DIR.resolve()}")
    for f in files:
        convert_workbook(f)