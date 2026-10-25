"""Download the UCI Online Retail II dataset and convert it to a single CSV.

Source: https://archive.ics.uci.edu/dataset/502/online+retail+ii
License: CC BY 4.0

The original file is an Excel workbook with one sheet per period. Reading
.xlsx with pandas is slow (minutes for ~1M rows), so this script does it once
and writes a CSV that later steps can load in seconds.

No cleaning happens here: rows, values and column names are kept exactly as
they come from the source. Cleaning decisions belong to the analysis stage,
where they are documented and justified.

Usage (from the project root, with the virtual environment active):
    python src/download_data.py
"""

from pathlib import Path
import urllib.request
import zipfile

import pandas as pd

DATA_URL = "https://archive.ics.uci.edu/static/public/502/online+retail+ii.zip"

PROJECT_ROOT = Path(__file__).resolve().parent.parent
RAW_DIR = PROJECT_ROOT / "data" / "raw"
ZIP_PATH = RAW_DIR / "online_retail_ii.zip"
CSV_PATH = RAW_DIR / "online_retail_ii.csv"


def download_zip() -> None:
    """Download the zip archive unless it is already present."""
    if ZIP_PATH.exists():
        print(f"Zip already present, skipping download: {ZIP_PATH.name}")
        return
    print(f"Downloading {DATA_URL} ...")
    urllib.request.urlretrieve(DATA_URL, ZIP_PATH)
    print(f"Saved {ZIP_PATH.name} ({ZIP_PATH.stat().st_size / 1e6:.1f} MB)")


def extract_xlsx() -> Path:
    """Extract the Excel workbook from the zip and return its path."""
    with zipfile.ZipFile(ZIP_PATH) as zf:
        xlsx_names = [n for n in zf.namelist() if n.lower().endswith(".xlsx")]
        if len(xlsx_names) != 1:
            raise RuntimeError(f"Expected one .xlsx in the zip, found: {xlsx_names}")
        zf.extract(xlsx_names[0], RAW_DIR)
        return RAW_DIR / xlsx_names[0]


def convert_to_csv(xlsx_path: Path) -> None:
    """Read every sheet, stack them, and write one CSV. Prints a short report."""
    print("Reading workbook (this can take a few minutes) ...")
    sheets = pd.read_excel(xlsx_path, sheet_name=None)  # dict: sheet name -> DataFrame

    print("\nSheets found:")
    for name, df in sheets.items():
        print(f"  - {name!r}: {len(df):,} rows, columns = {list(df.columns)}")

    combined = pd.concat(sheets.values(), ignore_index=True)
    combined.to_csv(CSV_PATH, index=False)
    print(f"\nWrote {CSV_PATH.name}: {len(combined):,} rows x {combined.shape[1]} columns")


def main() -> None:
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    if CSV_PATH.exists():
        print(f"CSV already exists, nothing to do: {CSV_PATH}")
        return
    download_zip()
    xlsx_path = extract_xlsx()
    convert_to_csv(xlsx_path)


if __name__ == "__main__":
    main()
