#!/usr/bin/env python3
"""
Download National Bridge Inventory (NBI) delimited data for years 1992-2025,
add a 'year' column to each dataset, and combine all years into a single Parquet file.

URL patterns:
  - 2018-2025: {year}hwybronefiledel.zip (single all-states file)
  - 1992-2017: {year}del.zip (per-state files, same schema)

Files are comma-delimited with single-quote (') text qualifiers.
"""

import io
import os
import sys
import time
import zipfile

import pandas as pd
import pyarrow as pa
import pyarrow.parquet as pq
import requests

BASE_URL = "https://www.fhwa.dot.gov/bridge/nbi"
START_YEAR = 1992
END_YEAR = 2025
OUTPUT_DIR = "nbi_data"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "nbi_all_years.parquet")
TEMP_DIR = os.path.join(OUTPUT_DIR, "yearly_parquet")

# Years where the single-file ZIP is available
SINGLE_FILE_YEARS = set(range(2018, END_YEAR + 1))


def download_with_retry(url: str, max_retries: int = 4, timeout: int = 300) -> bytes:
    """Download a URL with exponential backoff retries."""
    for attempt in range(max_retries + 1):
        try:
            resp = requests.get(url, timeout=timeout)
            resp.raise_for_status()
            return resp.content
        except (requests.RequestException, IOError) as e:
            if attempt == max_retries:
                raise
            wait = 2 ** (attempt + 1)
            print(f"  Retry {attempt + 1}/{max_retries} after {wait}s: {e}")
            time.sleep(wait)


def get_zip_url(year: int) -> str:
    """Return the download URL for a given year."""
    if year in SINGLE_FILE_YEARS:
        return f"{BASE_URL}/{year}hwybronefiledel.zip"
    return f"{BASE_URL}/{year}del.zip"


def read_year_from_zip(content: bytes, year: int) -> pd.DataFrame:
    """Read all CSV files from a ZIP archive and return a single DataFrame."""
    zf = zipfile.ZipFile(io.BytesIO(content))
    txt_files = [n for n in zf.namelist() if n.lower().endswith(".txt")]
    if not txt_files:
        raise ValueError(f"No .txt files found in ZIP for year {year}")

    frames = []
    for i, name in enumerate(sorted(txt_files)):
        raw = zf.read(name)
        # Decode as latin-1 to handle any byte values
        df = pd.read_csv(
            io.StringIO(raw.decode("latin-1")),
            sep=",",
            quotechar="'",
            dtype=str,          # keep everything as string to avoid type issues
            low_memory=False,
            on_bad_lines="warn",
        )
        frames.append(df)
        if (i + 1) % 10 == 0:
            print(f"    Read {i + 1}/{len(txt_files)} files...", flush=True)

    combined = pd.concat(frames, ignore_index=True)
    return combined


def process_year(year: int) -> str:
    """Download, parse, and save one year of NBI data as a Parquet file.
    Returns the path to the saved Parquet file."""
    out_path = os.path.join(TEMP_DIR, f"nbi_{year}.parquet")
    if os.path.exists(out_path):
        print(f"  [{year}] Already processed, skipping.", flush=True)
        return out_path

    url = get_zip_url(year)
    print(f"  [{year}] Downloading {url} ...", flush=True)
    content = download_with_retry(url)
    print(f"  [{year}] Downloaded {len(content) / 1e6:.1f} MB", flush=True)

    print(f"  [{year}] Parsing CSV ...", flush=True)
    df = read_year_from_zip(content, year)
    df.insert(0, "year", str(year))
    print(f"  [{year}] {len(df):,} rows, {len(df.columns)} columns", flush=True)

    df.to_parquet(out_path, engine="pyarrow", index=False)
    size_mb = os.path.getsize(out_path) / 1e6
    print(f"  [{year}] Saved {out_path} ({size_mb:.1f} MB)", flush=True)

    del df  # free memory
    return out_path


def build_union_schema(parquet_paths: list[str]) -> pa.Schema:
    """Build a union schema from all Parquet files, preserving column order."""
    seen = set()
    ordered_fields = []

    for path in sorted(parquet_paths):
        schema = pq.read_schema(path)
        for field in schema:
            if field.name not in seen:
                seen.add(field.name)
                ordered_fields.append(pa.field(field.name, pa.string()))

    return pa.schema(ordered_fields)


def align_table_to_schema(table: pa.Table, target_schema: pa.Schema) -> pa.Table:
    """Add missing columns (as null) and reorder to match the target schema."""
    columns = {}
    for field in target_schema:
        if field.name in table.column_names:
            columns[field.name] = table.column(field.name).cast(pa.string())
        else:
            columns[field.name] = pa.nulls(table.num_rows, type=pa.string())

    return pa.table(columns, schema=target_schema)


def combine_parquet_files(parquet_paths: list[str], output_path: str):
    """Combine multiple Parquet files into a single file using PyArrow.
    Handles differing schemas by building a union schema with all columns."""
    print(f"\nCombining {len(parquet_paths)} yearly files into {output_path} ...", flush=True)

    print("  Building union schema ...", flush=True)
    union_schema = build_union_schema(parquet_paths)
    print(f"  Union schema has {len(union_schema)} columns", flush=True)

    writer = pq.ParquetWriter(output_path, union_schema, compression="snappy")
    total_rows = 0

    for path in sorted(parquet_paths):
        table = pq.read_table(path)
        aligned = align_table_to_schema(table, union_schema)
        writer.write_table(aligned)
        total_rows += table.num_rows
        print(f"  Added {os.path.basename(path)}: {table.num_rows:,} rows", flush=True)
        del table, aligned

    writer.close()

    size_mb = os.path.getsize(output_path) / 1e6
    print(f"\nDone! Combined file: {output_path}", flush=True)
    print(f"  Total rows: {total_rows:,}", flush=True)
    print(f"  File size:  {size_mb:.1f} MB", flush=True)


def main():
    os.makedirs(TEMP_DIR, exist_ok=True)

    years = list(range(START_YEAR, END_YEAR + 1))
    print(f"Processing NBI data for {len(years)} years ({START_YEAR}-{END_YEAR})\n", flush=True)

    parquet_paths = []
    for year in years:
        try:
            path = process_year(year)
            parquet_paths.append(path)
        except Exception as e:
            print(f"  [{year}] ERROR: {e}", flush=True)
            continue

    if parquet_paths:
        combine_parquet_files(parquet_paths, OUTPUT_FILE)
    else:
        print("No data was downloaded successfully.", flush=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
