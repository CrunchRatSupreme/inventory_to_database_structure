# build_author_db.py
# ==================
# Step 1 of 2.
#
# Reads a bepress/Digital Commons inventory Excel file and produces three CSVs:
#   - authors.csv        : deduplicated author registry with unique author_id primary keys
#   - records.csv        : original records with all author columns removed
#   - record_authors.csv : junction table — one row per record-author link,
#                          with author_order preserving the original sequence
#
# These three files form a normalised relational structure:
#   records  <--  record_authors  -->  authors
#   (1 record can have many authors; 1 author can appear on many records)
#
# Usage:
#     python build_author_db.py --input inventory.xlsx --outdir ./output
#
# Requirements:
#     pip install pandas openpyxl
#
# Notes:
# - The inventory must have an 'authors' JSON column and/or individual
#   author[N]_fname / author[N]_lname / ... columns (up to author132).
# - The 'authors' JSON field is used as the primary source. Individual
#   author[N]_* columns are the fallback when the JSON is absent or empty.
# - Author deduplication is based on (last_name, first_name, middle_name)
#   for individuals, and (corporate_name,) for corporate authors — all
#   lowercased. Where duplicates exist, the richest record (most fields
#   populated) is kept.
# - Author IDs are assigned alphabetically by last/first name: AU000001,
#   AU000002, ...

import argparse
import gc
import json
import os
import sys

import pandas as pd


# ── Configuration ─────────────────────────────────────────────────────────────

# Suffixes that identify individual author sub-fields
AUTHOR_SUBFIELDS = ("_fname", "_mname", "_lname", "_suffix", "_email", "_institution")

# Maximum number of author slots in the inventory
MAX_AUTHORS = 132


# ── Helper functions ──────────────────────────────────────────────────────────

def clean(val) -> str:
    # Return val as a stripped string, or '' if null/NaN.
    if val is None:
        return ""
    if not isinstance(val, str) and pd.isna(val):
        return ""
    return str(val).strip()


def parse_authors_from_row(row: dict) -> list[dict]:
    # Extract a list of author dicts from a single inventory row.
    # Tries the 'authors' JSON field first; falls back to individual
    # author[N]_* columns. Returns a list of dicts with keys:
    #     fname, mname, lname, suffix, email, institution,
    #     corporate_name, is_corporate (bool)
    results = []

    # ── Primary source: 'authors' JSON column ──
    json_str = row.get("authors", "")
    if json_str and json_str != "[]":
        try:
            entries = json.loads(json_str)
            for e in entries:
                if e.get("IS_CORPORATE_AUTHOR", 0) == 1:
                    results.append({
                        "fname": "", "mname": "", "lname": "",
                        "suffix": "", "email": "", "institution": "",
                        "corporate_name": clean(e.get("CORPORATE_AUTHOR", "")),
                        "is_corporate": True,
                    })
                else:
                    results.append({
                        "fname":         clean(e.get("FNAME", "")),
                        "mname":         clean(e.get("MNAME", "")),
                        "lname":         clean(e.get("LNAME", "")),
                        "suffix":        clean(e.get("SUFFIX", "")),
                        "email":         clean(e.get("EMAIL", "")),
                        "institution":   clean(e.get("INSTITUTION", "")),
                        "corporate_name": "",
                        "is_corporate":  False,
                    })
            if results:
                return results
        except (json.JSONDecodeError, TypeError):
            pass  # fall through to individual fields

    # ── Fallback: individual author[N]_* columns ──
    for n in range(1, MAX_AUTHORS + 1):
        lname = clean(row.get(f"author{n}_lname", ""))
        fname = clean(row.get(f"author{n}_fname", ""))
        if not lname and not fname:
            break  # slots are sequential; stop at first empty slot
        results.append({
            "fname":         fname,
            "mname":         clean(row.get(f"author{n}_mname", "")),
            "lname":         lname,
            "suffix":        clean(row.get(f"author{n}_suffix", "")),
            "email":         clean(row.get(f"author{n}_email", "")),
            "institution":   clean(row.get(f"author{n}_institution", "")),
            "corporate_name": "",
            "is_corporate":  False,
        })

    return results


def make_dedup_key(author: dict) -> tuple:
    # Return a hashable deduplication key for an author dict.
    # Corporate authors are keyed on their name; individuals on last/first/middle.
    if author["is_corporate"]:
        return ("__corp__", "", "", author["corporate_name"].lower())
    return (
        author["lname"].lower(),
        author["fname"].lower(),
        author["mname"].lower(),
        "",
    )


def make_display_name(author: dict) -> str:
    # Return a 'Last, First Middle Suffix' display string.
    if author["is_corporate"]:
        return author["corporate_name"]
    parts = []
    if author["lname"]:
        parts.append(author["lname"] + ",")
    if author["fname"]:
        parts.append(author["fname"])
    if author["mname"]:
        parts.append(author["mname"])
    if author["suffix"]:
        parts.append(author["suffix"])
    return " ".join(parts).strip().rstrip(",")


# ── Main logic ────────────────────────────────────────────────────────────────

def build(input_path: str, outdir: str) -> None:
    os.makedirs(outdir, exist_ok=True)

    # ── Discover columns ──────────────────────────────────────────────────────
    print(f"Reading column headers from: {input_path}")
    all_cols = list(pd.read_excel(input_path, nrows=0).columns)

    author_detail_cols = [
        c for c in all_cols
        if c.startswith("author") and any(c.endswith(s) for s in AUTHOR_SUBFIELDS)
    ]
    author_read_cols = ["authors", "manuscript_id"] + author_detail_cols
    non_author_cols  = [c for c in all_cols if c not in author_detail_cols and c != "authors"]

    print(f"  {len(all_cols)} total columns")
    print(f"  {len(author_detail_cols)} author detail columns (author1_fname … author{MAX_AUTHORS}_institution)")
    print(f"  {len(non_author_cols)} non-author columns to keep in records output")

    # ── Pass 1: Build author registry ─────────────────────────────────────────
    print("\nPass 1: Loading author data...")
    df_authors = pd.read_excel(input_path, usecols=author_read_cols, dtype=str).fillna("")
    print(f"  {len(df_authors):,} rows loaded. Building registry...")

    registry: dict[tuple, dict] = {}  # dedup_key → best author dict

    for i, row in df_authors.iterrows():
        for author in parse_authors_from_row(row):
            key = make_dedup_key(author)
            if key not in registry:
                registry[key] = author.copy()
            else:
                # Merge: fill in any missing fields from this occurrence
                existing = registry[key]
                for field in ("email", "institution", "suffix"):
                    if not existing[field] and author[field]:
                        existing[field] = author[field]

        if (i + 1) % 20000 == 0:
            print(f"  {i+1:,} rows processed | {len(registry):,} unique authors so far")

    print(f"\n  Registry complete: {len(registry):,} unique authors")

    # ── Assign author IDs (sorted alphabetically) ─────────────────────────────
    sorted_keys = sorted(registry.keys(), key=lambda k: (k[0], k[1], k[2]))
    key_to_id   = {k: f"AU{str(i+1).zfill(6)}" for i, k in enumerate(sorted_keys)}

    # ── Save authors.csv ──────────────────────────────────────────────────────
    print("\nBuilding authors table...")
    author_rows = []
    for key in sorted_keys:
        a   = registry[key]
        aid = key_to_id[key]
        author_rows.append({
            "author_id":      aid,
            "display_name":   make_display_name(a),
            "last_name":      a["lname"],
            "first_name":     a["fname"],
            "middle_name":    a["mname"],
            "suffix":         a["suffix"],
            "is_corporate":   "Yes" if a["is_corporate"] else "No",
            "corporate_name": a["corporate_name"],
            "email":          a["email"],
            "institution":    a["institution"],
        })

    authors_df = pd.DataFrame(author_rows)
    authors_path = os.path.join(outdir, "authors.csv")
    authors_df.to_csv(authors_path, index=False)
    print(f"  Saved: {authors_path}  ({len(authors_df):,} rows)")

    # ── Pass 2: Build the record_authors junction table ───────────────────────
    # Instead of cramming all author IDs into one cell, we create one row per
    # record-author pair. author_order records the position of each author in
    # the original author list (1 = first author, 2 = second, etc.).
    print("\nPass 2: Building record_authors junction table...")

    junction_rows = []  # will become record_authors.csv

    for i, row in df_authors.iterrows():
        # manuscript_id is our stable record identifier
        manuscript_id = clean(row.get("manuscript_id", ""))

        authors = parse_authors_from_row(row)
        for order, author in enumerate(authors, start=1):
            # start=1 so the first author gets author_order=1, not 0
            key = make_dedup_key(author)
            if key in key_to_id:
                junction_rows.append({
                    "manuscript_id": manuscript_id,
                    "author_id":     key_to_id[key],
                    "author_order":  order,
                })

        if (i + 1) % 20000 == 0:
            print(f"  {i+1:,} rows processed")

    junction_df   = pd.DataFrame(junction_rows)
    junction_path = os.path.join(outdir, "record_authors.csv")
    junction_df.to_csv(junction_path, index=False)
    print(f"  Saved: {junction_path}  ({len(junction_df):,} rows)")

    del df_authors, junction_df
    gc.collect()

    # ── Load non-author record columns and save records.csv ───────────────────
    # Records now contain no author data at all — link to authors via
    # record_authors.csv using manuscript_id as the join key.
    print("\nLoading non-author record columns...")
    df_records = pd.read_excel(input_path, usecols=non_author_cols, dtype=str).fillna("")

    records_path = os.path.join(outdir, "records.csv")
    print(f"  Saving records ({len(df_records):,} rows × {len(df_records.columns)} cols)...")
    df_records.to_csv(records_path, index=False)
    print(f"  Saved: {records_path}")

    del df_records
    gc.collect()

    print("\n✓ Done.")
    print(f"  Authors        : {authors_path}")
    print(f"  Records        : {records_path}")
    print(f"  Record-Authors : {junction_path}")
    print("\nJoin example (pandas):")
    print("  import pandas as pd")
    print("  records       = pd.read_csv('records.csv')")
    print("  record_authors = pd.read_csv('record_authors.csv')")
    print("  authors       = pd.read_csv('authors.csv')")
    print("  # All authors for a record:")
    print("  pd.merge(record_authors, authors, on='author_id')")
    print("\nNext step: run  python format_authors_excel.py  to produce a styled Excel file.")


# ── CLI entry point ───────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Extract a normalised author registry from a bepress inventory Excel file."
    )
    parser.add_argument(
        "--input", "-i",
        required=True,
        help="Path to the inventory .xlsx file"
    )
    parser.add_argument(
        "--outdir", "-o",
        default="./output",
        help="Directory to write authors.csv, records.csv, and record_authors.csv (default: ./output)"
    )
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"Error: input file not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    build(args.input, args.outdir)