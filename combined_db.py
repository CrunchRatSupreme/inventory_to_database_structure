# build_all_dbs.py
# ================
# Combined script. Reads the inventory Excel file ONCE and produces five CSVs:
#
#   authors.csv          : deduplicated author registry  (primary key: author_id)
#   submitters.csv       : deduplicated submitter registry (primary key: submitter_id)
#   record_authors.csv   : junction table — one row per record-author link
#   record_submitters.csv: junction table — one row per record-submitter link
#   records.csv          : all other record columns, with author AND submitter
#                          columns removed
#
# Relational structure:
#
#   records  <--  record_authors   -->  authors
#   records  <--  record_submitters --> submitters
#
#   Join key for all four junction/registry tables: manuscript_id
#
# Usage:
#     python build_all_dbs.py --input inventory.xlsx --outdir ./output
#
# Requirements:
#     pip install pandas openpyxl
#
# Why one script?
#   Running build_author_db.py and build_submitter_db.py separately each writes
#   its own records.csv, stripping only its own columns. The second run would
#   overwrite the first. This script strips ALL normalised columns in a single
#   pass so records.csv is definitive and complete.

import argparse
import gc
import json
import os
import sys

import pandas as pd


# ── Configuration ─────────────────────────────────────────────────────────────

# Suffixes that mark individual author sub-field columns (author1_fname, etc.)
AUTHOR_SUBFIELDS = ("_fname", "_mname", "_lname", "_suffix", "_email", "_institution")

# Maximum number of numbered author slots in the inventory
MAX_AUTHORS = 132

# Column holding the submitter email address
SUBMITTER_COL = "submitted_by"

# Stable record identifier — used as the join key in all junction tables
RECORD_KEY = "manuscript_id"


# ═════════════════════════════════════════════════════════════════════════════
# SHARED HELPER
# ═════════════════════════════════════════════════════════════════════════════

def clean(val) -> str:
    # Return val as a stripped string, or '' if null/NaN.
    # Used by both the author and submitter sections.
    if val is None:
        return ""
    if not isinstance(val, str) and pd.isna(val):
        return ""
    return str(val).strip()


# ═════════════════════════════════════════════════════════════════════════════
# AUTHOR HELPERS  (ported from build_author_db.py)
# ═════════════════════════════════════════════════════════════════════════════

def parse_authors_from_row(row: dict) -> list[dict]:
    # Extract a list of author dicts from a single inventory row.
    # Tries the 'authors' JSON field first; falls back to individual
    # author[N]_* columns. Returns a list of dicts with keys:
    #   fname, mname, lname, suffix, email, institution,
    #   corporate_name, is_corporate (bool)
    results = []

    # Primary source: 'authors' JSON column
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
                        "fname":          clean(e.get("FNAME", "")),
                        "mname":          clean(e.get("MNAME", "")),
                        "lname":          clean(e.get("LNAME", "")),
                        "suffix":         clean(e.get("SUFFIX", "")),
                        "email":          clean(e.get("EMAIL", "")),
                        "institution":    clean(e.get("INSTITUTION", "")),
                        "corporate_name": "",
                        "is_corporate":   False,
                    })
            if results:
                return results
        except (json.JSONDecodeError, TypeError):
            pass  # fall through to individual fields

    # Fallback: individual author[N]_* columns
    for n in range(1, MAX_AUTHORS + 1):
        lname = clean(row.get(f"author{n}_lname", ""))
        fname = clean(row.get(f"author{n}_fname", ""))
        if not lname and not fname:
            break  # slots are sequential; stop at first empty slot
        results.append({
            "fname":          fname,
            "mname":          clean(row.get(f"author{n}_mname", "")),
            "lname":          lname,
            "suffix":         clean(row.get(f"author{n}_suffix", "")),
            "email":          clean(row.get(f"author{n}_email", "")),
            "institution":    clean(row.get(f"author{n}_institution", "")),
            "corporate_name": "",
            "is_corporate":   False,
        })

    return results


def make_author_dedup_key(author: dict) -> tuple:
    # Return a hashable fingerprint for an author dict.
    # Corporate authors keyed on name; individuals on last/first/middle.
    if author["is_corporate"]:
        return ("__corp__", "", "", author["corporate_name"].lower())
    return (
        author["lname"].lower(),
        author["fname"].lower(),
        author["mname"].lower(),
        "",
    )


def make_author_display_name(author: dict) -> str:
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


# ═════════════════════════════════════════════════════════════════════════════
# SUBMITTER HELPERS  (ported from build_submitter_db.py)
# ═════════════════════════════════════════════════════════════════════════════

def parse_email_parts(email: str) -> dict:
    # Split an email address into logical parts for the submitter registry.
    # e.g. "jane.smith@wright.edu"
    #   -> { "username": "jane.smith", "domain": "wright.edu",
    #        "display_name": "Jane Smith" }
    email = email.strip()
    if "@" in email:
        username, domain = email.split("@", 1)
    else:
        username, domain = email, ""
    return {
        "email":        email,
        "username":     username,
        "domain":       domain,
        # Replace dots in username with spaces, then title-case
        # e.g. "jane.smith" -> "Jane Smith"
        "display_name": username.replace(".", " ").title(),
    }


# ═════════════════════════════════════════════════════════════════════════════
# MAIN BUILD FUNCTION
# ═════════════════════════════════════════════════════════════════════════════

def build(input_path: str, outdir: str) -> None:
    os.makedirs(outdir, exist_ok=True)

    # ── Step 1: Discover all column names ─────────────────────────────────────
    # Read only the header row (nrows=0) — fast, no data loaded yet.
    print(f"Reading column headers from: {input_path}")
    all_cols = list(pd.read_excel(input_path, nrows=0).columns)

    # All author detail columns: author1_fname, author2_lname, etc.
    author_detail_cols = [
        c for c in all_cols
        if c.startswith("author") and any(c.endswith(s) for s in AUTHOR_SUBFIELDS)
    ]

    # Columns needed to build both author registries and junction tables
    author_read_cols = [RECORD_KEY, "authors"] + author_detail_cols

    # Columns needed to build the submitter registry and junction table
    submitter_read_cols = [RECORD_KEY, SUBMITTER_COL]

    # Columns to keep in records.csv — everything except author details,
    # the authors JSON blob, and the submitted_by column
    cols_to_drop = set(author_detail_cols) | {"authors", SUBMITTER_COL}
    record_cols  = [c for c in all_cols if c not in cols_to_drop]

    print(f"  {len(all_cols):,} total columns in inventory")
    print(f"  {len(author_detail_cols):,} author detail columns will be removed")
    print(f"  1 'authors' JSON column will be removed")
    print(f"  1 '{SUBMITTER_COL}' column will be removed")
    print(f"  {len(record_cols):,} columns will remain in records.csv")


    # ═════════════════════════════════════════════════════════════════════════
    # PASS 1 — AUTHORS
    # Read author columns, build registry, assign IDs, write authors.csv and
    # record_authors.csv. Then free the memory before moving on.
    # ═════════════════════════════════════════════════════════════════════════

    print("\n── Pass 1: Authors ──────────────────────────────────────────────────")
    print("  Loading author columns...")
    df_auth = pd.read_excel(
        input_path, usecols=author_read_cols, dtype=str
    ).fillna("")
    print(f"  {len(df_auth):,} rows loaded.")

    # Build the author deduplication registry
    # Key: tuple fingerprint  →  Value: best author dict seen so far
    author_registry: dict[tuple, dict] = {}

    for i, row in df_auth.iterrows():
        for author in parse_authors_from_row(row):
            key = make_author_dedup_key(author)
            if key not in author_registry:
                author_registry[key] = author.copy()
            else:
                # Enrich existing entry with any new fields this occurrence has
                existing = author_registry[key]
                for field in ("email", "institution", "suffix"):
                    if not existing[field] and author[field]:
                        existing[field] = author[field]

        if (i + 1) % 20000 == 0:
            print(f"  {i+1:,} rows scanned | {len(author_registry):,} unique authors so far")

    print(f"  Author registry complete: {len(author_registry):,} unique authors")

    # Assign author IDs sorted alphabetically by last/first name
    author_sorted_keys = sorted(author_registry.keys(), key=lambda k: (k[0], k[1], k[2]))
    author_key_to_id   = {
        k: f"AU{str(i + 1).zfill(6)}"
        for i, k in enumerate(author_sorted_keys)
    }

    # Write authors.csv
    author_rows = []
    for key in author_sorted_keys:
        a   = author_registry[key]
        aid = author_key_to_id[key]
        author_rows.append({
            "author_id":      aid,
            "display_name":   make_author_display_name(a),
            "last_name":      a["lname"],
            "first_name":     a["fname"],
            "middle_name":    a["mname"],
            "suffix":         a["suffix"],
            "is_corporate":   "Yes" if a["is_corporate"] else "No",
            "corporate_name": a["corporate_name"],
            "email":          a["email"],
            "institution":    a["institution"],
        })

    authors_df   = pd.DataFrame(author_rows)
    authors_path = os.path.join(outdir, "authors.csv")
    authors_df.to_csv(authors_path, index=False)
    print(f"  Saved: {authors_path}  ({len(authors_df):,} rows)")

    # Build and write record_authors.csv
    # One row per record-author pair, preserving author order.
    junction_auth_rows = []
    for _, row in df_auth.iterrows():
        manuscript_id = clean(row[RECORD_KEY])
        for order, author in enumerate(parse_authors_from_row(row), start=1):
            key = make_author_dedup_key(author)
            if key in author_key_to_id:
                junction_auth_rows.append({
                    "manuscript_id": manuscript_id,
                    "author_id":     author_key_to_id[key],
                    "author_order":  order,
                })

    junction_auth_df   = pd.DataFrame(junction_auth_rows)
    junction_auth_path = os.path.join(outdir, "record_authors.csv")
    junction_auth_df.to_csv(junction_auth_path, index=False)
    print(f"  Saved: {junction_auth_path}  ({len(junction_auth_df):,} rows)")

    # Free author data — no longer needed
    del df_auth, authors_df, junction_auth_df, author_rows, junction_auth_rows
    gc.collect()


    # ═════════════════════════════════════════════════════════════════════════
    # PASS 2 — SUBMITTERS
    # Read the submitted_by column, build registry, assign IDs, write
    # submitters.csv and record_submitters.csv. Then free the memory.
    # ═════════════════════════════════════════════════════════════════════════

    print("\n── Pass 2: Submitters ───────────────────────────────────────────────")
    print("  Loading submitter column...")
    df_sub = pd.read_excel(
        input_path, usecols=submitter_read_cols, dtype=str
    ).fillna("")
    print(f"  {len(df_sub):,} rows loaded.")

    # Build submitter registry: lowercase email -> parsed email dict
    submitter_registry: dict[str, dict] = {}

    for _, row in df_sub.iterrows():
        email = clean(row[SUBMITTER_COL])
        if email and email not in submitter_registry:
            submitter_registry[email] = parse_email_parts(email)

    print(f"  Submitter registry complete: {len(submitter_registry):,} unique submitters")

    # Assign submitter IDs sorted alphabetically by email
    sorted_emails  = sorted(submitter_registry.keys())
    email_to_id    = {
        email: f"SU{str(i + 1).zfill(6)}"
        for i, email in enumerate(sorted_emails)
    }

    # Write submitters.csv
    submitter_rows = []
    for email in sorted_emails:
        parts = submitter_registry[email]
        submitter_rows.append({
            "submitter_id": email_to_id[email],
            "email":        parts["email"],
            "display_name": parts["display_name"],
            "username":     parts["username"],
            "domain":       parts["domain"],
        })

    submitters_df   = pd.DataFrame(submitter_rows)
    submitters_path = os.path.join(outdir, "submitters.csv")
    submitters_df.to_csv(submitters_path, index=False)
    print(f"  Saved: {submitters_path}  ({len(submitters_df):,} rows)")

    # Build and write record_submitters.csv
    junction_sub_rows = []
    for _, row in df_sub.iterrows():
        manuscript_id = clean(row[RECORD_KEY])
        email         = clean(row[SUBMITTER_COL])
        if manuscript_id and email and email in email_to_id:
            junction_sub_rows.append({
                "manuscript_id": manuscript_id,
                "submitter_id":  email_to_id[email],
            })

    junction_sub_df   = pd.DataFrame(junction_sub_rows)
    junction_sub_path = os.path.join(outdir, "record_submitters.csv")
    junction_sub_df.to_csv(junction_sub_path, index=False)
    print(f"  Saved: {junction_sub_path}  ({len(junction_sub_df):,} rows)")

    # Free submitter data — no longer needed
    del df_sub, submitters_df, junction_sub_df, submitter_rows, junction_sub_rows
    gc.collect()


    # ═════════════════════════════════════════════════════════════════════════
    # PASS 3 — RECORDS
    # Load only the columns that belong in records.csv — everything except
    # author detail columns, the authors JSON, and submitted_by.
    # This is the single definitive records file for the whole project.
    # ═════════════════════════════════════════════════════════════════════════

    print("\n── Pass 3: Records ──────────────────────────────────────────────────")
    print("  Loading record columns (author and submitter columns excluded)...")
    df_records   = pd.read_excel(input_path, usecols=record_cols, dtype=str).fillna("")
    records_path = os.path.join(outdir, "records.csv")
    print(f"  Saving records ({len(df_records):,} rows × {len(df_records.columns)} cols)...")
    df_records.to_csv(records_path, index=False)
    print(f"  Saved: {records_path}")

    del df_records
    gc.collect()


    # ── Summary ───────────────────────────────────────────────────────────────
    print("\n✓ All done.")
    print(f"  authors.csv           : {authors_path}")
    print(f"  submitters.csv        : {submitters_path}")
    print(f"  record_authors.csv    : {junction_auth_path}")
    print(f"  record_submitters.csv : {junction_sub_path}")
    print(f"  records.csv           : {records_path}")
    print()
    print("Join examples (pandas):")
    print("  import pandas as pd")
    print("  records            = pd.read_csv('records.csv')")
    print("  record_authors     = pd.read_csv('record_authors.csv')")
    print("  authors            = pd.read_csv('authors.csv')")
    print("  record_submitters  = pd.read_csv('record_submitters.csv')")
    print("  submitters         = pd.read_csv('submitters.csv')")
    print()
    print("  # All authors for a given record:")
    print("  pd.merge(record_authors, authors, on='author_id')")
    print()
    print("  # All records submitted by a given person:")
    print("  merged = pd.merge(record_submitters, submitters, on='submitter_id')")
    print("  merged[merged['email'] == 'jane.wildermuth@wright.edu']")
    print()
    print("  # Full picture: record + its authors + its submitter:")
    print("  r_a  = pd.merge(records, record_authors,    on='manuscript_id')")
    print("  r_a  = pd.merge(r_a,     authors,           on='author_id')")
    print("  r_as = pd.merge(r_a,     record_submitters, on='manuscript_id')")
    print("  r_as = pd.merge(r_as,    submitters,        on='submitter_id')")


# ── CLI entry point ───────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description=(
            "Build normalised author, submitter, junction, and record tables "
            "from a bepress/Digital Commons inventory Excel file."
        )
    )
    parser.add_argument(
        "--input", "-i",
        required=True,
        help="Path to the inventory .xlsx file"
    )
    parser.add_argument(
        "--outdir", "-o",
        default="./output",
        help="Directory to write all five CSVs (default: ./output)"
    )
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"Error: input file not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    build(args.input, args.outdir)