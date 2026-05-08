# build_submitter_db.py
# =====================
# Reads a bepress/Digital Commons inventory Excel file and produces three CSVs:
#   - submitters.csv        : deduplicated submitter registry with unique
#                             submitter_id primary keys
#   - records.csv           : original records with the submitted_by column removed
#   - record_submitters.csv : junction table — one row per record-submitter link
#
# These three files form a normalised relational structure:
#   records  <--  record_submitters  -->  submitters
#
# Usage:
#     python build_submitter_db.py --input inventory.xlsx --outdir ./output
#
# Requirements:
#     pip install pandas openpyxl
#
# Notes:
# - submitted_by contains an email address (e.g. jane.smith@wright.edu).
# - Deduplication is case-insensitive on the email address.
# - Submitter IDs are assigned alphabetically by email: SU000001, SU000002, ...
# - 115 rows have a blank submitted_by — those rows produce no junction entry.
 
import argparse
import gc
import os
import sys
 
import pandas as pd
 
 
# ── Configuration ─────────────────────────────────────────────────────────────
 
# The column that identifies who submitted each record
SUBMITTER_COL = "submitted_by"
 
# The stable record identifier used as the join key across all three tables
RECORD_KEY    = "manuscript_id"
 
 
# ── Helper functions ──────────────────────────────────────────────────────────
 
def clean(val) -> str:
    # Return val as a stripped, lowercased string, or '' if null/NaN.
    if val is None:
        return ""
    if not isinstance(val, str) and pd.isna(val):
        return ""
    return str(val).strip().lower()
 
 
def parse_email_parts(email: str) -> dict:
    # Split an email address into logical parts for the registry.
    # e.g. "jane.smith@wright.edu"
    #   -> { "username": "jane.smith", "domain": "wright.edu",
    #        "display_name": "jane.smith" }
    # If the address has no '@', treat the whole string as the username.
    email = email.strip()
    if "@" in email:
        username, domain = email.split("@", 1)
    else:
        username, domain = email, ""
    return {
        "email":        email,
        "username":     username,
        "domain":       domain,
        # display_name: replace dots in username with spaces for readability
        # e.g. "jane.smith" -> "jane smith"
        "display_name": username.replace(".", " ").title(),
    }
 
 
# ── Main logic ────────────────────────────────────────────────────────────────
 
def build(input_path: str, outdir: str) -> None:
    os.makedirs(outdir, exist_ok=True)
 
    # ── Discover columns ──────────────────────────────────────────────────────
    print(f"Reading column headers from: {input_path}")
    all_cols     = list(pd.read_excel(input_path, nrows=0).columns)
    non_sub_cols = [c for c in all_cols if c != SUBMITTER_COL]
 
    print(f"  {len(all_cols)} total columns")
    print(f"  Submitter column : '{SUBMITTER_COL}'")
    print(f"  Record key       : '{RECORD_KEY}'")
 
    # ── Pass 1: Build the submitter registry ──────────────────────────────────
    # Read only the two columns we need — much faster than loading all 938.
    print(f"\nPass 1: Loading '{SUBMITTER_COL}' data...")
    df_sub = pd.read_excel(
        input_path,
        usecols=[RECORD_KEY, SUBMITTER_COL],
        dtype=str
    ).fillna("")
    print(f"  {len(df_sub):,} rows loaded. Building registry...")
 
    # registry maps lowercase email -> parsed email dict
    registry: dict[str, dict] = {}
 
    for _, row in df_sub.iterrows():
        email = clean(row[SUBMITTER_COL])
        if email and email not in registry:
            registry[email] = parse_email_parts(email)
 
    print(f"  Registry complete: {len(registry):,} unique submitters")
 
    # ── Assign submitter IDs (sorted alphabetically by email) ─────────────────
    sorted_emails = sorted(registry.keys())
    # SU000001 ... SU000757
    email_to_id = {
        email: f"SU{str(i + 1).zfill(6)}"
        for i, email in enumerate(sorted_emails)
    }
 
    # ── Save submitters.csv ───────────────────────────────────────────────────
    print("\nBuilding submitters table...")
    submitter_rows = []
    for email in sorted_emails:
        parts = registry[email]
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
 
    # ── Pass 2: Build the record_submitters junction table ────────────────────
    # One row per record that has a non-blank submitted_by value.
    print("\nPass 2: Building record_submitters junction table...")
 
    junction_rows = []
 
    for _, row in df_sub.iterrows():
        manuscript_id = clean(row[RECORD_KEY])
        email         = clean(row[SUBMITTER_COL])
 
        # Skip rows with no submitter or no record key
        if not email or not manuscript_id:
            continue
 
        if email in email_to_id:
            junction_rows.append({
                "manuscript_id": manuscript_id,
                "submitter_id":  email_to_id[email],
            })
 
    junction_df   = pd.DataFrame(junction_rows)
    junction_path = os.path.join(outdir, "record_submitters.csv")
    junction_df.to_csv(junction_path, index=False)
    print(f"  Saved: {junction_path}  ({len(junction_df):,} rows)")
 
    del df_sub, junction_df
    gc.collect()
 
    # ── Load non-submitter record columns and save records.csv ────────────────
    # Records now contain no submitted_by column — link via record_submitters.csv.
    print("\nLoading non-submitter record columns...")
    df_records   = pd.read_excel(input_path, usecols=non_sub_cols, dtype=str).fillna("")
    records_path = os.path.join(outdir, "records.csv")
    print(f"  Saving records ({len(df_records):,} rows × {len(df_records.columns)} cols)...")
    df_records.to_csv(records_path, index=False)
    print(f"  Saved: {records_path}")
 
    del df_records
    gc.collect()
 
    print("\n✓ Done.")
    print(f"  Submitters        : {submitters_path}")
    print(f"  Records           : {records_path}")
    print(f"  Record-Submitters : {junction_path}")
    print("\nJoin example (pandas):")
    print("  import pandas as pd")
    print("  records           = pd.read_csv('records.csv')")
    print("  record_submitters = pd.read_csv('record_submitters.csv')")
    print("  submitters        = pd.read_csv('submitters.csv')")
    print("  # All records submitted by a given person:")
    print("  merged = pd.merge(record_submitters, submitters, on='submitter_id')")
    print("  merged[merged['email'] == 'jane.wildermuth@wright.edu']")
    print("\nNext step: run  python format_submitters_excel.py  to produce a styled Excel file.")
 
 
# ── CLI entry point ───────────────────────────────────────────────────────────
 
if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Extract a normalised submitter registry from a bepress inventory Excel file."
    )
    parser.add_argument(
        "--input", "-i",
        required=True,
        help="Path to the inventory .xlsx file"
    )
    parser.add_argument(
        "--outdir", "-o",
        default="./output",
        help="Directory to write submitters.csv, records.csv, and record_submitters.csv (default: ./output)"
    )
    args = parser.parse_args()
 
    if not os.path.exists(args.input):
        print(f"Error: input file not found: {args.input}", file=sys.stderr)
        sys.exit(1)
 
    build(args.input, args.outdir)
