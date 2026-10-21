"""
Remove Duplicate Rows from MahaRERA Merged Files by Considering All Columns.

Features:
- Considers all data columns (ignoring source tracking metadata like source_file/source_row).
- Identifies exact duplicate rows across all fields.
- Preserves the first occurrence and removes redundant duplicates.
- Exports both deduplicated CSV (UTF-8 with BOM) and Excel (.xlsx) files.
"""

from pathlib import Path
import csv
import pandas as pd
import openpyxl

FOLDER = Path(r"D:\Mumbai\Thane Output files")

# Determine input file (supporting Thane_merged_csv.csv or maharera_merged.csv)
INPUT_CSV = FOLDER / "Thane_merged_csv.csv"
if not INPUT_CSV.exists():
    INPUT_CSV = FOLDER / "maharera_merged.csv"

OUTPUT_DEDUP_CSV = FOLDER / "Thane_merged_deduplicated.csv"
OUTPUT_DEDUP_XLSX = FOLDER / "Thane_merged_deduplicated.xlsx"
OUTPUT_REPORT_CSV = FOLDER / "duplicates_removed_report.csv"


def remove_duplicates(input_csv: Path, output_csv: Path, output_xlsx: Path):
    if not input_csv.exists():
        raise FileNotFoundError(f"Input file not found: {input_csv}")

    print(f"Reading merged file: {input_csv.name}...")
    df = pd.read_csv(input_csv, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    initial_count = len(df)
    print(f"Total rows before deduplication: {initial_count:,}")

    # Columns considered for deduplication (all data columns, ignoring file metadata)
    metadata_cols = {"source_file", "source_row"}
    dedup_cols = [c for c in df.columns if c not in metadata_cols]
    print(f"Deduplicating across all {len(dedup_cols)} data columns...")

    # Identify duplicate rows
    dup_mask = df.duplicated(subset=dedup_cols, keep="first")
    dup_rows = df[dup_mask]
    num_duplicates = len(dup_rows)

    print(f"\nDuplicates found: {num_duplicates:,}")

    if num_duplicates > 0:
        print("\nSummary of removed duplicate records:")
        removed_summary = []
        for reg_no, group in dup_rows.groupby("Registration Number"):
            project_name = group["Project Name"].iloc[0] if "Project Name" in group else ""
            source_file = group["source_file"].iloc[0] if "source_file" in group else ""
            print(f"  - Reg No: {reg_no} | Project: {project_name} | From: {source_file} ({len(group)} duplicate(s))")
            for idx, r in group.iterrows():
                removed_summary.append({
                    "Registration Number": reg_no,
                    "Project Name": project_name,
                    "Source File": r.get("source_file", ""),
                    "Source Row": r.get("source_row", ""),
                })
        
        # Save report of removed duplicates
        pd.DataFrame(removed_summary).to_csv(OUTPUT_REPORT_CSV, index=False, encoding="utf-8-sig")
        print(f"Duplicate removal log saved to: {OUTPUT_REPORT_CSV.name}")

    # Drop duplicates
    df_clean = df.drop_duplicates(subset=dedup_cols, keep="first").reset_index(drop=True)
    final_count = len(df_clean)
    print(f"\nTotal rows after deduplication: {final_count:,} (Removed: {num_duplicates:,})")

    # 1. Export Clean CSV
    print(f"\nSaving clean CSV to: {output_csv.name}...")
    df_clean.to_csv(output_csv, index=False, encoding="utf-8-sig", quoting=csv.QUOTE_MINIMAL)
    print("CSV saved successfully.")

    # 2. Export Clean Excel
    print(f"Saving clean Excel to: {output_xlsx.name}...")
    with pd.ExcelWriter(output_xlsx, engine="openpyxl") as writer:
        df_clean.to_excel(writer, sheet_name="MahaRERA_Clean", index=False)
    print("Excel saved successfully.")

    print("\n================ DEDUPLICATION COMPLETE ================")
    print(f"Original Records:     {initial_count:,}")
    print(f"Duplicates Removed:   {num_duplicates:,}")
    print(f"Remaining Unique:     {final_count:,}")
    print(f"Clean CSV Output:     {output_csv}")
    print(f"Clean Excel Output:   {output_xlsx}")
    print("=========================================================")


if __name__ == "__main__":
    remove_duplicates(INPUT_CSV, OUTPUT_DEDUP_CSV, OUTPUT_DEDUP_XLSX)
