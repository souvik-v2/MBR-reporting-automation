"""
Main entry point for the Excel/CSV/PPTX Merger & Sprint MBR Excel
Reporting tool.

Usage:
    python main.py [source_folder] [output_file] [dashboard_file]

If arguments are omitted, the paths in config.py are used.

This builds the merged master workbook, HTML dashboard, and (if
config.ENABLE_MBR_REPORTING is True) an Excel-based Sprint MBR KPI report
with charts (mbr_report.xlsx) -- all from the same merged tabular data.

NOTE: this is unrelated to the separate MBR PowerPoint-deck pipeline
(build_mbr.py). That pipeline updates a monthly MBR .pptx deck's charts
and consolidates team status slides; it has its own command-line
interface. Run it directly: `python build_mbr.py --help`. See
MBR_PIPELINE.md for details.
"""
from __future__ import annotations

import sys
from pathlib import Path
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))

from src.merger import config
from src.merger.readers_tabular import read_excel_file, read_csv_file
from src.merger.readers_pptx import read_pptx_file
from src.merger.dedup import dedup_latest
from src.merger.writer import write_master_workbook
from src.merger.dashboard import write_dashboard
from src.merger.metrics_calculator import calculate_metrics_for_dataframe
from src.merger.pptx_slide_extractor import extract_goal_slides
from src.merger.mbr_report_builder import build_mbr_report


def discover_files(folder: Path):
    excel_files, csv_files, pptx_files = [], [], []
    for path in sorted(folder.rglob("*")):
        if path.name.startswith("~$") or not path.is_file():
            continue
        suffix = path.suffix.lower()
        if suffix in config.EXCEL_EXTENSIONS:
            excel_files.append(path)
        elif suffix in config.CSV_EXTENSIONS:
            csv_files.append(path)
        elif suffix in config.PPTX_EXTENSIONS:
            pptx_files.append(path)
    return excel_files, csv_files, pptx_files


def main():
    source_folder = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(config.SOURCE_FOLDER)
    output_file = Path(sys.argv[2]) if len(sys.argv) > 2 else Path(config.OUTPUT_FILE)
    dashboard_file = Path(sys.argv[3]) if len(sys.argv) > 3 else Path(config.DASHBOARD_FILE)
    output_file.parent.mkdir(parents=True, exist_ok=True)
    dashboard_file.parent.mkdir(parents=True, exist_ok=True)

    log: list = []

    if not source_folder.exists():
        print(f"ERROR: source folder '{source_folder}' does not exist.")
        sys.exit(1)

    excel_files, csv_files, pptx_files = discover_files(source_folder)
    log.append(
        f"[INFO] Found {len(excel_files)} Excel file(s), {len(csv_files)} CSV/TSV file(s), "
        f"{len(pptx_files)} PPTX file(s) in '{source_folder}'."
    )

    all_records = []

    for path in excel_files:
        recs = read_excel_file(path, log)
        all_records.extend(recs)
        log.append(f"[INFO] Read {len(recs)} row(s) from Excel file '{path.name}'.")

    for path in csv_files:
        recs = read_csv_file(path, log)
        all_records.extend(recs)
        log.append(f"[INFO] Read {len(recs)} row(s) from CSV/TSV file '{path.name}'.")

    pptx_text_rows, pptx_chart_rows = [], []
    for path in pptx_files:
        table_recs, text_rows, chart_rows = read_pptx_file(path, log)
        all_records.extend(table_recs)
        pptx_text_rows.extend(text_rows)
        pptx_chart_rows.extend(chart_rows)
        log.append(
            f"[INFO] Read {len(table_recs)} table row(s), {len(text_rows)} text block(s), "
            f"{len(chart_rows)} chart data point(s) from PPTX file '{path.name}'."
        )

    if not all_records and not pptx_text_rows and not pptx_chart_rows:
        log.append("[WARN] No data found anywhere — check SOURCE_FOLDER in config.py.")

    log.append(
        f"[INFO] Total tabular rows collected before dedup: {len(all_records)}."
    )

    master_records = dedup_latest(all_records, log) if all_records else []
    log.append(f"[INFO] Master sheet row count after dedup: {len(master_records)}.")

    write_master_workbook(
        str(output_file),
        master_records=master_records,
        all_records=all_records,
        pptx_text_rows=pptx_text_rows,
        pptx_chart_rows=pptx_chart_rows,
        log=log,
        id_col=config.DEDUP_ID_COLUMN,
        date_col=config.DEDUP_DATE_COLUMN,
    )

    write_dashboard(
        str(dashboard_file),
        master_records=master_records,
        all_records=all_records,
        pptx_text_rows=pptx_text_rows,
        pptx_chart_rows=pptx_chart_rows,
        log=log,
        id_col=config.DEDUP_ID_COLUMN,
        date_col=config.DEDUP_DATE_COLUMN,
    )

    # ── Sprint MBR Reporting ──────────────────────────────────────────────
    if config.ENABLE_MBR_REPORTING:
        log.append("[INFO] Starting Sprint MBR reporting...")

        # Extract goal slides from presentations
        goal_slides = []
        for path in pptx_files:
            slides = extract_goal_slides(path, log=log)
            goal_slides.extend(slides)

        log.append(f"[INFO] Extracted {len(goal_slides)} goal achievement slides from presentations.")

        # Convert consolidated records to DataFrame for KPI calculation
        if all_records:
            df = pd.DataFrame(all_records)
            log.append(f"[INFO] Converting {len(df)} records to DataFrame for KPI calculation.")

            # Calculate metrics
            df_with_metrics = calculate_metrics_for_dataframe(df)
            log.append("[INFO] KPI calculations completed.")

            # Build MBR report
            mbr_output = Path(config.MBR_REPORT_FILE)
            mbr_output.parent.mkdir(parents=True, exist_ok=True)

            build_mbr_report(
                str(mbr_output),
                consolidated_df=df_with_metrics,
                goals_data=goal_slides,
                log=log
            )

            log.append(f"[INFO] MBR report written to: {mbr_output}")
            print(f"\nMBR Report written to: {mbr_output}")
            print(f"  Goal slides extracted: {len(goal_slides)}")
        else:
            log.append("[WARN] No data available for MBR reporting.")
            print("\nWARN: No data available for MBR reporting.")

    print(f"\nDone. Master file written to: {output_file}")
    print(f"Dashboard written to: {dashboard_file}")
    print(f"  Master rows (deduped, latest only): {len(master_records)}")
    print(f"  All_Records rows (everything found): {len(all_records)}")
    print(f"  PPTX text blocks: {len(pptx_text_rows)}")
    print(f"  PPTX chart data points: {len(pptx_chart_rows)}")
    print("  See the Run_Log sheet inside the output file for full details.")


if __name__ == "__main__":
    # NOTE: --mbr-workflow used to call mbr_workflow.py, which depended on
    # a Windows-COM-only PowerPoint consolidator that was never finished
    # (see archive/old_scripts/ and archive/old_docs/ and MBR_PIPELINE.md for why). That workflow has
    # been replaced by build_mbr.py, which has its own CLI (different
    # required arguments: --template, --graph-dir, --ppt-data-dir, --month,
    # --out) since it builds a PowerPoint deck rather than an Excel report.
    if len(sys.argv) > 1 and sys.argv[1] == "--mbr-workflow":
        print(
            "The old --mbr-workflow flag pointed at a Windows-COM-only "
            "PowerPoint automation that was never completed. It has been "
            "replaced by build_mbr.py, which has its own command-line "
            "interface. Run it directly instead, e.g.:\n\n"
            "  python build_mbr.py --template <last month's deck.pptx> "
            "--graph-dir <Graph folder> --ppt-data-dir <ppt-data folder> "
            "--month <Month> --out <output.pptx>\n\n"
            "See MBR_PIPELINE.md for full documentation."
        )
        sys.exit(1)
    else:
        main()
