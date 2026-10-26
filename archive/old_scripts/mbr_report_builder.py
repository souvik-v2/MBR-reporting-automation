"""
MBR Report Builder for Sprint MBR reporting.

Builds the final MBR Excel report with:
- Consolidated sprint data
- KPI summary
- Sprint trends
- Team comparison
- Quality metrics
- Goals achieved (from PPTs)
- Run log
"""
from __future__ import annotations
from typing import List, Dict, Any
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

import config
from chart_generator import (
    create_sprint_trend_chart,
    create_team_comparison_chart,
    create_planned_vs_actual_chart,
    create_quality_metrics_chart,
    create_rework_trend_chart,
)


def apply_header_style(ws, row_num: int = 1):
    """Apply styling to header row."""
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    
    for cell in ws[row_num]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")


def apply_border(ws, start_row: int, end_row: int, start_col: int, end_col: int):
    """Apply borders to a range of cells."""
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    for row in ws.iter_rows(min_row=start_row, max_row=end_row,
                            min_col=start_col, max_col=end_col):
        for cell in row:
            cell.border = thin_border


def format_percentage(value: Any) -> str:
    """Format a decimal value as percentage string."""
    try:
        return f"{float(value) * 100:.1f}%"
    except (TypeError, ValueError):
        return "0.0%"


def create_consolidated_sheet(wb: Workbook, df: pd.DataFrame, sheet_name: str):
    """
    Create consolidated data sheet with all sprint data.
    
    Args:
        wb: Workbook object
        df: DataFrame containing consolidated sprint data
        sheet_name: Name for the sheet
    """
    ws = wb.create_sheet(sheet_name)
    
    # Write data
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            ws.cell(row=r_idx, column=c_idx, value=value)
    
    # Apply styling
    apply_header_style(ws, 1)
    if len(df) > 0:
        apply_border(ws, 1, len(df) + 1, 1, len(df.columns))
    
    # Auto-adjust column widths
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column_letter].width = adjusted_width


def create_kpi_summary_sheet(wb: Workbook, df: pd.DataFrame, sheet_name: str):
    """
    Create KPI summary sheet with aggregated metrics.
    
    Args:
        wb: Workbook object
        df: DataFrame containing sprint data with calculated KPIs
        sheet_name: Name for the sheet
    """
    ws = wb.create_sheet(sheet_name)
    
    # Calculate summary statistics
    summary_data = []

    if not df.empty:
        # Convert to numeric, coercing errors to NaN
        total_planned = pd.to_numeric(df["Planned_SP"], errors='coerce').sum() if "Planned_SP" in df.columns else 0
        total_delivered = pd.to_numeric(df["Actual_Delivered"], errors='coerce').sum() if "Actual_Delivered" in df.columns else 0
        avg_completion = pd.to_numeric(df["Completion_Ratio"], errors='coerce').mean() if "Completion_Ratio" in df.columns else 0
        avg_lt_reopen = pd.to_numeric(df["LT_Reopen_Percentage"], errors='coerce').mean() if "LT_Reopen_Percentage" in df.columns else 0
        avg_lt_carryover = pd.to_numeric(df["LT_Carryover_Percentage"], errors='coerce').mean() if "LT_Carryover_Percentage" in df.columns else 0
        
        summary_data = [
            ["KPI", "Value"],
            ["Total Planned SP", total_planned],
            ["Total Delivered SP", total_delivered],
            ["Average Completion Ratio", format_percentage(avg_completion)],
            ["Average LT Reopen %", format_percentage(avg_lt_reopen)],
            ["Average LT Carryover %", format_percentage(avg_lt_carryover)],
            ["Total Sprints", len(df)],
        ]
    else:
        summary_data = [["KPI", "Value"], ["No data available", ""]]
    
    # Write summary data
    for r_idx, row in enumerate(summary_data, 1):
        for c_idx, value in enumerate(row, 1):
            ws.cell(row=r_idx, column=c_idx, value=value)
    
    # Apply styling
    apply_header_style(ws, 1)
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 20


def create_sprint_trends_sheet(wb: Workbook, df: pd.DataFrame, sheet_name: str):
    """Create sprint trends sheet with chart."""
    ws = wb.create_sheet(sheet_name)

    if df.empty:
        ws.cell(1, 1, "No sprint data available")
        return

    # Prepare data for charting
    trend_cols = ["Sprint", "Planned_SP", "Actual_Delivered", "Completion_Ratio"]
    trend_df = df[trend_cols].copy() if all(col in df.columns for col in trend_cols) else df

    # Write data
    for r_idx, row in enumerate(dataframe_to_rows(trend_df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            ws.cell(row=r_idx, column=c_idx, value=value)

    apply_header_style(ws, 1)

    # Add chart
    if len(trend_df) > 1 and config.GENERATE_CHARTS:
        chart = create_sprint_trend_chart(ws, 1, len(trend_df) + 1)
        ws.add_chart(chart, "F2")


def create_goals_achieved_sheet(wb: Workbook, goals_data: List[Dict], sheet_name: str):
    """Create goals achieved sheet from extracted PPT slides."""
    ws = wb.create_sheet(sheet_name)

    if not goals_data:
        ws.cell(1, 1, "No goals data extracted from presentations")
        return

    # Headers
    headers = ["Source File", "Slide Number", "Month", "Title", "Content"]
    for c_idx, header in enumerate(headers, 1):
        ws.cell(1, c_idx, header)

    apply_header_style(ws, 1)

    # Data rows
    row_idx = 2
    for goal in goals_data:
        ws.cell(row_idx, 1, goal.get("__source_file", ""))
        ws.cell(row_idx, 2, goal.get("Slide_Number", ""))
        ws.cell(row_idx, 3, goal.get("Month", ""))
        ws.cell(row_idx, 4, goal.get("Title", ""))

        # Combine text content
        text_content = "\n\n".join(goal.get("Text_Content", []))
        ws.cell(row_idx, 5, text_content)

        row_idx += 1

    # Auto-adjust columns
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 40
    ws.column_dimensions['E'].width = 60


def build_mbr_report(
    output_path: str,
    consolidated_df: pd.DataFrame,
    goals_data: List[Dict],
    log: List[str]
):
    """
    Build the complete MBR report Excel file.

    Args:
        output_path: Path to save the output Excel file
        consolidated_df: DataFrame with consolidated sprint data and calculated KPIs
        goals_data: List of extracted goal slides from presentations
        log: Log messages from processing
    """
    wb = Workbook()

    # Remove default sheet
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])

    # Create all sheets
    create_consolidated_sheet(wb, consolidated_df, config.MBR_SHEET_NAMES["consolidated"])
    create_kpi_summary_sheet(wb, consolidated_df, config.MBR_SHEET_NAMES["kpi_summary"])
    create_sprint_trends_sheet(wb, consolidated_df, config.MBR_SHEET_NAMES["sprint_trends"])
    create_goals_achieved_sheet(wb, goals_data, config.MBR_SHEET_NAMES["goals_achieved"])

    # Add run log sheet
    log_ws = wb.create_sheet(config.MBR_SHEET_NAMES["run_log"])
    log_ws.cell(1, 1, "Log Entry")
    apply_header_style(log_ws, 1)
    for idx, entry in enumerate(log, 2):
        log_ws.cell(idx, 1, entry)
    log_ws.column_dimensions['A'].width = 100

    # Save workbook
    wb.save(output_path)
