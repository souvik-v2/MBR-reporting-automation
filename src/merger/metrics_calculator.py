"""
Metrics calculator for Sprint MBR reporting.

Calculates KPIs:
- Reopened % = Reopened SP / Delivered SP
- Carryover % = Carryover / Planned SP
- Completion Ratio = Actual / Planned
- QA Pass Rate
- V2 vs LT metrics
"""
from __future__ import annotations
from typing import List, Dict, Any
import pandas as pd


def safe_float(value: Any, default: float = 0.0) -> float:
    """Safely convert a value to float."""
    try:
        if value is None or pd.isna(value) or value == "":
            return default
        # Try to extract numbers from strings (e.g., "5 points" -> 5)
        if isinstance(value, str):
            import re
            numbers = re.findall(r'[-+]?\d*\.?\d+', value)
            if numbers:
                return float(numbers[0])
            return default
        return float(value)
    except (TypeError, ValueError):
        return default


def safe_divide(numerator: float, denominator: float, default: float = 0.0) -> float:
    """Safely divide two numbers, returning default if denominator is 0."""
    try:
        if denominator == 0 or pd.isna(denominator):
            return default
        return numerator / denominator
    except (TypeError, ZeroDivisionError):
        return default


def calculate_reopened_percentage(reopened_sp: float, delivered_sp: float) -> float:
    """
    Calculate Reopened % = Reopened SP / Delivered SP
    
    Args:
        reopened_sp: Story points that were reopened
        delivered_sp: Story points that were delivered
    
    Returns:
        Percentage as decimal (e.g., 0.21 for 21%)
    """
    return safe_divide(reopened_sp, delivered_sp, 0.0)


def calculate_carryover_percentage(carryover_sp: float, planned_sp: float) -> float:
    """
    Calculate Carryover % = Carryover / Planned SP
    
    Args:
        carryover_sp: Story points carried over to next sprint
        planned_sp: Story points planned for the sprint
    
    Returns:
        Percentage as decimal (e.g., 0.33 for 33%)
    """
    return safe_divide(carryover_sp, planned_sp, 0.0)


def calculate_completion_ratio(actual_sp: float, planned_sp: float) -> float:
    """
    Calculate Completion Ratio = Actual / Planned
    
    Args:
        actual_sp: Story points actually delivered
        planned_sp: Story points planned for the sprint
    
    Returns:
        Ratio as decimal (e.g., 0.82 for 82%)
    """
    return safe_divide(actual_sp, planned_sp, 0.0)


def calculate_qa_pass_rate(qa_passed: float, total_delivered: float) -> float:
    """
    Calculate QA Pass Rate = QA Passed / Total Delivered
    
    Args:
        qa_passed: Number of items that passed QA
        total_delivered: Total number of items delivered
    
    Returns:
        Percentage as decimal
    """
    return safe_divide(qa_passed, total_delivered, 0.0)


def calculate_sprint_metrics(row: Dict[str, Any]) -> Dict[str, Any]:
    """
    Calculate all KPIs for a single sprint row.
    
    Args:
        row: Dictionary containing sprint data with keys:
            - Planned_SP, Actual_Delivered, V2, LT, Rework
            - Reopened_LT, V2_Carryover, LT_Carryover, QA_Passed
    
    Returns:
        Dictionary with calculated metrics added
    """
    result = row.copy()

    # Extract values with defaults using safe_float
    planned_sp = safe_float(row.get("Planned_SP"))
    actual_delivered = safe_float(row.get("Actual_Delivered"))
    v2 = safe_float(row.get("V2"))
    lt = safe_float(row.get("LT"))
    rework = safe_float(row.get("Rework"))
    reopened_lt = safe_float(row.get("Reopened_LT"))
    v2_carryover = safe_float(row.get("V2_Carryover"))
    lt_carryover = safe_float(row.get("LT_Carryover"))
    qa_passed = safe_float(row.get("QA_Passed"))
    
    # Calculate KPIs
    result["V2_Reopen_Percentage"] = calculate_reopened_percentage(0, v2)  # V2 typically has 0 reopens in sample
    result["LT_Reopen_Percentage"] = calculate_reopened_percentage(reopened_lt, lt)
    result["V2_Carryover_Percentage"] = calculate_carryover_percentage(v2_carryover, planned_sp)
    result["LT_Carryover_Percentage"] = calculate_carryover_percentage(lt_carryover, planned_sp)
    result["Completion_Ratio"] = calculate_completion_ratio(actual_delivered, planned_sp)
    result["QA_Pass_Rate"] = calculate_qa_pass_rate(qa_passed, actual_delivered)
    
    # Additional metrics
    result["Total_Carryover"] = v2_carryover + lt_carryover
    result["Total_Carryover_Percentage"] = calculate_carryover_percentage(
        result["Total_Carryover"], planned_sp
    )
    
    return result


def calculate_metrics_for_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Calculate metrics for all rows in a DataFrame.
    
    Args:
        df: DataFrame containing sprint data
    
    Returns:
        DataFrame with calculated metrics added
    """
    if df.empty:
        return df
    
    # Apply calculations to each row
    calculated_rows = []
    for _, row in df.iterrows():
        calculated_row = calculate_sprint_metrics(row.to_dict())
        calculated_rows.append(calculated_row)
    
    return pd.DataFrame(calculated_rows)
