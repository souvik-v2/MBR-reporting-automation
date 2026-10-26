"""
Chart generator for Sprint MBR reporting.

Generates all required charts:
- Sprint/Kanban trend (week-wise)
- Team comparison
- Planned vs Actual delivery
- Quality metrics trends
- Rework trend
- Reopened trend
- V2 quality metrics
"""
from __future__ import annotations
from typing import List, Dict, Any
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import (
    LineChart, BarChart, Reference, Series
)
from openpyxl.chart.label import DataLabelList


def create_sprint_trend_chart(ws, start_row: int, end_row: int, title: str = "Sprint Trend"):
    """
    Create a line chart showing sprint trends over time.
    
    Args:
        ws: Worksheet object
        start_row: Starting row number (1-based)
        end_row: Ending row number (1-based)
        title: Chart title
    
    Returns:
        Chart object
    """
    chart = LineChart()
    chart.title = title
    chart.style = 10
    chart.height = 10
    chart.width = 20
    chart.y_axis.title = "Story Points"
    chart.x_axis.title = "Sprint"
    
    # Data references - adjust column letters based on your sheet structure
    # Assuming: A=Sprint, B=Planned, C=Actual
    cats = Reference(ws, min_col=1, min_row=start_row+1, max_row=end_row)
    
    # Planned SP series
    planned_data = Reference(ws, min_col=2, min_row=start_row, max_row=end_row)
    planned_series = Series(planned_data, cats, title="Planned SP")
    chart.append(planned_series)
    
    # Actual Delivered series
    actual_data = Reference(ws, min_col=3, min_row=start_row, max_row=end_row)
    actual_series = Series(actual_data, cats, title="Actual Delivered")
    chart.append(actual_series)
    
    return chart


def create_team_comparison_chart(ws, start_row: int, end_row: int, title: str = "Team Comparison"):
    """
    Create a bar chart comparing team performance.
    
    Args:
        ws: Worksheet object
        start_row: Starting row number (1-based)
        end_row: Ending row number (1-based)
        title: Chart title
    
    Returns:
        Chart object
    """
    chart = BarChart()
    chart.type = "col"
    chart.title = title
    chart.style = 10
    chart.height = 10
    chart.width = 20
    chart.y_axis.title = "Story Points"
    chart.x_axis.title = "Team"
    
    # Team names as categories
    cats = Reference(ws, min_col=1, min_row=start_row+1, max_row=end_row)
    
    # Completion ratio
    data = Reference(ws, min_col=2, min_row=start_row, max_row=end_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    
    return chart


def create_planned_vs_actual_chart(ws, start_row: int, end_row: int):
    """
    Create a grouped bar chart showing Planned vs Actual delivery.
    
    Args:
        ws: Worksheet object
        start_row: Starting row number (1-based)
        end_row: Ending row number (1-based)
    
    Returns:
        Chart object
    """
    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.title = "Planned vs Actual Delivery"
    chart.style = 11
    chart.height = 10
    chart.width = 20
    chart.y_axis.title = "Story Points"
    chart.x_axis.title = "Sprint"
    
    # Sprint names
    cats = Reference(ws, min_col=1, min_row=start_row+1, max_row=end_row)
    
    # Planned and Actual data
    data = Reference(ws, min_col=2, min_row=start_row, max_col=3, max_row=end_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    
    return chart


def create_quality_metrics_chart(ws, start_row: int, end_row: int):
    """
    Create a line chart showing quality metrics trends.
    
    Args:
        ws: Worksheet object
        start_row: Starting row number (1-based)
        end_row: Ending row number (1-based)
    
    Returns:
        Chart object
    """
    chart = LineChart()
    chart.title = "Quality Metrics Trend"
    chart.style = 12
    chart.height = 10
    chart.width = 20
    chart.y_axis.title = "Percentage"
    chart.x_axis.title = "Sprint"
    
    # Sprint names
    cats = Reference(ws, min_col=1, min_row=start_row+1, max_row=end_row)
    
    # QA Pass Rate, Reopened %, etc.
    data = Reference(ws, min_col=2, min_row=start_row, max_col=4, max_row=end_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    
    return chart


def create_rework_trend_chart(ws, start_row: int, end_row: int):
    """Create a line chart showing rework trend over time."""
    chart = LineChart()
    chart.title = "Rework Trend"
    chart.style = 10
    chart.height = 8
    chart.width = 16
    chart.y_axis.title = "Rework SP"
    chart.x_axis.title = "Sprint"
    
    cats = Reference(ws, min_col=1, min_row=start_row+1, max_row=end_row)
    data = Reference(ws, min_col=2, min_row=start_row, max_row=end_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    
    return chart
