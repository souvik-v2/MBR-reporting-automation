"""
PowerPoint slide extractor for Sprint MBR reporting.

Extracts specific slides by title pattern (e.g., "Goals Achieved in May").
Useful for consolidating monthly achievements from multiple team presentations.
"""
from __future__ import annotations
from pathlib import Path
from typing import List, Dict, Any, Tuple
import re

from pptx import Presentation
from pptx.slide import Slide

from . import config


def extract_slide_title(slide: Slide) -> str:
    """
    Extract the title text from a slide.
    
    Args:
        slide: PowerPoint slide object
    
    Returns:
        Title text or empty string if no title found
    """
    for shape in slide.shapes:
        if shape.has_text_frame:
            if shape.is_placeholder:
                try:
                    # Title placeholders are usually index 0
                    if shape.placeholder_format.idx == 0:
                        return shape.text.strip()
                except Exception:
                    pass
            # Fallback: check if it's at the top of the slide
            if shape.top < 1000000:  # Less than 1 inch from top (in EMUs)
                text = shape.text.strip()
                if text:
                    return text
    return ""


def matches_goal_pattern(title: str, patterns: List[str] = None) -> bool:
    """
    Check if slide title matches any of the goal achievement patterns.
    
    Args:
        title: Slide title text
        patterns: List of patterns to match (uses config default if None)
    
    Returns:
        True if title matches any pattern
    """
    if patterns is None:
        patterns = config.PPTX_GOAL_SLIDE_PATTERNS
    
    title_lower = title.lower()
    for pattern in patterns:
        if pattern.lower() in title_lower:
            return True
    return False


def extract_month_from_title(title: str) -> str:
    """
    Extract month name from slide title (e.g., "Goals Achieved in May" -> "May").
    
    Args:
        title: Slide title text
    
    Returns:
        Month name if found, otherwise empty string
    """
    for month in config.MONTH_NAMES:
        if month.lower() in title.lower():
            return month
    return ""


def extract_slide_content(slide: Slide, slide_idx: int, source_file: str) -> Dict[str, Any]:
    """
    Extract all content from a slide including text, tables, and charts.
    
    Args:
        slide: PowerPoint slide object
        slide_idx: Slide number (1-based)
        source_file: Name of source PPTX file
    
    Returns:
        Dictionary containing slide content
    """
    content = {
        "__source_file": source_file,
        "Slide_Number": slide_idx,
        "Title": extract_slide_title(slide),
        "Text_Content": [],
        "Tables": [],
        "Charts": [],
    }
    
    for shape in slide.shapes:
        # Extract text
        if shape.has_text_frame:
            text = "\n".join(p.text for p in shape.text_frame.paragraphs if p.text.strip())
            if text.strip() and text.strip() != content["Title"]:
                content["Text_Content"].append(text.strip())
        
        # Extract tables
        if shape.has_table:
            table_data = []
            table = shape.table
            for row in table.rows:
                row_data = [cell.text.strip() for cell in row.cells]
                table_data.append(row_data)
            content["Tables"].append(table_data)
        
        # Extract chart titles
        try:
            if hasattr(shape, 'chart'):
                chart = shape.chart
                chart_title = chart.chart_title.text_frame.text if chart.has_title else "Untitled Chart"
                content["Charts"].append(chart_title)
        except Exception:
            pass
    
    return content


def extract_goal_slides(
    pptx_path: Path,
    patterns: List[str] = None,
    log: List[str] = None
) -> List[Dict[str, Any]]:
    """
    Extract all slides that match goal achievement patterns from a PPTX file.
    
    Args:
        pptx_path: Path to PowerPoint file
        patterns: List of title patterns to match (uses config default if None)
        log: Optional log list to append messages
    
    Returns:
        List of dictionaries containing slide content
    """
    if log is None:
        log = []
    
    goal_slides = []
    
    try:
        prs = Presentation(str(pptx_path))
    except Exception as exc:
        log.append(f"[ERROR] Could not read PPTX file '{pptx_path.name}': {exc}")
        return []
    
    for slide_idx, slide in enumerate(prs.slides, start=1):
        title = extract_slide_title(slide)
        
        if matches_goal_pattern(title, patterns):
            month = extract_month_from_title(title)
            content = extract_slide_content(slide, slide_idx, pptx_path.name)
            content["Month"] = month
            goal_slides.append(content)
            log.append(
                f"[INFO] Extracted goal slide '{title}' from '{pptx_path.name}' (slide {slide_idx})"
            )
    
    return goal_slides
