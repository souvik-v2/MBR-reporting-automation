"""
Complete MBR Workflow:
1. Extract graph data from LT-SimplifiedMetrics files
2. Extract 'Goals Achieved in May' slides from presentations
3. Create consolidated PowerPoint presentation
4. Generate summary Excel report with graph data
"""
from __future__ import annotations
import pandas as pd
from pathlib import Path
from datetime import datetime

from graph_data_extractor import extract_all_graph_data, create_graph_data_dataframe
from pptx_goals_extractor import create_goals_report

# Try to import automated PowerPoint consolidation
try:
    from pptx_auto_consolidator import consolidate_pptx_com
    AUTO_PPT_AVAILABLE = True
except ImportError:
    AUTO_PPT_AVAILABLE = False
    print("Note: pywin32 not available. PowerPoint consolidation will be manual.")


def run_mbr_workflow(
    graph_folder: str = "sample_data/Graph",
    ppt_folder: str = "sample_data/ppt-data",
    template_path: str = "sample_data/final-data/06-01-2026.pptx",
    output_folder: str = "output",
    month: str = "March",
    use_template: bool = True
):
    """
    Run complete MBR workflow to create consolidated deliverables.

    Args:
        graph_folder: Folder containing LT-SimplifiedMetrics files
continue        template_path: Path to template PPTX (optional if use_template=False)
        output_folder: Output folder for results
        month: Month to search for (e.g., "May", "June")
        use_template: If False, creates new PPT without template

    Creates:
    - consolidated_mbr.pptx: PowerPoint with goals slides inserted into template
    - graph_data_summary.xlsx: Excel file with extracted graph metrics
    - mbr_workflow_log.txt: Processing log
    """
    output_path = Path(output_folder)
    output_path.mkdir(parents=True, exist_ok=True)
    
    log_messages = []
    
    def log(message: str):
        print(message)
        log_messages.append(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {message}")
    
    log("=" * 80)
    log("MBR WORKFLOW STARTED")
    log("=" * 80)
    
    # Step 1: Extract graph data
    log("\n[STEP 1] Extracting graph data from LT-SimplifiedMetrics files...")
    graph_data = extract_all_graph_data(graph_folder)
    graph_df = create_graph_data_dataframe(graph_data)
    
    log(f"  ✓ Extracted data for {len(graph_df)} projects")
    log(f"  Columns: {', '.join(graph_df.columns.tolist())}")
    
    # Save graph data to Excel
    graph_excel_path = output_path / "graph_data_summary.xlsx"
    graph_df.to_excel(graph_excel_path, index=False, sheet_name="Graph Data")
    log(f"  ✓ Saved graph data to: {graph_excel_path}")
    
    # Step 2: Extract goals from PowerPoint presentations
    log(f"\n[STEP 2] Extracting '{month}' goals from presentations...")
    goals_excel_path = output_path / "goals_achieved_report.xlsx"
    ppt_instructions_path = output_path / "ppt_consolidation_instructions.txt"
    consolidated_ppt_path = output_path / "consolidated_mbr.pptx"

    try:
        goals_result = create_goals_report(
            ppt_data_folder=ppt_folder,
            output_excel=str(goals_excel_path),
            output_instructions=str(ppt_instructions_path),
            month=month
        )
        log(f"  ✓ Found {goals_result['goals_found']} goal slides")
        log(f"  ✓ Goals report saved to: {goals_excel_path}")
        log(f"  ✓ Consolidation instructions saved to: {ppt_instructions_path}")
    except Exception as e:
        log(f"  ✗ Error extracting goals: {e}")
        goals_result = {'goals_found': 0}

    # Step 3: Automated PowerPoint Consolidation (if available)
    if AUTO_PPT_AVAILABLE and goals_result.get('goals_found', 0) > 0:
        log(f"\n[STEP 3] Creating consolidated PowerPoint automatically...")

        # Check if template exists
        template_exists = Path(template_path).exists() if use_template else False
        actual_template = template_path if template_exists else None

        if not template_exists and use_template:
            log(f"  ⚠  Template not found: {template_path}")
            log(f"  ℹ  Creating new presentation from goal slides only")

        try:
            consolidate_pptx_com(
                template_path=actual_template,
                ppt_data_folder=ppt_folder,
                output_path=str(consolidated_ppt_path),
                month=month,
                insert_after_slide=5
            )
            log(f"  ✓ Consolidated PowerPoint created: {consolidated_ppt_path}")
        except Exception as e:
            log(f"  ✗ Automated PowerPoint failed: {e}")
            log(f"  ℹ  Use manual instructions in: {ppt_instructions_path}")
    elif not AUTO_PPT_AVAILABLE:
        log(f"\n[STEP 3] Automated PowerPoint not available (install pywin32)")
        log(f"  ℹ  Follow manual instructions in: {ppt_instructions_path}")
    
    # Step 4: Create summary report
    log("\n[STEP 4] Creating summary report...")
    summary_lines = [
        "MBR WORKFLOW SUMMARY",
        "=" * 80,
        f"Execution Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"Month: {month}",
        "",
        "GRAPH DATA EXTRACTED:",
        f"  Total Projects: {len(graph_df)}",
        ""
    ]
    
    for _, row in graph_df.iterrows():
        summary_lines.append(f"  {row['project']}:")
        summary_lines.append(f"    Week: {row.get('week', 'N/A')}")
        summary_lines.append(f"    V2 Carryover: {row.get('v2_carryover', 'N/A')}")
        summary_lines.append(f"    LT Carryover: {row.get('lt_carryover', 'N/A')}")
        summary_lines.append(f"    Rework: {row.get('rework', 'N/A')}")
        summary_lines.append(f"    Reopened SP (LT): {row.get('reopened_sp_lt', 'N/A')}")
        summary_lines.append("")
    
    outputs = [
        f"  1. {graph_excel_path.name} - Graph data summary",
        f"  2. {goals_excel_path.name} - Goals achieved report",
    ]

    if AUTO_PPT_AVAILABLE and consolidated_ppt_path.exists():
        outputs.append(f"  3. {consolidated_ppt_path.name} - AUTOMATED consolidated PowerPoint")
        outputs.append(f"  4. {ppt_instructions_path.name} - Manual instructions (if needed)")
    else:
        outputs.append(f"  3. {ppt_instructions_path.name} - PowerPoint consolidation guide")

    summary_lines.extend([
        "OUTPUTS CREATED:",
        *outputs,
        "",
        "=" * 80
    ])
    
    summary_text = "\n".join(summary_lines)
    summary_path = output_path / "mbr_workflow_summary.txt"
    summary_path.write_text(summary_text)
    log(f"  ✓ Summary report saved to: {summary_path}")
    
    # Save detailed log
    log_path = output_path / "mbr_workflow_log.txt"
    log_path.write_text("\n".join(log_messages), encoding='utf-8')
    
    log("\n" + "=" * 80)
    log("MBR WORKFLOW COMPLETED SUCCESSFULLY")
    log("=" * 80)
    log(f"\nAll outputs saved to: {output_path.absolute()}")
    
    return {
        'graph_data': graph_df,
        'graph_excel': graph_excel_path,
        'goals_excel': goals_excel_path,
        'ppt_instructions': ppt_instructions_path,
        'summary': summary_path,
        'log': log_path
    }


if __name__ == "__main__":
    results = run_mbr_workflow()
    
    print("\n" + "=" * 80)
    print("FILES CREATED:")
    print("=" * 80)
    for key, path in results.items():
        if key != 'graph_data':
            print(f"  {key}: {path}")
