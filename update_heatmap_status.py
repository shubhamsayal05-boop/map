#!/usr/bin/env python3
"""
Script to update HeatMap Sheet with evaluation results from Evaluation Results sheet.

This script:
1. Reads all evaluations from the "Evaluation Results" sheet
2. Groups them by Op Code
3. For each operation in "HeatMap Sheet", finds matching evaluations by Op Code
4. Determines the final status (worst status among sub-operations)
5. Updates the Status column (column R/18) in the HeatMap Sheet
"""

import openpyxl
from collections import defaultdict
import sys
import os

# Column constants for Evaluation Results sheet
EVAL_OP_CODE_COLUMN = 1
EVAL_OPERATION_COLUMN = 2
EVAL_FINAL_STATUS_COLUMN = 12

# Column constants for HeatMap Sheet
HEATMAP_OP_CODE_COLUMN = 1
HEATMAP_OPERATION_COLUMN = 2
HEATMAP_STATUS_COLUMN = 18

# Row constants
EVAL_DATA_START_ROW = 2
HEATMAP_DATA_START_ROW = 4


def get_status_priority(status):
    """
    Return priority for status values. Lower number = worse status.
    RED is worst, GREEN is best, N/A is neutral.
    """
    if not status:
        return 3  # Neutral (None or empty)
    
    status_str = str(status).upper().strip()
    
    if status_str == 'N/A':
        return 3  # Neutral
    elif status_str == 'RED':
        return 0  # Worst
    elif status_str == 'YELLOW':
        return 1  # Medium
    elif status_str == 'GREEN':
        return 2  # Good
    else:
        return 3  # Unknown/N/A


def determine_final_status(statuses):
    """
    Determine the final status from a list of statuses.
    Returns the worst status (RED > YELLOW > GREEN).
    If all are N/A or None, returns None.
    """
    # Filter out None and N/A values
    valid_statuses = []
    for s in statuses:
        if s is not None and str(s).upper().strip() != 'N/A':
            valid_statuses.append(s)
    
    if not valid_statuses:
        return None  # No valid status found
    
    # Find the worst status (minimum priority)
    worst_status = min(valid_statuses, key=get_status_priority)
    return worst_status


def format_status_with_dot(status):
    """
    Convert status to formatted text with colored dot.
    RED -> "● NOK", YELLOW -> "● Acceptable", GREEN -> "● OK"
    """
    if not status:
        return None
    
    status_str = str(status).upper().strip()
    
    if status_str == 'RED':
        return '● NOK'
    elif status_str == 'YELLOW':
        return '● Acceptable'
    elif status_str == 'GREEN':
        return '● OK'
    else:
        return None


def apply_status_color(cell, status):
    """
    Apply color formatting to status cell.
    """
    if not status:
        return
    
    status_str = str(status).upper().strip()
    
    # Import Font and Color from openpyxl.styles
    from openpyxl.styles import Font, Color
    
    if status_str == 'RED':
        cell.font = Font(color='FF0000')  # Red
    elif status_str == 'YELLOW':
        cell.font = Font(color='FFC000')  # Orange/Yellow
    elif status_str == 'GREEN':
        cell.font = Font(color='00B050')  # Green


def update_heatmap_with_evaluations(input_file, output_file=None):
    """
    Update the HeatMap Sheet with evaluation results.
    
    Args:
        input_file: Path to the input Excel file
        output_file: Path to the output Excel file (if None, overwrites input)
    """
    print(f"Loading workbook: {input_file}")
    
    # Load the workbook (keep_vba=True to preserve macros)
    wb = openpyxl.load_workbook(input_file, keep_vba=True)
    
    # Check if required sheets exist
    if 'Evaluation Results' not in wb.sheetnames:
        print("ERROR: 'Evaluation Results' sheet not found!")
        return False
    
    if 'HeatMap Sheet' not in wb.sheetnames:
        print("ERROR: 'HeatMap Sheet' sheet not found!")
        return False
    
    # Read Evaluation Results
    ws_eval = wb['Evaluation Results']
    eval_by_opcode = defaultdict(list)
    
    print("\nReading Evaluation Results...")
    for row_num in range(EVAL_DATA_START_ROW, ws_eval.max_row + 1):
        op_code = ws_eval.cell(row=row_num, column=EVAL_OP_CODE_COLUMN).value
        operation = ws_eval.cell(row=row_num, column=EVAL_OPERATION_COLUMN).value
        final_status = ws_eval.cell(row=row_num, column=EVAL_FINAL_STATUS_COLUMN).value
        
        # Only process rows with valid op codes
        if op_code and isinstance(op_code, (int, float)):
            op_code_int = int(op_code)
            eval_by_opcode[op_code_int].append({
                'row': row_num,
                'operation': operation,
                'status': final_status
            })
    
    print(f"Found {len(eval_by_opcode)} unique op codes with evaluations")
    
    # Update HeatMap Sheet
    ws_heatmap = wb['HeatMap Sheet']
    
    print("\nUpdating HeatMap Sheet Status column...")
    updates_made = 0
    no_match_count = 0
    
    for row_num in range(HEATMAP_DATA_START_ROW, ws_heatmap.max_row + 1):
        op_code = ws_heatmap.cell(row=row_num, column=HEATMAP_OP_CODE_COLUMN).value
        operation = ws_heatmap.cell(row=row_num, column=HEATMAP_OPERATION_COLUMN).value
        
        if not op_code or not operation:
            continue
        
        # Convert op_code to int for matching
        if isinstance(op_code, (int, float)):
            op_code_int = int(op_code)
            
            # Find matching evaluations
            if op_code_int in eval_by_opcode:
                # Get all statuses for this op code
                statuses = [item['status'] for item in eval_by_opcode[op_code_int]]
                
                # Determine final status
                final_status = determine_final_status(statuses)
                
                # Format status with colored dot
                formatted_status = format_status_with_dot(final_status)
                
                # Update the cell
                status_cell = ws_heatmap.cell(row=row_num, column=HEATMAP_STATUS_COLUMN)
                current_value = status_cell.value
                status_cell.value = formatted_status
                
                # Apply color formatting
                apply_status_color(status_cell, final_status)
                
                updates_made += 1
                print(f"  Row {row_num}: {op_code} | {operation}")
                print(f"    Sub-operations: {len(eval_by_opcode[op_code_int])}")
                print(f"    Statuses: {statuses}")
                print(f"    Final Status: {current_value} => {formatted_status}")
            else:
                no_match_count += 1
                print(f"  Row {row_num}: {op_code} | {operation} - No matching evaluation")
    
    print(f"\n=== Summary ===")
    print(f"Total updates made: {updates_made}")
    print(f"No matches found: {no_match_count}")
    
    # Save the workbook
    if output_file is None:
        output_file = input_file
    
    print(f"\nSaving workbook to: {output_file}")
    wb.save(output_file)
    wb.close()
    
    print("Done!")
    return True


def main():
    """Main entry point."""
    if len(sys.argv) < 2:
        print("Usage: python update_heatmap_status.py <input_file> [output_file]")
        print("\nIf output_file is not specified, the input file will be updated in place.")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    if not os.path.exists(input_file):
        print(f"ERROR: Input file not found: {input_file}")
        sys.exit(1)
    
    success = update_heatmap_with_evaluations(input_file, output_file)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
