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


def get_status_priority(status):
    """
    Return priority for status values. Lower number = worse status.
    RED is worst, GREEN is best, N/A is neutral.
    """
    if not status or status == 'N/A' or status is None:
        return 3  # Neutral
    
    status_str = str(status).upper().strip()
    
    if status_str == 'RED':
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
    valid_statuses = [s for s in statuses if s and str(s).upper().strip() != 'N/A']
    
    if not valid_statuses:
        return None  # No valid status found
    
    # Find the worst status (minimum priority)
    worst_status = min(valid_statuses, key=get_status_priority)
    return worst_status


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
    for row_num in range(2, ws_eval.max_row + 1):
        op_code = ws_eval.cell(row=row_num, column=1).value
        operation = ws_eval.cell(row=row_num, column=2).value
        final_status = ws_eval.cell(row=row_num, column=12).value  # Column L
        
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
    STATUS_COLUMN = 18  # Column R
    
    print("\nUpdating HeatMap Sheet Status column...")
    updates_made = 0
    no_match_count = 0
    
    for row_num in range(4, ws_heatmap.max_row + 1):
        op_code = ws_heatmap.cell(row=row_num, column=1).value
        operation = ws_heatmap.cell(row=row_num, column=2).value
        
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
                
                # Update the cell
                current_value = ws_heatmap.cell(row=row_num, column=STATUS_COLUMN).value
                ws_heatmap.cell(row=row_num, column=STATUS_COLUMN).value = final_status
                
                updates_made += 1
                print(f"  Row {row_num}: {op_code} | {operation}")
                print(f"    Sub-operations: {len(eval_by_opcode[op_code_int])}")
                print(f"    Statuses: {statuses}")
                print(f"    Final Status: {current_value} => {final_status}")
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
