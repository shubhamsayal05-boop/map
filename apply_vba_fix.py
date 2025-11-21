#!/usr/bin/env python3
"""
Utility script to extract and display VBA code that needs to be fixed.
This script does NOT automatically modify the Excel file - it provides
the information needed to manually apply the fix.
"""

from oletools.olevba import VBA_Parser
import sys
import os

def main():
    input_file = 'AVLDrive_Heatmap_Tool version3.1.xlsm'
    
    if not os.path.exists(input_file):
        print(f"ERROR: Could not find {input_file}")
        print("Please run this script from the directory containing the Excel file.")
        sys.exit(1)
    
    print("="*80)
    print("VBA CODE FIX UTILITY")
    print("="*80)
    print(f"\nAnalyzing: {input_file}")
    
    try:
        vba = VBA_Parser(input_file)
        
        evaluation_code = None
        for (filename, stream_path, vba_filename, vba_code) in vba.extract_all_macros():
            if 'Evaluation' in vba_filename and 'InferParentMode' in vba_code:
                evaluation_code = vba_code
                break
        
        if not evaluation_code:
            print("\nERROR: Could not find Evaluation module with InferParentMode function!")
            vba.close()
            sys.exit(1)
        
        # Find the function
        start_idx = evaluation_code.find('Private Function InferParentMode')
        if start_idx == -1:
            print("\nERROR: Could not find InferParentMode function!")
            vba.close()
            sys.exit(1)
        
        end_idx = evaluation_code.find('End Function', start_idx) + len('End Function')
        
        current_function = evaluation_code[start_idx:end_idx]
        
        # Check if already fixed
        if 'Left$(code, 4) = Left$(k, 4)' in current_function:
            print("\n✓ The fix has ALREADY been applied!")
            print("  The InferParentMode function is using the correct 4-digit matching logic.")
        else:
            print("\n✗ The fix has NOT been applied yet.")
            print("  The InferParentMode function needs to be updated.")
            print("\nCURRENT FUNCTION:")
            print("="*80)
            print(current_function)
            print("="*80)
            
            print("\n" + "="*80)
            print("INSTRUCTIONS:")
            print("="*80)
            print("1. Open the Excel file in Excel or LibreOffice")
            print("2. Press Alt+F11 to open VBA Editor")
            print("3. Find the 'Evaluation' module")
            print("4. Replace the InferParentMode function with:")
            print()
            print("="*80)
            print('''Private Function InferParentMode(code As String, modes As Object) As String
    If modes.Exists(code) Then
        InferParentMode = code
        Exit Function
    End If

    Dim k As Variant
    ' Match based on first 4 digits since all operation modes follow pattern "10XX0000"
    ' where XX identifies the mode (e.g., 1010 = Drive away)
    For Each k In modes.Keys
        If Len(code) >= 4 And Len(k) >= 4 Then
            If Left$(code, 4) = Left$(k, 4) Then
                InferParentMode = k
                Exit Function
            End If
        End If
    Next k

    InferParentMode = ""
End Function''')
            print("="*80)
            print("\n5. Save the file (Ctrl+S)")
            print("6. Close VBA Editor")
            print("\nFor detailed instructions, see VBA_CODE_FIX.md")
        
        vba.close()
        
    except Exception as e:
        print(f"\nERROR: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == '__main__':
    main()
