# AVLDrive Heatmap Tool

This repository contains the AVLDrive Heatmap Tool Excel workbook and automation tools to update the HeatMap sheet with evaluation results.

## Files

- `AVLDrive_Heatmap_Tool version3.2.xlsm` - The main Excel workbook containing evaluation data and heatmap visualizations
- `update_heatmap_status.py` - Python script to transfer evaluation results to the heatmap status column
- `UpdateHeatMapStatus.vba` - VBA code for Excel button automation
- `VBA_INSTALLATION_GUIDE.md` - Step-by-step guide to install VBA code in Excel

## Overview

The Excel workbook contains two main sheets:

1. **Evaluation Results** - Contains detailed evaluation data for various operations and sub-operations with final status ratings (RED, YELLOW, GREEN, N/A)
2. **HeatMap Sheet** - Contains a visual heatmap of operations with a Status column that needs to be populated from the evaluation results

## Usage Options

You can update the HeatMap Status column using either:
- **Option A**: VBA Macro (directly in Excel) - **Recommended for end users**
- **Option B**: Python Script (command line) - For automation and batch processing

### Option A: VBA Macro (In Excel)

**Best for**: Users who want to click a button directly in Excel

1. Follow the instructions in `VBA_INSTALLATION_GUIDE.md` to install the VBA code
2. Click the "Update HeatMap Status" button in the HeatMap Sheet
3. The Status column will be automatically updated

See `VBA_INSTALLATION_GUIDE.md` for complete installation steps.

### Option B: Python Script

### Prerequisites

```bash
pip install openpyxl
```

### Running the Script

To update the HeatMap sheet with evaluation results:

```bash
python update_heatmap_status.py "AVLDrive_Heatmap_Tool version3.2.xlsm"
```

This will:
1. Read all evaluations from the "Evaluation Results" sheet
2. Group them by Operation Code (Op Code)
3. For each operation in the "HeatMap Sheet", find matching evaluations by Op Code
4. Determine the final status using the worst status among all sub-operations (RED > YELLOW > GREEN)
5. Update the Status column (column R/18) in the HeatMap Sheet
6. Save the updated workbook (overwrites the original file)

### Creating a Backup

To save the output to a new file instead of overwriting:

```bash
python update_heatmap_status.py "AVLDrive_Heatmap_Tool version3.2.xlsm" "output_file.xlsm"
```

## How It Works

### Matching Logic

The script matches operations between sheets using the **Op Code** field:

- Each operation in the HeatMap sheet has an Op Code (e.g., 10101300, 10101100)
- The Evaluation Results sheet contains multiple sub-operations for each Op Code
- The script finds all sub-operations that match a given Op Code

### Status Determination

When multiple sub-operations exist for an Op Code, the script determines the final status by selecting the **worst status**:

1. **RED** - Worst (highest priority)
2. **YELLOW** - Medium
3. **GREEN** - Good (lowest priority)
4. **N/A** or **None** - Ignored (not considered in the worst status calculation)

For example:
- If sub-operations have statuses: RED, YELLOW, N/A → Final status: **RED**
- If sub-operations have statuses: YELLOW, GREEN, N/A → Final status: **YELLOW**
- If sub-operations have statuses: GREEN, N/A, N/A → Final status: **GREEN**
- If all sub-operations have N/A or None → Final status: **None**

### Results

After running the script on version 3.2:
- **28 operations** were updated with status values
- **18 operations** had no matching evaluations (remain as None/empty)

## Example Output

```
Loading workbook: AVLDrive_Heatmap_Tool version3.2.xlsm

Reading Evaluation Results...
Found 36 unique op codes with evaluations

Updating HeatMap Sheet Status column...
  Row 6: 10101300 | Creep
    Sub-operations: 5
    Statuses: ['RED', 'N/A', 'YELLOW', 'N/A', None]
    Final Status: None => RED

=== Summary ===
Total updates made: 28
No matches found: 18

Done!
```

## Technical Details

- The script preserves Excel macros (VBA code) when saving
- Status column is located at column R (column 18)
- The script uses openpyxl library to read and write Excel files
- Data starts at row 4 in the HeatMap Sheet (rows 1-3 are headers)
