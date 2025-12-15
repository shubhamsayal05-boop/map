# AVLDrive HeatMap Updater - User Guide

## Table of Contents
1. [Getting Started](#getting-started)
2. [Interface Overview](#interface-overview)
3. [Step-by-Step Tutorial](#step-by-step-tutorial)
4. [Understanding AI Features](#understanding-ai-features)
5. [Advanced Usage](#advanced-usage)
6. [FAQ](#faq)

## Getting Started

### First Time Setup

1. **Install Python** (if not already installed)
   - Download from https://www.python.org/downloads/
   - Version 3.7 or higher required
   - During installation, check "Add Python to PATH"

2. **Install Dependencies**
   ```bash
   cd heatmap_updater_app
   pip install -r requirements.txt
   ```

3. **Launch Application**
   ```bash
   python heatmap_updater_gui.py
   ```

### Quick Start

For experienced users:
1. Browse and select Excel file
2. Click "Analyze File"
3. Review AI insights
4. Click "Update HeatMap"
5. Done!

## Interface Overview

### Main Window Components

#### 1. Header Section
- **Title**: Application name and version
- **Subtitle**: Brief description

#### 2. File Selection Section
- **Excel File**: Input file browser
- **Output File**: Where updated file will be saved
- **Browse Buttons**: Open file dialogs

#### 3. Action Buttons
- **🔍 Analyze File**: Run AI analysis on selected file
- **▶ Update HeatMap**: Execute the update operation
- **📊 Export Report**: Save analysis results to file

#### 4. Progress Section
- **Progress Bar**: Visual indicator of operation progress
- **Status Label**: Current operation status

#### 5. Results Tabs
- **📝 Log**: Detailed operation log
- **🤖 AI Insights**: AI-generated recommendations
- **📊 Statistics**: Numerical breakdown of results

## Step-by-Step Tutorial

### Tutorial 1: Your First Update

**Objective**: Update a HeatMap sheet for the first time

**Steps**:

1. **Launch the Application**
   ```bash
   python heatmap_updater_gui.py
   ```
   You should see the main window appear.

2. **Select Your Excel File**
   - Click the "Browse" button next to "Excel File:"
   - Navigate to your AVLDrive HeatMap Tool file
   - Select the `.xlsm` file
   - Click "Open"
   
   **Expected Result**: File path appears in the text field, output path is auto-filled

3. **Run AI Analysis**
   - Click "🔍 Analyze File" button
   - Watch the progress bar
   - Wait for "Analysis complete" message
   
   **Expected Result**: Status turns green, tabs show analysis results

4. **Review AI Insights**
   - Click on "🤖 AI Insights" tab
   - Read the quality score
   - Review recommendations
   - Check for any warnings
   
   **What to Look For**:
   - Quality score above 70 is good
   - Red flags in recommendations need attention
   - Validation errors must be fixed

5. **Check Statistics**
   - Click on "📊 Statistics" tab
   - Review total evaluations
   - Check status distribution
   
   **Understanding the Numbers**:
   - High RED count = many failures
   - Balanced distribution = comprehensive testing

6. **Update the HeatMap**
   - Click "▶ Update HeatMap" button
   - Confirm the action in the dialog
   - Wait for completion
   
   **Expected Result**: Success dialog with update counts

7. **Verify Results**
   - Open the output Excel file
   - Check HeatMap Sheet column R
   - Verify colored dots for sub-operations
   - Verify status text for parent operations

8. **Export Report (Optional)**
   - Click "📊 Export Report"
   - Choose save location
   - Open the report file to review

**Congratulations!** You've completed your first HeatMap update.

### Tutorial 2: Handling Errors

**Objective**: Learn to handle common errors

**Scenario 1: Wrong File Selected**

**Problem**: Selected a non-Excel file or wrong Excel file

**Solution**:
1. Click "Browse" again
2. Select the correct file
3. Look for "AVLDrive_Heatmap_Tool" in the filename

**Scenario 2: Missing Sheets**

**Problem**: Excel file doesn't have required sheets

**Error Message**: "Evaluation Results sheet not found"

**Solution**:
1. Open the Excel file manually
2. Verify these sheets exist:
   - "Evaluation Results"
   - "HeatMap Sheet"
3. If missing, use the correct template file

**Scenario 3: Low Quality Score**

**Problem**: AI reports quality score below 60

**Solution**:
1. Read the AI recommendations carefully
2. Common issues:
   - High failure rate → Review test procedures
   - Limited data → Test more operations
   - Structure issues → Check Excel format
3. Fix issues before proceeding
4. Run analysis again

### Tutorial 3: Batch Processing Multiple Files

**Objective**: Update multiple HeatMap files efficiently

**Steps**:

1. **Prepare Files**
   - Organize all files in one folder
   - Ensure consistent structure
   - Backup originals

2. **Process First File**
   - Follow Tutorial 1
   - Note any issues

3. **Process Remaining Files**
   - For each file:
     - Browse and select
     - Analyze
     - Review AI insights
     - Update
     - Export report
   - Keep notes of any patterns

4. **Compare Results**
   - Review all exported reports
   - Compare quality scores
   - Identify trends

## Understanding AI Features

### Quality Score Calculation

The AI calculates quality score using:

```
Base Score: 100 points

Deductions:
- High failure rate (>50%): -30 points
- Moderate failure rate (30-50%): -15 points
- Limited evaluations (<20): -10 points
- Structure validation errors: varies

Final Score = Base Score - Total Deductions
```

**Example Calculation**:
- Start: 100 points
- 40% failure rate: -15 points
- 25 total evaluations: 0 deduction
- No errors: 0 deduction
- **Final Score: 85** (Good)

### AI Recommendations

The AI generates recommendations based on:

1. **Failure Rate Analysis**
   - >50% failures → "High failure rate detected"
   - 30-50% failures → "Moderate failure rate"
   - <30% failures → "Good test performance"

2. **Data Completeness**
   - <20 evaluations → "Limited evaluation data"
   - <10 evaluations → "Insufficient test coverage"

3. **Structure Validation**
   - Missing columns → "Invalid structure"
   - Wrong format → "Format error"

### Validation Checks

The AI performs these validations:

✓ **Sheet Existence**
- Evaluation Results sheet present
- HeatMap Sheet present

✓ **Column Structure**
- Minimum required columns
- Correct column positions

✓ **Data Integrity**
- OpCodes are numeric
- Status values are valid
- No critical missing data

## Advanced Usage

### Custom Output Paths

Instead of using the auto-generated output path:

1. Click "Browse" next to "Output File:"
2. Choose custom location
3. Enter custom filename
4. Click "Save"

**Tip**: Use descriptive names like `HeatMap_Updated_2024-12-15.xlsm`

### Reading Log Messages

Log messages are color-coded:

- **Black [INFO]**: Normal operations
- **Green [SUCCESS]**: Successful operations
- **Orange [WARNING]**: Non-critical issues
- **Red [ERROR]**: Failures or critical issues

**Example Log Interpretation**:
```
[10:30:15] [INFO] Selected file: file.xlsm
           → Normal file selection

[10:30:20] [SUCCESS] AI analysis completed
           → Analysis successful

[10:30:25] [WARNING] Limited evaluation data
           → Consider more testing

[10:30:30] [ERROR] Failed to load file
           → Critical issue, fix required
```

### Interpreting Statistics

**Status Distribution Example**:
```
RED: 10 (33.3%)    → 10 failures out of 30
YELLOW: 5 (16.7%)  → 5 marginal results
GREEN: 15 (50.0%)  → 15 passes
```

**Analysis**:
- Pass rate: 50% (15/30)
- Failure rate: 33.3%
- Action: Moderate concern, review RED items

### Export Report Formats

**Text Format (.txt)**:
- Human-readable
- Easy to share
- Can be opened in any text editor

**JSON Format (.json)**:
- Machine-readable
- Can be processed programmatically
- Good for automation

## FAQ

### General Questions

**Q: What Excel versions are supported?**
A: Excel 2007 and later (.xlsx, .xlsm files)

**Q: Can I process multiple files at once?**
A: Currently one at a time, but you can process sequentially (see Tutorial 3)

**Q: Does this work on Mac/Linux?**
A: Yes, Python and tkinter work on all platforms

**Q: Will my VBA macros be preserved?**
A: Yes, the application uses `keep_vba=True` to preserve macros

### Technical Questions

**Q: Why does analysis take time?**
A: Large Excel files with many rows require processing time

**Q: Can I run this without GUI?**
A: Yes, you can use the original `update_heatmap_status.py` script

**Q: What if I don't have admin rights?**
A: Use `pip install --user -r requirements.txt`

**Q: Can I customize colors?**
A: Yes, edit the color constants in the script:
```python
RED_COLOR = "FF0000"
GREEN_COLOR = "00FF00"
YELLOW_COLOR = "FFFF00"
```

### Troubleshooting Questions

**Q: Application window is blank**
A: Check if tkinter is installed: `python -m tkinter`

**Q: Progress bar stuck at 100%**
A: This is normal, it resets after completion

**Q: Can't find output file**
A: Check the "Output File:" path in the application

**Q: AI recommendations seem wrong**
A: The AI uses heuristics; apply human judgment to recommendations

### Best Practices

**Q: How often should I analyze before updating?**
A: Always analyze first to catch issues early

**Q: Should I always export reports?**
A: Yes, for documentation and tracking trends

**Q: What quality score is acceptable?**
A: 70+ is good, but review recommendations regardless of score

**Q: How do I know if update was successful?**
A: Check for:
1. Green success message
2. No errors in log
3. Update counts in success dialog
4. Visual inspection of output file

### Data Questions

**Q: What does N/A status mean?**
A: No evaluation data available for that operation

**Q: Why are some operations not updated?**
A: They may not have evaluation data or are not in the HeatMap template

**Q: Can I add custom operations?**
A: Add them to both Evaluation Results and HeatMap sheets first

**Q: What if OpCodes don't match?**
A: Ensure OpCodes are identical in both sheets (8-digit numbers)

## Tips for Success

### Before You Start
- ✓ Backup original files
- ✓ Ensure Excel file is closed
- ✓ Read AI recommendations carefully
- ✓ Keep files in accessible locations

### During Processing
- ✓ Wait for operations to complete
- ✓ Don't close the application mid-process
- ✓ Monitor the log tab for issues
- ✓ Review AI insights before updating

### After Completion
- ✓ Verify output file contents
- ✓ Export report for records
- ✓ Compare before/after if needed
- ✓ Archive old files

## Getting Help

If you encounter issues:

1. **Check the Log Tab**: Look for error messages
2. **Review AI Insights**: May contain helpful hints
3. **Read Error Dialogs**: They provide specific information
4. **Consult README.md**: For technical details
5. **Check File Permissions**: Ensure you can read/write files

## Keyboard Shortcuts

Currently supported:
- **Ctrl+O**: (Future) Open file
- **Ctrl+S**: (Future) Save
- **Ctrl+Q**: (Future) Quit
- **F5**: (Future) Refresh

## Version History

**v1.0.0** (Current)
- Initial release
- GUI interface
- AI-powered analysis
- Automated updates
- Export reports

---

**Need more help?** Check the README.md file for additional technical information.
