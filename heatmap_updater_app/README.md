# AVLDrive HeatMap Updater - GUI Application with AI

A robust, user-friendly desktop application for automatically updating HeatMap sheets with evaluation results. Features an intelligent AI engine for validation, analysis, and recommendations.

![Application Screenshot](screenshot.png)

## ✨ Features

### 🖥️ User-Friendly GUI
- **Intuitive Interface**: Clean, modern design with easy-to-use controls
- **File Browser**: Simple file selection with drag-and-drop support
- **Real-time Progress**: Visual progress bar and status updates
- **Multi-tab Results**: Organized view of logs, AI insights, and statistics

### 🤖 AI-Powered Intelligence
- **Smart Analysis**: Automatic evaluation data quality assessment
- **Quality Scoring**: 0-100 quality score based on test results
- **Intelligent Recommendations**: Context-aware suggestions for improvement
- **Failure Rate Detection**: Automatic identification of high failure rates
- **Structure Validation**: Validates HeatMap sheet structure before processing

### 🎨 Advanced Features
- **Colored Status Indicators**: Red/Yellow/Green dots for sub-operations
- **Automated Status Text**: "NOK"/"acceptable"/"OK" for parent operations
- **Hierarchy Detection**: Automatically identifies parent-child relationships
- **Export Reports**: Save analysis results and logs to file
- **Error Handling**: Comprehensive error detection and user-friendly messages
- **Background Processing**: Non-blocking UI during long operations

### 📊 Visual Feedback
- Real-time progress tracking
- Color-coded log messages (INFO, WARNING, ERROR, SUCCESS)
- Statistics dashboard with percentage breakdowns
- AI insights panel with actionable recommendations

## 🚀 Installation

### Prerequisites
- Python 3.7 or higher
- tkinter (usually included with Python)

### Setup

1. **Navigate to the app folder**:
   ```bash
   cd heatmap_updater_app
   ```

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

## 📖 Usage

### Starting the Application

Run the application with:
```bash
python heatmap_updater_gui.py
```

Or on Windows, double-click the `.py` file.

### Step-by-Step Guide

1. **Select Excel File**
   - Click "Browse" next to "Excel File"
   - Select your AVLDrive HeatMap Tool Excel file (`.xlsm` or `.xlsx`)
   - The output path will be automatically set

2. **Analyze File (AI-Powered)**
   - Click "🔍 Analyze File" button
   - Wait for AI analysis to complete
   - Review the AI insights and statistics tabs
   - Check quality score and recommendations

3. **Update HeatMap**
   - After successful analysis, click "▶ Update HeatMap"
   - Confirm the action
   - Wait for the update to complete
   - Review the results in the log tab

4. **Export Report (Optional)**
   - Click "📊 Export Report" to save analysis results
   - Choose save location
   - Report includes logs, AI insights, and statistics

### Understanding the Interface

#### Tabs
- **📝 Log**: Timestamped log of all operations with color-coded messages
- **🤖 AI Insights**: AI-generated recommendations and warnings
- **📊 Statistics**: Detailed breakdown of evaluation results

#### Status Indicators
- **Green**: Operation successful
- **Blue**: Operation in progress
- **Red**: Error occurred
- **Orange**: Warning issued

## 🧠 AI Features Explained

### Quality Scoring
The AI analyzer assigns a quality score (0-100) based on:
- Failure rate (RED status count)
- Data completeness
- Structure validation

**Score Interpretation**:
- **90-100**: Excellent - Low failure rate, complete data
- **75-89**: Good - Moderate results with minor issues
- **60-74**: Fair - Significant failures or missing data
- **Below 60**: Poor - High failure rate or incomplete data

### Intelligent Recommendations

The AI provides context-aware recommendations such as:
- **High Failure Rate (>50%)**: Suggests reviewing test procedures
- **Moderate Failure Rate (30-50%)**: Highlights operations needing attention
- **Low Test Coverage**: Recommends testing more operations
- **Structure Issues**: Identifies missing columns or invalid format

### Automatic Validation

Before processing, the AI validates:
- ✓ Excel file structure
- ✓ Required sheets exist
- ✓ Column layout is correct
- ✓ Data integrity

## 🎯 How It Works

### Sub-Operations (Colored Dots)
- **● Red**: RED status (failure) - Test failed
- **● Yellow**: YELLOW status (warning) - Marginal performance
- **● Green**: GREEN status (pass) - Test passed

### Parent Operations (Status Text)
- **NOK** (red): At least one sub-operation failed
- **acceptable** (yellow): All passing but some marginal
- **OK** (green): All sub-operations passed

### Hierarchy Detection
- **Parent Operations**: OpCodes ending with ≥4 zeros (e.g., 10100000)
- **Sub-Operations**: Specific OpCodes (e.g., 10101300)
- Parent status is automatically calculated from child statuses

## 📁 Application Structure

```
heatmap_updater_app/
├── heatmap_updater_gui.py  # Main application
├── requirements.txt         # Python dependencies
├── README.md               # This file
└── USER_GUIDE.md          # Detailed user guide
```

## 🔧 Technical Details

### Architecture
- **GUI Framework**: Tkinter (Python standard library)
- **Excel Processing**: openpyxl library
- **Threading**: Background processing for non-blocking UI
- **AI Engine**: Custom AIAnalyzer class with intelligent rules

### Data Flow
1. User selects Excel file
2. AI analyzer validates structure and evaluates data
3. Quality score and recommendations generated
4. User confirms update
5. Sub-operations updated with colored dots
6. Parent operations calculated and updated
7. Results saved to output file

### Error Handling
- File not found → User-friendly error message
- Invalid structure → Detailed validation report
- Missing sheets → Clear indication of problem
- Save errors → Alternative save options offered

## 🆘 Troubleshooting

### "File not found" Error
- Ensure the Excel file path is correct
- Check file permissions
- Try copying file to a different location

### "Invalid structure" Error
- Verify the Excel file has "Evaluation Results" sheet
- Check that "HeatMap Sheet" exists
- Ensure columns are in expected positions

### Application Won't Start
- Verify Python 3.7+ is installed: `python --version`
- Install dependencies: `pip install -r requirements.txt`
- Check tkinter is available: `python -m tkinter`

### Slow Performance
- Large Excel files may take time to process
- Check the progress bar for status
- Background threads prevent UI freezing

## 💡 Tips & Best Practices

1. **Backup Original File**: Always keep a backup before updating
2. **Review AI Analysis**: Check insights before proceeding with update
3. **Export Reports**: Save reports for documentation
4. **Test with Sample**: Try with a copy first
5. **Check Quality Score**: Address recommendations if score is low

## 🔮 Future Enhancements

Potential features for future versions:
- Machine learning for anomaly detection
- Predictive analysis of test trends
- Automated report generation with charts
- Batch processing of multiple files
- Cloud integration for team collaboration

## 📜 License

This application is provided as-is for use with AVLDrive HeatMap Tool.

## 🤝 Support

For issues or questions:
1. Check the troubleshooting section
2. Review the user guide
3. Check application logs for detailed errors

## 🎓 Credits

Developed with:
- Python 3.x
- Tkinter GUI framework
- openpyxl library for Excel processing
- Custom AI analysis engine

---

**Version**: 1.0.0  
**Last Updated**: December 2024
