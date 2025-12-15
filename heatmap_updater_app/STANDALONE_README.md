# AVLDrive HeatMap Manager - Standalone Application

## 🎉 Completely Independent of Excel!

A fully standalone desktop application for managing HeatMap evaluations with **NO Excel dependency**. Uses built-in SQLite database for data storage and provides Excel-like data entry experience.

## ✨ Key Features

### 🗄️ Built-in Database
- **SQLite database** for data persistence
- No need for Excel files
- Fast and reliable data storage
- Automatic data backup

### 📝 Excel-Like Data Entry
- Spreadsheet-style data grid
- **Paste from Excel** with one click (Ctrl+V)
- Add, edit, delete evaluations easily
- Familiar interface for Excel users

### 🗺️ HeatMap Generation
- Automatic HeatMap generation from database
- Color-coded status indicators (●)
- Parent operation status aggregation
- Same logic as Excel version

### 🤖 AI-Powered Analysis
- Quality scoring (0-100)
- Failure rate detection
- Smart recommendations
- Data completeness assessment

### 📊 Export Capabilities
- **Export to CSV** for Excel compatibility
- **Export to JSON** for data exchange
- **Export to PDF** (coming soon)
- HeatMap results export

### 🎨 User-Friendly Interface
- Modern GUI with tabs
- No command-line required
- Cross-platform (Windows, Mac, Linux)
- Real-time data updates

## 🚀 Quick Start

### Installation

**No Installation Required!** - Uses Python standard library only.

### Requirements

- Python 3.7 or higher
- **No external dependencies** (everything built-in)

### Launch Application

#### Windows
Double-click `launch_standalone.bat`

#### Mac/Linux
```bash
./launch_standalone.sh
```

Or directly:
```bash
python3 launch_standalone.py
```

## 📖 User Guide

### First Time Setup

1. **Launch the application**
2. **Initialize operations**: Go to `Data → Initialize Operations`
3. **Start entering data**: Use the Data Entry tab

### Data Entry

#### Method 1: Manual Entry
1. Click "➕ Add Evaluation" button
2. Fill in the form:
   - OpCode (e.g., 10101300)
   - Operation name
   - Tested AVL value
   - Driv Status (GREEN/YELLOW/RED)
   - Resp Status (GREEN/YELLOW/RED)
   - Final Status (GREEN/YELLOW/RED)
3. Click "Save"

#### Method 2: Paste from Excel
1. **Copy data from Excel** (select rows and Ctrl+C)
2. Click "📋 Paste from Excel" button
3. Data is automatically imported!

**Format**: Tab-separated values
```
OpCode   Operation   TestedAVL   DrivStatus   RespStatus   FinalStatus
10101300 Creep       7          GREEN        RED          RED
10101100 DASS        6.7        RED          YELLOW       RED
```

#### Method 3: Import CSV
1. Go to `File → Import from CSV`
2. Select your CSV file
3. Data is loaded into database

### Generating HeatMap

1. Go to **HeatMap** tab
2. Click "▶ Generate HeatMap"
3. View results with color coding:
   - **● (RED)**: Failed sub-operations
   - **● (YELLOW)**: Warning sub-operations
   - **● (GREEN)**: Passed sub-operations
   - **NOK**: Failed parent operations
   - **acceptable**: Warning parent operations
   - **OK**: Passed parent operations

### AI Analysis

1. Go to **AI Analysis** tab
2. Click "🔍 Run Analysis"
3. Review:
   - Quality score
   - Status distribution
   - Recommendations
   - Warnings

### Exporting Data

#### Export Evaluations
- `File → Export to CSV` - All evaluation data
- `File → Export to PDF` - Report format (coming soon)

#### Export HeatMap
- Click "📊 Export Results" in HeatMap tab
- Choose CSV or JSON format

## 🗂️ Database Structure

The application uses SQLite with three main tables:

### Operations Table
Stores operation definitions (OpCodes and names)
- OpCode (unique identifier)
- Operation name
- Is parent flag
- Parent OpCode (for sub-operations)

### Evaluations Table
Stores test evaluation results
- OpCode (foreign key)
- Operation name
- Tested AVL value
- Driver status
- Response status
- Final status
- Timestamps

### HeatMap Results Table
Stores generated HeatMap results
- OpCode
- Operation name
- Status (text or dot symbol)
- Status color
- Generation timestamp

## 📊 Data Management

### Clear Data
- `Data → Clear All Evaluations` - Remove all evaluation records
- Database structure remains intact

### Initialize Operations
- `Data → Initialize Operations` - Load default operation structure
- Includes all standard AVL operations

### Backup Database
- Database file: `heatmap_data.db`
- Simply copy this file to backup your data
- Portable - can move to other machines

## 🎯 Operation Hierarchy

Same as Excel version:

### Parent Operations (OpCodes ending with 0000)
- 10100000 - Drive away
- 10120000 - Acceleration
- 10030000 - Tip in
- 10040000 - Tip out
- 10070000 - Deceleration
- 10090000 - Gear shift

### Sub-Operations (Specific OpCodes)
- 10101300 - Creep (under Drive away)
- 10101100 - Standing start (under Drive away)
- 10120300 - Load increase (under Acceleration)
- etc.

## 🔄 Migration from Excel

### Step 1: Export from Excel
1. Open your Excel file
2. Select evaluation data
3. Copy (Ctrl+C)

### Step 2: Import to Standalone App
1. Launch standalone app
2. Go to Data Entry tab
3. Click "📋 Paste from Excel"
4. Done!

Or:
1. Save Excel data as CSV
2. Use `File → Import from CSV`

## 💡 Tips & Tricks

### Efficient Data Entry
- Use Tab key to move between fields
- Copy-paste multiple rows from Excel
- Use dropdowns for status values

### Data Validation
- Run AI Analysis regularly
- Check quality score
- Review recommendations

### Backup Strategy
- Copy `heatmap_data.db` regularly
- Export to CSV for external backup
- Keep multiple versions for history

## 🆚 Comparison with Excel Version

| Feature | Excel Version | Standalone App |
|---------|--------------|----------------|
| Dependencies | openpyxl required | None (built-in) |
| Data Storage | Excel files | SQLite database |
| Data Entry | Excel interface | GUI data grid |
| Performance | Slower with large files | Fast database queries |
| Portability | Need Excel files | Single .db file |
| Multi-user | File conflicts | Better concurrency |
| Backup | Copy .xlsm files | Copy .db file |
| Export | Excel only | CSV, JSON, PDF |

## 🔧 Technical Details

### Architecture
- **GUI**: tkinter (Python standard library)
- **Database**: SQLite3 (Python standard library)
- **Data Processing**: Pure Python
- **No external dependencies**

### Performance
- Fast database operations
- Efficient data queries
- Real-time updates
- Handles thousands of records

### File Locations
- **Database**: `heatmap_data.db` (in app folder)
- **Application**: `standalone_heatmap_app.py`
- **Launcher**: `launch_standalone.py`

## ❓ FAQ

**Q: Do I need to install anything?**
A: No! Just Python 3.7+. No external packages required.

**Q: Can I still use Excel data?**
A: Yes! Copy-paste from Excel or import CSV files.

**Q: Where is my data stored?**
A: In `heatmap_data.db` file in the app folder.

**Q: Can I move the database?**
A: Yes! Just copy the `.db` file to another machine.

**Q: What happened to openpyxl?**
A: Not needed anymore! This app is completely independent.

**Q: Can I export back to Excel?**
A: Yes! Export to CSV and open in Excel.

**Q: Is this the same HeatMap logic?**
A: Yes! Exactly the same evaluation and status calculation.

**Q: Can multiple users share the database?**
A: SQLite supports multiple readers, one writer at a time.

## 🐛 Troubleshooting

### Application won't start
- Check Python version: `python3 --version`
- Ensure Python 3.7+
- Try running directly: `python3 launch_standalone.py`

### Can't paste from Excel
- Ensure data is tab-separated
- Copy full rows including OpCode
- Check clipboard contains valid data

### Database errors
- Check `heatmap_data.db` exists
- Verify file permissions
- Try "Initialize Operations" if empty

### Missing operations
- Run `Data → Initialize Operations`
- This loads default operation structure

## 📞 Support

For issues:
1. Check this README
2. Verify Python version
3. Try reinitializing operations
4. Check database file exists

## 🎓 Tutorial

### Complete Workflow Example

1. **Start Application**
   ```bash
   ./launch_standalone.sh
   ```

2. **Initialize Data**
   - Go to `Data → Initialize Operations`

3. **Add Evaluations**
   Option A - Manual:
   - Click "➕ Add Evaluation"
   - Fill: OpCode=10101300, Operation=Creep, Status=RED
   - Click Save
   
   Option B - From Excel:
   - Copy rows from Excel
   - Click "📋 Paste from Excel"

4. **Generate HeatMap**
   - Go to HeatMap tab
   - Click "▶ Generate HeatMap"
   - View color-coded results

5. **Run Analysis**
   - Go to AI Analysis tab
   - Click "🔍 Run Analysis"
   - Review recommendations

6. **Export Results**
   - `File → Export to CSV`
   - Or HeatMap tab → "📊 Export Results"

## 🌟 Advantages

✅ **No Excel dependency**
✅ **Fast and efficient**
✅ **Built-in database**
✅ **Easy data entry**
✅ **AI-powered analysis**
✅ **Multiple export formats**
✅ **Cross-platform**
✅ **No installation required**
✅ **Same HeatMap logic**
✅ **Better data management**

## 🚀 Getting Started Now

1. Launch: `./launch_standalone.sh` (or `.bat` on Windows)
2. Initialize: `Data → Initialize Operations`
3. Add data: Paste from Excel or add manually
4. Generate: Click "▶ Generate HeatMap"
5. Analyze: Run AI Analysis
6. Export: Save results

**You're ready to go!** 🎉

---

**Version**: 2.0 (Standalone)
**No Dependencies**: Uses Python standard library only
**Database**: SQLite3 built-in
**Status**: Production Ready
