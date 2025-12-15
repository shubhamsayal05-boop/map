# Quick Start Guide

## 🚀 Get Started in 3 Minutes

### Windows Users

1. **Double-click** `launch.bat`
2. Click "Browse" to select your Excel file
3. Click "🔍 Analyze File"
4. Click "▶ Update HeatMap"
5. Done! ✓

### Mac/Linux Users

1. **Double-click** `launch.sh` (or run `./launch.sh` in terminal)
2. Click "Browse" to select your Excel file
3. Click "🔍 Analyze File"
4. Click "▶ Update HeatMap"
5. Done! ✓

## 📋 What You Need

- ✓ Python 3.7+ (will be checked automatically)
- ✓ Your AVLDrive HeatMap Tool Excel file (.xlsm)
- ✓ 5 minutes of time

## 🎯 First Time Setup

If launcher says "Python not found":

**Windows**:
1. Download Python from https://www.python.org/downloads/
2. Run installer
3. **Important**: Check "Add Python to PATH"
4. Click "Install Now"
5. Restart computer
6. Try launcher again

**Mac**:
```bash
# Install Homebrew (if not installed)
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

# Install Python
brew install python3
```

**Linux (Ubuntu/Debian)**:
```bash
sudo apt update
sudo apt install python3 python3-pip python3-tk
```

## 💡 Quick Tips

1. **Always Analyze First**: The AI checks for issues before updating
2. **Review Quality Score**: Above 70 is good
3. **Read Recommendations**: AI gives helpful suggestions
4. **Backup Original**: Keep a copy of your original file
5. **Export Reports**: Great for documentation

## 🆘 Common Issues

**"tkinter not found"**
- Windows: Reinstall Python with "tcl/tk and IDLE" checked
- Mac: `brew install python-tk`
- Linux: `sudo apt install python3-tk`

**"openpyxl not found"**
- Run: `pip install openpyxl`
- Or: Launcher will try to install automatically

**"File not found"**
- Make sure Excel file path is correct
- Try browsing to file instead of typing path
- Check file isn't open in Excel

## 📖 Need More Help?

- **Basic Usage**: See `USER_GUIDE.md`
- **Technical Details**: See `README.md`
- **All Features**: Check the app's built-in help

## 🎓 Video Tutorial

(Future: Link to video tutorial will be added here)

## ✅ Success Checklist

After running the app, verify:
- [ ] Green success message appeared
- [ ] No errors in log tab
- [ ] Output file created
- [ ] HeatMap Sheet column R has dots and text
- [ ] Colors look correct (red/yellow/green)

**Congratulations!** You're now using AI-powered HeatMap updates! 🎉

---

**Questions?** Check the User Guide or README for detailed information.
