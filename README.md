# Log Analyzer

A Windows desktop application for analyzing log files and exporting filtered results to Excel.

## Features

- **Filter Log Files**: Search for errors, exceptions, and custom keywords
- **Multiple Filter Tabs**: Create custom filter categories (Errors, Warnings, etc.)
- **Excel Export**: Generate Excel workbooks with multiple sheets:
  - All Logs - Complete log file contents
  - Filter tabs - Filtered results with hyperlinks back to original lines
  - Summary - Statistics and filter results overview
- **Multi-File Support**: Concatenate and analyze multiple log files chronologically
- **Drag & Drop**: Simply drag log files onto the application
- **Large File Handling**: Automatically handles files exceeding Excel's row limit (1M+ rows)
  - Option to export as CSV
  - Option to split across multiple sheets

## Download

Download the latest version from the [Releases](../../releases) page.

1. Go to **Releases** (on the right sidebar)
2. Download `LogAnalyzer.exe` from the latest release
3. Run the application - no installation required!

## Usage

1. **Select a log file**: Click "Browse" or drag & drop a file onto the input area
2. **Configure filters**: Add/remove filter tabs and keywords as needed
3. **Choose output location**: Select where to save the Excel report
4. **Analyze**: Click "Analyze Log File" to generate the report

### Concatenate Mode

Enable "Concatenate Mode" to analyze multiple log files together:
- Select up to 10 files
- Files are merged and sorted chronologically by timestamp
- Date separators help identify when logs are from different days

## Default Filters

The application comes with these default filter categories:
- **Errors**: Error, Exception, Fatal, Critical, Fail, Failed
- **Client Command**: Client Command
- **AET warning**: AET warning
- **Bigfoot General**: Various Bigfoot-specific log entries

You can customize these or add your own filters.

## Building from Source

### Requirements

- Python 3.8+
- Dependencies: `pip install -r requirements.txt`

### Run from Source

```bash
python log_analyzer_gui.py
```

### Build Executable

```bash
# Windows
build_exe.bat

# Or manually with PyInstaller
pip install pyinstaller
pyinstaller --clean --distpath "Log Analyzer" LogAnalyzer.spec
```

The executable will be created in the `Log Analyzer` folder.

## License

This project is provided as-is for personal and commercial use.
