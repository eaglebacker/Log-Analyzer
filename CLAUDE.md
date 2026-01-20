# Claude Project Instructions - Log Analyzer

## Project Overview
Log Analyzer is a Windows desktop application for analyzing log files and exporting filtered results to Excel.

## Important Workflows

### After Making Code Changes
Always perform these steps after modifying `log_analyzer_gui.py` or other source files:

1. **Rebuild the executable:**
   ```bash
   cd "C:\Claude Projects\Log reader" && python -m PyInstaller --clean --distpath "Log Analyzer" LogAnalyzer.spec
   ```

2. **Commit and push to GitHub:**
   ```bash
   cd "C:\Claude Projects\Log reader" && git add . && git commit -m "Description of changes" && git push
   ```

3. **Remind the user** to create a new Release on GitHub if this is a version they want to distribute.

## Project Structure
- **Source code:** `log_analyzer_gui.py` (main GUI), `log_analyzer.py` (CLI)
- **Executable output:** `Log Analyzer/LogAnalyzer.exe`
- **Build config:** `LogAnalyzer.spec`, `build_exe.bat`
- **Icon:** `sasquatch.ico`

## GitHub Repository
- **URL:** https://github.com/eaglebacker/Log-Analyzer
- **Releases page:** https://github.com/eaglebacker/Log-Analyzer/releases

## Build Notes
- Use `--distpath "Log Analyzer"` to output to the correct folder (not default `dist`)
- The executable is ~19MB and includes tkinterdnd2 for drag-and-drop support

## Key Features to Remember
- Drag and drop works on the entire input section (not just the text field)
- Filter tabs have hyperlinks that jump to the original line in "All Logs"
- Large files (>1M rows) trigger a dialog offering CSV export or split sheets
