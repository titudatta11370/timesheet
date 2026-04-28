# Solar Edge Timesheet Reconciliation Tool

## What it does
Compares Solar Edge roster emails against timesheet Excel files and generates a color-coded reconciliation report.

## Setup (one time)
1. Install Python from python.org (check "Add Python to PATH")
2. Open Command Prompt and run: pip install pandas openpyxl
3. Create these folders inside the project: emails\ and reports\

## Every week
1. Save the Solar Edge email as .txt in the emails\ folder
2. Make sure SharePoint timesheets are synced to your PC
3. Double-click run_reconcile.bat
4. Drag the .txt email file into the window and press Enter
5. Excel report opens automatically