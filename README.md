# Job-Application-Tracker - Excel VBA Macro

A comprehensive Excel VBA macro system for tracking job applications with an intuitive button-based interface and color-coded status management.

🚀 Features
> Complete Job Tracking: Track company, position, employment type, work location, and shift preferences
> Status Management: Color-coded status tracking from application to offer/rejection
> One-Click Interface: Easy-to-use buttons for all functions
> Smart Validation: Dropdown menus ensure consistent data entry
> Follow-up Reminders: Set and track follow-up dates
> Status Reports: Generate summary reports of your application progress
> Data Safety: Confirmation dialogs prevent accidental data loss

📊 Track Everything
> Core Information
  - Application ID (auto-generated)
  - Company Name & Job Title
  - Application Date (auto-filled)
  - Contact Person & Email
  - Salary Range & Notes

> Job Categorization
  - Employment Type: Full-time, Part-time, Contract, Temporary, Internship
  - Work Location: On-site, Hybrid, Work from Home, Remote
  - Work Shift: Day Shift, Night Shift, Graveyard Shift, Flexible, Rotating

> Application Status
  - Applied (Blue) - Initial application submitted
  - Phone Screen (Orange, Bold) - Phone screening scheduled/completed
  - Interview Scheduled (Red-Orange, Bold) - Interview confirmed
  - Interviewed (Purple, Bold) - Interview completed
  - Follow-up (Brown) - Waiting for response or following up
  - Offer (Green, Bold) - Job offer received ✅
  - Rejected (Red, Bold) - Application declined ❌
  - Withdrawn (Gray) - You withdrew from consideration

🔧 Installation
Step 1: Download the Code
1. Download Job Application Tracker.bas from this repository
2. Or copy the code from the file

Step 2: Set Up Excel
1. Open Excel and create a new workbook
2. Save it as .xlsm (Excel Macro-Enabled Workbook)
3. Press Alt + F11 to open the VBA Editor

Step 3: Import the Macro
1. In VBA Editor, right-click on your workbook name in the left panel
2. Select Insert > Module
3. Copy and paste the entire code into the new module
4. Press Ctrl + S to save

Step 4: Enable Macros
1. Close VBA Editor and return to Excel
2. If you see a security warning, click "Enable Content"
3. Or go to File > Options > Trust Center > Trust Center Settings > Macro Settings
4. Select "Enable all macros" (for testing) or "Disable all macros with notification"

🎯 Quick Start
> Initial Setup
1. Press Alt + F8 to open the macro dialog
2. Select SetupJobTracker and click "Run"
3. Select CreateControlButtons and click "Run"
4. Your tracker is ready to use!

> Using the Interface
The control panel provides these buttons:
  - Add New Application - Add a new job application with guided prompts
  - Update Status - Change the status of an existing application
  - Set Follow-up - Add follow-up reminders for applications
  - Status Report - Generate a summary report of all applications
  - Setup Tracker - Reset/recreate the tracker structure
  - Refresh Colors - Reapply color coding to all status entries
  - Clear All Data - Safely remove all application data (keeps structure)

📋 Usage Guide
Adding Applications
1. Click "Add New Application"
2. Fill in the prompted information:
  - Company Name (required)
  - Job Title (required)
  - Employment Type, Work Location, Work Shift
  - Status, Contact Info, Salary, Notes (optional)
3. Application is automatically added with color coding

Managing Applications
> Update Status: Use Application ID to change status
> Set Reminders: Add follow-up dates to track when to contact employers
> Generate Reports: View summary statistics of your job search progress

Data Management
> Export: Copy data to other applications
> Backup: Save copies of your .xlsm file regularly
> Clear: Use "Clear All Data" for fresh starts while keeping the structure

🎨 Color Coding System
The status column uses text colors to help you quickly identify application stages:
> 🔵 Applied - Initial submissions
> 🟠 Phone Screen - Active screening process
> 🔴 Interview Scheduled - Confirmed interviews
> 🟣 Interviewed - Waiting for results
> 🟤 Follow-up - Following up on applications
> 🟢 Offer - Successful applications
> 🔴 Rejected - Closed applications
> ⚫ Withdrawn - Self-removed applications

🛡️ Safety Features
> Confirmation Dialogs: Prevents accidental data deletion
> Input Validation: Dropdown menus ensure consistent data
> Auto-Backup: Excel's auto-save protects your work
> Structure Preservation: Clear function keeps headers and formatting

📁 File Structure
your-excel-file.xlsm
├── Job Applications (main tracking sheet)
├── Status Report (generated reports)
└── VBA Module (macro code)

🔧 Troubleshooting
> Macro Security Issues
  - Error: "Cannot run the macro"
  - Solution: Enable macros in Excel security settings or add your file location to trusted locations

> Button Not Working
  - Error: Buttons don't respond to clicks
  - Solution: Ensure macros are enabled and the VBA module is properly imported

> Missing Dropdown Options
  - Error: Status or other dropdowns are empty
  - Solution: Run SetupJobTracker again to restore validation lists

> Data Not Saving
  - Error: Information disappears after closing Excel
  - Solution: Ensure your file is saved as .xlsm format (macro-enabled)

🙏 Acknowledgments
> Designed for job seekers who want to stay organized
> Built with Excel VBA for maximum compatibility
> Inspired by the need for simple, effective job tracking
