# SETUP-INSTRUCTIONS
This guide will walk you through setting up the Job Application Tracker Excel macro step-by-step.

📋 Prerequisites
- Microsoft Excel (2010 or later)
- Windows or Mac computer
- Basic knowledge of Excel

📥 Download and Setup
Step 1: Download the Files
1. Download JobApplicationTracker.bas from this repository
2. Save it to a location you can easily find (like your Desktop)

Step 2: Create Your Excel Workbook
1. Open Microsoft Excel
2. Create a new blank workbook
3. Important: Save it as .xlsm format:
  - Click File > Save As
  - Choose .xlsm (Excel Macro-Enabled Workbook) from the dropdown
  - Name it something like MyJobTracker.xlsm

🔧 Import the Macro Code
Method 1: Copy-Paste (Recommended)
1. Press Alt + F11 to open the VBA Editor
2. In the left panel, right-click on your workbook name
3. Select Insert > Module
4. Open the JobApplicationTracker.bas file in a text editor
5. Copy all the code (Ctrl+A, then Ctrl+C)
6. Paste it into the new module window in Excel (Ctrl+V)
7. Press Ctrl + S to save

Method 2: Import File
1. Press Alt + F11 to open the VBA Editor
2. Right-click on your workbook name in the left panel
3. Select Import File...
4. Browse and select Job Application Tracker.bas
5. Press Ctrl + S to save

🔒 Enable Macros
Option 1: Enable for This File
1. Close the VBA Editor (Alt + F11)
2. Look for a yellow security warning bar at the top
3. Click Enable Content or Enable Macros

Option 2: Change Excel Security Settings
1. Go to File > Options
2. Click Trust Center in the left panel
3. Click Trust Center Settings...
4. Click Macro Settings
5. Select Disable all macros with notification
6. Click OK and restart Excel

Option 3: Add to Trusted Location (Most Secure)
1. Go to File > Options > Trust Center > Trust Center Settings
2. Click Trusted Locations
3. Click Add new location...
4. Browse to the folder containing your Excel file
5. Check Subfolders of this location are also trusted
6. Click OK

🚀 Initialize the Tracker
Step 1: Run Setup
1. Press Alt + F8 to open the macro dialog
2. Select SetupJobTracker from the list
3. Click Run
4. You should see a message: "Job Application Tracker setup complete!"

Step 2: Add Control Buttons
1. Press Alt + F8 again
2. Select CreateControlButtons from the list
3. Click Run
4. You should see buttons appear on the right side of your worksheet

✅ Test Your Setup
Quick Test
1. Click the Add New Application button
2. Enter test data when prompted:
  - Company: "Test Company"
  - Job Title: "Test Position"
  - Accept defaults for other fields
3. Verify the data appears in your spreadsheet
4. Try clicking Status Report to generate a test report

Clean Up Test Data
1. Click Clear All Data button
2. Confirm when prompted
3. Your tracker is now ready for real use!

🛠️ Troubleshooting
"Cannot run the macro" Error
Problem: Excel security is blocking macros
Solutions:
1. Enable macros using the yellow security bar
2. Check Trust Center settings (see Enable Macros section above)
3. Move your file to a Trusted Location

Buttons Don't Appear
Problem: CreateControlButtons didn't run properly
Solutions:
1. Make sure you ran SetupJobTracker first
2. Try running CreateControlButtons again
3. Check if macros are enabled

Dropdown Lists Are Empty
Problem: Data validation wasn't set up properly
Solutions:
1. Run SetupJobTracker again
2. Manually check if you're on the "Job Applications" worksheet

Data Doesn't Save
Problem: File wasn't saved as macro-enabled format
Solutions:
1. Save as .xlsm format (not .xlsx)
2. Enable macros when reopening the file

VBA Editor Issues
Problem: Can't access VBA Editor
Solutions:
1. Enable Developer tab: File > Options > Customize Ribbon > Check Developer
2. Or use Alt + F11 keyboard shortcut

📱 Platform-Specific Notes
Windows
- Use Alt + F11 for VBA Editor
- Most compatible with all features

Mac
- Use Option + F11 for VBA Editor
- Some button formatting may differ slightly
- All core functionality works the same

🎯 Next Steps
Once setup is complete:
1. Read the main README.md for usage instructions
2. Start adding your real job applications
3. Customize the categories if needed
4. Set up a backup routine for your data
