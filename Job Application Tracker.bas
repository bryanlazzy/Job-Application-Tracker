Attribute VB_Name = "Module1"
Sub SetupJobTracker()
    
    Dim ws As Worksheet
    Dim headers As Variant
    Dim i As Integer
    
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Job Applications")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Job Applications"
    End If
    
    
    ws.Cells.Clear
    
    
    headers = Array("Application ID", "Company Name", "Job Title", "Employment Type", _
                   "Work Location", "Work Shift", "Application Date", "Status", _
                   "Contact Person", "Contact Email", "Salary Range", "Notes", _
                   "Interview Date", "Follow-up Date", "Response Date")
    
    
    For i = 1 To UBound(headers) + 1
        ws.Cells(1, i).Value = headers(i - 1)
        ws.Cells(1, i).Font.Bold = True
        ws.Cells(1, i).Interior.Color = RGB(70, 130, 180)
        ws.Cells(1, i).Font.Color = RGB(255, 255, 255)
    Next i
    
   
    ws.Columns.AutoFit
    
    
    With ws.Range("D:D").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="Full-time,Part-time,Contract,Temporary,Internship"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    
    
    With ws.Range("E:E").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="On-site,Hybrid,Work from Home,Remote"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    
    
    With ws.Range("F:F").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="Day Shift,Night Shift,Graveyard Shift,Flexible,Rotating"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    
    
    With ws.Range("H:H").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="Applied,Phone Screen,Interview Scheduled,Interviewed,Follow-up,Offer,Rejected,Withdrawn"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    
    
    Call ApplyStatusColors
    
    
    ws.Columns("G:G").NumberFormat = "mm/dd/yyyy"
    ws.Columns("M:M").NumberFormat = "mm/dd/yyyy"
    ws.Columns("N:N").NumberFormat = "mm/dd/yyyy"
    ws.Columns("O:O").NumberFormat = "mm/dd/yyyy"
    
    
    ws.Range("A1:O1").Borders.Weight = xlThin
    
    MsgBox "Job Application Tracker setup complete! You can now start adding your applications."
End Sub

Sub AddJobApplication()
    
    Dim ws As Worksheet
    Dim nextRow As Long
    Dim company As String, jobTitle As String, employmentType As String
    Dim workLocation As String, workShift As String, status As String
    Dim contactPerson As String, contactEmail As String, salaryRange As String, notes As String
    
    Set ws = ThisWorkbook.Worksheets("Job Applications")
    nextRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row + 1
    
    
    company = InputBox("Enter Company Name:", "New Job Application")
    If company = "" Then Exit Sub
    
    jobTitle = InputBox("Enter Job Title:", "New Job Application")
    If jobTitle = "" Then Exit Sub
    
    employmentType = InputBox("Enter Employment Type:" & vbNewLine & vbNewLine & _
                             "Options:" & vbNewLine & _
                             "• Full-time" & vbNewLine & _
                             "• Part-time" & vbNewLine & _
                             "• Contract" & vbNewLine & _
                             "• Temporary" & vbNewLine & _
                             "• Internship", "New Job Application - Employment Type", "Full-time")
    If employmentType = "" Then employmentType = "Full-time"
    
    workLocation = InputBox("Enter Work Location:" & vbNewLine & vbNewLine & _
                           "Options:" & vbNewLine & _
                           "• On-site" & vbNewLine & _
                           "• Hybrid" & vbNewLine & _
                           "• Work from Home" & vbNewLine & _
                           "• Remote", "New Job Application - Work Location", "On-site")
    If workLocation = "" Then workLocation = "On-site"
    
    workShift = InputBox("Enter Work Shift:" & vbNewLine & vbNewLine & _
                        "Options:" & vbNewLine & _
                        "• Day Shift" & vbNewLine & _
                        "• Night Shift" & vbNewLine & _
                        "• Graveyard Shift" & vbNewLine & _
                        "• Flexible" & vbNewLine & _
                        "• Rotating", "New Job Application - Work Shift", "Day Shift")
    If workShift = "" Then workShift = "Day Shift"
    
    status = InputBox("Enter Status:" & vbNewLine & vbNewLine & _
                     "Options:" & vbNewLine & _
                     "• Applied" & vbNewLine & _
                     "• Phone Screen" & vbNewLine & _
                     "• Interview Scheduled" & vbNewLine & _
                     "• Interviewed" & vbNewLine & _
                     "• Follow-up" & vbNewLine & _
                     "• Offer" & vbNewLine & _
                     "• Rejected" & vbNewLine & _
                     "• Withdrawn", "New Job Application - Status", "Applied")
    If status = "" Then status = "Applied"
    
    contactPerson = InputBox("Enter Contact Person (optional):", "New Job Application")
    contactEmail = InputBox("Enter Contact Email (optional):", "New Job Application")
    salaryRange = InputBox("Enter Salary Range (optional):", "New Job Application")
    notes = InputBox("Enter Notes (optional):", "New Job Application")
    
    
    ws.Cells(nextRow, 1).Value = nextRow - 1
    ws.Cells(nextRow, 2).Value = company
    ws.Cells(nextRow, 3).Value = jobTitle
    ws.Cells(nextRow, 4).Value = employmentType
    ws.Cells(nextRow, 5).Value = workLocation
    ws.Cells(nextRow, 6).Value = workShift
    ws.Cells(nextRow, 7).Value = Date
    ws.Cells(nextRow, 8).Value = status
    ws.Cells(nextRow, 9).Value = contactPerson
    ws.Cells(nextRow, 10).Value = contactEmail
    ws.Cells(nextRow, 11).Value = salaryRange
    ws.Cells(nextRow, 12).Value = notes
    
    
    Call ColorCodeStatus(ws.Cells(nextRow, 8))
    
    
    ws.Columns.AutoFit
    
    MsgBox "Job application for " & company & " - " & jobTitle & " has been added!"
End Sub

Sub UpdateApplicationStatus()
    
    Dim ws As Worksheet
    Dim appID As String
    Dim newStatus As String
    Dim searchRange As Range
    Dim foundCell As Range
    Dim targetRow As Long
    
    Set ws = ThisWorkbook.Worksheets("Job Applications")
    
    appID = InputBox("Enter Application ID to update:", "Update Status")
    If appID = "" Then Exit Sub
    
    
    Set searchRange = ws.Range("A:A")
    Set foundCell = searchRange.Find(What:=appID, LookIn:=xlValues, LookAt:=xlWhole)
    
    If foundCell Is Nothing Then
        MsgBox "Application ID " & appID & " not found!"
        Exit Sub
    End If
    
    targetRow = foundCell.Row
    
    newStatus = InputBox("Enter New Status:" & vbNewLine & vbNewLine & _
                        "Current Status: " & ws.Cells(targetRow, 8).Value & vbNewLine & vbNewLine & _
                        "Available Options:" & vbNewLine & _
                        "• Applied" & vbNewLine & _
                        "• Phone Screen" & vbNewLine & _
                        "• Interview Scheduled" & vbNewLine & _
                        "• Interviewed" & vbNewLine & _
                        "• Follow-up" & vbNewLine & _
                        "• Offer" & vbNewLine & _
                        "• Rejected" & vbNewLine & _
                        "• Withdrawn", "Update Application Status")
    If newStatus = "" Then Exit Sub
    
    
    ws.Cells(targetRow, 8).Value = newStatus
    
    
    Call ColorCodeStatus(ws.Cells(targetRow, 8))
    
    
    If newStatus = "Interview Scheduled" Then
        Dim interviewDate As String
        interviewDate = InputBox("Enter interview date (MM/DD/YYYY):", "Interview Date")
        If IsDate(interviewDate) Then
            ws.Cells(targetRow, 13).Value = CDate(interviewDate)
        End If
    End If
    
    
    If newStatus = "Offer" Or newStatus = "Rejected" Then
        ws.Cells(targetRow, 15).Value = Date
    End If
    
    MsgBox "Status updated to: " & newStatus
End Sub

Sub GenerateStatusReport()
    
    Dim ws As Worksheet
    Dim reportWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim statusCount As Object
    Dim status As String
    Dim key As Variant
    Dim reportRow As Long
    
    Set ws = ThisWorkbook.Worksheets("Job Applications")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    If lastRow <= 1 Then
        MsgBox "No applications found to report on."
        Exit Sub
    End If
    
    
    On Error Resume Next
    Set reportWs = ThisWorkbook.Worksheets("Status Report")
    On Error GoTo 0
    
    If reportWs Is Nothing Then
        Set reportWs = ThisWorkbook.Worksheets.Add
        reportWs.Name = "Status Report"
    Else
        reportWs.Cells.Clear
    End If
    
    
    Set statusCount = CreateObject("Scripting.Dictionary")
    
    For i = 2 To lastRow
        status = ws.Cells(i, 8).Value
        If status <> "" Then
            If statusCount.exists(status) Then
                statusCount(status) = statusCount(status) + 1
            Else
                statusCount(status) = 1
            End If
        End If
    Next i
    
    
    reportWs.Cells(1, 1).Value = "Job Application Status Report"
    reportWs.Cells(1, 1).Font.Bold = True
    reportWs.Cells(1, 1).Font.Size = 14
    
    reportWs.Cells(2, 1).Value = "Generated on: " & Date
    
    reportWs.Cells(4, 1).Value = "Status"
    reportWs.Cells(4, 2).Value = "Count"
    reportWs.Range("A4:B4").Font.Bold = True
    
    reportRow = 5
    For Each key In statusCount.keys
        reportWs.Cells(reportRow, 1).Value = key
        reportWs.Cells(reportRow, 2).Value = statusCount(key)
        reportRow = reportRow + 1
    Next key
    
    reportWs.Cells(reportRow + 1, 1).Value = "Total Applications:"
    reportWs.Cells(reportRow + 1, 2).Value = lastRow - 1
    reportWs.Range("A" & (reportRow + 1) & ":B" & (reportRow + 1)).Font.Bold = True
    
    reportWs.Columns.AutoFit
    reportWs.Activate
    
    MsgBox "Status report generated successfully!"
End Sub

Sub AddFollowUpReminder()
    
    Dim ws As Worksheet
    Dim appID As String
    Dim followUpDate As String
    Dim searchRange As Range
    Dim foundCell As Range
    Dim targetRow As Long
    
    Set ws = ThisWorkbook.Worksheets("Job Applications")
    
    appID = InputBox("Enter Application ID to set follow-up reminder:", "Follow-up Reminder")
    If appID = "" Then Exit Sub
    
    
    Set searchRange = ws.Range("A:A")
    Set foundCell = searchRange.Find(What:=appID, LookIn:=xlValues, LookAt:=xlWhole)
    
    If foundCell Is Nothing Then
        MsgBox "Application ID " & appID & " not found!"
        Exit Sub
    End If
    
    targetRow = foundCell.Row
    
    followUpDate = InputBox("Enter follow-up date (MM/DD/YYYY):", "Follow-up Date")
    If IsDate(followUpDate) Then
        ws.Cells(targetRow, 14).Value = CDate(followUpDate)
        MsgBox "Follow-up reminder set for " & followUpDate
    Else
        MsgBox "Invalid date format entered."
    End If
End Sub

Sub CreateControlButtons()
    
    Dim ws As Worksheet
    Dim btn As Object
    Dim buttonTop As Double
    Dim buttonLeft As Double
    Dim buttonWidth As Double
    Dim buttonHeight As Double
    
    Set ws = ThisWorkbook.Worksheets("Job Applications")
    
    
    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo 0
    
    
    buttonWidth = 150
    buttonHeight = 25
    buttonLeft = ws.Columns("N").Left
    buttonTop = ws.Rows(3).Top
    
    
    Set btn = ws.Buttons.Add(buttonLeft, buttonTop, buttonWidth, buttonHeight)
    btn.Text = "Add New Application"
    btn.OnAction = "AddJobApplication"
    btn.Font.Size = 10
    btn.Font.Bold = True
    
    
    buttonTop = buttonTop + buttonHeight + 5
    Set btn = ws.Buttons.Add(buttonLeft, buttonTop, buttonWidth, buttonHeight)
    btn.Text = "Update Status"
    btn.OnAction = "UpdateApplicationStatus"
    btn.Font.Size = 10
    btn.Font.Bold = True
    
    
    buttonTop = buttonTop + buttonHeight + 5
    Set btn = ws.Buttons.Add(buttonLeft, buttonTop, buttonWidth, buttonHeight)
    btn.Text = "Set Follow-up"
    btn.OnAction = "AddFollowUpReminder"
    btn.Font.Size = 10
    btn.Font.Bold = True
    
    
    buttonTop = buttonTop + buttonHeight + 5
    Set btn = ws.Buttons.Add(buttonLeft, buttonTop, buttonWidth, buttonHeight)
    btn.Text = "Status Report"
    btn.OnAction = "GenerateStatusReport"
    btn.Font.Size = 10
    btn.Font.Bold = True
    
    
    buttonTop = buttonTop + buttonHeight + 5
    Set btn = ws.Buttons.Add(buttonLeft, buttonTop, buttonWidth, buttonHeight)
    btn.Text = "Setup Tracker"
    btn.OnAction = "SetupJobTracker"
    btn.Font.Size = 10
    btn.Font.Bold = True
    
    
    buttonTop = buttonTop + buttonHeight + 5
    Set btn = ws.Buttons.Add(buttonLeft, buttonTop, buttonWidth, buttonHeight)
    btn.Text = "Refresh Colors"
    btn.OnAction = "RefreshAllColors"
    btn.Font.Size = 10
    btn.Font.Bold = True
    
    
    buttonTop = buttonTop + buttonHeight + 5
    Set btn = ws.Buttons.Add(buttonLeft, buttonTop, buttonWidth, buttonHeight)
    btn.Text = "Clear All Data"
    btn.OnAction = "ClearAllData"
    btn.Font.Size = 10
    btn.Font.Bold = True
    
    
    With ws.Range("N1")
        .Value = "JOB TRACKER CONTROLS"
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    
    ws.Range("N1:P1").Merge
    ws.Range("N1:P1").HorizontalAlignment = xlCenter
    
    
    ws.Columns("N:P").AutoFit
    
    MsgBox "Control buttons have been added to your worksheet! You can now easily access all job tracking functions."
End Sub

Sub ShowPendingFollowUps()
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim followUpDate As Date
    Dim today As Date
    Dim pendingList As String
    Dim count As Integer
    
    Set ws = ThisWorkbook.Worksheets("Job Applications")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    today = Date
    count = 0
    pendingList = "Applications needing follow-up:" & vbNewLine & vbNewLine
    
    For i = 2 To lastRow
        If IsDate(ws.Cells(i, 14).Value) Then
            followUpDate = ws.Cells(i, 14).Value
            If followUpDate <= today And ws.Cells(i, 8).Value <> "Offer" And ws.Cells(i, 8).Value <> "Rejected" Then
                count = count + 1
                pendingList = pendingList & "• " & ws.Cells(i, 2).Value & " - " & ws.Cells(i, 3).Value & _
                             " (Due: " & Format(followUpDate, "mm/dd/yyyy") & ")" & vbNewLine
            End If
        End If
    Next i
    
    If count = 0 Then
        MsgBox "No pending follow-ups found!", vbInformation, "Follow-up Check"
    Else
        MsgBox pendingList, vbInformation, "Pending Follow-ups (" & count & ")"
    End If
End Sub

Sub ColorCodeStatus(cell As Range)
    
    Dim statusValue As String
    statusValue = cell.Value
    
    
    cell.Interior.Color = xlNone
    cell.Font.Color = RGB(0, 0, 0)
    cell.Font.Bold = False
    
    Select Case statusValue
        Case "Applied"
            cell.Font.Color = RGB(0, 100, 200)        ' Blue
            
        Case "Phone Screen"
            cell.Font.Color = RGB(255, 140, 0)        ' Orange
            cell.Font.Bold = True
            
        Case "Interview Scheduled"
            cell.Font.Color = RGB(255, 69, 0)         ' Red-Orange
            cell.Font.Bold = True
            
        Case "Interviewed"
            cell.Font.Color = RGB(148, 0, 211)        ' Purple
            cell.Font.Bold = True
            
        Case "Follow-up"
            cell.Font.Color = RGB(139, 69, 19)        ' Brown
            
        Case "Offer"
            cell.Font.Color = RGB(0, 150, 0)          ' Green
            cell.Font.Bold = True
            
        Case "Rejected"
            cell.Font.Color = RGB(220, 20, 60)        ' Red
            cell.Font.Bold = True
            
        Case "Withdrawn"
            cell.Font.Color = RGB(105, 105, 105)      ' Gray
            
    End Select
End Sub

Sub ApplyStatusColors()
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Worksheets("Job Applications")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    
    For i = 2 To lastRow
        If ws.Cells(i, 8).Value <> "" Then
            Call ColorCodeStatus(ws.Cells(i, 8))
        End If
    Next i
End Sub

Sub RefreshAllColors()
    
    Call ApplyStatusColors
    MsgBox "Status colors have been refreshed!", vbInformation, "Color Refresh"
End Sub

Sub ClearAllData()
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim response As Integer
    
    Set ws = ThisWorkbook.Worksheets("Job Applications")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    
    If lastRow <= 1 Then
        MsgBox "No data to clear. The tracker is already empty.", vbInformation, "Clear Data"
        Exit Sub
    End If
    
    
    response = MsgBox("Are you sure you want to clear ALL job application data?" & vbNewLine & vbNewLine & _
                     "This will delete " & (lastRow - 1) & " application(s) permanently." & vbNewLine & _
                     "Headers and formatting will be preserved.", _
                     vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Clear All Data")
    
    If response = vbYes Then
        
        If lastRow > 1 Then
            ws.Range("A2:O" & lastRow).Clear
        End If
        
        
        ws.Range("A:A").NumberFormat = "General"
        ws.Range("G:G").NumberFormat = "mm/dd/yyyy"
        ws.Range("M:M").NumberFormat = "mm/dd/yyyy"
        ws.Range("N:N").NumberFormat = "mm/dd/yyyy"
        ws.Range("O:O").NumberFormat = "mm/dd/yyyy"
        
        
        ws.Columns.AutoFit
        
        MsgBox "All job application data has been cleared successfully!" & vbNewLine & _
               "You can now start fresh with new applications.", vbInformation, "Data Cleared"
    Else
        MsgBox "Clear operation cancelled. Your data is safe.", vbInformation, "Operation Cancelled"
    End If
End Sub

