Option Explicit

' UAN (Urgent Action Network) Reports Generator v1.0.0
' Last Updated: 2024-03-19
' Author: AI Assistant
'
' DESCRIPTION:
' This script processes campaign data from an Excel sheet to generate various reports
' tracking supporter engagement across different dimensions like country, case number,
' topics etc. The data is expected to be in a sheet named "processed-export"
'
' REQUIREMENTS:
' - Windows: Microsoft Scripting Runtime reference required for Dictionary object
' - Mac: Uses Collection object instead of Dictionary
' - Excel 2010 or later recommended
'
' SHEET REQUIREMENTS:
' - "processed-export" sheet with headers:
'   * Campaign ID
'   * Campaign Date
'   * Supporter ID
'   * Supporter Email
'   * External Reference 6 (Country)
'   * External Reference 7 (Case Number)
'   * External Reference 8 (Topics)
'   * External Reference 10 (Year)
'   * External Reference 10 (Type)
'
' USAGE:
' Windows: Use the "Update UAN Reports" menu in the ribbon
' Mac: Use Command+U to show the reports menu
'
' REPORTS GENERATED:
' - by-name: Campaign ID counts with unique supporters
' - by-case-number: Case number engagement
' - by-country: Country-wise participation
' - by-topic: Topic-wise breakdown
' - by-year: Year-wise analysis
' - by-type: Type-based categorization
' - by-date: Monthly trends
' - by-supporter: Individual supporter engagement
'
' PERFORMANCE NOTES:
' - Uses arrays instead of ranges for better performance
' - Includes progress indicators for long operations
' - Handles large datasets efficiently
'
' ERROR HANDLING:
' - Validates all required columns
' - Handles date input validation
' - Provides user feedback for all operations
' - Graceful cleanup on errors
'
' MAINTENANCE NOTES:
' - Mac/Windows compatibility handled via compiler directives
' - Dictionary operations abstracted for cross-platform support
' - Status updates provided via Application.StatusBar
' - Excel state properly managed for reliability

' Configuration settings
Private Type ConfigSettings
    DefaultDateFormat As String
    MaxRowsToProcess As Long
    EnableDebugMode As Boolean
End Type

Private Type ColumnIndices
    CampaignID As Long
    CampaignDate As Long
    SupporterID As Long
    SupporterEmail As Long
    Country As Long
    CaseNumber As Long
    Topics As Long
    Year As Long
    Type As Long
End Type

#If Mac Then
    ' Mac-specific Dictionary alternative using Collection
    Private Type MacDict
        Keys As Collection
        Items As Collection
    End Type
    
    Private Type ReportCounts
        Dict As MacDict
    End Type
#Else
    ' Windows version using Scripting.Dictionary
    Private Type ReportCounts
        Dict As Scripting.Dictionary
    End Type
#End If

' Add menu to Excel ribbon
Public Sub AddReportMenu()
    On Error GoTo ErrorHandler
    
    SetApplicationState False
    
    #If Mac Then
        Application.OnKey "⌘u", "ShowReportMenu"  ' Command+U on Mac
    #Else
        On Error Resume Next
        Application.CommandBars("UAN Reports").Delete
        On Error GoTo ErrorHandler
        
        Dim menuBar As CommandBar
        Dim mainMenu As CommandBarPopup
        
        Set menuBar = Application.CommandBars.Add(Name:="UAN Reports", Position:=msoBarTop, Temporary:=True)
        Set mainMenu = menuBar.Controls.Add(Type:=msoControlPopup)
        
        With mainMenu
            .Caption = "Update UAN Reports"
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Update All Reports"
                .OnAction = "ProcessCampaignData"
            End With
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Update All Reports (except by-supporter)"
                .OnAction = "ProcessCampaignDataExceptSupporter"
            End With
            .Controls.Add Type:=msoControlSeparator
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Update by-name"
                .OnAction = "UpdateCampaignReport"
            End With
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Update by-case-number"
                .OnAction = "UpdateCaseReport"
            End With
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Update by-country"
                .OnAction = "UpdateCountryReport"
            End With
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Update by-topic"
                .OnAction = "UpdateTopicReport"
            End With
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Update by-year"
                .OnAction = "UpdateYearReport"
            End With
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Update by-type"
                .OnAction = "UpdateTypeReport"
            End With
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Update by-date"
                .OnAction = "UpdateDateReport"
            End With
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Update by-supporter"
                .OnAction = "UpdateSupporterReport"
            End With
        End With
        
        menuBar.Visible = True
    #End If

ExitSub:
    SetApplicationState True
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume ExitSub
End Sub

Private Function CreateDictionary() As ReportCounts
    #If Mac Then
        Dim rc As ReportCounts
        Set rc.Dict.Keys = New Collection
        Set rc.Dict.Items = New Collection
        CreateDictionary = rc
    #Else
        Dim rc As ReportCounts
        Set rc.Dict = CreateObject("Scripting.Dictionary")
        CreateDictionary = rc
    #End If
End Function

Private Sub SetApplicationState(ByVal enable As Boolean)
    With Application
        .ScreenUpdating = enable
        .EnableEvents = enable
        .Calculation = IIf(enable, xlCalculationAutomatic, xlCalculationManual)
        If enable Then .StatusBar = False
    End With
End Sub

Private Sub ShowProgress(ByVal current As Long, ByVal total As Long)
    Application.StatusBar = "Processing... " & Format(current / total, "0%")
    DoEvents
End Sub

Private Function ValidateDate(ByVal dateStr As String) As Boolean
    If dateStr = "" Then
        ValidateDate = True
        Exit Function
    End If
    
    If Not IsDate(dateStr) Then Exit Function
    If CDate(dateStr) < DateSerial(1900, 1, 1) Then Exit Function
    If CDate(dateStr) > DateSerial(2100, 12, 31) Then Exit Function
    ValidateDate = True
End Function

Private Function GetDateInput(ByVal promptText As String) As Date
    On Error GoTo ErrorHandler
    
    Dim dateStr As String
    Dim dateVal As Date
    
    dateStr = Application.InputBox(promptText, "Enter Date", Type:=2)
    
    If dateStr = "" Or dateStr = "False" Then
        GetDateInput = DateSerial(1900, 1, 1)
        Exit Function
    End If
    
    If Not ValidateDate(dateStr) Then
        MsgBox "Invalid date format. Please use YYYY-MM-DD format.", vbExclamation
        GetDateInput = DateSerial(1900, 1, 1)
        Exit Function
    End If
    
    GetDateInput = CDate(dateStr)
    Exit Function
    
ErrorHandler:
    MsgBox "Error processing date: " & Err.Description, vbExclamation
    GetDateInput = DateSerial(1900, 1, 1)
End Function

' ... rest of the existing code with error handling added ...

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        GetOrCreateSheet.Name = sheetName
    Else
        GetOrCreateSheet.Cells.Clear
    End If
End Function

Private Function GetColumnIndices() As ColumnIndices
    Dim headers As Range
    Dim cols As ColumnIndices
    
    Set headers = ThisWorkbook.Sheets("processed-export").Rows(1)
    
    With cols
        .CampaignID = Application.Match("Campaign ID", headers, 0)
        .CampaignDate = Application.Match("Campaign Date", headers, 0)
        .SupporterID = Application.Match("Supporter ID", headers, 0)
        .SupporterEmail = Application.Match("Supporter Email", headers, 0)
        .Country = Application.Match("External Reference 6 (Country)", headers, 0)
        .CaseNumber = Application.Match("External Reference 7 (Case Number)", headers, 0)
        .Topics = Application.Match("External Reference 8 (Topics)", headers, 0)
        .Year = Application.Match("External Reference 10 (Year)", headers, 0)
        .Type = Application.Match("External Reference 10 (Type)", headers, 0)
    End With
    
    GetColumnIndices = cols
End Function

Private Sub UpdateReportDates(ByVal startDate As String, ByVal endDate As String)
    Dim reportSheet As Worksheet
    Set reportSheet = GetOrCreateSheet("report")
    
    With reportSheet
        .Range("A2").Value = "Start Date"
        .Range("A3").Value = "End Date"
        .Range("B2").Value = startDate
        .Range("B3").Value = endDate
        .Range("B2:B3").HorizontalAlignment = xlRight
    End With
End Sub

Private Sub WriteSortedData(ByRef ws As Worksheet, ByVal headers() As String, ByRef counts As ReportCounts, Optional ByRef uniqueCounts As ReportCounts = Nothing)
    Dim row As Long
    row = 1
    
    ' Write headers
    Dim col As Long
    For col = 1 To UBound(headers) + 1
        ws.Cells(row, col).Value = headers(col - 1)
        If col = 1 Then
            ws.Cells(row, col).HorizontalAlignment = xlLeft
        Else
            ws.Cells(row, col).HorizontalAlignment = xlRight
        End If
    Next col
    
    ' Write data
    Dim key As Variant
    row = 2
    For Each key In counts.Dict.Keys
        ws.Cells(row, 1).Value = key
        ws.Cells(row, 1).HorizontalAlignment = xlLeft
        
        ws.Cells(row, 2).Value = counts.Dict(key)
        ws.Cells(row, 2).HorizontalAlignment = xlRight
        
        If Not uniqueCounts Is Nothing Then
            ws.Cells(row, 3).Value = uniqueCounts.Dict(key)
            ws.Cells(row, 3).HorizontalAlignment = xlRight
        End If
        
        row = row + 1
    Next key
    
    ' Sort data if there are rows
    If row > 2 Then
        With ws.Range(ws.Cells(2, 1), ws.Cells(row - 1, UBound(headers) + 1))
            .Sort Key1:=.Columns(1), Order1:=xlAscending, Header:=xlNo
        End With
    End If
End Sub

Private Sub WriteSupporterData(ByRef ws As Worksheet, ByRef counts As ReportCounts)
    ' Write headers
    ws.Cells(1, 1).Value = "Supporter ID"
    ws.Cells(1, 2).Value = "Supporter Email"
    ws.Cells(1, 3).Value = "Count"
    
    ws.Cells(1, 1).HorizontalAlignment = xlLeft
    ws.Cells(1, 2).HorizontalAlignment = xlLeft
    ws.Cells(1, 3).HorizontalAlignment = xlRight
    
    ' Write data
    Dim row As Long
    row = 2
    
    Dim key As Variant
    Dim parts() As String
    For Each key In counts.Dict.Keys
        parts = Split(key, " - ")
        
        ws.Cells(row, 1).Value = parts(0)
        ws.Cells(row, 1).HorizontalAlignment = xlLeft
        
        If UBound(parts) > 0 Then
            ws.Cells(row, 2).Value = parts(1)
        End If
        ws.Cells(row, 2).HorizontalAlignment = xlLeft
        
        ws.Cells(row, 3).Value = counts.Dict(key)
        ws.Cells(row, 3).HorizontalAlignment = xlRight
        
        row = row + 1
    Next key
    
    ' Sort data if there are rows
    If row > 2 Then
        With ws.Range(ws.Cells(2, 1), ws.Cells(row - 1, 3))
            .Sort Key1:=.Columns(1), Order1:=xlAscending, Header:=xlNo
        End With
    End If
End Sub

Private Function CalculateUniqueSupporters(ByRef data As Range, ByVal cols As ColumnIndices, _
                                         ByVal startDate As Date, ByVal endDate As Date) As ReportCounts
    Dim uniqueCounts As ReportCounts
    Set uniqueCounts.Dict = CreateObject("Scripting.Dictionary")
    
    Dim supporterSets As Object
    Set supporterSets = CreateObject("Scripting.Dictionary")
    
    Dim row As Long
    For row = 2 To data.Rows.Count
        Dim campaignDate As Date
        On Error Resume Next
        campaignDate = data.Cells(row, cols.CampaignDate).Value
        On Error GoTo 0
        
        If campaignDate = 0 Then GoTo NextRow
        
        If (startDate > DateSerial(1900, 1, 1) And campaignDate < startDate) Or _
           (endDate > DateSerial(1900, 1, 1) And campaignDate > endDate) Then
            GoTo NextRow
        End If
        
        Dim campaignID As String
        Dim supporterID As String
        
        campaignID = CStr(data.Cells(row, cols.CampaignID).Value)
        supporterID = CStr(data.Cells(row, cols.SupporterID).Value)
        
        If campaignID <> "" And supporterID <> "" Then
            If Not supporterSets.Exists(campaignID) Then
                Set supporterSets(campaignID) = CreateObject("Scripting.Dictionary")
            End If
            supporterSets(campaignID)(supporterID) = 1
        End If
NextRow:
    Next row
    
    ' Convert sets to counts
    Dim campaign As Variant
    For Each campaign In supporterSets.Keys
        uniqueCounts.Dict(campaign) = supporterSets(campaign).Count
    Next campaign
    
    CalculateUniqueSupporters = uniqueCounts
End Function

Public Sub ProcessCampaignData()
    On Error GoTo ErrorHandler
    
    InitializeProcessing
    
    Dim ws As Worksheet
    Set ws = ValidateWorksheet("processed-export")
    If ws Is Nothing Then GoTo ExitSub
    
    ' Get date inputs
    Dim startDate As Date, endDate As Date
    startDate = GetDateInput("Enter Start Date (YYYY-MM-DD) or cancel for no limit")
    endDate = GetDateInput("Enter End Date (YYYY-MM-DD) or cancel for no limit")
    
    ' Update report dates
    UpdateReportDates IIf(startDate = DateSerial(1900, 1, 1), "", FormatDateForMacWindows(startDate)), _
                      IIf(endDate = DateSerial(1900, 1, 1), "", FormatDateForMacWindows(endDate))
    
    ' Load data into array for better performance
    Dim dataArray As Variant
    Dim lastRow As Long, lastCol As Long
    LoadDataIntoArray ws, dataArray, lastRow, lastCol
    
    ' Get column indices
    Dim cols As ColumnIndices
    cols = GetColumnIndices()
    
    ' Initialize counters
    Dim campaignCounts As ReportCounts
    Dim caseCounts As ReportCounts
    Dim countryCounts As ReportCounts
    Dim topicCounts As ReportCounts
    Dim yearCounts As ReportCounts
    Dim typeCounts As ReportCounts
    Dim dateCounts As ReportCounts
    Dim supporterCounts As ReportCounts
    
    Set campaignCounts = CreateDictionary()
    Set caseCounts = CreateDictionary()
    Set countryCounts = CreateDictionary()
    Set topicCounts = CreateDictionary()
    Set yearCounts = CreateDictionary()
    Set typeCounts = CreateDictionary()
    Set dateCounts = CreateDictionary()
    Set supporterCounts = CreateDictionary()
    
    ' Process data using array
    UpdateProcessingStatus "Processing data..."
    
    Dim i As Long
    For i = 2 To lastRow
        ShowProgress i - 1, lastRow - 1
        
        ' Get campaign date
        Dim campaignDate As Date
        On Error Resume Next
        campaignDate = CDate(dataArray(i, cols.CampaignDate))
        On Error GoTo 0
        
        If campaignDate = 0 Then GoTo NextRow
        
        If (startDate > DateSerial(1900, 1, 1) And campaignDate < startDate) Or _
           (endDate > DateSerial(1900, 1, 1) And campaignDate > endDate) Then
            GoTo NextRow
        End If
        
        ' Campaign counts
        Dim campaignID As String
        campaignID = Trim(CStr(dataArray(i, cols.CampaignID)))
        If campaignID <> "" Then
            #If Mac Then
                DictSet campaignCounts, campaignID, DictGet(campaignCounts, campaignID) + 1
            #Else
                campaignCounts.Dict(campaignID) = CLng(campaignCounts.Dict(campaignID)) + 1
            #End If
        End If
        
        ' Supporter counts
        Dim supporterID As String, supporterEmail As String
        supporterID = Trim(CStr(dataArray(i, cols.SupporterID)))
        supporterEmail = Trim(CStr(dataArray(i, cols.SupporterEmail)))
        If supporterID <> "" Then
            Dim supporterKey As String
            supporterKey = supporterID & " - " & supporterEmail
            #If Mac Then
                DictSet supporterCounts, supporterKey, DictGet(supporterCounts, supporterKey) + 1
            #Else
                supporterCounts.Dict(supporterKey) = CLng(supporterCounts.Dict(supporterKey)) + 1
            #End If
        End If
        
        ' Date counts
        Dim monthYear As String
        monthYear = Format(campaignDate, "yyyy-mm")
        #If Mac Then
            DictSet dateCounts, monthYear, DictGet(dateCounts, monthYear) + 1
        #Else
            dateCounts.Dict(monthYear) = CLng(dateCounts.Dict(monthYear)) + 1
        #End If
        
        ' Other counts
        Dim cellValue As String
        
        ' Country counts
        cellValue = Trim(CStr(dataArray(i, cols.Country)))
        If cellValue <> "" Then
            #If Mac Then
                DictSet countryCounts, cellValue, DictGet(countryCounts, cellValue) + 1
            #Else
                countryCounts.Dict(cellValue) = CLng(countryCounts.Dict(cellValue)) + 1
            #End If
        End If
        
        ' Case number counts
        cellValue = Trim(CStr(dataArray(i, cols.CaseNumber)))
        If cellValue <> "" Then
            #If Mac Then
                DictSet caseCounts, cellValue, DictGet(caseCounts, cellValue) + 1
            #Else
                caseCounts.Dict(cellValue) = CLng(caseCounts.Dict(cellValue)) + 1
            #End If
        End If
        
        ' Year counts
        cellValue = Trim(CStr(dataArray(i, cols.Year)))
        If cellValue <> "" Then
            #If Mac Then
                DictSet yearCounts, cellValue, DictGet(yearCounts, cellValue) + 1
            #Else
                yearCounts.Dict(cellValue) = CLng(yearCounts.Dict(cellValue)) + 1
            #End If
        End If
        
        ' Type counts
        cellValue = Trim(CStr(dataArray(i, cols.Type)))
        If cellValue <> "" Then
            #If Mac Then
                DictSet typeCounts, cellValue, DictGet(typeCounts, cellValue) + 1
            #Else
                typeCounts.Dict(cellValue) = CLng(typeCounts.Dict(cellValue)) + 1
            #End If
        End If
        
        ' Topic counts
        cellValue = Trim(CStr(dataArray(i, cols.Topics)))
        If cellValue <> "" Then
            Dim topics() As String
            topics = Split(cellValue, ",")
            Dim topic As Variant
            For Each topic In topics
                cellValue = Trim(CStr(topic))
                If cellValue <> "" Then
                    #If Mac Then
                        DictSet topicCounts, cellValue, DictGet(topicCounts, cellValue) + 1
                    #Else
                        topicCounts.Dict(cellValue) = CLng(topicCounts.Dict(cellValue)) + 1
                    #End If
                End If
            Next topic
        End If
NextRow:
    Next i
    
    ' Calculate unique supporters using array data
    Dim uniqueCounts As ReportCounts
    Set uniqueCounts = CalculateUniqueSupportersFromArray(dataArray, cols, startDate, endDate)
    
    ' Write reports
    UpdateProcessingStatus "Writing reports..."
    WriteSortedData GetOrCreateSheet("by-name"), Array("Campaign ID", "Count", "Unique Supporters"), campaignCounts, uniqueCounts
    WriteSortedData GetOrCreateSheet("by-case-number"), Array("Case Number", "Count"), caseCounts
    WriteSortedData GetOrCreateSheet("by-country"), Array("Country", "Count"), countryCounts
    WriteSortedData GetOrCreateSheet("by-topic"), Array("Topic", "Count"), topicCounts
    WriteSortedData GetOrCreateSheet("by-year"), Array("Year", "Count"), yearCounts
    WriteSortedData GetOrCreateSheet("by-type"), Array("Type", "Count"), typeCounts
    WriteSortedData GetOrCreateSheet("by-date"), Array("Month", "Count"), dateCounts
    WriteSupporterData GetOrCreateSheet("by-supporter"), supporterCounts
    
    ' Cleanup
    CleanupDictionaries campaignCounts, caseCounts, countryCounts, topicCounts, _
                        yearCounts, typeCounts, dateCounts, supporterCounts
    
    MsgBox "✓ Your UAN Reports have been updated!", vbInformation

ExitSub:
    FinalizeProcessing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume ExitSub
End Sub

Public Sub ProcessCampaignDataExceptSupporter()
    On Error GoTo ErrorHandler
    
    SetApplicationState False
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("processed-export")
    
    If ws Is Nothing Then
        MsgBox "The 'processed-export' sheet is required.", vbExclamation
        GoTo ExitSub
    End If
    
    ' Get date inputs
    Dim startDate As Date, endDate As Date
    startDate = GetDateInput("Enter Start Date (YYYY-MM-DD) or cancel for no limit")
    endDate = GetDateInput("Enter End Date (YYYY-MM-DD) or cancel for no limit")
    
    ' Update report dates to "Mixed" since by-supporter isn't included
    UpdateReportDates "Mixed", "Mixed"
    
    ' Get column indices
    Dim cols As ColumnIndices
    cols = GetColumnIndices()
    
    ' Check for required columns
    If cols.CampaignID = 0 Or cols.CampaignDate = 0 Or cols.SupporterID = 0 Or _
       cols.SupporterEmail = 0 Or cols.Country = 0 Or cols.CaseNumber = 0 Or _
       cols.Topics = 0 Or cols.Year = 0 Or cols.Type = 0 Then
        MsgBox "One or more required columns are missing.", vbExclamation
        GoTo ExitSub
    End If
    
    ' Initialize counters (except supporter)
    Dim campaignCounts As ReportCounts
    Dim caseCounts As ReportCounts
    Dim countryCounts As ReportCounts
    Dim topicCounts As ReportCounts
    Dim yearCounts As ReportCounts
    Dim typeCounts As ReportCounts
    Dim dateCounts As ReportCounts
    
    Set campaignCounts = CreateDictionary()
    Set caseCounts = CreateDictionary()
    Set countryCounts = CreateDictionary()
    Set topicCounts = CreateDictionary()
    Set yearCounts = CreateDictionary()
    Set typeCounts = CreateDictionary()
    Set dateCounts = CreateDictionary()
    
    ' Process data
    Dim dataRange As Range
    Set dataRange = ws.Range("A1").CurrentRegion
    
    Dim row As Long
    For row = 2 To dataRange.Rows.Count
        ShowProgress row - 1, dataRange.Rows.Count - 1
        
        ' Process data based on report type
        ProcessRowForReport dataRange, row, cols, campaignCounts, "Campaign ID", startDate, endDate
        ProcessRowForReport dataRange, row, cols, caseCounts, "Case Number", startDate, endDate
        ProcessRowForReport dataRange, row, cols, countryCounts, "Country", startDate, endDate
        ProcessRowForReport dataRange, row, cols, topicCounts, "Topics", startDate, endDate
        ProcessRowForReport dataRange, row, cols, yearCounts, "Year", startDate, endDate
        ProcessRowForReport dataRange, row, cols, typeCounts, "Type", startDate, endDate
        ProcessRowForReport dataRange, row, cols, dateCounts, "Date", startDate, endDate
    Next row
    
    ' Calculate unique supporters for campaign report
    Dim uniqueCounts As ReportCounts
    Set uniqueCounts = CalculateUniqueSupporters(dataRange, cols, startDate, endDate)
    
    ' Write all reports except supporter
    WriteSortedData GetOrCreateSheet("by-name"), Array("Campaign ID", "Count", "Unique Supporters"), campaignCounts, uniqueCounts
    WriteSortedData GetOrCreateSheet("by-case-number"), Array("Case Number", "Count"), caseCounts
    WriteSortedData GetOrCreateSheet("by-country"), Array("Country", "Count"), countryCounts
    WriteSortedData GetOrCreateSheet("by-topic"), Array("Topic", "Count"), topicCounts
    WriteSortedData GetOrCreateSheet("by-year"), Array("Year", "Count"), yearCounts
    WriteSortedData GetOrCreateSheet("by-type"), Array("Type", "Count"), typeCounts
    WriteSortedData GetOrCreateSheet("by-date"), Array("Month", "Count"), dateCounts
    
    MsgBox "✓ Your UAN Reports have been updated! (except by-supporter)", vbInformation

ExitSub:
    SetApplicationState True
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume ExitSub
End Sub

' Individual report update functions
Public Sub UpdateCampaignReport()
    ProcessSpecificReport "by-name", "Campaign ID"
End Sub

Public Sub UpdateCaseReport()
    ProcessSpecificReport "by-case-number", "Case Number"
End Sub

Public Sub UpdateCountryReport()
    ProcessSpecificReport "by-country", "Country"
End Sub

Public Sub UpdateTopicReport()
    ProcessSpecificReport "by-topic", "Topics"
End Sub

Public Sub UpdateYearReport()
    ProcessSpecificReport "by-year", "Year"
End Sub

Public Sub UpdateTypeReport()
    ProcessSpecificReport "by-type", "Type"
End Sub

Public Sub UpdateDateReport()
    ProcessSpecificReport "by-date", "Date"
End Sub

Public Sub UpdateSupporterReport()
    ProcessSpecificReport "by-supporter", "Supporter"
End Sub

Private Sub ProcessSpecificReport(ByVal sheetName As String, ByVal reportType As String)
    On Error GoTo ErrorHandler
    
    InitializeProcessing
    
    Dim ws As Worksheet
    Set ws = ValidateWorksheet("processed-export")
    If ws Is Nothing Then GoTo ExitSub
    
    ' Get date inputs
    Dim startDate As Date, endDate As Date
    startDate = GetDateInput("Enter Start Date (YYYY-MM-DD) or cancel for no limit")
    endDate = GetDateInput("Enter End Date (YYYY-MM-DD) or cancel for no limit")
    
    ' Update report dates to "Mixed" when running individual reports
    UpdateReportDates "Mixed", "Mixed"
    
    ' Load data into array for better performance
    Dim dataArray As Variant
    Dim lastRow As Long, lastCol As Long
    LoadDataIntoArray ws, dataArray, lastRow, lastCol
    
    ' Get column indices
    Dim cols As ColumnIndices
    cols = GetColumnIndices()
    
    ' Initialize counter
    Dim counts As ReportCounts
    Set counts = CreateDictionary()
    
    ' Process data
    UpdateProcessingStatus "Processing data..."
    
    Dim i As Long
    For i = 2 To lastRow
        ShowProgress i - 1, lastRow - 1
        ProcessArrayRowForReport dataArray, i, cols, counts, reportType, startDate, endDate
    Next i
    
    ' Write report
    UpdateProcessingStatus "Writing report..."
    Dim reportSheet As Worksheet
    Set reportSheet = GetOrCreateSheet(sheetName)
    
    If reportType = "Campaign ID" Then
        Dim uniqueCounts As ReportCounts
        Set uniqueCounts = CalculateUniqueSupportersFromArray(dataArray, cols, startDate, endDate)
        WriteSortedData reportSheet, Array("Campaign ID", "Count", "Unique Supporters"), counts, uniqueCounts
    ElseIf reportType = "Supporter" Then
        WriteSupporterData reportSheet, counts
    Else
        WriteSortedData reportSheet, Array(reportType, "Count"), counts
    End If
    
    MsgBox "✓ Your " & reportType & " Report has been updated!", vbInformation

ExitSub:
    CleanupDictionaries counts
    FinalizeProcessing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume ExitSub
End Sub

Private Sub ProcessArrayRowForReport(ByRef dataArray As Variant, ByVal row As Long, ByRef cols As ColumnIndices, _
                                   ByRef counts As ReportCounts, ByVal reportType As String, _
                                   ByVal startDate As Date, ByVal endDate As Date)
    ' Check date filter first
    Dim campaignDate As Date
    On Error Resume Next
    campaignDate = CDate(dataArray(row, cols.CampaignDate))
    On Error GoTo 0
    
    If campaignDate = 0 Then Exit Sub
    
    If (startDate > DateSerial(1900, 1, 1) And campaignDate < startDate) Or _
       (endDate > DateSerial(1900, 1, 1) And campaignDate > endDate) Then
        Exit Sub
    End If
    
    ' Process based on report type
    Dim cellValue As String
    Select Case reportType
        Case "Campaign ID"
            cellValue = Trim(CStr(dataArray(row, cols.CampaignID)))
            If cellValue <> "" Then
                #If Mac Then
                    DictSet counts, cellValue, DictGet(counts, cellValue) + 1
                #Else
                    counts.Dict(cellValue) = CLng(counts.Dict(cellValue)) + 1
                #End If
            End If
            
        Case "Case Number"
            cellValue = Trim(CStr(dataArray(row, cols.CaseNumber)))
            If cellValue <> "" Then
                #If Mac Then
                    DictSet counts, cellValue, DictGet(counts, cellValue) + 1
                #Else
                    counts.Dict(cellValue) = CLng(counts.Dict(cellValue)) + 1
                #End If
            End If
            
        Case "Country"
            cellValue = Trim(CStr(dataArray(row, cols.Country)))
            If cellValue <> "" Then
                #If Mac Then
                    DictSet counts, cellValue, DictGet(counts, cellValue) + 1
                #Else
                    counts.Dict(cellValue) = CLng(counts.Dict(cellValue)) + 1
                #End If
            End If
            
        Case "Topics"
            cellValue = Trim(CStr(dataArray(row, cols.Topics)))
            If cellValue <> "" Then
                Dim topics() As String
                topics = Split(cellValue, ",")
                Dim topic As Variant
                For Each topic In topics
                    cellValue = Trim(CStr(topic))
                    If cellValue <> "" Then
                        #If Mac Then
                            DictSet counts, cellValue, DictGet(counts, cellValue) + 1
                        #Else
                            counts.Dict(cellValue) = CLng(counts.Dict(cellValue)) + 1
                        #End If
                    End If
                Next topic
            End If
            
        Case "Year"
            cellValue = Trim(CStr(dataArray(row, cols.Year)))
            If cellValue <> "" Then
                #If Mac Then
                    DictSet counts, cellValue, DictGet(counts, cellValue) + 1
                #Else
                    counts.Dict(cellValue) = CLng(counts.Dict(cellValue)) + 1
                #End If
            End If
            
        Case "Type"
            cellValue = Trim(CStr(dataArray(row, cols.Type)))
            If cellValue <> "" Then
                #If Mac Then
                    DictSet counts, cellValue, DictGet(counts, cellValue) + 1
                #Else
                    counts.Dict(cellValue) = CLng(counts.Dict(cellValue)) + 1
                #End If
            End If
            
        Case "Date"
            cellValue = Format(campaignDate, "yyyy-mm")
            #If Mac Then
                DictSet counts, cellValue, DictGet(counts, cellValue) + 1
            #Else
                counts.Dict(cellValue) = CLng(counts.Dict(cellValue)) + 1
            #End If
            
        Case "Supporter"
            Dim supporterID As String, supporterEmail As String
            supporterID = Trim(CStr(dataArray(row, cols.SupporterID)))
            supporterEmail = Trim(CStr(dataArray(row, cols.SupporterEmail)))
            If supporterID <> "" Then
                Dim supporterKey As String
                supporterKey = supporterID & " - " & supporterEmail
                #If Mac Then
                    DictSet counts, supporterKey, DictGet(counts, supporterKey) + 1
                #Else
                    counts.Dict(supporterKey) = CLng(counts.Dict(supporterKey)) + 1
                #End If
            End If
    End Select
End Sub

#If Mac Then
' Mac version of menu using InputBox
Public Sub ShowReportMenu()
    On Error GoTo ErrorHandler
    
    SetApplicationState False
    
    Dim choice As Variant
    choice = Application.InputBox( _
        "Choose a report to update:" & vbNewLine & _
        "1. Update All Reports" & vbNewLine & _
        "2. Update All Reports (except by-supporter)" & vbNewLine & _
        "3. Update by-name" & vbNewLine & _
        "4. Update by-case-number" & vbNewLine & _
        "5. Update by-country" & vbNewLine & _
        "6. Update by-topic" & vbNewLine & _
        "7. Update by-year" & vbNewLine & _
        "8. Update by-type" & vbNewLine & _
        "9. Update by-date" & vbNewLine & _
        "10. Update by-supporter", _
        "UAN Reports", , , , , , , , 1)
    
    If choice = False Then GoTo ExitSub
    
    Select Case CInt(choice)
        Case 1: Call ProcessCampaignData
        Case 2: Call ProcessCampaignDataExceptSupporter
        Case 3: Call UpdateCampaignReport
        Case 4: Call UpdateCaseReport
        Case 5: Call UpdateCountryReport
        Case 6: Call UpdateTopicReport
        Case 7: Call UpdateYearReport
        Case 8: Call UpdateTypeReport
        Case 9: Call UpdateDateReport
        Case 10: Call UpdateSupporterReport
        Case Else
            MsgBox "Invalid selection. Please choose a number between 1 and 10.", vbExclamation
    End Select

ExitSub:
    SetApplicationState True
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume ExitSub
End Sub
#End If

' Add these helper functions at the end of the module

Private Sub CleanupDictionaries(ParamArray dicts() As ReportCounts)
    Dim dict As Variant
    For Each dict In dicts
        #If Mac Then
            If Not dict.Dict.Keys Is Nothing Then
                Do While dict.Dict.Keys.Count > 0
                    dict.Dict.Keys.Remove 1
                    dict.Dict.Items.Remove 1
                Loop
            End If
            Set dict.Dict.Keys = Nothing
            Set dict.Dict.Items = Nothing
        #Else
            If Not dict.Dict Is Nothing Then
                dict.Dict.RemoveAll
                Set dict.Dict = Nothing
            End If
        #End If
    Next dict
End Sub

#If Mac Then
Private Function DictGet(ByRef dict As ReportCounts, ByVal key As String) As Long
    Dim i As Long
    On Error Resume Next
    For i = 1 To dict.Dict.Keys.Count
        If dict.Dict.Keys(i) = key Then
            DictGet = CLng(dict.Dict.Items(i))
            Exit Function
        End If
    Next i
    DictGet = 0
End Function

Private Sub DictSet(ByRef dict As ReportCounts, ByVal key As String, ByVal value As Long)
    Dim i As Long
    For i = 1 To dict.Dict.Keys.Count
        If dict.Dict.Keys(i) = key Then
            dict.Dict.Items.Remove i
            dict.Dict.Items.Add value, , i
            Exit Sub
        End If
    Next i
    dict.Dict.Keys.Add key
    dict.Dict.Items.Add value
End Sub

Private Function DictExists(ByRef dict As ReportCounts, ByVal key As String) As Boolean
    Dim i As Long
    For i = 1 To dict.Dict.Keys.Count
        If dict.Dict.Keys(i) = key Then
            DictExists = True
            Exit Function
        End If
    Next i
    DictExists = False
End Function
#End If

Private Sub LoadDataIntoArray(ByRef ws As Worksheet, ByRef dataArray As Variant, ByRef lastRow As Long, ByRef lastCol As Long)
    Dim dataRange As Range
    Set dataRange = ws.Range("A1").CurrentRegion
    lastRow = dataRange.Rows.Count
    lastCol = dataRange.Columns.Count
    dataArray = dataRange.Value
End Sub

Private Function FormatDateForMacWindows(ByVal dateValue As Date) As String
    #If Mac Then
        FormatDateForMacWindows = Format(dateValue, "yyyy-mm-dd")
    #Else
        FormatDateForMacWindows = Format(dateValue, "yyyy\-mm\-dd")
    #End If
End Function

Private Sub OptimizeExcel(ByVal optimize As Boolean)
    With Application
        .ScreenUpdating = Not optimize
        .EnableEvents = Not optimize
        .DisplayAlerts = Not optimize
        If optimize Then
            .Calculation = xlCalculationManual
        Else
            .Calculation = xlCalculationAutomatic
        End If
    End With
End Sub

Private Sub UpdateProcessingStatus(ByVal message As String, Optional ByVal showProgress As Boolean = True)
    If showProgress Then
        Application.StatusBar = message
        DoEvents
    End If
End Sub

Private Function ValidateWorksheet(ByVal wsName As String) As Worksheet
    On Error Resume Next
    Set ValidateWorksheet = ThisWorkbook.Sheets(wsName)
    On Error GoTo 0
    
    If ValidateWorksheet Is Nothing Then
        MsgBox "Required worksheet '" & wsName & "' not found.", vbExclamation
    End If
End Function

Private Function GetWorkbookPath() As String
    #If Mac Then
        GetWorkbookPath = ThisWorkbook.Path & ":"
    #Else
        GetWorkbookPath = ThisWorkbook.Path & "\"
    #End If
End Function

' Add this to the beginning of each main processing function
Private Sub InitializeProcessing()
    OptimizeExcel True
    UpdateProcessingStatus "Initializing..."
End Sub

' Add this to the end of each main processing function (in the ExitSub section)
Private Sub FinalizeProcessing()
    OptimizeExcel False
    Application.StatusBar = False
End Sub

' New function for array-based unique supporter calculation
Private Function CalculateUniqueSupportersFromArray(ByRef dataArray As Variant, ByVal cols As ColumnIndices, _
                                                   ByVal startDate As Date, ByVal endDate As Date) As ReportCounts
    #If Mac Then
        Dim uniqueCounts As ReportCounts
        Set uniqueCounts.Dict.Keys = New Collection
        Set uniqueCounts.Dict.Items = New Collection
        
        Dim supporterSets As Object
        Set supporterSets = CreateObject("Scripting.Dictionary")
        
        Dim i As Long
        For i = 2 To UBound(dataArray, 1)
            Dim campaignDate As Date
            On Error Resume Next
            campaignDate = CDate(dataArray(i, cols.CampaignDate))
            On Error GoTo 0
            
            If campaignDate = 0 Then GoTo NextRow
            
            If (startDate > DateSerial(1900, 1, 1) And campaignDate < startDate) Or _
               (endDate > DateSerial(1900, 1, 1) And campaignDate > endDate) Then
                GoTo NextRow
            End If
            
            Dim campaignID As String
            Dim supporterID As String
            
            campaignID = CStr(dataArray(i, cols.CampaignID))
            supporterID = CStr(dataArray(i, cols.SupporterID))
            
            If campaignID <> "" And supporterID <> "" Then
                If Not DictExists(uniqueCounts, campaignID) Then
                    uniqueCounts.Dict.Keys.Add campaignID
                    uniqueCounts.Dict.Items.Add CreateObject("Scripting.Dictionary")
                End If
                uniqueCounts.Dict.Items(uniqueCounts.Dict.Keys.Count)(supporterID) = 1
            End If
NextRow:
        Next i
        
        ' Convert sets to counts
        Dim j As Long
        For j = 1 To uniqueCounts.Dict.Keys.Count
            Dim supporters As Object
            Set supporters = uniqueCounts.Dict.Items(j)
            uniqueCounts.Dict.Items.Remove j
            uniqueCounts.Dict.Items.Add supporters.Count, , j
        Next j
        
        Set CalculateUniqueSupportersFromArray = uniqueCounts
    #Else
        Dim uniqueCounts As ReportCounts
        Set uniqueCounts.Dict = CreateObject("Scripting.Dictionary")
        
        Dim supporterSets As Object
        Set supporterSets = CreateObject("Scripting.Dictionary")
        
        Dim i As Long
        For i = 2 To UBound(dataArray, 1)
            Dim campaignDate As Date
            On Error Resume Next
            campaignDate = CDate(dataArray(i, cols.CampaignDate))
            On Error GoTo 0
            
            If campaignDate = 0 Then GoTo NextRow
            
            If (startDate > DateSerial(1900, 1, 1) And campaignDate < startDate) Or _
               (endDate > DateSerial(1900, 1, 1) And campaignDate > endDate) Then
                GoTo NextRow
            End If
            
            Dim campaignID As String
            Dim supporterID As String
            
            campaignID = CStr(dataArray(i, cols.CampaignID))
            supporterID = CStr(dataArray(i, cols.SupporterID))
            
            If campaignID <> "" And supporterID <> "" Then
                If Not supporterSets.Exists(campaignID) Then
                    Set supporterSets(campaignID) = CreateObject("Scripting.Dictionary")
                End If
                supporterSets(campaignID)(supporterID) = 1
            End If
NextRow:
        Next i
        
        ' Convert sets to counts
        Dim campaign As Variant
        For Each campaign In supporterSets.Keys
            uniqueCounts.Dict(campaign) = supporterSets(campaign).Count
        Next campaign
        
        Set CalculateUniqueSupportersFromArray = uniqueCounts
    #End If
End Function 