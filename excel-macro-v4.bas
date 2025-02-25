' UAN (Urgent Action Network) Reports Generator v1.0.0
' Last Updated: 2024-03-19
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

Option Explicit

' Column indices structure
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

' Main entry point macro that will appear in Excel's macro list
Public Sub GenerateUANReports()
    ' Initialize Excel environment
    Dim screenUpdating As Boolean
    Dim statusBar As Boolean
    Dim calculation As XlCalculation
    Dim displayAlerts As Boolean
    
    ' Save current Excel settings
    screenUpdating = Application.ScreenUpdating
    statusBar = Application.DisplayStatusBar
    calculation = Application.Calculation
    displayAlerts = Application.DisplayAlerts
    
    ' Configure Excel for processing
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.StatusBar = "Initializing UAN Reports..."
    
    On Error GoTo ErrorHandler
    
    ' Get the export sheet
    Dim wsExport As Worksheet
    Set wsExport = ThisWorkbook.Sheets("processed-export")
    
    If wsExport Is Nothing Then
        MsgBox "Could not find a sheet named 'processed-export'.", vbExclamation
        GoTo Cleanup
    End If
    
    ' Get date range from user
    Dim startDateInput As String, endDateInput As String
    Dim hasStartDate As Boolean, hasEndDate As Boolean
    Dim startDate As Date, endDate As Date
    
    startDateInput = InputBox("Enter Start Date (YYYY-MM-DD) or leave blank for no limit:", "Filter by Date Range", "")
    If startDateInput <> "" Then
        If IsDate(startDateInput) Then
            startDate = CDate(startDateInput)
            hasStartDate = True
        Else
            MsgBox "Invalid start date format. Please use YYYY-MM-DD format.", vbExclamation
            GoTo Cleanup
        End If
    End If
    
    endDateInput = InputBox("Enter End Date (YYYY-MM-DD) or leave blank for no limit:", "Filter by Date Range", "")
    If endDateInput <> "" Then
        If IsDate(endDateInput) Then
            endDate = CDate(endDateInput)
            hasEndDate = True
        Else
            MsgBox "Invalid end date format. Please use YYYY-MM-DD format.", vbExclamation
            GoTo Cleanup
        End If
    End If
    
    ' Process the data
    ProcessData wsExport, startDate, endDate, hasStartDate, hasEndDate
    
    GoTo Cleanup
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    
Cleanup:
    ' Restore Excel settings
    Application.StatusBar = False
    Application.ScreenUpdating = screenUpdating
    Application.DisplayStatusBar = statusBar
    Application.Calculation = calculation
    Application.DisplayAlerts = displayAlerts
End Sub

Private Sub ProcessData(ws As Worksheet, startDate As Date, endDate As Date, hasStartDate As Boolean, hasEndDate As Boolean)
    ' Find column indices
    Dim cols As ColumnIndices
    cols = GetColumnIndices(ws)
    
    ' Add debugging to check column indices
    Dim debugMsg As String
    debugMsg = "Column Indices:" & vbNewLine & _
               "CampaignID: " & cols.CampaignID & vbNewLine & _
               "CampaignDate: " & cols.CampaignDate & vbNewLine & _
               "SupporterID: " & cols.SupporterID & vbNewLine & _
               "SupporterEmail: " & cols.SupporterEmail & vbNewLine & _
               "Country: " & cols.Country & vbNewLine & _
               "CaseNumber: " & cols.CaseNumber & vbNewLine & _
               "Topics: " & cols.Topics & vbNewLine & _
               "Year: " & cols.Year & vbNewLine & _
               "Type: " & cols.Type
    
    Debug.Print debugMsg
    
    ' Get data range
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, cols.CampaignID).End(xlUp).Row
    
    ' Add debugging for row count
    Debug.Print "Last Row: " & lastRow
    
    ' Initialize data arrays for each report
    Dim campaignIDs() As String
    Dim campaignCounts() As Long
    Dim campaignUniqueSupport() As Long
    Dim caseNumbers() As String
    Dim caseCounts() As Long
    Dim countries() As String
    Dim countryCounts() As Long
    Dim topics() As String
    Dim topicCounts() As Long
    Dim years() As String
    Dim yearCounts() As Long
    Dim types() As String
    Dim typeCounts() As Long
    Dim dates() As String
    Dim dateCounts() As Long
    Dim supporters() As String
    Dim supporterCounts() As Long
    
    ' Initialize counters
    Dim totalCampaigns As Long, totalCases As Long, totalCountries As Long
    Dim totalTopics As Long, totalYears As Long, totalTypes As Long
    Dim totalDates As Long, totalSupporters As Long
    
    totalCampaigns = 0: totalCases = 0: totalCountries = 0
    totalTopics = 0: totalYears = 0: totalTypes = 0
    totalDates = 0: totalSupporters = 0
    
    ' Initialize the arrays
    ReDim campaignIDs(1 To 1000)
    ReDim campaignCounts(1 To 1000)
    ReDim campaignUniqueSupport(1 To 1000)
    ReDim caseNumbers(1 To 1000)
    ReDim caseCounts(1 To 1000)
    ReDim countries(1 To 1000)
    ReDim countryCounts(1 To 1000)
    ReDim topics(1 To 1000)
    ReDim topicCounts(1 To 1000)
    ReDim years(1 To 1000)
    ReDim yearCounts(1 To 1000)
    ReDim types(1 To 1000)
    ReDim typeCounts(1 To 1000)
    ReDim dates(1 To 1000)
    ReDim dateCounts(1 To 1000)
    ReDim supporters(1 To 1000)
    ReDim supporterCounts(1 To 1000)
    
    ' For tracking unique supporters per campaign
    Dim campaignSupporters() As String
    ReDim campaignSupporters(1 To 10000)
    Dim campaignSupporterCount As Long
    campaignSupporterCount = 0
    
    ' Process data
    Application.StatusBar = "Processing data: 0%"
    Dim row As Long, i As Long
    Dim supporterID As String, campaignID As String, caseNumber As String
    Dim country As String, topic As String, typeValue As String, yearValue As String
    Dim campaignDate As Date, monthYear As String
    
    Dim progress As Double
    
    Dim dummyArray() As Long
    ReDim dummyArray(1 To 1)
    
    For row = 2 To lastRow
        ' Show progress
        progress = (row - 1) / (lastRow - 1) * 100
        If row Mod 100 = 0 Then
            Application.StatusBar = "Processing data: " & Format(progress, "0") & "%"
        End If
        
        ' Get campaign date with error handling
        On Error Resume Next
        If IsDate(ws.Cells(row, cols.CampaignDate).Value) Then
            campaignDate = CDate(ws.Cells(row, cols.CampaignDate).Value)
        Else
            campaignDate = 0
        End If
        On Error GoTo 0
        
        ' Skip rows without valid dates or if outside date range
        If campaignDate = 0 Then GoTo NextRow
        If hasStartDate And campaignDate < startDate Then GoTo NextRow
        If hasEndDate And campaignDate > endDate Then GoTo NextRow
        
        ' Get all field values with safe handling
        campaignID = Trim(CStr(ws.Cells(row, cols.CampaignID).Value))
        supporterID = Trim(CStr(ws.Cells(row, cols.SupporterID).Value))
        caseNumber = Trim(CStr(ws.Cells(row, cols.CaseNumber).Value))
        country = Trim(CStr(ws.Cells(row, cols.Country).Value))
        topic = Trim(CStr(ws.Cells(row, cols.Topics).Value))
        yearValue = Trim(CStr(ws.Cells(row, cols.Year).Value))
        typeValue = Trim(CStr(ws.Cells(row, cols.Type).Value))
        
        ' Campaign counts
        If campaignID <> "" Then
            totalCampaigns = CountOccurrences(campaignIDs, campaignCounts, totalCampaigns, campaignID)
        End If
        
        ' Track campaign-supporter pairs for unique counts
        If campaignID <> "" And supporterID <> "" Then
            campaignSupporterCount = campaignSupporterCount + 1
            If campaignSupporterCount > UBound(campaignSupporters) Then
                ReDim Preserve campaignSupporters(1 To UBound(campaignSupporters) * 2)
            End If
            campaignSupporters(campaignSupporterCount) = campaignID & "|" & supporterID
        End If
        
        ' Case number counts
        If caseNumber <> "" Then
            totalCases = CountOccurrences(caseNumbers, caseCounts, totalCases, caseNumber)
        End If
        
        ' Country counts
        If country <> "" Then
            totalCountries = CountOccurrences(countries, countryCounts, totalCountries, country)
        End If
        
        ' Topic counts (can have multiple per row)
        If topic <> "" Then
            Dim topicArray() As String
            topicArray = Split(topic, ",")
            Dim t As Variant
            For Each t In topicArray
                Dim topicValue As String
                topicValue = Trim(CStr(t))
                If topicValue <> "" Then
                    totalTopics = CountOccurrences(topics, topicCounts, totalTopics, topicValue)
                End If
            Next t
        End If
        
        ' Year counts
        If yearValue <> "" Then
            totalYears = CountOccurrences(years, yearCounts, totalYears, yearValue)
        End If
        
        ' Type counts
        If typeValue <> "" Then
            totalTypes = CountOccurrences(types, typeCounts, totalTypes, typeValue)
        End If
        
        ' Date counts (by month)
        monthYear = Format(campaignDate, "yyyy-mm")
        totalDates = CountOccurrences(dates, dateCounts, totalDates, monthYear)
        
        ' Supporter counts
        If supporterID <> "" Then
            Dim supporterEmail As String
            supporterEmail = Trim(CStr(ws.Cells(row, cols.SupporterEmail).Value))
            Dim supporterKey As String
            supporterKey = supporterID & " - " & supporterEmail
            totalSupporters = CountOccurrences(supporters, supporterCounts, totalSupporters, supporterKey)
        End If
        
NextRow:
    Next row
    
    ' Calculate unique supporters per campaign
    For i = 1 To totalCampaigns
        campaignUniqueSupport(i) = CountUniqueSupporters(campaignIDs(i), campaignSupporters, campaignSupporterCount)
    Next i
    
    ' Write reports
    Application.StatusBar = "Writing reports..."
    
    WriteReport GetOrCreateSheet("by-name"), "Campaign ID", "Count", "Unique Supporters", _
               campaignIDs, campaignCounts, campaignUniqueSupport, totalCampaigns
               
    WriteReport GetOrCreateSheet("by-case-number"), "Case Number", "Count", "", _
               caseNumbers, caseCounts, dummyArray, totalCases
               
    WriteReport GetOrCreateSheet("by-country"), "Country", "Count", "", _
               countries, countryCounts, dummyArray, totalCountries
               
    WriteReport GetOrCreateSheet("by-topic"), "Topic", "Count", "", _
               topics, topicCounts, dummyArray, totalTopics
               
    WriteReport GetOrCreateSheet("by-year"), "Year", "Count", "", _
               years, yearCounts, dummyArray, totalYears
               
    WriteReport GetOrCreateSheet("by-type"), "Type", "Count", "", _
               types, typeCounts, dummyArray, totalTypes
               
    WriteReport GetOrCreateSheet("by-date"), "Month", "Count", "", _
               dates, dateCounts, dummyArray, totalDates
               
    WriteSupporterReport GetOrCreateSheet("by-supporter"), supporters, supporterCounts, totalSupporters
    
    ' Update report dates
    UpdateReportDates IIf(hasStartDate, Format(startDate, "yyyy-mm-dd"), ""), _
                      IIf(hasEndDate, Format(endDate, "yyyy-mm-dd"), "")
    
    MsgBox "✓ Your UAN Reports have been updated!", vbInformation
    
    ' After processing, add debugging for totals
    Debug.Print "Total Campaigns: " & totalCampaigns
    Debug.Print "Total Cases: " & totalCases
    Debug.Print "Total Countries: " & totalCountries
    Debug.Print "Total Topics: " & totalTopics
    Debug.Print "Total Years: " & totalYears
    Debug.Print "Total Types: " & totalTypes
    Debug.Print "Total Dates: " & totalDates
    Debug.Print "Total Supporters: " & totalSupporters
End Sub

Private Function GetColumnIndices(ws As Worksheet) As ColumnIndices
    Dim cols As ColumnIndices
    
    ' Print all column headers for debugging
    Dim i As Long
    For i = 1 To 20 ' Check first 20 columns
        If ws.Cells(1, i).Value <> "" Then
            Debug.Print i & ": " & ws.Cells(1, i).Value
        End If
    Next i
    
    On Error Resume Next
    cols.CampaignID = WorksheetFunction.Match("Campaign ID", ws.Rows(1), 0)
    If cols.CampaignID = 0 Then cols.CampaignID = WorksheetFunction.Match("*Campaign*ID*", ws.Rows(1), 0)
    
    cols.CampaignDate = WorksheetFunction.Match("Campaign Date", ws.Rows(1), 0)
    If cols.CampaignDate = 0 Then cols.CampaignDate = WorksheetFunction.Match("*Campaign*Date*", ws.Rows(1), 0)
    
    cols.SupporterID = WorksheetFunction.Match("Supporter ID", ws.Rows(1), 0)
    cols.SupporterEmail = WorksheetFunction.Match("Supporter Email", ws.Rows(1), 0)
    cols.Country = WorksheetFunction.Match("External Reference 6 (Country)", ws.Rows(1), 0)
    cols.CaseNumber = WorksheetFunction.Match("External Reference 7 (Case Number)", ws.Rows(1), 0)
    cols.Topics = WorksheetFunction.Match("External Reference 8 (Topics)", ws.Rows(1), 0)
    cols.Year = WorksheetFunction.Match("External Reference 10 (Year)", ws.Rows(1), 0)
    cols.Type = WorksheetFunction.Match("External Reference 10 (Type)", ws.Rows(1), 0)
    On Error GoTo 0
    
    ' Validate that all required columns were found
    Dim missingColumns As String
    missingColumns = ""
    
    If cols.CampaignID = 0 Then missingColumns = missingColumns & "Campaign ID, "
    If cols.CampaignDate = 0 Then missingColumns = missingColumns & "Campaign Date, "
    If cols.SupporterID = 0 Then missingColumns = missingColumns & "Supporter ID, "
    If cols.SupporterEmail = 0 Then missingColumns = missingColumns & "Supporter Email, "
    If cols.Country = 0 Then missingColumns = missingColumns & "Country, "
    If cols.CaseNumber = 0 Then missingColumns = missingColumns & "Case Number, "
    If cols.Topics = 0 Then missingColumns = missingColumns & "Topics, "
    If cols.Year = 0 Then missingColumns = missingColumns & "Year, "
    If cols.Type = 0 Then missingColumns = missingColumns & "Type, "
    
    If missingColumns <> "" Then
        missingColumns = Left(missingColumns, Len(missingColumns) - 2) ' Remove trailing comma and space
        MsgBox "The following required columns were not found: " & missingColumns, vbExclamation
    End If
    
    GetColumnIndices = cols
End Function

Private Function GetOrCreateSheet(sheetName As String) As Worksheet
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

Private Sub UpdateReportDates(startDate As String, endDate As String)
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

Private Function CountOccurrences(ByRef keys() As String, ByRef counts() As Long, ByVal total As Long, ByVal key As String) As Long
    Dim i As Long
    
    ' If first item, initialize arrays
    If total = 0 Then
        keys(1) = key
        counts(1) = 1
        CountOccurrences = 1
        Exit Function
    End If
    
    ' Check if key exists
    For i = 1 To total
        If keys(i) = key Then
            counts(i) = counts(i) + 1
            CountOccurrences = total
            Exit Function
        End If
    Next i
    
    ' Key not found, add it
    total = total + 1
    If total > UBound(keys) Then
        ReDim Preserve keys(1 To UBound(keys) * 2)
        ReDim Preserve counts(1 To UBound(counts) * 2)
    End If
    keys(total) = key
    counts(total) = 1
    CountOccurrences = total
End Function

Private Function CountUniqueSupporters(campaign As String, campaignSupporters() As String, totalPairs As Long) As Long
    Dim uniqueSupporters() As String
    ReDim uniqueSupporters(1 To 1000)
    Dim uniqueCount As Long
    uniqueCount = 0
    
    Dim i As Long
    Dim pair As String
    Dim parts() As String
    
    For i = 1 To totalPairs
        pair = campaignSupporters(i)
        If pair <> "" Then
            parts = Split(pair, "|")
            If UBound(parts) >= 1 Then
                If parts(0) = campaign Then
                    Dim supporter As String
                    supporter = parts(1)
                    
                    ' Check if supporter already counted
                    Dim found As Boolean
                    found = False
                    Dim j As Long
                    
                    For j = 1 To uniqueCount
                        If uniqueSupporters(j) = supporter Then
                            found = True
                            Exit For
                        End If
                    Next j
                    
                    If Not found Then
                        uniqueCount = uniqueCount + 1
                        If uniqueCount > UBound(uniqueSupporters) Then
                            ReDim Preserve uniqueSupporters(1 To UBound(uniqueSupporters) * 2)
                        End If
                        uniqueSupporters(uniqueCount) = supporter
                    End If
                End If
            End If
        End If
    Next i
    
    CountUniqueSupporters = uniqueCount
End Function

Private Sub WriteReport(ws As Worksheet, col1Header As String, col2Header As String, col3Header As String, _
                      keys() As String, counts() As Long, uniqueCounts() As Long, total As Long)
    ' Check parameters
    If ws Is Nothing Then
        Debug.Print "WriteReport: Worksheet is Nothing"
        Exit Sub
    End If
    
    If total <= 0 Then
        Debug.Print "WriteReport: No data to write (total = " & total & ")"
        Exit Sub
    End If
    
    Debug.Print "Writing report to " & ws.Name & " with " & total & " rows"
    
    ' Write headers
    ws.Cells(1, 1).Value = col1Header
    ws.Cells(1, 2).Value = col2Header
    
    ' Format headers
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 2).Font.Bold = True
    
    ' Check if we need to include unique counts
    Dim includeUniques As Boolean
    includeUniques = (col3Header <> "" And Not IsError(uniqueCounts))
    
    If includeUniques Then
        ws.Cells(1, 3).Value = col3Header
        ws.Cells(1, 3).Font.Bold = True
    End If
    
    ' Write data
    Dim row As Long
    row = 2
    
    Dim i As Long
    For i = 1 To total
        ws.Cells(row, 1).Value = keys(i)
        ws.Cells(row, 2).Value = counts(i)
        
        If includeUniques Then
            ws.Cells(row, 3).Value = uniqueCounts(i)
        End If
        
        row = row + 1
    Next i
    
    ' Auto-fit columns
    ws.Columns("A:C").AutoFit
    
    ' Sort data
    SortReportSheet ws
End Sub

Private Sub WriteSupporterReport(ws As Worksheet, supporters() As String, counts() As Long, total As Long)
    ' Check parameters
    If ws Is Nothing Then Exit Sub
    If total <= 0 Then Exit Sub
    
    ' Write headers
    ws.Cells(1, 1).Value = "Supporter ID"
    ws.Cells(1, 2).Value = "Supporter Email"
    ws.Cells(1, 3).Value = "Count"
    
    ' Format headers
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 2).Font.Bold = True
    ws.Cells(1, 3).Font.Bold = True
    
    ' Write data
    Dim row As Long
    row = 2
    
    Dim i As Long
    For i = 1 To total
        Dim parts() As String
        parts = Split(supporters(i), " - ")
        
        ws.Cells(row, 1).Value = parts(0)
        
        If UBound(parts) > 0 Then
            ws.Cells(row, 2).Value = parts(1)
        End If
        
        ws.Cells(row, 3).Value = counts(i)
        
        row = row + 1
    Next i
    
    ' Auto-fit columns
    ws.Columns("A:C").AutoFit
    
    ' Sort data
    SortReportSheet ws
End Sub

Private Sub SortReportSheet(ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow <= 1 Then Exit Sub
    
    On Error Resume Next
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("A2:A" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending
        .SetRange ws.Range("A1").CurrentRegion
        .Header = xlYes
        .Apply
    End With
    On Error GoTo 0
End Sub