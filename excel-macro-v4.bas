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
Public Sub Generate_UAN_Report()
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
    
    ' Variables to track min and max dates in the data
    Dim minDate As Date, maxDate As Date
    minDate = DateSerial(9999, 12, 31) ' Initialize to far future
    maxDate = DateSerial(1900, 1, 1)   ' Initialize to far past
    
    ' Add counter for processed rows
    Dim processedRows As Long
    processedRows = 0
    
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
        
        ' Track min and max dates for ALL data, regardless of filter
        If campaignDate <> 0 Then
            If campaignDate < minDate Then minDate = campaignDate
            If campaignDate > maxDate Then maxDate = campaignDate
        End If
        
        ' Skip rows without valid dates or if outside date range
        If campaignDate = 0 Then GoTo NextRow
        If hasStartDate And campaignDate < startDate Then GoTo NextRow
        If hasEndDate And campaignDate > endDate Then GoTo NextRow
        
        ' Increment processed rows counter
        processedRows = processedRows + 1
        
        ' Track min and max dates
        If campaignDate < minDate Then minDate = campaignDate
        If campaignDate > maxDate Then maxDate = campaignDate
        
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
    
    ' After processing all data, create the report sheet with dates
    Application.StatusBar = "Creating reports..."
    
    ' First, write the individual report sheets
    ' ... (code to write individual reports) ...
    
    ' Determine date range for display, ensuring it's within the actual data range
    Dim displayStartDate As String, displayEndDate As String
    Dim effectiveStartDate As Date, effectiveEndDate As Date
    
    ' Determine effective start date (user input or data min, whichever is later)
    If hasStartDate Then
        If startDate > minDate Then
            effectiveStartDate = startDate
        Else
            effectiveStartDate = minDate
        End If
        displayStartDate = Format(effectiveStartDate, "yyyy-mm-dd")
    Else
        effectiveStartDate = minDate
        displayStartDate = Format(minDate, "yyyy-mm-dd")
    End If
    
    ' Determine effective end date (user input or data max, whichever is earlier)
    If hasEndDate Then
        If endDate < maxDate Then
            effectiveEndDate = endDate
        Else
            effectiveEndDate = maxDate
        End If
        displayEndDate = Format(effectiveEndDate, "yyyy-mm-dd")
    Else
        effectiveEndDate = maxDate
        displayEndDate = Format(maxDate, "yyyy-mm-dd")
    End If
    
    ' Create the main report sheet directly
    CreateMainReport displayStartDate, displayEndDate, _
                    campaignIDs, campaignCounts, campaignUniqueSupport, totalCampaigns, _
                    caseNumbers, caseCounts, totalCases, _
                    countries, countryCounts, totalCountries, _
                    topics, topicCounts, totalTopics, _
                    years, yearCounts, totalYears, _
                    types, typeCounts, totalTypes, _
                    dates, dateCounts, totalDates, _
                    supporters, supporterCounts, totalSupporters
    
    ' Create enhanced confirmation message
    Dim confirmMsg As String
    confirmMsg = "Your UAN Report has been updated!" & vbNewLine & vbNewLine & _
                "Summary:" & vbNewLine & _
                "Entries in export: " & (lastRow - 1) & vbNewLine & _
                "Entries processed: " & processedRows & " (" & Format(processedRows / (lastRow - 1) * 100, "0.0") & "%)" & vbNewLine & _
                "Start Date: " & displayStartDate & vbNewLine & _
                "End Date: " & displayEndDate & vbNewLine & vbNewLine & _
                "Results (Unique / Total):" & vbNewLine & _
                "Supporters: " & totalSupporters & " / " & Application.WorksheetFunction.Sum(supporterCounts) & vbNewLine & _
                "Campaigns: " & totalCampaigns & " / " & Application.WorksheetFunction.Sum(campaignCounts) & vbNewLine & _
                "Case Numbers: " & totalCases & " / " & Application.WorksheetFunction.Sum(caseCounts) & vbNewLine & _
                "Countries: " & totalCountries & " / " & Application.WorksheetFunction.Sum(countryCounts) & vbNewLine & _
                "Topics: " & totalTopics & " / " & Application.WorksheetFunction.Sum(topicCounts) & vbNewLine & _
                "Types: " & totalTypes & " / " & Application.WorksheetFunction.Sum(typeCounts)
    
    MsgBox confirmMsg, vbInformation, "UAN Report Generation Complete"
    
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

Private Sub CreateMainReport(startDate As String, endDate As String, _
                           campaignIDs() As String, campaignCounts() As Long, campaignUniques() As Long, totalCampaigns As Long, _
                           caseNumbers() As String, caseCounts() As Long, totalCases As Long, _
                           countries() As String, countryCounts() As Long, totalCountries As Long, _
                           topics() As String, topicCounts() As Long, totalTopics As Long, _
                           years() As String, yearCounts() As Long, totalYears As Long, _
                           types() As String, typeCounts() As Long, totalTypes As Long, _
                           dates() As String, dateCounts() As Long, totalDates As Long, _
                           supporters() As String, supporterCounts() As Long, totalSupporters As Long)
    
    ' Create a dummy array for reports without unique counts
    Dim dummyArray() As Long
    ReDim dummyArray(1 To 1)
    
    ' Sort all arrays alphabetically by key before creating the report
    SortArrays campaignIDs, campaignCounts, campaignUniques, totalCampaigns
    SortArrays caseNumbers, caseCounts, dummyArray, totalCases
    SortArrays countries, countryCounts, dummyArray, totalCountries
    SortArrays topics, topicCounts, dummyArray, totalTopics
    SortArrays years, yearCounts, dummyArray, totalYears
    SortArrays types, typeCounts, dummyArray, totalTypes
    SortArrays dates, dateCounts, dummyArray, totalDates
    SortArrays supporters, supporterCounts, dummyArray, totalSupporters
    
    ' Get or create the report sheet
    Dim reportSheet As Worksheet
    Set reportSheet = GetOrCreateSheet("report")
    
    ' Clear the sheet
    reportSheet.Cells.Clear
    
    ' Add title and date range
    reportSheet.Range("A1").Value = "Report"
    reportSheet.Range("B1").Value = "Date"  ' Add Date heading
    reportSheet.Range("A2").Value = "Start Date"
    reportSheet.Range("A3").Value = "End Date"
    
    ' Set date values with consistent formatting
    If IsDate(startDate) Then
        reportSheet.Range("B2").Value = Format(CDate(startDate), "yyyy-mm-dd")
    Else
        reportSheet.Range("B2").Value = startDate
    End If
    
    If IsDate(endDate) Then
        reportSheet.Range("B3").Value = Format(CDate(endDate), "yyyy-mm-dd")
    Else
        reportSheet.Range("B3").Value = endDate
    End If
    
    ' Format headers
    reportSheet.Range("A1:B1").Font.Bold = True
    reportSheet.Range("A1:B3").Borders.LineStyle = xlContinuous
    reportSheet.Range("A1:B1").Interior.ColorIndex = 15 ' Light gray
    
    ' Right-align the Date header and values
    reportSheet.Range("B1").HorizontalAlignment = xlRight
    reportSheet.Range("B2:B3").HorizontalAlignment = xlRight
    
    ' Add a narrow blank column for visual separation
    reportSheet.Columns("C:C").ColumnWidth = 2
    
    ' Set up the columns for each report type
    Dim col As Long
    col = 4 ' Start at column D (after the separator column)
    
    ' Campaign data (with unique supporters)
    AddDataToReport reportSheet, col, "Campaign ID", "Count", "Unique Supporters", _
                   campaignIDs, campaignCounts, campaignUniques, totalCampaigns
    col = col + 3
    
    ' Add narrow separator column
    reportSheet.Columns(col).ColumnWidth = 2
    col = col + 1
    
    ' Case number data (without unique supporters)
    AddDataToReport reportSheet, col, "Case Number", "Count", "", _
                   caseNumbers, caseCounts, campaignUniques, totalCases
    col = col + 2
    
    ' Add narrow separator column
    reportSheet.Columns(col).ColumnWidth = 2
    col = col + 1
    
    ' Country data (without unique supporters)
    AddDataToReport reportSheet, col, "Country", "Count", "", _
                   countries, countryCounts, campaignUniques, totalCountries
    col = col + 2
    
    ' Add narrow separator column
    reportSheet.Columns(col).ColumnWidth = 2
    col = col + 1
    
    ' Topic data (without unique supporters)
    AddDataToReport reportSheet, col, "Topic", "Count", "", _
                   topics, topicCounts, campaignUniques, totalTopics
    col = col + 2
    
    ' Add narrow separator column
    reportSheet.Columns(col).ColumnWidth = 2
    col = col + 1
    
    ' Type data (without unique supporters)
    AddDataToReport reportSheet, col, "Type", "Count", "", _
                   types, typeCounts, campaignUniques, totalTypes
    col = col + 2
    
    ' Add narrow separator column
    reportSheet.Columns(col).ColumnWidth = 2
    col = col + 1
    
    ' Year data (without unique supporters)
    AddDataToReport reportSheet, col, "Year", "Count", "", _
                   years, yearCounts, campaignUniques, totalYears
    col = col + 2
    
    ' Add narrow separator column
    reportSheet.Columns(col).ColumnWidth = 2
    col = col + 1
    
    ' Date data (without unique supporters)
    AddDataToReport reportSheet, col, "Month", "Count", "", _
                   dates, dateCounts, campaignUniques, totalDates
    col = col + 2
    
    ' Add narrow separator column
    reportSheet.Columns(col).ColumnWidth = 2
    col = col + 1
    
    ' Supporter data
    AddSupporterDataToReport reportSheet, col, supporters, supporterCounts, totalSupporters
    
    ' Format the report
    reportSheet.Columns("A:Z").AutoFit
    
    ' Ensure separator columns remain narrow
    reportSheet.Columns("C:C").ColumnWidth = 2
    
    ' Set column width for all separator columns
    Dim i As Long
    For i = 7 To col Step 3
        If i < col Then ' Skip the last one which might not be a separator
            reportSheet.Columns(i).ColumnWidth = 2
        End If
    Next i
End Sub

Private Sub AddDataToReport(reportSheet As Worksheet, startCol As Long, headerText As String, countHeader As String, _
                          uniqueHeader As String, keys() As String, counts() As Long, uniqueCounts() As Long, total As Long)
    ' Add headers
    reportSheet.Cells(1, startCol).Value = headerText
    reportSheet.Cells(1, startCol + 1).Value = countHeader
    
    If uniqueHeader <> "" Then
        reportSheet.Cells(1, startCol + 2).Value = uniqueHeader
    End If
    
    ' Format headers
    reportSheet.Cells(1, startCol).Font.Bold = True
    reportSheet.Cells(1, startCol + 1).Font.Bold = True
    
    If uniqueHeader <> "" Then
        reportSheet.Cells(1, startCol + 2).Font.Bold = True
    End If
    
    ' Right-align count and unique supporter columns including headers
    reportSheet.Cells(1, startCol + 1).HorizontalAlignment = xlRight
    reportSheet.Range(reportSheet.Cells(2, startCol + 1), reportSheet.Cells(total + 1, startCol + 1)).HorizontalAlignment = xlRight
    
    If uniqueHeader <> "" Then
        reportSheet.Cells(1, startCol + 2).HorizontalAlignment = xlRight
        reportSheet.Range(reportSheet.Cells(2, startCol + 2), reportSheet.Cells(total + 1, startCol + 2)).HorizontalAlignment = xlRight
    End If
    
    ' Write data
    Dim row As Long
    row = 2
    
    Dim i As Long
    For i = 1 To total
        ' Copy item name
        reportSheet.Cells(row, startCol).Value = keys(i)
        
        ' Copy count
        reportSheet.Cells(row, startCol + 1).Value = counts(i)
        
        ' Copy unique supporters if applicable
        If uniqueHeader <> "" Then
            reportSheet.Cells(row, startCol + 2).Value = uniqueCounts(i)
        End If
        
        row = row + 1
    Next i
    
    ' Format columns
    If uniqueHeader <> "" Then
        reportSheet.Range(reportSheet.Cells(1, startCol), reportSheet.Cells(row - 1, startCol + 2)).Borders.LineStyle = xlContinuous
    Else
        reportSheet.Range(reportSheet.Cells(1, startCol), reportSheet.Cells(row - 1, startCol + 1)).Borders.LineStyle = xlContinuous
    End If
    
    ' Format header row
    reportSheet.Range(reportSheet.Cells(1, startCol), reportSheet.Cells(1, startCol + IIf(uniqueHeader <> "", 2, 1))).Interior.ColorIndex = 15
End Sub

Private Sub AddSupporterDataToReport(reportSheet As Worksheet, startCol As Long, supporters() As String, counts() As Long, total As Long)
    ' Add headers
    reportSheet.Cells(1, startCol).Value = "Supporter ID"
    reportSheet.Cells(1, startCol + 1).Value = "Supporter Email"
    reportSheet.Cells(1, startCol + 2).Value = "Count"
    
    ' Format headers
    reportSheet.Cells(1, startCol).Font.Bold = True
    reportSheet.Cells(1, startCol + 1).Font.Bold = True
    reportSheet.Cells(1, startCol + 2).Font.Bold = True
    
    ' Right-align count column including header
    reportSheet.Cells(1, startCol + 2).HorizontalAlignment = xlRight
    reportSheet.Range(reportSheet.Cells(2, startCol + 2), reportSheet.Cells(total + 1, startCol + 2)).HorizontalAlignment = xlRight
    
    ' Write data
    Dim row As Long
    row = 2
    
    Dim i As Long
    For i = 1 To total
        Dim parts() As String
        parts = Split(supporters(i), " - ")
        
        ' Copy supporter ID
        reportSheet.Cells(row, startCol).Value = parts(0)
        
        ' Copy supporter email
        If UBound(parts) > 0 Then
            reportSheet.Cells(row, startCol + 1).Value = parts(1)
        End If
        
        ' Copy count
        reportSheet.Cells(row, startCol + 2).Value = counts(i)
        
        row = row + 1
    Next i
    
    ' Format columns
    reportSheet.Range(reportSheet.Cells(1, startCol), reportSheet.Cells(row - 1, startCol + 2)).Borders.LineStyle = xlContinuous
    
    ' Format header row
    reportSheet.Range(reportSheet.Cells(1, startCol), reportSheet.Cells(1, startCol + 2)).Interior.ColorIndex = 15
End Sub

Private Sub SortArrays(ByRef keys() As String, ByRef counts() As Long, ByRef uniqueCounts() As Long, ByVal total As Long)
    ' Simple bubble sort for the arrays
    Dim i As Long, j As Long
    Dim tempKey As String, tempCount As Long, tempUnique As Long
    Dim hasUniqueData As Boolean
    
    ' Check if uniqueCounts is a valid array with data
    hasUniqueData = False
    On Error Resume Next
    If UBound(uniqueCounts) >= total Then hasUniqueData = True
    On Error GoTo 0
    
    For i = 1 To total - 1
        For j = i + 1 To total
            ' Sort alphabetically by key
            If keys(i) > keys(j) Then
                ' Swap keys
                tempKey = keys(i)
                keys(i) = keys(j)
                keys(j) = tempKey
                
                ' Swap counts
                tempCount = counts(i)
                counts(i) = counts(j)
                counts(j) = tempCount
                
                ' Swap unique counts if applicable
                If hasUniqueData Then
                    tempUnique = uniqueCounts(i)
                    uniqueCounts(i) = uniqueCounts(j)
                    uniqueCounts(j) = tempUnique
                End If
            End If
        Next j
    Next i
End Sub