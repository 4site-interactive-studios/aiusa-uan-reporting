' UAN (Urgent Action Network) Reports Generator
'
' DESCRIPTION:
' This script processes campaign data from an Excel sheet to generate various reports
' tracking supporter engagement across different dimensions like country, case number,
' topics etc. All data is expected to be in a sheet named "processed-export".
' The "export" and "processed-export" sheets can have over 100 columns and are expected to be in the same workbook.

' SHEET REQUIREMENTS:
' - "export" sheet with headers:
'   * Campaign Data 33 (URL with Page Title Argument)
'
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
'   * Campaign Data 33 (URL with Page Title Argument)


' REPORTS GENERATED:
' - by-name: Campaign ID counts with unique supporters
' - by-page-title: Page Title count and unique count
' - by-case-number: Case number engagement
' - by-country: Country-wise participation
' - by-topic: Topic-wise breakdown
' - by-year: Year-wise analysis
' - by-type: Type-based categorization
' - by-date: Monthly trends
' - by-supporter: Individual supporter engagement


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
    CampaignData33 As Long
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
    Set wsExport = ThisWorkbook.Sheets("export")
    
    If wsExport Is Nothing Then
        MsgBox "Could not find a sheet named 'export'.", vbExclamation
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
               "Type: " & cols.Type & vbNewLine & _
               "CampaignData33: " & cols.CampaignData33
    
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
    
    ' Add arrays for page titles
    Dim pageTitles() As String
    Dim pageTitleCounts() As Long
    Dim pageTitleUniqueSupport() As Long
    Dim totalPageTitles As Long
    
    totalPageTitles = 0
    
    ' Initialize the arrays
    ReDim pageTitles(1 To 1000)
    ReDim pageTitleCounts(1 To 1000)
    ReDim pageTitleUniqueSupport(1 To 1000)
    
    ' For tracking unique supporters per campaign
    Dim campaignSupporters() As String
    ReDim campaignSupporters(1 To 10000)
    Dim campaignSupporterCount As Long
    campaignSupporterCount = 0
    
    ' For tracking unique supporters per page title
    Dim pageTitleSupporters() As String
    ReDim pageTitleSupporters(1 To 10000)
    Dim pageTitleSupporterCount As Long
    pageTitleSupporterCount = 0
    
    ' Process data
    Application.StatusBar = "Processing data: 0%"
    Dim row As Long, i As Long
    Dim supporterID As String, campaignID As String, caseNumber As String
    Dim country As String, topic As String, typeValue As String, yearValue As String
    Dim campaignDate As Date, monthYear As String
    Dim pageTitle As String, campaignData33 As String
    
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
        
        ' Extract Country (removing "Country: " prefix)
        Dim rawCountry As String
        rawCountry = Trim(CStr(ws.Cells(row, cols.Country).Value))
        
        ' Try different variations of the Country prefix
        If InStr(1, rawCountry, "Country: ", vbTextCompare) = 1 Then
            country = Trim(Mid(rawCountry, 9)) ' Remove "Country: " prefix
        ElseIf InStr(1, rawCountry, "Country :", vbTextCompare) = 1 Then
            country = Trim(Mid(rawCountry, 10)) ' Remove "Country :" prefix
        ElseIf InStr(1, rawCountry, "Country:", vbTextCompare) = 1 Then
            country = Trim(Mid(rawCountry, 8)) ' Remove "Country:" prefix
        Else
            country = rawCountry
        End If
        
        ' If only a colon remains, treat as empty
        If country = ":" Then country = ""
        
        ' Extract Case Number (removing "CaseNumber: " prefix)
        Dim rawCaseNumber As String
        rawCaseNumber = Trim(CStr(ws.Cells(row, cols.CaseNumber).Value))
        If InStr(1, rawCaseNumber, "CaseNumber: ", vbTextCompare) = 1 Then
            caseNumber = Trim(Mid(rawCaseNumber, 12)) ' Remove "CaseNumber: " prefix
        Else
            caseNumber = rawCaseNumber
        End If
        
        ' Extract Topics (removing "Topic: " prefix)
        Dim rawTopic As String
        rawTopic = Trim(CStr(ws.Cells(row, cols.Topics).Value))
        
        ' Try different variations of the Topic prefix
        If InStr(1, rawTopic, "Topic: ", vbTextCompare) = 1 Then
            topic = Trim(Mid(rawTopic, 7)) ' Remove "Topic: " prefix
        ElseIf InStr(1, rawTopic, "Topic :", vbTextCompare) = 1 Then
            topic = Trim(Mid(rawTopic, 8)) ' Remove "Topic :" prefix
        ElseIf InStr(1, rawTopic, "Topic:", vbTextCompare) = 1 Then
            topic = Trim(Mid(rawTopic, 6)) ' Remove "Topic:" prefix
        Else
            topic = rawTopic
        End If
        
        ' If only a colon remains, treat as empty
        If topic = ":" Then topic = ""
        
        ' Extract Year and Type from External Reference 10
        Dim rawYearType As String
        rawYearType = Trim(CStr(ws.Cells(row, cols.Year).Value))
        
        ' Default empty values
        yearValue = ""
        typeValue = ""
        
        If InStr(1, rawYearType, "YearType: ", vbTextCompare) = 1 Then
            ' Remove "YearType: " prefix
            Dim yearTypeContent As String
            yearTypeContent = Trim(Mid(rawYearType, 10))
            
            ' Check if there's a comma (separating year and type)
            Dim commaPos As Long
            commaPos = InStr(1, yearTypeContent, ",")
            
            If commaPos > 0 Then
                ' Extract year (before comma)
                Dim possibleYear As String
                possibleYear = Trim(Left(yearTypeContent, commaPos - 1))
                
                ' Check if it's a 4-digit year
                If Len(possibleYear) = 4 And IsNumeric(possibleYear) Then
                    yearValue = possibleYear
                    ' Type is everything after the comma
                    typeValue = Trim(Mid(yearTypeContent, commaPos + 1))
                Else
                    ' If not a year, it's all type
                    typeValue = yearTypeContent
                End If
            Else
                ' No comma - check if it's a 4-digit year
                If Len(yearTypeContent) = 4 And IsNumeric(yearTypeContent) Then
                    yearValue = yearTypeContent
                Else
                    ' If not a year, it's type
                    typeValue = yearTypeContent
                End If
            End If
        Else
            ' No prefix, use raw value
            If Len(rawYearType) = 4 And IsNumeric(rawYearType) Then
                yearValue = rawYearType
            Else
                typeValue = rawYearType
            End If
        End If
        
        ' Get Campaign Data 33 value and extract page title
        If cols.CampaignData33 > 0 Then
            campaignData33 = Trim(CStr(ws.Cells(row, cols.CampaignData33).Value))
            pageTitle = ExtractPageTitle(campaignData33)
            
            ' Page title counts
            If pageTitle <> "" Then
                totalPageTitles = CountOccurrences(pageTitles, pageTitleCounts, totalPageTitles, pageTitle)
                
                ' Track page title-supporter pairs for unique counts
                If supporterID <> "" Then
                    pageTitleSupporterCount = pageTitleSupporterCount + 1
                    If pageTitleSupporterCount > UBound(pageTitleSupporters) Then
                        ReDim Preserve pageTitleSupporters(1 To UBound(pageTitleSupporters) * 2)
                    End If
                    pageTitleSupporters(pageTitleSupporterCount) = pageTitle & "|" & supporterID
                End If
            End If
        End If
        
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
    
    ' Calculate unique supporters per page title
    For i = 1 To totalPageTitles
        pageTitleUniqueSupport(i) = CountUniqueSupporters(pageTitles(i), pageTitleSupporters, pageTitleSupporterCount)
    Next i
    
    ' After processing all data, create the report sheet with dates
    Application.StatusBar = "Creating reports..."
    
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
                    pageTitles, pageTitleCounts, pageTitleUniqueSupport, totalPageTitles, _
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
                "Page Titles: " & totalPageTitles & " / " & Application.WorksheetFunction.Sum(pageTitleCounts) & vbNewLine & _
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
    Debug.Print "Total Page Titles: " & totalPageTitles
End Sub

Private Function GetColumnIndices(ws As Worksheet) As ColumnIndices
    Dim cols As ColumnIndices
    
    ' Print all column headers for debugging
    Dim i As Long
    For i = 1 To 50 ' Check first 50 columns
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
    
    ' Find the External Reference columns
    cols.Country = WorksheetFunction.Match("External Reference 6", ws.Rows(1), 0)
    cols.CaseNumber = WorksheetFunction.Match("External Reference 7", ws.Rows(1), 0)
    cols.Topics = WorksheetFunction.Match("External Reference 8", ws.Rows(1), 0)
    cols.Year = WorksheetFunction.Match("External Reference 10", ws.Rows(1), 0)
    cols.Type = cols.Year ' Both Year and Type are in External Reference 10
    
    cols.CampaignData33 = WorksheetFunction.Match("Campaign Data 33", ws.Rows(1), 0)
    If cols.CampaignData33 = 0 Then cols.CampaignData33 = WorksheetFunction.Match("*Campaign*Data*33*", ws.Rows(1), 0)
    On Error GoTo 0
    
    ' Validate that all required columns were found
    Dim missingColumns As String
    missingColumns = ""
    
    If cols.CampaignID = 0 Then missingColumns = missingColumns & "Campaign ID, "
    If cols.CampaignDate = 0 Then missingColumns = missingColumns & "Campaign Date, "
    If cols.SupporterID = 0 Then missingColumns = missingColumns & "Supporter ID, "
    If cols.SupporterEmail = 0 Then missingColumns = missingColumns & "Supporter Email, "
    If cols.Country = 0 Then missingColumns = missingColumns & "External Reference 6, "
    If cols.CaseNumber = 0 Then missingColumns = missingColumns & "External Reference 7, "
    If cols.Topics = 0 Then missingColumns = missingColumns & "External Reference 8, "
    If cols.Year = 0 Then missingColumns = missingColumns & "External Reference 10, "
    ' Campaign Data 33 is not required, so we don't check for it
    
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
                           pageTitles() As String, pageTitleCounts() As Long, pageTitleUniques() As Long, totalPageTitles As Long, _
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
    
    ' Page title data (with unique supporters)
    AddDataToReport reportSheet, col, "Page Title", "Count", "Unique Supporters", _
                   pageTitles, pageTitleCounts, pageTitleUniques, totalPageTitles
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
    ' First, auto-fit all columns to ensure content is properly sized
    reportSheet.Columns("A:Z").AutoFit
    
    ' Then explicitly set the width for separator columns
    reportSheet.Columns("C:C").ColumnWidth = 2
    
    ' Set column width for all separator columns
    Dim j As Long
    For j = 7 To col Step 3
        If j < col Then ' Skip the last one which might not be a separator
            reportSheet.Columns(j).ColumnWidth = 2
        End If
    Next j
    
    ' Now auto-fit all columns that have headers (content columns)
    Dim contentCol As Long
    For contentCol = 1 To col
        ' If the column has a header, it's a content column and should be auto-sized
        If Trim(reportSheet.Cells(1, contentCol).Value) <> "" Then
            reportSheet.Columns(contentCol).AutoFit
        End If
    Next contentCol
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

' Function to extract and decode page title from URL
Private Function ExtractPageTitle(urlString As String) As String
    ' Default return value
    ExtractPageTitle = ""
    
    ' Check if the URL string is empty
    If urlString = "" Then Exit Function
    
    ' Look for page_title parameter
    Dim startPos As Long
    startPos = InStr(1, urlString, "page_title=")
    
    If startPos > 0 Then
        ' Move to the start of the value
        startPos = startPos + 11 ' Length of "page_title="
        
        ' Find the end of the value (either & or end of string)
        Dim endPos As Long
        endPos = InStr(startPos, urlString, "&")
        
        ' If no & found, use the end of the string
        If endPos = 0 Then endPos = Len(urlString) + 1
        
        ' Extract the encoded page title
        Dim encodedTitle As String
        encodedTitle = Mid(urlString, startPos, endPos - startPos)
        
        ' Decode the URL-encoded title, strip unwanted characters, and apply title casing
        Dim decodedTitle As String
        decodedTitle = URLDecode(encodedTitle)
        decodedTitle = Replace(decodedTitle, ",ÄôS", "'s")
        decodedTitle = Replace(decodedTitle, "Äô", "'")
        decodedTitle = Replace(decodedTitle, "¬", "")
        decodedTitle = Replace(decodedTitle, "†", "")
        decodedTitle = Replace(decodedTitle, "‚äôs", "'s")
        decodedTitle = Replace(decodedTitle, "‚'", "'")
        decodedTitle = Replace(decodedTitle, "‚", "'")  ' Replace just the single character
        decodedTitle = Replace(decodedTitle, "&#8217;", "'")  ' Replace HTML entity for right single quote
        
        ExtractPageTitle = SmartTitleCase(decodedTitle)
    End If
End Function

' Function to apply smart title casing to text
Private Function SmartTitleCase(ByVal Txt As String) As String
    Dim words As Variant
    Dim i As Integer
    Dim result As String
    
    ' Convert the text to lowercase, then split into words
    words = Split(LCase(Txt), " ")
    
    ' List of words to keep lowercase (articles, conjunctions, prepositions)
    Dim lowerWords As String
    lowerWords = "|a|an|and|as|at|but|by|for|if|in|nor|of|on|or|so|the|to|up|with|"
    
    ' Capitalize each word unless it's in the exception list
    For i = LBound(words) To UBound(words)
        ' Check if the word is in our lowercase list
        ' First word is always capitalized regardless
        If i = 0 Or InStr(1, lowerWords, "|" & LCase(words(i)) & "|") = 0 Then
            ' Capitalize first letter
            If Len(words(i)) > 0 Then
                words(i) = UCase(Left(words(i), 1)) & Mid(words(i), 2)
            End If
        End If
    Next i
    
    ' Rejoin words into a single string
    SmartTitleCase = Join(words, " ")
End Function

' Function to decode URL-encoded strings
Private Function URLDecode(encodedString As String) As String
    Dim result As String
    result = encodedString
    
    ' Replace + with space
    result = Replace(result, "+", " ")
    
    ' Replace %xx hex codes with their character equivalents
    Dim i As Long, hexCode As String, charCode As Long
    i = 1
    
    Do While i <= Len(result)
        If Mid(result, i, 1) = "%" And i + 2 <= Len(result) Then
            hexCode = Mid(result, i + 1, 2)
            
            ' Try to convert hex to decimal
            On Error Resume Next
            charCode = CLng("&H" & hexCode)
            
            If Err.Number = 0 Then
                ' Replace the %xx with the character
                result = Left(result, i - 1) & Chr(charCode) & Mid(result, i + 3)
            Else
                ' If conversion failed, just move on
                i = i + 1
            End If
            On Error GoTo 0
        Else
            i = i + 1
        End If
    Loop
    
    URLDecode = result
End Function

' Clear all UAN report data
Public Sub ZZZ_Clear_UAN_Reports_ZZZ()
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
    Application.StatusBar = "Clearing UAN Reports..."
    
    On Error GoTo ErrorHandler
    
    ' List of report sheets to clear
    Dim reportSheets As Variant
    reportSheets = Array("report")  ' Removed "by-page-title"
    
    ' Clear each report sheet
    Dim i As Long
    Dim ws As Worksheet
    Dim sheetsCleared As Long
    
    sheetsCleared = 0
    
    For i = LBound(reportSheets) To UBound(reportSheets)
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(reportSheets(i))
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            ws.Cells.Clear
            sheetsCleared = sheetsCleared + 1
        End If
    Next i
    
    ' Show confirmation message
    MsgBox "UAN Reports cleared successfully." & vbNewLine & _
           sheetsCleared & " report sheets were cleared.", _
           vbInformation, "UAN Reports Cleared"
    
    GoTo Cleanup
    
ErrorHandler:
    MsgBox "An error occurred while clearing reports: " & Err.Description, vbCritical
    
Cleanup:
    ' Restore Excel settings
    Application.StatusBar = False
    Application.ScreenUpdating = screenUpdating
    Application.DisplayStatusBar = statusBar
    Application.Calculation = calculation
    Application.DisplayAlerts = displayAlerts
End Sub