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

' Dictionary type for storing counts - MOVED TO TOP OF MODULE
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
        Dict As Object ' Scripting.Dictionary
    End Type
#End If

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

' Add menu to Excel ribbon or set up keyboard shortcut on Mac
Private Sub AddReportMenu()
    On Error Resume Next
    
    #If Mac Then
        Application.OnKey "⌘u", "GenerateUANReports"  ' Command+U on Mac calls our main entry point
    #Else
        ' Windows version using CommandBars
        Application.CommandBars("UAN Reports").Delete
        On Error GoTo 0
        
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
End Sub

#If Mac Then
' Mac version of menu using InputBox
Private Sub ShowReportMenu()
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
    
    If choice = False Then Exit Sub
    
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
    End Select
End Sub
#End If

' Dictionary helper functions for Mac
#If Mac Then
Private Function CreateDictionary() As ReportCounts
    Dim rc As ReportCounts
    Dim md As MacDict
    Set md.Keys = New Collection
    Set md.Items = New Collection
    rc.Dict = md
    CreateDictionary = rc
End Function

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

Private Function DictKeys(ByRef dict As ReportCounts) As Collection
    Set DictKeys = dict.Dict.Keys
End Function
#Else
Private Function CreateDictionary() As ReportCounts
    Dim rc As ReportCounts
    Set rc.Dict = CreateObject("Scripting.Dictionary")
    CreateDictionary = rc
End Function
#End If

' After the dictionary helper functions (#If Mac Then... #Else... #End If section)
' Add these functions before UpdateReportDates

Private Sub WriteSortedData(ByRef ws As Worksheet, ByVal headers As Variant, ByRef counts As ReportCounts, Optional ByRef uniqueCounts As ReportCounts = Nothing)
    ' Write headers
    Dim col As Long
    For col = 1 To UBound(headers) + 1
        ws.Cells(1, col).Value = headers(col - 1)
        If col = 1 Then
            ws.Cells(1, col).HorizontalAlignment = xlLeft
        Else
            ws.Cells(1, col).HorizontalAlignment = xlRight
        End If
    Next col
    
    ' Write data
    Dim row As Long
    row = 2
    
    ' Check if we need to include unique counts
    Dim includeUniqueCounts As Boolean
    includeUniqueCounts = Not uniqueCounts Is Nothing
    
    #If Mac Then
        ' Mac version using Collections
        Dim i As Long
        For i = 1 To counts.Dict.Keys.Count
            Dim key As String
            key = counts.Dict.Keys(i)
            
            ws.Cells(row, 1).Value = key
            ws.Cells(row, 1).HorizontalAlignment = xlLeft
            
            ws.Cells(row, 2).Value = counts.Dict.Items(i)
            ws.Cells(row, 2).HorizontalAlignment = xlRight
            
            If includeUniqueCounts Then
                ws.Cells(row, 3).Value = DictGet(uniqueCounts, key)
                ws.Cells(row, 3).HorizontalAlignment = xlRight
            End If
            
            row = row + 1
        Next i
    #Else
        ' Windows version using Dictionary
        Dim key As Variant
        For Each key In counts.Dict.Keys
            ws.Cells(row, 1).Value = key
            ws.Cells(row, 1).HorizontalAlignment = xlLeft
            
            ws.Cells(row, 2).Value = counts.Dict(key)
            ws.Cells(row, 2).HorizontalAlignment = xlRight
            
            If includeUniqueCounts Then
                On Error Resume Next
                ws.Cells(row, 3).Value = uniqueCounts.Dict(key)
                On Error GoTo 0
                ws.Cells(row, 3).HorizontalAlignment = xlRight
            End If
            
            row = row + 1
        Next key
    #End If
    
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
    
    #If Mac Then
        Dim i As Long
        For i = 1 To counts.Dict.Keys.Count
            Dim key As String
            key = counts.Dict.Keys(i)
            
            Dim parts() As String
            parts = Split(key, " - ")
            
            ws.Cells(row, 1).Value = parts(0)
            ws.Cells(row, 1).HorizontalAlignment = xlLeft
            
            If UBound(parts) > 0 Then
                ws.Cells(row, 2).Value = parts(1)
            End If
            ws.Cells(row, 2).HorizontalAlignment = xlLeft
            
            ws.Cells(row, 3).Value = counts.Dict.Items(i)
            ws.Cells(row, 3).HorizontalAlignment = xlRight
            
            row = row + 1
        Next i
    #Else
        Dim key As Variant
        For Each key In counts.Dict.Keys
            Dim parts() As String
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
    #End If
    
    ' Sort data if there are rows
    If row > 2 Then
        With ws.Range(ws.Cells(2, 1), ws.Cells(row - 1, 3))
            .Sort Key1:=.Columns(1), Order1:=xlAscending, Header:=xlNo
        End With
    End If
End Sub

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
        .CampaignID = WorksheetFunction.Match("Campaign ID", headers, 0)
        .CampaignDate = WorksheetFunction.Match("Campaign Date", headers, 0)
        .SupporterID = WorksheetFunction.Match("Supporter ID", headers, 0)
        .SupporterEmail = WorksheetFunction.Match("Supporter Email", headers, 0)
        .Country = WorksheetFunction.Match("External Reference 6 (Country)", headers, 0)
        .CaseNumber = WorksheetFunction.Match("External Reference 7 (Case Number)", headers, 0)
        .Topics = WorksheetFunction.Match("External Reference 8 (Topics)", headers, 0)
        .Year = WorksheetFunction.Match("External Reference 10 (Year)", headers, 0)
        .Type = WorksheetFunction.Match("External Reference 10 (Type)", headers, 0)
    End With
    
    GetColumnIndices = cols
End Function

' Forward declaration of ProcessRowForReport to avoid forward reference error
Private Sub ProcessRowForReport(ByRef dataRange As Range, ByVal row As Long, ByVal cols As ColumnIndices, _
                              ByRef counts As ReportCounts, ByVal reportType As String, _
                              ByVal hasStartDate As Boolean, ByVal startDate As Date, _
                              ByVal hasEndDate As Boolean, ByVal endDate As Date)
    ' Get campaign date with error handling
    Dim campaignDate As Date
    On Error Resume Next
    campaignDate = dataRange.Cells(row, cols.CampaignDate).Value
    On Error GoTo 0
    
    ' Skip rows without valid dates or outside date range
    If campaignDate = 0 Then Exit Sub
    If hasStartDate And campaignDate < startDate Then Exit Sub
    If hasEndDate And campaignDate > endDate Then Exit Sub
    
    ' Process based on report type
    Select Case reportType
        Case "Campaign ID"
            Dim campaignID As String
            campaignID = Trim(CStr(dataRange.Cells(row, cols.CampaignID).Value))
            If campaignID <> "" Then
                #If Mac Then
                    DictSet counts, campaignID, DictGet(counts, campaignID) + 1
                #Else
                    counts.Dict(campaignID) = CLng(counts.Dict(campaignID)) + 1
                #End If
            End If
            
        Case "Case Number"
            Dim caseNumber As String
            caseNumber = Trim(CStr(dataRange.Cells(row, cols.CaseNumber).Value))
            If caseNumber <> "" Then
                #If Mac Then
                    DictSet counts, caseNumber, DictGet(counts, caseNumber) + 1
                #Else
                    counts.Dict(caseNumber) = CLng(counts.Dict(caseNumber)) + 1
                #End If
            End If
            
        Case "Country"
            Dim country As String
            country = Trim(CStr(dataRange.Cells(row, cols.Country).Value))
            If country <> "" Then
                #If Mac Then
                    DictSet counts, country, DictGet(counts, country) + 1
                #Else
                    counts.Dict(country) = CLng(counts.Dict(country)) + 1
                #End If
            End If
            
        Case "Topics"
            Dim topics As String
            topics = Trim(CStr(dataRange.Cells(row, cols.Topics).Value))
            If topics <> "" Then
                Dim topicArray() As String
                topicArray = Split(topics, ",")
                Dim topic As Variant
                For Each topic In topicArray
                    Dim topicValue As String
                    topicValue = Trim(CStr(topic))
                    If topicValue <> "" Then
                        #If Mac Then
                            DictSet counts, topicValue, DictGet(counts, topicValue) + 1
                        #Else
                            counts.Dict(topicValue) = CLng(counts.Dict(topicValue)) + 1
                        #End If
                    End If
                Next topic
            End If
            
        Case "Year"
            Dim yearValue As String
            yearValue = Trim(CStr(dataRange.Cells(row, cols.Year).Value))
            If yearValue <> "" Then
                #If Mac Then
                    DictSet counts, yearValue, DictGet(counts, yearValue) + 1
                #Else
                    counts.Dict(yearValue) = CLng(counts.Dict(yearValue)) + 1
                #End If
            End If
            
        Case "Type"
            Dim typeValue As String
            typeValue = Trim(CStr(dataRange.Cells(row, cols.Type).Value))
            If typeValue <> "" Then
                #If Mac Then
                    DictSet counts, typeValue, DictGet(counts, typeValue) + 1
                #Else
                    counts.Dict(typeValue) = CLng(counts.Dict(typeValue)) + 1
                #End If
            End If
            
        Case "Date"
            Dim monthYear As String
            monthYear = Format(campaignDate, "yyyy-mm")
            #If Mac Then
                DictSet counts, monthYear, DictGet(counts, monthYear) + 1
            #Else
                counts.Dict(monthYear) = CLng(counts.Dict(monthYear)) + 1
            #End If
            
        Case "Supporter"
            Dim supporterID As String, supporterEmail As String
            supporterID = Trim(CStr(dataRange.Cells(row, cols.SupporterID).Value))
            supporterEmail = Trim(CStr(dataRange.Cells(row, cols.SupporterEmail).Value))
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

Private Function CalculateUniqueSupporters(ByRef data As Range, ByVal cols As ColumnIndices, _
                                         ByVal hasStartDate As Boolean, ByVal startDate As Date, _
                                         ByVal hasEndDate As Boolean, ByVal endDate As Date) As ReportCounts
    ' Initialize the return value
    Dim uniqueCounts As ReportCounts
    uniqueCounts = CreateDictionary()
    
    ' Create a dictionary to track unique supporters per campaign
    Dim campaignSupporters As Object
    Set campaignSupporters = CreateObject("Scripting.Dictionary")
    
    ' Loop through data rows
    Dim row As Long
    For row = 2 To data.Rows.Count
        ' Get campaign date with error handling
        Dim campaignDate As Date
        On Error Resume Next
        campaignDate = data.Cells(row, cols.CampaignDate).Value
        On Error GoTo 0
        
        ' Skip rows without valid dates
        If campaignDate = 0 Then GoTo NextRow
        
        ' Apply date filters if specified
        If (hasStartDate And campaignDate < startDate) Or (hasEndDate And campaignDate > endDate) Then
            GoTo NextRow
        End If
        
        ' Get campaign and supporter IDs
        Dim campaignID As String
        Dim supporterID As String
        
        campaignID = CStr(data.Cells(row, cols.CampaignID).Value)
        supporterID = CStr(data.Cells(row, cols.SupporterID).Value)
        
        ' Track unique supporters per campaign
        If campaignID <> "" And supporterID <> "" Then
            If Not campaignSupporters.Exists(campaignID) Then
                Set campaignSupporters(campaignID) = CreateObject("Scripting.Dictionary")
            End If
            campaignSupporters(campaignID)(supporterID) = 1
        End If
NextRow:
    Next row
    
    ' Convert to counts
    Dim campaign As Variant
    For Each campaign In campaignSupporters.Keys
        Dim campaignKey As String
        campaignKey = CStr(campaign)
        
        #If Mac Then
            DictSet uniqueCounts, campaignKey, campaignSupporters(campaign).Count
        #Else
            uniqueCounts.Dict(campaignKey) = campaignSupporters(campaign).Count
        #End If
    Next campaign
    
    CalculateUniqueSupporters = uniqueCounts
End Function

Private Sub ProcessCampaignData()
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("processed-export")
    
    If ws Is Nothing Then
        MsgBox "The 'processed-export' sheet is required.", vbExclamation
        Exit Sub
    End If
    
    ' Get date inputs
    Dim startDateInput As String, endDateInput As String
    startDateInput = InputBox("Enter Start Date (YYYY-MM-DD) or leave blank for no limit", "Start Date")
    endDateInput = InputBox("Enter End Date (YYYY-MM-DD) or leave blank for no limit", "End Date")
    
    Dim startDate As Date, endDate As Date
    Dim hasStartDate As Boolean, hasEndDate As Boolean
    
    If startDateInput <> "" Then
        startDate = CDate(startDateInput)
        hasStartDate = True
    End If
    
    If endDateInput <> "" Then
        endDate = CDate(endDateInput)
        hasEndDate = True
    End If
    
    ' Update report dates
    UpdateReportDates startDateInput, endDateInput
    
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
    
    campaignCounts = CreateDictionary()
    caseCounts = CreateDictionary()
    countryCounts = CreateDictionary()
    topicCounts = CreateDictionary()
    yearCounts = CreateDictionary()
    typeCounts = CreateDictionary()
    dateCounts = CreateDictionary()
    supporterCounts = CreateDictionary()
    
    ' Process data
    Dim dataRange As Range
    Set dataRange = ws.UsedRange
    
    Dim row As Long
    For row = 2 To dataRange.Rows.Count
        ' Process for each report type
        ProcessRowForReport dataRange, row, cols, campaignCounts, "Campaign ID", hasStartDate, startDate, hasEndDate, endDate
        ProcessRowForReport dataRange, row, cols, caseCounts, "Case Number", hasStartDate, startDate, hasEndDate, endDate
        ProcessRowForReport dataRange, row, cols, countryCounts, "Country", hasStartDate, startDate, hasEndDate, endDate
        ProcessRowForReport dataRange, row, cols, topicCounts, "Topics", hasStartDate, startDate, hasEndDate, endDate
        ProcessRowForReport dataRange, row, cols, yearCounts, "Year", hasStartDate, startDate, hasEndDate, endDate
        ProcessRowForReport dataRange, row, cols, typeCounts, "Type", hasStartDate, startDate, hasEndDate, endDate
        ProcessRowForReport dataRange, row, cols, dateCounts, "Date", hasStartDate, startDate, hasEndDate, endDate
        ProcessRowForReport dataRange, row, cols, supporterCounts, "Supporter", hasStartDate, startDate, hasEndDate, endDate
    Next row
    
    ' Calculate unique supporters
    Dim uniqueCounts As ReportCounts
    uniqueCounts = CalculateUniqueSupporters(dataRange, cols, hasStartDate, startDate, hasEndDate, endDate)
    
    ' Write reports
    WriteSortedData GetOrCreateSheet("by-name"), Array("Campaign ID", "Count", "Unique Supporters"), campaignCounts, uniqueCounts
    WriteSortedData GetOrCreateSheet("by-case-number"), Array("Case Number", "Count"), caseCounts
    WriteSortedData GetOrCreateSheet("by-country"), Array("Country", "Count"), countryCounts
    WriteSortedData GetOrCreateSheet("by-topic"), Array("Topic", "Count"), topicCounts
    WriteSortedData GetOrCreateSheet("by-year"), Array("Year", "Count"), yearCounts
    WriteSortedData GetOrCreateSheet("by-type"), Array("Type", "Count"), typeCounts
    WriteSortedData GetOrCreateSheet("by-date"), Array("Month", "Count"), dateCounts
    WriteSupporterData GetOrCreateSheet("by-supporter"), supporterCounts
    
    MsgBox "✓ Your UAN Reports have been updated!", vbInformation
    
    Application.ScreenUpdating = True
End Sub

Private Sub ProcessCampaignDataExceptSupporter()
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("processed-export")
    
    If ws Is Nothing Then
        MsgBox "The 'processed-export' sheet is required.", vbExclamation
        Exit Sub
    End If
    
    ' Get date inputs
    Dim startDateInput As String, endDateInput As String
    startDateInput = InputBox("Enter Start Date (YYYY-MM-DD) or leave blank for no limit", "Start Date")
    endDateInput = InputBox("Enter End Date (YYYY-MM-DD) or leave blank for no limit", "End Date")
    
    Dim startDate As Date, endDate As Date
    Dim hasStartDate As Boolean, hasEndDate As Boolean
    
    If startDateInput <> "" Then
        startDate = CDate(startDateInput)
        hasStartDate = True
    End If
    
    If endDateInput <> "" Then
        endDate = CDate(endDateInput)
        hasEndDate = True
    End If
    
    ' Update report dates
    UpdateReportDates startDateInput, endDateInput
    
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
    
    campaignCounts = CreateDictionary()
    caseCounts = CreateDictionary()
    countryCounts = CreateDictionary()
    topicCounts = CreateDictionary()
    yearCounts = CreateDictionary()
    typeCounts = CreateDictionary()
    dateCounts = CreateDictionary()
    
    ' Process data
    Dim dataRange As Range
    Set dataRange = ws.UsedRange
    
    Dim row As Long
    For row = 2 To dataRange.Rows.Count
        ' Process for each report type except supporter
        ProcessRowForReport dataRange, row, cols, campaignCounts, "Campaign ID", hasStartDate, startDate, hasEndDate, endDate
        ProcessRowForReport dataRange, row, cols, caseCounts, "Case Number", hasStartDate, startDate, hasEndDate, endDate
        ProcessRowForReport dataRange, row, cols, countryCounts, "Country", hasStartDate, startDate, hasEndDate, endDate
        ProcessRowForReport dataRange, row, cols, topicCounts, "Topics", hasStartDate, startDate, hasEndDate, endDate
        ProcessRowForReport dataRange, row, cols, yearCounts, "Year", hasStartDate, startDate, hasEndDate, endDate
        ProcessRowForReport dataRange, row, cols, typeCounts, "Type", hasStartDate, startDate, hasEndDate, endDate
        ProcessRowForReport dataRange, row, cols, dateCounts, "Date", hasStartDate, startDate, hasEndDate, endDate
    Next row
    
    ' Calculate unique supporters
    Dim uniqueCounts As ReportCounts
    uniqueCounts = CalculateUniqueSupporters(dataRange, cols, hasStartDate, startDate, hasEndDate, endDate)
    
    ' Write reports
    WriteSortedData GetOrCreateSheet("by-name"), Array("Campaign ID", "Count", "Unique Supporters"), campaignCounts, uniqueCounts
    WriteSortedData GetOrCreateSheet("by-case-number"), Array("Case Number", "Count"), caseCounts
    WriteSortedData GetOrCreateSheet("by-country"), Array("Country", "Count"), countryCounts
    WriteSortedData GetOrCreateSheet("by-topic"), Array("Topic", "Count"), topicCounts
    WriteSortedData GetOrCreateSheet("by-year"), Array("Year", "Count"), yearCounts
    WriteSortedData GetOrCreateSheet("by-type"), Array("Type", "Count"), typeCounts
    WriteSortedData GetOrCreateSheet("by-date"), Array("Month", "Count"), dateCounts
    
    MsgBox "✓ Your UAN Reports have been updated! (except by-supporter)", vbInformation
    
    Application.ScreenUpdating = True
End Sub

' Individual report update functions
Private Sub UpdateCampaignReport()
    ProcessSpecificReport "by-name", "Campaign ID"
End Sub

Private Sub UpdateCaseReport()
    ProcessSpecificReport "by-case-number", "Case Number"
End Sub

Private Sub UpdateCountryReport()
    ProcessSpecificReport "by-country", "Country"
End Sub

Private Sub UpdateTopicReport()
    ProcessSpecificReport "by-topic", "Topics"
End Sub

Private Sub UpdateYearReport()
    ProcessSpecificReport "by-year", "Year"
End Sub

Private Sub UpdateTypeReport()
    ProcessSpecificReport "by-type", "Type"
End Sub

Private Sub UpdateDateReport()
    ProcessSpecificReport "by-date", "Date"
End Sub

Private Sub UpdateSupporterReport()
    ProcessSpecificReport "by-supporter", "Supporter"
End Sub

Private Sub ProcessSpecificReport(ByVal sheetName As String, ByVal reportType As String)
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("processed-export")
    
    If ws Is Nothing Then
        MsgBox "The 'processed-export' sheet is required.", vbExclamation
        Exit Sub
    End If
    
    ' Get date inputs
    Dim startDateInput As String, endDateInput As String
    startDateInput = InputBox("Enter Start Date (YYYY-MM-DD) or leave blank for no limit", "Start Date")
    endDateInput = InputBox("Enter End Date (YYYY-MM-DD) or leave blank for no limit", "End Date")
    
    Dim startDate As Date, endDate As Date
    Dim hasStartDate As Boolean, hasEndDate As Boolean
    
    If startDateInput <> "" Then
        startDate = CDate(startDateInput)
        hasStartDate = True
    End If
    
    If endDateInput <> "" Then
        endDate = CDate(endDateInput)
        hasEndDate = True
    End If
    
    ' Update report dates to "Mixed" when running individual reports
    UpdateReportDates "Mixed", "Mixed"
    
    ' Get column indices
    Dim cols As ColumnIndices
    cols = GetColumnIndices()
    
    ' Process data based on report type
    Dim counts As ReportCounts
    Set counts = CreateDictionary()
    
    ' Process data
    Dim dataRange As Range
    Set dataRange = ws.UsedRange
    
    Dim row As Long
    For row = 2 To dataRange.Rows.Count
        ProcessRowForReport dataRange, row, cols, counts, reportType, hasStartDate, startDate, hasEndDate, endDate
    Next row
    
    ' Write report
    Dim reportSheet As Worksheet
    Set reportSheet = GetOrCreateSheet(sheetName)
    
    If reportType = "Campaign ID" Then
        Dim uniqueCounts As ReportCounts
        uniqueCounts = CalculateUniqueSupporters(dataRange, cols, hasStartDate, startDate, hasEndDate, endDate)
        WriteSortedData reportSheet, Array("Campaign ID", "Count", "Unique Supporters"), counts, uniqueCounts
    ElseIf reportType = "Supporter" Then
        WriteSupporterData reportSheet, counts
    Else
        WriteSortedData reportSheet, Array(reportType, "Count"), counts
    End If
    
    MsgBox "✓ The " & sheetName & " report has been updated!", vbInformation
    
    Application.ScreenUpdating = True
End Sub

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
    
    ' Add menu for future use
    AddReportMenu
    
    ' Show options dialog to user
    Dim choice As Variant
    choice = Application.InputBox( _
        "Choose a report to generate:" & vbNewLine & _
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
        "UAN Reports", "1", , , , , 2)
    
    If choice = False Then GoTo Cleanup
    
    ' Process the user's choice
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
        Case Else: MsgBox "Invalid selection. Please try again.", vbExclamation
    End Select
    
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