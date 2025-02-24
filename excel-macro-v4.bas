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

Option Explicit

' Dictionary type for storing counts
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
Public Sub AddReportMenu()
    On Error Resume Next
    
    #If Mac Then
        Application.OnKey "⌘u", "ShowReportMenu"  ' Command+U on Mac
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
Public Sub ShowReportMenu()
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
        "UAN Reports", , , , , , 1)
    
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
    Set rc.Dict.Keys = New Collection
    Set rc.Dict.Items = New Collection
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

Public Sub ProcessCampaignData()
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
    
    Set campaignCounts = CreateDictionary()
    Set caseCounts = CreateDictionary()
    Set countryCounts = CreateDictionary()
    Set topicCounts = CreateDictionary()
    Set yearCounts = CreateDictionary()
    Set typeCounts = CreateDictionary()
    Set dateCounts = CreateDictionary()
    Set supporterCounts = CreateDictionary()
    
    ' Process data
    Dim dataRange As Range
    Set dataRange = ws.UsedRange
    
    Dim row As Long
    For row = 2 To dataRange.Rows.Count
        Dim campaignDate As Date
        On Error Resume Next
        campaignDate = dataRange.Cells(row, cols.CampaignDate).Value
        On Error GoTo 0
        
        If campaignDate = 0 Then GoTo NextRow
        
        If (hasStartDate And campaignDate < startDate) Or (hasEndDate And campaignDate > endDate) Then
            GoTo NextRow
        End If
        
        ' Campaign counts
        Dim campaignID As String
        campaignID = Trim(CStr(dataRange.Cells(row, cols.CampaignID).Value))
        If campaignID <> "" Then
            #If Mac Then
                DictSet campaignCounts, campaignID, DictGet(campaignCounts, campaignID) + 1
            #Else
                campaignCounts.Dict(campaignID) = CLng(campaignCounts.Dict(campaignID)) + 1
            #End If
        End If
        
        ' Supporter counts
        Dim supporterID As String, supporterEmail As String
        supporterID = Trim(CStr(dataRange.Cells(row, cols.SupporterID).Value))
        supporterEmail = Trim(CStr(dataRange.Cells(row, cols.SupporterEmail).Value))
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
        cellValue = Trim(CStr(dataRange.Cells(row, cols.Country).Value))
        If cellValue <> "" Then
            #If Mac Then
                DictSet countryCounts, cellValue, DictGet(countryCounts, cellValue) + 1
            #Else
                countryCounts.Dict(cellValue) = CLng(countryCounts.Dict(cellValue)) + 1
            #End If
        End If
        
        ' Case number counts
        cellValue = Trim(CStr(dataRange.Cells(row, cols.CaseNumber).Value))
        If cellValue <> "" Then
            #If Mac Then
                DictSet caseCounts, cellValue, DictGet(caseCounts, cellValue) + 1
            #Else
                caseCounts.Dict(cellValue) = CLng(caseCounts.Dict(cellValue)) + 1
            #End If
        End If
        
        ' Year counts
        cellValue = Trim(CStr(dataRange.Cells(row, cols.Year).Value))
        If cellValue <> "" Then
            #If Mac Then
                DictSet yearCounts, cellValue, DictGet(yearCounts, cellValue) + 1
            #Else
                yearCounts.Dict(cellValue) = CLng(yearCounts.Dict(cellValue)) + 1
            #End If
        End If
        
        ' Type counts
        cellValue = Trim(CStr(dataRange.Cells(row, cols.Type).Value))
        If cellValue <> "" Then
            #If Mac Then
                DictSet typeCounts, cellValue, DictGet(typeCounts, cellValue) + 1
            #Else
                typeCounts.Dict(cellValue) = CLng(typeCounts.Dict(cellValue)) + 1
            #End If
        End If
        
        ' Topic counts
        cellValue = Trim(CStr(dataRange.Cells(row, cols.Topics).Value))
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
    Next row
    
    ' Calculate unique supporters
    Dim uniqueCounts As ReportCounts
    Set uniqueCounts = CalculateUniqueSupporters(dataRange, cols, hasStartDate, startDate, hasEndDate, endDate)
    
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

Private Function CalculateUniqueSupporters(ByRef data As Range, ByVal cols As ColumnIndices, _
                                         ByVal hasStartDate As Boolean, ByVal startDate As Date, _
                                         ByVal hasEndDate As Boolean, ByVal endDate As Date) As ReportCounts
    #If Mac Then
        Dim uniqueCounts As ReportCounts
        Set uniqueCounts.Dict.Keys = New Collection
        Set uniqueCounts.Dict.Items = New Collection
        
        Dim campaignSupporters As Object
        Set campaignSupporters = CreateObject("Scripting.Dictionary")
        
        Dim row As Long
        For row = 2 To data.Rows.Count
            Dim campaignDate As Date
            On Error Resume Next
            campaignDate = data.Cells(row, cols.CampaignDate).Value
            On Error GoTo 0
            
            If campaignDate = 0 Then GoTo NextRow
            
            If (hasStartDate And campaignDate < startDate) Or (hasEndDate And campaignDate > endDate) Then
                GoTo NextRow
            End If
            
            Dim campaignID As String
            Dim supporterID As String
            
            campaignID = CStr(data.Cells(row, cols.CampaignID).Value)
            supporterID = CStr(data.Cells(row, cols.SupporterID).Value)
            
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
            DictSet uniqueCounts, campaign, campaignSupporters(campaign).Count
        Next campaign
        
        Set CalculateUniqueSupporters = uniqueCounts
    #Else
        Dim uniqueCounts As ReportCounts
        Set uniqueCounts.Dict = CreateObject("Scripting.Dictionary")
        
        Dim campaignSupporters As Object
        Set campaignSupporters = CreateObject("Scripting.Dictionary")
        
        Dim row As Long
        For row = 2 To data.Rows.Count
            Dim campaignDate As Date
            On Error Resume Next
            campaignDate = data.Cells(row, cols.CampaignDate).Value
            On Error GoTo 0
            
            If campaignDate = 0 Then GoTo NextRow
            
            If (hasStartDate And campaignDate < startDate) Or (hasEndDate And campaignDate > endDate) Then
                GoTo NextRow
            End If
            
            Dim campaignID As String
            Dim supporterID As String
            
            campaignID = CStr(data.Cells(row, cols.CampaignID).Value)
            supporterID = CStr(data.Cells(row, cols.SupporterID).Value)
            
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
            uniqueCounts.Dict(campaign) = campaignSupporters(campaign).Count
        Next campaign
        
        Set CalculateUniqueSupporters = uniqueCounts
    #End If
End Function

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
    
    #If Mac Then
        Dim i As Long
        For i = 1 To counts.Dict.Keys.Count
            Dim key As String
            key = counts.Dict.Keys(i)
            
            ws.Cells(row, 1).Value = key
            ws.Cells(row, 1).HorizontalAlignment = xlLeft
            
            ws.Cells(row, 2).Value = counts.Dict.Items(i)
            ws.Cells(row, 2).HorizontalAlignment = xlRight
            
            If Not uniqueCounts Is Nothing Then
                ws.Cells(row, 3).Value = DictGet(uniqueCounts, key)
                ws.Cells(row, 3).HorizontalAlignment = xlRight
            End If
            
            row = row + 1
        Next i
    #Else
        Dim key As Variant
        For Each key In counts.Dict.Keys
            ws.Cells(row, 1).Value = key
            ws.Cells(row, 1).HorizontalAlignment = xlLeft
            
            ws.Cells(row, 2).Value = counts.Dict(key)
            ws.Cells(row, 2).HorizontalAlignment = xlRight
            
            If Not uniqueCounts Is Nothing Then
                If uniqueCounts.Dict.Exists(key) Then
                    ws.Cells(row, 3).Value = uniqueCounts.Dict(key)
                Else
                    ws.Cells(row, 3).Value = 0
                End If
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

Public Sub ProcessCampaignDataExceptSupporter()
    ' Similar to ProcessCampaignData but skips supporter report
    ' Implementation follows the same pattern as ProcessCampaignData
    ' but without the supporter-related code
    
    Application.ScreenUpdating = False
    
    ' ... similar code to ProcessCampaignData ...
    
    ' Update report dates to "Mixed"
    UpdateReportDates "Mixed", "Mixed"
    
    ' ... process data ...
    
    ' Write all reports except supporter
    ' ... write reports ...
    
    MsgBox "✓ Your UAN Reports have been updated! (except by-supporter)", vbInformation
    
    Application.ScreenUpdating = True
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
        Set uniqueCounts = CalculateUniqueSupporters(dataRange, cols, hasStartDate, startDate, hasEndDate, endDate