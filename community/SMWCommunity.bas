Attribute VB_Name = "SMWCommunity"
Option Explicit

Private Const PWD As String = "SMWCommunity2026"
Private Const DASHBOARD_SHEET As String = "Dashboard"
Private Const TRACKER_SHEET As String = "Tracker"
Private Const ONLINE_SHEET As String = "Online Database"
Private Const MANUAL_SHEET As String = "Manual Database"
Private Const COMBINED_SHEET As String = "Hack Database"
Private Const LISTS_SHEET As String = "Lists"
Private Const SETTINGS_SHEET As String = "Settings"
Private Const README_SHEET As String = "Read Me"
Private Const TRACKER_LAST_ROW As Long = 501

Public Sub PrepareCommunityWorkbook()
    On Error GoTo Failed

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    EnsureSupportSheetsHidden
    ApplyWorkbookProtection
    InstallCommunityButtons
    UpdateListsAndNames

    With ThisWorkbook.Worksheets(SETTINGS_SHEET)
        .Range("B5").Value = OfficialHackCount()
        If Len(Trim$(CStr(.Range("B6").Value))) = 0 Then
            .Range("B6").Value = Format$(Date, "yyyy.mm.dd")
        End If
        If Len(Trim$(CStr(.Range("B7").Value))) = 0 Then
            .Range("B7").Value = "Initial embedded catalog"
        End If
    End With

    ThisWorkbook.Worksheets(DASHBOARD_SHEET).Activate
    Application.CalculateFull

CleanExit:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

Failed:
    MsgBox "The Community Tracker could not be prepared:" & vbCrLf & Err.Description, _
        vbExclamation, "Setup Failed"
    Resume CleanExit
End Sub

Public Sub InitializeWorkbook()
    On Error Resume Next
    Application.ScreenUpdating = False
    EnsureSupportSheetsHidden
    ApplyWorkbookProtection
    If Not ShapeExists(ThisWorkbook.Worksheets(DASHBOARD_SHEET), "btnRefreshHacks") Then
        InstallCommunityButtons
    End If
    Application.Calculate
    Application.ScreenUpdating = True
End Sub

Public Sub RefreshHacks()
    Dim catalogUrl As String
    Dim tempPath As String
    Dim sourceBook As Workbook
    Dim errText As String

    On Error GoTo Failed

    catalogUrl = Trim$(CStr(ThisWorkbook.Worksheets(SETTINGS_SHEET).Range("B2").Value))
    If Len(catalogUrl) = 0 Then
        MsgBox "The catalog URL is missing from the Settings sheet.", vbExclamation, "Refresh Failed"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Downloading the latest SMWCentral hack catalog..."

    UnprotectSupportSheets
    tempPath = Environ$("TEMP") & "\SMWCentral_All_Moderated_Hacks.csv"
    DownloadBinaryFile catalogUrl, tempPath

    Workbooks.OpenText Filename:=tempPath, Origin:=65001, StartRow:=1, _
        DataType:=xlDelimited, TextQualifier:=xlTextQualifierDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, _
        Comma:=True, Space:=False, Other:=False, Local:=False

    Set sourceBook = ActiveWorkbook
    ValidateCatalogHeaders sourceBook.Worksheets(1)
    CopyCatalogIntoOnlineDatabase sourceBook.Worksheets(1)
    sourceBook.Close SaveChanges:=False
    Set sourceBook = Nothing

    RebuildCombinedCatalog
    UpdateListsAndNames

    With ThisWorkbook.Worksheets(SETTINGS_SHEET)
        .Range("B5").Value = OfficialHackCount()
        .Range("B6").Value = Format$(Date, "yyyy.mm.dd")
        .Range("B7").Value = Format$(Now, "mmm d, yyyy h:mm AM/PM")
    End With

    ApplyWorkbookProtection
    Application.CalculateFull
    ThisWorkbook.Worksheets(DASHBOARD_SHEET).Activate

    On Error Resume Next
    Kill tempPath
    On Error GoTo 0

    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "The SMWCentral catalog has been refreshed." & vbCrLf & vbCrLf & _
        "Your Tracker entries and custom hacks were not changed.", _
        vbInformation, "Refresh Complete"
    Exit Sub

Failed:
    errText = Err.Description
    On Error Resume Next
    If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    If Len(tempPath) > 0 Then Kill tempPath
    ApplyWorkbookProtection
    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0

    MsgBox "The catalog could not be refreshed:" & vbCrLf & errText & vbCrLf & vbCrLf & _
        "Check your internet connection and try again.", _
        vbExclamation, "Refresh Failed"
End Sub

Public Sub AddCustomHack()
    Dim formWs As Worksheet
    Dim manualWs As Worksheet
    Dim title As String
    Dim label As String
    Dim manualId As String
    Dim nextRow As Long
    Dim errText As String

    On Error GoTo Failed

    Set formWs = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    Set manualWs = ThisWorkbook.Worksheets(MANUAL_SHEET)

    title = Trim$(CStr(formWs.Range("C21").Value))
    If Len(title) = 0 Then
        MsgBox "Enter a custom hack title first.", vbExclamation, "Title Required"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    UnprotectSupportSheets

    manualId = "MANUAL-" & Format$(Now, "yyyymmddhhnnss")
    label = title
    If Application.CountIf(ThisWorkbook.Worksheets(COMBINED_SHEET).Columns("A"), label) > 0 Then
        label = title & " [Manual " & Right$(manualId, 6) & "]"
    End If

    nextRow = manualWs.Cells(manualWs.Rows.Count, "A").End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    With manualWs
        .Cells(nextRow, 1).Value = label
        .Cells(nextRow, 2).Value = title
        .Cells(nextRow, 3).Value = formWs.Range("C22").Value
        .Cells(nextRow, 4).Value = formWs.Range("C23").Value
        .Cells(nextRow, 5).Value = formWs.Range("G23").Value
        .Cells(nextRow, 6).Value = formWs.Range("C24").Value
        If IsDate(formWs.Range("C25").Value) Then
            .Cells(nextRow, 7).Value = CDate(formWs.Range("C25").Value)
        Else
            .Cells(nextRow, 7).Value = Date
        End If
        .Cells(nextRow, 8).Value = manualId
        .Cells(nextRow, 9).Value = formWs.Range("C26").Value
        .Cells(nextRow, 10).Value = formWs.Range("C27").Value
    End With

    RebuildCombinedCatalog
    UpdateListsAndNames
    ClearCustomHackForm
    ApplyWorkbookProtection
    Application.CalculateFull

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "The custom hack was added." & vbCrLf & vbCrLf & _
        "It will remain in this workbook after every future catalog refresh.", _
        vbInformation, "Custom Hack Added"
    Exit Sub

Failed:
    errText = Err.Description
    ApplyWorkbookProtection
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "The custom hack could not be added:" & vbCrLf & errText, _
        vbExclamation, "Add Failed"
End Sub

Public Sub InstallCommunityButtons()
    Dim ws As Worksheet
    Dim target As Range
    Dim button As Shape

    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    ws.Unprotect PWD

    On Error Resume Next
    ws.Shapes("btnRefreshHacks").Delete
    ws.Shapes("btnAddCustomHack").Delete
    On Error GoTo 0

    Set target = ws.Range("G13:L17")
    Set button = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        target.Left + 5, target.Top + 5, target.Width - 10, target.Height - 10)
    With button
        .Name = "btnRefreshHacks"
        .OnAction = "'" & ThisWorkbook.Name & "'!SMWCommunity.RefreshHacks"
        .Placement = xlMoveAndSize
        .Fill.ForeColor.RGB = RGB(20, 125, 146)
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .Line.Weight = 2
        .TextFrame2.TextRange.Text = "REFRESH HACKS"
        .TextFrame2.TextRange.Font.Name = "Aptos Display"
        .TextFrame2.TextRange.Font.Size = 17
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .Locked = True
    End With

    Set target = ws.Range("I20:L23")
    Set button = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        target.Left + 5, target.Top + 5, target.Width - 10, target.Height - 10)
    With button
        .Name = "btnAddCustomHack"
        .OnAction = "'" & ThisWorkbook.Name & "'!SMWCommunity.AddCustomHack"
        .Placement = xlMoveAndSize
        .Fill.ForeColor.RGB = RGB(20, 125, 146)
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .Line.Weight = 2
        .TextFrame2.TextRange.Text = "ADD CUSTOM HACK"
        .TextFrame2.TextRange.Font.Name = "Aptos Display"
        .TextFrame2.TextRange.Font.Size = 14
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .Locked = True
    End With

    ProtectDashboardSheet
End Sub

Private Sub DownloadBinaryFile(ByVal url As String, ByVal destination As String)
    Dim request As Object
    Dim stream As Object

    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.Open "GET", url, False
    request.SetRequestHeader "Cache-Control", "no-cache"
    request.SetRequestHeader "User-Agent", "SMW-Community-Tracker/1.0"
    request.Send

    If request.Status < 200 Or request.Status >= 300 Then
        Err.Raise vbObjectError + 1000, , "The server returned HTTP status " & request.Status
    End If

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write request.ResponseBody
    stream.SaveToFile destination, 2
    stream.Close
End Sub

Private Sub ValidateCatalogHeaders(ByVal ws As Worksheet)
    Dim expected As Variant
    Dim i As Long

    expected = Array( _
        "Dropdown Selection", "ROM Hack Title", "Created By", "Exits", _
        "Difficulty", "Type", "Added Date", "SMWC ID", _
        "SMWCentral Page URL", "Direct Download URL")

    For i = 0 To 9
        If Trim$(CStr(ws.Cells(1, i + 1).Value)) <> expected(i) Then
            Err.Raise vbObjectError + 1001, , _
                "The downloaded catalog has an unexpected column layout."
        End If
    Next i

    If ws.Cells(ws.Rows.Count, "A").End(xlUp).Row < 2 Then
        Err.Raise vbObjectError + 1002, , "The downloaded catalog contains no hacks."
    End If
End Sub

Private Sub CopyCatalogIntoOnlineDatabase(ByVal sourceWs As Worksheet)
    Dim targetWs As Worksheet
    Dim lastRow As Long

    Set targetWs = ThisWorkbook.Worksheets(ONLINE_SHEET)
    lastRow = sourceWs.Cells(sourceWs.Rows.Count, "A").End(xlUp).Row

    targetWs.Range("A1:J10000").ClearContents
    targetWs.Range("A1").Resize(lastRow, 10).Value = _
        sourceWs.Range("A1").Resize(lastRow, 10).Value
    targetWs.Columns("G").NumberFormat = "yyyy-mm-dd"
End Sub

Private Sub RebuildCombinedCatalog()
    Dim onlineWs As Worksheet
    Dim manualWs As Worksheet
    Dim combinedWs As Worksheet
    Dim onlineData As Variant
    Dim manualData As Variant
    Dim uniqueKeys As Object
    Dim rows As Collection
    Dim currentRow As Variant
    Dim output() As Variant
    Dim i As Long
    Dim c As Long
    Dim lastRow As Long

    Set onlineWs = ThisWorkbook.Worksheets(ONLINE_SHEET)
    Set manualWs = ThisWorkbook.Worksheets(MANUAL_SHEET)
    Set combinedWs = ThisWorkbook.Worksheets(COMBINED_SHEET)
    Set uniqueKeys = CreateObject("Scripting.Dictionary")
    Set rows = New Collection

    combinedWs.Range("A1:J10000").ClearContents
    combinedWs.Range("A1:J1").Value = onlineWs.Range("A1:J1").Value

    onlineData = ReadDataRows(onlineWs)
    manualData = ReadDataRows(manualWs)

    AppendUniqueRows onlineData, uniqueKeys, rows
    AppendUniqueRows manualData, uniqueKeys, rows

    If rows.Count > 0 Then
        ReDim output(1 To rows.Count, 1 To 10)
        For i = 1 To rows.Count
            currentRow = rows(i)
            For c = 0 To 9
                output(i, c + 1) = currentRow(c)
            Next c
        Next i
        combinedWs.Range("A2").Resize(rows.Count, 10).Value = output
    End If

    lastRow = combinedWs.Cells(combinedWs.Rows.Count, "A").End(xlUp).Row
    If lastRow >= 2 Then
        combinedWs.Range("A1:J" & lastRow).Sort _
            Key1:=combinedWs.Range("A2"), Order1:=xlAscending, Header:=xlYes
    End If
    combinedWs.Columns("G").NumberFormat = "yyyy-mm-dd"
End Sub

Private Sub AppendUniqueRows(ByVal data As Variant, ByVal uniqueKeys As Object, ByVal rows As Collection)
    Dim i As Long
    Dim key As String

    If IsEmpty(data) Then Exit Sub

    For i = 1 To UBound(data, 1)
        If Len(Trim$(CStr(data(i, 1)))) > 0 Then
            key = CatalogKey(data, i)
            If Len(key) > 0 And Not uniqueKeys.Exists(key) Then
                uniqueKeys.Add key, True
                rows.Add Array( _
                    data(i, 1), data(i, 2), data(i, 3), data(i, 4), data(i, 5), _
                    data(i, 6), data(i, 7), data(i, 8), data(i, 9), data(i, 10))
            End If
        End If
    Next i
End Sub

Private Function ReadDataRows(ByVal ws As Worksheet) As Variant
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    If lastRow < 2 Then
        ReadDataRows = Empty
    Else
        ReadDataRows = ws.Range("A2:J" & lastRow).Value
    End If
End Function

Private Function CatalogKey(ByVal data As Variant, ByVal rowNumber As Long) As String
    Dim idValue As String
    Dim labelValue As String

    idValue = Trim$(CStr(data(rowNumber, 8)))
    labelValue = Trim$(CStr(data(rowNumber, 1)))

    If Len(idValue) > 0 Then
        CatalogKey = UCase$(idValue)
    ElseIf Len(labelValue) > 0 Then
        CatalogKey = "LABEL:" & UCase$(labelValue)
    End If
End Function

Private Sub UpdateListsAndNames()
    Dim dbWs As Worksheet
    Dim listWs As Worksheet
    Dim lastRow As Long

    Set dbWs = ThisWorkbook.Worksheets(COMBINED_SHEET)
    Set listWs = ThisWorkbook.Worksheets(LISTS_SHEET)

    lastRow = dbWs.Cells(dbWs.Rows.Count, "A").End(xlUp).Row
    listWs.Range("A2:A10000").ClearContents

    If lastRow >= 2 Then
        listWs.Range("A2").Resize(lastRow - 1, 1).Value = dbWs.Range("A2:A" & lastRow).Value
    End If

    listWs.Range("B2:B3").Value = Application.Transpose(Array("Yes", "No"))
    listWs.Range("C2:C9").Value = Application.Transpose(Array( _
        "Newcomer", "Casual", "Intermediate", "Advanced", _
        "Expert", "Master", "Grandmaster", "Unranked"))
    listWs.Range("D2:D7").Value = Application.Transpose(Array( _
        "Standard", "Kaizo", "Puzzle", "Tool-Assisted", "Pit", "Troll"))

    On Error Resume Next
    ThisWorkbook.Names("HackList").Delete
    ThisWorkbook.Names("DifficultyList").Delete
    ThisWorkbook.Names("TypeList").Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add Name:="HackList", _
        RefersTo:="='Lists'!$A$2:INDEX('Lists'!$A:$A,COUNTA('Lists'!$A:$A))"
    ThisWorkbook.Names.Add Name:="DifficultyList", RefersTo:="='Lists'!$C$2:$C$9"
    ThisWorkbook.Names.Add Name:="TypeList", RefersTo:="='Lists'!$D$2:$D$7"
End Sub

Private Sub ClearCustomHackForm()
    With ThisWorkbook.Worksheets(DASHBOARD_SHEET)
        .Range("C21").ClearContents
        .Range("C22").ClearContents
        .Range("C23").ClearContents
        .Range("G23").ClearContents
        .Range("C24").ClearContents
        .Range("C25").ClearContents
        .Range("C26").ClearContents
        .Range("C27").ClearContents
    End With
End Sub

Private Function OfficialHackCount() As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(ONLINE_SHEET)
    OfficialHackCount = Application.Max(0, ws.Cells(ws.Rows.Count, "A").End(xlUp).Row - 1)
End Function

Private Sub UnprotectSupportSheets()
    On Error Resume Next
    ThisWorkbook.Worksheets(ONLINE_SHEET).Unprotect PWD
    ThisWorkbook.Worksheets(MANUAL_SHEET).Unprotect PWD
    ThisWorkbook.Worksheets(COMBINED_SHEET).Unprotect PWD
    ThisWorkbook.Worksheets(LISTS_SHEET).Unprotect PWD
    ThisWorkbook.Worksheets(SETTINGS_SHEET).Unprotect PWD
    On Error GoTo 0
End Sub

Private Sub ApplyWorkbookProtection()
    Dim ws As Worksheet

    On Error Resume Next
    ThisWorkbook.Unprotect PWD
    On Error GoTo 0

    ProtectTrackerSheet
    ProtectDashboardSheet

    For Each ws In ThisWorkbook.Worksheets
        Select Case ws.Name
            Case TRACKER_SHEET, DASHBOARD_SHEET
                ' Protected separately.
            Case Else
                ws.Unprotect PWD
                ws.Protect Password:=PWD, DrawingObjects:=True, Contents:=True, _
                    Scenarios:=True, UserInterfaceOnly:=True
        End Select
    Next ws

    EnsureSupportSheetsHidden
    ThisWorkbook.Protect Password:=PWD, Structure:=True, Windows:=False
End Sub

Private Sub ProtectTrackerSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(TRACKER_SHEET)

    ws.Unprotect PWD
    ws.Cells.Locked = True
    ws.Cells.FormulaHidden = False
    ws.Range("B2:B" & TRACKER_LAST_ROW).Locked = False
    ws.Range("H2:J" & TRACKER_LAST_ROW).Locked = False
    ws.Range("L2:M" & TRACKER_LAST_ROW).Locked = False
    ws.Range("P2:P" & TRACKER_LAST_ROW).Locked = False
    ws.Range("A2:A" & TRACKER_LAST_ROW).FormulaHidden = True
    ws.Range("C2:G" & TRACKER_LAST_ROW).FormulaHidden = True
    ws.Range("K2:K" & TRACKER_LAST_ROW).FormulaHidden = True
    ws.Range("N2:O" & TRACKER_LAST_ROW).FormulaHidden = True
    ws.Protect Password:=PWD, DrawingObjects:=True, Contents:=True, _
        Scenarios:=True, UserInterfaceOnly:=True, _
        AllowFiltering:=True, AllowSorting:=True
    ws.EnableSelection = xlNoRestrictions
End Sub

Private Sub ProtectDashboardSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SHEET)

    ws.Unprotect PWD
    ws.Cells.Locked = True
    ws.Range("C21:H22").Locked = False
    ws.Range("C23:D23").Locked = False
    ws.Range("G23:H23").Locked = False
    ws.Range("C24:H24").Locked = False
    ws.Range("C25:D25").Locked = False
    ws.Range("C26:H27").Locked = False
    ws.Protect Password:=PWD, DrawingObjects:=True, Contents:=True, _
        Scenarios:=True, UserInterfaceOnly:=True
    ws.EnableSelection = xlNoRestrictions
End Sub

Private Sub EnsureSupportSheetsHidden()
    Dim supportSheets As Variant
    Dim name As Variant

    supportSheets = Array(ONLINE_SHEET, MANUAL_SHEET, COMBINED_SHEET, _
        LISTS_SHEET, SETTINGS_SHEET, README_SHEET)

    On Error Resume Next
    ThisWorkbook.Unprotect PWD
    For Each name In supportSheets
        ThisWorkbook.Worksheets(CStr(name)).Visible = xlSheetVeryHidden
    Next name
    ThisWorkbook.Worksheets(DASHBOARD_SHEET).Visible = xlSheetVisible
    ThisWorkbook.Worksheets(TRACKER_SHEET).Visible = xlSheetVisible
    On Error GoTo 0
End Sub

Private Function ShapeExists(ByVal ws As Worksheet, ByVal shapeName As String) As Boolean
    Dim shape As Shape
    On Error Resume Next
    Set shape = ws.Shapes(shapeName)
    ShapeExists = Not shape Is Nothing
    On Error GoTo 0
End Function
