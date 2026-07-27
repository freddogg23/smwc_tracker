Attribute VB_Name = "SMWCommunity"
Option Explicit

Private Const DASHBOARD_SHEET As String = "Dashboard"
Private Const TRACKER_SHEET As String = "Tracker"
Private Const ONLINE_SHEET As String = "Online Database"
Private Const MANUAL_SHEET As String = "Manual Database"
Private Const DATABASE_SHEET As String = "Hack Database"
Private Const FINDER_SHEET As String = "Hack Finder"
Private Const LISTS_SHEET As String = "Lists"
Private Const SETTINGS_SHEET As String = "Settings"
Private Const PROTECTION_PASSWORD As String = "SMWCommunity2026"
Private Const SETTINGS_SEQUENCE_ROW As Long = 10
Private Const SETTINGS_MODE_ROW As Long = 11

Public Sub InstallButtons()
    Dim ws As Worksheet
    Dim shp As Shape

    EnsureIncrementalSettings
    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SHEET)

    On Error Resume Next
    ws.Shapes("btnRefreshHacks").Delete
    ws.Shapes("btnAddCustomHack").Delete
    On Error GoTo 0

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        ws.Range("G13").Left, ws.Range("G13").Top, _
        ws.Range("G13:L17").Width, ws.Range("G13:L17").Height)
    With shp
        .Name = "btnRefreshHacks"
        .TextFrame2.TextRange.Text = "REFRESH HACKS"
        .OnAction = "RefreshHacks"
        .Fill.ForeColor.RGB = RGB(20, 125, 146)
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .Line.Weight = 2
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Size = 18
    End With

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        ws.Range("I20").Left, ws.Range("I20").Top, _
        ws.Range("I20:L23").Width, ws.Range("I20:L23").Height)
    With shp
        .Name = "btnAddCustomHack"
        .TextFrame2.TextRange.Text = "ADD CUSTOM HACK"
        .OnAction = "AddCustomHack"
        .Fill.ForeColor.RGB = RGB(20, 125, 146)
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .Line.Weight = 2
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Size = 14
    End With

    If SheetExists(FINDER_SHEET) Then
        Set ws = ThisWorkbook.Worksheets(FINDER_SHEET)
        On Error Resume Next
        ws.Shapes("btnRandomHackFinder").Delete
        On Error GoTo 0

        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
            ws.Range("A19").Left, ws.Range("A19").Top, _
            ws.Range("A19:F21").Width, ws.Range("A19:F21").Height)
        With shp
            .Name = "btnRandomHackFinder"
            .TextFrame2.TextRange.Text = "RANDOM HACK"
            .OnAction = "RandomHackFinder"
            .Fill.ForeColor.RGB = RGB(20, 125, 146)
            .Line.ForeColor.RGB = RGB(0, 0, 0)
            .Line.Weight = 2
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextFrame2.TextRange.Font.Bold = msoTrue
            .TextFrame2.TextRange.Font.Size = 18
        End With
    End If
End Sub

Public Sub UpgradeCommunityWorkbook()
    EnsureRatingColumns
    ConfigureLists
    ConfigureTrackerRatingColumn
    ConfigureHackFinder
    UpdateDropdownList
End Sub

Private Sub EnsureRatingColumns()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = ONLINE_SHEET Or ws.Name = MANUAL_SHEET Or ws.Name = DATABASE_SHEET Then
            ws.Range("K1").Value = "SMWCentral Rating"
            ws.Columns("K").ColumnWidth = 18
        End If
    Next ws
End Sub

Private Sub ConfigureLists()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(LISTS_SHEET)

    ws.Range("H1:H7").Value = Application.Transpose(Array("SMWCentral Rating Filter", "Any", "1+", "2+", "3+", "4+", "5"))
    ws.Columns("H").ColumnWidth = 24

    On Error Resume Next
    ThisWorkbook.Names("DifficultyList").Delete
    ThisWorkbook.Names("TypeList").Delete
    ThisWorkbook.Names("SMWCRatingFilterList").Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add Name:="DifficultyList", RefersTo:="='" & LISTS_SHEET & "'!$C$2:$C$9"
    ThisWorkbook.Names.Add Name:="TypeList", RefersTo:="='" & LISTS_SHEET & "'!$D$2:$D$7"
    ThisWorkbook.Names.Add Name:="SMWCRatingFilterList", RefersTo:="='" & LISTS_SHEET & "'!$H$2:$H$7"
End Sub

Private Sub ConfigureTrackerRatingColumn()
    Dim ws As Worksheet
    Dim lastRow As Long
    Set ws = ThisWorkbook.Worksheets(TRACKER_SHEET)

    ws.Range("S1").Value = "SMWCentral Rating"
    ws.Range("S1").Interior.Color = RGB(23, 54, 93)
    ws.Range("S1").Font.Color = RGB(255, 255, 255)
    ws.Range("S1").Font.Bold = True
    ws.Range("S1").HorizontalAlignment = xlCenter
    ws.Range("S1").VerticalAlignment = xlCenter
    ws.Columns("S").ColumnWidth = 18

    lastRow = 501
    ws.Range("S2:S" & lastRow).Formula = "=IFERROR(VLOOKUP($B2,'Hack Database'!$A:$K,11,FALSE),"""")"
    ws.Range("S2:S" & lastRow).HorizontalAlignment = xlCenter
End Sub

Private Sub ConfigureHackFinder()
    If Not SheetExists(FINDER_SHEET) Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(FINDER_SHEET)

    ws.Range("B5").Formula = "=IFERROR(VLOOKUP($B$3,'Hack Database'!$A:$K,2,FALSE),"""")"
    ws.Range("B6").Formula = "=IFERROR(VLOOKUP($B$3,'Hack Database'!$A:$K,3,FALSE),"""")"
    ws.Range("B7").Formula = "=IFERROR(VLOOKUP($B$3,'Hack Database'!$A:$K,4,FALSE),"""")"
    ws.Range("B8").Formula = "=IFERROR(VLOOKUP($B$3,'Hack Database'!$A:$K,5,FALSE),"""")"
    ws.Range("B9").Formula = "=IFERROR(VLOOKUP($B$3,'Hack Database'!$A:$K,6,FALSE),"""")"
    ws.Range("B10").Formula = "=IFERROR(VLOOKUP($B$3,'Hack Database'!$A:$K,7,FALSE),"""")"
    ws.Range("B11").Formula = "=IFERROR(VLOOKUP($B$3,'Hack Database'!$A:$K,8,FALSE),"""")"
    ws.Range("B12").Formula = "=IFERROR(HYPERLINK(VLOOKUP($B$3,'Online Database'!$A:$K,9,FALSE),""Open SMWCentral Page""),"""")"
    ws.Range("B13").Formula = "=IFERROR(HYPERLINK(VLOOKUP($B$3,'Online Database'!$A:$K,10,FALSE),""Download Patch""),"""")"
    ws.Range("A14").Value = "SMWCentral Rating"
    ws.Range("B14").Formula = "=IFERROR(VLOOKUP($B$3,'Hack Database'!$A:$K,11,FALSE),"""")"

    ws.Range("A16:F16").Merge
    ws.Range("A16").Value = "Random Hack Filters"
    ws.Range("A16").Interior.Color = RGB(23, 54, 93)
    ws.Range("A16").Font.Color = RGB(255, 255, 255)
    ws.Range("A16").Font.Bold = True
    ws.Range("A16").HorizontalAlignment = xlCenter

    ws.Range("A17").Value = "Difficulty"
    ws.Range("C17").Value = "Type"
    ws.Range("E17").Value = "SMWC Rating"
    ws.Range("A17:F17").Interior.Color = RGB(255, 248, 231)
    ws.Range("A17:F17").Font.Bold = True
    ws.Range("A17:F17").HorizontalAlignment = xlCenter

    With ws.Range("B17").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=DifficultyList"
    End With
    With ws.Range("D17").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=TypeList"
    End With
    With ws.Range("F17").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=SMWCRatingFilterList"
    End With
    If Len(Trim$(CStr(ws.Range("F17").Value))) = 0 Then ws.Range("F17").Value = "Any"

    ws.Range("A19:F21").Merge
    ws.Range("A19").Value = "RANDOM HACK"
    ws.Range("A19").Interior.Color = RGB(20, 125, 146)
    ws.Range("A19").Font.Color = RGB(255, 255, 255)
    ws.Range("A19").Font.Bold = True
    ws.Range("A19").Font.Size = 18
    ws.Range("A19").HorizontalAlignment = xlCenter
    ws.Range("A19").VerticalAlignment = xlCenter
End Sub

Public Sub RandomHackFinder()
    If Not SheetExists(FINDER_SHEET) Then
        MsgBox "Hack Finder sheet is missing.", vbExclamation
        Exit Sub
    End If

    Dim finder As Worksheet, db As Worksheet
    Set finder = ThisWorkbook.Worksheets(FINDER_SHEET)
    Set db = ThisWorkbook.Worksheets(DATABASE_SHEET)

    Dim diffFilter As String, typeFilter As String, ratingFilter As String
    diffFilter = Trim$(CStr(finder.Range("B17").Value))
    typeFilter = Trim$(CStr(finder.Range("D17").Value))
    ratingFilter = Trim$(CStr(finder.Range("F17").Value))
    If Len(ratingFilter) = 0 Then ratingFilter = "Any"

    Dim lastRow As Long
    lastRow = LastUsedRow(db, "A")

    Dim matches As Collection
    Set matches = New Collection

    Dim r As Long
    For r = 2 To lastRow
        If RowMatchesRandomFilters(db, r, diffFilter, typeFilter, ratingFilter) Then
            matches.Add db.Cells(r, 1).Value
        End If
    Next r

    If matches.Count = 0 Then
        MsgBox "No hacks matched those filters.", vbInformation, "Random Hack Finder"
        Exit Sub
    End If

    Randomize
    Dim pickIndex As Long
    pickIndex = Int(matches.Count * Rnd) + 1

    finder.Range("B3").Value = matches.Item(pickIndex)
    MsgBox "Selected: " & CStr(matches.Item(pickIndex)) & vbCrLf & _
           "Matches found: " & matches.Count, vbInformation, "Random Hack Finder"
End Sub

Private Function RowMatchesRandomFilters(ByVal db As Worksheet, ByVal rowNum As Long, ByVal diffFilter As String, ByVal typeFilter As String, ByVal ratingFilter As String) As Boolean
    Dim dbDiff As String, dbType As String, dbRating As String
    dbDiff = Trim$(CStr(db.Cells(rowNum, 5).Value))
    dbType = Trim$(CStr(db.Cells(rowNum, 6).Value))
    dbRating = Trim$(CStr(db.Cells(rowNum, 11).Value))

    If Len(diffFilter) > 0 And StrComp(diffFilter, dbDiff, vbTextCompare) <> 0 Then
        RowMatchesRandomFilters = False
        Exit Function
    End If

    If Len(typeFilter) > 0 And InStr(1, dbType, typeFilter, vbTextCompare) = 0 Then
        RowMatchesRandomFilters = False
        Exit Function
    End If

    If Not RatingMatches(dbRating, ratingFilter) Then
        RowMatchesRandomFilters = False
        Exit Function
    End If

    RowMatchesRandomFilters = True
End Function

Private Function RatingMatches(ByVal ratingText As String, ByVal filterText As String) As Boolean
    filterText = Trim$(filterText)
    If Len(filterText) = 0 Or StrComp(filterText, "Any", vbTextCompare) = 0 Then
        RatingMatches = True
        Exit Function
    End If

    Dim rating As Double
    rating = ExtractNumber(ratingText)

    If rating <= 0 Then
        RatingMatches = False
        Exit Function
    End If

    Select Case filterText
        Case "1+"
            RatingMatches = (rating >= 1)
        Case "2+"
            RatingMatches = (rating >= 2)
        Case "3+"
            RatingMatches = (rating >= 3)
        Case "4+"
            RatingMatches = (rating >= 4)
        Case "5"
            RatingMatches = (rating >= 5)
        Case Else
            RatingMatches = True
    End Select
End Function

Private Function ExtractNumber(ByVal text As String) As Double
    Dim i As Long, ch As String, buf As String
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "." Then
            buf = buf & ch
        ElseIf Len(buf) > 0 Then
            Exit For
        End If
    Next i
    If Len(buf) = 0 Then
        ExtractNumber = 0
    Else
        ExtractNumber = Val(buf)
    End If
End Function

Public Sub RefreshHacks()
    On Error GoTo Failed

    Dim settingsWs As Worksheet
    Dim versionUrl As String
    Dim versionText As String
    Dim remoteSequence As Long
    Dim localSequence As Long
    Dim remoteCount As Long
    Dim remoteVersion As String
    Dim nextSequence As Long
    Dim deltaUrl As String
    Dim addedCount As Long
    Dim updatedCount As Long
    Dim removedCount As Long
    Dim answer As VbMsgBoxResult

    Set settingsWs = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    EnsureIncrementalSettings

    versionUrl = Trim$(CStr(settingsWs.Range("B3").Value))
    If Len(versionUrl) = 0 Or InStr(1, versionUrl, "raw.githubusercontent.com", vbTextCompare) = 0 Then
        MsgBox "Version JSON URL is missing or invalid in Settings!B3.", vbExclamation, "Refresh Hacks"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Checking SMWCentral catalog version..."

    versionText = HttpGetText(versionUrl)
    remoteSequence = JsonLong(versionText, "sequence", -1)
    remoteCount = JsonLong(versionText, "hack_count", 0)
    remoteVersion = JsonString(versionText, "catalog_version", "")
    localSequence = CLng(Val(settingsWs.Cells(SETTINGS_SEQUENCE_ROW, 2).Value))

    If remoteSequence < 0 Then
        answer = MsgBox( _
            "The repository is not publishing incremental updates yet." & vbCrLf & vbCrLf & _
            "Download the full catalog this one time?", _
            vbYesNo + vbQuestion, "Refresh Hacks")
        If answer = vbYes Then
            ForceFullCatalogRefreshInternal versionText
        Else
            GoTo CleanExit
        End If
        GoTo CompleteRefresh
    End If

    If localSequence > remoteSequence Then
        answer = MsgBox( _
            "This workbook's catalog sequence is newer than the repository." & vbCrLf & vbCrLf & _
            "Perform a full catalog synchronization?", _
            vbYesNo + vbQuestion, "Refresh Hacks")
        If answer = vbYes Then
            ForceFullCatalogRefreshInternal versionText
        Else
            GoTo CleanExit
        End If
        GoTo CompleteRefresh
    End If

    If localSequence = remoteSequence Then
        UpdateRefreshMetadata remoteSequence, remoteCount, remoteVersion
        MsgBox "Your hack catalog is already up to date." & vbCrLf & _
               "Catalog sequence: " & remoteSequence, vbInformation, "Refresh Hacks"
        GoTo CleanExit
    End If

    For nextSequence = localSequence + 1 To remoteSequence
        Application.StatusBar = "Applying catalog update " & nextSequence & " of " & remoteSequence & "..."
        deltaUrl = BuildDeltaUrl(versionUrl, nextSequence)
        ApplyDeltaCsv deltaUrl, addedCount, updatedCount, removedCount
        settingsWs.Cells(SETTINGS_SEQUENCE_ROW, 2).Value = nextSequence
    Next nextSequence

    SortDatabaseSheet ThisWorkbook.Worksheets(ONLINE_SHEET)

CompleteRefresh:
    RebuildCombinedDatabase
    SortDatabaseSheet ThisWorkbook.Worksheets(DATABASE_SHEET)
    UpdateDropdownList
    UpgradeCommunityWorkbook
    UpdateRefreshMetadata remoteSequence, remoteCount, remoteVersion
    PrepareWorkbookForRelease

    MsgBox "Incremental refresh complete." & vbCrLf & vbCrLf & _
           "Added: " & addedCount & vbCrLf & _
           "Updated: " & updatedCount & vbCrLf & _
           "Removed: " & removedCount & vbCrLf & vbCrLf & _
           "The full catalog was not downloaded.", _
           vbInformation, "Refresh Complete"

CleanExit:
    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

Failed:
    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    answer = MsgBox( _
        "Incremental refresh failed:" & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
        "Would you like to download the full catalog as a recovery step?", _
        vbYesNo + vbExclamation, "Refresh Hacks")

    If answer = vbYes Then
        On Error GoTo FullFailed
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.DisplayAlerts = False
        ForceFullCatalogRefreshInternal versionText
        RebuildCombinedDatabase
        SortDatabaseSheet ThisWorkbook.Worksheets(DATABASE_SHEET)
        UpdateDropdownList
        UpgradeCommunityWorkbook
        UpdateRefreshMetadata remoteSequence, remoteCount, remoteVersion
        PrepareWorkbookForRelease
        Application.StatusBar = False
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        MsgBox "Full catalog recovery refresh completed.", vbInformation, "Refresh Complete"
        Exit Sub
    End If
    Exit Sub

FullFailed:
    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Full catalog recovery also failed:" & vbCrLf & Err.Description, vbCritical, "Refresh Hacks"
End Sub

Public Sub ForceFullCatalogRefresh()
    On Error GoTo Failed
    Dim versionText As String
    Dim versionUrl As String
    Dim remoteSequence As Long
    Dim remoteCount As Long
    Dim remoteVersion As String

    EnsureIncrementalSettings
    versionUrl = Trim$(CStr(ThisWorkbook.Worksheets(SETTINGS_SHEET).Range("B3").Value))
    versionText = HttpGetText(versionUrl)
    remoteSequence = JsonLong(versionText, "sequence", 0)
    remoteCount = JsonLong(versionText, "hack_count", 0)
    remoteVersion = JsonString(versionText, "catalog_version", "")

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ForceFullCatalogRefreshInternal versionText
    RebuildCombinedDatabase
    SortDatabaseSheet ThisWorkbook.Worksheets(DATABASE_SHEET)
    UpdateDropdownList
    UpgradeCommunityWorkbook
    UpdateRefreshMetadata remoteSequence, remoteCount, remoteVersion
    PrepareWorkbookForRelease

    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Full catalog refresh completed.", vbInformation, "Refresh Complete"
    Exit Sub

Failed:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Full catalog refresh failed:" & vbCrLf & Err.Description, vbCritical, "Refresh Hacks"
End Sub

Private Sub EnsureIncrementalSettings()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)

    If Len(Trim$(CStr(ws.Cells(SETTINGS_SEQUENCE_ROW, 1).Value))) = 0 Then
        ws.Cells(SETTINGS_SEQUENCE_ROW, 1).Value = "Catalog Sequence"
    End If
    If Len(Trim$(CStr(ws.Cells(SETTINGS_SEQUENCE_ROW, 2).Value))) = 0 Then
        ws.Cells(SETTINGS_SEQUENCE_ROW, 2).Value = 0
    End If
    If Len(Trim$(CStr(ws.Cells(SETTINGS_MODE_ROW, 1).Value))) = 0 Then
        ws.Cells(SETTINGS_MODE_ROW, 1).Value = "Refresh Mode"
    End If
    ws.Cells(SETTINGS_MODE_ROW, 2).Value = "Incremental"
End Sub

Private Sub UpdateRefreshMetadata(ByVal sequence As Long, ByVal hackCount As Long, ByVal catalogVersion As String)
    Dim settingsWs As Worksheet
    Set settingsWs = ThisWorkbook.Worksheets(SETTINGS_SHEET)

    settingsWs.Cells(SETTINGS_SEQUENCE_ROW, 2).Value = sequence
    settingsWs.Cells(SETTINGS_MODE_ROW, 2).Value = "Incremental"
    If hackCount > 0 Then settingsWs.Range("B5").Value = hackCount
    If Len(catalogVersion) > 0 Then settingsWs.Range("B6").Value = catalogVersion
    settingsWs.Range("B7").Value = Now
    settingsWs.Range("B7").NumberFormat = "mmm d, yyyy h:mm AM/PM"

    With ThisWorkbook.Worksheets(DASHBOARD_SHEET)
        .Range("J9").Value = Now
        .Range("J9").NumberFormat = "mmm d, yyyy h:mm AM/PM"
    End With
End Sub

Private Function BuildDeltaUrl(ByVal versionUrl As String, ByVal sequence As Long) As String
    Dim slashPos As Long
    slashPos = InStrRev(versionUrl, "/")
    If slashPos = 0 Then Err.Raise vbObjectError + 1100, , "Invalid version URL."
    BuildDeltaUrl = Left$(versionUrl, slashPos) & "deltas/" & Format$(sequence, "00000000") & ".csv"
End Function

Private Sub ApplyDeltaCsv(ByVal url As String, ByRef addedCount As Long, ByRef updatedCount As Long, ByRef removedCount As Long)
    Dim tempPath As String
    Dim csvBook As Workbook
    Dim csvSheet As Worksheet
    Dim onlineWs As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim operation As String
    Dim smwcId As String
    Dim targetRow As Long

    tempPath = Environ$("TEMP") & "\SMWCatalogDelta_" & Format$(Now, "yyyymmddhhnnss") & ".csv"
    DownloadBinaryFile url, tempPath

    Workbooks.OpenText Filename:=tempPath, Origin:=65001, _
        DataType:=xlDelimited, Comma:=True, TextQualifier:=xlTextQualifierDoubleQuote

    Set csvBook = ActiveWorkbook
    Set csvSheet = csvBook.Worksheets(1)
    Set onlineWs = ThisWorkbook.Worksheets(ONLINE_SHEET)

    If UCase$(Trim$(CStr(csvSheet.Cells(1, 1).Value))) <> "OPERATION" Then
        Err.Raise vbObjectError + 1101, , "Invalid delta CSV header."
    End If

    lastRow = LastUsedRow(csvSheet, "A")
    For r = 2 To lastRow
        operation = UCase$(Trim$(CStr(csvSheet.Cells(r, 1).Value)))
        smwcId = Trim$(CStr(csvSheet.Cells(r, 9).Value))

        If operation = "UPSERT" Then
            targetRow = FindOnlineRowById(onlineWs, smwcId)
            If targetRow = 0 Then
                targetRow = LastUsedRow(onlineWs, "A") + 1
                addedCount = addedCount + 1
            Else
                updatedCount = updatedCount + 1
            End If
            onlineWs.Cells(targetRow, 1).Resize(1, 11).Value = csvSheet.Cells(r, 2).Resize(1, 11).Value

        ElseIf operation = "DELETE" Then
            targetRow = FindOnlineRowById(onlineWs, smwcId)
            If targetRow > 0 Then
                onlineWs.Rows(targetRow).Delete
                removedCount = removedCount + 1
            End If
        End If
    Next r

    csvBook.Close SaveChanges:=False
    On Error Resume Next
    Kill tempPath
    On Error GoTo 0
End Sub

Private Function FindOnlineRowById(ByVal ws As Worksheet, ByVal smwcId As String) As Long
    Dim found As Range
    If Len(smwcId) = 0 Then Exit Function

    Set found = ws.Columns(8).Find( _
        What:=smwcId, _
        After:=ws.Cells(1, 8), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False)

    If Not found Is Nothing Then FindOnlineRowById = found.Row
End Function

Private Function HttpGetText(ByVal url As String) As String
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.setRequestHeader "Cache-Control", "no-cache"
    http.send

    If http.Status < 200 Or http.Status >= 300 Then
        Err.Raise vbObjectError + 1102, , "HTTP " & http.Status & " while downloading " & url
    End If
    HttpGetText = CStr(http.responseText)
End Function

Private Sub DownloadBinaryFile(ByVal url As String, ByVal destination As String)
    Dim http As Object
    Dim stream As Object

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.setRequestHeader "Cache-Control", "no-cache"
    http.send

    If http.Status < 200 Or http.Status >= 300 Then
        Err.Raise vbObjectError + 1103, , "HTTP " & http.Status & " while downloading " & url
    End If

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile destination, 2
    stream.Close
End Sub

Private Function JsonLong(ByVal jsonText As String, ByVal key As String, ByVal defaultValue As Long) As Long
    Dim marker As String
    Dim p As Long
    Dim startPos As Long
    Dim endPos As Long
    Dim ch As String
    Dim numberText As String

    marker = Chr$(34) & key & Chr$(34)
    p = InStr(1, jsonText, marker, vbTextCompare)
    If p = 0 Then
        JsonLong = defaultValue
        Exit Function
    End If

    p = InStr(p + Len(marker), jsonText, ":")
    If p = 0 Then
        JsonLong = defaultValue
        Exit Function
    End If

    startPos = p + 1
    Do While startPos <= Len(jsonText)
        ch = Mid$(jsonText, startPos, 1)
        If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then Exit Do
        startPos = startPos + 1
    Loop

    endPos = startPos
    Do While endPos <= Len(jsonText)
        ch = Mid$(jsonText, endPos, 1)
        If (ch < "0" Or ch > "9") And ch <> "-" Then Exit Do
        endPos = endPos + 1
    Loop

    numberText = Mid$(jsonText, startPos, endPos - startPos)
    If Len(numberText) = 0 Then
        JsonLong = defaultValue
    Else
        JsonLong = CLng(numberText)
    End If
End Function

Private Function JsonString(ByVal jsonText As String, ByVal key As String, ByVal defaultValue As String) As String
    Dim marker As String
    Dim p As Long
    Dim colonPos As Long
    Dim quoteStart As Long
    Dim quoteEnd As Long

    marker = Chr$(34) & key & Chr$(34)
    p = InStr(1, jsonText, marker, vbTextCompare)
    If p = 0 Then
        JsonString = defaultValue
        Exit Function
    End If

    colonPos = InStr(p + Len(marker), jsonText, ":")
    quoteStart = InStr(colonPos + 1, jsonText, Chr$(34))
    quoteEnd = InStr(quoteStart + 1, jsonText, Chr$(34))

    If quoteStart = 0 Or quoteEnd = 0 Then
        JsonString = defaultValue
    Else
        JsonString = Mid$(jsonText, quoteStart + 1, quoteEnd - quoteStart - 1)
    End If
End Function

Private Sub ForceFullCatalogRefreshInternal(ByVal versionText As String)
    Dim csvUrl As String
    Dim remoteSequence As Long
    csvUrl = Trim$(CStr(ThisWorkbook.Worksheets(SETTINGS_SHEET).Range("B2").Value))
    If Len(csvUrl) = 0 Then Err.Raise vbObjectError + 1104, , "Catalog CSV URL is missing."

    DownloadCatalogToOnlineDatabase csvUrl
    remoteSequence = JsonLong(versionText, "sequence", 0)
    ThisWorkbook.Worksheets(SETTINGS_SHEET).Cells(SETTINGS_SEQUENCE_ROW, 2).Value = remoteSequence
End Sub

Private Sub DownloadCatalogToOnlineDatabase(ByVal url As String)
    Dim tempPath As String
    Dim csvBook As Workbook
    Dim csvSheet As Worksheet
    Dim target As Worksheet
    Dim lastRow As Long

    tempPath = Environ$("TEMP") & "\SMWCentral_All_Moderated_Hacks.csv"
    DownloadBinaryFile url, tempPath

    Workbooks.OpenText Filename:=tempPath, Origin:=65001, _
        DataType:=xlDelimited, Comma:=True, TextQualifier:=xlTextQualifierDoubleQuote

    Set csvBook = ActiveWorkbook
    Set csvSheet = csvBook.Worksheets(1)
    Set target = ThisWorkbook.Worksheets(ONLINE_SHEET)

    target.Range("A1:K10000").ClearContents
    lastRow = LastUsedRow(csvSheet, "A")
    target.Range("A1").Resize(lastRow, 11).Value = csvSheet.Range("A1").Resize(lastRow, 11).Value

    csvBook.Close SaveChanges:=False
    On Error Resume Next
    Kill tempPath
    On Error GoTo 0
End Sub

Private Sub RebuildCombinedDatabase()
    Dim onlineWs As Worksheet
    Dim manualWs As Worksheet
    Dim dbWs As Worksheet
    Dim dict As Object
    Dim sourceData As Variant
    Dim output() As Variant
    Dim r As Long, c As Long, outRow As Long
    Dim lastOnline As Long, lastManual As Long, maxRows As Long
    Dim key As String

    Set onlineWs = ThisWorkbook.Worksheets(ONLINE_SHEET)
    Set manualWs = ThisWorkbook.Worksheets(MANUAL_SHEET)
    Set dbWs = ThisWorkbook.Worksheets(DATABASE_SHEET)
    Set dict = CreateObject("Scripting.Dictionary")

    dbWs.Hyperlinks.Delete
    dbWs.Range("A1:K10000").ClearContents
    dbWs.Range("A1:K1").Value = onlineWs.Range("A1:K1").Value

    lastOnline = LastUsedRow(onlineWs, "A")
    lastManual = LastUsedRow(manualWs, "A")
    maxRows = Application.Max(0, lastOnline - 1) + Application.Max(0, lastManual - 1)
    If maxRows = 0 Then Exit Sub

    ReDim output(1 To maxRows, 1 To 11)
    outRow = 0

    If lastOnline >= 2 Then
        sourceData = onlineWs.Range("A2:K" & lastOnline).Value
        For r = 1 To UBound(sourceData, 1)
            key = MakeCatalogKey(sourceData, r)
            If Len(key) > 0 And Not dict.Exists(key) Then
                dict.Add key, True
                outRow = outRow + 1
                For c = 1 To 11
                    output(outRow, c) = sourceData(r, c)
                Next c
            End If
        Next r
    End If

    If lastManual >= 2 Then
        sourceData = manualWs.Range("A2:K" & lastManual).Value
        For r = 1 To UBound(sourceData, 1)
            key = MakeCatalogKey(sourceData, r)
            If Len(key) > 0 And Not dict.Exists(key) Then
                dict.Add key, True
                outRow = outRow + 1
                For c = 1 To 11
                    output(outRow, c) = sourceData(r, c)
                Next c
            End If
        Next r
    End If

    If outRow > 0 Then
        dbWs.Range("A2").Resize(outRow, 11).Value = output
        MakeHackDatabaseLinksFriendly dbWs, outRow + 1
    End If
End Sub

Private Function MakeCatalogKey(ByVal data As Variant, ByVal rowNumber As Long) As String
    Dim idValue As String
    Dim labelValue As String

    idValue = Trim$(CStr(data(rowNumber, 8)))
    labelValue = Trim$(CStr(data(rowNumber, 1)))

    If Len(idValue) > 0 Then
        MakeCatalogKey = UCase$(idValue)
    ElseIf Len(labelValue) > 0 Then
        MakeCatalogKey = "LABEL:" & UCase$(labelValue)
    Else
        MakeCatalogKey = ""
    End If
End Function

Private Sub MakeHackDatabaseLinksFriendly(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim r As Long
    Dim pageUrl As String
    Dim downloadUrl As String

    For r = 2 To lastRow
        pageUrl = Trim$(CStr(ws.Cells(r, 9).Value))
        downloadUrl = Trim$(CStr(ws.Cells(r, 10).Value))

        If Len(pageUrl) > 0 Then
            ws.Hyperlinks.Add Anchor:=ws.Cells(r, 9), Address:=pageUrl, TextToDisplay:="Open Page"
        End If
        If Len(downloadUrl) > 0 Then
            ws.Hyperlinks.Add Anchor:=ws.Cells(r, 10), Address:=downloadUrl, TextToDisplay:="Download"
        End If
    Next r
End Sub

Private Sub SortDatabaseSheet(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = LastUsedRow(ws, "A")
    If lastRow < 3 Then Exit Sub

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("A2:A" & lastRow), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range("A1:K" & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

Private Sub UpdateDropdownList()
    Dim dbWs As Worksheet
    Dim listWs As Worksheet
    Dim lastRow As Long

    Set dbWs = ThisWorkbook.Worksheets(DATABASE_SHEET)
    Set listWs = ThisWorkbook.Worksheets(LISTS_SHEET)

    lastRow = LastUsedRow(dbWs, "A")
    listWs.Range("A2:A10000").ClearContents
    If lastRow >= 2 Then
        listWs.Range("A2").Resize(lastRow - 1, 1).Value = dbWs.Range("A2:A" & lastRow).Value
    End If

    On Error Resume Next
    ThisWorkbook.Names("HackList").Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="HackList", _
        RefersTo:="='" & LISTS_SHEET & "'!$A$2:INDEX('" & LISTS_SHEET & "'!$A:$A,COUNTA('" & LISTS_SHEET & "'!$A:$A))"
End Sub

Public Sub AddCustomHack()
    On Error GoTo Failed

    Dim formWs As Worksheet
    Dim manualWs As Worksheet
    Dim title As String
    Dim label As String
    Dim nextRow As Long
    Dim manualId As String

    Set formWs = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    Set manualWs = ThisWorkbook.Worksheets(MANUAL_SHEET)

    title = Trim$(CStr(formWs.Range("C21").Value))
    If Len(title) = 0 Then
        MsgBox "Enter a custom hack title first.", vbExclamation, "Add Custom Hack"
        Exit Sub
    End If

    label = title
    If Not IsError(Application.Match(label, ThisWorkbook.Worksheets(DATABASE_SHEET).Range("A:A"), 0)) Then
        label = title & " [Manual]"
    End If

    manualId = "MANUAL-" & Format(Now, "yyyymmddhhnnss")
    nextRow = LastUsedRow(manualWs, "A") + 1
    If nextRow < 2 Then nextRow = 2

    With manualWs
        .Cells(nextRow, 1).Value = label
        .Cells(nextRow, 2).Value = title
        .Cells(nextRow, 3).Value = formWs.Range("C22").Value
        .Cells(nextRow, 4).Value = formWs.Range("C23").Value
        .Cells(nextRow, 5).Value = formWs.Range("G23").Value
        .Cells(nextRow, 6).Value = formWs.Range("C24").Value
        .Cells(nextRow, 7).Value = formWs.Range("C25").Value
        .Cells(nextRow, 8).Value = manualId
        .Cells(nextRow, 9).Value = formWs.Range("C26").Value
        .Cells(nextRow, 10).Value = formWs.Range("C27").Value
        .Cells(nextRow, 11).Value = ""
    End With

    RebuildCombinedDatabase
    SortDatabaseSheet ThisWorkbook.Worksheets(DATABASE_SHEET)
    UpdateDropdownList
    UpgradeCommunityWorkbook
    formWs.Range("C21:H27").ClearContents

    MsgBox "Custom hack added. It will not be removed by Refresh Hacks.", vbInformation, "Add Custom Hack"
    Exit Sub

Failed:
    MsgBox "Custom hack could not be added:" & vbCrLf & Err.Description, vbExclamation, "Add Custom Hack"
End Sub

Public Sub PrepareWorkbookForRelease()
    Dim ws As Worksheet

    EnsureIncrementalSettings
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=PROTECTION_PASSWORD
    On Error GoTo 0

    For Each ws In ThisWorkbook.Worksheets
        Select Case ws.Name
            Case DASHBOARD_SHEET, TRACKER_SHEET, DATABASE_SHEET, FINDER_SHEET
                ws.Visible = xlSheetVisible
            Case Else
                ws.Visible = xlSheetVeryHidden
        End Select
    Next ws

    ThisWorkbook.Protect Password:=PROTECTION_PASSWORD, Structure:=True
End Sub

Private Function LastUsedRow(ByVal ws As Worksheet, ByVal colLetter As String) As Long
    LastUsedRow = ws.Cells(ws.Rows.Count, colLetter).End(xlUp).Row
End Function

Private Function SheetExists(ByVal sheetName As String) As Boolean
    On Error Resume Next
    SheetExists = Not ThisWorkbook.Worksheets(sheetName) Is Nothing
    On Error GoTo 0
End Function
