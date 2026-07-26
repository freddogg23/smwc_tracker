Attribute VB_Name = "SMWHackCatalog"
Option Explicit

Private Const DASH As String = "Dashboard"
Private Const ONLINE As String = "Online Database"
Private Const MANUAL As String = "Manual Database"
Private Const COMBINED As String = "Hack Database"
Private Const LISTS As String = "Lists"

Public Sub InstallButtons()
    Dim ws As Worksheet, shp As Shape
    Set ws = ThisWorkbook.Worksheets(DASH)
    On Error Resume Next
    ws.Shapes("btnRefreshHacks").Delete
    ws.Shapes("btnAddCustomHack").Delete
    On Error GoTo 0

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("I4").Left, ws.Range("I4").Top, ws.Range("I4:K4").Width, ws.Range("I4:K4").Height)
    With shp
        .Name = "btnRefreshHacks": .TextFrame2.TextRange.Text = "Refresh Hacks": .OnAction = "RefreshHacks"
        .Fill.ForeColor.RGB = RGB(31, 78, 121): .Line.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255): .TextFrame2.TextRange.Font.Bold = msoTrue
    End With

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("I21").Left, ws.Range("I21").Top, ws.Range("I21:K21").Width, ws.Range("I21:K21").Height)
    With shp
        .Name = "btnAddCustomHack": .TextFrame2.TextRange.Text = "Add Custom Hack": .OnAction = "AddCustomHack"
        .Fill.ForeColor.RGB = RGB(31, 109, 122): .Line.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255): .TextFrame2.TextRange.Font.Bold = msoTrue
    End With
End Sub

Public Sub RefreshHacks()
    Dim url As String
    On Error GoTo Failed
    url = Trim$(CStr(ThisWorkbook.Worksheets(DASH).Range("J10").Value))
    If Len(url) = 0 Or InStr(1, url, "PASTE_RAW_", vbTextCompare) > 0 Then
        MsgBox "Paste the raw GitHub CSV URL into Dashboard cell J10 first.", vbExclamation: Exit Sub
    End If
    Application.ScreenUpdating = False: Application.EnableEvents = False: Application.DisplayAlerts = False
    Application.StatusBar = "Downloading the latest SMWCentral hack catalog..."
    DownloadCatalog url
    RebuildCombined
    UpdateLists
    With ThisWorkbook.Worksheets(DASH)
        .Range("J2").Value = "Refresh complete": .Range("J3").Value = Format(Now, "mmm d, yyyy h:mm AM/PM")
    End With
    Application.StatusBar = False: Application.DisplayAlerts = True: Application.EnableEvents = True: Application.ScreenUpdating = True
    MsgBox "Catalog refreshed. Tracker entries and custom hacks were preserved.", vbInformation
    Exit Sub
Failed:
    Application.StatusBar = False: Application.DisplayAlerts = True: Application.EnableEvents = True: Application.ScreenUpdating = True
    MsgBox "Refresh failed:" & vbCrLf & Err.Description, vbExclamation
End Sub

Private Sub DownloadCatalog(ByVal url As String)
    Dim http As Object, stream As Object, tempPath As String
    Dim csvBook As Workbook, csvSheet As Worksheet, target As Worksheet, lastRow As Long
    tempPath = Environ$("TEMP") & "\SMWCentral_All_Moderated_Hacks.csv"
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False: http.setRequestHeader "Cache-Control", "no-cache": http.send
    If http.Status < 200 Or http.Status >= 300 Then Err.Raise vbObjectError + 1000, , "HTTP status " & http.Status
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1: stream.Open: stream.Write http.responseBody: stream.SaveToFile tempPath, 2: stream.Close
    Workbooks.OpenText Filename:=tempPath, Origin:=65001, DataType:=xlDelimited, Comma:=True, TextQualifier:=xlTextQualifierDoubleQuote
    Set csvBook = ActiveWorkbook: Set csvSheet = csvBook.Worksheets(1): Set target = ThisWorkbook.Worksheets(ONLINE)
    target.Range("A1:J10000").ClearContents
    lastRow = csvSheet.Cells(csvSheet.Rows.Count, "A").End(xlUp).Row
    target.Range("A1").Resize(lastRow, 10).Value = csvSheet.Range("A1").Resize(lastRow, 10).Value
    csvBook.Close False
    On Error Resume Next: Kill tempPath: On Error GoTo 0
End Sub

Private Sub RebuildCombined()
    Dim onWs As Worksheet, manWs As Worksheet, outWs As Worksheet, dict As Object
    Dim data As Variant, output() As Variant, i As Long, c As Long, n As Long, key As String, maxRows As Long
    Set onWs = ThisWorkbook.Worksheets(ONLINE): Set manWs = ThisWorkbook.Worksheets(MANUAL): Set outWs = ThisWorkbook.Worksheets(COMBINED)
    Set dict = CreateObject("Scripting.Dictionary")
    maxRows = Application.Max(onWs.Cells(onWs.Rows.Count, "A").End(xlUp).Row - 1, 0) + Application.Max(manWs.Cells(manWs.Rows.Count, "A").End(xlUp).Row - 1, 0)
    If maxRows = 0 Then Exit Sub
    ReDim output(1 To maxRows, 1 To 10)
    outWs.Range("A2:J10000").ClearContents: outWs.Range("A1:J1").Value = onWs.Range("A1:J1").Value
    data = ReadRows(onWs): AppendUnique data, dict, output, n
    data = ReadRows(manWs): AppendUnique data, dict, output, n
    If n > 0 Then outWs.Range("A2").Resize(n, 10).Value = output
End Sub

Private Function ReadRows(ByVal ws As Worksheet) As Variant
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then ReadRows = Empty Else ReadRows = ws.Range("A2:J" & lastRow).Value
End Function

Private Sub AppendUnique(ByVal data As Variant, ByVal dict As Object, ByRef output As Variant, ByRef n As Long)
    Dim i As Long, c As Long, key As String
    If IsEmpty(data) Then Exit Sub
    For i = 1 To UBound(data, 1)
        key = Trim$(CStr(data(i, 8)))
        If Len(key) = 0 Then key = "LABEL:" & UCase$(Trim$(CStr(data(i, 1)))) Else key = UCase$(key)
        If Len(key) > 0 And Not dict.Exists(key) Then
            dict.Add key, True: n = n + 1
            For c = 1 To 10: output(n, c) = data(i, c): Next c
        End If
    Next i
End Sub

Private Sub UpdateLists()
    Dim db As Worksheet, ls As Worksheet, lastRow As Long
    Set db = ThisWorkbook.Worksheets(COMBINED): Set ls = ThisWorkbook.Worksheets(LISTS)
    lastRow = db.Cells(db.Rows.Count, "A").End(xlUp).Row
    ls.Range("A2:A10000").ClearContents
    If lastRow >= 2 Then ls.Range("A2").Resize(lastRow - 1, 1).Value = db.Range("A2:A" & lastRow).Value
    On Error Resume Next: ThisWorkbook.Names("HackList").Delete: On Error GoTo 0
    ThisWorkbook.Names.Add Name:="HackList", RefersTo:="='Lists'!$A$2:INDEX('Lists'!$A:$A,COUNTA('Lists'!$A:$A))"
End Sub

Public Sub AddCustomHack()
    Dim f As Worksheet, m As Worksheet, nextRow As Long, title As String, label As String, id As String
    Set f = ThisWorkbook.Worksheets(DASH): Set m = ThisWorkbook.Worksheets(MANUAL)
    title = Trim$(CStr(f.Range("J13").Value))
    If Len(title) = 0 Then MsgBox "Enter a title first.", vbExclamation: Exit Sub
    label = title
    If Not IsError(Application.Match(label, ThisWorkbook.Worksheets(COMBINED).Range("A:A"), 0)) Then label = title & " [Manual]"
    id = "MANUAL-" & Format(Now, "yyyymmddhhnnss")
    nextRow = m.Cells(m.Rows.Count, "A").End(xlUp).Row + 1: If nextRow < 2 Then nextRow = 2
    m.Cells(nextRow, 1).Value = label: m.Cells(nextRow, 2).Value = title: m.Cells(nextRow, 3).Value = f.Range("J14").Value
    m.Cells(nextRow, 4).Value = f.Range("J15").Value: m.Cells(nextRow, 5).Value = f.Range("J16").Value: m.Cells(nextRow, 6).Value = f.Range("J17").Value
    m.Cells(nextRow, 7).Value = f.Range("J18").Value: m.Cells(nextRow, 8).Value = id: m.Cells(nextRow, 9).Value = f.Range("J19").Value: m.Cells(nextRow, 10).Value = f.Range("J20").Value
    RebuildCombined: UpdateLists: f.Range("J13:K20").ClearContents
    MsgBox "Custom hack added. It will remain after future refreshes.", vbInformation
End Sub
