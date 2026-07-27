Attribute VB_Name = "SMWCommunity"
Option Explicit

Private Const DASHBOARD_SHEET As String = "Dashboard"
Private Const TRACKER_SHEET As String = "Tracker"
Private Const ONLINE_SHEET As String = "Online Database"
Private Const MANUAL_SHEET As String = "Manual Database"
Private Const DATABASE_SHEET As String = "Hack Database"
Private Const LISTS_SHEET As String = "Lists"
Private Const SETTINGS_SHEET As String = "Settings"
Private Const PROTECTION_PASSWORD As String = "SMWCommunity2026"

Public Sub InstallButtons()
    Dim ws As Worksheet
    Dim shp As Shape

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
End Sub

Public Sub RefreshHacks()
    On Error GoTo Failed

    Dim url As String
    url = Trim$(CStr(ThisWorkbook.Worksheets(SETTINGS_SHEET).Range("B2").Value))

    If Len(url) = 0 Or InStr(1, url, "raw.githubusercontent.com", vbTextCompare) = 0 Then
        MsgBox "Catalog CSV URL is missing or invalid in Settings!B2.", vbExclamation, "Refresh Hacks"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Downloading SMWCentral catalog..."

    DownloadCatalogToOnlineDatabase url
    RebuildCombinedDatabase
    UpdateDropdownList

    With ThisWorkbook.Worksheets(SETTINGS_SHEET)
        .Range("B5").Value = LastUsedRow(ThisWorkbook.Worksheets(ONLINE_SHEET), "A") - 1
        .Range("B7").Value = Now
        .Range("B7").NumberFormat = "mmm d, yyyy h:mm AM/PM"
    End With

    With ThisWorkbook.Worksheets(DASHBOARD_SHEET)
        .Range("J9").Value = Now
        .Range("J9").NumberFormat = "mmm d, yyyy h:mm AM/PM"
    End With

    PrepareWorkbookForRelease

    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "SMWCentral catalog refreshed. Tracker entries and custom hacks were preserved.", vbInformation, "Refresh Complete"
    Exit Sub

Failed:
    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Refresh failed:" & vbCrLf & Err.Description, vbExclamation, "Refresh Hacks"
End Sub

Private Sub DownloadCatalogToOnlineDatabase(ByVal url As String)
    Dim http As Object
    Dim stream As Object
    Dim tempPath As String
    Dim csvBook As Workbook
    Dim csvSheet As Worksheet
    Dim target As Worksheet
    Dim lastRow As Long

    tempPath = Environ$("TEMP") & "\SMWCentral_All_Moderated_Hacks.csv"

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.setRequestHeader "Cache-Control", "no-cache"
    http.send

    If http.Status < 200 Or http.Status >= 300 Then
        Err.Raise vbObjectError + 1000, , "HTTP " & http.Status & " while downloading catalog."
    End If

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile tempPath, 2
    stream.Close

    Workbooks.OpenText Filename:=tempPath, Origin:=65001, _
        DataType:=xlDelimited, Comma:=True, TextQualifier:=xlTextQualifierDoubleQuote

    Set csvBook = ActiveWorkbook
    Set csvSheet = csvBook.Worksheets(1)
    Set target = ThisWorkbook.Worksheets(ONLINE_SHEET)

    target.Range("A1:J10000").ClearContents
    lastRow = csvSheet.Cells(csvSheet.Rows.Count, "A").End(xlUp).Row
    target.Range("A1").Resize(lastRow, 10).Value = csvSheet.Range("A1").Resize(lastRow, 10).Value

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

    dbWs.Range("A1:J10000").ClearContents
    dbWs.Range("A1:J1").Value = onlineWs.Range("A1:J1").Value

    lastOnline = LastUsedRow(onlineWs, "A")
    lastManual = LastUsedRow(manualWs, "A")
    maxRows = Application.Max(0, lastOnline - 1) + Application.Max(0, lastManual - 1)

    If maxRows = 0 Then Exit Sub

    ReDim output(1 To maxRows, 1 To 10)
    outRow = 0

    If lastOnline >= 2 Then
        sourceData = onlineWs.Range("A2:J" & lastOnline).Value
        For r = 1 To UBound(sourceData, 1)
            key = MakeCatalogKey(sourceData, r)
            If Len(key) > 0 And Not dict.Exists(key) Then
                dict.Add key, True
                outRow = outRow + 1
                For c = 1 To 10
                    output(outRow, c) = sourceData(r, c)
                Next c
            End If
        Next r
    End If

    If lastManual >= 2 Then
        sourceData = manualWs.Range("A2:J" & lastManual).Value
        For r = 1 To UBound(sourceData, 1)
            key = MakeCatalogKey(sourceData, r)
            If Len(key) > 0 And Not dict.Exists(key) Then
                dict.Add key, True
                outRow = outRow + 1
                For c = 1 To 10
                    output(outRow, c) = sourceData(r, c)
                Next c
            End If
        Next r
    End If

    If outRow > 0 Then
        dbWs.Range("A2").Resize(outRow, 10).Value = output
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
    End With

    RebuildCombinedDatabase
    UpdateDropdownList
    formWs.Range("C21:H27").ClearContents

    MsgBox "Custom hack added. It will not be removed by Refresh Hacks.", vbInformation, "Add Custom Hack"
    Exit Sub

Failed:
    MsgBox "Custom hack could not be added:" & vbCrLf & Err.Description, vbExclamation, "Add Custom Hack"
End Sub

Public Sub PrepareWorkbookForRelease()
    Dim ws As Worksheet

    On Error Resume Next
    ThisWorkbook.Unprotect Password:=PROTECTION_PASSWORD
    On Error GoTo 0

    For Each ws In ThisWorkbook.Worksheets
        Select Case ws.Name
            Case DASHBOARD_SHEET, TRACKER_SHEET, DATABASE_SHEET
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
