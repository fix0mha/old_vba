Attribute VB_Name = "modUI"
Option Explicit
'===============================================================
' modUI - Oberflaeche auf Arbeitsblaettern
'   'alle Projekte'     Projektliste (Kurzbeschreibung/Bearbeiter editierbar)
'   'Projekt'           Felder des geoeffneten Projekts
'   'Felder anpassen'   globale Feldliste (F-, S-, B-Felder)
'   'Admin Center'      Datenpfad, wiederherstellen, loeschen, Selbsttest, purge
' Setup_Erstellen baut Blaetter und Buttons automatisch auf.
'===============================================================
Private Const OLD_SHEET_START As String = "Start"
Private Const OLD_SHEET_FELDER As String = "Felder bearbeiten|Projekt bearbeiten"
Private Const START_FIRST_ROW As Long = 9
Private Const PROJ_FIRST_ROW As Long = 9
Private Const FELD_FIRST_ROW As Long = 7
Private Const MAX_ROWS As Long = 5000
Private Const MAX_CELL_CHARS As Long = 32767
Private Const BTN_SAVE_LIST As String = "btnProjektSpeichern"
Private Const COLOR_OWN As Long = &HCCFFCC          ' hellgruen: eigene Sperre

Private Const COLOR_CHANGED As Long = &H99CCFF      ' hellorange (BGR)
Private Const COLOR_READONLY As Long = &HF2F2F2     ' hellgrau
Private Const COLOR_ERROR As Long = &H9696FF        ' rot

' Stand der Projektliste (Schluessel = UCase(Ticket ID))
Private mOrig As Object     ' -> Array(Kurz, Bearb, Angelegt, AngelegtVon, Ordner, Gesperrt)
Private mRowId As Object    ' Zeilennummer (Text) -> Ticket ID
Private mListEnd As Long

' Stand des geoeffneten Projekts im Blatt 'Projekt' (Schluessel = UCase(Field ID) -> Wert)
Private mPOrig As Object

' Stand der Feldliste beim Laden (Schluessel = UCase(Field ID))
Private mFKey As Object
Private mFDesc As Object
Private mFLoaded As Boolean

'===============================================================
' SETUP
'===============================================================
Public Sub Setup_Erstellen()
    Dim wsS As Worksheet, wsP As Worksheet, wsF As Worksheet, wsA As Worksheet, root As String
    On Error GoTo EH
    If MsgBox("Das Setup erstellt bzw. erneuert die Blätter '" & SHEET_START & "', '" & SHEET_PROJEKT & "', '" & SHEET_FELDER & "' und '" & SHEET_ADMIN & "'." & vbCrLf & _
              "Daten auf dem Netzlaufwerk bleiben unberührt." & vbCrLf & vbCrLf & "Fortfahren?", vbYesNo + vbQuestion, APP_TITLE) <> vbYes Then Exit Sub
    Application.ScreenUpdating = False
    Set wsS = SheetOrNew(SHEET_START, OLD_SHEET_START)
    Set wsP = SheetOrNew(SHEET_PROJEKT, "")
    Set wsF = SheetOrNew(SHEET_FELDER, OLD_SHEET_FELDER)
    Set wsA = SheetOrNew(SHEET_ADMIN, "")
    On Error Resume Next
    root = Trim$(CStr(wsA.Range(CELL_DATAROOT).Value))
    If Len(root) = 0 Then root = Trim$(CStr(wsS.Range("C3").Value))
    On Error GoTo EH
    BuildStart wsS
    BuildProjekt wsP
    BuildFelder wsF
    BuildAdmin wsA, root
    mFLoaded = False
    Set mOrig = Nothing
    mListEnd = 0
    HookEvents
    wsA.Activate
    wsA.Range("C3").Select
    Application.ScreenUpdating = True
    MsgBox "Setup abgeschlossen." & vbCrLf & vbCrLf & _
           "Bitte jetzt im Blatt '" & SHEET_ADMIN & "' den Datenpfad in Zelle C3 eintragen" & vbCrLf & _
           "(oder Button 'Datenpfad wählen'), danach im Blatt '" & SHEET_START & "' 'Aktualisieren'" & vbCrLf & _
           "und im Admin Center den 'Selbsttest' ausführen.", vbInformation, APP_TITLE
    Exit Sub
EH:
    HandleError "Setup_Erstellen"
End Sub

' Liefert das Blatt nm; benennt sonst ein altes oder ein leeres Standardblatt um
' oder legt ein neues an.
Private Function SheetOrNew(ByVal nm As String, ByVal oldNm As String) As Worksheet
    Dim ws As Worksheet, o As Variant
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, nm, vbTextCompare) = 0 Then
            Set SheetOrNew = ws
            Exit Function
        End If
    Next ws
    If Len(oldNm) > 0 Then
        For Each o In Split(oldNm, "|")
            For Each ws In ThisWorkbook.Worksheets
                If StrComp(ws.Name, CStr(o), vbTextCompare) = 0 Then
                    ws.Name = nm
                    Set SheetOrNew = ws
                    Exit Function
                End If
            Next ws
        Next o
    End If
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "Tabelle#*" Or ws.Name Like "Sheet#*" Then
            If Application.WorksheetFunction.CountA(ws.Cells) = 0 And ws.Shapes.Count = 0 Then
                ws.Name = nm
                Set SheetOrNew = ws
                Exit Function
            End If
        End If
    Next ws
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = nm
    Set SheetOrNew = ws
End Function

Private Sub ClearSheet(ByVal ws As Worksheet)
    ws.Cells.Clear
    Do While ws.Shapes.Count > 0
        ws.Shapes(1).Delete
    Loop
    ws.Rows.RowHeight = ws.StandardHeight
End Sub

Private Sub AddButton(ByVal ws As Worksheet, ByVal caption As String, ByVal macroName As String, _
                      ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, _
                      Optional ByVal shapeName As String = "")
    Dim b As Object
    Set b = ws.Buttons.Add(x, y, w, h)
    b.Caption = caption
    b.OnAction = macroName
    On Error Resume Next
    b.Font.Size = 9
    b.Placement = 3   ' xlFreeFloating
    If Len(shapeName) > 0 Then b.Name = shapeName
End Sub

Private Sub StyleHeader(ByVal rng As Range)
    rng.Font.Bold = True
    rng.Interior.Color = RGB(221, 235, 247)
    rng.Borders(9).LineStyle = 1   ' xlEdgeBottom, xlContinuous
End Sub

Private Sub BuildStart(ByVal ws As Worksheet)
    Dim x As Double, y As Double
    ClearSheet ws
    ws.Columns("A").ColumnWidth = 26
    ws.Columns("B").ColumnWidth = 44
    ws.Columns("C").ColumnWidth = 28
    ws.Columns("D").ColumnWidth = 20
    ws.Columns("E").ColumnWidth = 24
    ws.Columns("F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 40
    ws.Range("A1").Value = APP_TITLE
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").Font.Bold = True
    ws.Rows(1).RowHeight = 26
    ws.Range("A2").Value = "Version " & APP_VERSION
    ws.Range("A3").Value = "Datenpfad (siehe Admin Center):"
    ws.Range("A3").Font.Bold = True
    ws.Range("B3").Formula = "='" & SHEET_ADMIN & "'!" & CELL_DATAROOT
    ws.Rows(5).RowHeight = 28
    ws.Rows(7).RowHeight = 28
    x = ws.Range("A5").Left + 2
    y = ws.Range("A5").Top + 2
    AddButton ws, "Neues Projekt", "Start_NeuesProjekt", x, y, 110, 24
    AddButton ws, "Projekt öffnen", "Start_Oeffnen", x + 116, y, 110, 24
    AddButton ws, "Projekt entfernen", "Start_Delete", x + 232, y, 120, 24
    AddButton ws, "Update Projekte", "Start_UpdateProjekte", x + 358, y, 120, 24
    AddButton ws, "Projekt speichern", "Liste_Speichern", x + 484, y, 130, 24, BTN_SAVE_LIST
    AddButton ws, "Aktualisieren", "Start_Aktualisieren", ws.Range("A7").Left + 2, ws.Range("A7").Top + 2, 110, 24
    ws.Range("A8").Value = "Ticket ID"
    ws.Range("B8").Value = KEY_KURZ
    ws.Range("C8").Value = KEY_BEARB
    ws.Range("D8").Value = "Angelegt"
    ws.Range("E8").Value = "Angelegt von"
    ws.Range("F8").Value = "Ordner"
    ws.Range("G8").Value = "Gesperrt von"
    StyleHeader ws.Range("A8:G8")
    ws.Range("A" & START_FIRST_ROW & ":G" & MAX_ROWS).NumberFormat = "@"
    ws.Range("A" & START_FIRST_ROW & ":A" & MAX_ROWS).Interior.Color = COLOR_READONLY
    ws.Range("D" & START_FIRST_ROW & ":G" & MAX_ROWS).Interior.Color = COLOR_READONLY
    ws.Range("A4").Value = "Kurzbeschreibung und Bearbeiter sind hier änderbar (orange = ungespeichert, dann 'Projekt speichern'). Die übrigen Spalten sind schreibgeschützt."
    ws.Range("A4").Font.Italic = True
    ws.Range("A4").Font.Color = RGB(100, 100, 100)
    On Error Resume Next
    ws.Shapes(BTN_SAVE_LIST).Visible = False
End Sub

Private Sub BuildProjekt(ByVal ws As Worksheet)
    Dim x As Double, y As Double, i As Long, names As Variant
    ClearSheet ws
    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 80
    ws.Columns("D").ColumnWidth = 16
    ws.Range("A1:A2").NumberFormat = "@"
    ws.Range("A1").Value = "(kein Projekt geöffnet)"
    ws.Range("A1:A2").Font.Size = 14
    ws.Range("A1:A2").Font.Bold = True
    ws.Rows(1).RowHeight = 22
    ws.Rows(2).RowHeight = 22
    ws.Rows(4).RowHeight = 28
    ws.Rows(5).RowHeight = 28
    x = ws.Range("A4").Left + 2
    y = ws.Range("A4").Top + 2
    AddButton ws, "Speichern", "Projekt_Speichern", x, y, 110, 24
    AddButton ws, "Schließen", "Projekt_Schliessen", x + 116, y, 110, 24
    names = Array("1 Dokument B einlesen", "2 Vorlagen kopieren", "3 Dokument A füllen", "4 Score füllen", "5 SH füllen")
    y = ws.Range("A5").Top + 2
    For i = 0 To 4
        AddButton ws, CStr(names(i)), "Aufgabe_" & (i + 1), x + i * 140, y, 134, 24
    Next i
    ws.Range("A6").Value = "Aufgaben-Status: -"
    ws.Range("A3").Font.Bold = True
    ws.Range("A3").Font.Color = RGB(204, 102, 0)
    ws.Range("A8").Value = "Field ID"
    ws.Range("B8").Value = "Name"
    ws.Range("C8").Value = "Wert"
    ws.Range("D8").Value = "Quelle"
    StyleHeader ws.Range("A8:D8")
    ws.Range("A" & PROJ_FIRST_ROW & ":D" & MAX_ROWS).NumberFormat = "@"
    ws.Range("A" & PROJ_FIRST_ROW & ":D" & MAX_ROWS).VerticalAlignment = -4160   ' xlTop
    ws.Columns("C").WrapText = True
End Sub

Private Sub BuildFelder(ByVal ws As Worksheet)
    Dim x As Double, y As Double
    ClearSheet ws
    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 40
    ws.Columns("C").ColumnWidth = 60
    ws.Columns("D").ColumnWidth = 16
    ws.Range("A1").Value = SHEET_FELDER
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    ws.Rows(1).RowHeight = 22
    ws.Range("A2").Value = "F = lokale Felder (änderbar), S = aus System und B = aus BeKo (schreibgeschützt). Zeile löschen oder Button 'Feld deaktivieren' = Feld deaktivieren. Button 'Neues Feld' oder neue Zeile ohne Field ID = neues lokales Feld."
    ws.Range("A3").Value = "Jeder Name (Key) muss eindeutig sein. '" & KEY_KURZ & "' und '" & KEY_BEARB & "' sind Standardfelder. Endgültig löschen: Admin Center, 'purge inactive Fields'."
    ws.Rows(4).RowHeight = 28
    x = ws.Range("A4").Left + 2
    y = ws.Range("A4").Top + 2
    AddButton ws, "Neu laden", "Felder_Laden", x, y, 110, 24
    AddButton ws, "Speichern", "Felder_Speichern", x + 116, y, 110, 24
    AddButton ws, "Neues Feld", "Felder_Neu", x + 232, y, 110, 24
    AddButton ws, "Feld deaktivieren", "Felder_Deaktivieren", x + 348, y, 130, 24
    ws.Range("A6").Value = "Field ID"
    ws.Range("B6").Value = "Name"
    ws.Range("C6").Value = "Beschreibung"
    ws.Range("D6").Value = "Quelle"
    StyleHeader ws.Range("A6:D6")
    ws.Range("A" & FELD_FIRST_ROW & ":D" & MAX_ROWS).NumberFormat = "@"
End Sub

Private Sub BuildAdmin(ByVal ws As Worksheet, ByVal root As String)
    Dim x As Double, y As Double
    ClearSheet ws
    ws.Columns("A").ColumnWidth = 28
    ws.Columns("B").ColumnWidth = 20
    ws.Columns("C").ColumnWidth = 70
    ws.Range("A1").Value = SHEET_ADMIN
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").Font.Bold = True
    ws.Rows(1).RowHeight = 26
    ws.Range("A3").Value = "Datenpfad (Netzlaufwerk):"
    ws.Range("A3").Font.Bold = True
    ws.Range("C3").NumberFormat = "@"
    ws.Range("C3").Value = root
    ws.Range("C3").Interior.Color = RGB(255, 249, 196)
    ws.Rows(3).RowHeight = 28
    ws.Rows(5).RowHeight = 28
    ws.Rows(6).RowHeight = 28
    ws.Rows(7).RowHeight = 28
    x = ws.Range("A5").Left + 2
    y = ws.Range("A5").Top + 2
    AddButton ws, "Projekt wiederherstellen", "Start_Verknuepfen", x, y, 170, 24
    AddButton ws, "Projekt vollständig löschen", "Start_Purge", x + 176, y, 170, 24
    y = ws.Range("A6").Top + 2
    AddButton ws, "Selbsttest", "Start_Selbsttest", x, y, 170, 24
    AddButton ws, "Datenpfad wählen", "Start_DatenpfadWaehlen", x + 176, y, 170, 24
    y = ws.Range("A7").Top + 2
    AddButton ws, "purge inactive Fields", "Admin_PurgeInaktiveFelder", x, y, 170, 24
    ws.Range("A9").Value = "Projekt wiederherstellen: stellt den Verweis zu einem vorhandenen Projektordner wieder her (z. B. nach 'Projekt entfernen')."
    ws.Range("A10").Value = "Projekt vollständig löschen: löscht Verweis UND alle Projektdaten endgültig (nicht rückgängig zu machen)."
    ws.Range("A11").Value = "Selbsttest: prüft JSON, Dateizugriff, Sperren, Feldliste und die Word-Automation auf diesem Rechner."
    ws.Range("A12").Value = "Datenpfad wählen: legt den Datenordner auf dem Netzlaufwerk fest (Zelle C3)."
    ws.Range("A13").Value = "purge inactive Fields: löscht deaktivierte Felder (F, S, B) endgültig samt Werten in allen Projekten."
End Sub

'===============================================================
' ALLE PROJEKTE: Projektliste
'===============================================================
Private Function SelectedProjectId(ByVal prompt As String) As String
    Dim ws As Worksheet, s As String, hint As String
    Set ws = ThisWorkbook.Worksheets(SHEET_START)
    If ActiveSheet.Name = SHEET_START Then
        If ActiveCell.Row >= START_FIRST_ROW Then s = Trim$(CStr(ws.Cells(ActiveCell.Row, 1).Value))
        hint = vbCrLf & vbCrLf & "(Tipp: vorher eine Zeile in der Liste anklicken)"
    End If
    If Len(s) = 0 Then
        s = Trim$(InputBox(prompt & hint, APP_TITLE))
    End If
    SelectedProjectId = s
End Function

' Anzahl der Zeilen mit ungespeicherten Aenderungen (Kurzbeschreibung/Bearbeiter).
Public Function DirtyRowCount() As Long
    Dim ws As Worksheet, v As Variant, r As Long, n As Long, o As Variant, key As String
    If mOrig Is Nothing Then Exit Function
    If mListEnd < START_FIRST_ROW Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SHEET_START)
    v = ws.Range(ws.Cells(START_FIRST_ROW, 1), ws.Cells(mListEnd, 3)).Value
    For r = 1 To UBound(v, 1)
        key = UCase$(Trim$(CStr(v(r, 1))))
        If mOrig.Exists(key) Then
            o = mOrig(key)
            If CStr(v(r, 2)) <> CStr(o(0)) Or CStr(v(r, 3)) <> CStr(o(1)) Then n = n + 1
        End If
    Next r
    DirtyRowCount = n
End Function

' Fragt nach, wenn ungespeicherte Listenaenderungen verloren gingen.
Public Function ConfirmDiscardListChanges() As Boolean
    Dim n As Long
    n = DirtyRowCount()
    If n = 0 Then
        ConfirmDiscardListChanges = True
    Else
        ConfirmDiscardListChanges = (MsgBox("In der Projektliste gibt es " & n & " ungespeicherte Änderung(en)." & vbCrLf & _
            "Sie gehen verloren, wenn Sie fortfahren." & vbCrLf & vbCrLf & "Änderungen verwerfen und fortfahren?", _
            vbYesNo + vbExclamation + vbDefaultButton2, APP_TITLE) = vbYes)
    End If
End Function

Private Sub UpdateSaveButton()
    Dim vis As Boolean
    vis = (DirtyRowCount() > 0)
    On Error Resume Next
    ThisWorkbook.Worksheets(SHEET_START).Shapes(BTN_SAVE_LIST).Visible = vis
End Sub

Public Sub ListeNeuLaden()
    Dim ws As Worksheet, col As Collection, d As Object, n As Long, i As Long, arr() As Variant
    Dim id As String, errNum As Long, errSrc As String, errDesc As String
    Set ws = ThisWorkbook.Worksheets(SHEET_START)
    InitDataRoot
    LoadFieldsDoc
    Set col = ListProjects()
    n = col.Count
    On Error GoTo EH
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ws.Range("A" & START_FIRST_ROW & ":G" & MAX_ROWS).ClearContents
    ws.Range("B" & START_FIRST_ROW & ":C" & MAX_ROWS).Interior.ColorIndex = -4142   ' xlNone
    Set mOrig = NewDict()
    Set mRowId = NewDict()
    If n > 0 Then
        ReDim arr(1 To n, 1 To 7)
        i = 0
        For Each d In col
            i = i + 1
            id = CStr(d("id"))
            arr(i, 1) = id
            arr(i, 2) = CStr(d("kurzbeschreibung"))
            arr(i, 3) = CStr(d("bearbeiter"))
            arr(i, 4) = CStr(d("angelegt"))
            arr(i, 5) = CStr(d("angelegt_von"))
            arr(i, 6) = CStr(d("ordner"))
            arr(i, 7) = CStr(d("gesperrt"))
            mOrig(UCase$(id)) = Array(arr(i, 2), arr(i, 3), arr(i, 4), arr(i, 5), arr(i, 6), arr(i, 7))
            mRowId(CStr(START_FIRST_ROW + i - 1)) = id
        Next d
        ws.Range(ws.Cells(START_FIRST_ROW, 1), ws.Cells(START_FIRST_ROW + n - 1, 7)).Value = arr
    End If
    mListEnd = START_FIRST_ROW + n - 1
    ws.Range("G" & START_FIRST_ROW & ":G" & MAX_ROWS).Interior.Color = COLOR_READONLY
    For i = 1 To n
        If Left$(CStr(arr(i, 7)), 4) = "Ich " Then ws.Cells(START_FIRST_ROW + i - 1, 7).Interior.Color = COLOR_OWN
    Next i
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    UpdateSaveButton
    Exit Sub
EH:
    errNum = Err.Number: errSrc = Err.Source: errDesc = Err.Description
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Err.Raise errNum, errSrc, errDesc
End Sub

Public Sub Start_Aktualisieren()
    On Error GoTo EH
    EnsureHooked
    If Not ConfirmDiscardListChanges() Then Exit Sub
    ListeNeuLaden
    Exit Sub
EH:
    HandleError "Start_Aktualisieren"
End Sub

'--- Aenderungen in der Liste (Aufruf aus clsSheetEvents) ---------
Public Sub List_Changed(ByVal Target As Range)
    Dim ws As Worksheet, rng As Range, area As Range, r As Long, rows As Object, k As Variant
    Dim restored As Boolean
    On Error GoTo EH
    If mOrig Is Nothing Then Exit Sub
    If mListEnd < START_FIRST_ROW Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SHEET_START)
    Set rng = Intersect(Target, ws.Range("A" & START_FIRST_ROW & ":G" & mListEnd))
    If rng Is Nothing Then Exit Sub
    Set rows = NewDict()
    For Each area In rng.Areas
        For r = area.Row To area.Row + area.Rows.Count - 1
            rows(CStr(r)) = True
        Next r
    Next area
    Application.EnableEvents = False
    For Each k In rows.Keys
        If ProcessRow(ws, CLng(k)) Then restored = True
    Next k
    Application.EnableEvents = True
    UpdateSaveButton
    If restored Then MsgBox "Die Ticket ID kann nicht geändert werden und wurde zurückgesetzt.", vbInformation, APP_TITLE
    Exit Sub
EH:
    HandleError "List_Changed"
End Sub

' Prueft eine Zeile; True, wenn die Ticket ID zurueckgesetzt wurde.
Private Function ProcessRow(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    Dim v As Variant, id As String, back As String, key As String, o As Variant, c As Long
    Dim bad As Boolean
    v = ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).Value
    id = Trim$(CStr(v(1, 1)))
    key = UCase$(id)
    If mRowId.Exists(CStr(r)) Then back = CStr(mRowId(CStr(r)))
    bad = (Len(id) = 0) Or Not mOrig.Exists(key)
    If Not bad Then
        If Application.WorksheetFunction.CountIf(ws.Range("A" & START_FIRST_ROW & ":A" & mListEnd), id) > 1 Then bad = True
    End If
    If bad Then
        If Len(back) = 0 Or Not mOrig.Exists(UCase$(back)) Then Exit Function
        ws.Cells(r, 1).Value = back
        id = back
        key = UCase$(id)
        v = ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).Value
        ProcessRow = True
    End If
    o = mOrig(key)
    mRowId(CStr(r)) = id
    For c = 4 To 7
        If CStr(v(1, c)) <> CStr(o(c - 2)) Then ws.Cells(r, c).Value = o(c - 2)
    Next c
    For c = 2 To 3
        If CStr(v(1, c)) <> CStr(o(c - 2)) Then
            ws.Cells(r, c).Interior.Color = COLOR_CHANGED
        Else
            ws.Cells(r, c).Interior.ColorIndex = -4142
        End If
    Next c
End Function

'--- Button 'Projekt speichern' -----------------------------------
Public Sub Liste_Speichern()
    Dim ws As Worksheet, v As Variant, r As Long, key As String, o As Variant, id As String
    Dim kurz As String, bearb As String, res As Long, info As String, ans As VbMsgBoxResult
    Dim nOk As Long, nUnch As Long, failed As Collection, msg As String, cancelled As Boolean
    Dim origK As String, origB As String
    On Error GoTo EH
    If mOrig Is Nothing Then Exit Sub
    If DirtyRowCount() = 0 Then
        UpdateSaveButton
        MsgBox "Es gibt keine ungespeicherten Änderungen.", vbInformation, APP_TITLE
        Exit Sub
    End If
    Set ws = ThisWorkbook.Worksheets(SHEET_START)
    Set failed = New Collection
    v = ws.Range(ws.Cells(START_FIRST_ROW, 1), ws.Cells(mListEnd, 3)).Value
    For r = 1 To UBound(v, 1)
        key = UCase$(Trim$(CStr(v(r, 1))))
        If mOrig.Exists(key) And Not cancelled Then
            o = mOrig(key)
            If CStr(v(r, 2)) <> CStr(o(0)) Or CStr(v(r, 3)) <> CStr(o(1)) Then
                id = Trim$(CStr(v(r, 1)))
                kurz = Trim$(CStr(v(r, 2)))
                bearb = Trim$(CStr(v(r, 3)))
                origK = CStr(o(0))
                origB = CStr(o(1))
                res = SetStandardFields(id, kurz, bearb, origK, origB, False, info)
                If res = PR_CONFLICT Then
                    ans = MsgBox("Projekt " & id & " " & info & "." & vbCrLf & vbCrLf & _
                                 "Ja = meine Änderung trotzdem speichern" & vbCrLf & _
                                 "Nein = dieses Projekt überspringen" & vbCrLf & _
                                 "Abbrechen = Rest nicht mehr speichern", vbYesNoCancel + vbQuestion, APP_TITLE)
                    If ans = vbYes Then
                        res = SetStandardFields(id, kurz, bearb, origK, origB, True, info)
                    ElseIf ans = vbNo Then
                        failed.Add id & ": übersprungen (geändert)"
                        res = PR_CONFLICT
                    Else
                        cancelled = True
                        res = PR_CONFLICT
                    End If
                End If
                Select Case res
                    Case PR_OK, PR_UNCHANGED
                        If res = PR_OK Then nOk = nOk + 1 Else nUnch = nUnch + 1
                        mOrig(key) = Array(kurz, bearb, o(2), o(3), o(4), o(5))
                        Application.EnableEvents = False
                        ws.Cells(START_FIRST_ROW + r - 1, 2).Value = kurz
                        ws.Cells(START_FIRST_ROW + r - 1, 3).Value = bearb
                        ws.Cells(START_FIRST_ROW + r - 1, 2).Interior.ColorIndex = -4142
                        ws.Cells(START_FIRST_ROW + r - 1, 3).Interior.ColorIndex = -4142
                        Application.EnableEvents = True
                    Case PR_LOCKED, PR_ERROR
                        failed.Add id & ": " & info
                End Select
            End If
        End If
    Next r
    UpdateSaveButton
    msg = nOk & " Projekt(e) gespeichert."
    If nUnch > 0 Then msg = msg & vbCrLf & nUnch & " ohne Änderung (Wert war schon gespeichert)."
    If cancelled Then msg = msg & vbCrLf & "Vorgang abgebrochen."
    If failed.Count > 0 Then
        msg = msg & vbCrLf & vbCrLf & failed.Count & " nicht gespeichert (bleiben orange markiert):" & vbCrLf & JoinLines(failed, 8)
        For r = 1 To failed.Count
            WriteLog "Liste nicht gespeichert: " & CStr(failed(r))
        Next r
    End If
    MsgBox msg, vbInformation, APP_TITLE
    Exit Sub
EH:
    HandleError "Liste_Speichern"
End Sub

'--- Buttons ------------------------------------------------------
Public Sub Start_DatenpfadWaehlen()
    Dim fd As Object
    On Error GoTo EH
    If IsProjectOpen() Then UserError "Bitte zuerst das geöffnete Projekt schließen."
    If Not ConfirmDiscardListChanges() Then Exit Sub
    Set fd = Application.FileDialog(4)   ' msoFileDialogFolderPicker
    fd.Title = "Datenordner auf dem Netzlaufwerk wählen"
    If fd.Show = -1 Then
        ThisWorkbook.Worksheets(SHEET_ADMIN).Range(CELL_DATAROOT).Value = fd.SelectedItems(1)
        mFLoaded = False
        ListeNeuLaden
    End If
    Exit Sub
EH:
    HandleError "Start_DatenpfadWaehlen"
End Sub

Public Sub Start_NeuesProjekt()
    Dim id As String, kurz As String
    On Error GoTo EH
    If Not ConfirmDiscardListChanges() Then Exit Sub
    id = Trim$(InputBox("Neue Ticket ID:" & vbCrLf & "Erlaubt: Buchstaben, Ziffern, - und _", APP_TITLE))
    If Len(id) = 0 Then Exit Sub
    kurz = Trim$(InputBox("Kurzbeschreibung (Projektname) für '" & id & "' - optional:", APP_TITLE))
    CreateProject id, kurz
    ListeNeuLaden
    MsgBox "Projekt '" & id & "' wurde angelegt.", vbInformation, APP_TITLE
    Exit Sub
EH:
    HandleError "Start_NeuesProjekt"
End Sub

' Setzt 'Gesperrt von' fuer ein Projekt in der Liste (ohne Neuladen).
Private Sub SetListLock(ByVal id As String, ByVal txt As String)
    Dim ws As Worksheet, r As Long, o As Variant, key As String
    If mOrig Is Nothing Then Exit Sub
    If mListEnd < START_FIRST_ROW Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SHEET_START)
    key = UCase$(Trim$(id))
    For r = START_FIRST_ROW To mListEnd
        If UCase$(Trim$(CStr(ws.Cells(r, 1).Value))) = key Then
            Application.EnableEvents = False
            ws.Cells(r, 7).Value = txt
            If Len(txt) > 0 Then ws.Cells(r, 7).Interior.Color = COLOR_OWN Else ws.Cells(r, 7).Interior.Color = COLOR_READONLY
            Application.EnableEvents = True
            If mOrig.Exists(key) Then
                o = mOrig(key)
                o(5) = txt
                mOrig(key) = o
            End If
            Exit For
        End If
    Next r
End Sub

Public Sub Start_Oeffnen()
    Dim id As String, opened As Boolean
    On Error GoTo EH
    If IsProjectOpen() Then
        UserError "Es ist bereits das Projekt '" & CurrentId() & "' geöffnet." & vbCrLf & "Bitte zuerst im Blatt '" & SHEET_PROJEKT & "' schließen."
    End If
    id = SelectedProjectId("Ticket ID zum Öffnen:")
    If Len(id) = 0 Then Exit Sub
    OpenProject id
    opened = True
    ShowProject
    SetListLock id, OwnLockText()
    ThisWorkbook.Worksheets(SHEET_PROJEKT).Activate
    Exit Sub
EH:
    If opened Then CloseProject
    HandleError "Start_Oeffnen"
End Sub

Public Sub Start_Delete()
    Dim id As String
    On Error GoTo EH
    If Not ConfirmDiscardListChanges() Then Exit Sub
    id = SelectedProjectId("Ticket ID zum Entfernen (nur der Verweis, die Daten bleiben erhalten):")
    If Len(id) = 0 Then Exit Sub
    If MsgBox("Projekt '" & id & "' aus der Liste entfernen?" & vbCrLf & vbCrLf & _
              "Nur der Verweis wird gelöscht, die Projektdaten im Ordner bleiben erhalten." & vbCrLf & _
              "(Wiederherstellen: Admin Center, 'Projekt wiederherstellen')", vbYesNo + vbQuestion, APP_TITLE) <> vbYes Then Exit Sub
    DeleteProjectRef id
    ListeNeuLaden
    MsgBox "Das Projekt wurde entfernt. Die Daten sind weiterhin vorhanden.", vbInformation, APP_TITLE
    Exit Sub
EH:
    HandleError "Start_Delete"
End Sub

Public Sub Start_Purge()
    Dim id As String, chk As String
    On Error GoTo EH
    If Not ConfirmDiscardListChanges() Then Exit Sub
    id = SelectedProjectId("Ticket ID für 'Projekt vollständig löschen':")
    If Len(id) = 0 Then Exit Sub
    If MsgBox("ACHTUNG: Das Projekt '" & id & "' wird mit ALLEN gespeicherten Daten und Dateien endgültig gelöscht!" & vbCrLf & vbCrLf & _
              "Dieser Vorgang kann nicht rückgängig gemacht werden. Fortfahren?", vbYesNo + vbExclamation + vbDefaultButton2, APP_TITLE) <> vbYes Then Exit Sub
    chk = Trim$(InputBox("Zur Bestätigung die Ticket ID eingeben:" & vbCrLf & id, APP_TITLE))
    If chk <> id Then
        MsgBox "Die Eingabe stimmt nicht überein. Es wurde nichts gelöscht.", vbInformation, APP_TITLE
        Exit Sub
    End If
    PurgeProject id
    ListeNeuLaden
    MsgBox "Das Projekt '" & id & "' wurde vollständig gelöscht.", vbInformation, APP_TITLE
    Exit Sub
EH:
    HandleError "Start_Purge"
End Sub

Public Sub Start_Verknuepfen()
    Dim id As String
    On Error GoTo EH
    If Not ConfirmDiscardListChanges() Then Exit Sub
    id = Trim$(InputBox("Ticket ID des vorhandenen Projektordners (Ordnername unter 'projekte'):", APP_TITLE))
    If Len(id) = 0 Then Exit Sub
    LinkExistingProject id
    ListeNeuLaden
    MsgBox "Das Projekt '" & id & "' wurde wiederhergestellt und steht wieder in der Liste.", vbInformation, APP_TITLE
    Exit Sub
EH:
    HandleError "Start_Verknuepfen"
End Sub

'===============================================================
' PROJEKT: Felder anzeigen, bearbeiten, speichern
'===============================================================
Private Function CellText(ByVal s As String, ByVal keyName As String) As String
    If Len(s) > MAX_CELL_CHARS Then
        UserError "Der Wert von '" & keyName & "' ist zu lang für die Tabellenansicht (" & Len(s) & " Zeichen)."
    End If
    CellText = s
End Function

' Formatiert eine Zeile im Blatt 'Projekt' (bestehende Felder grau, S/B komplett schreibgeschuetzt).
Private Sub FormatProjektRow(ByVal ws As Worksheet, ByVal r As Long, ByVal cls As String)
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)).Interior.ColorIndex = -4142
    ws.Cells(r, 1).Interior.Color = COLOR_READONLY
    ws.Cells(r, 2).Interior.Color = COLOR_READONLY
    ws.Cells(r, 4).Interior.Color = COLOR_READONLY
    If cls <> "F" Then ws.Cells(r, 3).Interior.Color = COLOR_READONLY
End Sub

Public Sub ShowProject()
    Dim ws As Worksheet, p As Object, felder As Object, st As Object, doc As Object
    Dim ordered As Collection, it As Variant, r As Long, i As Long, txt As String, id As String, cls As String
    Dim q As String, n As Long, arr() As Variant
    Set ws = ThisWorkbook.Worksheets(SHEET_PROJEKT)
    Application.EnableEvents = False
    Set p = CurrentProject()
    Set felder = p("felder")
    Set doc = LoadFieldsDoc()
    Set ordered = ActiveFieldsOrdered(doc)
    Application.ScreenUpdating = False
    ws.Range("A" & PROJ_FIRST_ROW & ":D" & MAX_ROWS).ClearContents
    ws.Range("A" & PROJ_FIRST_ROW & ":D" & MAX_ROWS).Interior.ColorIndex = -4142
    ws.Range("A1:A2").NumberFormat = "@"
    ws.Range("A1:A2").Font.Size = 14
    ws.Range("A1:A2").Font.Bold = True
    ws.Rows(2).RowHeight = 22
    Set mPOrig = NewDict()
    ws.Range("A1").Value = CurrentId()
    ws.Range("A2").Value = FeldWert(felder, FID_KURZ)
    Set st = p("status")
    txt = "Aufgaben-Status:  "
    For i = 1 To TASK_COUNT
        txt = txt & i & ": "
        If JBool(st, "aufgabe_" & i) Then txt = txt & "erledigt" Else txt = txt & "offen"
        If i < TASK_COUNT Then txt = txt & "  |  "
    Next i
    ws.Range("A6").Value = txt
    n = ordered.Count
    If n > 0 Then
        ReDim arr(1 To n, 1 To 4)
        i = 0
        For Each it In ordered
            i = i + 1
            id = JStr(it, "id")
            arr(i, 1) = id
            arr(i, 2) = JStr(it, "key")
            arr(i, 3) = CellText(FeldWert(felder, id), id)
            If felder.Exists(id) Then arr(i, 4) = FeldQuelle(felder, id) Else arr(i, 4) = ""
            mPOrig(UCase$(id)) = CStr(arr(i, 3))
        Next it
        ws.Range(ws.Cells(PROJ_FIRST_ROW, 1), ws.Cells(PROJ_FIRST_ROW + n - 1, 4)).Value = arr
        r = PROJ_FIRST_ROW
        For Each it In ordered
            FormatProjektRow ws, r, FieldClass(JStr(it, "id"))
            r = r + 1
        Next it
    End If
    ws.Rows(PROJ_FIRST_ROW & ":" & (PROJ_FIRST_ROW + n + 1)).AutoFit
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    UpdateProjektFlag
End Sub

' Anzahl ungespeicherter Aenderungen im Blatt 'Projekt'
' (geaenderte lokale Werte, neue Zeilen, Aufgabenstatus)
Public Function ProjektDirtyCount() As Long
    Dim ws As Worksheet, v As Variant, lastRow As Long, i As Long, n As Long, id As String
    If IsProjectOpen() Then
        If TaskDirty() Then n = 1
        If Not mPOrig Is Nothing Then
            Set ws = ThisWorkbook.Worksheets(SHEET_PROJEKT)
            lastRow = LastUsedRow(ws, PROJ_FIRST_ROW, 4)
            If lastRow >= PROJ_FIRST_ROW Then
                v = ws.Range(ws.Cells(PROJ_FIRST_ROW, 1), ws.Cells(lastRow, 3)).Value
                For i = 1 To UBound(v, 1)
                    id = UCase$(Trim$(CStr(v(i, 1))))
                    If Len(id) > 0 Then
                        If mPOrig.Exists(id) Then
                            If FieldClass(id) = "F" Then
                                If CStr(v(i, 3)) <> CStr(mPOrig(id)) Then n = n + 1
                            End If
                        End If
                    ElseIf Len(Trim$(CStr(v(i, 2)))) > 0 Or Len(CStr(v(i, 3))) > 0 Then
                        n = n + 1
                    End If
                Next i
            End If
        End If
    End If
    ProjektDirtyCount = n
End Function

' Zeigt/entfernt den Hinweis 'Ungespeicherte Änderungen' (Zelle A3)
Public Sub UpdateProjektFlag()
    Dim ws As Worksheet, n As Long, ev As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_PROJEKT)
    n = ProjektDirtyCount()
    ev = Application.EnableEvents
    Application.EnableEvents = False
    If n > 0 Then
        ws.Range("A3").Value = ChrW(9679) & " Ungespeicherte Änderungen (" & n & ")"
    Else
        ws.Range("A3").ClearContents
    End If
    Application.EnableEvents = ev
End Sub

' Ereignis: Aenderung im Blatt 'Projekt' (ueber clsSheetEvents)
Public Sub Projekt_Changed(ByVal Target As Range)
    Dim ws As Worksheet, rng As Range, area As Range, r As Long, rws As Object, k As Variant
    Dim id As String, hasText As Boolean
    On Error GoTo EH
    If Not IsProjectOpen() Then Exit Sub
    If mPOrig Is Nothing Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SHEET_PROJEKT)
    Set rng = Intersect(Target, ws.Range("A" & PROJ_FIRST_ROW & ":D" & MAX_ROWS))
    If rng Is Nothing Then Exit Sub
    Set rws = NewDict()
    For Each area In rng.Areas
        For r = area.Row To area.Row + area.Rows.Count - 1
            rws(CStr(r)) = True
        Next r
    Next area
    Application.EnableEvents = False
    For Each k In rws.Keys
        r = CLng(k)
        id = UCase$(Trim$(CStr(ws.Cells(r, 1).Value)))
        If Len(id) > 0 Then
            If mPOrig.Exists(id) And FieldClass(id) = "F" Then
                If CStr(ws.Cells(r, 3).Value) <> CStr(mPOrig(id)) Then
                    ws.Cells(r, 3).Interior.Color = COLOR_CHANGED
                Else
                    ws.Cells(r, 3).Interior.ColorIndex = -4142
                End If
            End If
        Else
            hasText = (Len(Trim$(CStr(ws.Cells(r, 2).Value))) > 0 Or Len(CStr(ws.Cells(r, 3).Value)) > 0)
            If hasText Then
                ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).Interior.Color = COLOR_CHANGED
            Else
                ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).Interior.ColorIndex = -4142
            End If
        End If
    Next k
    Application.EnableEvents = True
    UpdateProjektFlag
    Exit Sub
EH:
    HandleError "Projekt_Changed"
End Sub

Private Function LastUsedRow(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal nCols As Long) As Long
    Dim c As Long, lr As Long, x As Long
    lr = firstRow - 1
    For c = 1 To nCols
        x = ws.Cells(ws.Rows.Count, c).End(-4162).Row   ' xlUp
        If x > lr Then lr = x
    Next c
    LastUsedRow = lr
End Function

' Liest das Blatt 'Projekt' in das geoeffnete Projekt zurueck.
' Zwei Durchgaenge: erst alles pruefen (nichts wird geaendert), dann uebernehmen.
Private Sub ReadSheetIntoProject(ByRef nNew As Long, ByRef nIgnored As Long)
    Dim ws As Worksheet, p As Object, felder As Object, doc As Object, fl As Collection
    Dim v As Variant, lastRow As Long, i As Long, r As Long
    Dim id As String, nm As String, wert As String, cls As String, fld As Object
    Dim errors As Collection, seenIds As Object, usedKeys As Object, newKeys As Object, it As Variant
    Dim newIds As Object, react As Boolean, newId As String, oldV As String
    nNew = 0
    nIgnored = 0
    Set ws = ThisWorkbook.Worksheets(SHEET_PROJEKT)
    Set p = CurrentProject()
    Set felder = p("felder")
    Set doc = LoadFieldsDoc()
    Set fl = doc("felder")
    lastRow = LastUsedRow(ws, PROJ_FIRST_ROW, 4)
    If lastRow < PROJ_FIRST_ROW Then Exit Sub
    v = ws.Range(ws.Cells(PROJ_FIRST_ROW, 1), ws.Cells(lastRow, 3)).Value
    ' alte Fehlermarkierungen entfernen
    For i = 1 To UBound(v, 1)
        r = PROJ_FIRST_ROW + i - 1
        If ws.Cells(r, 2).Interior.Color = COLOR_ERROR Then
            If Len(Trim$(CStr(v(i, 1)))) > 0 Then
                ws.Cells(r, 2).Interior.Color = COLOR_READONLY
            Else
                ws.Cells(r, 2).Interior.ColorIndex = -4142
            End If
        End If
    Next i
    Set errors = New Collection
    Set seenIds = NewDict()
    Set usedKeys = NewDict()
    Set newKeys = NewDict()
    For Each it In fl
        If FieldClass(JStr(it, "id")) = "F" And IsFieldActive(it) Then usedKeys(LCase$(Trim$(JStr(it, "key")))) = JStr(it, "id")
    Next it
    ' --- Durchgang 1: pruefen ---
    For i = 1 To UBound(v, 1)
        r = PROJ_FIRST_ROW + i - 1
        id = Trim$(CStr(v(i, 1)))
        nm = Trim$(CStr(v(i, 2)))
        wert = CStr(v(i, 3))
        If Len(id) > 0 Then
            Set fld = FindFieldById(fl, id)
            If fld Is Nothing Then
                errors.Add "Zeile " & r & ": Field ID '" & id & "' ist unbekannt."
            ElseIf seenIds.Exists(UCase$(id)) Then
                errors.Add "Zeile " & r & ": Field ID " & id & " steht schon in Zeile " & seenIds(UCase$(id)) & "."
            End If
            seenIds(UCase$(id)) = r
        ElseIf Len(nm) = 0 Then
            If Len(wert) > 0 Then errors.Add "Zeile " & r & ": Es gibt einen Wert, aber weder Field ID noch Name."
        Else
            If Len(nm) > 100 Then
                errors.Add "Zeile " & r & ": Der Name ist zu lang (max. 100 Zeichen)."
                ws.Cells(r, 2).Interior.Color = COLOR_ERROR
            ElseIf usedKeys.Exists(LCase$(nm)) Then
                errors.Add "Zeile " & r & ": Der Name '" & Left$(nm, 40) & "' ist schon vergeben (" & usedKeys(LCase$(nm)) & ")."
                ws.Cells(r, 2).Interior.Color = COLOR_ERROR
            ElseIf newKeys.Exists(LCase$(nm)) Then
                errors.Add "Zeile " & r & ": Der Name '" & Left$(nm, 40) & "' kommt auch in Zeile " & newKeys(LCase$(nm)) & " vor."
                ws.Cells(r, 2).Interior.Color = COLOR_ERROR
                ws.Cells(newKeys(LCase$(nm)), 2).Interior.Color = COLOR_ERROR
            Else
                newKeys(LCase$(nm)) = r
            End If
        End If
    Next i
    If errors.Count > 0 Then
        UserError "Es wurde nichts gespeichert. Bitte korrigieren:" & vbCrLf & vbCrLf & JoinLines(errors, 12)
    End If
    ' --- neue lokale Felder anlegen (Feldliste) ---
    Set newIds = NewDict()
    For i = 1 To UBound(v, 1)
        id = Trim$(CStr(v(i, 1)))
        nm = Trim$(CStr(v(i, 2)))
        If Len(id) = 0 And Len(nm) > 0 Then
            newId = RegisterNewFField(nm, react)
            newIds(CStr(i)) = newId
            nNew = nNew + 1
        End If
    Next i
    ' --- Durchgang 2: Werte uebernehmen ---
    For i = 1 To UBound(v, 1)
        id = Trim$(CStr(v(i, 1)))
        nm = Trim$(CStr(v(i, 2)))
        wert = CStr(v(i, 3))
        If Len(id) > 0 Then
            cls = FieldClass(id)
            oldV = FeldWert(felder, UCase$(id))
            If cls = "F" Then
                If oldV <> wert Then
                    If felder.Exists(UCase$(id)) Or Len(wert) > 0 Then
                        SetFeld felder, UCase$(id), JStr(FindFieldById(fl, id), "key"), wert, SRC_F
                    End If
                End If
            ElseIf oldV <> wert Then
                nIgnored = nIgnored + 1
            End If
        ElseIf newIds.Exists(CStr(i)) Then
            If Len(wert) > 0 Then SetFeld felder, CStr(newIds(CStr(i))), nm, wert, SRC_F
        End If
    Next i
End Sub

Public Sub Projekt_Speichern()
    Dim nNew As Long, nIgn As Long, msg As String
    On Error GoTo EH
    RequireOpenProject
    ReadSheetIntoProject nNew, nIgn
    SaveProject
    ShowProject
    msg = "Das Projekt '" & CurrentId() & "' wurde gespeichert."
    If nNew > 0 Then msg = msg & vbCrLf & nNew & " neue(s) lokale(s) Feld(er) steht/stehen jetzt in allen Projekten zur Verfügung."
    If nIgn > 0 Then msg = msg & vbCrLf & nIgn & " Änderung(en) an System-/BeKo-Feldern wurden ignoriert (schreibgeschützt)."
    MsgBox msg, vbInformation, APP_TITLE
    Exit Sub
EH:
    HandleError "Projekt_Speichern"
End Sub

Public Sub Projekt_Schliessen()
    Dim a As VbMsgBoxResult, nNew As Long, nIgn As Long, ws As Worksheet
    On Error GoTo EH
    If Not IsProjectOpen() Then
        ThisWorkbook.Worksheets(SHEET_START).Activate
        Exit Sub
    End If
    If ProjektDirtyCount() > 0 Then
        a = MsgBox("Es gibt ungespeicherte Änderungen. Vor dem Schließen speichern?" & vbCrLf & vbCrLf & _
                   "Ja = speichern, Nein = verwerfen, Abbrechen = Projekt geöffnet lassen.", vbYesNoCancel + vbQuestion, APP_TITLE)
        If a = vbCancel Then Exit Sub
        If a = vbYes Then
            ReadSheetIntoProject nNew, nIgn
            SaveProject
        End If
    End If
    Set mPOrig = Nothing
    CloseProject
    Application.EnableEvents = False
    Set ws = ThisWorkbook.Worksheets(SHEET_PROJEKT)
    ws.Range("A3").ClearContents
    ws.Range("A" & PROJ_FIRST_ROW & ":D" & MAX_ROWS).ClearContents
    ws.Range("A" & PROJ_FIRST_ROW & ":D" & MAX_ROWS).Interior.ColorIndex = -4142
    ws.Range("A1").Value = "(kein Projekt geöffnet)"
    ws.Range("A2").ClearContents
    ws.Range("A6").Value = "Aufgaben-Status: -"
    Application.EnableEvents = True
    ThisWorkbook.Worksheets(SHEET_START).Activate
    Start_Aktualisieren
    Exit Sub
EH:
    HandleError "Projekt_Schliessen"
End Sub

'===============================================================
' FELDER BEARBEITEN: globale Feldliste
'===============================================================
Public Sub Felder_Laden()
    Dim ws As Worksheet, doc As Object, ordered As Collection, it As Variant, r As Long, n As Long, i As Long
    Dim arr() As Variant, id As String, cls As String
    On Error GoTo EH
    Set ws = ThisWorkbook.Worksheets(SHEET_FELDER)
    InitDataRoot
    Set doc = LoadFieldsDoc()
    Set ordered = ActiveFieldsOrdered(doc)
    Application.ScreenUpdating = False
    ws.Range("A" & FELD_FIRST_ROW & ":D" & MAX_ROWS).ClearContents
    ws.Range("A" & FELD_FIRST_ROW & ":D" & MAX_ROWS).Interior.ColorIndex = -4142
    Set mFKey = NewDict()
    Set mFDesc = NewDict()
    n = ordered.Count
    If n > 0 Then
        ReDim arr(1 To n, 1 To 4)
        i = 0
        For Each it In ordered
            i = i + 1
            id = JStr(it, "id")
            cls = FieldClass(id)
            arr(i, 1) = id
            arr(i, 2) = JStr(it, "key")
            arr(i, 3) = JStr(it, "beschreibung")
            arr(i, 4) = SourceText(cls)
            mFKey(UCase$(id)) = JStr(it, "key")
            mFDesc(UCase$(id)) = JStr(it, "beschreibung")
        Next it
        ws.Range(ws.Cells(FELD_FIRST_ROW, 1), ws.Cells(FELD_FIRST_ROW + n - 1, 4)).Value = arr
        r = FELD_FIRST_ROW
        For Each it In ordered
            FormatFeldRow ws, r, FieldClass(JStr(it, "id")), JBool(it, "standard")
            r = r + 1
        Next it
    End If
    mFLoaded = True
    Application.ScreenUpdating = True
    Exit Sub
EH:
    HandleError "Felder_Laden"
End Sub

Private Sub FormatFeldRow(ByVal ws As Worksheet, ByVal r As Long, ByVal cls As String, ByVal isStandard As Boolean)
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)).Interior.ColorIndex = -4142
    ws.Cells(r, 1).Interior.Color = COLOR_READONLY
    ws.Cells(r, 4).Interior.Color = COLOR_READONLY
    If cls <> "F" Then
        ws.Cells(r, 2).Interior.Color = COLOR_READONLY
        ws.Cells(r, 3).Interior.Color = COLOR_READONLY
    ElseIf isStandard Then
        ws.Cells(r, 2).Interior.Color = COLOR_READONLY
    End If
End Sub

Public Sub Felder_Speichern()
    Dim ws As Worksheet, v As Variant, lastRow As Long, i As Long, r As Long
    Dim id As String, nm As String, ds As String, cls As String
    Dim errors As Collection, seenIds As Object, keyRows As Object, rows As Collection
    Dim renamed As Object, nNew As Long, nDeact As Long, nReact As Long, warnings As Collection
    Dim nUpd As Long, skipped As Collection, msg As String, firstRow As Long, nDesc As Long
    On Error GoTo EH
    If Not mFLoaded Then UserError "Bitte zuerst 'Neu laden' ausführen, damit Änderungen anderer Clients nicht überschrieben werden."
    Set ws = ThisWorkbook.Worksheets(SHEET_FELDER)
    InitDataRoot
    lastRow = LastUsedRow(ws, FELD_FIRST_ROW, 3)
    Set errors = New Collection
    Set seenIds = NewDict()
    Set keyRows = NewDict()
    Set rows = New Collection
    ' alte Fehlermarkierungen entfernen
    For r = FELD_FIRST_ROW To lastRow
        If ws.Cells(r, 2).Interior.Color = COLOR_ERROR Then
            If Len(Trim$(CStr(ws.Cells(r, 1).Value))) > 0 Then
                FormatFeldRow ws, r, FieldClass(CStr(ws.Cells(r, 1).Value)), False
            Else
                ws.Cells(r, 2).Interior.ColorIndex = -4142
            End If
        End If
    Next r
    If lastRow >= FELD_FIRST_ROW Then
        v = ws.Range(ws.Cells(FELD_FIRST_ROW, 1), ws.Cells(lastRow, 3)).Value
        For i = 1 To UBound(v, 1)
            r = FELD_FIRST_ROW + i - 1
            id = Trim$(CStr(v(i, 1)))
            nm = Trim$(CStr(v(i, 2)))
            ds = Trim$(CStr(v(i, 3)))
            If Len(id) > 0 Or Len(nm) > 0 Or Len(ds) > 0 Then
                cls = ""
                If Len(id) > 0 Then
                    cls = FieldClass(id)
                    If Len(cls) = 0 Then
                        errors.Add "Zeile " & r & ": ungültige Field ID '" & Left$(id, 20) & "'."
                    ElseIf seenIds.Exists(UCase$(id)) Then
                        errors.Add "Zeile " & r & ": Field ID " & id & " steht schon in Zeile " & seenIds(UCase$(id)) & "."
                    End If
                    seenIds(UCase$(id)) = r
                End If
                ' Key-Pruefung fuer lokale Felder (bestehende F und neue Zeilen)
                If cls = "F" Or Len(id) = 0 Then
                    If Len(nm) = 0 Then
                        errors.Add "Zeile " & r & ": Der Name (Key) fehlt."
                        ws.Cells(r, 2).Interior.Color = COLOR_ERROR
                    ElseIf Len(nm) > 100 Then
                        errors.Add "Zeile " & r & ": Der Name ist zu lang (max. 100 Zeichen)."
                        ws.Cells(r, 2).Interior.Color = COLOR_ERROR
                    ElseIf keyRows.Exists(LCase$(nm)) Then
                        errors.Add "Zeile " & r & ": Der Name '" & Left$(nm, 40) & "' kommt auch in Zeile " & keyRows(LCase$(nm)) & " vor."
                        ws.Cells(r, 2).Interior.Color = COLOR_ERROR
                        ws.Cells(keyRows(LCase$(nm)), 2).Interior.Color = COLOR_ERROR
                    Else
                        keyRows(LCase$(nm)) = r
                    End If
                End If
                rows.Add Array(id, nm, ds, r)
            End If
        Next i
    End If
    If errors.Count > 0 Then
        UserError "Es wurde nichts gespeichert. Bitte korrigieren (rot markiert):" & vbCrLf & vbCrLf & JoinLines(errors, 12)
    End If
    ApplyFieldRows rows, mFKey, mFDesc, renamed, nNew, nDeact, nReact, warnings, nDesc
    Set skipped = PropagateRenames(renamed, nUpd)
    Felder_Laden
    msg = "Die Feldliste wurde gespeichert."
    If nNew + nReact + nDeact + nDesc + renamed.Count > 0 Then
        msg = msg & vbCrLf & nNew & " neu, " & nReact & " reaktiviert, " & nDeact & " deaktiviert, " & renamed.Count & " umbenannt, " & nDesc & " Beschreibung(en) geändert."
    Else
        msg = msg & vbCrLf & "Es gab keine Änderungen."
    End If
    If skipped.Count > 0 Then msg = msg & vbCrLf & vbCrLf & skipped.Count & " Projekt(e) gesperrt - der neue Name wird dort beim nächsten Speichern übernommen."
    If warnings.Count > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Hinweise:" & vbCrLf & JoinLines(warnings, 6)
        For i = 1 To warnings.Count
            WriteLog "Felder speichern, Hinweis: " & CStr(warnings(i))
        Next i
    End If
    WriteLog "Feldliste gespeichert: " & nNew & " neu, " & nReact & " reaktiviert, " & nDeact & " deaktiviert, " & renamed.Count & " umbenannt, " & nDesc & " Beschreibung(en)"
    MsgBox msg, vbInformation, APP_TITLE
    Exit Sub
EH:
    HandleError "Felder_Speichern"
End Sub

' Button 'Neues Feld': fragt den Namen, legt die Zeile an und speichert.
Public Sub Felder_Neu()
    Dim ws As Worksheet, nm As String, r As Long
    On Error GoTo EH
    If Not mFLoaded Then UserError "Bitte zuerst 'Neu laden' ausführen."
    Set ws = ThisWorkbook.Worksheets(SHEET_FELDER)
    nm = Trim$(InputBox("Name (Key) des neuen lokalen Feldes:", APP_TITLE))
    If Len(nm) = 0 Then Exit Sub
    r = LastUsedRow(ws, FELD_FIRST_ROW, 3) + 1
    ws.Cells(r, 1).ClearContents
    ws.Cells(r, 2).Value = nm
    ws.Cells(r, 4).Value = SRC_F
    ws.Cells(r, 4).Interior.Color = COLOR_READONLY
    Felder_Speichern
    Exit Sub
EH:
    HandleError "Felder_Neu"
End Sub

' Button 'Feld deaktivieren': markierte Zeilen (nur Verweis, Werte bleiben erhalten).
Public Sub Felder_Deaktivieren()
    Dim ws As Worksheet, sel As Range, area As Range, r As Long, rws As Object, k As Variant
    Dim id As String, ids As Collection, skipStd As Long, lst As String
    On Error GoTo EH
    If Not mFLoaded Then UserError "Bitte zuerst 'Neu laden' ausführen."
    Set ws = ThisWorkbook.Worksheets(SHEET_FELDER)
    If ActiveSheet.Name <> SHEET_FELDER Then UserError "Bitte im Blatt '" & SHEET_FELDER & "' die Zeile(n) markieren."
    If TypeName(Selection) <> "Range" Then UserError "Bitte zuerst die Zeile(n) der Felder markieren."
    Set sel = Selection
    Set rws = NewDict()
    For Each area In sel.Areas
        For r = area.Row To area.Row + area.Rows.Count - 1
            If r >= FELD_FIRST_ROW And r <= MAX_ROWS Then rws(CStr(r)) = True
        Next r
    Next area
    Set ids = New Collection
    For Each k In rws.Keys
        id = Trim$(CStr(ws.Cells(CLng(k), 1).Value))
        If Len(id) > 0 Then
            If mFKey.Exists(UCase$(id)) Then
                If UCase$(id) = FID_KURZ Or UCase$(id) = FID_BEARB Then
                    skipStd = skipStd + 1
                Else
                    ids.Add CLng(k)
                    lst = lst & id & " "
                End If
            End If
        End If
    Next k
    If ids.Count = 0 Then
        If skipStd > 0 Then UserError "Standardfelder können nicht deaktiviert werden."
        UserError "In der Markierung ist kein Feld mit Field ID enthalten."
    End If
    If MsgBox(ids.Count & " Feld(er) deaktivieren: " & Left$(lst, 200) & vbCrLf & vbCrLf & _
              "Die Felder werden nur ausgeblendet; die Werte in den Projekten bleiben erhalten." & vbCrLf & _
              "Weitere Änderungen in der Tabelle werden dabei mitgespeichert.", vbYesNo + vbQuestion, APP_TITLE) <> vbYes Then Exit Sub
    For Each k In ids
        ws.Range(ws.Cells(CLng(k), 1), ws.Cells(CLng(k), 4)).ClearContents
    Next k
    Felder_Speichern
    Exit Sub
EH:
    HandleError "Felder_Deaktivieren"
End Sub
