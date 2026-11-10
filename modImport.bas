Attribute VB_Name = "modImport"
Option Explicit
'===============================================================
' modImport - 'Update Projekte': Excel-Tabelle -> S-Felder
' Zeile 1 = Spaltentitel (Keys), jede weitere Zeile = ein Projekt.
' ID-Spalte: 'Ticket ID' (Vorrang) oder 'Project ID'.
' Spalten 'Kurzbeschreibung' / 'Bearbeiter' fuellen die Standard-
' felder F001/F002 (Quelle 'aus System'), alle anderen Spalten
' werden S-Felder (nur Import darf sie aendern).
'===============================================================

Private Const MAX_KEY_LEN As Long = 100

' Normalisiert Spaltentitel: ohne Gross/Klein, Leerzeichen, _ und -
Private Function NormKey(ByVal s As String) As String
    s = LCase$(Trim$(s))
    s = Replace(s, " ", "")
    s = Replace(s, "_", "")
    s = Replace(s, "-", "")
    s = Replace(s, ChrW(160), "")
    NormKey = s
End Function

' Zelle -> Text (Zahlen mit Punkt, Datum ISO, Boolean true/false)
Private Function CellToText(ByVal v As Variant) As String
    Dim s As String, dsep As String
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function
    Select Case VarType(v)
        Case vbDate
            If CDbl(v) = Int(CDbl(v)) Then
                s = Format$(v, "yyyy-mm-dd")
            Else
                s = Format$(v, "yyyy-mm-dd") & "T" & Format$(v, "hh:nn:ss")
            End If
        Case vbBoolean
            If v Then s = "true" Else s = "false"
        Case vbDouble, vbSingle, vbCurrency, vbDecimal
            s = CStr(v)
            dsep = Application.International(3)   ' xlDecimalSeparator
            If dsep <> "." Then s = Replace(s, dsep, ".")
        Case Else
            s = CStr(v)
    End Select
    CellToText = Trim$(s)
End Function

Private Function PickExcelFile() As String
    Dim fd As Object
    Set fd = Application.FileDialog(3)             ' msoFileDialogFilePicker
    With fd
        .Title = "Excel-Datei für 'Update Projekte' wählen"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel-Dateien", "*.xlsx;*.xlsm;*.xls"
        If .Show = -1 Then PickExcelFile = .SelectedItems(1)
    End With
End Function

' Sucht die ID-Spalte. Rueckgabe: Spaltenindex (0 = nicht gefunden)
Private Function FindIdColumn(ByVal titles As Variant, ByVal nCols As Long) As Long
    Dim c As Long, tk As String, cfg As String
    For c = 1 To nCols
        If NormKey(CStr(titles(c))) = "ticketid" Then FindIdColumn = c: Exit Function
    Next c
    For c = 1 To nCols
        tk = NormKey(CStr(titles(c)))
        If tk = "projectid" Or tk = "projektid" Then FindIdColumn = c: Exit Function
    Next c
    cfg = GetSetting("import_id_spalte", "")
    If Len(cfg) > 0 Then
        For c = 1 To nCols
            If NormKey(CStr(titles(c))) = NormKey(cfg) Then FindIdColumn = c: Exit Function
        Next c
    End If
End Function

Private Sub ReleaseQuiet(ByVal lp As String)
    On Error Resume Next
    ReleaseLock lp
End Sub

'---------------------------------------------------------------
' Aktualisiert ein Projekt. Rueckgabe: PR_OK, PR_UNCHANGED, PR_LOCKED, PR_ERROR
'---------------------------------------------------------------
Private Function UpdateOne(ByVal id As String, ByVal sIds As Collection, ByVal sKeys As Object, _
                           ByVal colOf As Object, ByVal rowVals As Variant, _
                           ByVal colKurz As Long, ByVal colBearb As Long, ByRef info As String) As Long
    Dim held As String, lp As String, p As Object, felder As Object, i As Long
    Dim fid As String, v As String, wr As Boolean, n As Long, d As String
    info = ""
    If IsProjectOpen() Then
        If StrComp(CurrentId(), id, vbTextCompare) = 0 Then
            info = "ist im Blatt 'Projekt' geöffnet"
            UpdateOne = PR_LOCKED
            Exit Function
        End If
    End If
    lp = LockPath(id)
    If Not TryAcquireLock(lp, held) Then
        info = "gesperrt von " & held
        UpdateOne = PR_LOCKED
        Exit Function
    End If
    On Error GoTo EH
    Set p = ReadJsonFile(ProjectFile(id))
    CheckProjectSchema p
    EnsureStructure p
    Set felder = p("felder")
    For i = 1 To sIds.Count
        fid = CStr(sIds(i))
        v = CellToText(rowVals(CLng(colOf(fid))))
        If FeldWert(felder, fid) <> v Or FeldQuelle(felder, fid) <> SRC_S Or Not felder.Exists(fid) Then
            SetFeld felder, fid, CStr(sKeys(fid)), v, SRC_S
            wr = True
        End If
    Next i
    If colKurz > 0 Then
        v = CellToText(rowVals(colKurz))
        If Len(v) > 0 Then
            If FeldWert(felder, FID_KURZ) <> v Or FeldQuelle(felder, FID_KURZ) <> SRC_S Then
                SetFeld felder, FID_KURZ, KEY_KURZ, v, SRC_S
                wr = True
            End If
        End If
    End If
    If colBearb > 0 Then
        v = CellToText(rowVals(colBearb))
        If Len(v) > 0 Then
            If FeldWert(felder, FID_BEARB) <> v Or FeldQuelle(felder, FID_BEARB) <> SRC_S Then
                SetFeld felder, FID_BEARB, KEY_BEARB, v, SRC_S
                wr = True
            End If
        End If
    End If
    If wr Then
        RefreshKeyCopies felder
        p("geaendert") = IsoNow()
        p("geaendert_von") = UserTag()
        WriteJsonFile ProjectFile(id), p
        UpdateOne = PR_OK
    Else
        UpdateOne = PR_UNCHANGED
    End If
    ReleaseLock lp
    Exit Function
EH:
    n = Err.Number
    d = Err.Description
    ReleaseQuiet lp
    info = "[" & n & "] " & Left$(d, 150)
    UpdateOne = PR_ERROR
End Function

'---------------------------------------------------------------
' Button 'Update Projekte'
'---------------------------------------------------------------
Public Sub Start_UpdateProjekte()
    Dim path As String, wb As Workbook, ws As Worksheet, v As Variant
    Dim nRows As Long, nCols As Long, r As Long, c As Long
    Dim idCol As Long, colKurz As Long, colBearb As Long, tk As String, t As String
    Dim keys As Collection, colOfKey As Object, seen As Object, errs As Collection
    Dim idOf As Object, keyOf As Object, sIds As Collection, colOf As Object
    Dim nNewF As Long, nReactF As Long, fid As String, lk As String
    Dim ids As Collection, rowVals As Variant, id As String, k As Long
    Dim unknown As Collection, createUnknown As Boolean, ans As VbMsgBoxResult
    Dim cOk As Long, cSame As Long, cLocked As Long, cErr As Long, cNew As Long, cSkip As Long, cDup As Long
    Dim lockedList As Collection, errList As Collection, rc As Long, info As String, exists As Boolean
    Dim msg As String, opened As Boolean, seenId As Object, rowOfId As Object
    On Error GoTo EH
    InitDataRoot
    If IsProjectOpen() Then
        UserError "Es ist ein Projekt geöffnet (" & CurrentId() & ")." & vbCrLf & "Bitte zuerst im Blatt 'Projekt' speichern und schließen."
    End If
    If Not ConfirmDiscardListChanges() Then Exit Sub
    LoadFieldsDoc
    path = PickExcelFile()
    If Len(path) = 0 Then Exit Sub

    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(path, 0, True)       ' UpdateLinks=0, ReadOnly
    opened = True
    Set ws = wb.Worksheets(1)
    nRows = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    nCols = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    If nRows < 2 Or nCols < 1 Then
        wb.Close False
        opened = False
        Application.ScreenUpdating = True
        UserError "Die erste Tabelle enthält keine Daten (Zeile 1 = Spaltentitel, ab Zeile 2 Projekte)."
    End If
    v = ws.Range(ws.Cells(1, 1), ws.Cells(nRows, nCols)).Value
    wb.Close False
    opened = False
    Application.ScreenUpdating = True
    If nRows = 1 Then
        ' einzelne Zeile liefert kein 2D-Array - oben abgefangen
        UserError "Die erste Tabelle enthält keine Daten."
    End If

    ' Titel aus Zeile 1
    Dim titles() As Variant
    ReDim titles(1 To nCols)
    For c = 1 To nCols
        titles(c) = Trim$(CStr(v(1, c)))
    Next c
    idCol = FindIdColumn(titles, nCols)
    If idCol = 0 Then
        t = InputBox("In Zeile 1 wurde weder 'Ticket ID' noch 'Project ID' gefunden." & vbCrLf & _
                     "Wie heißt die Spalte mit der Projekt-ID (genauer Spaltentitel)?", APP_TITLE)
        If Len(Trim$(t)) = 0 Then Exit Sub
        For c = 1 To nCols
            If NormKey(CStr(titles(c))) = NormKey(t) Then idCol = c: Exit For
        Next c
        If idCol = 0 Then UserError "Die Spalte '" & t & "' wurde nicht gefunden."
        SetSetting "import_id_spalte", Trim$(t)
    End If

    ' Spalten einordnen und Keys pruefen
    Set keys = New Collection
    Set colOfKey = NewDict()
    Set seen = NewDict()
    Set errs = New Collection
    For c = 1 To nCols
        If c <> idCol Then
            t = CStr(titles(c))
            If Len(t) > 0 Then
                tk = NormKey(t)
                If tk = "kurzbeschreibung" Then
                    If colKurz > 0 Then errs.Add "Spalte " & c & ": 'Kurzbeschreibung' kommt mehrfach vor"
                    colKurz = c
                ElseIf tk = "bearbeiter" Then
                    If colBearb > 0 Then errs.Add "Spalte " & c & ": 'Bearbeiter' kommt mehrfach vor"
                    colBearb = c
                Else
                    lk = LCase$(t)
                    If Len(t) > MAX_KEY_LEN Then
                        errs.Add "Spalte " & c & ": Titel länger als " & MAX_KEY_LEN & " Zeichen"
                    ElseIf seen.Exists(lk) Then
                        errs.Add "Spalte " & c & ": Titel '" & t & "' kommt doppelt vor (Spalte " & seen(lk) & ")"
                    Else
                        seen(lk) = c
                        keys.Add t
                        colOfKey(lk) = c
                    End If
                End If
            End If
        End If
    Next c
    If errs.Count > 0 Then
        WriteLog "Update Projekte, Titelfehler: " & JoinLines(errs, 200)
        UserError "Die Spaltentitel sind nicht eindeutig:" & vbCrLf & JoinLines(errs, 8) & vbCrLf & vbCrLf & "Nichts wurde geändert."
    End If

    ' S-Felder bestimmen/anlegen
    ResolveImportFields keys, idOf, keyOf, nNewF, nReactF
    Set sIds = New Collection
    Set colOf = NewDict()
    For k = 1 To keys.Count
        lk = LCase$(CStr(keys(k)))
        fid = CStr(idOf(lk))
        sIds.Add fid
        colOf(fid) = CLng(colOfKey(lk))
    Next k

    ' Projekt-IDs einsammeln
    Set ids = New Collection
    Set seenId = NewDict()
    Set unknown = New Collection
    Set errList = New Collection
    Set lockedList = New Collection
    For r = 2 To nRows
        id = CellToText(v(r, idCol))
        If Len(id) > 0 Then
            If Not IsValidProjectId(id) Then
                cSkip = cSkip + 1
                errList.Add "Zeile " & r & ": ungültige ID '" & Left$(id, 40) & "'"
            ElseIf seenId.Exists(UCase$(id)) Then
                cDup = cDup + 1
                errList.Add "Zeile " & r & ": ID '" & id & "' kommt doppelt vor (erste Zeile " & seenId(UCase$(id)) & " gilt)"
            Else
                seenId(UCase$(id)) = r
                ids.Add id
                If Not FileExists(ProjectFile(id)) Then unknown.Add id
            End If
        End If
    Next r
    If ids.Count = 0 Then
        UserError "Es wurden keine gültigen Projekt-IDs gefunden (Spalte " & idCol & ")."
    End If

    createUnknown = False
    If unknown.Count > 0 Then
        ans = MsgBox(unknown.Count & " Projekt-ID(s) sind noch nicht angelegt, z. B. " & CStr(unknown(1)) & "." & vbCrLf & vbCrLf & _
                     "Diese Projekte neu anlegen?" & vbCrLf & "Ja = anlegen, Nein = überspringen, Abbrechen = alles abbrechen.", _
                     vbYesNoCancel + vbQuestion, APP_TITLE)
        If ans = vbCancel Then Exit Sub
        createUnknown = (ans = vbYes)
    End If

    Application.StatusBar = "Update Projekte läuft ..."
    ReDim rowVals(1 To nCols)
    For k = 1 To ids.Count
        id = CStr(ids(k))
        r = CLng(seenId(UCase$(id)))
        exists = FileExists(ProjectFile(id))
        If Not exists Then
            If createUnknown Then
                On Error Resume Next
                CreateProject id
                If Err.Number <> 0 Then
                    cErr = cErr + 1
                    errList.Add "Zeile " & r & " (" & id & "): " & Left$(Err.Description, 100)
                    Err.Clear
                    On Error GoTo EH
                    GoTo NextProject
                End If
                On Error GoTo EH
                cNew = cNew + 1
            Else
                cSkip = cSkip + 1
                GoTo NextProject
            End If
        End If
        For c = 1 To nCols
            rowVals(c) = v(r, c)
        Next c
        rc = UpdateOne(id, sIds, keyOf, colOf, rowVals, colKurz, colBearb, info)
        Select Case rc
            Case PR_OK: cOk = cOk + 1
            Case PR_UNCHANGED: cSame = cSame + 1
            Case PR_LOCKED
                cLocked = cLocked + 1
                lockedList.Add id & " (" & info & ")"
            Case Else
                cErr = cErr + 1
                errList.Add "Zeile " & r & " (" & id & "): " & info
        End Select
NextProject:
    Next k
    Application.StatusBar = False

    msg = "Update Projekte abgeschlossen." & vbCrLf & vbCrLf & _
          cOk & " Projekt(e) aktualisiert, " & cSame & " unverändert." & vbCrLf
    If cNew > 0 Then msg = msg & cNew & " neu angelegt." & vbCrLf
    If nNewF + nReactF > 0 Then msg = msg & nNewF & " neue(s) S-Feld(er), " & nReactF & " reaktiviert." & vbCrLf
    If cSkip > 0 Then msg = msg & cSkip & " Zeile(n) übersprungen." & vbCrLf
    If cDup > 0 Then msg = msg & cDup & " doppelte ID(s) ignoriert." & vbCrLf
    If cLocked > 0 Then msg = msg & vbCrLf & cLocked & " Projekt(e) gesperrt und NICHT aktualisiert:" & vbCrLf & JoinLines(lockedList, 4) & vbCrLf
    If cErr > 0 Then msg = msg & vbCrLf & cErr & " Fehler:" & vbCrLf & JoinLines(errList, 4) & vbCrLf
    If cLocked + cErr + cSkip + cDup > 0 Then msg = msg & vbCrLf & "Details stehen im Protokoll."
    WriteLog "Update Projekte aus '" & path & "': ok=" & cOk & ", unverändert=" & cSame & ", neu=" & cNew & _
             ", gesperrt=" & cLocked & ", Fehler=" & cErr & ", übersprungen=" & cSkip & ", doppelt=" & cDup
    If lockedList.Count > 0 Then WriteLog "Update Projekte gesperrt: " & JoinLines(lockedList, 500)
    If errList.Count > 0 Then WriteLog "Update Projekte Fehler/Hinweise: " & JoinLines(errList, 500)
    ListeNeuLaden
    MsgBox msg, IIf(cLocked + cErr > 0, vbExclamation, vbInformation), APP_TITLE
    Exit Sub
EH:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    On Error Resume Next
    If opened Then wb.Close False
    On Error GoTo 0
    HandleError "Start_UpdateProjekte"
End Sub
