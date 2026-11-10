Attribute VB_Name = "modFelder"
Option Explicit
'===============================================================
' modFelder - globale Feldliste (config\fields.json), Schema 3
'
' Drei getrennte Feldklassen, die nie vermischt werden:
'   F001...  lokale Felder   (Quelle "lokales Feld", niedrigste Integritaet)
'   S001...  System-Felder   (Quelle "aus System", nur ueber 'Update Projekte')
'   B001...  BeKo-Felder     (Quelle "aus BeKo", Position im Word-Dokument, Phase 2)
' Standardfelder: F001 Kurzbeschreibung, F002 Bearbeiter.
' IDs werden nie wiederverwendet (Zaehler 'naechste_id' je Klasse F und S).
'
' fields.json:
'  { "schema_version": 3, "naechste_id": {"F": 3, "S": 1},
'    "felder": [ {"id":"F001","klasse":"F","key":"...","beschreibung":"...",
'                 "aktiv":true,"standard":true}, ... ] }
'===============================================================

Public Function FieldClass(ByVal id As String) As String
    Dim c As String
    id = Trim$(id)
    If Len(id) < 2 Then Exit Function
    c = UCase$(Left$(id, 1))
    If c = "F" Or c = "S" Or c = "B" Then
        If IsDigits(Mid$(id, 2)) Then FieldClass = c
    End If
End Function

Public Function SourceText(ByVal cls As String) As String
    Select Case UCase$(cls)
        Case "F": SourceText = SRC_F
        Case "S": SourceText = SRC_S
        Case "B": SourceText = SRC_B
    End Select
End Function

Public Function BFieldId(ByVal pos As Long) As String
    BFieldId = "B" & Format$(pos, "000")
End Function

Private Function IdNumber(ByVal id As String) As Long
    If Len(id) > 1 Then
        If IsDigits(Mid$(id, 2)) Then IdNumber = CLng(Mid$(id, 2))
    End If
End Function

Public Function IsFieldActive(ByVal fld As Object) As Boolean
    If fld Is Nothing Then Exit Function
    If Not fld.Exists("aktiv") Then
        IsFieldActive = True
    Else
        IsFieldActive = JBool(fld, "aktiv")
    End If
End Function

Public Function FieldsPath() As String
    FieldsPath = PathConfig() & "\fields.json"
End Function

Public Function FieldsLockPath() As String
    FieldsLockPath = PathLocks() & "\_fields.lock"
End Function

Public Function OldDataMessage() As String
    OldDataMessage = "Die Daten im Datenordner stammen aus einer älteren Version (Schema ungleich " & SCHEMA_VERSION & ")." & vbCrLf & vbCrLf & _
        "Bitte die Testdaten löschen (Ordner 'config', 'registry', 'projekte' und 'locks' im Datenpfad) oder im Admin Center einen neuen, leeren Datenordner wählen."
End Function

Private Sub AcquireFields()
    Dim held As String
    If Not TryAcquireLock(FieldsLockPath(), held) Then
        UserError "Die Feldliste wird gerade bearbeitet von: " & held & vbCrLf & "Bitte später erneut versuchen."
    End If
End Sub

'---------------------------------------------------------------
' Dokument anlegen / laden / schreiben
'---------------------------------------------------------------
Public Function NewFieldEntry(ByVal id As String, ByVal key As String, ByVal desc As String, ByVal isStandard As Boolean) As Object
    Dim e As Object
    Set e = NewDict()
    e("id") = id
    e("klasse") = FieldClass(id)
    e("key") = key
    e("beschreibung") = desc
    e("aktiv") = True
    e("standard") = isStandard
    Set NewFieldEntry = e
End Function

Private Function NewFieldsDoc() As Object
    Dim doc As Object, nid As Object, fl As Collection
    Set doc = NewDict()
    doc("schema_version") = SCHEMA_VERSION
    Set nid = NewDict()
    nid("F") = 3
    nid("S") = 1
    Set doc("naechste_id") = nid
    Set fl = NewList()
    fl.Add NewFieldEntry(FID_KURZ, KEY_KURZ, "Projektname", True)
    fl.Add NewFieldEntry(FID_BEARB, KEY_BEARB, "zuständiger Bearbeiter", True)
    Set doc("felder") = fl
    Set NewFieldsDoc = doc
End Function

Public Sub EnsureFieldsFile()
    If Not FileExists(FieldsPath()) Then WriteJsonFile FieldsPath(), NewFieldsDoc()
End Sub

Public Sub WriteFieldsDoc(ByVal doc As Object)
    WriteJsonFile FieldsPath(), doc
End Sub

' Liest fields.json, prueft das Schema und korrigiert die Zaehler.
Public Function LoadFieldsDoc() As Object
    Dim doc As Object, nid As Object, fl As Collection, it As Variant
    Dim cls As String, nr As Long
    EnsureFieldsFile
    Set doc = ReadJsonFile(FieldsPath())
    If Not doc.Exists("naechste_id") Or Not doc.Exists("felder") Then UserError OldDataMessage()
    If JStr(doc, "schema_version") <> CStr(SCHEMA_VERSION) Then UserError OldDataMessage()
    If Not IsObject(doc("naechste_id")) Or Not IsObject(doc("felder")) Then UserError OldDataMessage()
    Set nid = doc("naechste_id")
    If Not nid.Exists("F") Then nid("F") = 3
    If Not nid.Exists("S") Then nid("S") = 1
    Set fl = doc("felder")
    For Each it In fl
        cls = FieldClass(JStr(it, "id"))
        If cls = "F" Or cls = "S" Then
            nr = IdNumber(JStr(it, "id"))
            If nr >= CLng(nid(cls)) Then nid(cls) = nr + 1
        End If
    Next it
    If FindFieldById(fl, FID_KURZ) Is Nothing Then fl.Add NewFieldEntry(FID_KURZ, KEY_KURZ, "Projektname", True)
    If FindFieldById(fl, FID_BEARB) Is Nothing Then fl.Add NewFieldEntry(FID_BEARB, KEY_BEARB, "zuständiger Bearbeiter", True)
    Set LoadFieldsDoc = doc
End Function

Public Function FindFieldById(ByVal fl As Collection, ByVal id As String) As Object
    Dim it As Variant
    id = UCase$(Trim$(id))
    For Each it In fl
        If UCase$(JStr(it, "id")) = id Then
            Set FindFieldById = it
            Exit Function
        End If
    Next it
End Function

' Sucht ein Feld einer Klasse ueber den Key (ohne Gross/Kleinschreibung, getrimmt).
Public Function FindFieldByKey(ByVal fl As Collection, ByVal cls As String, ByVal key As String) As Object
    Dim it As Variant, k As String
    k = LCase$(Trim$(key))
    For Each it In fl
        If FieldClass(JStr(it, "id")) = cls Then
            If LCase$(Trim$(JStr(it, "key"))) = k Then
                Set FindFieldByKey = it
                Exit Function
            End If
        End If
    Next it
End Function

Public Function TakeNextId(ByVal doc As Object, ByVal cls As String) As String
    Dim nid As Object, n As Long
    Set nid = doc("naechste_id")
    n = CLng(nid(cls))
    TakeNextId = cls & Format$(n, "000")
    nid(cls) = n + 1
End Function

' Aktive Felder in Anzeigereihenfolge: F, dann S, dann B (je nach ID-Nummer).
Public Function ActiveFieldsOrdered(ByVal doc As Object, Optional ByVal includeInactive As Boolean = False) As Collection
    Dim res As Collection, cls As Variant, it As Variant, arr() As Object, n As Long
    Dim i As Long, j As Long, t As Object
    Set res = New Collection
    For Each cls In Array("F", "S", "B")
        n = 0
        ReDim arr(0 To doc("felder").Count)
        For Each it In doc("felder")
            If FieldClass(JStr(it, "id")) = CStr(cls) Then
                If includeInactive Or IsFieldActive(it) Then
                    Set arr(n) = it
                    n = n + 1
                End If
            End If
        Next it
        For i = 1 To n - 1
            Set t = arr(i)
            j = i - 1
            Do While j >= 0
                If IdNumber(JStr(arr(j), "id")) > IdNumber(JStr(t, "id")) Then
                    Set arr(j + 1) = arr(j)
                    j = j - 1
                Else
                    Exit Do
                End If
            Loop
            Set arr(j + 1) = t
        Next i
        For i = 0 To n - 1
            res.Add arr(i)
        Next i
    Next cls
    Set ActiveFieldsOrdered = res
End Function

'---------------------------------------------------------------
' Import ('Update Projekte'): Spaltentitel -> S-Felder
' keys: Spaltentitel (ohne ID-Spalte und ohne Standardfelder)
' idOf: lcase(Titel) -> S-ID   keyOf: S-ID -> Key
'---------------------------------------------------------------
Public Sub ResolveImportFields(ByVal keys As Collection, ByRef idOf As Object, ByRef keyOf As Object, _
                               ByRef nNew As Long, ByRef nReact As Long)
    Dim doc As Object, fl As Collection, k As Variant, fld As Object, changed As Boolean
    Dim n As Long, s As String, d As String, id As String
    Set idOf = NewDict()
    Set keyOf = NewDict()
    nNew = 0
    nReact = 0
    If keys.Count = 0 Then Exit Sub
    AcquireFields
    On Error GoTo EH
    Set doc = LoadFieldsDoc()
    Set fl = doc("felder")
    For Each k In keys
        Set fld = FindFieldByKey(fl, "S", CStr(k))
        If fld Is Nothing Then
            id = TakeNextId(doc, "S")
            Set fld = NewFieldEntry(id, Trim$(CStr(k)), "", False)
            fl.Add fld
            nNew = nNew + 1
            changed = True
        ElseIf Not IsFieldActive(fld) Then
            fld("aktiv") = True
            nReact = nReact + 1
            changed = True
        End If
        idOf(LCase$(Trim$(CStr(k)))) = JStr(fld, "id")
        keyOf(JStr(fld, "id")) = JStr(fld, "key")
    Next k
    If changed Then WriteFieldsDoc doc
    ReleaseLock FieldsLockPath()
    Exit Sub
EH:
    n = Err.Number: s = Err.Source: d = Err.Description
    ReleaseLock FieldsLockPath()
    Err.Raise n, s, d
End Sub

'---------------------------------------------------------------
' Neues lokales Feld (aus Blatt 'Projekt'); gleicher Key wie ein
' inaktives F-Feld -> dieses wird reaktiviert.
'---------------------------------------------------------------
Public Function RegisterNewFField(ByVal key As String, ByRef reactivated As Boolean) As String
    Dim doc As Object, fl As Collection, fld As Object, id As String
    Dim n As Long, s As String, d As String
    reactivated = False
    key = Trim$(key)
    AcquireFields
    On Error GoTo EH
    Set doc = LoadFieldsDoc()
    Set fl = doc("felder")
    Set fld = FindFieldByKey(fl, "F", key)
    If fld Is Nothing Then
        id = TakeNextId(doc, "F")
        fl.Add NewFieldEntry(id, key, "", False)
        WriteFieldsDoc doc
    ElseIf Not IsFieldActive(fld) Then
        fld("aktiv") = True
        reactivated = True
        id = JStr(fld, "id")
        WriteFieldsDoc doc
    Else
        id = JStr(fld, "id")
    End If
    ReleaseLock FieldsLockPath()
    RegisterNewFField = id
    Exit Function
EH:
    n = Err.Number: s = Err.Source: d = Err.Description
    ReleaseLock FieldsLockPath()
    Err.Raise n, s, d
End Function

'---------------------------------------------------------------
' Blatt 'Felder anpassen' speichern.
' rows: Collection von Array(id, name, beschreibung, zeilennummer)
' loadedKey / loadedDesc: Stand beim Laden des Blatts (ID -> Text).
' Es werden nur Aenderungen gegenueber diesem Stand angewendet.
'---------------------------------------------------------------
Public Sub ApplyFieldRows(ByVal rows As Collection, ByVal loadedKey As Object, ByVal loadedDesc As Object, _
                          ByRef renamed As Object, ByRef nNew As Long, ByRef nDeact As Long, _
                          ByRef nReact As Long, ByRef warnings As Collection, ByRef nDesc As Long)
    Dim doc As Object, fl As Collection, row As Variant, fld As Object, other As Object, it As Variant
    Dim id As String, nm As String, ds As String, present As Object
    Dim changed As Boolean, n As Long, s As String, d As String
    Set renamed = NewDict()
    Set warnings = New Collection
    nNew = 0
    nDeact = 0
    nReact = 0
    nDesc = 0
    AcquireFields
    On Error GoTo EH
    Set doc = LoadFieldsDoc()
    Set fl = doc("felder")
    Set present = NewDict()
    For Each row In rows
        id = Trim$(CStr(row(0)))
        nm = Trim$(CStr(row(1)))
        ds = Trim$(CStr(row(2)))
        If Len(id) > 0 Then
            present(UCase$(id)) = True
            Set fld = FindFieldById(fl, id)
            If fld Is Nothing Then
                warnings.Add "Zeile " & row(3) & ": Feld-ID " & id & " gibt es nicht mehr - ignoriert."
            ElseIf FieldClass(id) = "F" Then
                If loadedKey.Exists(UCase$(id)) Then
                    If nm <> CStr(loadedKey(UCase$(id))) Then
                        If JBool(fld, "standard") Then
                            warnings.Add id & ": Standardfeld, der Name kann nicht geändert werden."
                        ElseIf JStr(fld, "key") <> CStr(loadedKey(UCase$(id))) Then
                            warnings.Add id & ": wurde inzwischen von einem anderen Client umbenannt - nicht überschrieben."
                        Else
                            Set other = FindFieldByKey(fl, "F", nm)
                            If other Is Nothing Then
                                fld("key") = nm
                                renamed(UCase$(id)) = nm
                                changed = True
                            ElseIf JStr(other, "id") = JStr(fld, "id") Then
                                fld("key") = nm
                                renamed(UCase$(id)) = nm
                                changed = True
                            Else
                                warnings.Add id & ": der Name '" & nm & "' ist bereits vergeben (" & JStr(other, "id") & ")."
                            End If
                        End If
                    End If
                End If
                If loadedDesc.Exists(UCase$(id)) Then
                    If ds <> CStr(loadedDesc(UCase$(id))) Then
                        If JStr(fld, "beschreibung") = CStr(loadedDesc(UCase$(id))) Then
                            fld("beschreibung") = ds
                            nDesc = nDesc + 1
                            changed = True
                        Else
                            warnings.Add id & ": Beschreibung wurde inzwischen von einem anderen Client geändert - nicht überschrieben."
                        End If
                    End If
                End If
            Else
                If loadedKey.Exists(UCase$(id)) Then
                    If nm <> CStr(loadedKey(UCase$(id))) Then warnings.Add id & ": " & SourceText(FieldClass(id)) & " ist schreibgeschützt - Änderung ignoriert."
                End If
            End If
        End If
    Next row
    ' Entfernte Zeilen -> Feld inaktiv
    For Each it In fl
        id = JStr(it, "id")
        If IsFieldActive(it) And loadedKey.Exists(UCase$(id)) And Not present.Exists(UCase$(id)) Then
            If JBool(it, "standard") Then
                warnings.Add id & ": Standardfeld kann nicht entfernt werden."
            Else
                it("aktiv") = False
                nDeact = nDeact + 1
                changed = True
            End If
        End If
    Next it
    ' Neue Zeilen -> neue (oder reaktivierte) F-Felder
    For Each row In rows
        If Len(Trim$(CStr(row(0)))) = 0 And Len(Trim$(CStr(row(1)))) > 0 Then
            nm = Trim$(CStr(row(1)))
            ds = Trim$(CStr(row(2)))
            Set other = FindFieldByKey(fl, "F", nm)
            If other Is Nothing Then
                id = TakeNextId(doc, "F")
                fl.Add NewFieldEntry(id, nm, ds, False)
                nNew = nNew + 1
                changed = True
            ElseIf Not IsFieldActive(other) Then
                other("aktiv") = True
                If Len(ds) > 0 Then other("beschreibung") = ds
                nReact = nReact + 1
                changed = True
            Else
                warnings.Add "Zeile " & row(3) & ": '" & nm & "' gibt es schon (" & JStr(other, "id") & ") - nicht neu angelegt."
            End If
        End If
    Next row
    If changed Then WriteFieldsDoc doc
    ReleaseLock FieldsLockPath()
    Exit Sub
EH:
    n = Err.Number: s = Err.Source: d = Err.Description
    ReleaseLock FieldsLockPath()
    Err.Raise n, s, d
End Sub

'---------------------------------------------------------------
' Key-Kopien in einem Projekt an die globale Feldliste angleichen
'---------------------------------------------------------------
Public Sub RefreshKeyCopies(ByVal felder As Object)
    Dim doc As Object, fl As Collection, k As Variant, fld As Object, e As Object
    If felder.Count = 0 Then Exit Sub
    Set doc = LoadFieldsDoc()
    Set fl = doc("felder")
    For Each k In felder.Keys
        Set fld = FindFieldById(fl, CStr(k))
        If Not fld Is Nothing Then
            If IsObject(felder(CStr(k))) Then
                Set e = felder(CStr(k))
                e("key") = JStr(fld, "key")
            End If
        End If
    Next k
End Sub

' Schreibt umbenannte Keys in alle Projekte (gesperrte werden uebersprungen
' und beim naechsten Speichern per RefreshKeyCopies angeglichen).
Public Function PropagateRenames(ByVal renamed As Object, ByRef nUpdated As Long) As Collection
    Dim skipped As Collection, ids As Collection, id As Variant, lp As String, held As String
    Dim r As Long
    Set skipped = New Collection
    nUpdated = 0
    If renamed.Count > 0 Then
        Set ids = AllProjectIds()
        For Each id In ids
            lp = LockPath(CStr(id))
            If HasLock(lp) Then
                skipped.Add CStr(id) & " (hier geöffnet)"
            ElseIf Not TryAcquireLock(lp, held) Then
                skipped.Add CStr(id) & " (" & held & ")"
            Else
                r = RenameInProject(CStr(id), renamed)
                If r > 0 Then
                    nUpdated = nUpdated + 1
                ElseIf r < 0 Then
                    skipped.Add CStr(id) & " (nicht lesbar)"
                End If
            End If
        Next id
    End If
    Set PropagateRenames = skipped
End Function

' Sperre liegt bereits bei uns und wird hier freigegeben. 1 = geaendert, 0 = nichts, -1 = Fehler
Private Function RenameInProject(ByVal id As String, ByVal renamed As Object) As Long
    Dim p As Object, felder As Object, k As Variant, e As Object, changed As Boolean
    On Error GoTo EH
    Set p = ReadJsonFile(ProjectFile(id))
    CheckProjectSchema p
    If p.Exists("felder") Then
        Set felder = p("felder")
        For Each k In renamed.Keys
            If felder.Exists(CStr(k)) Then
                Set e = felder(CStr(k))
                e("key") = CStr(renamed(k))
                changed = True
            End If
        Next k
    End If
    If changed Then
        WriteJsonFile ProjectFile(id), p
        RenameInProject = 1
    End If
    ReleaseLock LockPath(id)
    Exit Function
EH:
    RenameInProject = -1
    ReleaseLock LockPath(id)
End Function

'---------------------------------------------------------------
' Admin Center: 'purge inactive Fields'
' Loescht inaktive Felder (F, S, B) endgueltig samt Werten in allen
' Projekten. Alles oder nichts: ist ein betroffenes Projekt gesperrt,
' bleibt alles unveraendert. IDs werden nie wieder vergeben.
'---------------------------------------------------------------
Public Sub Admin_PurgeInaktiveFelder()
    Dim doc As Object, fl As Collection, it As Variant, inact As Collection
    Dim ids As Collection, id As Variant, usage As Object, pj As Object, felder As Object
    Dim lines As Collection, txt As String, ans As String, sel As Object, tok As Variant
    Dim affected As Collection, locked As Collection, busy As Collection, lp As String, held As String
    Dim nProj As Long, nVal As Long, i As Long, unreadable As Long, fid As Variant, cnt As Long
    Dim nDel As Long, nChanged As Long, tmp As String, n As Long
    On Error GoTo EH
    If IsProjectOpen() Then UserError "Bitte zuerst das geöffnete Projekt schließen."
    InitDataRoot
    Set doc = LoadFieldsDoc()
    Set inact = New Collection
    For Each it In ActiveFieldsOrdered(doc, True)
        If Not IsFieldActive(it) And Not JBool(it, "standard") Then inact.Add it
    Next it
    If inact.Count = 0 Then
        MsgBox "Es gibt keine inaktiven Felder.", vbInformation, APP_TITLE
        Exit Sub
    End If
    ' Verwendung zaehlen
    Set usage = NewDict()
    For Each it In inact
        usage(UCase$(JStr(it, "id"))) = 0
    Next it
    Application.StatusBar = "Zähle Verwendung der Felder ..."
    Set ids = AllProjectIds()
    For Each id In ids
        Set pj = Nothing
        On Error Resume Next
        Set pj = ReadJsonFile(ProjectFile(CStr(id)))
        If Err.Number <> 0 Then unreadable = unreadable + 1
        Err.Clear
        On Error GoTo EH
        If Not pj Is Nothing Then
            If pj.Exists("felder") Then
                Set felder = pj("felder")
                For Each it In inact
                    If felder.Exists(JStr(it, "id")) Then usage(UCase$(JStr(it, "id"))) = usage(UCase$(JStr(it, "id"))) + 1
                Next it
            End If
        End If
    Next id
    Application.StatusBar = False
    Set lines = New Collection
    For Each it In inact
        lines.Add JStr(it, "id") & "  " & Left$(JStr(it, "key"), 30) & "  (" & usage(UCase$(JStr(it, "id"))) & " Projekt(e))"
        WriteLog "Inaktives Feld: " & JStr(it, "id") & " '" & JStr(it, "key") & "' in " & usage(UCase$(JStr(it, "id"))) & " Projekt(en)"
    Next it
    txt = "Inaktive Felder (Löschen ist endgültig):" & vbCrLf & vbCrLf & JoinLines(lines, 14) & vbCrLf & _
          "Feld-IDs eingeben (mit Komma getrennt) oder ALLE:"
    ans = Trim$(InputBox(txt, APP_TITLE & " - purge inactive Fields"))
    If Len(ans) = 0 Then Exit Sub
    Set sel = NewDict()
    If UCase$(ans) = "ALLE" Then
        For Each it In inact
            sel(UCase$(JStr(it, "id"))) = True
        Next it
    Else
        ans = Replace(Replace(Replace(ans, ";", ","), " ", ","), vbTab, ",")
        For Each tok In Split(ans, ",")
            If Len(Trim$(CStr(tok))) > 0 Then
                If Not usage.Exists(UCase$(Trim$(CStr(tok)))) Then UserError "'" & Trim$(CStr(tok)) & "' ist kein inaktives Feld (Standardfelder und aktive Felder können nicht gelöscht werden)."
                sel(UCase$(Trim$(CStr(tok)))) = True
            End If
        Next tok
    End If
    If sel.Count = 0 Then Exit Sub
    nVal = 0
    For Each fid In sel.Keys
        nVal = nVal + usage(CStr(fid))
    Next fid
    If MsgBox(sel.Count & " Feld(er) endgültig löschen?" & vbCrLf & nVal & " gespeicherte(r) Wert(e) in den Projekten gehen verloren." & vbCrLf & _
              "Die Feld-IDs werden nie wieder vergeben." & vbCrLf & vbCrLf & "Das kann nicht rückgängig gemacht werden.", _
              vbYesNo + vbExclamation + vbDefaultButton2, APP_TITLE) <> vbYes Then Exit Sub

    ' betroffene Projekte sperren (alles oder nichts)
    Set affected = New Collection
    Set locked = New Collection
    Set busy = New Collection
    For Each id In ids
        Set pj = Nothing
        On Error Resume Next
        Set pj = ReadJsonFile(ProjectFile(CStr(id)))
        Err.Clear
        On Error GoTo EH
        If Not pj Is Nothing Then
            If pj.Exists("felder") Then
                cnt = 0
                For Each fid In sel.Keys
                    If pj("felder").Exists(CStr(fid)) Then cnt = cnt + 1
                Next fid
                If cnt > 0 Then affected.Add CStr(id)
            End If
        End If
    Next id
    If affected.Count > 200 Then UserError "Zu viele betroffene Projekte (" & affected.Count & ") für einen Durchlauf. Bitte weniger Felder auswählen."
    For Each id In affected
        lp = LockPath(CStr(id))
        If Not TryAcquireLock(lp, held) Then
            busy.Add CStr(id) & " (" & held & ")"
        Else
            locked.Add lp
        End If
    Next id
    If busy.Count > 0 Then
        ReleaseLockList locked
        WriteLog "Purge abgebrochen, gesperrte Projekte: " & busy.Count
        UserError "Nichts gelöscht: " & busy.Count & " betroffene(s) Projekt(e) ist/sind gesperrt:" & vbCrLf & vbCrLf & JoinLines(busy, 10) & vbCrLf & "Bitte später erneut versuchen."
    End If

    ' Werte entfernen
    Application.StatusBar = "Lösche Werte in den Projekten ..."
    For Each id In affected
        Set pj = ReadJsonFile(ProjectFile(CStr(id)))
        Set felder = pj("felder")
        For Each fid In sel.Keys
            If felder.Exists(CStr(fid)) Then
                felder.Remove CStr(fid)
                nDel = nDel + 1
            End If
        Next fid
        WriteJsonFile ProjectFile(CStr(id)), pj
        nChanged = nChanged + 1
    Next id
    ReleaseLockList locked
    Set locked = Nothing
    Application.StatusBar = False

    ' Felder aus fields.json entfernen
    PurgeFromFieldsFile sel
    WriteLog "Purge inactive Fields: " & sel.Count & " Feld(er), " & nDel & " Wert(e) in " & nChanged & " Projekt(en) gelöscht"
    MsgBox sel.Count & " Feld(er) endgültig gelöscht." & vbCrLf & nDel & " Wert(e) in " & nChanged & " Projekt(en) entfernt.", vbInformation, APP_TITLE
    Exit Sub
EH:
    n = Err.Number
    tmp = Err.Description
    On Error Resume Next
    Application.StatusBar = False
    If Not locked Is Nothing Then ReleaseLockList locked
    Err.Clear
    Err.Raise n, APP_TITLE, tmp
    HandleError "Admin_PurgeInaktiveFelder"
End Sub

Private Sub ReleaseLockList(ByVal locked As Collection)
    Dim lp As Variant
    For Each lp In locked
        ReleaseLock CStr(lp)
    Next lp
End Sub

Private Sub PurgeFromFieldsFile(ByVal sel As Object)
    Dim doc As Object, fl As Collection, nfl As Collection, it As Variant
    Dim n As Long, s As String, d As String
    AcquireFields
    On Error GoTo EH
    Set doc = LoadFieldsDoc()
    Set nfl = NewList()
    For Each it In doc("felder")
        If sel.Exists(UCase$(JStr(it, "id"))) And Not IsFieldActive(it) And Not JBool(it, "standard") Then
            ' entfernt
        Else
            nfl.Add it
        End If
    Next it
    Set doc("felder") = nfl
    WriteFieldsDoc doc
    ReleaseLock FieldsLockPath()
    Exit Sub
EH:
    n = Err.Number: s = Err.Source: d = Err.Description
    ReleaseLock FieldsLockPath()
    Err.Raise n, s, d
End Sub
