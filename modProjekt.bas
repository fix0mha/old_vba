Attribute VB_Name = "modProjekt"
Option Explicit
'===============================================================
' modProjekt - Projektlogik: anlegen, oeffnen, speichern,
'              entfernen (nur Verweis), vollstaendig loeschen
'
' Aufbau auf dem Netzlaufwerk:
'   registry\<ID>.ref            Verweis auf das Projekt
'   locks\<ID>.lock              Sperre waehrend der Bearbeitung
'   projekte\<ID>\project.json   alle Daten des Projekts (Schema 3)
'
' project.json:  "felder": { "F001": {"key":"...","wert":"...","quelle":"..."} }
' Die Projekt-ID (= Ticket ID = Ordnername) ist kein Feld.
'===============================================================
Private mTaskDirty As Boolean   ' Aufgabenstatus geaendert, noch nicht gespeichert

Public Const PR_OK As Long = 0
Public Const PR_UNCHANGED As Long = 1
Public Const PR_CONFLICT As Long = 2
Public Const PR_LOCKED As Long = -1
Public Const PR_ERROR As Long = -2

Private gProject As Object
Private gId As String

'---------------------------------------------------------------
' Pfade und Pruefungen
'---------------------------------------------------------------
Public Function IsValidProjectId(ByVal id As String) As Boolean
    Dim i As Long, c As String
    If Len(id) < 1 Or Len(id) > 64 Then Exit Function
    For i = 1 To Len(id)
        c = Mid$(id, i, 1)
        If Not (c Like "[A-Za-z0-9]" Or c = "-" Or c = "_") Then Exit Function
    Next i
    If InStr("|CON|PRN|AUX|NUL|COM1|COM2|COM3|COM4|COM5|COM6|COM7|COM8|COM9|LPT1|LPT2|LPT3|LPT4|LPT5|LPT6|LPT7|LPT8|LPT9|", "|" & UCase$(id) & "|") > 0 Then Exit Function
    IsValidProjectId = True
End Function

Private Sub CheckId(ByVal id As String)
    If Not IsValidProjectId(id) Then
        UserError "Ungültige Ticket ID: '" & id & "'" & vbCrLf & "Erlaubt sind Buchstaben, Ziffern, '-' und '_' (max. 64 Zeichen)."
    End If
End Sub

Public Function ProjectDir(ByVal id As String) As String
    ProjectDir = PathProjects() & "\" & id
End Function

Public Function ProjectFile(ByVal id As String) As String
    ProjectFile = ProjectDir(id) & "\project.json"
End Function

Public Function RefPath(ByVal id As String) As String
    RefPath = PathRegistry() & "\" & id & ".ref"
End Function

Public Function LockPath(ByVal id As String) As String
    LockPath = PathLocks() & "\" & id & ".lock"
End Function

' Text fuer 'Gesperrt von' bei eigener Sperre (Windows-Anmeldename)
Public Function OwnLockText() As String
    OwnLockText = "Ich (" & Environ$("USERNAME") & ")"
End Function

Public Function TaskDirty() As Boolean
    TaskDirty = mTaskDirty
End Function

Public Function IsDigits(ByVal s As String) As Boolean
    IsDigits = (Len(s) > 0) And Not (s Like "*[!0-9]*")
End Function

' Alle Projektordner mit project.json (unabhaengig vom Verweis).
Public Function AllProjectIds() As Collection
    Dim res As Collection, nm As Variant
    Set res = New Collection
    For Each nm In ListSubFolderNames(PathProjects())
        If IsValidProjectId(CStr(nm)) Then
            If FileExists(ProjectFile(CStr(nm))) Then res.Add CStr(nm)
        End If
    Next nm
    Set AllProjectIds = res
End Function

'---------------------------------------------------------------
' Zustand des geoeffneten Projekts
'---------------------------------------------------------------
Public Function IsProjectOpen() As Boolean
    IsProjectOpen = Not (gProject Is Nothing)
End Function

Public Function CurrentId() As String
    CurrentId = gId
End Function

Public Function CurrentProject() As Object
    Set CurrentProject = gProject
End Function

Public Sub RequireOpenProject()
    If gProject Is Nothing Then UserError "Es ist kein Projekt geöffnet."
End Sub

'---------------------------------------------------------------
' Struktur
'---------------------------------------------------------------
Private Function NewProjectDict(ByVal id As String) As Object
    Dim p As Object, st As Object, i As Long
    Set p = NewDict()
    p("schema_version") = SCHEMA_VERSION
    p("projekt_id") = id
    p("angelegt") = IsoNow()
    p("angelegt_von") = UserTag()
    p("geaendert") = IsoNow()
    p("geaendert_von") = UserTag()
    Set st = NewDict()
    For i = 1 To TASK_COUNT
        st("aufgabe_" & i) = False
    Next i
    Set p("status") = st
    Set p("felder") = NewDict()
    Set NewProjectDict = p
End Function

' Meldet Daten aus einer aelteren Version (Schema ungleich 3).
Public Sub CheckProjectSchema(ByVal p As Object)
    If JStr(p, "schema_version") <> CStr(SCHEMA_VERSION) Or p.Exists("daten") Or p.Exists("zusatzfelder") Or Not p.Exists("felder") Then
        UserError OldDataMessage()
    End If
End Sub

Public Sub EnsureStructure(ByVal p As Object)
    Dim st As Object, i As Long
    If Not p.Exists("status") Then Set p("status") = NewDict()
    Set st = p("status")
    For i = 1 To TASK_COUNT
        If Not st.Exists("aufgabe_" & i) Then st("aufgabe_" & i) = False
    Next i
    For i = TASK_COUNT + 1 To 9
        If st.Exists("aufgabe_" & i) Then st.Remove "aufgabe_" & i
    Next i
    If Not p.Exists("felder") Then Set p("felder") = NewDict()
End Sub

Public Function FeldWert(ByVal felder As Object, ByVal fid As String) As String
    If felder Is Nothing Then Exit Function
    If Not felder.Exists(fid) Then Exit Function
    If IsObject(felder(fid)) Then FeldWert = JStr(felder(fid), "wert")
End Function

Public Function FeldQuelle(ByVal felder As Object, ByVal fid As String) As String
    If felder Is Nothing Then Exit Function
    If Not felder.Exists(fid) Then Exit Function
    If IsObject(felder(fid)) Then FeldQuelle = JStr(felder(fid), "quelle")
End Function

Public Sub SetFeld(ByVal felder As Object, ByVal fid As String, ByVal key As String, ByVal wert As String, ByVal quelle As String)
    Dim e As Object
    Set e = NewDict()
    e("key") = key
    e("wert") = wert
    e("quelle") = quelle
    Set felder(fid) = e
End Sub

Private Sub WriteRef(ByVal id As String)
    Dim r As Object
    Set r = NewDict()
    r("schema_version") = SCHEMA_VERSION
    r("projekt_id") = id
    r("pfad") = "projekte/" & id
    r("angelegt") = IsoNow()
    r("angelegt_von") = UserTag()
    WriteJsonFile RefPath(id), r
End Sub

'---------------------------------------------------------------
' Anlegen / Verknuepfen
'---------------------------------------------------------------
Public Sub CreateProject(ByVal id As String, Optional ByVal kurz As String = "")
    Dim held As String, n As Long, s As String, d As String, pj As Object
    id = Trim$(id)
    CheckId id
    InitDataRoot
    LoadFieldsDoc
    If Not TryAcquireLock(LockPath(id), held) Then
        UserError "Das Projekt '" & id & "' wird gerade bearbeitet von: " & held
    End If
    On Error GoTo EH
    If FileExists(RefPath(id)) Then UserError "Das Projekt '" & id & "' existiert bereits."
    If FolderExists(ProjectDir(id)) Then
        UserError "Der Ordner für '" & id & "' existiert bereits, aber es gibt keinen Verweis." & vbCrLf & _
                  "Bitte im Admin Center 'Projekt wiederherstellen' verwenden."
    End If
    EnsureFolder ProjectDir(id)
    Set pj = NewProjectDict(id)
    If Len(Trim$(kurz)) > 0 Then SetFeld pj("felder"), FID_KURZ, KEY_KURZ, Trim$(kurz), SRC_F
    WriteJsonFile ProjectFile(id), pj
    WriteRef id
    WriteLog "Projekt angelegt: " & id
    ReleaseLock LockPath(id)
    Exit Sub
EH:
    n = Err.Number: s = Err.Source: d = Err.Description
    ReleaseLock LockPath(id)
    Err.Raise n, s, d
End Sub

' Stellt den Verweis zu einem vorhandenen Projektordner wieder her.
Public Sub LinkExistingProject(ByVal id As String)
    Dim held As String, n As Long, s As String, d As String, p As Object
    id = Trim$(id)
    CheckId id
    InitDataRoot
    If FileExists(RefPath(id)) Then UserError "Zu '" & id & "' gibt es bereits einen Verweis."
    If Not FileExists(ProjectFile(id)) Then UserError "Im Ordner '" & ProjectDir(id) & "' wurde keine project.json gefunden."
    If Not TryAcquireLock(LockPath(id), held) Then UserError "Das Projekt wird gerade bearbeitet von: " & held
    On Error GoTo EH
    Set p = ReadJsonFile(ProjectFile(id))
    CheckProjectSchema p
    If JStr(p, "projekt_id") <> id Then UserError "Die project.json gehört zu einer anderen Ticket ID (" & JStr(p, "projekt_id") & ")."
    WriteRef id
    WriteLog "Projekt verknüpft: " & id
    ReleaseLock LockPath(id)
    Exit Sub
EH:
    n = Err.Number: s = Err.Source: d = Err.Description
    ReleaseLock LockPath(id)
    Err.Raise n, s, d
End Sub

'---------------------------------------------------------------
' Liste
'---------------------------------------------------------------
Public Function ListProjects() As Collection
    Dim names As Collection, ids() As String, n As Long, i As Long, j As Long, t As String
    Dim res As Collection, d As Object, ref As Object, info As String, nm As Variant, nm2 As String
    Dim pj As Object
    Set res = New Collection
    Set names = ListFileNames(PathRegistry(), ".ref")
    n = names.Count
    If n = 0 Then
        Set ListProjects = res
        Exit Function
    End If
    ReDim ids(0 To n - 1)
    i = 0
    For Each nm In names
        nm2 = CStr(nm)
        ids(i) = Left$(nm2, Len(nm2) - 4)
        i = i + 1
    Next nm
    For i = 1 To n - 1
        t = ids(i)
        j = i - 1
        Do While j >= 0
            If StrComp(t, ids(j), vbTextCompare) < 0 Then
                ids(j + 1) = ids(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        ids(j + 1) = t
    Next i
    For i = 0 To n - 1
        If IsValidProjectId(ids(i)) Then
            Set d = NewDict()
            d("id") = ids(i)
            d("angelegt") = ""
            d("angelegt_von") = ""
            If FolderExists(ProjectDir(ids(i))) Then d("ordner") = "ok" Else d("ordner") = "FEHLT"
            Set ref = Nothing
            On Error Resume Next
            Set ref = ReadJsonFile(RefPath(ids(i)))
            If Err.Number = 0 Then
                d("angelegt") = JStr(ref, "angelegt")
                d("angelegt_von") = JStr(ref, "angelegt_von")
            Else
                d("ordner") = "Verweis defekt"
            End If
            Err.Clear
            On Error GoTo 0
            d("kurzbeschreibung") = ""
            d("bearbeiter") = ""
            Set pj = Nothing
            On Error Resume Next
            Set pj = ReadJsonFile(ProjectFile(ids(i)))
            If Err.Number = 0 Then
                If pj.Exists("felder") Then
                    d("kurzbeschreibung") = FeldWert(pj("felder"), FID_KURZ)
                    d("bearbeiter") = FeldWert(pj("felder"), FID_BEARB)
                End If
            End If
            Err.Clear
            On Error GoTo 0
            info = ""
            If HasLock(LockPath(ids(i))) Then
                d("gesperrt") = OwnLockText()
            ElseIf IsLockedByOther(LockPath(ids(i)), info) Then
                d("gesperrt") = info
            Else
                d("gesperrt") = ""
            End If
            res.Add d
        End If
    Next i
    Set ListProjects = res
End Function

'---------------------------------------------------------------
' Kurzbeschreibung / Bearbeiter direkt in project.json aendern
' (Blatt 'alle Projekte', Button 'Projekt speichern').
' exp... = Wert beim Laden der Liste. Ohne force wird bei einer
' Aenderung durch Dritte PR_CONFLICT gemeldet (nichts geschrieben).
'---------------------------------------------------------------
Public Function SetStandardFields(ByVal id As String, ByVal kurz As String, ByVal bearb As String, _
                                  ByVal expKurz As String, ByVal expBearb As String, _
                                  ByVal force As Boolean, ByRef info As String) As Long
    Dim held As String, lp As String, p As Object, felder As Object
    Dim curK As String, curB As String, chK As Boolean, chB As Boolean, wr As Boolean
    Dim n As Long, d As String
    info = ""
    If Not (gProject Is Nothing) Then
        If StrComp(gId, id, vbTextCompare) = 0 Then
            info = "ist im Blatt 'Projekt' geöffnet"
            SetStandardFields = PR_LOCKED
            Exit Function
        End If
    End If
    lp = LockPath(id)
    If Not TryAcquireLock(lp, held) Then
        info = "gesperrt von " & held
        SetStandardFields = PR_LOCKED
        Exit Function
    End If
    On Error GoTo EH
    Set p = ReadJsonFile(ProjectFile(id))
    CheckProjectSchema p
    EnsureStructure p
    Set felder = p("felder")
    curK = FeldWert(felder, FID_KURZ)
    curB = FeldWert(felder, FID_BEARB)
    chK = (kurz <> expKurz)
    chB = (bearb <> expBearb)
    If Not force Then
        If (chK And curK <> expKurz) Or (chB And curB <> expBearb) Then
            info = "wurde inzwischen geändert (gespeichert: '" & Left$(curK, 30) & "' / '" & Left$(curB, 30) & "')"
            ReleaseLock lp
            SetStandardFields = PR_CONFLICT
            Exit Function
        End If
    End If
    If chK And (curK <> kurz Or FeldQuelle(felder, FID_KURZ) <> SRC_F) Then
        SetFeld felder, FID_KURZ, KEY_KURZ, kurz, SRC_F
        wr = True
    End If
    If chB And (curB <> bearb Or FeldQuelle(felder, FID_BEARB) <> SRC_F) Then
        SetFeld felder, FID_BEARB, KEY_BEARB, bearb, SRC_F
        wr = True
    End If
    If wr Then
        RefreshKeyCopies felder
        p("geaendert") = IsoNow()
        p("geaendert_von") = UserTag()
        WriteJsonFile ProjectFile(id), p
        WriteLog "Liste gespeichert: " & id
        SetStandardFields = PR_OK
    Else
        SetStandardFields = PR_UNCHANGED
    End If
    ReleaseLock lp
    Exit Function
EH:
    n = Err.Number
    d = Err.Description
    ReleaseLock lp
    info = "[" & n & "] " & Left$(d, 120)
    SetStandardFields = PR_ERROR
End Function

'---------------------------------------------------------------
' Oeffnen / Speichern / Schliessen
'---------------------------------------------------------------
Public Sub OpenProject(ByVal id As String)
    Dim held As String, p As Object, n As Long, s As String, d As String
    id = Trim$(id)
    CheckId id
    If Not (gProject Is Nothing) Then UserError "Es ist bereits das Projekt '" & gId & "' geöffnet. Bitte zuerst schließen."
    InitDataRoot
    LoadFieldsDoc
    If Not FileExists(RefPath(id)) Then UserError "Es gibt keinen Verweis auf ein Projekt '" & id & "'."
    If Not FileExists(ProjectFile(id)) Then UserError "Die Datei project.json fehlt im Projektordner von '" & id & "'."
    If Not TryAcquireLock(LockPath(id), held) Then
        UserError "Das Projekt '" & id & "' wird gerade bearbeitet von:" & vbCrLf & held
    End If
    On Error GoTo EH
    Set p = ReadJsonFile(ProjectFile(id))
    CheckProjectSchema p
    EnsureStructure p
    Set gProject = p
    gId = id
    mTaskDirty = False
    WriteLog "Projekt geöffnet: " & id
    Exit Sub
EH:
    n = Err.Number: s = Err.Source: d = Err.Description
    ReleaseLock LockPath(id)
    Err.Raise n, s, d
End Sub

Public Sub SaveProject()
    RequireOpenProject
    If Not HasLock(LockPath(gId)) Then UserError "Die Sperre für das Projekt ist verloren gegangen. Bitte Projekt schließen und neu öffnen."
    RefreshKeyCopies gProject("felder")
    gProject("geaendert") = IsoNow()
    gProject("geaendert_von") = UserTag()
    WriteJsonFile ProjectFile(gId), gProject
    mTaskDirty = False
    WriteLog "Projekt gespeichert: " & gId
End Sub

Public Sub CloseProject()
    If gProject Is Nothing Then Exit Sub
    ReleaseLock LockPath(gId)
    WriteLog "Projekt geschlossen: " & gId
    Set gProject = Nothing
    gId = ""
    mTaskDirty = False
End Sub

Public Sub SetTaskDone(ByVal n As Long, ByVal done As Boolean)
    Dim st As Object
    RequireOpenProject
    Set st = gProject("status")
    If JBool(st, "aufgabe_" & n) <> done Then mTaskDirty = True
    st("aufgabe_" & n) = done
    UpdateProjektFlag
End Sub

'---------------------------------------------------------------
' Projekt entfernen: nur Verweis loeschen, Daten bleiben erhalten
'---------------------------------------------------------------
Public Sub DeleteProjectRef(ByVal id As String)
    Dim held As String, n As Long, s As String, d As String
    id = Trim$(id)
    CheckId id
    If gId = id And Not (gProject Is Nothing) Then UserError "Das Projekt ist geöffnet. Bitte zuerst schließen."
    If Not TryAcquireLock(LockPath(id), held) Then UserError "Das Projekt wird gerade bearbeitet von: " & held
    On Error GoTo EH
    DeleteFileSafe RefPath(id)
    WriteLog "Projekt entfernt (nur Verweis): " & id
    ReleaseLock LockPath(id)
    Exit Sub
EH:
    n = Err.Number: s = Err.Source: d = Err.Description
    ReleaseLock LockPath(id)
    Err.Raise n, s, d
End Sub

'---------------------------------------------------------------
' Projekt vollstaendig loeschen: Verweis und gesamten Projektordner
'---------------------------------------------------------------
Public Sub PurgeProject(ByVal id As String)
    Dim held As String, n As Long, s As String, d As String, dir As String
    id = Trim$(id)
    CheckId id
    If gId = id And Not (gProject Is Nothing) Then UserError "Das Projekt ist geöffnet. Bitte zuerst schließen."
    If Not TryAcquireLock(LockPath(id), held) Then UserError "Das Projekt wird gerade bearbeitet von: " & held
    On Error GoTo EH
    dir = ProjectDir(id)
    ' Sicherheitspruefung: Der Ordner muss direkt unter 'projekte' liegen.
    If LCase$(ParentFolder(dir)) <> LCase$(PathProjects()) Then
        UserError "Sicherheitsabbruch: Der Projektordner liegt nicht direkt unter '" & PathProjects() & "'."
    End If
    DeleteFolderTree dir
    DeleteFileSafe RefPath(id)
    WriteLog "Projekt vollständig gelöscht: " & id
    ReleaseLock LockPath(id)
    Exit Sub
EH:
    n = Err.Number: s = Err.Source: d = Err.Description
    ReleaseLock LockPath(id)
    Err.Raise n, s, d
End Sub
