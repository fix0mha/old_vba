Attribute VB_Name = "modConfig"
Option Explicit
'===============================================================
' modConfig - Konstanten, Pfade, Einstellungen, Fehlerbehandlung
'===============================================================
Public Const APP_TITLE As String = "Projektverwaltung"
Public Const APP_VERSION As String = "2.1 (Phase 1)"
Public Const SCHEMA_VERSION As Long = 3
Public Const ERR_USER As Long = vbObjectError + 2000

Public Const SHEET_START As String = "alle Projekte"
Public Const SHEET_PROJEKT As String = "Projekt"
Public Const SHEET_FELDER As String = "Felder anpassen"
Public Const SHEET_ADMIN As String = "Admin Center"
Public Const CELL_DATAROOT As String = "C3"      ' im Blatt 'Admin Center'

' Standardfelder (lokale Felder, immer vorhanden, nicht umbenennbar)
Public Const FID_KURZ As String = "F001"
Public Const FID_BEARB As String = "F002"
Public Const KEY_KURZ As String = "Kurzbeschreibung"
Public Const KEY_BEARB As String = "Bearbeiter"

' Quelle der Werte (Feldklassen F = lokal, S = System, B = BeKo)
Public Const SRC_F As String = "lokales Feld"
Public Const SRC_S As String = "aus System"
Public Const SRC_B As String = "aus BeKo"

Public Const TASK_COUNT As Long = 5

'---------------------------------------------------------------
' Hilfen
'---------------------------------------------------------------
Public Function IsoNow() As String
    IsoNow = Format$(Now, "yyyy-mm-dd\Thh:nn:ss")
End Function

Public Sub UserError(ByVal msg As String)
    Err.Raise ERR_USER, APP_TITLE, msg
End Sub

Public Sub HandleError(ByVal procName As String)
    Dim n As Long, d As String
    n = Err.Number: d = Err.Description
    On Error Resume Next
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    If n = ERR_USER Then
        WriteLog "Hinweis in " & procName & ": " & d
        MsgBox d, vbExclamation, APP_TITLE
    Else
        WriteLog "FEHLER in " & procName & ": [" & n & "] " & d
        MsgBox "Unerwarteter Fehler in " & procName & ":" & vbCrLf & vbCrLf & _
               "[" & n & "] " & d & vbCrLf & vbCrLf & _
               "Details stehen im Log (Ordner 'logs' im Datenpfad).", vbCritical, APP_TITLE
    End If
End Sub

' Fuegt die ersten maxLines Eintraege zeilenweise zusammen.
Public Function JoinLines(ByVal items As Collection, ByVal maxLines As Long) As String
    Dim i As Long, s As String
    For i = 1 To items.Count
        If i > maxLines Then
            s = s & "... und " & (items.Count - maxLines) & " weitere" & vbCrLf
            Exit For
        End If
        s = s & Left$(CStr(items(i)), 100) & vbCrLf
    Next i
    JoinLines = s
End Function

'---------------------------------------------------------------
' Pfade (alle relativ zum Datenpfad auf dem Netzlaufwerk)
'---------------------------------------------------------------
Public Function DataRoot() As String
    Dim s As String
    s = Trim$(CStr(ThisWorkbook.Worksheets(SHEET_ADMIN).Range(CELL_DATAROOT).Value))
    s = Replace(s, "/", "\")
    Do While Len(s) > 0 And Right$(s, 1) = "\"
        s = Left$(s, Len(s) - 1)
    Loop
    If Len(s) = 0 Then UserError "Es ist noch kein Datenpfad eingetragen (Blatt '" & SHEET_ADMIN & "', Zelle " & CELL_DATAROOT & ")."
    DataRoot = s
End Function

Public Function PathConfig() As String
    PathConfig = DataRoot() & "\config"
End Function

Public Function PathRegistry() As String
    PathRegistry = DataRoot() & "\registry"
End Function

Public Function PathLocks() As String
    PathLocks = DataRoot() & "\locks"
End Function

Public Function PathProjects() As String
    PathProjects = DataRoot() & "\projekte"
End Function

Public Function PathTemplates() As String
    PathTemplates = DataRoot() & "\vorlagen"
End Function

Public Function PathLogs() As String
    PathLogs = DataRoot() & "\logs"
End Function

Private Function DefaultDoc(ByVal listKey As String) As Object
    Dim d As Object
    Set d = NewDict()
    d("schema_version") = SCHEMA_VERSION
    Set d(listKey) = NewList()
    Set DefaultDoc = d
End Function

' Legt Ordner und Konfigurationsdateien an, falls sie fehlen.
Public Sub InitDataRoot()
    Dim root As String, p As String
    root = DataRoot()
    On Error GoTo EH
    EnsureFolder root
    EnsureFolder PathConfig()
    EnsureFolder PathRegistry()
    EnsureFolder PathLocks()
    EnsureFolder PathProjects()
    EnsureFolder PathTemplates()
    EnsureFolder PathLogs()
    p = PathConfig() & "\required.json"
    If Not FileExists(p) Then WriteJsonFile p, DefaultDoc("pflichtfelder")
    p = PathConfig() & "\mapping.json"
    If Not FileExists(p) Then WriteJsonFile p, DefaultDoc("mappings")
    On Error GoTo 0
    EnsureFieldsFile
    Exit Sub
EH:
    UserError "Der Datenpfad ist nicht erreichbar oder nicht beschreibbar:" & vbCrLf & root & vbCrLf & vbCrLf & Err.Description
End Sub

'---------------------------------------------------------------
' Einstellungen (config\settings.json), gelten fuer alle Clients
'---------------------------------------------------------------
Private Function SettingsPath() As String
    SettingsPath = PathConfig() & "\settings.json"
End Function

Public Function GetSetting(ByVal settingKey As String, ByVal defaultValue As String) As String
    Dim doc As Object
    GetSetting = defaultValue
    On Error Resume Next
    If FileExists(SettingsPath()) Then
        Set doc = ReadJsonFile(SettingsPath())
        If Err.Number = 0 Then
            If doc.Exists(settingKey) Then
                If Len(JStr(doc, settingKey)) > 0 Then GetSetting = JStr(doc, settingKey)
            End If
        End If
    End If
    Err.Clear
    On Error GoTo 0
End Function

Public Sub SetSetting(ByVal settingKey As String, ByVal newValue As String)
    Dim doc As Object
    Set doc = Nothing
    On Error Resume Next
    If FileExists(SettingsPath()) Then Set doc = ReadJsonFile(SettingsPath())
    Err.Clear
    On Error GoTo 0
    If doc Is Nothing Then
        Set doc = NewDict()
        doc("schema_version") = SCHEMA_VERSION
    End If
    doc(settingKey) = newValue
    WriteJsonFile SettingsPath(), doc
End Sub
