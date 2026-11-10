Attribute VB_Name = "modFileIO"
Option Explicit
'===============================================================
' modFileIO - Dateizugriff (UTF-8 ohne BOM), atomares Schreiben,
'             Ordnerfunktionen und Sperren (Lock-Dateien)
' Sperren: Die Lock-Datei bleibt waehrend der Bearbeitung exklusiv
' geoeffnet. Bricht ein Client ab, gibt Windows die Sperre frei.
'===============================================================
#If VBA7 Then
    Private Declare PtrSafe Function MoveFileExW Lib "kernel32" (ByVal lpExistingFileName As LongPtr, ByVal lpNewFileName As LongPtr, ByVal dwFlags As Long) As Long
#Else
    Private Declare Function MoveFileExW Lib "kernel32" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long, ByVal dwFlags As Long) As Long
#End If

Private mLocks As Object      ' Pfad -> Dateinummer
Private mCounter As Long

Private Function Fso() As Object
    Static o As Object
    If o Is Nothing Then Set o = CreateObject("Scripting.FileSystemObject")
    Set Fso = o
End Function

Public Function UserTag() As String
    UserTag = Environ$("USERNAME") & "@" & Environ$("COMPUTERNAME")
End Function

'---------------------------------------------------------------
' Dateien und Ordner
'---------------------------------------------------------------
Public Function FileExists(ByVal path As String) As Boolean
    FileExists = Fso().FileExists(path)
End Function

Public Function FolderExists(ByVal path As String) As Boolean
    FolderExists = Fso().FolderExists(path)
End Function

Public Sub EnsureFolder(ByVal path As String)
    Dim parent As String
    Do While Len(path) > 3 And Right$(path, 1) = "\"
        path = Left$(path, Len(path) - 1)
    Loop
    If Fso().FolderExists(path) Then Exit Sub
    parent = Fso().GetParentFolderName(path)
    If Len(parent) > 0 Then
        If Not Fso().FolderExists(parent) Then EnsureFolder parent
    End If
    Fso().CreateFolder path
End Sub

Public Sub DeleteFileSafe(ByVal path As String)
    If Fso().FileExists(path) Then Fso().DeleteFile path, True
End Sub

Public Sub DeleteFolderTree(ByVal path As String)
    If Fso().FolderExists(path) Then Fso().DeleteFolder path, True
End Sub

Public Function ParentFolder(ByVal path As String) As String
    ParentFolder = Fso().GetParentFolderName(path)
End Function

Public Function ListFileNames(ByVal folder As String, ByVal ext As String) As Collection
    Dim out As Collection, f As Object
    Set out = New Collection
    If Fso().FolderExists(folder) Then
        For Each f In Fso().GetFolder(folder).Files
            If LCase$(Right$(f.Name, Len(ext))) = LCase$(ext) Then out.Add f.Name
        Next f
    End If
    Set ListFileNames = out
End Function

Public Function ListSubFolderNames(ByVal folder As String) As Collection
    Dim out As Collection, f As Object
    Set out = New Collection
    If Fso().FolderExists(folder) Then
        For Each f In Fso().GetFolder(folder).SubFolders
            out.Add f.Name
        Next f
    End If
    Set ListSubFolderNames = out
End Function

'---------------------------------------------------------------
' Text lesen/schreiben (UTF-8)
'---------------------------------------------------------------
Public Function ReadTextFile(ByVal path As String) As String
    Dim st As Object, v As Variant
    Dim n As Long, d As String
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2          ' Text
    st.Charset = "UTF-8"
    st.Open
    On Error GoTo EH
    st.LoadFromFile path
    v = st.ReadText(-1)
    st.Close
    If IsNull(v) Then v = ""
    ReadTextFile = CStr(v)
    Exit Function
EH:
    n = Err.Number: d = Err.Description
    On Error Resume Next
    st.Close
    On Error GoTo 0
    Err.Raise n, "ReadTextFile", "Datei konnte nicht gelesen werden: " & path & vbCrLf & d
End Function

Public Sub WriteTextUtf8(ByVal path As String, ByVal s As String)
    Dim st As Object, bin As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "UTF-8"
    st.Open
    st.WriteText s
    st.Position = 0
    st.Type = 1          ' binaer
    st.Position = 3      ' BOM (EF BB BF) ueberspringen
    Set bin = CreateObject("ADODB.Stream")
    bin.Type = 1
    bin.Open
    st.CopyTo bin
    bin.SaveToFile path, 2   ' adSaveCreateOverWrite
    bin.Close
    st.Close
End Sub

' Schreibt zuerst in eine temporaere Datei und ersetzt dann das Ziel.
Public Sub WriteTextFileAtomic(ByVal path As String, ByVal s As String)
    Dim tmp As String
    mCounter = mCounter + 1
    tmp = path & ".tmp_" & Environ$("USERNAME") & "_" & Format$(Now, "yymmddhhnnss") & "_" & mCounter
    WriteTextUtf8 tmp, s
    If MoveFileExW(StrPtr(tmp), StrPtr(path), 9) = 0 Then   ' 1 = ersetzen, 8 = Durchschreiben
        On Error Resume Next
        Kill path
        On Error GoTo 0
        Name tmp As path
    End If
End Sub

Public Function ReadJsonFile(ByVal path As String) As Object
    Set ReadJsonFile = JsonParse(ReadTextFile(path))
End Function

Public Sub WriteJsonFile(ByVal path As String, ByVal doc As Object)
    WriteTextFileAtomic path, JsonStringify(doc, True) & vbLf
End Sub

Public Function FileHasBom(ByVal path As String) As Boolean
    Dim fn As Long, b(0 To 2) As Byte
    fn = FreeFile
    Open path For Binary Access Read Shared As #fn
    Get #fn, 1, b
    Close #fn
    FileHasBom = (b(0) = &HEF And b(1) = &HBB And b(2) = &HBF)
End Function

'---------------------------------------------------------------
' Sperren
'---------------------------------------------------------------
Private Function LockDict() As Object
    If mLocks Is Nothing Then Set mLocks = CreateObject("Scripting.Dictionary")
    Set LockDict = mLocks
End Function

Public Function HasLock(ByVal lockPath As String) As Boolean
    HasLock = LockDict().Exists(lockPath)
End Function

' Liefert True, wenn die Sperre erhalten wurde (oder schon bei uns liegt).
' Sonst False und heldBy = Info des aktuellen Inhabers.
Public Function TryAcquireLock(ByVal lockPath As String, ByRef heldBy As String) As Boolean
    Dim fn As Long, b() As Byte, raw() As Byte, info As String, i As Long
    heldBy = ""
    If LockDict().Exists(lockPath) Then
        TryAcquireLock = True
        Exit Function
    End If
    fn = FreeFile
    On Error Resume Next
    Open lockPath For Binary Access Read Write Lock Write As #fn
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        heldBy = ReadLockInfo(lockPath)
        Exit Function
    End If
    On Error GoTo 0
    info = "Benutzer=" & UserTag() & "; seit=" & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ReDim b(0 To 511)
    For i = 0 To 511
        b(i) = 32
    Next i
    raw = StrConv(info, vbFromUnicode)
    For i = 0 To UBound(raw)
        If i > 511 Then Exit For
        b(i) = raw(i)
    Next i
    Put #fn, 1, b
    LockDict().Add lockPath, fn
    TryAcquireLock = True
End Function

Public Sub ReleaseLock(ByVal lockPath As String)
    Dim fn As Long
    If Not LockDict().Exists(lockPath) Then Exit Sub
    fn = LockDict()(lockPath)
    LockDict().Remove lockPath
    On Error Resume Next
    Close #fn
    Kill lockPath
    On Error GoTo 0
End Sub

Public Sub ReleaseAllLocks()
    Dim k As Variant, arr As Variant
    If mLocks Is Nothing Then Exit Sub
    If mLocks.Count = 0 Then Exit Sub
    arr = mLocks.Keys
    For Each k In arr
        ReleaseLock CStr(k)
    Next k
End Sub

' True, wenn ein anderer Client die Sperre haelt (ohne sie zu uebernehmen).
Public Function IsLockedByOther(ByVal lockPath As String, ByRef info As String) As Boolean
    Dim fn As Long
    info = ""
    If LockDict().Exists(lockPath) Then Exit Function
    If Not Fso().FileExists(lockPath) Then Exit Function
    fn = FreeFile
    On Error Resume Next
    Open lockPath For Binary Access Read Write Lock Write As #fn
    If Err.Number = 0 Then
        Close #fn
    Else
        Err.Clear
        On Error GoTo 0
        IsLockedByOther = True
        info = ReadLockInfo(lockPath)
        Exit Function
    End If
    On Error GoTo 0
End Function

Private Function ReadLockInfo(ByVal lockPath As String) As String
    Dim fn As Long, b() As Byte
    On Error GoTo EH
    fn = FreeFile
    Open lockPath For Binary Access Read Shared As #fn
    ReDim b(0 To 511)
    Get #fn, 1, b
    Close #fn
    ReadLockInfo = Trim$(Replace(StrConv(b, vbUnicode), vbNullChar, ""))
    If Len(ReadLockInfo) = 0 Then ReadLockInfo = "(unbekannt)"
    Exit Function
EH:
    On Error Resume Next
    Close #fn
    ReadLockInfo = "(unbekannt)"
End Function
