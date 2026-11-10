Attribute VB_Name = "modSelfTest"
Option Explicit
'===============================================================
' modSelfTest - Selbsttest fuer Umgebung und Grundfunktionen
' Prueft JSON, UTF-8-Dateien, Sperren und Word-Verfuegbarkeit.
' Es werden nur Testdateien im Ordner 'logs' bzw. 'locks' benutzt.
'===============================================================
Private Sub Chk(ByRef sb As String, ByVal cond As Boolean, ByVal label As String)
    If cond Then
        sb = sb & "  OK       " & label & vbCrLf
    Else
        sb = sb & "  FEHLER   " & label & vbCrLf
    End If
End Sub

Public Sub Start_Selbsttest()
    Dim res As String
    On Error GoTo EH
    InitDataRoot
    res = "Selbsttest " & APP_VERSION & " - " & UserTag() & vbCrLf & vbCrLf
    res = res & "JSON" & vbCrLf & TestJson()
    res = res & vbCrLf & "Dateien (UTF-8, atomares Schreiben)" & vbCrLf & TestFiles()
    res = res & vbCrLf & "Sperren" & vbCrLf & TestLock()
    res = res & vbCrLf & "Felder" & vbCrLf & TestFelder()
    res = res & vbCrLf & "Office" & vbCrLf & TestOffice()
    WriteLog "Selbsttest:" & vbCrLf & res
    MsgBox res, vbInformation, APP_TITLE & " - Selbsttest"
    Exit Sub
EH:
    HandleError "Start_Selbsttest"
End Sub

Private Function TestJson() As String
    Dim sb As String, d As Object, l As Collection, d2 As Object, d3 As Object, l2 As Collection
    Dim u As String, s As String, s2 As String
    On Error GoTo EH
    u = "A" & ChrW(228) & ChrW(246) & ChrW(252) & ChrW(223) & ChrW(8364) & " ""q"" \ " & vbTab & "x" & vbLf & "y"
    Set d = NewDict()
    d("text") = u
    d("zahl") = 42
    d("kommazahl") = 3.5
    d("ja") = True
    d("leer") = Null
    Set l = NewList()
    l.Add 1
    l.Add "zwei"
    Set d3 = NewDict()
    d3("x") = "y"
    l.Add d3
    Set d("liste") = l
    s = JsonStringify(d)
    Set d2 = JsonParse(s)
    Chk sb, (d2("text") = u), "Text mit Umlauten, Anführungszeichen, Backslash, Tab, Zeilenumbruch"
    Chk sb, (d2("zahl") = 42), "Ganzzahl"
    Chk sb, (d2("kommazahl") = 3.5), "Dezimalzahl (Punkt, unabhängig von der Ländereinstellung)"
    Chk sb, (d2("ja") = True), "Boolean"
    Chk sb, IsNull(d2("leer")), "null"
    Set l2 = d2("liste")
    Set d3 = l2(3)
    Chk sb, (l2.Count = 3 And d3("x") = "y"), "Array mit verschachteltem Objekt"
    s2 = JsonStringify(d2)
    Chk sb, (s = s2), "Schreiben - Lesen - Schreiben ergibt identischen Text"
    Set d2 = JsonParse("{""a"":[1,2,{""b"":null}],""c"":""\u00e4\n""}")
    Chk sb, (d2("c") = ChrW(228) & vbLf), "Unicode-Escape und \n beim Lesen"
    TestJson = sb
    Exit Function
EH:
    TestJson = sb & "  FEHLER   Ausnahme: " & Err.Description & vbCrLf
End Function

Private Function TestFiles() As String
    Dim sb As String, p As String, t As String
    On Error GoTo EH
    p = PathLogs() & "\_selftest.json"
    t = "Gr" & ChrW(252) & ChrW(223) & "e " & ChrW(8364) & " - Test"
    WriteTextFileAtomic p, t
    Chk sb, FileExists(p), "Datei geschrieben"
    Chk sb, (ReadTextFile(p) = t), "UTF-8-Roundtrip mit Sonderzeichen"
    Chk sb, Not FileHasBom(p), "Datei ohne BOM (für Python u. a. geeignet)"
    WriteTextFileAtomic p, "zweite Version"
    Chk sb, (ReadTextFile(p) = "zweite Version"), "Vorhandene Datei atomar ersetzt"
    DeleteFileSafe p
    Chk sb, Not FileExists(p), "Testdatei gelöscht"
    TestFiles = sb
    Exit Function
EH:
    TestFiles = sb & "  FEHLER   Ausnahme: " & Err.Description & vbCrLf
End Function

Private Function TestLock() As String
    Dim sb As String, lp As String, held As String, fn As Long, second As Boolean
    On Error GoTo EH
    lp = PathLocks() & "\_selftest.lock"
    Chk sb, TryAcquireLock(lp, held), "Sperre erhalten"
    fn = FreeFile
    On Error Resume Next
    Open lp For Binary Access Read Write Lock Write As #fn
    second = (Err.Number = 0)
    If second Then Close #fn
    Err.Clear
    On Error GoTo EH
    Chk sb, Not second, "Zweiter Zugriff auf dieselbe Sperre wird blockiert"
    ReleaseLock lp
    Chk sb, Not FileExists(lp), "Sperre freigegeben und Datei entfernt"
    TestLock = sb
    Exit Function
EH:
    ReleaseLock lp
    TestLock = sb & "  FEHLER   Ausnahme: " & Err.Description & vbCrLf
End Function

Private Function TestFelder() As String
    Dim sb As String, doc As Object
    On Error GoTo EH
    Chk sb, (FieldClass("F001") = "F" And FieldClass("S012") = "S" And FieldClass("B003") = "B" And BFieldId(12) = "B012" And SourceText("S") = SRC_S), "Feld-IDs und Quellen (F/S/B)"
    Set doc = LoadFieldsDoc()
    Chk sb, Not (FindFieldById(doc("felder"), FID_KURZ) Is Nothing Or FindFieldById(doc("felder"), FID_BEARB) Is Nothing), "Feldliste lesbar, Standardfelder F001/F002 vorhanden"
    TestFelder = sb
    Exit Function
EH:
    TestFelder = sb & "  FEHLER   Ausnahme: " & Err.Description & vbCrLf
End Function

Private Function TestOffice() As String
    Dim sb As String, w As Object, v As String
    On Error Resume Next
    Set w = CreateObject("Word.Application")
    If Err.Number <> 0 Or w Is Nothing Then
        Err.Clear
        sb = sb & "  FEHLER   Word ließ sich nicht per Automation starten" & vbCrLf
    Else
        v = CStr(w.Version)
        w.Quit
        sb = sb & "  OK       Word per Automation erreichbar (Version " & v & ")" & vbCrLf
    End If
    sb = sb & "  OK       Excel Version " & Application.Version & vbCrLf
    TestOffice = sb
End Function
