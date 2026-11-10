Attribute VB_Name = "modAufgaben"
Option Explicit
'===============================================================
' modAufgaben - Aufgaben 1 bis 5 (Buttons im Blatt 'Projekt')
' Phase 1: nur Platzhalter. Umsetzung folgt in Phase 2 und 3:
'   Phase 2: Aufgabe 1 (Dokument B einlesen, Vollstaendigkeitspruefung)
'   Phase 3: Aufgabe 2 bis 5 (Vorlagen kopieren, Mapping, Befuellen)
' Die Excel-Daten (frueher Aufgabe 'RFC-Zeile') kommen jetzt ueber
' den Button 'Update Projekte' im Blatt 'alle Projekte' (modImport).
'===============================================================
Private Sub NichtUmgesetzt(ByVal nr As Long, ByVal titel As String, ByVal phase As String)
    On Error GoTo EH
    RequireOpenProject
    MsgBox "Aufgabe " & nr & " (" & titel & ") ist in dieser Version noch nicht enthalten." & vbCrLf & _
           "Sie folgt in " & phase & ".", vbInformation, APP_TITLE
    Exit Sub
EH:
    HandleError "Aufgabe_" & nr
End Sub

Public Sub Aufgabe_1()
    NichtUmgesetzt 1, "Dokument B einlesen", "Phase 2"
End Sub

Public Sub Aufgabe_2()
    NichtUmgesetzt 2, "Vorlagen kopieren", "Phase 3"
End Sub

Public Sub Aufgabe_3()
    NichtUmgesetzt 3, "Dokument A füllen", "Phase 3"
End Sub

Public Sub Aufgabe_4()
    NichtUmgesetzt 4, "Score füllen", "Phase 3"
End Sub

Public Sub Aufgabe_5()
    NichtUmgesetzt 5, "SH füllen", "Phase 3"
End Sub
