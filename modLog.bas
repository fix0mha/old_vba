Attribute VB_Name = "modLog"
Option Explicit
'===============================================================
' modLog - einfaches Protokoll (eine Datei je Rechner und Benutzer)
' Fehler beim Protokollieren werden bewusst ignoriert.
'===============================================================
Public Sub WriteLog(ByVal msg As String)
    Dim fn As Long, p As String, folder As String
    On Error Resume Next
    Debug.Print Format$(Now, "hh:nn:ss") & " " & msg
    folder = PathLogs()
    If Len(folder) = 0 Then Exit Sub
    EnsureFolder folder
    p = folder & "\" & Environ$("COMPUTERNAME") & "_" & Environ$("USERNAME") & ".log"
    fn = FreeFile
    Open p For Append As #fn
    Print #fn, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " " & Replace(msg, vbCrLf, " | ")
    Close #fn
End Sub
