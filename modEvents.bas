Attribute VB_Name = "modEvents"
Option Explicit
'===============================================================
' modEvents - Ereignisbehandlung ohne Code im Tabellenblatt
' HookEvents verbindet die Klasse clsSheetEvents mit dem Blatt
' 'alle Projekte'. Auto_Open/Auto_Close laufen beim Oeffnen und
' Schliessen der Arbeitsmappe (Makros muessen aktiviert sein).
'===============================================================
Private mHook As clsSheetEvents
Private mHookP As clsSheetEvents

Public Sub HookEvents()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_START)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    Set mHook = New clsSheetEvents
    mHook.Attach ws, "L"
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_PROJEKT)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    Set mHookP = New clsSheetEvents
    mHookP.Attach ws, "P"
End Sub

' Stellt die Verbindung wieder her, falls sie verloren ging
' (z. B. nach einem VBA-Reset).
Public Sub EnsureHooked()
    If mHook Is Nothing Or mHookP Is Nothing Then
        HookEvents
    ElseIf Not (mHook.IsAttached And mHookP.IsAttached) Then
        HookEvents
    End If
End Sub

Public Sub Auto_Open()
    On Error Resume Next
    HookEvents
End Sub

Public Sub Auto_Close()
    On Error Resume Next
    Set mHook = Nothing
    Set mHookP = Nothing
    ReleaseAllLocks
End Sub
