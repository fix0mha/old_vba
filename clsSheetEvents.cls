VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSheetEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'===============================================================
' clsSheetEvents - leitet Aenderungen im Blatt 'alle Projekte' weiter
'===============================================================
Private WithEvents mWs As Worksheet

Private mKind As String

Public Sub Attach(ByVal ws As Worksheet, ByVal kind As String)
    Set mWs = ws
    mKind = kind
End Sub

Public Function IsAttached() As Boolean
    IsAttached = Not (mWs Is Nothing)
End Function

Private Sub mWs_Change(ByVal Target As Range)
    If mKind = "P" Then
        Projekt_Changed Target
    Else
        List_Changed Target
    End If
End Sub
