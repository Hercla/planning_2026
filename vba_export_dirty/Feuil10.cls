VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo SafeExit
    
    Dim lo As ListObject
    Dim rngTable As Range
    
    ' Table de configuration (nom confirmé : tblCFG)
    On Error Resume Next
    Set lo = Me.ListObjects("tblCFG")
    On Error GoTo 0
    
    If Not lo Is Nothing Then
        Set rngTable = lo.Range
        If Not Intersect(Target, rngTable) Is Nothing Then
            CFG_OnChange_RequestViewApply
            GoTo SafeExit
        End If
    End If
    
SafeExit:
End Sub

