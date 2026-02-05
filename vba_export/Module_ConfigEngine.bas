Attribute VB_Name = "Module_ConfigEngine"
Option Explicit
' ==========================================================
' Module_ConfigEngine — version adaptée Feuil_Config propre
' ==========================================================

Private Const CFG_SHEET As String = "Feuil_Config"
Private Const COL_KEY As Long = 1
Private Const COL_VAL As Long = 2
Private Const FIRST_ROW As Long = 2

Private mCfg As Object
Private mLoaded As Boolean

' ===================== PUBLIC ======================

Public Sub CFG_Reset()
    mLoaded = False
    Set mCfg = Nothing
End Sub

Public Sub CFG_Load()
    If mLoaded Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CFG_SHEET)

    Set mCfg = CreateObject("Scripting.Dictionary")
    mCfg.CompareMode = vbTextCompare

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, COL_KEY).End(xlUp).row

    Dim r As Long, k As String, v As String

    For r = FIRST_ROW To lastRow
        k = Trim$(CStr(ws.Cells(r, COL_KEY).value))
        If k = "" Then GoTo NextRow

        v = Trim$(CStr(ws.Cells(r, COL_VAL).value))

        If mCfg.Exists(k) Then
            Err.Raise vbObjectError + 1001, , _
                "Doublon clé config : " & k & " (ligne " & r & ")"
        End If

        mCfg.Add k, v

NextRow:
    Next r

    mLoaded = True
End Sub

Public Function CFG_Str(ByVal key As String) As String
    CFG_Load
    If Not mCfg.Exists(key) Then
        Err.Raise vbObjectError + 1002, , "Clé config manquante : " & key
    End If
    CFG_Str = CStr(mCfg(key))
End Function

Public Function CFG_Long(ByVal key As String) As Long
    Dim v As String
    v = CFG_Str(key)

    If Not IsNumeric(v) Then
        Err.Raise vbObjectError + 1003, , _
            "Valeur non numérique pour " & key & " : " & v
    End If

    CFG_Long = CLng(v)
End Function

Public Function CFG_Double(ByVal key As String) As Double
    Dim v As String
    v = CFG_Str(key)

    If Not IsNumeric(v) Then
        Err.Raise vbObjectError + 1004, , _
            "Valeur non numérique pour " & key & " : " & v
    End If

    CFG_Double = CDbl(v)
End Function

Public Function CFG_Bool(ByVal key As String) As Boolean
    Dim v As String
    v = LCase$(CFG_Str(key))

    Select Case v
        Case "1", "true", "oui", "yes"
            CFG_Bool = True
        Case "0", "false", "non", "no"
            CFG_Bool = False
        Case Else
            Err.Raise vbObjectError + 1005, , _
                "Valeur booléenne invalide pour " & key & " : " & v
    End Select
End Function

Public Function CFG_List(ByVal key As String, Optional sep As String = ",") As Variant
    CFG_List = Split(CFG_Str(key), sep)
End Function

Public Sub CFG_Dump()
    CFG_Load
    Dim k As Variant
    For Each k In mCfg.keys
        Debug.Print k & " = " & mCfg(k)
    Next
    Debug.Print "TOTAL:", mCfg.count
End Sub

