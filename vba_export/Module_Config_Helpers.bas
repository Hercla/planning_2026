Attribute VB_Name = "Module_Config_Helpers"
Option Explicit

'================================================================================================
'  CONFIG ACCESS (LOCAL, NO DEPENDENCIES)
'  Source: Sheet "Feuil_Config" (A=Key, B=Value)
'================================================================================================

Private Const CFG_SHEET As String = "Feuil_Config"
Private Const CFG_KEY_COL As Long = 1   'A
Private Const CFG_VAL_COL As Long = 2   'B

' Cache (key -> value) to speed up repeated reads
Private mCfgCache As Object
Private mCfgCacheBuilt As Boolean

Private Sub CfgCache_Ensure()
    If mCfgCacheBuilt Then Exit Sub
    Set mCfgCache = CreateObject("Scripting.Dictionary")
    mCfgCache.CompareMode = 1 ' vbTextCompare

    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim k As String, v As String

    On Error GoTo EH
    Set ws = ThisWorkbook.Worksheets(CFG_SHEET)
    lastRow = ws.Cells(ws.Rows.count, CFG_KEY_COL).End(xlUp).row

    For r = 1 To lastRow
        k = Trim$(CStr(ws.Cells(r, CFG_KEY_COL).Value2))
        If Len(k) > 0 Then
            v = Trim$(CStr(ws.Cells(r, CFG_VAL_COL).Value2))
            If Not mCfgCache.Exists(k) Then mCfgCache.Add k, v
        End If
    Next r

    mCfgCacheBuilt = True
    Exit Sub

EH:
    Set mCfgCache = CreateObject("Scripting.Dictionary")
    mCfgCache.CompareMode = 1
    mCfgCacheBuilt = True
End Sub

Public Sub CfgCache_Reset()
    mCfgCacheBuilt = False
    Set mCfgCache = Nothing
End Sub

Public Function CfgGetRaw(ByVal key As String) As String
    CfgCache_Ensure
    key = Trim$(key)
    If Len(key) = 0 Then
        CfgGetRaw = vbNullString
    ElseIf mCfgCache.Exists(key) Then
        CfgGetRaw = CStr(mCfgCache(key))
    Else
        CfgGetRaw = vbNullString
    End If
End Function

Public Function CfgTextOr(ByVal key As String, ByVal defaultVal As String) As String
    Dim raw As String
    raw = CfgGetRaw(key)
    If Len(raw) > 0 Then
        CfgTextOr = raw
    Else
        CfgTextOr = defaultVal
    End If
End Function

Public Function CfgValueOr(ByVal key As String, ByVal defaultVal As Variant) As Variant
    Dim raw As String
    raw = CfgGetRaw(key)
    
    If Len(raw) = 0 Then
        CfgValueOr = defaultVal
        Exit Function
    End If
    
    ' Conversion selon type defaultVal
    Select Case VarType(defaultVal)
        Case vbLong, vbInteger
            If IsNumeric(raw) Then
                CfgValueOr = CLng(raw)
            Else
                CfgValueOr = defaultVal
            End If
        Case vbDouble, vbSingle
            If IsNumeric(raw) Then
                CfgValueOr = CDbl(raw)
            Else
                CfgValueOr = defaultVal
            End If
        Case vbBoolean
            If UCase(raw) = "TRUE" Or raw = "1" Then
                CfgValueOr = True
            ElseIf UCase(raw) = "FALSE" Or raw = "0" Then
                CfgValueOr = False
            Else
                CfgValueOr = defaultVal
            End If
        Case Else
            CfgValueOr = raw
    End Select
End Function
