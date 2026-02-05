Attribute VB_Name = "MODULEMODES_CONFIGDRIVEN"
'===============================================================================
' MODULEMODES_CONFIGDRIVEN - VERSION CORRIGEE
' - Séparateur ";"
' - Protection contre les erreurs de type
'===============================================================================
Option Explicit

Public Enum ViewMode
    ViewJour = 1
    ViewNuit = 2
End Enum

'===============================================================================
' PUBLIC SUBS
'===============================================================================
Public Sub Mode_Jour()
    ApplyMode ViewJour
End Sub

Public Sub Mode_Nuit()
    ApplyMode ViewNuit
End Sub

'===============================================================================
' MAIN ROUTINE - ROBUST
'===============================================================================
Private Sub ApplyMode(mode As ViewMode)
    Dim ws As Worksheet
    Dim hideBlocks As String
    Dim parts As Variant
    Dim i As Long
    Dim startRow As Long, endRow As Long
    Dim blockParts() As String
    Dim z As Variant
    Dim hdr As String
    Dim mc As String

    On Error GoTo CleanUp

    ' Freeze updates
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set ws = ActiveSheet

    ' 1) RESET: Show all rows
    ws.Rows.Hidden = False

    ' 2) Get hide blocks from config
    If mode = ViewJour Then
        hideBlocks = Module_Config.CfgText("VIEW_Jour_HideBlocks")
    Else
        hideBlocks = Module_Config.CfgText("VIEW_Nuit_HideBlocks")
    End If

    ' 3) Parse and hide each block - SEPARATEUR ";"
    If Len(hideBlocks) > 0 Then
        parts = Split(hideBlocks, ";")
        For i = LBound(parts) To UBound(parts)
            If Len(Trim$(CStr(parts(i)))) > 0 Then
                blockParts = Split(Trim$(CStr(parts(i))), ":")
                If UBound(blockParts) >= 1 Then
                    If IsNumeric(Trim$(blockParts(0))) And IsNumeric(Trim$(blockParts(1))) Then
                        startRow = CLng(Trim$(blockParts(0)))
                        endRow = CLng(Trim$(blockParts(1)))
                        If startRow > 0 And endRow > 0 And endRow >= startRow Then
                            ws.Rows(startRow & ":" & endRow).Hidden = True
                        End If
                    End If
                ElseIf UBound(blockParts) = 0 Then
                    If IsNumeric(Trim$(blockParts(0))) Then
                        ws.Rows(CLng(Trim$(blockParts(0)))).Hidden = True
                    End If
                End If
            End If
        Next i
    End If

    ' 4) Auto-hide empty name rows
    AutoHideEmpty ws, mode

    ' 5) Headers always visible
    hdr = Module_Config.CfgText("VIEW_HeaderRows_Keep")
    If Len(hdr) > 0 Then
        On Error Resume Next
        ws.Rows(hdr).Hidden = False
        On Error GoTo CleanUp
    End If

    ' 6) Column B
    On Error Resume Next
    ws.Columns("B").Hidden = Module_Config.CfgBool("VIEW_HideColumnB")
    On Error GoTo CleanUp

    ' 7) Menu columns
    mc = Module_Config.CfgText("VIEW_MenuCols")
    If Len(mc) > 0 Then
        On Error Resume Next
        ws.Columns(mc).Hidden = True
        On Error GoTo CleanUp
    End If

    ' 8) Zoom - avec protection
    On Error Resume Next
    z = Module_Config.CfgText("VIEW_Zoom")
    If IsNumeric(z) And val(z) > 0 Then
        ActiveWindow.Zoom = CLng(z)
    Else
        ActiveWindow.Zoom = 70
    End If
    On Error GoTo CleanUp

    ' 9) Scroll position
    If mode = ViewJour Then
        Application.GoTo ws.Range("A1"), Scroll:=True
    Else
        Application.GoTo ws.Range("A30"), Scroll:=True
    End If

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        MsgBox "Erreur Mode " & IIf(mode = ViewJour, "Jour", "Nuit") & ": " & Err.description, vbCritical
    End If
End Sub

'===============================================================================
' AUTO-HIDE EMPTY NAMES
'===============================================================================
Private Sub AutoHideEmpty(ws As Worksheet, mode As ViewMode)
    Dim dataArr As Variant
    Dim startRow As Long, endRow As Long
    Dim colName As String
    Dim i As Long, rowNum As Long

    On Error Resume Next
    
    colName = Module_Config.CfgText("VIEW_NameCol_A")
    If Len(colName) = 0 Then colName = "A"

    If mode = ViewJour Then
        startRow = 6: endRow = 28
    Else
        startRow = 31: endRow = 38
    End If

    dataArr = ws.Range(colName & startRow & ":" & colName & endRow).value
    
    If Not IsArray(dataArr) Then Exit Sub

    For i = 1 To UBound(dataArr, 1)
        If Len(Trim$(CStr(dataArr(i, 1) & ""))) = 0 Then
            rowNum = startRow + i - 1
            ws.Rows(rowNum).Hidden = True
        End If
    Next i
    
    On Error GoTo 0
End Sub

'===============================================================================
' UTILITY
'===============================================================================
Public Sub ToggleMode()
    Static currentMode As ViewMode
    If currentMode = ViewJour Then
        currentMode = ViewNuit
    Else
        currentMode = ViewJour
    End If
    ApplyMode currentMode
End Sub

Public Sub ResetAllRows()
    Application.ScreenUpdating = False
    On Error Resume Next
    ActiveSheet.Rows.Hidden = False
    ActiveSheet.Columns.Hidden = False
    On Error GoTo 0
    Application.ScreenUpdating = True
End Sub

'===============================================================================
' CONFIG HELPERS - Safe versions
'===============================================================================
Private Function CfgText(key As String) As String
    On Error Resume Next
    CfgText = ""
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Feuil_Config")
    Dim rng As Range
    Set rng = ws.Columns(1).Find(What:=key, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rng Is Nothing Then
        CfgText = CStr(rng.offset(0, 1).value & "")
    End If
    On Error GoTo 0
End Function

Private Function CfgBool(key As String) As Boolean
    On Error Resume Next
    CfgBool = False
    Dim val As String
    val = UCase(Module_Config.CfgText(key))
    CfgBool = (val = "VRAI" Or val = "TRUE" Or val = "1" Or val = "OUI" Or val = "YES")
    On Error GoTo 0
End Function

Private Function CfgLong(key As String) As Long
    On Error Resume Next
    CfgLong = 0
    Dim val As String
    val = Module_Config.CfgText(key)
    If IsNumeric(val) Then CfgLong = CLng(val)
    On Error GoTo 0
End Function
