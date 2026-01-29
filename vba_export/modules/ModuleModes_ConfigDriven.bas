Attribute VB_Name = "ModuleModes_ConfigDriven"
'===============================================================================
' MODULEMODES_CONFIGDRIVEN - ROBUST + FAST VERSION
'===============================================================================
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
    
    ' Freeze updates
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo CleanUp
    
    Set ws = ActiveSheet
    
    ' 1) RESET: Show all rows
    ws.Rows.Hidden = False
    
    ' 2) Get hide blocks from config
    If mode = ViewJour Then
        hideBlocks = CfgTextOr("VIEW_Jour_HideBlocks", "")
    Else
        hideBlocks = CfgTextOr("VIEW_Nuit_HideBlocks", "")
    End If
    
    ' Debug: Uncomment to verify config is loading
    ' MsgBox "Mode: " & IIf(mode = ViewJour, "JOUR", "NUIT") & vbCrLf & "HideBlocks: " & hideBlocks
    
    ' 3) Parse and hide each block
    If Len(hideBlocks) > 0 Then
        parts = Split(hideBlocks, "|")
        For i = LBound(parts) To UBound(parts)
            If Len(Trim$(parts(i))) > 0 Then
                blockParts = Split(Trim$(parts(i)), ":")
                If UBound(blockParts) >= 1 Then
                    startRow = CLng(Trim$(blockParts(0)))
                    endRow = CLng(Trim$(blockParts(1)))
                    ws.Rows(startRow & ":" & endRow).Hidden = True
                ElseIf UBound(blockParts) = 0 Then
                    ws.Rows(CLng(Trim$(blockParts(0)))).Hidden = True
                End If
            End If
        Next i
    End If
    
    ' 4) Auto-hide empty name rows
    AutoHideEmpty ws, mode
    
    ' 5) Headers always visible
    Dim hdr As String
    hdr = CfgTextOr("VIEW_HeaderRows_Keep", "")
    If Len(hdr) > 0 Then ws.Rows(hdr).Hidden = False
    
    ' 6) Column B
    ws.Columns("B").Hidden = CfgBool("VIEW_HideColumnB")
    
    ' 7) Menu columns
    Dim mc As String
    mc = CfgTextOr("VIEW_MenuCols", "")
    If Len(mc) > 0 Then ws.Columns(mc).Hidden = True
    
    ' 8) Zoom
    Dim z As Long
    z = CfgLong("VIEW_Zoom")
    If z > 0 Then ActiveWindow.Zoom = z Else ActiveWindow.Zoom = 70
    
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
    
    colName = CfgTextOr("VIEW_NameCol_A", "")
    If Len(colName) = 0 Then colName = "A"
    
    If mode = ViewJour Then
        startRow = 6: endRow = 28
    Else
        startRow = 31: endRow = 38
    End If
    
    On Error Resume Next
    dataArr = ws.Range(colName & startRow & ":" & colName & endRow).value
    On Error GoTo 0
    
    If Not IsArray(dataArr) Then Exit Sub
    
    For i = 1 To UBound(dataArr, 1)
        If Len(Trim$(CStr(dataArr(i, 1) & ""))) = 0 Then
            rowNum = startRow + i - 1
            ws.Rows(rowNum).Hidden = True
        End If
    Next i
End Sub
'===============================================================================
' UTILITY
'===============================================================================
Public Sub ToggleMode()
    Static currentMode As ViewMode
    currentMode = IIf(currentMode = ViewJour, ViewNuit, ViewJour)
    ApplyMode currentMode
End Sub
Public Sub ResetAllRows()
    Application.ScreenUpdating = False
    ActiveSheet.Rows.Hidden = False
    ActiveSheet.Columns.Hidden = False
    Application.ScreenUpdating = True
End Sub
