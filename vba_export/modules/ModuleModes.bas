Attribute VB_Name = "ModuleModes"
' ExportedAt: 2026-01-12 20:22:00 | Workbook: Planning_2026.xlsm
Option Explicit

'===============================================================================
' MODULEMODES - FIXED VERSION
'===============================================================================

Private Enum ViewMode
    ViewJour
    ViewNuit
End Enum

Private Const MENU_COLS As String = "AH:AO"

'===============================================================================
' PUBLIC SUBS
'===============================================================================

Public Sub Mode_Jour_Legacy()
    AdjustView ViewJour
End Sub

Public Sub Mode_Nuit_Legacy()
    AdjustView ViewNuit
End Sub

'===============================================================================
' CORE - FIXED
'===============================================================================

Private Sub AdjustView(mode As ViewMode)
    Dim ws As Worksheet
    
    Set ws = ActiveSheet
    
    ' Performance optimization
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo CleanUp
    
    ' 1. Reset: Show all rows first
    ws.Rows.Hidden = False
    ws.Columns.Hidden = False
    
    ' 2. Hide rows based on mode
    If mode = ViewJour Then
        ' MODE JOUR: Hide night rows
        ws.Rows("5:5").Hidden = True
        ws.Rows("31:39").Hidden = True
        ws.Rows("43:45").Hidden = True
        ws.Rows("46:47").Hidden = True
        ws.Rows("48:58").Hidden = True
        ws.Rows("71:150").Hidden = True
    Else
        ' MODE NUIT: Hide day rows
        ws.Rows("5:28").Hidden = True
        ws.Rows("39:45").Hidden = True
        ws.Rows("48:58").Hidden = True
        ws.Rows("60:62").Hidden = True
        ws.Rows("64:70").Hidden = True
    End If
    
    ' 3. Auto-hide empty name rows
    AutoHideEmptyNames ws, mode
    
    ' 4. Hide columns
    ws.Columns("B").Hidden = True
    ws.Columns(MENU_COLS).Hidden = True
    
    ' 5. Zoom and position
    ActiveWindow.Zoom = 70
    If mode = ViewJour Then
        Application.GoTo ws.Range("A1"), Scroll:=True
    Else
        Application.GoTo ws.Range("A30"), Scroll:=True
    End If

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

'===============================================================================
' AUTO-HIDE EMPTY NAMES - FAST VERSION
'===============================================================================

Private Sub AutoHideEmptyNames(ws As Worksheet, mode As ViewMode)
    Dim dataArr As Variant
    Dim startRow As Long, endRow As Long
    Dim i As Long, rowNum As Long
    
    ' Define range based on mode
    If mode = ViewJour Then
        startRow = 6: endRow = 28
    Else
        startRow = 31: endRow = 38
    End If
    
    ' Read all values at once
    dataArr = ws.Range("A" & startRow & ":A" & endRow).value
    
    ' Hide empty rows
    On Error Resume Next
    For i = 1 To UBound(dataArr, 1)
        If Len(Trim$(CStr(dataArr(i, 1) & ""))) = 0 Then
            rowNum = startRow + i - 1
            ws.Rows(rowNum).Hidden = True
        End If
    Next i
    On Error GoTo 0
End Sub
