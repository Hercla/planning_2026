' ExportedAt: 2026-01-12 15:37:10 | Workbook: Planning_2026.xlsm
Attribute VB_Name = "PASS2A_Config_Init"
Option Explicit

'===============================================================================
' MODULE: PASS2A_Config_Init
' PURPOSE:
'   PASS 2A – Fondation CONFIG-DRIVEN
'   - Crée / vérifie Feuil_Config
'   - Crée / normalise tblCFG
'   - API robuste CFG_* (SAFE même table vide)
'
' GARANTIES:
'   - Idempotent
'   - Non destructif
'   - Zéro crash ListObject vide
'===============================================================================

Private Const CFG_SHEET_NAME As String = "Feuil_Config"
Private Const CFG_TABLE_NAME As String = "tblCFG"

'===============================================================================
' ENTRYPOINT – PASS 2A
'===============================================================================
Public Sub PASS2A_Init_ConfigFoundation()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim createdSheet As Boolean
    Dim createdTable As Boolean

    Set ws = EnsureConfigSheet(createdSheet)
    Set lo = EnsureConfigTable(ws, createdTable)

    ' Seed defaults (ONLY if missing)
    CFG_AddIfMissing "PLANNING_FIRST_ROW", "5", "Long", "Première ligne du planning"
    CFG_AddIfMissing "DAY_HEADER_ROW", "3", "Long", "Ligne des jours"
    CFG_AddIfMissing "TEAM_MODE", "JOUR", "String", "JOUR / NUIT"
    CFG_AddIfMissing "DEBUG_MODE", "FALSE", "Boolean", "Activer logs debug"

    lo.Range.Columns.AutoFit

    Debug.Print "PASS2A OK | SheetCreated=" & createdSheet & " | TableCreated=" & createdTable
End Sub

'===============================================================================
' ENSURE STRUCTURE
'===============================================================================
Private Function EnsureConfigSheet(ByRef created As Boolean) As Worksheet
    Dim ws As Worksheet
    created = False

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CFG_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        created = True
        Set ws = ThisWorkbook.Worksheets.Add(After:=Sheets(Sheets.Count))
        ws.Name = CFG_SHEET_NAME
    End If

    ws.Visible = xlSheetVisible
    Set EnsureConfigSheet = ws
End Function

Private Function EnsureConfigTable(ByVal ws As Worksheet, ByRef created As Boolean) As ListObject
    Dim lo As ListObject
    created = False

    On Error Resume Next
    Set lo = ws.ListObjects(CFG_TABLE_NAME)
    On Error GoTo 0

    If lo Is Nothing Then
        created = True
        ws.Cells.Clear

        ws.Range("A1:D1").value = Array("Cle", "Valeur", "Type", "Description")
        ' NOTE: ne pas forcer une fausse ligne A2:D2.
        ' La table peut exister sans DataBodyRange -> l'API gère.

        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:D1"), , xlYes)
        lo.Name = CFG_TABLE_NAME

    Else
        ' Normalisation headers (non destructive) + sécurité nb colonnes
        If lo.ListColumns.Count < 4 Then
            ' Si quelqu'un a cassé la table, on l'étend proprement à 4 colonnes
            lo.Resize lo.Range.Resize(lo.Range.Rows.Count, 4)
        End If

        lo.HeaderRowRange.Cells(1, 1).value = "Cle"
        lo.HeaderRowRange.Cells(1, 2).value = "Valeur"
        lo.HeaderRowRange.Cells(1, 3).value = "Type"
        lo.HeaderRowRange.Cells(1, 4).value = "Description"
    End If

    Set EnsureConfigTable = lo
End Function

'===============================================================================
' CORE ACCESS
'===============================================================================
Private Function CFG_Table() As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim dummy As Boolean

    Set ws = EnsureConfigSheet(dummy)
    Set lo = EnsureConfigTable(ws, dummy)
    Set CFG_Table = lo
End Function

'===============================================================================
' API – SAFE (table vide OK)
'===============================================================================
Public Function CFG_Exists(ByVal key As String) As Boolean
    Dim lo As ListObject
    Dim r As Range

    Set lo = CFG_Table()

    If lo.DataBodyRange Is Nothing Then
        CFG_Exists = False
        Exit Function
    End If

    For Each r In lo.DataBodyRange.Rows
        If StrComp(CStr(r.Cells(1, 1).value), key, vbTextCompare) = 0 Then
            CFG_Exists = True
            Exit Function
        End If
    Next r

    CFG_Exists = False
End Function

Public Function CFG_Get(ByVal key As String, Optional defaultValue As Variant) As Variant
    Dim lo As ListObject
    Dim r As Range
    Dim hits As Long

    Set lo = CFG_Table()

    If lo.DataBodyRange Is Nothing Then
        CFG_Get = defaultValue
        Exit Function
    End If

    For Each r In lo.DataBodyRange.Rows
        If StrComp(CStr(r.Cells(1, 1).value), key, vbTextCompare) = 0 Then
            CFG_Get = r.Cells(1, 2).value
            hits = hits + 1
            ' On continue pour détecter doublons (sans casser le comportement)
        End If
    Next r

    If hits = 0 Then
        CFG_Get = defaultValue
    ElseIf hits > 1 Then
        Debug.Print "WARN CFG_Get: key en doublon -> " & key & " (hits=" & hits & ")"
    End If
End Function

Public Sub CFG_Set(ByVal key As String, ByVal value As String, _
                   Optional ByVal typ As String = "String", _
                   Optional ByVal desc As String = "")
    Dim lo As ListObject
    Dim r As Range
    Dim newRow As ListRow
    Dim found As Boolean

    Set lo = CFG_Table()

    If Not lo.DataBodyRange Is Nothing Then
        For Each r In lo.DataBodyRange.Rows
            If StrComp(CStr(r.Cells(1, 1).value), key, vbTextCompare) = 0 Then
                ' Si doublon existe, on met à jour le 1er rencontré, et on log si plusieurs
                If Not found Then
                    r.Cells(1, 2).value = value
                    r.Cells(1, 3).value = typ
                    If Len(desc) > 0 Then r.Cells(1, 4).value = desc
                    found = True
                Else
                    Debug.Print "WARN CFG_Set: doublon ignoré pour key=" & key
                End If
            End If
        Next r

        If found Then Exit Sub
    End If

    ' Add new
    Set newRow = lo.ListRows.Add
    newRow.Range.Cells(1, 1).value = key
    newRow.Range.Cells(1, 2).value = value
    newRow.Range.Cells(1, 3).value = typ
    newRow.Range.Cells(1, 4).value = desc
End Sub

Public Sub CFG_AddIfMissing(ByVal key As String, ByVal value As String, _
                            ByVal typ As String, ByVal desc As String)
    If CFG_Exists(key) Then Exit Sub
    CFG_Set key, value, typ, desc
End Sub

'===============================================================================
' TYPED GETTERS (ROBUSTES)
'===============================================================================
Public Function CFG_GetAsLong(ByVal key As String, Optional defaultValue As Long = 0) As Long
    Dim v As Variant
    v = CFG_Get(key, defaultValue)
    If IsNumeric(v) Then CFG_GetAsLong = CLng(v) Else CFG_GetAsLong = defaultValue
End Function

Public Function CFG_GetAsDouble(ByVal key As String, Optional defaultValue As Double = 0#) As Double
    Dim v As Variant
    v = CFG_Get(key, defaultValue)
    If IsNumeric(v) Then CFG_GetAsDouble = CDbl(v) Else CFG_GetAsDouble = defaultValue
End Function

Public Function CFG_GetAsBoolean(ByVal key As String, Optional defaultValue As Boolean = False) As Boolean
    Dim s As String
    s = UCase$(Trim$(CStr(CFG_Get(key, IIf(defaultValue, "TRUE", "FALSE")))))

    Select Case s
        Case "TRUE", "VRAI", "1", "YES", "OUI"
            CFG_GetAsBoolean = True
        Case "FALSE", "FAUX", "0", "NO", "NON"
            CFG_GetAsBoolean = False
        Case Else
            CFG_GetAsBoolean = defaultValue
    End Select
End Function

Public Function CFG_GetAsString(ByVal key As String, Optional defaultValue As String = "") As String
    CFG_GetAsString = CStr(CFG_Get(key, defaultValue))
End Function


