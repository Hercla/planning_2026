Attribute VB_Name = "ModuleCheckAFCMonthly"
' ExportedAt: 2026-01-12 15:37:09 | Workbook: Planning_2026.xlsm
Option Explicit

' Vérifie que certains employés ont le nombre requis de codes "DP"
' dans la feuille de planning actuellement ouverte.
Sub CheckDPMonthlyCodes()
    On Error GoTo ErrorHandler

    ' Optimisation de l'affichage pour la vitesse
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim ws As Worksheet
    Dim wsConfig As Worksheet
    Dim shiftType As String
    Dim startRow As Long, lastRow As Long
    Dim startCol As Long, endCol As Long
    Dim row As Long
    Dim employeeName As String, countDP As Long
    Dim configCol As Long
    Dim expectedCounts As Object
    Dim report As String
    Dim rngSearch As Range

    Set ws = ActiveSheet
    
    ' --- Chargement de la configuration ---
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("Configuration_CTR_CheckWeek")
    On Error GoTo ErrorHandler
    
    If wsConfig Is Nothing Then
        MsgBox "La feuille 'Configuration_CTR_CheckWeek' est introuvable.", vbCritical, "Erreur de configuration"
        GoTo CleanUp
    End If
    
    ' --- Détection du type de planning (jour/nuit) ---
    Dim startRowJour As Long, startRowNuit As Long
    startRowJour = wsConfig.Cells(2, 2).value
    startRowNuit = wsConfig.Cells(2, 3).value

    shiftType = ""

    ' Vérification basée sur les lignes masquées
    If ws.Rows(startRowJour).Hidden = False Then
        shiftType = "jour"
    ElseIf ws.Rows(startRowNuit).Hidden = False Then
        shiftType = "nuit"
    End If
    
    ' Fallback sur le nom de l'onglet si les lignes ne suffisent pas
    If shiftType = "" Then
        If InStr(1, ws.Name, "nuit", vbTextCompare) > 0 Then
            shiftType = "nuit"
        ElseIf InStr(1, ws.Name, "jour", vbTextCompare) > 0 Then
            shiftType = "jour"
        End If
    End If

    If shiftType = "" Then
        MsgBox "Impossible de déterminer si le planning est de type Jour ou Nuit." & vbNewLine & _
               "Vérifiez l'affichage des lignes ou le nom de l'onglet.", vbExclamation, "Vérification DP"
        GoTo CleanUp
    End If

    ' --- Configuration des limites de lecture ---
    If shiftType = "jour" Then configCol = 2 Else configCol = 3
    startRow = wsConfig.Cells(2, configCol).value
    lastRow = wsConfig.Cells(3, configCol).value
    startCol = wsConfig.Cells(5, configCol).value
    endCol = wsConfig.Cells(6, configCol).value

    ' --- Chargement du dictionnaire (Employés et Quotas DP) ---
    Set expectedCounts = CreateObject("Scripting.Dictionary")
    
    Dim lastConfigRow As Long
    Dim i As Long
    Dim configEmployeeName As String
    Dim configExpectedCount As Long
    Dim configShiftType As String
    
    lastConfigRow = wsConfig.Cells(wsConfig.Rows.count, "G").End(xlUp).row
    
    For i = 2 To lastConfigRow
        configEmployeeName = LCase(Trim(CStr(wsConfig.Cells(i, "G").value)))
        configExpectedCount = wsConfig.Cells(i, "H").value
        configShiftType = LCase(Trim(CStr(wsConfig.Cells(i, "I").value)))
        
        If configShiftType = shiftType And configEmployeeName <> "" Then
            If Not expectedCounts.Exists(configEmployeeName) Then
                expectedCounts.Add configEmployeeName, configExpectedCount
            End If
        End If
    Next i
    
    If expectedCounts.count = 0 Then
        MsgBox "Aucun employé à vérifier n'a été trouvé dans la config pour l'équipe de " & shiftType & ".", vbInformation, "Vérification DP"
        GoTo CleanUp
    End If

    ' --- Vérification des codes DP sur le planning (Optimisée) ---
    report = ""
    
    For row = startRow To lastRow
        ' Vérifie si la cellule nom n'est pas vide
        If Not IsEmpty(ws.Cells(row, 1).value) Then
            employeeName = LCase(Trim(CStr(ws.Cells(row, 1).value)))
            
            If expectedCounts.Exists(employeeName) Then
                
                ' OPTIMISATION : Utilisation de CountIf au lieu de boucler sur chaque cellule
                Set rngSearch = ws.Range(ws.Cells(row, startCol), ws.Cells(row, endCol))
                countDP = Application.WorksheetFunction.CountIf(rngSearch, "DP")
                
                If countDP <> expectedCounts(employeeName) Then
                    report = report & ws.Cells(row, 1).value & " : " & countDP & _
                             " DP (attendu " & expectedCounts(employeeName) & ")" & vbNewLine
                End If
                
            End If
        End If
    Next row

    ' --- Affichage du rapport final ---
    If report <> "" Then
        MsgBox "Vérification DP - écarts détectés (" & shiftType & ") :" & vbNewLine & vbNewLine & report, vbExclamation, "Rapport DP"
    Else
        MsgBox "Tous les employés ciblés (" & shiftType & ") possèdent le nombre requis de codes DP.", vbInformation, "Vérification DP"
    End If
    
    GoTo CleanUp

ErrorHandler:
    MsgBox "Erreur inattendue : " & Err.description & " (Code: " & Err.Number & ")", vbCritical, "Erreur"

CleanUp:
    Set ws = Nothing
    Set wsConfig = Nothing
    Set expectedCounts = Nothing
    Set rngSearch = Nothing
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

