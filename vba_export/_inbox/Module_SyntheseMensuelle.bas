' ExportedAt: 2026-01-12 15:37:10 | Workbook: Planning_2026.xlsm
Attribute VB_Name = "Module_SyntheseMensuelle"
Option Explicit

'===================================================================================
' MODULE :      Module_SyntheseMensuelle (Version Finale Définitive)
' DESCRIPTION : Calcule les totaux de synthèse en se basant EXCLUSIVEMENT
'               sur les informations de la feuille "Config_Codes".
'===================================================================================

Public Sub CalculerSynthesePlannings()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsConfig As Worksheet, wsPlan As Worksheet
    Dim dictCodes As Object, mois As Variant, userChoice As VbMsgBoxResult
    Dim arrOngletsMois As Variant, arrListeMois As Variant
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    arrOngletsMois = Array("Janv", "Fev", "Mars", "Avril", "Mai", "Juin", "Juil", "Aout", "Sept", "Oct", "Nov", "Dec")
    userChoice = MsgBox("Voulez-vous mettre à jour les soldes pour TOUTE l'année ?" & vbCrLf & vbCrLf & _
                        "?? 'Oui' = Rapport Annuel (12 mois)" & vbCrLf & _
                        "?? 'Non' = Uniquement pour le mois actif", _
                        vbYesNoCancel + vbQuestion, "Périmètre de la Mise à Jour")

    Select Case userChoice
        Case vbYes: arrListeMois = arrOngletsMois
        Case vbNo:  arrListeMois = Array(ActiveSheet.Name)
        Case vbCancel: GoTo CleanUp
    End Select
    
    On Error Resume Next
    Set wsConfig = wb.Sheets("Config_Codes")
    On Error GoTo 0
    If wsConfig Is Nothing Then
        MsgBox "ERREUR : La feuille 'Config_Codes' est introuvable.", vbCritical
        GoTo CleanUp
    End If
    
    Set dictCodes = CreateObject("Scripting.Dictionary")
    dictCodes.CompareMode = vbTextCompare
    Dim lastConfigRow As Long: lastConfigRow = wsConfig.Cells(wsConfig.Rows.Count, "A").End(xlUp).row
    Dim i As Long
    For i = 2 To lastConfigRow
        Dim code As String: code = Trim(wsConfig.Cells(i, "A").value)
        If Len(code) > 0 Then
            ' Stocke le Type_Code (Colonne C) et les Heures Digitales (Colonne R)
            dictCodes(code) = Array(wsConfig.Cells(i, "C").value, wsConfig.Cells(i, "R").value)
        End If
    Next i

    For Each mois In arrListeMois
        On Error Resume Next
        Set wsPlan = wb.Sheets(mois)
        On Error GoTo 0
        If Not wsPlan Is Nothing Then
            CalculerPourUneFeuille wsPlan, dictCodes
        Else
            Debug.Print "Onglet '" & mois & "' introuvable. Ignoré."
        End If
    Next mois

    MsgBox "Synthèse des plannings terminée !", vbInformation

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub


Private Sub CalculerPourUneFeuille(ByVal ws As Worksheet, ByVal dictCodes As Object)
    Dim colHeuresPrestees As Long, colHeuresAPrester As Long, colSoldeMois As Long
    Dim colHeuresRecup As Long, colJoursMaladie As Long, colJoursConge As Long, colJoursAbsence As Long
    
    colHeuresPrestees = TrouverColonne(ws, "Heures prestées")
    colHeuresAPrester = TrouverColonne(ws, "Heures à prester")
    colSoldeMois = TrouverColonne(ws, "Solde du mois")
    colHeuresRecup = TrouverColonne(ws, "Heures à récupérer")
    colJoursMaladie = TrouverColonne(ws, "Jours maladie")
    colJoursConge = TrouverColonne(ws, "Jours congé")
    colJoursAbsence = TrouverColonne(ws, "Jours d’absence")
    
    If colHeuresPrestees = 0 Or colHeuresAPrester = 0 Or colSoldeMois = 0 Then
        MsgBox "ERREUR sur l'onglet '" & ws.Name & "':" & vbCrLf & "Colonnes de synthèse introuvables.", vbCritical
        Exit Sub
    End If

    Dim lastRow As Long, lastCol As Long, startRow As Long
    Dim arrData As Variant, arrResultats As Variant
    Dim r As Long, c As Long
    Dim codeJour As String, typeCode As String, info As Variant
    Dim heuresJour As Double
    Dim hPrestees As Double, hRecup As Double, jMaladie As Long, jConge As Long, jAbsence As Long
    
    startRow = 6
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    If lastRow < startRow Then Exit Sub
    lastCol = ws.Cells(startRow - 2, ws.Columns.Count).End(xlToLeft).Column
    
    arrData = ws.Range(ws.Cells(startRow, 1), ws.Cells(lastRow, lastCol)).value
    ReDim arrResultats(1 To UBound(arrData, 1), 1 To 5)

    For r = 1 To UBound(arrData, 1)
        '--- CORRECTION CLÉ : Réinitialisation des compteurs pour CHAQUE agent ---
        hPrestees = 0: hRecup = 0: jMaladie = 0: jConge = 0: jAbsence = 0
        
        For c = 4 To UBound(arrData, 2)
            codeJour = Trim(CStr(arrData(r, c)))
            
            If dictCodes.Exists(codeJour) Then
                info = dictCodes(codeJour)
                typeCode = info(0)
                heuresJour = val(info(1)) ' Lit la valeur de la colonne R
                
                hPrestees = hPrestees + heuresJour
                
                Select Case typeCode
                    Case "Recup":         hRecup = hRecup + 1
                    Case "Maladie":       jMaladie = jMaladie + 1
                    Case "Congé":         jConge = jConge + 1
                    Case "SansSolde", "Externe", "Famille", "Exceptionnel": jAbsence = jAbsence + 1
                End Select
            End If
        Next c
        
        arrResultats(r, 1) = hPrestees: arrResultats(r, 2) = hRecup: arrResultats(r, 3) = jMaladie
        arrResultats(r, 4) = jConge: arrResultats(r, 5) = jAbsence
    Next r

    ws.Cells(startRow, colHeuresPrestees).Resize(UBound(arrResultats, 1), 1).value = Application.index(arrResultats, 0, 1)
    ws.Cells(startRow, colHeuresRecup).Resize(UBound(arrResultats, 1), 1).value = Application.index(arrResultats, 0, 2)
    ws.Cells(startRow, colJoursMaladie).Resize(UBound(arrResultats, 1), 1).value = Application.index(arrResultats, 0, 3)
    ws.Cells(startRow, colJoursConge).Resize(UBound(arrResultats, 1), 1).value = Application.index(arrResultats, 0, 4)
    ws.Cells(startRow, colJoursAbsence).Resize(UBound(arrResultats, 1), 1).value = Application.index(arrResultats, 0, 5)

    ws.Range(ws.Cells(startRow, colSoldeMois), ws.Cells(lastRow, colSoldeMois)).FormulaR1C1 = _
        "=RC" & colHeuresPrestees & "-RC" & colHeuresAPrester
End Sub


Private Function TrouverColonne(ws As Worksheet, nomHeader As String) As Long
    Dim searchRange As Range, foundCell As Range
    Set searchRange = ws.Range("A1:AZ5")
    Set foundCell = searchRange.Find(What:=nomHeader, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not foundCell Is Nothing Then TrouverColonne = foundCell.Column Else TrouverColonne = 0
End Function

