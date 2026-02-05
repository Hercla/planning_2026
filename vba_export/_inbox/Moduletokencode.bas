' ExportedAt: 2026-01-12 15:37:09 | Workbook: Planning_2026.xlsm
Attribute VB_Name = "Moduletokencode"
Sub CalculerSyntheseHeuresPresteesUnCodeParJour()
    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim ongletsMois As Variant
    ongletsMois = Array("Janv", "Fev", "Mars", "Avril", "Mai", "Juin", "Juillet", "Aout", "Sept", "Oct", "Nov", "Dec")
    
    ' Demande à l'utilisateur
    Dim choix As Integer
    choix = MsgBox("Voulez-vous mettre à jour les soldes pour TOUTE l'année (12 mois) ?" & vbCrLf & _
                   "Oui = tous les mois, Non = seulement le mois actif.", _
                   vbYesNoCancel + vbQuestion, "Mise à jour des soldes")
    Dim listeMois As Variant
    If choix = vbYes Then
        listeMois = ongletsMois
    ElseIf choix = vbNo Then
        listeMois = Array(ActiveSheet.Name)
    Else
        MsgBox "Opération annulée."
        Exit Sub
    End If

    ' Charger la table des codes
    Dim wsConfig As Worksheet
    Set wsConfig = wb.Sheets("Config_Codes")
    Dim lastConfigRow As Long
    lastConfigRow = wsConfig.Cells(wsConfig.Rows.Count, 1).End(xlUp).row
    Dim codesTable As Object
    Set codesTable = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 2 To lastConfigRow
        Dim code As String
        code = Trim(wsConfig.Cells(i, 1).value)
        If Len(code) > 0 Then
            codesTable(code) = Array(wsConfig.Cells(i, 3).value, wsConfig.Cells(i, 18).value) ' 3=Type_Code, 18=colonne R Heures digital
        End If
    Next i

    Dim mois As Variant
    For Each mois In listeMois
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = wb.Sheets(mois)
        On Error GoTo 0
        If ws Is Nothing Then
            MsgBox "Onglet '" & mois & "' introuvable. Passage au suivant.", vbExclamation
            GoTo next_mois
        End If

        ' Trouver les colonnes jours (détection automatique)
        Dim firstDayCol As Integer: firstDayCol = 3 ' colonne C
        Dim joursLigne As Integer: joursLigne = 2
        Do While Not IsNumeric(ws.Cells(joursLigne, firstDayCol).value)
            joursLigne = joursLigne + 1
        Loop
        Dim lastDayCol As Integer: lastDayCol = firstDayCol
        Do While IsNumeric(ws.Cells(joursLigne, lastDayCol).value) And ws.Cells(joursLigne, lastDayCol).value <> ""
            lastDayCol = lastDayCol + 1
        Loop
        lastDayCol = lastDayCol - 1

        ' Première ligne du personnel
        Dim firstStaffRow As Integer: firstStaffRow = joursLigne + 1
        Do While ws.Cells(firstStaffRow, 2).value = ""
            firstStaffRow = firstStaffRow + 1
        Loop

        ' Pour chaque personne
        Dim staffRow As Long
        staffRow = firstStaffRow
        Do While ws.Cells(staffRow, 2).value <> ""
            Dim heuresPrestees As Double: heuresPrestees = 0
            Dim joursConge As Integer: joursConge = 0
            Dim joursMaladie As Integer: joursMaladie = 0
            Dim joursAbsence As Integer: joursAbsence = 0

            ' Pour chaque jour
            Dim col As Integer
            For col = firstDayCol To lastDayCol
                Dim codeJour As String
                codeJour = Trim(ws.Cells(staffRow, col).value)
                If Len(codeJour) > 0 And codesTable.Exists(codeJour) Then
                    Dim info As Variant
                    info = codesTable(codeJour)
                    Dim typeCode As String: typeCode = info(0)
                    Dim heuresDigital As Double: heuresDigital = val(info(1))
                    Select Case typeCode
                        Case "Travail"
                            heuresPrestees = heuresPrestees + heuresDigital
                        Case "Congé"
                            joursConge = joursConge + 1
                        Case "Maladie"
                            joursMaladie = joursMaladie + 1
                        Case "SansSolde", "Externe", "Famille", "Exceptionnel"
                            joursAbsence = joursAbsence + 1
                        ' Ajouter autres cas si besoin
                    End Select
                End If
            Next col
            
            ws.Cells(staffRow, 33).value = heuresPrestees   ' AG = Heures prestées
            ws.Cells(staffRow, 40).value = joursConge       ' AN = Jours congé
            ws.Cells(staffRow, 39).value = joursMaladie     ' AM = Jours maladie
            ws.Cells(staffRow, 41).value = joursAbsence     ' AO = Jours d'absence
            staffRow = staffRow + 1
        Loop
next_mois:
    Next mois

    MsgBox "Calcul des heures prestées et jours d'absence terminé !"
End Sub


