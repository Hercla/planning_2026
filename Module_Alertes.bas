Attribute VB_Name = "Module_Alertes"
'====================================================================
' MODULE ALERTES - Planning_2026_RUNTIME
' Systeme d'alertes centralise: soldes conges, heures, absenteisme,
' staffing minimum, fin de periode, mise en forme conditionnelle
'
' DEPENDANCES:
'   - Module_HeuresTravaillees: HeuresPresteesMois(), HeuresTheoriquesMois(),
'     JoursOuvrablesMois(), DureeEffectiveCode()
'   - Module_HeuresSup: CalculerHeuresSupMois(), GetSoldeHeuresSup()
'   - Module_Planning_Core: BuildFeriesBE(), ChargerFonctionsPersonnel()
'   - Module_SuiviRH: GetQuotas() (pour soldes conges)
'   - Feuille "Personnel": col B=Nom, C=Prenom, D=HeuresStdJour, E=Fonction
'   - Feuilles mensuelles: Janv..Dec
'
' INSTALLATION:
'   1. Ouvrir Planning_2026_RUNTIME.xlsm
'   2. Alt+F11 > Menu Fichier > Importer > Module_Alertes.bas
'   3. Verifier que les modules dependants sont presents
'   4. Optionnel: appeler AfficherAlertesAuDemarrage depuis Workbook_Open
'====================================================================

Option Explicit

' ---- SEUILS CONFIGURABLES ----
Private Const SEUIL_SOLDE_CA_BAS As Long = 3     ' Jours CA restants (alerte si <)
Private Const SEUIL_DETTE_HORAIRE As Double = -10 ' Heures (alerte si delta <)
Private Const SEUIL_CREDIT_HORAIRE As Double = 50 ' Heures (alerte si delta >)
Private Const SEUIL_MALADIE_MOIS As Long = 5     ' Jours maladie par mois
Private Const SEUIL_MALADIE_AN As Long = 30      ' Jours maladie par an
Private Const SEUIL_CA_FIN_ANNEE As Long = 5     ' CA restants en Q4 (alerte si >)
Private Const STAFFING_MIN_INF As Long = 3        ' Infirmiers minimum par jour
Private Const STAFFING_MIN_AS As Long = 2         ' Aides-soignants minimum par jour

' ---- CONSTANTES STRUCTURE ----
Private Const MONTH_SHEETS As String = "Janv,Fev,Mars,Avril,Mai,Juin,Juil,Aout,Sept,Oct,Nov,Dec"
Private Const FIRST_EMP_ROW As Long = 5
Private Const FIRST_DAY_COL As Long = 3
Private Const MAX_DAY_COL As Long = 33

' ---- CODES D'ABSENCE ET MALADIE ----
Private Const ABSENCE_CODES As String = "CA,EL,ANC,C SOC,DP,CTR,RCT,RV,WE,DECES,RHS,JF"
Private Const MALADIE_PREFIXES As String = "MAL-,MUT,MAT-,PAT-"

' ---- CRITICITE ----
Private Const CRIT_HAUTE As String = "HAUTE"
Private Const CRIT_MOYENNE As String = "MOYENNE"
Private Const CRIT_BASSE As String = "BASSE"

' ---- HELPER ----
Private Function SafeCStr(ByVal v As Variant) As String
    If IsError(v) Then
        SafeCStr = ""
    ElseIf IsNull(v) Or IsEmpty(v) Then
        SafeCStr = ""
    Else
        SafeCStr = CStr(v)
    End If
End Function

'====================================================================
' FONCTION PUBLIQUE 1: VerifierAlertesAgent
' Retourne toutes les alertes actives pour un agent
'
' @param nom   String   Nom agent (format "Nom_Prenom")
' @return Collection   Collection de strings (messages d'alerte)
'====================================================================
Public Function VerifierAlertesAgent(ByVal nom As String) As Collection
    Dim alertes As New Collection
    Dim annee As Long
    Dim m As Long
    Dim mSheets() As String
    Dim totalMaladie As Long
    Dim maladieMois As Long
    Dim deltaTotal As Double
    Dim hPrest As Double, hTheo As Double
    Dim heuresStd As Double, pctTemps As Double
    Dim ws As Worksheet
    Dim rowAgent As Long
    Dim soldeCA As Double
    Dim caQuota As Double, caPris As Double

    annee = Year(Now)
    mSheets = Split(MONTH_SHEETS, ",")
    heuresStd = GetHeuresStdJourAgent(nom)
    If heuresStd <= 0 Then heuresStd = 7.6

    ' --- Alerte Solde CA ---
    caQuota = 24  ' Defaut
    caPris = 0
    For m = 0 To 11
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(mSheets(m))
        On Error GoTo 0
        If Not ws Is Nothing Then
            rowAgent = TrouverLigneAgent(ws, nom)
            If rowAgent > 0 Then
                caPris = caPris + CompterCodeDansLigne(ws, rowAgent, "CA")
            End If
        End If
        Set ws = Nothing
    Next m
    soldeCA = caQuota - caPris
    If soldeCA < SEUIL_SOLDE_CA_BAS And soldeCA >= 0 Then
        alertes.Add "[" & CRIT_MOYENNE & "] " & nom & " : Solde CA bas (" & soldeCA & "j restants)"
    ElseIf soldeCA < 0 Then
        alertes.Add "[" & CRIT_HAUTE & "] " & nom & " : Solde CA NEGATIF (" & soldeCA & "j)"
    End If

    ' --- Alerte CA restants en Q4 ---
    If Month(Now) >= 10 Then
        If soldeCA > SEUIL_CA_FIN_ANNEE Then
            alertes.Add "[" & CRIT_MOYENNE & "] " & nom & " : " & soldeCA & "j CA restants en Q4 (risque perte)"
        End If
    End If

    ' --- Alerte Delta Horaire ---
    deltaTotal = 0#
    For m = 0 To 11
        Dim moisNum As Long
        moisNum = m + 1
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(mSheets(m))
        On Error GoTo 0
        If Not ws Is Nothing Then
            rowAgent = TrouverLigneAgent(ws, nom)
            If rowAgent > 0 Then
                hPrest = Module_HeuresTravaillees.HeuresPresteesMois(mSheets(m), rowAgent)
                pctTemps = GetPctTempsAgent(nom, moisNum)
                If pctTemps <= 0 Then pctTemps = 1#
                hTheo = Module_HeuresTravaillees.HeuresTheoriquesMois(annee, moisNum, pctTemps, heuresStd)
                deltaTotal = deltaTotal + (hPrest - hTheo)
            End If
        End If
        Set ws = Nothing
    Next m

    If deltaTotal < SEUIL_DETTE_HORAIRE Then
        alertes.Add "[" & CRIT_HAUTE & "] " & nom & " : Dette horaire " & Round(deltaTotal, 1) & "h (seuil: " & SEUIL_DETTE_HORAIRE & "h)"
    ElseIf deltaTotal > SEUIL_CREDIT_HORAIRE Then
        alertes.Add "[" & CRIT_MOYENNE & "] " & nom & " : Credit horaire " & Round(deltaTotal, 1) & "h (seuil: " & SEUIL_CREDIT_HORAIRE & "h)"
    End If

    ' --- Alerte Absenteisme Maladie ---
    totalMaladie = 0
    For m = 0 To 11
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(mSheets(m))
        On Error GoTo 0
        If Not ws Is Nothing Then
            rowAgent = TrouverLigneAgent(ws, nom)
            If rowAgent > 0 Then
                maladieMois = CompterMaladieDansLigne(ws, rowAgent)
                totalMaladie = totalMaladie + maladieMois
                If maladieMois > SEUIL_MALADIE_MOIS Then
                    alertes.Add "[" & CRIT_MOYENNE & "] " & nom & " : " & maladieMois & "j maladie en " & mSheets(m) & " (seuil: " & SEUIL_MALADIE_MOIS & "j)"
                End If
            End If
        End If
        Set ws = Nothing
    Next m

    If totalMaladie > SEUIL_MALADIE_AN Then
        alertes.Add "[" & CRIT_HAUTE & "] " & nom & " : " & totalMaladie & "j maladie/an (seuil: " & SEUIL_MALADIE_AN & "j)"
    End If

    Set VerifierAlertesAgent = alertes
End Function

'====================================================================
' FONCTION PUBLIQUE 2: VerifierAlertesSoldeConge
' Verifie les soldes CA de tous les agents
'
' @return Collection   Alertes sur les soldes conges
'====================================================================
Public Function VerifierAlertesSoldeConge() As Collection
    Dim alertes As New Collection
    Dim agents As Collection
    Dim agName As Variant
    Dim soldeCA As Double
    Dim caPris As Double
    Dim caQuota As Double
    Dim ws As Worksheet
    Dim mSheets() As String
    Dim m As Long, rowAgent As Long

    Set agents = CollecterAgents()
    mSheets = Split(MONTH_SHEETS, ",")

    For Each agName In agents
        caQuota = 24
        caPris = 0
        For m = 0 To 11
            On Error Resume Next
            Set ws = ThisWorkbook.Sheets(mSheets(m))
            On Error GoTo 0
            If Not ws Is Nothing Then
                rowAgent = TrouverLigneAgent(ws, CStr(agName))
                If rowAgent > 0 Then
                    caPris = caPris + CompterCodeDansLigne(ws, rowAgent, "CA")
                End If
            End If
            Set ws = Nothing
        Next m
        soldeCA = caQuota - caPris
        If soldeCA < SEUIL_SOLDE_CA_BAS And soldeCA >= 0 Then
            alertes.Add "[" & CRIT_MOYENNE & "] " & CStr(agName) & " : Solde CA bas (" & soldeCA & "j)"
        ElseIf soldeCA < 0 Then
            alertes.Add "[" & CRIT_HAUTE & "] " & CStr(agName) & " : Solde CA NEGATIF (" & soldeCA & "j)"
        End If
    Next agName

    Set VerifierAlertesSoldeConge = alertes
End Function

'====================================================================
' FONCTION PUBLIQUE 3: VerifierAlertesHeures
' Verifie les deltas horaires de tous les agents
'
' @return Collection   Alertes sur les heures (dette ou credit excessif)
'====================================================================
Public Function VerifierAlertesHeures() As Collection
    Dim alertes As New Collection
    Dim agents As Collection
    Dim agName As Variant
    Dim deltaTotal As Double
    Dim heuresStd As Double, pctTemps As Double
    Dim hPrest As Double, hTheo As Double
    Dim ws As Worksheet
    Dim mSheets() As String
    Dim m As Long, moisNum As Long, rowAgent As Long
    Dim annee As Long

    annee = Year(Now)
    Set agents = CollecterAgents()
    mSheets = Split(MONTH_SHEETS, ",")

    For Each agName In agents
        heuresStd = GetHeuresStdJourAgent(CStr(agName))
        If heuresStd <= 0 Then heuresStd = 7.6
        deltaTotal = 0#
        For m = 0 To 11
            moisNum = m + 1
            On Error Resume Next
            Set ws = ThisWorkbook.Sheets(mSheets(m))
            On Error GoTo 0
            If Not ws Is Nothing Then
                rowAgent = TrouverLigneAgent(ws, CStr(agName))
                If rowAgent > 0 Then
                    hPrest = Module_HeuresTravaillees.HeuresPresteesMois(mSheets(m), rowAgent)
                    pctTemps = GetPctTempsAgent(CStr(agName), moisNum)
                    If pctTemps <= 0 Then pctTemps = 1#
                    hTheo = Module_HeuresTravaillees.HeuresTheoriquesMois(annee, moisNum, pctTemps, heuresStd)
                    deltaTotal = deltaTotal + (hPrest - hTheo)
                End If
            End If
            Set ws = Nothing
        Next m

        If deltaTotal < SEUIL_DETTE_HORAIRE Then
            alertes.Add "[" & CRIT_HAUTE & "] " & CStr(agName) & " : Dette " & Round(deltaTotal, 1) & "h"
        ElseIf deltaTotal > SEUIL_CREDIT_HORAIRE Then
            alertes.Add "[" & CRIT_MOYENNE & "] " & CStr(agName) & " : Credit " & Round(deltaTotal, 1) & "h"
        End If
    Next agName

    Set VerifierAlertesHeures = alertes
End Function

'====================================================================
' FONCTION PUBLIQUE 4: VerifierAlertesAbsenteisme
' Verifie les jours de maladie par mois et cumul annuel
'
' @return Collection   Alertes absenteisme
'====================================================================
Public Function VerifierAlertesAbsenteisme() As Collection
    Dim alertes As New Collection
    Dim agents As Collection
    Dim agName As Variant
    Dim totalMaladie As Long
    Dim maladieMois As Long
    Dim ws As Worksheet
    Dim mSheets() As String
    Dim m As Long, rowAgent As Long

    Set agents = CollecterAgents()
    mSheets = Split(MONTH_SHEETS, ",")

    For Each agName In agents
        totalMaladie = 0
        For m = 0 To 11
            On Error Resume Next
            Set ws = ThisWorkbook.Sheets(mSheets(m))
            On Error GoTo 0
            If Not ws Is Nothing Then
                rowAgent = TrouverLigneAgent(ws, CStr(agName))
                If rowAgent > 0 Then
                    maladieMois = CompterMaladieDansLigne(ws, rowAgent)
                    totalMaladie = totalMaladie + maladieMois
                    If maladieMois > SEUIL_MALADIE_MOIS Then
                        alertes.Add "[" & CRIT_MOYENNE & "] " & CStr(agName) & " : " & maladieMois & "j maladie en " & mSheets(m)
                    End If
                End If
            End If
            Set ws = Nothing
        Next m

        If totalMaladie > SEUIL_MALADIE_AN Then
            alertes.Add "[" & CRIT_HAUTE & "] " & CStr(agName) & " : " & totalMaladie & "j maladie/an (seuil: " & SEUIL_MALADIE_AN & "j)"
        End If
    Next agName

    Set VerifierAlertesAbsenteisme = alertes
End Function

'====================================================================
' FONCTION PUBLIQUE 5: VerifierAlertesFinPeriode
' Verifie si des agents ont trop de CA restants en Q4
'
' @return Collection   Alertes fin de periode
'====================================================================
Public Function VerifierAlertesFinPeriode() As Collection
    Dim alertes As New Collection
    Dim agents As Collection
    Dim agName As Variant
    Dim caPris As Double, soldeCA As Double
    Dim ws As Worksheet
    Dim mSheets() As String
    Dim m As Long, rowAgent As Long

    ' Ne generer ces alertes qu'en Q4 (Oct-Dec)
    If Month(Now) < 10 Then
        Set VerifierAlertesFinPeriode = alertes
        Exit Function
    End If

    Set agents = CollecterAgents()
    mSheets = Split(MONTH_SHEETS, ",")

    For Each agName In agents
        caPris = 0
        For m = 0 To 11
            On Error Resume Next
            Set ws = ThisWorkbook.Sheets(mSheets(m))
            On Error GoTo 0
            If Not ws Is Nothing Then
                rowAgent = TrouverLigneAgent(ws, CStr(agName))
                If rowAgent > 0 Then
                    caPris = caPris + CompterCodeDansLigne(ws, rowAgent, "CA")
                End If
            End If
            Set ws = Nothing
        Next m
        soldeCA = 24 - caPris
        If soldeCA > SEUIL_CA_FIN_ANNEE Then
            alertes.Add "[" & CRIT_MOYENNE & "] " & CStr(agName) & " : " & soldeCA & "j CA restants (risque perte fin d'annee)"
        End If
    Next agName

    Set VerifierAlertesFinPeriode = alertes
End Function

'====================================================================
' FONCTION PUBLIQUE 6: VerifierStaffingMinimum
' Verifie si le nombre minimum d'agents par fonction est respecte
' pour chaque jour d'un mois donne
'
' @param nomMois   String   Nom de l'onglet mois
' @param jour      Long     Numero du jour (1-31), 0 = tous les jours
' @return Collection   Alertes staffing
'====================================================================
Public Function VerifierStaffingMinimum(ByVal nomMois As String, _
                                         Optional ByVal jour As Long = 0) As Collection
    Dim alertes As New Collection
    Dim ws As Worksheet
    Dim fonctions As Object
    Dim r As Long, col As Long, lastR As Long
    Dim numJours As Long
    Dim agName As String, cellVal As String
    Dim jourDebut As Long, jourFin As Long
    Dim j As Long
    Dim countINF As Long, countAS As Long
    Dim fonctionAgent As String

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nomMois)
    On Error GoTo 0
    If ws Is Nothing Then
        Set VerifierStaffingMinimum = alertes
        Exit Function
    End If

    ' Charger les fonctions du personnel
    Set fonctions = Module_Planning_Core.ChargerFonctionsPersonnel()
    numJours = CompterJoursMois(ws)
    lastR = GetLastEmployeeRow(ws)

    If jour > 0 And jour <= numJours Then
        jourDebut = jour
        jourFin = jour
    Else
        jourDebut = 1
        jourFin = numJours
    End If

    For j = jourDebut To jourFin
        col = FIRST_DAY_COL + j - 1
        countINF = 0
        countAS = 0

        For r = FIRST_EMP_ROW To lastR
            agName = Trim(SafeCStr(ws.Cells(r, 1).Value))
            If Len(agName) = 0 Then GoTo NextRowStaff
            If InStr(agName, "Remplacement") > 0 Then GoTo NextRowStaff
            If agName = "Us Nuit" Then GoTo NextRowStaff

            cellVal = Trim(SafeCStr(ws.Cells(r, col).Value))
            If Len(cellVal) = 0 Or cellVal = "0" Then GoTo NextRowStaff

            ' Verifier si c'est une prestation effective (pas absence ni maladie)
            If EstPresent(cellVal) Then
                ' Determiner la fonction
                fonctionAgent = ""
                If fonctions.Exists(agName) Then
                    fonctionAgent = UCase(Trim(CStr(fonctions(agName))))
                End If

                Select Case fonctionAgent
                    Case "INF"
                        countINF = countINF + 1
                    Case "AS"
                        countAS = countAS + 1
                    Case "CEFA"
                        ' Les CEFA ne comptent pas dans les minima
                End Select
            End If
NextRowStaff:
        Next r

        ' Verifier les seuils
        If countINF < STAFFING_MIN_INF Then
            alertes.Add "[" & CRIT_HAUTE & "] " & nomMois & " jour " & j & " : " & _
                        countINF & " INF presents (minimum: " & STAFFING_MIN_INF & ")"
        End If
        If countAS < STAFFING_MIN_AS Then
            alertes.Add "[" & CRIT_HAUTE & "] " & nomMois & " jour " & j & " : " & _
                        countAS & " AS presents (minimum: " & STAFFING_MIN_AS & ")"
        End If
    Next j

    Set VerifierStaffingMinimum = alertes
End Function

'====================================================================
' SUB PUBLIQUE 7: AfficherAlertesAuDemarrage
' Verification globale au lancement, affiche un resume si alertes
'====================================================================
Public Sub AfficherAlertesAuDemarrage()
    Dim alertesCA As Collection
    Dim alertesHeures As Collection
    Dim alertesMaladie As Collection
    Dim alertesFinPeriode As Collection
    Dim totalAlertes As Long
    Dim msg As String
    Dim al As Variant

    On Error Resume Next

    Set alertesCA = VerifierAlertesSoldeConge()
    Set alertesHeures = VerifierAlertesHeures()
    Set alertesMaladie = VerifierAlertesAbsenteisme()
    Set alertesFinPeriode = VerifierAlertesFinPeriode()

    On Error GoTo 0

    totalAlertes = 0
    If Not alertesCA Is Nothing Then totalAlertes = totalAlertes + alertesCA.Count
    If Not alertesHeures Is Nothing Then totalAlertes = totalAlertes + alertesHeures.Count
    If Not alertesMaladie Is Nothing Then totalAlertes = totalAlertes + alertesMaladie.Count
    If Not alertesFinPeriode Is Nothing Then totalAlertes = totalAlertes + alertesFinPeriode.Count

    If totalAlertes = 0 Then Exit Sub

    msg = totalAlertes & " alerte(s) detectee(s) :" & vbCrLf & vbCrLf

    ' Alertes HAUTE en premier (max 10 lignes dans la MsgBox)
    Dim countShown As Long
    countShown = 0

    If Not alertesCA Is Nothing Then
        For Each al In alertesCA
            If InStr(CStr(al), CRIT_HAUTE) > 0 And countShown < 10 Then
                msg = msg & CStr(al) & vbCrLf
                countShown = countShown + 1
            End If
        Next al
    End If

    If Not alertesHeures Is Nothing Then
        For Each al In alertesHeures
            If InStr(CStr(al), CRIT_HAUTE) > 0 And countShown < 10 Then
                msg = msg & CStr(al) & vbCrLf
                countShown = countShown + 1
            End If
        Next al
    End If

    If Not alertesMaladie Is Nothing Then
        For Each al In alertesMaladie
            If InStr(CStr(al), CRIT_HAUTE) > 0 And countShown < 10 Then
                msg = msg & CStr(al) & vbCrLf
                countShown = countShown + 1
            End If
        Next al
    End If

    If countShown < totalAlertes Then
        msg = msg & vbCrLf & "... et " & (totalAlertes - countShown) & " alerte(s) supplementaire(s)."
        msg = msg & vbCrLf & "Voir onglet 'Alertes' pour le rapport complet."
    End If

    MsgBox msg, vbExclamation, "Alertes Planning " & Year(Now)
End Sub

'====================================================================
' SUB PUBLIQUE 8: GenererRapportAlertes
' Cree un onglet "Alertes" avec toutes les alertes actives
' Colonnes: Agent, Type, Message, Criticite, Date
'====================================================================
Public Sub GenererRapportAlertes()
    Dim wsAlertes As Worksheet
    Dim agents As Collection
    Dim agName As Variant
    Dim alertesAgent As Collection
    Dim al As Variant
    Dim rowIdx As Long
    Dim parts() As String
    Dim criticite As String, typeAlerte As String, message As String
    Dim alerteStr As String

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Creer/recreer la feuille
    DeleteSheetSafe "Alertes"
    Set wsAlertes = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsAlertes.Name = "Alertes"

    ' En-tetes
    wsAlertes.Cells(1, 1).Value = "Agent"
    wsAlertes.Cells(1, 2).Value = "Type"
    wsAlertes.Cells(1, 3).Value = "Message"
    wsAlertes.Cells(1, 4).Value = "Criticite"
    wsAlertes.Cells(1, 5).Value = "Date"

    With wsAlertes.Range("A1:E1")
        .Interior.Color = RGB(192, 0, 0)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
    End With

    rowIdx = 2
    Set agents = CollecterAgents()

    For Each agName In agents
        Set alertesAgent = VerifierAlertesAgent(CStr(agName))
        If alertesAgent.Count > 0 Then
            For Each al In alertesAgent
                alerteStr = CStr(al)

                ' Extraire la criticite [HAUTE] / [MOYENNE] / [BASSE]
                criticite = ""
                If InStr(alerteStr, "[" & CRIT_HAUTE & "]") > 0 Then
                    criticite = CRIT_HAUTE
                ElseIf InStr(alerteStr, "[" & CRIT_MOYENNE & "]") > 0 Then
                    criticite = CRIT_MOYENNE
                ElseIf InStr(alerteStr, "[" & CRIT_BASSE & "]") > 0 Then
                    criticite = CRIT_BASSE
                End If

                ' Determiner le type d'alerte
                typeAlerte = ClassifierAlerte(alerteStr)

                ' Nettoyer le message (retirer le tag criticite)
                message = alerteStr
                If Len(criticite) > 0 Then
                    message = Trim(Mid(message, InStr(message, "]") + 1))
                End If

                wsAlertes.Cells(rowIdx, 1).Value = CStr(agName)
                wsAlertes.Cells(rowIdx, 2).Value = typeAlerte
                wsAlertes.Cells(rowIdx, 3).Value = message
                wsAlertes.Cells(rowIdx, 4).Value = criticite
                wsAlertes.Cells(rowIdx, 5).Value = Format(Now, "dd/mm/yyyy hh:nn")

                ' Coloration par criticite
                If criticite = CRIT_HAUTE Then
                    wsAlertes.Cells(rowIdx, 4).Font.Color = RGB(204, 0, 0)
                    wsAlertes.Cells(rowIdx, 4).Font.Bold = True
                    wsAlertes.Range(wsAlertes.Cells(rowIdx, 1), wsAlertes.Cells(rowIdx, 5)).Interior.Color = RGB(255, 230, 230)
                ElseIf criticite = CRIT_MOYENNE Then
                    wsAlertes.Cells(rowIdx, 4).Font.Color = RGB(204, 102, 0)
                    wsAlertes.Cells(rowIdx, 4).Font.Bold = True
                    wsAlertes.Range(wsAlertes.Cells(rowIdx, 1), wsAlertes.Cells(rowIdx, 5)).Interior.Color = RGB(255, 243, 230)
                End If

                rowIdx = rowIdx + 1
            Next al
        End If
    Next agName

    ' Ajouter les alertes de staffing pour le mois courant
    Dim mSheets() As String
    Dim moisActuel As Long
    Dim alertesStaff As Collection

    mSheets = Split(MONTH_SHEETS, ",")
    moisActuel = Month(Now)
    If moisActuel >= 1 And moisActuel <= 12 Then
        Set alertesStaff = VerifierStaffingMinimum(mSheets(moisActuel - 1))
        If alertesStaff.Count > 0 Then
            For Each al In alertesStaff
                alerteStr = CStr(al)
                wsAlertes.Cells(rowIdx, 1).Value = "STAFFING"
                wsAlertes.Cells(rowIdx, 2).Value = "Staffing"
                wsAlertes.Cells(rowIdx, 3).Value = Trim(Mid(alerteStr, InStr(alerteStr, "]") + 1))
                wsAlertes.Cells(rowIdx, 4).Value = CRIT_HAUTE
                wsAlertes.Cells(rowIdx, 5).Value = Format(Now, "dd/mm/yyyy hh:nn")
                wsAlertes.Cells(rowIdx, 4).Font.Color = RGB(204, 0, 0)
                wsAlertes.Cells(rowIdx, 4).Font.Bold = True
                wsAlertes.Range(wsAlertes.Cells(rowIdx, 1), wsAlertes.Cells(rowIdx, 5)).Interior.Color = RGB(255, 230, 230)
                rowIdx = rowIdx + 1
            Next al
        End If
    End If

    ' Formatage
    wsAlertes.Columns("A").ColumnWidth = 28
    wsAlertes.Columns("B").ColumnWidth = 14
    wsAlertes.Columns("C").ColumnWidth = 55
    wsAlertes.Columns("D").ColumnWidth = 12
    wsAlertes.Columns("E").ColumnWidth = 18

    If rowIdx > 2 Then
        Dim rng As Range
        Set rng = wsAlertes.Range(wsAlertes.Cells(1, 1), wsAlertes.Cells(rowIdx - 1, 5))
        rng.Borders.LineStyle = xlContinuous
        rng.Borders.Weight = xlThin
        rng.Font.Name = "Calibri"
    End If

    wsAlertes.Range("A2").Select
    ActiveWindow.FreezePanes = True

    wsAlertes.Activate

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox (rowIdx - 2) & " alerte(s) generee(s) dans l'onglet 'Alertes'.", _
           vbInformation, "Rapport Alertes"
End Sub

'====================================================================
' SUB PUBLIQUE 9: AppliquerMiseEnFormeConditionnelle
' Applique une coloration conditionnelle sur une feuille planning
'
' @param nomMois   String   Nom de l'onglet mois
'====================================================================
Public Sub AppliquerMiseEnFormeConditionnelle(ByVal nomMois As String)
    Dim ws As Worksheet
    Dim r As Long, col As Long, lastR As Long
    Dim numJours As Long
    Dim cellVal As String
    Dim agName As String

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nomMois)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    numJours = CompterJoursMois(ws)
    lastR = GetLastEmployeeRow(ws)

    For r = FIRST_EMP_ROW To lastR
        agName = Trim(SafeCStr(ws.Cells(r, 1).Value))
        If Len(agName) = 0 Then GoTo NextRowFormat
        If InStr(agName, "Remplacement") > 0 Then GoTo NextRowFormat

        For col = FIRST_DAY_COL To FIRST_DAY_COL + numJours - 1
            cellVal = UCase(Trim(SafeCStr(ws.Cells(r, col).Value)))
            If Len(cellVal) = 0 Then GoTo NextColFormat

            ' Maladie -> fond rose clair
            If EstCodeMaladie(cellVal) Then
                ws.Cells(r, col).Interior.Color = RGB(255, 199, 206)
            End If

            ' Conge annuel -> fond bleu clair
            If cellVal = "CA" Then
                ws.Cells(r, col).Interior.Color = RGB(189, 215, 238)
            End If

            ' Absence justifiee (EL, ANC, C SOC, DP) -> fond vert clair
            If cellVal = "EL" Or cellVal = "ANC" Or cellVal = "C SOC" Or cellVal = "DP" Then
                ws.Cells(r, col).Interior.Color = RGB(198, 224, 180)
            End If

            ' Jour ferie -> fond jaune
            If Left(cellVal, 2) = "F-" Or Left(cellVal, 2) = "R-" Or cellVal = "JF" Then
                ws.Cells(r, col).Interior.Color = RGB(255, 255, 153)
            End If
NextColFormat:
        Next col
NextRowFormat:
    Next r

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

'====================================================================
' FONCTIONS PRIVEES
'====================================================================

'--------------------------------------------------------------------
' CollecterAgents: collecte tous les agents uniques depuis les 12 mois
'--------------------------------------------------------------------
Private Function CollecterAgents() As Collection
    Dim agents As New Collection
    Dim agentDict As Object
    Dim mSheets() As String
    Dim ws As Worksheet
    Dim r As Long, m As Long
    Dim agName As String

    Set agentDict = CreateObject("Scripting.Dictionary")
    mSheets = Split(MONTH_SHEETS, ",")

    For m = 0 To 11
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(mSheets(m))
        On Error GoTo 0
        If ws Is Nothing Then GoTo NextMonthCollect

        For r = FIRST_EMP_ROW To GetLastEmployeeRow(ws)
            If IsError(ws.Cells(r, 1).Value) Then GoTo NextRowCollect
            agName = Trim(SafeCStr(ws.Cells(r, 1).Value))
            If Len(agName) = 0 Then GoTo NextRowCollect
            If InStr(agName, "Remplacement") > 0 Then GoTo NextRowCollect
            If agName = "Us Nuit" Then GoTo NextRowCollect

            If Not agentDict.Exists(agName) Then
                agentDict.Add agName, True
                agents.Add agName
            End If
NextRowCollect:
        Next r
NextMonthCollect:
        Set ws = Nothing
    Next m

    Set CollecterAgents = agents
End Function

'--------------------------------------------------------------------
' CompterCodeDansLigne: compte les occurrences d'un code dans une ligne
'--------------------------------------------------------------------
Private Function CompterCodeDansLigne(ByVal ws As Worksheet, _
                                       ByVal rowIndex As Long, _
                                       ByVal code As String) As Long
    Dim col As Long, numJours As Long, compteur As Long
    Dim cellVal As String

    numJours = CompterJoursMois(ws)
    compteur = 0
    For col = FIRST_DAY_COL To FIRST_DAY_COL + numJours - 1
        cellVal = UCase(Trim(SafeCStr(ws.Cells(rowIndex, col).Value)))
        If cellVal = UCase(Trim(code)) Then
            compteur = compteur + 1
        End If
    Next col
    CompterCodeDansLigne = compteur
End Function

'--------------------------------------------------------------------
' CompterMaladieDansLigne: compte les jours de maladie dans une ligne
'--------------------------------------------------------------------
Private Function CompterMaladieDansLigne(ByVal ws As Worksheet, _
                                          ByVal rowIndex As Long) As Long
    Dim col As Long, numJours As Long, compteur As Long
    Dim cellVal As String

    numJours = CompterJoursMois(ws)
    compteur = 0
    For col = FIRST_DAY_COL To FIRST_DAY_COL + numJours - 1
        cellVal = UCase(Trim(SafeCStr(ws.Cells(rowIndex, col).Value)))
        If EstCodeMaladie(cellVal) Then
            compteur = compteur + 1
        End If
    Next col
    CompterMaladieDansLigne = compteur
End Function

'--------------------------------------------------------------------
' EstCodeMaladie: verifie si le code est un code maladie
'--------------------------------------------------------------------
Private Function EstCodeMaladie(ByVal code As String) As Boolean
    Dim c As String
    Dim prefixes() As String
    Dim i As Long

    c = Trim(UCase(code))
    prefixes = Split(MALADIE_PREFIXES, ",")
    For i = 0 To UBound(prefixes)
        If Left(c, Len(Trim(prefixes(i)))) = Trim(prefixes(i)) Then
            EstCodeMaladie = True
            Exit Function
        End If
    Next i
    EstCodeMaladie = False
End Function

'--------------------------------------------------------------------
' EstPresent: verifie si un code represente une presence effective
' (ni absence, ni maladie, ni ferie, ni WE)
'--------------------------------------------------------------------
Private Function EstPresent(ByVal code As String) As Boolean
    Dim c As String
    Dim absCodes() As String
    Dim i As Long

    c = Trim(UCase(code))
    EstPresent = False

    If Len(c) = 0 Or c = "0" Then Exit Function

    ' Verifier absences
    absCodes = Split(ABSENCE_CODES, ",")
    For i = 0 To UBound(absCodes)
        If c = Trim(absCodes(i)) Then Exit Function
    Next i

    ' Verifier maladie
    If EstCodeMaladie(c) Then Exit Function

    ' Verifier ferie
    If Len(c) >= 2 Then
        If (Left(c, 1) = "F" Or Left(c, 1) = "R") And Mid(c, 2, 1) = "-" Then Exit Function
    End If

    ' Si on arrive ici, c'est une presence effective
    EstPresent = True
End Function

'--------------------------------------------------------------------
' ClassifierAlerte: determine le type d'alerte a partir du message
'--------------------------------------------------------------------
Private Function ClassifierAlerte(ByVal alerteMsg As String) As String
    Dim msg As String
    msg = UCase(alerteMsg)

    If InStr(msg, "SOLDE CA") > 0 Or InStr(msg, "CA RESTANTS") > 0 Then
        ClassifierAlerte = "Conges"
    ElseIf InStr(msg, "DETTE") > 0 Or InStr(msg, "CREDIT") > 0 Then
        ClassifierAlerte = "Heures"
    ElseIf InStr(msg, "MALADIE") > 0 Then
        ClassifierAlerte = "Absenteisme"
    ElseIf InStr(msg, "STAFFING") > 0 Or InStr(msg, "PRESENTS") > 0 Then
        ClassifierAlerte = "Staffing"
    ElseIf InStr(msg, "Q4") > 0 Or InStr(msg, "FIN D'ANNEE") > 0 Then
        ClassifierAlerte = "Fin Periode"
    Else
        ClassifierAlerte = "Autre"
    End If
End Function

'--------------------------------------------------------------------
' TrouverLigneAgent: trouve la ligne d'un agent dans une feuille mois
'--------------------------------------------------------------------
Private Function TrouverLigneAgent(ByVal ws As Worksheet, ByVal nom As String) As Long
    Dim r As Long, lastR As Long
    Dim agName As String
    TrouverLigneAgent = 0
    lastR = GetLastEmployeeRow(ws)
    For r = FIRST_EMP_ROW To lastR
        agName = Trim(SafeCStr(ws.Cells(r, 1).Value))
        If UCase(agName) = UCase(Trim(nom)) Then
            TrouverLigneAgent = r
            Exit Function
        End If
    Next r
End Function

'--------------------------------------------------------------------
' GetHeuresStdJourAgent: recupere heuresStdJour depuis feuille Personnel
'--------------------------------------------------------------------
Private Function GetHeuresStdJourAgent(ByVal agentName As String) As Double
    Dim wsPers As Worksheet, r As Long, lastRow As Long
    Dim nom As String, prenom As String, clef As String
    On Error Resume Next
    Set wsPers = ThisWorkbook.Sheets("Personnel")
    On Error GoTo 0
    If wsPers Is Nothing Then GetHeuresStdJourAgent = 7.6: Exit Function
    lastRow = wsPers.Cells(wsPers.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        nom = Trim(SafeCStr(wsPers.Cells(r, 2).Value))
        prenom = Trim(SafeCStr(wsPers.Cells(r, 3).Value))
        clef = nom & "_" & prenom
        If UCase(clef) = UCase(Trim(agentName)) Then
            If IsNumeric(wsPers.Cells(r, 4).Value) And wsPers.Cells(r, 4).Value > 0 Then
                GetHeuresStdJourAgent = CDbl(wsPers.Cells(r, 4).Value)
            Else
                GetHeuresStdJourAgent = 7.6
            End If
            Exit Function
        End If
    Next r
    GetHeuresStdJourAgent = 7.6
End Function

'--------------------------------------------------------------------
' GetPctTempsAgent: recupere le % temps mensuel depuis Personnel
'--------------------------------------------------------------------
Private Function GetPctTempsAgent(ByVal agentName As String, ByVal moisNum As Long) As Double
    Dim wsPers As Worksheet, r As Long, lastRow As Long
    Dim nom As String, prenom As String, clef As String
    Dim colPct As Long, val As Variant
    On Error Resume Next
    Set wsPers = ThisWorkbook.Sheets("Personnel")
    On Error GoTo 0
    If wsPers Is Nothing Then GetPctTempsAgent = 1#: Exit Function
    colPct = 29 + (moisNum - 1) * 2
    lastRow = wsPers.Cells(wsPers.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        nom = Trim(SafeCStr(wsPers.Cells(r, 2).Value))
        prenom = Trim(SafeCStr(wsPers.Cells(r, 3).Value))
        clef = nom & "_" & prenom
        If UCase(clef) = UCase(Trim(agentName)) Then
            val = wsPers.Cells(r, colPct).Value
            If IsNumeric(val) And val > 0 Then
                GetPctTempsAgent = CDbl(val)
                If GetPctTempsAgent > 1.5 Then GetPctTempsAgent = GetPctTempsAgent / 100#
            Else
                GetPctTempsAgent = 1#
            End If
            Exit Function
        End If
    Next r
    GetPctTempsAgent = 1#
End Function

'--------------------------------------------------------------------
' CompterJoursMois: nombre de jours dans le mois via en-tetes
'--------------------------------------------------------------------
Private Function CompterJoursMois(ByVal ws As Worksheet) As Long
    Dim col As Long, compteur As Long
    compteur = 0
    For col = FIRST_DAY_COL To MAX_DAY_COL
        If IsNumeric(ws.Cells(4, col).Value) And _
           Len(Trim(SafeCStr(ws.Cells(4, col).Value))) > 0 Then
            compteur = compteur + 1
        Else
            If compteur > 0 Then Exit For
        End If
    Next col
    If compteur < 28 Then compteur = 31
    If compteur > 31 Then compteur = 31
    CompterJoursMois = compteur
End Function

'--------------------------------------------------------------------
' GetLastEmployeeRow: derniere ligne agent
'--------------------------------------------------------------------
Private Function GetLastEmployeeRow(ByVal ws As Worksheet) As Long
    Dim r As Long, emptyCount As Long, lastFound As Long
    lastFound = FIRST_EMP_ROW - 1
    emptyCount = 0
    For r = FIRST_EMP_ROW To 200
        If Len(Trim(SafeCStr(ws.Cells(r, 1).Value))) > 0 Then
            lastFound = r
            emptyCount = 0
        Else
            emptyCount = emptyCount + 1
            If emptyCount >= 5 Then Exit For
        End If
    Next r
    GetLastEmployeeRow = lastFound
End Function

'--------------------------------------------------------------------
' DeleteSheetSafe: supprime un onglet sans alerte
'--------------------------------------------------------------------
Private Sub DeleteSheetSafe(ByVal sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
End Sub
