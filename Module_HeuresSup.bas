Attribute VB_Name = "Module_HeuresSup"
'====================================================================
' MODULE HEURES SUPPLEMENTAIRES - Planning_2026_RUNTIME
' Gestion complete des heures supplementaires: calcul, classification
' par type (normal/nuit/WE/ferie), majorations legales belges,
' conversion en jours RCT, valorisation financiere, bilan annuel
'
' DEPENDANCES:
'   - Module_HeuresTravaillees: HeuresPresteesMois(), DureeEffectiveCode(),
'     HeuresTheoriquesMois(), JoursOuvrablesMois()
'   - Module_Planning_Core: BuildFeriesBE()
'   - Feuille "Personnel": HeuresStdJour (col D), %Temps mensuels
'   - Feuilles mensuelles: Janv..Dec (Row 6+ = agents, Col 3-33 = jours)
'
' REGLEMENTATION:
'   - Seuil hebdomadaire legal belge: 38h/semaine
'   - Nuit = prestation traversant la plage 22h-6h
'   - WE = samedi (vbMonday Weekday=6) ou dimanche (vbMonday Weekday=7)
'   - Ferie = jour present dans BuildFeriesBE()
'   - Majorations: 150% normal, 200% nuit/WE/ferie
'
' INSTALLATION:
'   1. Ouvrir Planning_2026_RUNTIME.xlsm
'   2. Alt+F11 > Menu Fichier > Importer > Module_HeuresSup.bas
'   3. Verifier que Module_HeuresTravaillees et Module_Planning_Core sont presents
'   4. Lancer: Alt+F8 > GenererBilanHeuresSup > Executer
'====================================================================

Option Explicit

' ---- CONSTANTES PUBLIQUES ----
Public Const SEUIL_HEBDO As Double = 38         ' heures/semaine standard Belgique
Public Const TAUX_NORMAL As Double = 1.5         ' 150%
Public Const TAUX_NUIT As Double = 2#            ' 200% (22h-6h)
Public Const TAUX_WE As Double = 2#              ' 200% samedi/dimanche
Public Const TAUX_FERIE As Double = 2#           ' 200%
Public Const HEURES_PAR_JOUR_RCT As Double = 7.6

' ---- CONSTANTES PRIVEES: STRUCTURE FEUILLES ----
Private Const MONTH_SHEETS As String = "Janv,Fev,Mars,Avril,Mai,Juin,Juil,Aout,Sept,Oct,Nov,Dec"
Private Const FIRST_EMP_ROW As Long = 6          ' Premiere ligne agent
Private Const FIRST_DAY_COL As Long = 3          ' Colonne C = jour 1
Private Const MAX_DAY_COL As Long = 33           ' Colonne AG = jour 31
Private Const BILAN_SHEET_NAME As String = "Bilan Heures Sup"

' ---- SEUILS HORAIRES NUIT ----
Private Const NUIT_DEBUT As Double = 22#         ' 22h00
Private Const NUIT_FIN As Double = 6#            ' 06h00

' ---- SEUIL ALERTE ----
Private Const SEUIL_ALERTE_HS_MOIS As Double = 20#  ' HS > 20h/mois = alerte rouge

'====================================================================
' FONCTION PUBLIQUE 1: CalculerHeuresSupMois
' Calcule les heures supplementaires d'un agent pour un mois donne.
' HS = Max(0, heures prestees - heures theoriques)
'
' @param nom      String   Nom agent (format "Nom_Prenom")
' @param annee    Long     Annee (ex: 2026)
' @param moisNum  Long     Numero du mois (1-12)
' @return Double  Heures supplementaires (>= 0)
'====================================================================
Public Function CalculerHeuresSupMois(ByVal nom As String, _
                                       ByVal annee As Long, _
                                       ByVal moisNum As Long) As Double
    Dim nomMois As String
    Dim rowAgent As Long
    Dim hPrest As Double
    Dim hTheo As Double
    Dim heuresStd As Double
    Dim pctTemps As Double
    Dim ws As Worksheet

    ' Validation
    If Len(Trim(nom)) = 0 Or moisNum < 1 Or moisNum > 12 Then
        CalculerHeuresSupMois = 0#
        Exit Function
    End If

    ' Recuperer la feuille du mois
    nomMois = GetMonthName(moisNum)
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nomMois)
    On Error GoTo 0
    If ws Is Nothing Then
        CalculerHeuresSupMois = 0#
        Exit Function
    End If

    ' Trouver la ligne de l'agent
    rowAgent = FindAgentRow(nomMois, nom)
    If rowAgent = 0 Then
        CalculerHeuresSupMois = 0#
        Exit Function
    End If

    ' Heures prestees via Module_HeuresTravaillees
    hPrest = Module_HeuresTravaillees.HeuresPresteesMois(nomMois, rowAgent)

    ' Heures theoriques via Module_HeuresTravaillees
    heuresStd = GetHeuresStdJourAgent(nom)
    If heuresStd <= 0 Then heuresStd = 7.6
    pctTemps = GetPctTempsAgent(nom, moisNum)
    If pctTemps <= 0 Then pctTemps = 1#
    hTheo = Module_HeuresTravaillees.HeuresTheoriquesMois(annee, moisNum, pctTemps, heuresStd)

    ' Delta: si positif = heures sup
    If hPrest > hTheo Then
        CalculerHeuresSupMois = Round(hPrest - hTheo, 2)
    Else
        CalculerHeuresSupMois = 0#
    End If
End Function

'====================================================================
' FONCTION PUBLIQUE 2: ClassifierHeuresSup
' Classe les heures supplementaires par type de majoration.
' Scanne le planning jour par jour et repartit proportionnellement
' les HS selon le type de prestation (ferie > nuit > WE > normal).
'
' @param nom      String   Nom agent (format "Nom_Prenom")
' @param annee    Long     Annee (ex: 2026)
' @param moisNum  Long     Numero du mois (1-12)
' @return Object  Dictionary: "normal"->X, "nuit"->Y, "we"->Z, "ferie"->W
'====================================================================
Public Function ClassifierHeuresSup(ByVal nom As String, _
                                     ByVal annee As Long, _
                                     ByVal moisNum As Long) As Object
    Dim dict As Object
    Dim totalHS As Double
    Dim nomMois As String
    Dim rowIdx As Long
    Dim ws As Worksheet
    Dim feries As Object
    Dim numJours As Long
    Dim col As Long
    Dim cellVal As String
    Dim heuresStd As Double
    Dim dureeJour As Double
    Dim dateJour As Date
    Dim jourNum As Long
    Dim cumNuit As Double, cumWE As Double, cumFerie As Double, cumNormal As Double
    Dim totalClasse As Double

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "normal", 0#
    dict.Add "nuit", 0#
    dict.Add "we", 0#
    dict.Add "ferie", 0#

    ' Calculer le total HS du mois
    totalHS = CalculerHeuresSupMois(nom, annee, moisNum)
    If totalHS <= 0 Then
        Set ClassifierHeuresSup = dict
        Exit Function
    End If

    ' Recuperer la feuille et la ligne agent
    nomMois = GetMonthName(moisNum)
    rowIdx = FindAgentRow(nomMois, nom)
    If rowIdx <= 0 Then
        dict("normal") = totalHS
        Set ClassifierHeuresSup = dict
        Exit Function
    End If

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nomMois)
    On Error GoTo 0
    If ws Is Nothing Then
        dict("normal") = totalHS
        Set ClassifierHeuresSup = dict
        Exit Function
    End If

    ' Charger les jours feries belges
    Set feries = Module_Planning_Core.BuildFeriesBE(annee)

    heuresStd = GetHeuresStdJourAgent(nom)
    If heuresStd <= 0 Then heuresStd = 7.6

    numJours = CompterJoursMois(ws)

    cumNuit = 0#
    cumWE = 0#
    cumFerie = 0#
    cumNormal = 0#

    ' Scanner chaque jour du mois
    For col = FIRST_DAY_COL To FIRST_DAY_COL + numJours - 1
        cellVal = Trim(SafeCStr(ws.Cells(rowIdx, col).Value))
        If Len(cellVal) = 0 Or cellVal = "0" Then GoTo NextDayClass

        dureeJour = Module_HeuresTravaillees.DureeEffectiveCode(cellVal, heuresStd)
        If dureeJour <= 0 Then GoTo NextDayClass

        ' Determiner la date du jour
        jourNum = col - FIRST_DAY_COL + 1
        If jourNum < 1 Or jourNum > 31 Then GoTo NextDayClass

        On Error Resume Next
        dateJour = DateSerial(annee, moisNum, jourNum)
        On Error GoTo 0

        ' Classification par priorite: ferie > nuit > WE > normal
        If EstFerie(dateJour, feries) Then
            cumFerie = cumFerie + dureeJour
        ElseIf EstNuit(cellVal) Then
            cumNuit = cumNuit + dureeJour
        ElseIf EstWeekend(dateJour) Then
            cumWE = cumWE + dureeJour
        Else
            cumNormal = cumNormal + dureeJour
        End If
NextDayClass:
    Next col

    ' Repartir les HS proportionnellement aux types de prestation
    totalClasse = cumNuit + cumWE + cumFerie + cumNormal
    If totalClasse > 0 Then
        dict("nuit") = Round(totalHS * (cumNuit / totalClasse), 2)
        dict("we") = Round(totalHS * (cumWE / totalClasse), 2)
        dict("ferie") = Round(totalHS * (cumFerie / totalClasse), 2)
        ' Normal absorbe le reste (evite erreurs d'arrondi)
        dict("normal") = Round(totalHS - dict("nuit") - dict("we") - dict("ferie"), 2)
        If dict("normal") < 0 Then dict("normal") = 0#
    Else
        dict("normal") = totalHS
    End If

    Set ClassifierHeuresSup = dict
End Function

'====================================================================
' FONCTION PUBLIQUE 3: GetTauxMajoration
' Retourne le taux de majoration pour un type d'heure supplementaire.
'
' @param typeHS  String   Type: "normal", "nuit", "we", "ferie"
' @return Double  Taux de majoration (1.5 ou 2.0)
'====================================================================
Public Function GetTauxMajoration(ByVal typeHS As String) As Double
    Select Case LCase(Trim(typeHS))
        Case "normal"
            GetTauxMajoration = TAUX_NORMAL
        Case "nuit"
            GetTauxMajoration = TAUX_NUIT
        Case "we", "weekend", "samedi", "dimanche"
            GetTauxMajoration = TAUX_WE
        Case "ferie", "jf"
            GetTauxMajoration = TAUX_FERIE
        Case Else
            GetTauxMajoration = TAUX_NORMAL
    End Select
End Function

'====================================================================
' FONCTION PUBLIQUE 4: ConvertirEnJoursRCT
' Convertit des heures supplementaires en jours de repos compensatoire.
'
' @param heuresSup     Double   Nombre d'heures supplementaires
' @param heuresStdJour Double   Heures standard par jour (7.6 par defaut)
' @return Double   Nombre de jours RCT equivalents
'====================================================================
Public Function ConvertirEnJoursRCT(ByVal heuresSup As Double, _
                                     ByVal heuresStdJour As Double) As Double
    If heuresStdJour <= 0 Then heuresStdJour = HEURES_PAR_JOUR_RCT
    If heuresSup <= 0 Then
        ConvertirEnJoursRCT = 0#
        Exit Function
    End If
    ConvertirEnJoursRCT = Round(heuresSup / heuresStdJour, 2)
End Function

'====================================================================
' FONCTION PUBLIQUE 5: ValorisationHeuresSup
' Calcule la valorisation financiere totale des heures supplementaires.
' Chaque type d'HS est multiplie par son taux de majoration et le
' taux horaire brut de l'agent.
'
' @param heuresNormales Double   HS taux normal (150%)
' @param heuresNuit     Double   HS nuit (200%)
' @param heuresWE       Double   HS WE (200%)
' @param heuresFerie    Double   HS ferie (200%)
' @param tauxHoraire    Double   Taux horaire brut de l'agent (EUR/h)
' @return Double   Montant total de la valorisation (EUR)
'====================================================================
Public Function ValorisationHeuresSup(ByVal heuresNormales As Double, _
                                       ByVal heuresNuit As Double, _
                                       ByVal heuresWE As Double, _
                                       ByVal heuresFerie As Double, _
                                       ByVal tauxHoraire As Double) As Double
    Dim total As Double

    If tauxHoraire <= 0 Then
        ValorisationHeuresSup = 0#
        Exit Function
    End If

    total = (heuresNormales * TAUX_NORMAL * tauxHoraire) _
          + (heuresNuit * TAUX_NUIT * tauxHoraire) _
          + (heuresWE * TAUX_WE * tauxHoraire) _
          + (heuresFerie * TAUX_FERIE * tauxHoraire)

    ValorisationHeuresSup = Round(total, 2)
End Function

'====================================================================
' SUB PUBLIQUE 6: GenererBilanHeuresSup
' Cree/MAJ l'onglet "Bilan Heures Sup" avec recap annuel complet.
' Structure: 1 ligne par agent, 12 mois x (HS Normal, HS Nuit,
'   HS WE, HS Ferie, Total HS, Jours RCT equiv), + TOTAL ANNUEL.
' En-tetes colorees, totaux gras, HS > 20h/mois en rouge.
'====================================================================
Public Sub GenererBilanHeuresSup()
    Dim mSheets() As String
    Dim ws As Worksheet, wsBilan As Worksheet
    Dim r As Long, m As Long, rowIdx As Long
    Dim agName As String
    Dim annee As Long
    Dim moisNum As Long
    Dim heuresStd As Double
    Dim dict As Object
    Dim hsNormal As Double, hsNuit As Double, hsWE As Double, hsFerie As Double
    Dim totalMois As Double, joursRCT As Double
    Dim totAnnNormal As Double, totAnnNuit As Double
    Dim totAnnWE As Double, totAnnFerie As Double
    Dim totAnnHS As Double, totAnnRCT As Double
    Dim baseCol As Long, totCol As Long, lastHdrCol As Long
    Dim agents As Object
    Dim agentList As New Collection
    Dim ag As Variant
    Dim rng As Range
    Dim c2 As Long
    Dim sc As Long
    Dim dataRow As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo CleanupBilan

    mSheets = Split(MONTH_SHEETS, ",")
    annee = Year(Now)

    ' Creer/recreer la feuille Bilan Heures Sup
    DeleteSheetSafe BILAN_SHEET_NAME
    Set wsBilan = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsBilan.Name = BILAN_SHEET_NAME

    ' ---- COLLECTER LES AGENTS ----
    Set agents = CreateObject("Scripting.Dictionary")
    agents.CompareMode = vbTextCompare

    For m = 0 To 11
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(mSheets(m))
        On Error GoTo CleanupBilan
        If Not ws Is Nothing Then
            For r = FIRST_EMP_ROW To GetLastEmployeeRow(ws)
                agName = Trim(SafeCStr(ws.Cells(r, 1).Value))
                If Len(agName) > 0 Then
                    If InStr(agName, "Remplacement") = 0 And agName <> "Us Nuit" Then
                        If Not agents.Exists(agName) Then
                            agents.Add agName, True
                            agentList.Add agName
                        End If
                    End If
                End If
            Next r
        End If
    Next m

    ' ---- EN-TETES ----
    Dim moisLabels As Variant
    moisLabels = Array("Janvier", "Fevrier", "Mars", "Avril", "Mai", "Juin", _
                       "Juillet", "Aout", "Septembre", "Octobre", "Novembre", "Decembre")

    wsBilan.Cells(1, 1).Value = "Agent"
    wsBilan.Range(wsBilan.Cells(1, 1), wsBilan.Cells(2, 1)).Merge
    FormatHeaderCell wsBilan.Cells(1, 1), RGB(31, 78, 121)

    For m = 0 To 11
        baseCol = 2 + m * 6
        wsBilan.Range(wsBilan.Cells(1, baseCol), wsBilan.Cells(1, baseCol + 5)).Merge
        wsBilan.Cells(1, baseCol).Value = moisLabels(m)
        FormatHeaderCell wsBilan.Cells(1, baseCol), RGB(31, 78, 121)

        ' Sous-en-tetes (ligne 2)
        wsBilan.Cells(2, baseCol).Value = "HS Norm"
        wsBilan.Cells(2, baseCol + 1).Value = "HS Nuit"
        wsBilan.Cells(2, baseCol + 2).Value = "HS WE"
        wsBilan.Cells(2, baseCol + 3).Value = "HS Fer"
        wsBilan.Cells(2, baseCol + 4).Value = "Total HS"
        wsBilan.Cells(2, baseCol + 5).Value = "RCT (j)"

        For sc = 0 To 5
            FormatSubHeaderCell wsBilan.Cells(2, baseCol + sc), RGB(68, 114, 196)
        Next sc
    Next m

    ' Totaux annuels (6 colonnes)
    totCol = 2 + 12 * 6
    lastHdrCol = totCol + 5

    wsBilan.Range(wsBilan.Cells(1, totCol), wsBilan.Cells(1, lastHdrCol)).Merge
    wsBilan.Cells(1, totCol).Value = "TOTAL ANNUEL"
    FormatHeaderCell wsBilan.Cells(1, totCol), RGB(146, 56, 52)

    wsBilan.Cells(2, totCol).Value = "HS Norm"
    wsBilan.Cells(2, totCol + 1).Value = "HS Nuit"
    wsBilan.Cells(2, totCol + 2).Value = "HS WE"
    wsBilan.Cells(2, totCol + 3).Value = "HS Fer"
    wsBilan.Cells(2, totCol + 4).Value = "Total HS"
    wsBilan.Cells(2, totCol + 5).Value = "RCT (j)"

    Dim tc As Long
    For tc = 0 To 5
        FormatSubHeaderCell wsBilan.Cells(2, totCol + tc), RGB(192, 80, 77)
    Next tc

    ' ---- REMPLISSAGE DES DONNEES ----
    rowIdx = 3

    For Each ag In agentList
        agName = CStr(ag)
        wsBilan.Cells(rowIdx, 1).Value = agName
        wsBilan.Cells(rowIdx, 1).Font.Bold = True

        heuresStd = GetHeuresStdJourAgent(agName)
        If heuresStd <= 0 Then heuresStd = 7.6

        totAnnNormal = 0#
        totAnnNuit = 0#
        totAnnWE = 0#
        totAnnFerie = 0#
        totAnnHS = 0#
        totAnnRCT = 0#

        For m = 0 To 11
            moisNum = m + 1

            Set dict = ClassifierHeuresSup(agName, annee, moisNum)

            hsNormal = dict("normal")
            hsNuit = dict("nuit")
            hsWE = dict("we")
            hsFerie = dict("ferie")
            totalMois = hsNormal + hsNuit + hsWE + hsFerie
            joursRCT = ConvertirEnJoursRCT(totalMois, heuresStd)

            baseCol = 2 + m * 6

            wsBilan.Cells(rowIdx, baseCol).Value = Round(hsNormal, 2)
            wsBilan.Cells(rowIdx, baseCol).NumberFormat = "0.00"

            wsBilan.Cells(rowIdx, baseCol + 1).Value = Round(hsNuit, 2)
            wsBilan.Cells(rowIdx, baseCol + 1).NumberFormat = "0.00"

            wsBilan.Cells(rowIdx, baseCol + 2).Value = Round(hsWE, 2)
            wsBilan.Cells(rowIdx, baseCol + 2).NumberFormat = "0.00"

            wsBilan.Cells(rowIdx, baseCol + 3).Value = Round(hsFerie, 2)
            wsBilan.Cells(rowIdx, baseCol + 3).NumberFormat = "0.00"

            wsBilan.Cells(rowIdx, baseCol + 4).Value = Round(totalMois, 2)
            wsBilan.Cells(rowIdx, baseCol + 4).NumberFormat = "0.00"
            wsBilan.Cells(rowIdx, baseCol + 4).Font.Bold = True

            wsBilan.Cells(rowIdx, baseCol + 5).Value = Round(joursRCT, 2)
            wsBilan.Cells(rowIdx, baseCol + 5).NumberFormat = "0.00"

            If totalMois > SEUIL_ALERTE_HS_MOIS Then
                wsBilan.Cells(rowIdx, baseCol + 4).Font.Color = RGB(204, 0, 0)
                wsBilan.Cells(rowIdx, baseCol + 4).Interior.Color = RGB(255, 230, 230)
            End If

            totAnnNormal = totAnnNormal + hsNormal
            totAnnNuit = totAnnNuit + hsNuit
            totAnnWE = totAnnWE + hsWE
            totAnnFerie = totAnnFerie + hsFerie
            totAnnHS = totAnnHS + totalMois
            totAnnRCT = totAnnRCT + joursRCT
        Next m

        wsBilan.Cells(rowIdx, totCol).Value = Round(totAnnNormal, 2)
        wsBilan.Cells(rowIdx, totCol).NumberFormat = "0.00"
        wsBilan.Cells(rowIdx, totCol).Font.Bold = True

        wsBilan.Cells(rowIdx, totCol + 1).Value = Round(totAnnNuit, 2)
        wsBilan.Cells(rowIdx, totCol + 1).NumberFormat = "0.00"
        wsBilan.Cells(rowIdx, totCol + 1).Font.Bold = True

        wsBilan.Cells(rowIdx, totCol + 2).Value = Round(totAnnWE, 2)
        wsBilan.Cells(rowIdx, totCol + 2).NumberFormat = "0.00"
        wsBilan.Cells(rowIdx, totCol + 2).Font.Bold = True

        wsBilan.Cells(rowIdx, totCol + 3).Value = Round(totAnnFerie, 2)
        wsBilan.Cells(rowIdx, totCol + 3).NumberFormat = "0.00"
        wsBilan.Cells(rowIdx, totCol + 3).Font.Bold = True

        wsBilan.Cells(rowIdx, totCol + 4).Value = Round(totAnnHS, 2)
        wsBilan.Cells(rowIdx, totCol + 4).NumberFormat = "0.00"
        wsBilan.Cells(rowIdx, totCol + 4).Font.Bold = True

        wsBilan.Cells(rowIdx, totCol + 5).Value = Round(totAnnRCT, 2)
        wsBilan.Cells(rowIdx, totCol + 5).NumberFormat = "0.00"
        wsBilan.Cells(rowIdx, totCol + 5).Font.Bold = True

        If totAnnHS > SEUIL_ALERTE_HS_MOIS * 12 Then
            wsBilan.Cells(rowIdx, totCol + 4).Font.Color = RGB(204, 0, 0)
            wsBilan.Cells(rowIdx, totCol + 4).Interior.Color = RGB(255, 230, 230)
        End If

        rowIdx = rowIdx + 1
    Next ag

    ' ---- FORMATAGE GLOBAL ----
    wsBilan.Columns("A").ColumnWidth = 28
    For c2 = 2 To lastHdrCol
        wsBilan.Columns(c2).ColumnWidth = 8
    Next c2

    If rowIdx > 3 Then
        Set rng = wsBilan.Range(wsBilan.Cells(1, 1), wsBilan.Cells(rowIdx - 1, lastHdrCol))
        rng.Borders.LineStyle = xlContinuous
        rng.Borders.Weight = xlThin
        rng.Font.Name = "Calibri"
        rng.Font.Size = 9

        For dataRow = 3 To rowIdx - 1
            If (dataRow Mod 2) = 1 Then
                wsBilan.Range(wsBilan.Cells(dataRow, 1), _
                              wsBilan.Cells(dataRow, lastHdrCol)).Interior.Color = RGB(242, 242, 242)
            End If
        Next dataRow
    End If

    wsBilan.Range("B3").Select
    ActiveWindow.FreezePanes = True
    wsBilan.Activate

CleanupBilan:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

'====================================================================
' SUB PUBLIQUE 7: RecalculerHeuresSupMois
' Recalcule les heures sup pour un seul mois et met a jour le bilan.
' Si le bilan n'existe pas, genere le bilan complet.
'
' @param nomMois  String   Nom du mois ("Janv", "Fev", etc.)
'====================================================================
Public Sub RecalculerHeuresSupMois(ByVal nomMois As String)
    Dim wsBilan As Worksheet
    Dim moisNum As Long
    Dim annee As Long
    Dim rowIdx As Long
    Dim agName As String
    Dim heuresStd As Double
    Dim dict As Object
    Dim hsNormal As Double, hsNuit As Double, hsWE As Double, hsFerie As Double
    Dim totalMois As Double, joursRCT As Double
    Dim baseCol As Long
    Dim ws As Worksheet

    moisNum = GetMonthNum(nomMois)
    If moisNum = 0 Then
        MsgBox "Mois non reconnu: " & nomMois, vbExclamation, "Heures Sup"
        Exit Sub
    End If

    On Error Resume Next
    Set wsBilan = ThisWorkbook.Sheets(BILAN_SHEET_NAME)
    On Error GoTo 0
    If wsBilan Is Nothing Then
        GenererBilanHeuresSup
        Exit Sub
    End If

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nomMois)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    annee = GetAnneeFromSheet(ws)
    If annee = 0 Then annee = Year(Now)

    baseCol = 2 + (moisNum - 1) * 6

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo CleanupRecalc

    rowIdx = 3
    Do While Len(Trim(SafeCStr(wsBilan.Cells(rowIdx, 1).Value))) > 0
        agName = Trim(SafeCStr(wsBilan.Cells(rowIdx, 1).Value))

        heuresStd = GetHeuresStdJourAgent(agName)
        If heuresStd <= 0 Then heuresStd = 7.6

        Set dict = ClassifierHeuresSup(agName, annee, moisNum)

        hsNormal = dict("normal")
        hsNuit = dict("nuit")
        hsWE = dict("we")
        hsFerie = dict("ferie")
        totalMois = hsNormal + hsNuit + hsWE + hsFerie
        joursRCT = ConvertirEnJoursRCT(totalMois, heuresStd)

        wsBilan.Cells(rowIdx, baseCol).Value = Round(hsNormal, 2)
        wsBilan.Cells(rowIdx, baseCol + 1).Value = Round(hsNuit, 2)
        wsBilan.Cells(rowIdx, baseCol + 2).Value = Round(hsWE, 2)
        wsBilan.Cells(rowIdx, baseCol + 3).Value = Round(hsFerie, 2)
        wsBilan.Cells(rowIdx, baseCol + 4).Value = Round(totalMois, 2)
        wsBilan.Cells(rowIdx, baseCol + 5).Value = Round(joursRCT, 2)

        wsBilan.Cells(rowIdx, baseCol + 4).Font.Color = RGB(0, 0, 0)
        wsBilan.Cells(rowIdx, baseCol + 4).Font.Bold = True
        wsBilan.Cells(rowIdx, baseCol + 4).Interior.ColorIndex = xlNone
        If totalMois > SEUIL_ALERTE_HS_MOIS Then
            wsBilan.Cells(rowIdx, baseCol + 4).Font.Color = RGB(204, 0, 0)
            wsBilan.Cells(rowIdx, baseCol + 4).Interior.Color = RGB(255, 230, 230)
        End If

        RecalculerTotauxAnnuelsAgent wsBilan, rowIdx

        rowIdx = rowIdx + 1
    Loop

CleanupRecalc:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

'====================================================================
' FONCTIONS PRIVEES: HELPERS
'====================================================================

'--------------------------------------------------------------------
' SafeCStr: conversion safe cellule -> String (evite Type Mismatch)
'--------------------------------------------------------------------
Private Function SafeCStr(ByVal v As Variant) As String
    If IsError(v) Then
        SafeCStr = ""
    ElseIf IsNull(v) Or IsEmpty(v) Then
        SafeCStr = ""
    Else
        SafeCStr = CStr(v)
    End If
End Function

'--------------------------------------------------------------------
' SafeDbl: conversion safe cellule -> Double
'--------------------------------------------------------------------
Private Function SafeDbl(ByVal v As Variant) As Double
    If IsError(v) Then
        SafeDbl = 0#
    ElseIf IsNull(v) Or IsEmpty(v) Then
        SafeDbl = 0#
    ElseIf IsNumeric(v) Then
        SafeDbl = CDbl(v)
    Else
        SafeDbl = 0#
    End If
End Function

'--------------------------------------------------------------------
' EstNuit: detecte si un code horaire est un shift de nuit.
' Nuit si passage minuit (debut > fin) ou debut >= 22h.
' Ex: "19:45 6:45" -> True, "7 15:30" -> False, "22 6" -> True
'--------------------------------------------------------------------
Private Function EstNuit(ByVal code As String) As Boolean
    Dim c As String
    Dim parts() As String
    Dim hDebut As Double, hFin As Double

    EstNuit = False
    c = Trim(code)

    Do While InStr(c, "  ") > 0
        c = Replace(c, "  ", " ")
    Loop

    parts = Split(c, " ")
    If UBound(parts) < 1 Then Exit Function

    hDebut = ParseHeure(parts(0))
    hFin = ParseHeure(parts(UBound(parts)))

    If hDebut < 0 Or hFin < 0 Then Exit Function

    If hDebut > hFin Then
        EstNuit = True
    ElseIf hDebut >= NUIT_DEBUT Then
        EstNuit = True
    End If
End Function

'--------------------------------------------------------------------
' EstWeekend: detecte si une date tombe un samedi ou dimanche.
' Weekday avec vbMonday: Lun=1, ..., Sam=6, Dim=7
'--------------------------------------------------------------------
Private Function EstWeekend(ByVal d As Date) As Boolean
    Dim wd As Long
    wd = Weekday(d, vbMonday)
    EstWeekend = (wd = 6 Or wd = 7)
End Function

'--------------------------------------------------------------------
' EstFerie: verifie si une date est un jour ferie belge
'--------------------------------------------------------------------
Private Function EstFerie(ByVal d As Date, ByVal feries As Object) As Boolean
    EstFerie = feries.Exists(CStr(d))
End Function

'--------------------------------------------------------------------
' ParseHeure: "15:30" -> 15.5, "7" -> 7.0, "19:45" -> 19.75
' Retourne -1 si non parsable
'--------------------------------------------------------------------
Private Function ParseHeure(ByVal txt As String) As Double
    Dim parts() As String
    Dim h As Double, m As Double

    txt = Trim(txt)
    If Len(txt) = 0 Then
        ParseHeure = -1#
        Exit Function
    End If

    If InStr(txt, ":") > 0 Then
        parts = Split(txt, ":")
        If UBound(parts) <> 1 Then
            ParseHeure = -1#
            Exit Function
        End If
        If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Then
            ParseHeure = -1#
            Exit Function
        End If
        h = CDbl(parts(0))
        m = CDbl(parts(1))
        If h < 0 Or h > 23 Or m < 0 Or m > 59 Then
            ParseHeure = -1#
            Exit Function
        End If
        ParseHeure = h + (m / 60#)
    Else
        If Not IsNumeric(txt) Then
            ParseHeure = -1#
            Exit Function
        End If
        h = CDbl(txt)
        If h < 0 Or h > 23 Then
            ParseHeure = -1#
            Exit Function
        End If
        ParseHeure = h
    End If
End Function

'--------------------------------------------------------------------
' GetMonthName: numero de mois -> nom d'onglet
'--------------------------------------------------------------------
Private Function GetMonthName(ByVal moisNum As Long) As String
    Dim mSheets() As String
    mSheets = Split(MONTH_SHEETS, ",")
    If moisNum >= 1 And moisNum <= 12 Then
        GetMonthName = mSheets(moisNum - 1)
    Else
        GetMonthName = ""
    End If
End Function

'--------------------------------------------------------------------
' GetMonthNum: nom d'onglet -> numero de mois
'--------------------------------------------------------------------
Private Function GetMonthNum(ByVal nomMois As String) As Long
    Dim mSheets() As String
    Dim i As Long

    mSheets = Split(MONTH_SHEETS, ",")
    For i = 0 To 11
        If UCase(Trim(mSheets(i))) = UCase(Trim(nomMois)) Then
            GetMonthNum = i + 1
            Exit Function
        End If
    Next i
    GetMonthNum = 0
End Function

'--------------------------------------------------------------------
' FindAgentRow: cherche la ligne d'un agent dans une feuille mensuelle.
' Retourne 0 si non trouve.
'--------------------------------------------------------------------
Private Function FindAgentRow(ByVal nomMois As String, ByVal agentName As String) As Long
    Dim ws As Worksheet
    Dim r As Long
    Dim cellName As String

    FindAgentRow = 0

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nomMois)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    For r = FIRST_EMP_ROW To GetLastEmployeeRow(ws)
        cellName = Trim(SafeCStr(ws.Cells(r, 1).Value))
        If UCase(cellName) = UCase(Trim(agentName)) Then
            FindAgentRow = r
            Exit Function
        End If
    Next r
End Function

'--------------------------------------------------------------------
' GetLastEmployeeRow: derniere ligne agent (5 lignes vides = fin)
'--------------------------------------------------------------------
Private Function GetLastEmployeeRow(ByVal ws As Worksheet) As Long
    Dim r As Long
    Dim emptyCount As Long
    Dim lastFound As Long

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
' CompterJoursMois: nombre de jours via en-tetes ligne 4
'--------------------------------------------------------------------
Private Function CompterJoursMois(ByVal ws As Worksheet) As Long
    Dim col As Long
    Dim compteur As Long

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
' GetHeuresStdJourAgent: heuresStdJour depuis feuille Personnel
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
' GetPctTempsAgent: % temps mensuel depuis Personnel
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
' GetAnneeFromSheet: annee depuis cellules A1/A2
'--------------------------------------------------------------------
Private Function GetAnneeFromSheet(ByVal ws As Worksheet) As Long
    Dim val1 As String, val2 As String

    val1 = SafeCStr(ws.Cells(1, 1).Value)
    val2 = SafeCStr(ws.Cells(2, 1).Value)

    GetAnneeFromSheet = ExtractYear(val1)
    If GetAnneeFromSheet > 0 Then Exit Function

    GetAnneeFromSheet = ExtractYear(val2)
    If GetAnneeFromSheet > 0 Then Exit Function

    GetAnneeFromSheet = Year(Now)
End Function

'--------------------------------------------------------------------
' ExtractYear: extrait une annee (2020-2035) d'une chaine
'--------------------------------------------------------------------
Private Function ExtractYear(ByVal txt As String) As Long
    Dim i As Long
    Dim chunk As String

    ExtractYear = 0
    If Len(txt) < 4 Then Exit Function

    For i = 1 To Len(txt) - 3
        chunk = Mid(txt, i, 4)
        If IsNumeric(chunk) Then
            If CLng(chunk) >= 2020 And CLng(chunk) <= 2035 Then
                ExtractYear = CLng(chunk)
                Exit Function
            End If
        End If
    Next i
End Function

'--------------------------------------------------------------------
' RecalculerTotauxAnnuelsAgent: somme des 12 mois -> totaux annuels
'--------------------------------------------------------------------
Private Sub RecalculerTotauxAnnuelsAgent(ByVal wsBilan As Worksheet, ByVal rowIdx As Long)
    Dim m As Long
    Dim baseCol As Long, totCol As Long
    Dim totNormal As Double, totNuit As Double
    Dim totWE As Double, totFerie As Double
    Dim totHS As Double, totRCT As Double
    Dim tc2 As Long

    totCol = 2 + 12 * 6

    totNormal = 0#
    totNuit = 0#
    totWE = 0#
    totFerie = 0#
    totHS = 0#
    totRCT = 0#

    For m = 0 To 11
        baseCol = 2 + m * 6
        totNormal = totNormal + SafeDbl(wsBilan.Cells(rowIdx, baseCol).Value)
        totNuit = totNuit + SafeDbl(wsBilan.Cells(rowIdx, baseCol + 1).Value)
        totWE = totWE + SafeDbl(wsBilan.Cells(rowIdx, baseCol + 2).Value)
        totFerie = totFerie + SafeDbl(wsBilan.Cells(rowIdx, baseCol + 3).Value)
        totHS = totHS + SafeDbl(wsBilan.Cells(rowIdx, baseCol + 4).Value)
        totRCT = totRCT + SafeDbl(wsBilan.Cells(rowIdx, baseCol + 5).Value)
    Next m

    wsBilan.Cells(rowIdx, totCol).Value = Round(totNormal, 2)
    wsBilan.Cells(rowIdx, totCol + 1).Value = Round(totNuit, 2)
    wsBilan.Cells(rowIdx, totCol + 2).Value = Round(totWE, 2)
    wsBilan.Cells(rowIdx, totCol + 3).Value = Round(totFerie, 2)
    wsBilan.Cells(rowIdx, totCol + 4).Value = Round(totHS, 2)
    wsBilan.Cells(rowIdx, totCol + 5).Value = Round(totRCT, 2)

    For tc2 = 0 To 5
        wsBilan.Cells(rowIdx, totCol + tc2).Font.Bold = True
        wsBilan.Cells(rowIdx, totCol + tc2).NumberFormat = "0.00"
    Next tc2

    wsBilan.Cells(rowIdx, totCol + 4).Font.Color = RGB(0, 0, 0)
    wsBilan.Cells(rowIdx, totCol + 4).Interior.ColorIndex = xlNone
    If totHS > SEUIL_ALERTE_HS_MOIS * 12 Then
        wsBilan.Cells(rowIdx, totCol + 4).Font.Color = RGB(204, 0, 0)
        wsBilan.Cells(rowIdx, totCol + 4).Interior.Color = RGB(255, 230, 230)
    End If
End Sub

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

'--------------------------------------------------------------------
' FormatHeaderCell: en-tete principale (fond, blanc, gras, centre)
'--------------------------------------------------------------------
Private Sub FormatHeaderCell(ByVal cell As Range, ByVal bgColor As Long)
    With cell
        .Interior.Color = bgColor
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 10
        .Font.Name = "Calibri"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
End Sub

'--------------------------------------------------------------------
' FormatSubHeaderCell: sous-en-tete (fond, blanc, gras, taille 8)
'--------------------------------------------------------------------
Private Sub FormatSubHeaderCell(ByVal cell As Range, ByVal bgColor As Long)
    With cell
        .Interior.Color = bgColor
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 8
        .Font.Name = "Calibri"
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
End Sub
