Attribute VB_Name = "Module_ListeDropdown"
Option Explicit

' ============================================================================
' MACRO: Genere automatiquement la liste des codes pour les listes deroulantes
' Codes tries par heure de debut PUIS par heure de fin
' ============================================================================

Public Sub GenererListeCodesDropdown()
    Dim wsConfig As Worksheet, wsSpec As Worksheet, wsListe As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim code As String, hStart As String, hEnd As String
    Dim arrCodes() As String
    Dim arrHeures() As Double
    Dim arrHeuresFin() As Double
    Dim cntConfig As Long, cntSpec As Long
    Dim tempCode As String, tempHeure As Double, tempFin As Double
    
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("Config_Codes")
    Set wsSpec = ThisWorkbook.Sheets("Codes_Speciaux")
    On Error GoTo 0
    
    If wsConfig Is Nothing And wsSpec Is Nothing Then
        MsgBox "Aucune feuille source trouvee", vbExclamation
        Exit Sub
    End If
    
    On Error Resume Next
    Set wsListe = ThisWorkbook.Sheets("Liste_Codes")
    On Error GoTo 0
    
    If wsListe Is Nothing Then
        Set wsListe = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsListe.Name = "Liste_Codes"
    Else
        wsListe.Cells.Clear
    End If
    
    ' Charger Config_Codes avec heures de debut ET fin
    cntConfig = 0
    If Not wsConfig Is Nothing Then
        lastRow = wsConfig.Cells(wsConfig.Rows.Count, "A").End(xlUp).Row
        ReDim arrCodes(1 To lastRow)
        ReDim arrHeures(1 To lastRow)
        ReDim arrHeuresFin(1 To lastRow)
        
        For i = 2 To lastRow
            code = Trim(CStr(wsConfig.Cells(i, 1).Value))
            If code <> "" And UCase(code) <> "CODE" Then
                cntConfig = cntConfig + 1
                arrCodes(cntConfig) = code
                ' Lire heure de debut (colonne F) et fin (colonne I)
                hStart = Trim(CStr(wsConfig.Cells(i, 6).Value))
                hEnd = Trim(CStr(wsConfig.Cells(i, 9).Value))
                arrHeures(cntConfig) = ConvertirHeureSimple(hStart)
                arrHeuresFin(cntConfig) = ConvertirHeureSimple(hEnd)
            End If
        Next i
    End If
    
    ' Trier par heure de debut, puis par heure de fin
    If cntConfig > 1 Then
        For i = 1 To cntConfig - 1
            For j = i + 1 To cntConfig
                ' Comparer d'abord par heure de debut, puis par heure de fin
                If arrHeures(j) < arrHeures(i) Or _
                   (arrHeures(j) = arrHeures(i) And arrHeuresFin(j) < arrHeuresFin(i)) Then
                    ' Echanger codes
                    tempCode = arrCodes(i)
                    tempHeure = arrHeures(i)
                    tempFin = arrHeuresFin(i)
                    
                    arrCodes(i) = arrCodes(j)
                    arrHeures(i) = arrHeures(j)
                    arrHeuresFin(i) = arrHeuresFin(j)
                    
                    arrCodes(j) = tempCode
                    arrHeures(j) = tempHeure
                    arrHeuresFin(j) = tempFin
                End If
            Next j
        Next i
    End If
    
    ' Ecrire codes horaires tries (format TEXTE pour eviter interpretation date)
    wsListe.Columns("A").NumberFormat = "@"  ' Format texte
    wsListe.Cells(1, 1).Value = "Code"
    For i = 1 To cntConfig
        wsListe.Cells(i + 1, 1).Value = "'" & arrCodes(i)  ' Apostrophe force texte
    Next i
    
    ' Ajouter Codes_Speciaux a la fin
    cntSpec = 0
    If Not wsSpec Is Nothing Then
        lastRow = wsSpec.Cells(wsSpec.Rows.Count, "A").End(xlUp).Row
        For i = 2 To lastRow
            code = Trim(CStr(wsSpec.Cells(i, 1).Value))
            If code <> "" And UCase(code) <> "CODE" Then
                cntSpec = cntSpec + 1
                wsListe.Cells(cntConfig + cntSpec + 1, 1).Value = "'" & code
            End If
        Next i
    End If
    
    ' Creer la plage nommee
    Dim totalRows As Long
    totalRows = cntConfig + cntSpec + 1
    
    On Error Resume Next
    ThisWorkbook.Names("ListeCodes").Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="ListeCodes", RefersTo:=wsListe.Range("A2:A" & totalRows)
    
    wsListe.Columns("A").AutoFit
    MsgBox "Liste generee avec " & (cntConfig + cntSpec) & " codes!" & vbCrLf & _
           "Codes horaires: " & cntConfig & " (tries par heure)" & vbCrLf & _
           "Codes speciaux: " & cntSpec, vbInformation
End Sub

Private Function ConvertirHeureSimple(ByVal s As String) As Double
    If s = "" Then
        ConvertirHeureSimple = 99
        Exit Function
    End If
    
    s = Trim(s)
    s = Replace(s, ",", ".")
    
    Dim res As Double
    If InStr(s, ":") > 0 Then
        Dim p() As String
        p = Split(s, ":")
        res = Val(p(0)) + Val(p(1)) / 60
    Else
        res = Val(s)
    End If
    
    ' Normalisation: Si < 1 (format date/heure Excel), convertir en heures (x24)
    ' Attention: 0 reste 0
    If res > 0 And res <= 1.05 Then ' Tolerance pour 1h (0.04) jusqu'a ~25h? Non, 1.0 = 24h.
                                    ' 0.5 = 12h.
                                    ' Cas limite : Code "1" (1h du matin). Res=1.
                                    ' Est-ce 1h (texte) ou 24h (serie 1.0)?
                                    ' En general, H_Start <= 24.
                                    ' Si on a un melange, "1" texte = 1. "1:00" texte = 1.
                                    ' Serie 1:00 = 0.041.
                                    ' Donc seuil ~1.1 semble sur.
        ' Mais attention si code = "0.5" (30 min duration). Ici c'est H_Start.
        res = res * 24
    End If
    
    ConvertirHeureSimple = res
End Function

' ============================================================================
' MACRO: Rafraichir les couleurs sur la feuille planning active
' ============================================================================

Public Sub RafraichirCouleursPlanning()
    Dim ws As Worksheet
    Dim rngPlanning As Range
    Dim data As Variant, arrNoms As Variant
    Dim r As Long, c As Long
    Dim code As String, nom As String, jour As String
    Dim colDebut As Long, colFin As Long
    Dim ligDebut As Long, ligFin As Long
    
    ' Plages pour groupement (optimisation massive)
    
    Set ws = ActiveSheet
    
    colDebut = 2
    colFin = 32
    ligDebut = 5 ' Jours/Dates en 3-4, Donnees commencent en 5
    ligFin = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If ligFin < ligDebut Then
        MsgBox "Aucune donnee trouvee", vbExclamation
        Exit Sub
    End If
    
    ' Desactiver TOUT ce qui ralentit
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Charger En-tetes (Jours Ligne 3, Dates Ligne 4)
    Dim arrJours As Variant, arrDates As Variant
    arrJours = ws.Range(ws.Cells(3, colDebut), ws.Cells(3, colFin)).Value
    arrDates = ws.Range(ws.Cells(4, colDebut), ws.Cells(4, colFin)).Value
    
    ' Charger Planning + Noms en memoire
    Set rngPlanning = ws.Range(ws.Cells(ligDebut, colDebut), ws.Cells(ligFin, colFin))
    data = rngPlanning.Value
    
    ' Charger colonne A (Noms) correspondante
    arrNoms = ws.Range(ws.Cells(ligDebut, 1), ws.Cells(ligFin, 1)).Value
    
    ' --- CHARGEMENT EXCEPTIONS (Multi-personnes) ---
    Dim arrExceptions() As Variant
    Dim nbExceptions As Long
    Dim wsEx As Worksheet
    
    nbExceptions = 0
    
    ' OPTIMISATION: Verif rapide si les regles globales existent deja (on cherche "WE" comme indicateur)
    Dim wsCheck As Worksheet
    Dim needsInit As Boolean
    needsInit = True
    
    On Error Resume Next
    Set wsCheck = ThisWorkbook.Sheets("Config_Exceptions")
    On Error GoTo 0
    
    If Not wsCheck Is Nothing Then
        Dim rngFind As Range
        Set rngFind = wsCheck.Columns("B").Find(What:="WE", LookIn:=xlValues, LookAt:=xlWhole)
        If Not rngFind Is Nothing Then needsInit = False
    End If
    
    If needsInit Then
        On Error Resume Next
        InitialiserReglesDefaut
        On Error GoTo 0
    End If
    
    On Error Resume Next
    Set wsEx = ThisWorkbook.Sheets("Config_Exceptions")
    On Error GoTo 0
    
    If Not wsEx Is Nothing Then
        ' Format attendu: Col A=NomPattern, B=Code, C=Jours, D=DateDeb, E=DateFin, F=Couleur
        Dim lEx As Long
        lEx = wsEx.Cells(wsEx.Rows.Count, "A").End(xlUp).Row
        If lEx >= 2 Then
            ' On charge A jusqu'a F (6 colonnes)
            arrExceptions = wsEx.Range("A2:F" & lEx).Value
            nbExceptions = UBound(arrExceptions, 1)
        End If
    Else
        ' Tentative d'auto-generation si module present
        On Error Resume Next
        InitialiserReglesDefaut
        Set wsEx = ThisWorkbook.Sheets("Config_Exceptions")
        If Not wsEx Is Nothing Then
             Dim lEx2 As Long
             lEx2 = wsEx.Cells(wsEx.Rows.Count, "A").End(xlUp).Row
             If lEx2 >= 2 Then
                arrExceptions = wsEx.Range("A2:F" & lEx2).Value
                nbExceptions = UBound(arrExceptions, 1)
             End If
        End If
        On Error GoTo 0
    End If
    
    ' Fallback : Si pas de feuille ou vide, s'assurer que les regles globales existent
    If nbExceptions = 0 Then
        ' Appel auto-generation si disponible
        On Error Resume Next
        InitialiserReglesDefaut
        On Error GoTo 0
        
        ' Recharger apres generation
        Set wsEx = ThisWorkbook.Sheets("Config_Exceptions")
        If Not wsEx Is Nothing Then
            Dim lExRetry As Long
            lExRetry = wsEx.Cells(wsEx.Rows.Count, "A").End(xlUp).Row
            If lExRetry >= 2 Then
                arrExceptions = wsEx.Range("A2:F" & lExRetry).Value
                nbExceptions = UBound(arrExceptions, 1)
            End If
        End If
    End If
    
    ' Si toujours rien, on met une regle bidon pour eviter erreur
    If nbExceptions = 0 Then
        ReDim arrExceptions(1 To 1, 1 To 6)
        arrExceptions(1, 1) = "*"
        arrExceptions(1, 2) = "WE"
        arrExceptions(1, 3) = ""
        arrExceptions(1, 4) = Empty
        arrExceptions(1, 5) = Empty
        arrExceptions(1, 6) = "BLEU"
        nbExceptions = 1
    End If
    
    ' Initialisation des plages de couleurs pour application en masse
    Dim rngBleu As Range, rngRouge As Range, rngJaune As Range
    Dim rngCyan As Range, rngGris As Range, rngOrange As Range, rngRose As Range
    Dim rngBleuClair As Range
    
    Set rngBleu = Nothing: Set rngRouge = Nothing: Set rngJaune = Nothing
    Set rngCyan = Nothing: Set rngGris = Nothing: Set rngOrange = Nothing: Set rngRose = Nothing
    Set rngBleuClair = Nothing

    ' Boucle sur le tableau en memoire
    For r = 1 To UBound(data, 1)
        nom = UCase(Trim(CStr(arrNoms(r, 1)))) ' Nom de la ligne
        
        For c = 1 To UBound(data, 2)
            code = UCase(Trim(CStr(data(r, c))))
            If code <> "" Then
                Dim cell As Range
                Set cell = rngPlanning.Cells(r, c)
                
                ' Identifier le jour et la date depuis nos tableaux d'en-tetes
                jour = UCase(Trim(CStr(arrJours(1, c))))
                Dim dateCourante As Variant
                dateCourante = arrDates(1, c) ' Peut etre un Double (Date) ou String
                
                ' =========================================================
                ' REGLES PRIORITAIRES (Exceptions Multiples)
                ' =========================================================
                Dim iEx As Long
                For iEx = 1 To nbExceptions
                    Dim pCode As String, pNom As String, pJours As String, pCoul As String
                    Dim pDateDeb As Variant, pDateFin As Variant
                    
                    pNom = UCase(Trim(CStr(arrExceptions(iEx, 1))))
                    pCode = UCase(Trim(CStr(arrExceptions(iEx, 2))))
                    pJours = UCase(Trim(CStr(arrExceptions(iEx, 3))))
                    pDateDeb = arrExceptions(iEx, 4)
                    pDateFin = arrExceptions(iEx, 5)
                    ' Gestion couleur dynamique
                    If UBound(arrExceptions, 2) >= 6 Then
                        pCoul = UCase(Trim(CStr(arrExceptions(iEx, 6))))
                    End If
                    If pCoul = "" Then pCoul = "JAUNE" ' Securite
                    
                    Dim IsMatch As Boolean
                    IsMatch = False

                    
                    ' ---------------------------------------------------------
                    ' LOGIQUE UNIFIEE (Exceptions + Regles Globales)
                    ' ---------------------------------------------------------
                    If code = pCode Then ' Cas 1: Code Exact (ex: WE)
                        IsMatch = True
                    ElseIf InStr(pCode, "*") > 0 Or InStr(pCode, ",") > 0 Then 
                        ' Cas 2: Code avec Jokers OU Liste (ex: MAL*,MUT* ou CA,RTT)
                        Dim codeArr() As String
                        Dim cItem As Variant
                        
                        ' Si presence de virgule, on split d'abord
                        If InStr(pCode, ",") > 0 Then
                           codeArr = Split(pCode, ",")
                        Else
                           ReDim codeArr(0)
                           codeArr(0) = pCode
                        End If
                        
                        For Each cItem In codeArr
                            Dim cTrim As String
                            cTrim = Trim(cItem)
                            If code Like cTrim Then
                                IsMatch = True
                                Exit For
                            End If
                        Next cItem
                    End If
                    
                    If IsMatch Then
                        ' Verif NOM (Si "*" alors All, sinon Like)
                        If pNom = "*" Or pNom = "" Or nom Like pNom Then
                            
                            ' Verif Dates (Si definies)
                            Dim dateOk As Boolean
                            dateOk = True
                            
                            If IsDate(pDateDeb) Or IsNumeric(pDateDeb) Then
                                If Not IsEmpty(pDateDeb) And pDateDeb <> "" Then
                                    If dateCourante < DCd(pDateDeb) Then dateOk = False
                                End If
                            End If
                            
                            If dateOk And (IsDate(pDateFin) Or IsNumeric(pDateFin)) Then
                                If Not IsEmpty(pDateFin) And pDateFin <> "" Then
                                    If dateCourante > DCd(pDateFin) Then dateOk = False
                                End If
                            End If
                            
                            If dateOk Then
                                ' Verif Jours
                                Dim jArr() As String, jItem As Variant, matchJour As Boolean
                                matchJour = False
                                If pJours = "" Then 
                                    matchJour = True ' Pas de restriction jour
                                Else
                                    jArr = Split(pJours, ",")
                                    For Each jItem In jArr
                                        If jour Like Trim(jItem) & "*" Then
                                            matchJour = True
                                            Exit For
                                        End If
                                    Next jItem
                                End If
                                
                                If matchJour Then
                                    ' APPLICATION COULEUR DYNAMIQUE
                                    Select Case pCoul
                                        Case "JAUNE": If rngJaune Is Nothing Then Set rngJaune = cell Else Set rngJaune = Union(rngJaune, cell)
                                        Case "ORANGE": If rngOrange Is Nothing Then Set rngOrange = cell Else Set rngOrange = Union(rngOrange, cell)
                                        Case "ROUGE": If rngRouge Is Nothing Then Set rngRouge = cell Else Set rngRouge = Union(rngRouge, cell)
                                        Case "BLEU": If rngBleu Is Nothing Then Set rngBleu = cell Else Set rngBleu = Union(rngBleu, cell)
                                        Case "CYAN": If rngCyan Is Nothing Then Set rngCyan = cell Else Set rngCyan = Union(rngCyan, cell)
                                        Case "ROSE": If rngRose Is Nothing Then Set rngRose = cell Else Set rngRose = Union(rngRose, cell)
                                        Case "GRIS": If rngGris Is Nothing Then Set rngGris = cell Else Set rngGris = Union(rngGris, cell)
                                        Case "BLEU_CLAIR": If rngBleuClair Is Nothing Then Set rngBleuClair = cell Else Set rngBleuClair = Union(rngBleuClair, cell)
                                        Case Else: If rngJaune Is Nothing Then Set rngJaune = cell Else Set rngJaune = Union(rngJaune, cell) ' Defaut
                                    End Select
                                    
                                    GoTo NextCol ' Stop processing rules for this cell
                                End If
                            End If
                        End If
                    End If
                Next iEx
            End If
NextCol:
        Next c
    Next r
    
    ' Appliquer les couleurs en UNE SEULE FOIS par couleur
    With rngPlanning
        .Interior.ColorIndex = xlNone
        .Font.ColorIndex = xlAutomatic
        .Font.Bold = False
    End With
    
    If Not rngBleu Is Nothing Then
        rngBleu.Interior.Color = RGB(0, 0, 255)
        rngBleu.Font.Color = vbWhite
        rngBleu.Font.Bold = True
    End If
    
    If Not rngRouge Is Nothing Then
        rngRouge.Interior.Color = RGB(255, 0, 0)
        rngRouge.Font.Color = vbWhite
        rngRouge.Font.Bold = True
    End If
    
    If Not rngJaune Is Nothing Then
        rngJaune.Interior.Color = RGB(255, 255, 0)
        rngJaune.Font.Color = vbBlack
        rngJaune.Font.Bold = True
    End If
    
    If Not rngCyan Is Nothing Then
        rngCyan.Interior.Color = RGB(0, 255, 255)
        rngCyan.Font.Color = vbBlack
        rngCyan.Font.Bold = True
    End If
    
    If Not rngGris Is Nothing Then
        rngGris.Interior.Color = RGB(220, 220, 220)
        rngGris.Font.Color = vbBlack
        rngGris.Font.Bold = True
    End If
    
    If Not rngOrange Is Nothing Then
        rngOrange.Interior.Color = RGB(255, 140, 0) ' Orange
        rngOrange.Font.Color = vbBlack
        rngOrange.Font.Bold = True
    End If
    
    If Not rngRose Is Nothing Then
        rngRose.Interior.Color = RGB(255, 192, 203)
        rngRose.Font.Color = vbBlack
        rngRose.Font.Bold = True
    End If
    
    If Not rngBleuClair Is Nothing Then
        rngBleuClair.Interior.Color = RGB(173, 216, 230) ' Bleu Clair demande (LightBlue)
        rngBleuClair.Font.Color = vbBlack
        rngBleuClair.Font.Bold = True
    End If
    
    ' Retablir
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    MsgBox "Couleurs appliquees avec regles specifiques!", vbInformation
End Sub

' Helper pour conversion date sure
Private Function DCd(v As Variant) As Double
    On Error Resume Next
    DCd = CDbl(CDate(v))
    On Error GoTo 0
End Function
