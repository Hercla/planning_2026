' ExportedAt: 2026-01-12 15:37:09 | Workbook: Planning_2026.xlsm
Attribute VB_Name = "Module3calculfinal"
Option Explicit

' =========================================================================================
'   MACRO DE CALCUL DES TOTAUX - VERSION FINALE AVEC RÈGLE D'EXCEPTION CHIRURGICALE
'
'   Règle métier implémentée :
'   - Ignore "7 15:30" ou "6:45 15:15" SI ET SEULEMENT SI le fond est jaune
'     ET l'employé est "Bourgeois_Aurore" ou "Diallo_Mamadou".
' =========================================================================================

Sub Calculer_Totaux_Avec_Exception_Specifique()
    Dim ws As Worksheet, wsListe As Worksheet
    Dim dictCodes As Object
    Dim lastRowListe As Long, listeData As Variant
    Dim r As Long, code As String
    Dim col As Long, i As Long, j As Long
    Dim cell As Range, codeHoraire As String, personName As String
    Dim totals(1 To 10) As Double
    Dim storedValues As Variant
    Dim ignorerCellule As Boolean

    ' --- INITIALISATION ET VÉRIFICATIONS ---
    Set ws = ActiveSheet
    On Error Resume Next
    Set wsListe = ThisWorkbook.Sheets("Liste")
    On Error GoTo 0
    If wsListe Is Nothing Then
        MsgBox "ERREUR : La feuille ""Liste"" est introuvable.", vbCritical
        Exit Sub
    End If

    ' --- PARAMÈTRES ---
    Const PREMIERE_LIGNE_PERSONNEL As Long = 6, DERNIERE_LIGNE_PERSONNEL As Long = 26
    Const PREMIERE_COLONNE_JOUR As Long = 2, DERNIERE_COLONNE_JOUR As Long = 32
    Const COULEUR_JAUNE As Long = 65535 ' La valeur RGB pour le jaune pur est 255, 255, 0
    
    Const LIGNE_TOTAL_MATIN As Long = 60, LIGNE_TOTAL_APRESMIDI As Long = 61, LIGNE_TOTAL_SOIR As Long = 62
    Const LIGNE_PRESENCE_06H45 As Long = 64, LIGNE_PRESENCE_7H_8H As Long = 65, LIGNE_PRESENCE_8H_16H30 As Long = 66
    Const LIGNE_PRESENCE_C15 As Long = 67, LIGNE_PRESENCE_C20 As Long = 68, LIGNE_PRESENCE_C20E As Long = 69, LIGNE_PRESENCE_C19 As Long = 70

    Application.ScreenUpdating = False

    ' --- ÉTAPE 1: Charger la "Liste" en mémoire avec nettoyage ---
    Set dictCodes = CreateObject("Scripting.Dictionary")
    dictCodes.CompareMode = vbTextCompare
    lastRowListe = wsListe.Cells(wsListe.Rows.Count, "A").End(xlUp).row
    If lastRowListe < 2 Then Exit Sub
    listeData = wsListe.Range("A2:O" & lastRowListe).value
    For r = 1 To UBound(listeData, 1)
        code = Trim(Replace(CStr(listeData(r, 1)), Chr(160), " "))
        If code <> "" And Not dictCodes.Exists(code) Then
            dictCodes.Add code, Application.index(listeData, r, 0)
        End If
    Next r

    ' --- ÉTAPE 2: PARCOURIR LE PLANNING ET CALCULER ---
    For col = PREMIERE_COLONNE_JOUR To DERNIERE_COLONNE_JOUR
        For j = 1 To 10: totals(j) = 0: Next j
        For i = PREMIERE_LIGNE_PERSONNEL To DERNIERE_LIGNE_PERSONNEL
            Set cell = ws.Cells(i, col)
            codeHoraire = Trim(Replace(cell.text, Chr(160), " "))
            personName = ws.Cells(i, 1).text ' On récupère le nom de l'employé dans la colonne A
            
            ignorerCellule = False ' Par défaut, on n'ignore pas la cellule

            ' --- RÈGLE D'EXCEPTION CHIRURGICALE ---
            If cell.Interior.Color = COULEUR_JAUNE Then
                ' La cellule est jaune. On vérifie si c'est un de nos cas spéciaux.
                Select Case personName
                    Case "Bourgeois_Aurore", "Diallo_Mamadou"
                        ' C'est le bon employé. On vérifie si c'est le bon code.
                        Select Case codeHoraire
                            Case "7 15:30", "6:45 15:15"
                                ' C'est le bon employé ET le bon code ET la bonne couleur. On ignore.
                                ignorerCellule = True
                        End Select
                End Select
            End If
            
            ' On ne calcule que si la cellule ne doit pas être ignorée
            If Not ignorerCellule Then
                If dictCodes.Exists(codeHoraire) Then
                    storedValues = dictCodes(codeHoraire)
                    totals(1) = totals(1) + CDbl(IIf(IsNumeric(storedValues(3)), storedValues(3), 0))
                    totals(2) = totals(2) + CDbl(IIf(IsNumeric(storedValues(4)), storedValues(4), 0))
                    totals(3) = totals(3) + CDbl(IIf(IsNumeric(storedValues(5)), storedValues(5), 0))
                    totals(4) = totals(4) + CDbl(IIf(IsNumeric(storedValues(7)), storedValues(7), 0))
                    totals(5) = totals(5) + CDbl(IIf(IsNumeric(storedValues(8)), storedValues(8), 0))
                    totals(6) = totals(6) + CDbl(IIf(IsNumeric(storedValues(9)), storedValues(9), 0))
                    totals(7) = totals(7) + CDbl(IIf(IsNumeric(storedValues(10)), storedValues(10), 0))
                    totals(8) = totals(8) + CDbl(IIf(IsNumeric(storedValues(11)), storedValues(11), 0))
                    totals(9) = totals(9) + CDbl(IIf(IsNumeric(storedValues(12)), storedValues(12), 0))
                    totals(10) = totals(10) + CDbl(IIf(IsNumeric(storedValues(13)), storedValues(13), 0))
                End If
            End If
        Next i

        ' --- ÉTAPE 3: ÉCRIRE LES TOTAUX ---
        ws.Cells(LIGNE_TOTAL_MATIN, col).value = IIf(totals(1) > 0, totals(1), "")
        ws.Cells(LIGNE_TOTAL_APRESMIDI, col).value = IIf(totals(2) > 0, totals(2), "")
        ws.Cells(LIGNE_TOTAL_SOIR, col).value = IIf(totals(3) > 0, totals(3), "")
        ws.Cells(LIGNE_PRESENCE_06H45, col).value = IIf(totals(4) > 0, totals(4), "")
        ws.Cells(LIGNE_PRESENCE_7H_8H, col).value = IIf(totals(5) > 0, totals(5), "")
        ws.Cells(LIGNE_PRESENCE_8H_16H30, col).value = IIf(totals(6) > 0, totals(6), "")
        ws.Cells(LIGNE_PRESENCE_C15, col).value = IIf(totals(7) > 0, totals(7), "")
        ws.Cells(LIGNE_PRESENCE_C20, col).value = IIf(totals(8) > 0, totals(8), "")
        ws.Cells(LIGNE_PRESENCE_C20E, col).value = IIf(totals(9) > 0, totals(9), "")
        ws.Cells(LIGNE_PRESENCE_C19, col).value = IIf(totals(10) > 0, totals(10), "")
    Next col

    Application.ScreenUpdating = True
    MsgBox "Calcul des totaux terminé avec succès !", vbInformation
End Sub
