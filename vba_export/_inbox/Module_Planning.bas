' ExportedAt: 2026-01-12 15:37:09 | Workbook: Planning_2026.xlsm
Attribute VB_Name = "Module_Planning"
Option Explicit

'================================================================================================
' MODULE :          Module_Planning (Optimized)
' DESCRIPTION :     Totaux journaliers + diagnostics pour la feuille de planning principale.
'                   S'appuie sur Module_CodeProcessor (GetCodeInfo / clsCodeInfo).
'                   Règle spéciale : ignorer certains codes si fond JAUNE ou BLEU CLAIR
'                   pour Bourgeois_Aurore, Diallo_Mamadou, Dela Vega_Edelyn.
' DATE MODIFIED :   12/10/2025 - Règle couleur (jaune/bleu) Aurore = Mamadou = Edelyn
'                   31/07/2025 - Ajout décompte nuits (lignes 31–38).
'================================================================================================

' --- CONSTANTES POUR LE PLANNING PRINCIPAL (PERSONNEL DE JOUR) ---
Private Const START_ROW As Long = 6
Private Const END_ROW As Long = 26
Private Const START_COL As Long = 3      ' Colonne C
Private Const END_COL As Long = 33       ' Colonne AG

' --- CONSTANTES POUR LES LIGNES DE TOTAUX (PERSONNEL DE JOUR) ---
Private Const TOTAL_ROW_MATIN As Long = 60
Private Const TOTAL_ROW_APRESMIDI As Long = 61
Private Const TOTAL_ROW_SOIR As Long = 62
Private Const PRESENCE_ROW_P06H45 As Long = 64
Private Const PRESENCE_ROW_P07H8H As Long = 65
Private Const PRESENCE_ROW_P8H1630 As Long = 66
Private Const PRESENCE_ROW_C15 As Long = 67
Private Const PRESENCE_ROW_C20 As Long = 68
Private Const PRESENCE_ROW_C20E As Long = 69
Private Const PRESENCE_ROW_C19 As Long = 70

' --- CONSTANTES POUR LE DÉCOMPTE SPÉCIFIQUE DES NUITS ---
Private Const NIGHT_SHIFT_START_ROW As Long = 31 ' Ligne 31
Private Const NIGHT_SHIFT_END_ROW As Long = 38   ' Ligne 38
Private Const NIGHT_CODE_1 As String = "19:45 6:45"
Private Const NIGHT_CODE_2 As String = "20 7"
Private Const PRESENCE_ROW_NIGHT_1 As Long = 71  ' "19:45 6:45"
Private Const PRESENCE_ROW_NIGHT_2 As Long = 72  ' "20 7"
Private Const TOTAL_ROW_NUIT As Long = 73        ' Somme des nuits

' --- Exceptions : ignorer si JAUNE ou BLEU CLAIR (si blanc => compter)
Private ignoreIfYellowOrBlue As Object

'================================================================================================
'   PROCÉDURE PRINCIPALE
'================================================================================================
Public Sub UpdateDailyTotals()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim oldCalc As XlCalculation

    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim nbRows As Long: nbRows = END_ROW - START_ROW + 1
    Dim nbCols As Long: nbCols = END_COL - START_COL + 1

    Dim schedule As Variant
    schedule = ws.Range(ws.Cells(START_ROW, START_COL), ws.Cells(END_ROW, END_COL)).value

    Dim names As Variant
    names = ws.Range(ws.Cells(START_ROW, 1), ws.Cells(END_ROW, 1)).value

    Dim colIndex As Long, rowIndex As Long
    Dim codeHoraire As String, personName As String
    Dim totals(1 To 10) As Double
    Dim cell As Range
    Dim codeInfo As clsCodeInfo

    Debug.Print "--------------------------------------------------"
    Debug.Print "Lancement du Diagnostic... " & Now()
    Debug.Print "--------------------------------------------------"

    For colIndex = 1 To nbCols
        Dim i As Long: For i = 1 To 10: totals(i) = 0: Next i

        For rowIndex = 1 To nbRows
            Set cell = ws.Cells(START_ROW + rowIndex - 1, START_COL + colIndex - 1)
            codeHoraire = Trim(Replace(CStr(schedule(rowIndex, colIndex)), Chr(160), " "))
            personName = CStr(names(rowIndex, 1))

            If codeHoraire <> "" Then
                If Not ShouldBeIgnored(cell, personName, codeHoraire) Then
                    Set codeInfo = GetCodeInfo(codeHoraire)
                    If codeInfo.code <> "INCONNU" Then
                        totals(1) = totals(1) + codeInfo.Fractions(1)
                        totals(2) = totals(2) + codeInfo.Fractions(2)
                        totals(3) = totals(3) + codeInfo.Fractions(3)
                        totals(4) = totals(4) + codeInfo.Fractions(5)
                        totals(5) = totals(5) + codeInfo.Fractions(6)
                        totals(6) = totals(6) + codeInfo.Fractions(7)
                        totals(7) = totals(7) + codeInfo.Fractions(8)
                        totals(8) = totals(8) + codeInfo.Fractions(9)
                        totals(9) = totals(9) + codeInfo.Fractions(10)
                        totals(10) = totals(10) + codeInfo.Fractions(11)
                    End If
                End If
            End If
        Next rowIndex

        ' Écrit les totaux (jour + nuit) pour la colonne en cours
        WriteTotalsToSheet ws, START_COL + colIndex - 1, totals
    Next colIndex

    Application.ScreenUpdating = True
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    MsgBox "Calcul des totaux terminé avec succès !", vbInformation
End Sub

'================================================================================================
'   FONCTIONS DE SUPPORT
'================================================================================================

Private Function ShouldBeIgnored(ByVal cell As Range, ByVal personName As String, ByVal code As String) As Boolean
    ' Règle : pour Aurore / Mamadou / Edelyn, ignorer les codes cibles si fond JAUNE ou BLEU CLAIR.
    If ignoreIfYellowOrBlue Is Nothing Then InitIgnoreDicts

    Dim key As String: key = personName & "|" & code

    If ignoreIfYellowOrBlue.Exists(key) Then
        If IsYellow(cell) Or IsLightBlue(cell) Then
            ShouldBeIgnored = True      ' jaune/bleu clair => on n'additionne pas
            Exit Function
        End If
    End If

    ShouldBeIgnored = False             ' blanc (ou autre) => on compte
End Function

Private Sub InitIgnoreDicts()
    Set ignoreIfYellowOrBlue = CreateObject("Scripting.Dictionary")
    ignoreIfYellowOrBlue.CompareMode = vbTextCompare

    ' --- Codes concernés (toutes couleurs MAIS action seulement si jaune/bleu clair)
    ' Bourgeois_Aurore
    ignoreIfYellowOrBlue("Bourgeois_Aurore|7 15:30") = True
    ignoreIfYellowOrBlue("Bourgeois_Aurore|6:45 15:15") = True
    ' Diallo_Mamadou
    ignoreIfYellowOrBlue("Diallo_Mamadou|7 15:30") = True
    ignoreIfYellowOrBlue("Diallo_Mamadou|6:45 15:15") = True
    ' Dela Vega_Edelyn
    ignoreIfYellowOrBlue("Dela Vega_Edelyn|7 15:30") = True
    ignoreIfYellowOrBlue("Dela Vega_Edelyn|6:45 15:15") = True
End Sub

' === Détection des couleurs ===
Private Function IsYellow(c As Range) As Boolean
    ' Jaune standard Excel
    IsYellow = (c.Interior.Color = vbYellow) Or (c.Interior.ColorIndex = 6)
End Function

Private Function IsLightBlue(c As Range) As Boolean
    ' Détection robuste de “bleu clair” (thème + quelques ColorIndex usuels)
    On Error Resume Next
    Dim themec As Long: themec = c.Interior.ThemeColor
    Dim tint As Double: tint = c.Interior.TintAndShade
    Dim idx As Long: idx = c.Interior.ColorIndex
    Dim rgbv As Long: rgbv = c.Interior.Color

    ' Accent1 éclairci (thème) OU bleus pâles fréquents
    IsLightBlue = (themec = xlThemeColorAccent1 And tint > 0) _
                  Or (idx = 37 Or idx = 34 Or idx = 41) _
                  Or (rgbv = RGB(221, 235, 247) Or rgbv = RGB(204, 232, 255) Or rgbv = RGB(198, 239, 255))
End Function

'================================================================================================
'   ÉCRITURE DES RÉSULTATS (jour + nuit)
'================================================================================================
Private Sub WriteTotalsToSheet(ByVal ws As Worksheet, ByVal col As Long, ByRef totals() As Double)
    ' --- Totaux jour ---
    ws.Cells(TOTAL_ROW_MATIN, col).value = IIf(totals(1) > 0, totals(1), "")
    ws.Cells(TOTAL_ROW_APRESMIDI, col).value = IIf(totals(2) > 0, totals(2), "")
    ws.Cells(TOTAL_ROW_SOIR, col).value = IIf(totals(3) > 0, totals(3), "")
    ws.Cells(PRESENCE_ROW_P06H45, col).value = IIf(totals(4) > 0, totals(4), "")
    ws.Cells(PRESENCE_ROW_P07H8H, col).value = IIf(totals(5) > 0, totals(5), "")
    ws.Cells(PRESENCE_ROW_P8H1630, col).value = IIf(totals(6) > 0, totals(6), "")
    ws.Cells(PRESENCE_ROW_C15, col).value = IIf(totals(7) > 0, totals(7), "")
    ws.Cells(PRESENCE_ROW_C20, col).value = IIf(totals(8) > 0, totals(8), "")
    ws.Cells(PRESENCE_ROW_C20E, col).value = IIf(totals(9) > 0, totals(9), "")
    ws.Cells(PRESENCE_ROW_C19, col).value = IIf(totals(10) > 0, totals(10), "")

    ' --- Totaux nuit (décompte direct sur lignes 31–38) ---
    Dim nightRange As Range
    Set nightRange = ws.Range(ws.Cells(NIGHT_SHIFT_START_ROW, col), ws.Cells(NIGHT_SHIFT_END_ROW, col))

    Dim countNight1 As Double, countNight2 As Double
    countNight1 = Application.WorksheetFunction.CountIf(nightRange, NIGHT_CODE_1)
    countNight2 = Application.WorksheetFunction.CountIf(nightRange, NIGHT_CODE_2)

    Dim totalNight As Double: totalNight = countNight1 + countNight2

    ws.Cells(PRESENCE_ROW_NIGHT_1, col).value = IIf(countNight1 > 0, countNight1, "")
    ws.Cells(PRESENCE_ROW_NIGHT_2, col).value = IIf(countNight2 > 0, countNight2, "")
    ws.Cells(TOTAL_ROW_NUIT, col).value = IIf(totalNight > 0, totalNight, "")
End Sub


