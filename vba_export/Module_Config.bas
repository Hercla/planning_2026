Attribute VB_Name = "Module_Config"
Option Explicit

' ============================================================================
' Module de Configuration (MAJ)
' - Garde la génération des règles par défaut dans Config_Exceptions
' - Remplace l'ancien lecteur Feuil_Config (CfgText/CfgLong/...) par des wrappers
'   qui utilisent Module_ConfigEngine (CFG_*)
' ============================================================================

' =====================================================================
' PARTIE 1 — Config_Exceptions (inchangé)
' =====================================================================

Public Sub InitialiserReglesDefaut()
    Dim ws As Worksheet

    ' 1. Vérifier/Créer la feuille
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Config_Exceptions")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Config_Exceptions"
        ws.Range("A1:F1").value = Array("Nom", "Code", "Jours", "DateDeb", "DateFin", "Couleur")
        ws.Range("A1:F1").Font.Bold = True
        ws.Range("A1:F1").Interior.Color = RGB(220, 220, 220)
    End If

    ' 2. Mode Mise à jour intelligente (pas d'écrasement massif)
    Application.ScreenUpdating = False

    ' --- AJOUT SI MANQUANT ---

    ' --- BLEU (WE) ---
    VerifEtAjouterRegle ws, "*", "WE", "", "", "", "BLEU"

    ' --- ROUGE (Maladie, etc) ---
    VerifEtAjouterRegle ws, "*", "MAL*,MUT*,MAT*,PAT*,F 1-1,R *-*", "", "", "", "ROUGE"

    ' --- JAUNE (Congés, etc) ---
    VerifEtAjouterRegle ws, "*", "CA,RCT,RV,RHS,ANC,EL,C SOC,CRP*,*/*", "", "", "", "JAUNE"

    ' --- ORANGE (CTR) ---
    VerifEtAjouterRegle ws, "*", "CTR", "", "", "", "ORANGE"

    ' --- CYAN (DP) ---
    VerifEtAjouterRegle ws, "*", "DP", "", "", "", "CYAN"

    ' --- GRIS (Absences diverses) ---
    VerifEtAjouterRegle ws, "*", "CSS,PREAVIS,VJ,DECES,PETIT CHOM", "", "", "", "GRIS"

    ' --- ROSE (ASBD) ---
    VerifEtAjouterRegle ws, "*", "ASBD", "", "", "", "ROSE"

    ' Ajustement Colonnes
    ws.Columns("A:F").AutoFit
    Application.ScreenUpdating = True

    MsgBox "Vérification terminée : Les règles manquantes ont été ajoutées à Config_Exceptions.", vbInformation
End Sub

' Vérifie si une règle avec ce Code existe déjà (paire Nom + Code)
Private Sub VerifEtAjouterRegle(ws As Worksheet, nom As String, code As String, jours As String, dd As String, df As String, coul As String)
    Dim lastRow As Long, i As Long
    Dim existe As Boolean

    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    existe = False

    If lastRow >= 2 Then
        For i = 2 To lastRow
            ' On compare Nom et Code
            If UCase$(Trim$(ws.Cells(i, 1).value)) = UCase$(Trim$(nom)) And _
               UCase$(Trim$(ws.Cells(i, 2).value)) = UCase$(Trim$(code)) Then
                existe = True
                Exit For
            End If
        Next i
    End If

    If Not existe Then
        AjouterRegle ws, nom, code, jours, dd, df, coul
    End If
End Sub

Private Sub AjouterRegle(ws As Worksheet, nom As String, code As String, jours As String, dd As String, df As String, coul As String)
    Dim r As Long
    r = ws.Cells(ws.Rows.count, "A").End(xlUp).row + 1

    ws.Cells(r, 1).value = nom
    ws.Cells(r, 2).value = code
    ws.Cells(r, 3).value = jours
    ws.Cells(r, 4).value = dd
    ws.Cells(r, 5).value = df
    ws.Cells(r, 6).value = coul
End Sub

' =====================================================================
' PARTIE 2 — WRAPPERS vers Module_ConfigEngine (CFG_*)
' IMPORTANT: Module_ConfigEngine doit contenir CFG_Str/CFG_Long/CFG_Bool...
' =====================================================================

Public Function CfgText(ByVal key As String) As String
    ' Compat: renvoie "" si clé absente / erreur
    On Error GoTo Safe
    CfgText = CFG_Str(key)
    Exit Function
Safe:
    CfgText = vbNullString
End Function

Public Function CfgLong(ByVal key As String) As Long
    ' Compat: renvoie 0 si clé absente / erreur
    On Error GoTo Safe
    CfgLong = CFG_Long(key)
    Exit Function
Safe:
    CfgLong = 0
End Function

Public Function CfgBool(ByVal key As String) As Boolean
    ' Compat: renvoie False si clé absente / erreur
    On Error GoTo Safe
    CfgBool = CFG_Bool(key)
    Exit Function
Safe:
    CfgBool = False
End Function

Public Function CfgTextOr(ByVal key As String, ByVal defaultVal As String) As String
    On Error GoTo Safe
    Dim res As String
    res = CFG_Str(key)
    If Len(res) = 0 Then res = defaultVal
    CfgTextOr = res
    Exit Function
Safe:
    CfgTextOr = defaultVal
End Function

Public Function CfgLongOr(ByVal key As String, ByVal defaultVal As Long) As Long
    On Error GoTo Safe
    CfgLongOr = CFG_Long(key)
    Exit Function
Safe:
    CfgLongOr = defaultVal
End Function

Public Function CfgValueOr(ByVal key As String, ByVal defaultVal As Variant) As Variant
    On Error GoTo Safe

    Dim s As String
    s = CFG_Str(key)

    If Len(s) = 0 Then
        CfgValueOr = defaultVal
        Exit Function
    End If

    Select Case VarType(defaultVal)
        Case vbBoolean
            Dim u As String
            u = UCase$(Trim$(s))
            CfgValueOr = (u = "TRUE" Or u = "VRAI" Or u = "1" Or u = "OUI" Or u = "YES")

        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency
            If IsNumeric(s) Then
                CfgValueOr = CDbl(s)
            Else
                CfgValueOr = defaultVal
            End If

        Case Else
            CfgValueOr = s
    End Select

    Exit Function

Safe:
    CfgValueOr = defaultVal
End Function


