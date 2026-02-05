Attribute VB_Name = "Module_Config"
Option Explicit

' ============================================================================
' Module de Configuration
' Permet de generer les regles par defaut dans la feuille Config_Exceptions
' ============================================================================

Public Sub InitialiserReglesDefaut()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' 1. Verifier/Creer la feuille
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Config_Exceptions")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Config_Exceptions"
        ws.Range("A1:F1").Value = Array("Nom", "Code", "Jours", "DateDeb", "DateFin", "Couleur")
        ws.Range("A1:F1").Font.Bold = True
        ws.Range("A1:F1").Interior.Color = RGB(220, 220, 220)
    End If
    
    ' 2. Mode Mise a jour intelligente (pas d'ecrasement massif)
    Application.ScreenUpdating = False
    
    ' --- AJOUT SI MANQUANT ---
    
    ' --- BLEU (WE) ---
    VerifEtAjouterRegle ws, "*", "WE", "", "", "", "BLEU"
    
    ' --- ROUGE (Maladie, etc) ---
    VerifEtAjouterRegle ws, "*", "MAL*,MUT*,MAT*,PAT*,F 1-1,R *-*", "", "", "", "ROUGE"
    
    ' --- JAUNE (Conges, etc) ---
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
    
    MsgBox "Verification terminee : Les regles manquantes ont ete ajoutees a Config_Exceptions.", vbInformation
End Sub

' Verifie si une regle avec ce Code existe deja (peu importe le Nom si c'est global, ou simplification)
' Ici on verifie la paire Nom + Code pour eviter les doublons
Private Sub VerifEtAjouterRegle(ws As Worksheet, nom As String, code As String, jours As String, dd As String, df As String, coul As String)
    Dim lastRow As Long, i As Long
    Dim existe As Boolean
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    existe = False
    
    If lastRow >= 2 Then
        For i = 2 To lastRow
            ' On compare Nom et Code (sensible a la casse ou non, ici UCASE pour etre sur)
            If UCase(Trim(ws.Cells(i, 1).Value)) = UCase(Trim(nom)) And _
               UCase(Trim(ws.Cells(i, 2).Value)) = UCase(Trim(code)) Then
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
    r = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    
    ws.Cells(r, 1).Value = nom
    ws.Cells(r, 2).Value = code
    ws.Cells(r, 3).Value = jours
    ws.Cells(r, 4).Value = dd
    ws.Cells(r, 5).Value = df
    ws.Cells(r, 6).Value = coul
End Sub

' ============================================================================
' RESTAURATION DES FONCTIONS UTILITAIRES (CfgText, CfgBool, CfgLong)
' Necéssaires pour d'autres modules (ex: Module_ViewApply)
' Lit la configuration depuis la feuille "Feuil_Config" (Col A=Clé, Col B=Valeur)
' ============================================================================

Public Function CfgText(key As String) As String
    Dim ws As Worksheet
    Dim rng As Range
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Feuil_Config")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Function
    
    ' Recherche de la clé en Colonne A
    On Error Resume Next
    Set rng = ws.Columns("A").Find(What:=key, LookIn:=xlValues, LookAt:=xlWhole)
    On Error GoTo 0
    
    If Not rng Is Nothing Then
        CfgText = CStr(rng.Offset(0, 1).Value)
    End If
End Function

Public Function CfgLong(key As String) As Long
    Dim sVal As String
    sVal = CfgText(key)
    If IsNumeric(sVal) And sVal <> "" Then
        CfgLong = CLng(sVal)
    Else
        CfgLong = 0
    End If
End Function

Public Function CfgBool(key As String) As Boolean
    Dim sVal As String
    sVal = UCase(Trim(CfgText(key)))
    If sVal = "TRUE" Or sVal = "VRAI" Or sVal = "1" Or sVal = "OUI" Then
        CfgBool = True
    Else
        CfgBool = False
    End If
End Function

' ============================================================================
' HELPERS WITH DEFAULTS (Added for compatibility with check_presence)
' ============================================================================

Public Function CfgTextOr(key As String, defaultVal As String) As String
    Dim res As String
    res = CfgText(key)
    If res = "" Then
        CfgTextOr = defaultVal
    Else
        CfgTextOr = res
    End If
End Function

Public Function CfgLongOr(key As String, defaultVal As Long) As Long
    Dim res As String
    res = CfgText(key)
    If IsNumeric(res) And res <> "" Then
        CfgLongOr = CLng(res)
    Else
        CfgLongOr = defaultVal
    End If
End Function



Public Function CfgValueOr(key As String, defaultVal As Variant) As Variant
    Dim s As String
    s = CfgText(key)
    If s = "" Then
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
End Function

