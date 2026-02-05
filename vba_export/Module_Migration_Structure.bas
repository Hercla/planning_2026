Attribute VB_Name = "Module_Migration_Structure"
Option Explicit

' =========================================================================================
' SCRIPT DE MIGRATION STRUCTURELLE - ONE SHOT
' Objectif : Adapter les feuilles "Mois" pour la nouvelle synthese "Low-Flow"
' 1. Insere une ligne "Meteo" en 58
' 2. Dedouble les lignes Matin/AM/Soir pour avoir Total + INF separes
' 3. Met a jour Feuil_Config avec les nouveaux index de lignes
' =========================================================================================

Sub Lancer_Migration_Totale()
    If MsgBox("Voulez-vous lancer la migration de la structure du planning ?" & vbCrLf & _
              "Cela va INSERER des lignes dans tous les onglets mois (Janv-Dec) et mettre a jour Config." & vbCrLf & _
              "A faire UNE SEULE FOIS.", vbQuestion + vbYesNo, "Confirmation Migration") = vbNo Then Exit Sub
              
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False ' CRUCIAL : Evite les boucles infinies sur Worksheet_Change
    
    On Error GoTo CleanExit
    
    MigrerStructureMois
    MettreAJourConfig
    
CleanExit:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    If Err.Number <> 0 Then
        MsgBox "Erreur durant la migration: " & Err.description, vbCritical
    Else
        MsgBox "Migration terminee avec succes !" & vbCrLf & "Les lignes ont ete ajoutees et la config mise a jour.", vbInformation
    End If
End Sub

Private Sub MigrerStructureMois()
    Dim moisArr As Variant
    moisArr = Array("Janv", "Fev", "Mars", "Avril", "Mai", "Juin", "Juil", "Aout", "Sept", "Oct", "Nov", "Dec")
    
    Dim ws As Worksheet
    Dim m As Variant
    
    For Each m In moisArr
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(m)
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            TraiterFeuilleMois ws
        End If
    Next m
End Sub

Private Sub TraiterFeuilleMois(ws As Worksheet)
    ' On suppose que la structure actuelle commence a 60 (Matin)
    ' On veut arriver a :
    ' 58 : Meteo (New)
    ' 59 : Reunion/Formation (Exist)
    ' 60 : Matin Total (Exist)
    ' 61 : Matin INF (New insert)
    ' 62 : AM Total (Old 61)
    ' 63 : AM INF (New insert)
    ' 64 : Soir Total (Old 62)
    ' 65 : Soir INF (New Insert)
    ' ...
    
    ' Detection si deja migre pour eviter double insertion
    If ws.Cells(58, 1).value = "Météo / Status" Then Exit Sub ' Deja fait
    
    ' 1. Inserer Ligne Meteo en 58 (Insere en 58, decale le reste vers le bas)
    ws.Rows(58).Insert Shift:=xlDown
    ws.Cells(58, 1).value = "Météo / Status"
    ws.Cells(58, 1).Font.Bold = True
    ws.Cells(58, 1).Interior.Color = RGB(240, 240, 240)
    
    ' Maintenance apres insertion 58
    ' L'ancienne 59 (Reunion) est devenue 60
    ' L'ancienne 60 (Matin) est devenue 61
    
    ' On veut inserer SOUS Matin (qui est en 61 maintenant)
    ' Donc on insere en 62
    ws.Rows(62).Insert Shift:=xlDown
    ws.Cells(62, 1).value = "   dont Infirmiers"
    ws.Cells(62, 1).Font.Italic = True
    
    ' Maintenant :
    ' 61: Matin Total
    ' 62: Matin INF
    ' 63: AM Total (etait 61 avant decale x2)
    
    ' On veut inserer SOUS AM (63) -> insere en 64
    ws.Rows(64).Insert Shift:=xlDown
    ws.Cells(64, 1).value = "   dont Infirmiers"
    ws.Cells(64, 1).Font.Italic = True
    
    ' Maintenant :
    ' 63: AM Total
    ' 64: AM INF
    ' 65: Soir Total (etait 62 avant decale x3)
    
    ' On veut inserer SOUS Soir (65) -> insere en 66
    ws.Rows(66).Insert Shift:=xlDown
    ws.Cells(66, 1).value = "   dont Infirmiers"
    ws.Cells(66, 1).Font.Italic = True
    
    ' Maintenant :
    ' 65: Soir Total
    ' 66: Soir INF
    ' 67: Nuit Total (etait 63 avant...)
    
    ' On veut inserer SOUS Nuit (67) -> insere en 68 (Optionnel, mais pour consistance)
    ws.Rows(68).Insert Shift:=xlDown
    ws.Cells(68, 1).value = "   dont Infirmiers"
    ws.Cells(68, 1).Font.Italic = True
    
End Sub

Private Sub MettreAJourConfig()
    Dim wsCfg As Worksheet
    Set wsCfg = ThisWorkbook.Sheets("Feuil_Config")
    
    ' Mise a jour des valeurs existantes (decals) et ajout des nouvelles cles
    ' Nouvelle structure:
    ' 58: Meteo
    ' 60: Reunion/Formation (Attention, c'est une etiquette ?) Non, image montre Reunion en rouge ligne 59 (avant). Devenue 60.
    ' 61: Matin Total (Old 60 -> +1 pour weather +0 car insert INF est APRES)
    ' Wait.
    ' Original:
    ' 59: Reunion
    ' 60: Matin
    ' 61: AM
    ' 62: Soir
    
    ' Step 1 (Insert 58 Meteo):
    ' 58: Meteo
    ' 60: Reunion
    ' 61: Matin
    ' 62: AM
    ' 63: Soir
    
    ' Step 2 (Insert 62 Matin Inf):
    ' 61: Matin Total
    ' 62: Matin INF
    ' 63: AM Total
    ' 64: Soir
    
    ' Step 3 (Insert 64 AM Inf):
    ' 63: AM Total
    ' 64: AM INF
    ' 65: Soir Total
    
    ' Step 4 (Insert 66 Soir Inf):
    ' 65: Soir Total
    ' 66: Soir INF
    ' 67: Nuit Total
    
    ' Step 5 (Insert 68 Nuit Inf):
    ' 67: Nuit Total
    ' 68: Nuit INF
    
    ' Les suivantes (P_0645...) sont decalees de 5 lignes au total (1 meteo + 4 INFs)
    ' Old 64 (P_0645) -> New 69 ?
    ' Check:
    ' 68: Nuit INF
    ' 69: P_0645 (Old 64 -> +5 = 69). Correct.
    
    ' Ecriture Config
    UpdateOrAddConfig wsCfg, "CALC_ROW_Meteo", 58
    UpdateOrAddConfig wsCfg, "CALC_ROW_Matin", 61
    UpdateOrAddConfig wsCfg, "CALC_ROW_Matin_INF", 62
    UpdateOrAddConfig wsCfg, "CALC_ROW_AM", 63
    UpdateOrAddConfig wsCfg, "CALC_ROW_AM_INF", 64
    UpdateOrAddConfig wsCfg, "CALC_ROW_Soir", 65
    UpdateOrAddConfig wsCfg, "CALC_ROW_Soir_INF", 66
    UpdateOrAddConfig wsCfg, "CALC_ROW_Nuit", 67
    UpdateOrAddConfig wsCfg, "CALC_ROW_Nuit_INF", 68
    
    ' Decalage des suivants (+5 par rapport a l'original)
    UpdateOrAddConfig wsCfg, "CALC_ROW_P_0645", 69
    UpdateOrAddConfig wsCfg, "CALC_ROW_P_7H8H", 70
    UpdateOrAddConfig wsCfg, "CALC_ROW_P_8H1630", 71
    UpdateOrAddConfig wsCfg, "CALC_ROW_C15", 72
    UpdateOrAddConfig wsCfg, "CALC_ROW_C20", 73
    UpdateOrAddConfig wsCfg, "CALC_ROW_C20E", 74
    UpdateOrAddConfig wsCfg, "CALC_ROW_C19", 75
    
End Sub

Private Sub UpdateOrAddConfig(ws As Worksheet, key As String, val As Long)
    Dim rng As Range
    Set rng = ws.Columns("A").Find(key, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not rng Is Nothing Then
        rng.offset(0, 1).value = val
    Else
        Dim lr As Long
        lr = ws.Cells(ws.Rows.count, "A").End(xlUp).row + 1
        ws.Cells(lr, 1).value = key
        ws.Cells(lr, 2).value = val
        ws.Cells(lr, 3).value = "Auto-Migrated"
    End If
End Sub
