Attribute VB_Name = "Module_Fix_Style"
Option Explicit

Sub Finaliser_Migration_Style()
    Dim ws As Worksheet
    Dim arrMois As Variant
    Dim m As Variant
    
    arrMois = Array("Janv", "Fev", "Mars", "Avril", "Mai", "Juin", "Juil", "Aout", "Sept", "Oct", "Nov", "Dec")
    
    Application.ScreenUpdating = False
    
    For Each m In arrMois
        On Error Resume Next
        Set ws = Sheets(m)
        If Not ws Is Nothing Then
            ' 1. Corriger l'encodage "MÃ©tÃ©o"
            If Left(ws.Cells(58, 1).value, 2) = "MÃ" Or Left(ws.Cells(58, 1).value, 5) = "Météo" Then
                ws.Cells(58, 1).value = "Météo / Status"
            End If
            
            ' 2. Corriger le fond rouge hérité (surtout ligne 68 sous FRACTIONS HORAIRES)
            Dim r As Variant
            For Each r In Array(62, 64, 66, 68)
                If ws.Cells(r, 1).value Like "*Infirmiers*" Then
                    ws.Cells(r, 1).Interior.ColorIndex = xlNone
                    ws.Cells(r, 1).Font.Italic = True
                End If
            Next r
        End If
        On Error GoTo 0
    Next m
    
    Application.ScreenUpdating = True
    
    ' 3. Verifier et ajouter les cles de config manquantes pour la nouvelle structure
    Verif_Config_Keys
    
    ' 4. Lancer le calcul sur la feuille active pour voir le resultat immediat
    MsgBox "Corrections visuelles appliquées." & vbCrLf & "Lancement du calcul pour remplir les nouvelles cases...", vbInformation
    Application.Run "Module_Calculer_Totaux.Calculer_Totaux_Planning"
End Sub

Private Sub Verif_Config_Keys()
    ' --- MASTER CONFIG CHECKER ---
    ' Verifie config cles ET structure Config_Codes
    
    Dim wsCfg As Worksheet
    On Error Resume Next
    Set wsCfg = ThisWorkbook.Sheets("Feuil_Config")
    On Error GoTo 0
    If wsCfg Is Nothing Then
        Set wsCfg = ThisWorkbook.Sheets.Add
        wsCfg.Name = "Feuil_Config"
    End If
    
    Dim cles As Object
    Set cles = CreateObject("Scripting.Dictionary")
    
    ' 1. ROWS PLANING (Alignement Screenshot)
    cles("CALC_ROW_Meteo") = 58
    cles("CALC_ROW_Matin") = 60
    cles("CALC_ROW_Matin_INF") = 0
    cles("CALC_ROW_AM") = 61
    cles("CALC_ROW_AM_INF") = 0
    cles("CALC_ROW_Soir") = 62
    cles("CALC_ROW_Soir_INF") = 0
    cles("CALC_ROW_Nuit") = 0
    cles("CALC_ROW_Nuit_INF") = 0
    cles("CALC_ROW_Dates") = 63
    
    cles("CALC_ROW_P_0645") = 64
    cles("CALC_ROW_P_7H8H") = 65
    cles("CALC_ROW_P_8H1630") = 66
    cles("CALC_ROW_C15") = 67
    cles("CALC_ROW_C20") = 68
    cles("CALC_ROW_C20E") = 69
    cles("CALC_ROW_C19") = 70
    
    ' 2. GENERAL PARAMS
    cles("CHK_FirstPersonnelRow") = 6
    cles("ligneFin") = 28
    cles("PLN_FirstDayCol") = 3
    cles("PLN_LastDayCol") = 33
    cles("PLN_Row_DayNumbers") = 4
    cles("CFG_Year") = Year(Date)
    
    ' 3. CODES & COULEURS
    cles("CodesInfirmiere") = "INF;IDE;IC"
    cles("CHK_InfFunctions") = "INF,AS,CEFA"
    cles("CHK_IgnoreColor") = 15849925
    cles("COULEUR_INF_ADMIN") = 65535
    cles("COULEUR_BLEU_CLAIR") = 15128749
    cles("ALERT_SEUIL_MIN_INF") = 2
    
    ' 4. EFFECTIFS (Full Default)
    cles("EFF_SEM_Matin") = 7
    cles("EFF_SEM_PM") = 3
    cles("EFF_SEM_Soir") = 3
    cles("EFF_SEM_Nuit") = 2
    cles("EFF_WE_Matin") = 5
    cles("EFF_WE_PM") = 2
    cles("EFF_WE_Soir") = 3
    cles("EFF_WE_Nuit") = 2
    cles("EFF_FER_Matin") = 5
    cles("EFF_FER_PM") = 2
    cles("EFF_FER_Soir") = 3
    cles("EFF_FER_Nuit") = 2

    Dim k As Variant
    Dim lr As Long
    Dim found As Range
    
    ' VERIF ET AJOUT CLES
    For Each k In cles.keys
        Set found = wsCfg.Columns("A").Find(What:=k, LookIn:=xlValues, LookAt:=xlWhole)
        If found Is Nothing Then
            lr = wsCfg.Cells(wsCfg.Rows.count, "A").End(xlUp).row + 1
            If lr < 2 Then lr = 2
            wsCfg.Cells(lr, 1).value = k
            wsCfg.Cells(lr, 2).value = cles(k)
        Else
            ' FORCE UPDATE pour Row Mapping critique uniquement
            If InStr(k, "CALC_ROW") > 0 Then
                wsCfg.Cells(found.row, 2).value = cles(k)
            End If
        End If
    Next k
    
    ' VERIF STRUCTURE Config_Codes
    Dim wsCodes As Worksheet
    On Error Resume Next
    Set wsCodes = ThisWorkbook.Sheets("Config_Codes")
    On Error GoTo 0
    If wsCodes Is Nothing Then
        Set wsCodes = ThisWorkbook.Sheets.Add
        wsCodes.Name = "Config_Codes"
    End If
    
    ' Headers attendus (A-O)
    Dim headers As Variant
    headers = Array("Code", "Description", "Type_Code", "Heures_normales", "TopCode", _
                    "H_Start", "H_Pause_Start", "H_Pause_End", "H_End", _
                    "F_6h45", "F_7h_8h", "Matin", "PM", "Soir", "Nuit")
    
    wsCodes.Range("A1:O1").value = headers
    wsCodes.Range("A1:O1").Font.Bold = True
    wsCodes.Range("A1:O1").Interior.Color = RGB(200, 220, 240)
    wsCodes.Columns("A:O").AutoFit
End Sub


