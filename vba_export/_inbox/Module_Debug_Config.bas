Attribute VB_Name = "Module_Debug_Config"
Option Explicit

Sub Debug_Afficher_Config()
    ' Affiche toutes les clés lues depuis Feuil_Config
    Dim wsConfig As Worksheet
    Dim d As Object
    Dim k As Variant
    Dim msg As String
    Dim count As Long
    
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("Feuil_Config")
    On Error GoTo 0
    
    If wsConfig Is Nothing Then
        MsgBox "Feuille 'Feuil_Config' introuvable!", vbCritical
        Exit Sub
    End If
    
    ' Charger la config
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    
    Dim lr As Long, arr As Variant, i As Long
    Dim cle As String, valeur As Variant
    lr = wsConfig.Cells(wsConfig.Rows.Count, "A").End(xlUp).Row
    
    msg = "=== FEUIL_CONFIG ===" & vbCrLf
    msg = msg & "Lignes lues: 2 à " & lr & vbCrLf & vbCrLf
    
    If lr >= 2 Then
        arr = wsConfig.Range("A2:B" & lr).Value
        For i = 1 To UBound(arr, 1)
            cle = Trim(CStr(arr(i, 1)))
            valeur = arr(i, 2)
            
            If cle <> "" Then
                d(cle) = valeur
                count = count + 1
                If count <= 30 Then
                    msg = msg & cle & " = " & CStr(valeur) & vbCrLf
                End If
            End If
        Next i
    End If
    
    msg = msg & vbCrLf & "Total: " & count & " clés trouvées"
    
    ' Vérifier les clés importantes
    msg = msg & vbCrLf & vbCrLf & "=== VÉRIFICATION CLÉS CRITIQUES ==="
    
    ' Lignes destination
    msg = msg & vbCrLf & "CALC_ROW_Matin: " & IIf(d.Exists("CALC_ROW_Matin"), d("CALC_ROW_Matin"), "[MANQUANT]")
    msg = msg & vbCrLf & "CALC_ROW_PM: " & IIf(d.Exists("CALC_ROW_PM"), d("CALC_ROW_PM"), "[MANQUANT]")
    msg = msg & vbCrLf & "CALC_ROW_Soir: " & IIf(d.Exists("CALC_ROW_Soir"), d("CALC_ROW_Soir"), "[MANQUANT]")
    msg = msg & vbCrLf & "CALC_ROW_Nuit: " & IIf(d.Exists("CALC_ROW_Nuit"), d("CALC_ROW_Nuit"), "[MANQUANT]")
    
    ' Paramètres généraux
    msg = msg & vbCrLf & vbCrLf & "ligneDebut: " & IIf(d.Exists("ligneDebut"), d("ligneDebut"), "[MANQUANT]")
    msg = msg & vbCrLf & "ligneFin: " & IIf(d.Exists("ligneFin"), d("ligneFin"), "[MANQUANT]")
    msg = msg & vbCrLf & "colDebut: " & IIf(d.Exists("colDebut"), d("colDebut"), "[MANQUANT]")
    msg = msg & vbCrLf & "colFin: " & IIf(d.Exists("colFin"), d("colFin"), "[MANQUANT]")
    
    ' Effectifs
    msg = msg & vbCrLf & vbCrLf & "EFF_SEM_Matin: " & IIf(d.Exists("EFF_SEM_Matin"), d("EFF_SEM_Matin"), "[MANQUANT]")
    msg = msg & vbCrLf & "seuilMinINF: " & IIf(d.Exists("seuilMinINF"), d("seuilMinINF"), "[MANQUANT]")
    msg = msg & vbCrLf & "ALERT_SEUIL_MIN_INF: " & IIf(d.Exists("ALERT_SEUIL_MIN_INF"), d("ALERT_SEUIL_MIN_INF"), "[MANQUANT]")
    
    MsgBox msg, vbInformation, "Debug Config"
End Sub
