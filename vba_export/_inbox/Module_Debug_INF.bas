' Module de DEBUG pour vérifier le matching des noms
' Importer dans Excel et exécuter Debug_Matching()

Sub Debug_Matching()
    ' Créer une feuille de debug avec les correspondances
    Dim wsDebug As Worksheet
    Dim wsPlanning As Worksheet
    Dim dictFonctions As Object
    
    Set wsPlanning = ActiveSheet
    Set dictFonctions = ChargerFonctionsDebug()
    
    ' Créer ou vider la feuille debug
    On Error Resume Next
    Set wsDebug = ThisWorkbook.Sheets("DEBUG_INF")
    If wsDebug Is Nothing Then
        Set wsDebug = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDebug.Name = "DEBUG_INF"
    Else
        wsDebug.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Entêtes
    wsDebug.Cells(1, 1).Value = "Nom Planning"
    wsDebug.Cells(1, 2).Value = "Trouve?"
    wsDebug.Cells(1, 3).Value = "Fonction"
    wsDebug.Cells(1, 4).Value = "Est INF?"
    
    ' Parcourir les noms du planning (colonne A, lignes 5-30)
    Dim r As Long, nomPlanning As String, rowDebug As Long
    rowDebug = 2
    
    For r = 5 To 30
        nomPlanning = Trim(CStr(wsPlanning.Cells(r, 1).Value))
        If nomPlanning <> "" Then
            wsDebug.Cells(rowDebug, 1).Value = nomPlanning
            
            If dictFonctions.Exists(nomPlanning) Then
                wsDebug.Cells(rowDebug, 2).Value = "OUI"
                wsDebug.Cells(rowDebug, 3).Value = dictFonctions(nomPlanning)
                If UCase(dictFonctions(nomPlanning)) = "INF" Then
                    wsDebug.Cells(rowDebug, 4).Value = "OUI"
                    wsDebug.Cells(rowDebug, 4).Interior.Color = RGB(144, 238, 144)
                Else
                    wsDebug.Cells(rowDebug, 4).Value = "NON"
                    wsDebug.Cells(rowDebug, 4).Interior.Color = RGB(255, 200, 100)
                End If
            Else
                wsDebug.Cells(rowDebug, 2).Value = "NON"
                wsDebug.Cells(rowDebug, 2).Interior.Color = RGB(255, 100, 100)
                wsDebug.Cells(rowDebug, 3).Value = "?"
                wsDebug.Cells(rowDebug, 4).Value = "?"
            End If
            rowDebug = rowDebug + 1
        End If
    Next r
    
    ' Stats
    wsDebug.Cells(rowDebug + 1, 1).Value = "STATS:"
    wsDebug.Cells(rowDebug + 2, 1).Value = "Total dans dico:"
    wsDebug.Cells(rowDebug + 2, 2).Value = dictFonctions.Count
    
    Dim nbINF As Long, cle As Variant
    For Each cle In dictFonctions.Keys
        If UCase(dictFonctions(cle)) = "INF" Then nbINF = nbINF + 1
    Next cle
    wsDebug.Cells(rowDebug + 3, 1).Value = "INF dans dico:"
    wsDebug.Cells(rowDebug + 3, 2).Value = nbINF
    
    ' Auto-ajuster colonnes
    wsDebug.Columns("A:D").AutoFit
    
    wsDebug.Activate
    MsgBox "Debug termine! Voir feuille DEBUG_INF", vbInformation
End Sub

Private Function ChargerFonctionsDebug() As Object
    ' Copie de la fonction du module principal
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    
    Dim wsPersonnel As Worksheet
    On Error Resume Next
    Set wsPersonnel = ThisWorkbook.Sheets("Personnel")
    On Error GoTo 0
    
    If wsPersonnel Is Nothing Then
        MsgBox "Feuille Personnel introuvable!", vbCritical
        Set ChargerFonctionsDebug = d
        Exit Function
    End If
    
    Dim lr As Long, arr As Variant, i As Long
    Dim nom As String, prenom As String, cleNomPrenom As String, fonction As String
    
    lr = wsPersonnel.Cells(wsPersonnel.Rows.Count, "B").End(xlUp).Row
    If lr < 2 Then Set ChargerFonctionsDebug = d: Exit Function
    
    arr = wsPersonnel.Range("B2:E" & lr).Value
    For i = 1 To UBound(arr, 1)
        nom = Trim(CStr(arr(i, 1)))      ' Colonne B = Nom
        prenom = Trim(CStr(arr(i, 2)))   ' Colonne C = Prénom
        fonction = Trim(CStr(arr(i, 4))) ' Colonne E = Fonction
        
        ' Créer la clé au format "Nom_Prénom" comme dans le planning
        cleNomPrenom = nom & "_" & prenom
        
        If cleNomPrenom <> "_" And Not d.Exists(cleNomPrenom) Then
            d.Add cleNomPrenom, fonction
        End If
    Next i
    
    Set ChargerFonctionsDebug = d
End Function
