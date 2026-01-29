Attribute VB_Name = "Module_Interface"
Option Explicit

' ========================================================================================
' MODULE INTERFACE : Gestion du formulaire de saisie des exceptions
' Permet d'ajouter des regles dans Config_Exceptions sans manipuler le tableau
' ========================================================================================

Public Sub OuvrirInterfaceSaisie()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Interface_Exceptions")
    On Error GoTo 0
    
    If ws Is Nothing Then
        InitInterfaceSaisie
    Else
        ' Verification si mise a jour necessaire (Si B13 n'est pas "5. Couleur")
        If ws.Range("B13").value <> "5. Couleur" Then
            Dim resp As VbMsgBoxResult
            resp = MsgBox("Mise a jour de l'interface requise pour les couleurs. Reinitialiser ?", vbYesNo + vbQuestion)
            If resp = vbYes Then
                InitInterfaceSaisie
            Else
                ws.Activate
                MettreAJourListeNoms
            End If
        Else
            ws.Activate
            MettreAJourListeNoms
            MsgBox "Interface prete !", vbInformation
        End If
    End If
End Sub

' Initialise la feuille ex nihilo (Layout, contoles, boutons)
Public Sub InitInterfaceSaisie()
    Dim ws As Worksheet
    Dim btn As Button
    Dim chk As CheckBox
    Dim i As Integer
    Dim jours() As Variant
    jours = Array("LUN", "MAR", "MER", "JEU", "VEN", "SAM", "DIM")
    
    Application.ScreenUpdating = False
    
    ' 1. Creer ou Reset Feuille
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Interface_Exceptions").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    ws.Name = "Interface_Exceptions"
    ActiveWindow.DisplayGridlines = False
    
    ' 2. Styles et Titres
    With ws
        .Range("B2").value = "SAISIE EXCEPTION PLANNING"
        .Range("B2").Font.Size = 16
        .Range("B2").Font.Bold = True
        .Range("B2:H2").Merge
        .Range("B2:H2").HorizontalAlignment = xlCenter
        .Range("B2:H2").Interior.Color = RGB(50, 50, 50)
        .Range("B2:H2").Font.Color = vbWhite
        
        ' --- CHAMP NOM ---
        .Range("B5").value = "1. Qui ?"
        .Range("B5").Font.Bold = True
        .Range("C5:F5").Merge
        .Range("C5").Interior.Color = RGB(240, 240, 240)
        .Range("C5").Borders.LineStyle = xlContinuous
        
        ' Validation de donnees pour les Noms (Liste dynamique)
        MettreAJourListeNoms ' Cree la plage "ListeNomsSource"
        With .Range("C5").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=ListeNomsSource"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        
        ' --- CHAMP CODE ---
        .Range("B7").value = "2. Quel Code ?"
        .Range("B7").Font.Bold = True
        .Range("C7:F7").Merge
        .Range("C7").Interior.Color = RGB(240, 240, 240)
        .Range("C7").Borders.LineStyle = xlContinuous
        
        ' Validation Code (Source existante ListeCodes)
        On Error Resume Next
        With .Range("C7").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=ListeCodes"
        End With
        On Error GoTo 0
        
        ' --- CHAMP JOURS ---
        .Range("B9").value = "3. Quels Jours ?"
        .Range("B9").Font.Bold = True
        
        ' Creer Checkboxes
        Dim leftPos As Double
        leftPos = .Range("C9").Left
        For i = 0 To 6
            Set chk = .CheckBoxes.Add(leftPos, .Range("C9").Top, 40, 15)
            chk.Caption = jours(i)
            chk.Name = "chk_" & jours(i)
            leftPos = leftPos + 50
        Next i
        
        ' --- CHAMP DATES ---
        .Range("B11").value = "4. Dates (Optionnel)"
        .Range("B11").Font.Bold = True
        
        .Range("C11").value = "Debut:"
        .Range("D11").Interior.Color = RGB(240, 240, 240)
        .Range("D11").Borders.LineStyle = xlContinuous
        .Range("D11").NumberFormat = "dd/mm/yyyy"
        
        .Range("F11").value = "Fin:"
        .Range("G11").Interior.Color = RGB(240, 240, 240)
        .Range("G11").Borders.LineStyle = xlContinuous
        .Range("G11").NumberFormat = "dd/mm/yyyy"
        
        ' --- CHAMP COULEUR ---
        .Range("B13").value = "5. Couleur"
        .Range("B13").Font.Bold = True
        .Range("C13:E13").Merge
        .Range("C13").Interior.Color = RGB(240, 240, 240)
        .Range("C13").Borders.LineStyle = xlContinuous
        
        ' Validation Couleur
        With .Range("C13").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="JAUNE,ORANGE,ROUGE,BLEU,VERT,ROSE,CYAN,GRIS,BLEU_CLAIR"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        .Range("C13").value = "JAUNE" ' Defaut
        
        ' --- BOUTON AJOUTER ---
        Set btn = .Buttons.Add(.Range("C15").Left, .Range("C15").Top, 200, 30)
        btn.Caption = "AJOUTER L'EXCEPTION"
        btn.OnAction = "ActionAjouterException"
        btn.Font.Bold = True
        btn.Font.Size = 12
        
        ' Ajustement largeurs
        .Columns("A").ColumnWidth = 2
        .Columns("B").ColumnWidth = 15
        .Columns("B:H").EntireColumn.AutoFit
    End With
    
    Application.ScreenUpdating = True
    MsgBox "Formulaire de saisie genere avec Gestion Couleurs !", vbInformation
End Sub

' Met a jour la plage nommee "ListeNomsSource" a partir de la feuille "Personnel"
Private Sub MettreAJourListeNoms()
    Dim ws As Worksheet, wsPerso As Worksheet
    Dim lastRow As Long, i As Long
    Dim arrNoms As Variant
    Dim listeFinale() As String
    Dim count As Long
    
    ' Cible: Feuille Interface pour stocker la liste cachee (Col AA)
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Interface_Exceptions")
    Set wsPerso = ThisWorkbook.Sheets("Personnel")
    On Error GoTo 0
    
    If wsPerso Is Nothing Then
        ' Fallback si Personnel n'existe pas : on prend la Sheet 1 Col A
        Set wsPerso = ThisWorkbook.Sheets(1)
        lastRow = wsPerso.Cells(wsPerso.Rows.count, "A").End(xlUp).row
        If lastRow < 2 Then lastRow = 2
        arrNoms = wsPerso.Range("A2:A" & lastRow).value
    Else
        ' Mode Personnel : Col B=Nom, Col C=Prenom
        lastRow = wsPerso.Cells(wsPerso.Rows.count, "B").End(xlUp).row
        If lastRow < 2 Then lastRow = 2
        arrNoms = wsPerso.Range("B2:C" & lastRow).value
    End If
    
    ' Construire liste "NOM Prenom"
    ReDim listeFinale(1 To UBound(arrNoms, 1))
    count = 0
    
    For i = 1 To UBound(arrNoms, 1)
        Dim s As String
        
        If wsPerso.Name = "Personnel" Then
             s = Trim(CStr(arrNoms(i, 1))) & " " & Trim(CStr(arrNoms(i, 2)))
        Else
             s = Trim(CStr(arrNoms(i, 1)))
        End If
        
        s = Trim(s)
        If s <> "" And UCase(s) <> "NOM" And UCase(s) <> "MATRICULE" Then
            count = count + 1
            listeFinale(count) = s
        End If
    Next i
    
    ' Ecrire en colonne AA de Interface (Cache)
    ws.Columns("AA").Clear
    If count > 0 Then
        ' Transpose pour ecrire en colonne (limite 65k lignes mais ok ici)
        ws.Range("AA1").Resize(count, 1).value = Application.Transpose(listeFinale)
        
        ' Nommer la plage
        ThisWorkbook.names.Add Name:="ListeNomsSource", RefersTo:="='" & ws.Name & "'!$AA$1:$AA$" & count
    End If
End Sub

' --- ACTION DU BOUTON ---
Public Sub ActionAjouterException()
    Dim ws As Worksheet, wsConf As Worksheet
    Set ws = ActiveSheet
    
    ' 1. LIRE LES VALEURS
    Dim sNom As String, sCode As String, sJours As String, sCoul As String
    Dim dDeb As Variant, dFin As Variant
    
    sNom = Trim(ws.Range("C5").value)
    sCode = Trim(ws.Range("C7").value)
    dDeb = ws.Range("D11").value
    dFin = ws.Range("G11").value
    sCoul = Trim(UCase(ws.Range("C13").value))
    If sCoul = "" Then sCoul = "JAUNE" ' Defaut
    
    ' 2. VALIDATION
    If sNom = "" Or sCode = "" Then
        MsgBox "Merci de remplir Nom et Code !", vbExclamation
        Exit Sub
    End If
    
    ' Lire Checkboxes
    Dim chk As Object ' CheckBox
    Dim jList As String
    jList = ""
    
    For Each chk In ws.CheckBoxes
        If chk.value = 1 Then ' Checked
            ' Le nom est chk_LUN, on veut LUN
            Dim n As String
            n = Replace(chk.Name, "chk_", "")
            n = Replace(n, " ", "") ' Nettoyage
            If jList <> "" Then jList = jList & ","
            jList = jList & n
        End If
    Next chk
    sJours = jList
    
    ' Validation Dates
    If Not IsEmpty(dDeb) And dDeb <> "" Then
        If Not IsDate(dDeb) Then
             MsgBox "Date de debut invalide", vbExclamation
             Exit Sub
        End If
    End If
    If Not IsEmpty(dFin) And dFin <> "" Then
        If Not IsDate(dFin) Then
             MsgBox "Date de fin invalide", vbExclamation
             Exit Sub
        End If
    End If
    
    ' 3. AJOUT DANS CONFIG_EXCEPTIONS
    On Error Resume Next
    Set wsConf = ThisWorkbook.Sheets("Config_Exceptions")
    On Error GoTo 0
    
    If wsConf Is Nothing Then
        Set wsConf = ThisWorkbook.Sheets.Add
        wsConf.Name = "Config_Exceptions"
        wsConf.Range("A1:F1").value = Array("Nom", "Code", "Jours", "DateDeb", "DateFin", "Couleur")
    End If
    
    ' ECRITURE (Pattern intelligent)
    If InStr(sNom, "*") = 0 Then
        ' Transforme "NOM Prenom" (espace ou underscore) en *NOM*Prenom*
        sNom = Replace(sNom, " ", "*")
        sNom = Replace(sNom, "_", "*")
        sNom = "*" & sNom & "*"
    End If
    
    ' RECHERCHE DOUBLON (Update si existe)
    Dim rowToWrite As Long
    Dim i As Long
    Dim lastRowConf As Long
    
    rowToWrite = 0
    lastRowConf = wsConf.Cells(wsConf.Rows.count, "A").End(xlUp).row
    
    If lastRowConf >= 2 Then
        For i = 2 To lastRowConf
            ' On compare Nom et Code pour identifier la regle
            If UCase(Trim(wsConf.Cells(i, 1).value)) = UCase(sNom) And _
               UCase(Trim(wsConf.Cells(i, 2).value)) = UCase(sCode) Then
                rowToWrite = i
                Exit For
            End If
        Next i
    End If
    
    ' Si pas trouve, on ecrit a la suite
    If rowToWrite = 0 Then
        rowToWrite = lastRowConf + 1
    End If
    
    wsConf.Cells(rowToWrite, 1).value = sNom
    wsConf.Cells(rowToWrite, 2).value = sCode
    wsConf.Cells(rowToWrite, 3).value = sJours
    wsConf.Cells(rowToWrite, 4).value = dDeb
    wsConf.Cells(rowToWrite, 5).value = dFin
    wsConf.Cells(rowToWrite, 6).value = sCoul ' Nouvelle Colonne
    
    ' 4. CONFIRMATION ET RESET
    If rowToWrite > lastRowConf Then
        MsgBox "Exception AJOUTEE pour " & sNom & " en " & sCoul, vbInformation
    Else
        MsgBox "Exception MISE A JOUR pour " & sNom & " en " & sCoul, vbInformation
    End If
    
End Sub
