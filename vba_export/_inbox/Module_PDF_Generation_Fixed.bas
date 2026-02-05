' Module: Module_PDF_Generation_Fixed
' Description: Version corrigée utilisant CfgText au lieu de LireParametre
'              pour éviter les doublons de configuration
Option Explicit

' --- Boutons publics ---

Public Sub Generate_PDF_Jour()
    ProcessPDFGeneration "Jour"
End Sub

Public Sub Generate_PDF_Nuit()
    ProcessPDFGeneration "Nuit"
End Sub

' --- Orchestrateur principal ---

Private Sub ProcessPDFGeneration(ByVal equipe As String)
    If Application.Ready = False Then
        MsgBox "Valide/annule l'édition de cellule (Entrée/Echap) puis relance.", vbExclamation
        Exit Sub
    End If
    If Not TypeOf ActiveSheet Is Worksheet Then
        MsgBox "Sélectionne la feuille du mois (OCT, NOV, DEC...).", vbExclamation
        Exit Sub
    End If

    On Error GoTo FinalErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    UpdateStatusBar "Étape 1/3 : Archivage (" & equipe & ")..."
    If Not ArchivePreviousMonthPDFs(equipe) Then GoTo CleanUp

    UpdateStatusBar "Étape 2/3 : Nettoyage (" & equipe & ")..."
    If Not CleanupArchivedPDFs(equipe) Then GoTo CleanUp

    UpdateStatusBar "Étape 3/3 : Export PDF (" & equipe & ")..."
    If Not ExportHorairePDF(ActiveSheet, equipe) Then GoTo CleanUp

    GoTo CleanUp

FinalErrorHandler:
    MsgBox "Processus interrompu : " & Err.Description, vbCritical

CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    UpdateStatusBar False
End Sub


' ================================================================================
'                     ARCHIVE DU MOIS PRÉCÉDENT
' ================================================================================

Private Function ArchivePreviousMonthPDFs(ByVal equipe As String) As Boolean
    Dim parentPathRel As String: parentPathRel = CfgTextOr("PDF_CheminParentRelatif", "")
    Dim teamFolder As String:    teamFolder = CfgTextOr("PDF_Dossier_" & equipe, "")
    Dim archiveSub As String:    archiveSub = CfgTextOr("PDF_Archive_SousDossier_" & equipe, "")

    If parentPathRel = "" Or teamFolder = "" Or archiveSub = "" Then
        MsgBox "Config manquante (Feuil_Config) pour l'archivage " & equipe, vbCritical
        Exit Function
    End If

    Dim baseOneDrivePath As String: baseOneDrivePath = GetOneDriveBasePath()
    If baseOneDrivePath = "" Then
        ArchivePreviousMonthPDFs = True: Exit Function
    End If

    Dim fullParentPath As String
    fullParentPath = baseOneDrivePath & parentPathRel
    If Right(fullParentPath, 1) <> "\" Then fullParentPath = fullParentPath & "\"

    Dim prevMonthDate As Date
    Dim planningYear As Integer: planningYear = GetPlanningYear()
    Dim currentMonth As Integer: currentMonth = Month(Date)
    
    If currentMonth = 1 Then
        prevMonthDate = DateSerial(planningYear - 1, 12, 1)
    Else
        prevMonthDate = DateSerial(planningYear, currentMonth - 1, 1)
    End If
    
    Dim prevMonthName As String: prevMonthName = DateToFrenchMonthName(prevMonthDate)

    Dim sourceFolder As String:  sourceFolder = fullParentPath & teamFolder & "\"
    Dim archiveFolder As String: archiveFolder = sourceFolder & archiveSub & "\"
    EnsurePathExists archiveFolder

    Dim pdfName As String
    pdfName = "Horaire " & prevMonthName & "_" & equipe & ".pdf"

    If Dir(sourceFolder & pdfName) <> "" Then
        On Error Resume Next
        Name sourceFolder & pdfName As archiveFolder & pdfName
        On Error GoTo 0
    End If

    ArchivePreviousMonthPDFs = True
End Function


' ================================================================================
'                     NETTOYAGE DES ARCHIVES
' ================================================================================

Private Function CleanupArchivedPDFs(ByVal equipe As String) As Boolean
    Dim parentPathRel As String: parentPathRel = CfgTextOr("PDF_CheminParentRelatif", "")
    Dim teamFolder As String:    teamFolder = CfgTextOr("PDF_Dossier_" & equipe, "")
    Dim archiveSub As String:    archiveSub = CfgTextOr("PDF_Archive_SousDossier_" & equipe, "")

    If parentPathRel = "" Or teamFolder = "" Or archiveSub = "" Then
        CleanupArchivedPDFs = True: Exit Function
    End If

    Dim baseOneDrivePath As String: baseOneDrivePath = GetOneDriveBasePath()
    If baseOneDrivePath = "" Then
        CleanupArchivedPDFs = True: Exit Function
    End If

    Dim fullParentPath As String
    fullParentPath = baseOneDrivePath & parentPathRel
    If Right(fullParentPath, 1) <> "\" Then fullParentPath = fullParentPath & "\"

    Dim planningYear As Integer: planningYear = GetPlanningYear()
    Dim currentMonth As Integer: currentMonth = Month(Date)
    
    Dim cleanupMonth As Integer, cleanupYear As Integer
    cleanupMonth = currentMonth - 3
    cleanupYear = planningYear
    
    If cleanupMonth <= 0 Then
        cleanupMonth = cleanupMonth + 12
        cleanupYear = cleanupYear - 1
    End If
    
    Dim cleanupDate As Date: cleanupDate = DateSerial(cleanupYear, cleanupMonth, 1)
    Dim cleanupMonthName As String: cleanupMonthName = DateToFrenchMonthName(cleanupDate)

    Dim archiveFolder As String
    archiveFolder = fullParentPath & teamFolder & "\" & archiveSub & "\"

    Dim fileToDelete As String
    fileToDelete = archiveFolder & "Horaire " & cleanupMonthName & "_" & equipe & ".pdf"

    If Dir(fileToDelete) <> "" Then
        On Error Resume Next
        Kill fileToDelete
        On Error GoTo 0
    End If

    CleanupArchivedPDFs = True
End Function


' ================================================================================
'                     EXPORT PDF DU MOIS ACTIF
' ================================================================================

Private Function ExportHorairePDF(ByVal ws As Worksheet, ByVal equipe As String) As Boolean
    Dim parentPathRel As String: parentPathRel = CfgTextOr("PDF_CheminParentRelatif", "")
    Dim folderConfig As String:  folderConfig = CfgTextOr("PDF_Dossier_" & equipe, "")
    Dim archiveSub As String:    archiveSub = CfgTextOr("PDF_Archive_SousDossier_" & equipe, "")

    Dim printArea As String
    printArea = CfgTextOr("PDF_PrintArea_" & equipe, "")
    If printArea = "" Then printArea = CfgTextOr("PDF_PrintArea", "")

    Dim missing As String
    If parentPathRel = "" Then missing = missing & vbCrLf & "- PDF_CheminParentRelatif"
    If folderConfig = "" Then missing = missing & vbCrLf & "- PDF_Dossier_" & equipe
    If printArea = "" Then missing = missing & vbCrLf & "- PDF_PrintArea(_" & equipe & ")"
    If missing <> "" Then
        MsgBox "Export annulé. Paramètres manquants :" & missing, vbCritical: Exit Function
    End If

    Dim baseOneDrivePath As String: baseOneDrivePath = GetOneDriveBasePath()
    If baseOneDrivePath = "" Then
        MsgBox "Chemin OneDrive introuvable.", vbCritical: Exit Function
    End If

    Dim overrideBase As String
    overrideBase = Trim(CfgTextOr("PDF_BasePath_Override", ""))
    If overrideBase <> "" Then
        If Right(overrideBase, 1) <> "\" Then overrideBase = overrideBase & "\"
        baseOneDrivePath = overrideBase
    End If

    Dim sheetMonthDate As Date: sheetMonthDate = ParseSheetNameToDate(ws.Name)
    If sheetMonthDate = 0 Then
        MsgBox "Nom d'onglet non reconnu : " & ws.Name, vbCritical: Exit Function
    End If

    Dim teamFolderPath As String
    teamFolderPath = baseOneDrivePath & parentPathRel & folderConfig & "\"

    Dim currentMonthStart As Date: currentMonthStart = DateSerial(Year(Date), Month(Date), 1)
    Dim alwaysLive As String: alwaysLive = UCase(Trim(CfgTextOr("PDF_AlwaysLive", "")))
    Dim isPastMonth As Boolean: isPastMonth = (sheetMonthDate < currentMonthStart) And (alwaysLive <> "1")

    Dim targetPdfFolderPath As String
    If isPastMonth Then
        targetPdfFolderPath = teamFolderPath & archiveSub & "\"
    Else
        targetPdfFolderPath = teamFolderPath
    End If
    EnsurePathExists targetPdfFolderPath

    Dim formattedMonthName As String: formattedMonthName = DateToFrenchMonthName(sheetMonthDate)
    Dim pdfFileName As String: pdfFileName = "Horaire " & formattedMonthName & "_" & equipe & ".pdf"
    Dim fullPdfPath As String: fullPdfPath = targetPdfFolderPath & pdfFileName

    ' Debug: Afficher le chemin cible
    Debug.Print "PDF Target: " & fullPdfPath
    
    ' Masquage automatique pour Planning Nuit
    Dim rowsToHide As Variant
    Dim wasHidden() As Boolean
    Dim i As Long
    Dim needsHiding As Boolean: needsHiding = False
    
    If UCase(equipe) = "NUIT" Then
        rowsToHide = Array(5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, _
                           21, 22, 23, 24, 25, 26, 27, 28, 29, 30, _
                           48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, _
                           60, 61, 62, _
                           64, 65, 66, 67, 68, 69, 70)
        
        needsHiding = True
        ReDim wasHidden(LBound(rowsToHide) To UBound(rowsToHide))
        
        For i = LBound(rowsToHide) To UBound(rowsToHide)
            wasHidden(i) = ws.Rows(rowsToHide(i)).Hidden
            ws.Rows(rowsToHide(i)).Hidden = True
        Next i
        
        printArea = "$A$1:$AG$73"
    End If
    
    printArea = Replace(printArea, ";", ",")
    
    ' Configuration PageSetup
    With ws.PageSetup
        .printArea = printArea
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0.3)
        .FooterMargin = Application.CentimetersToPoints(0.3)
    End With

    On Error GoTo ExportError

    ' Supprimer ancien PDF
    If Dir(fullPdfPath) <> "" Then
        On Error Resume Next
        Kill fullPdfPath
        On Error GoTo ExportError
    End If

    ' Export PDF
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fullPdfPath, _
        Quality:=xlQualityStandard, OpenAfterPublish:=False

    ' Restaurer lignes masquées
    If needsHiding Then
        For i = LBound(rowsToHide) To UBound(rowsToHide)
            ws.Rows(rowsToHide(i)).Hidden = wasHidden(i)
        Next i
    End If

    MsgBox "PDF généré ici :" & vbCrLf & fullPdfPath, vbInformation, "Destination PDF"
    On Error Resume Next
    Shell "explorer.exe /select,""" & fullPdfPath & """", vbNormalFocus
    On Error GoTo 0

    ExportHorairePDF = True
    Exit Function

ExportError:
    If needsHiding Then
        For i = LBound(rowsToHide) To UBound(rowsToHide)
            ws.Rows(rowsToHide(i)).Hidden = wasHidden(i)
        Next i
    End If
    
    MsgBox "Erreur export PDF (" & equipe & ") : " & Err.Description & vbCrLf & _
           "Chemin cible: " & fullPdfPath & vbCrLf & vbCrLf & _
           "Solutions:" & vbCrLf & _
           "1. Sauvegardez le classeur (Ctrl+S)" & vbCrLf & _
           "2. Fermez le PDF s'il est ouvert" & vbCrLf & _
           "3. Vérifiez que le dossier OneDrive existe", vbCritical
    ExportHorairePDF = False
End Function


' ================================================================================
'                     OUTILS SUPPORT
' ================================================================================

Private Sub UpdateStatusBar(ByVal message As Variant)
    If VarType(message) = vbBoolean Then
        Application.StatusBar = False
    Else
        Application.StatusBar = CStr(message)
    End If
    DoEvents
End Sub

Private Function GetPlanningYear() As Integer
    Dim anneeStr As String
    anneeStr = Trim(CfgTextOr("AnneePlanning", ""))
    
    If anneeStr <> "" And IsNumeric(anneeStr) Then
        GetPlanningYear = CInt(anneeStr)
    Else
        GetPlanningYear = Year(Date)
    End If
End Function

Private Function GetOneDriveBasePath() As String
    Dim fso As Object, pathGuess As String
    On Error Resume Next
    pathGuess = Environ("OneDrive"): If pathGuess = "" Then pathGuess = Environ("OneDriveConsumer")
    If pathGuess = "" Then
        Dim userProfile As String: userProfile = Environ("UserProfile")
        If userProfile <> "" Then
            Set fso = CreateObject("Scripting.FileSystemObject")
            If fso.FolderExists(userProfile) Then
                Dim baseFolder As Object, subFolder As Object
                Set baseFolder = fso.GetFolder(userProfile)
                For Each subFolder In baseFolder.SubFolders
                    Dim upperName As String: upperName = UCase(subFolder.Name)
                    If upperName = "ONEDRIVE" Or Left(upperName, 9) = "ONEDRIVE -" Then
                        pathGuess = subFolder.Path: Exit For
                    End If
                Next subFolder
            End If
        End If
    End If
    If pathGuess = "" Then
        Dim fallback As String: fallback = Environ("UserProfile") & "\OneDrive"
        If fallback <> "" Then pathGuess = fallback
    End If
    On Error GoTo 0

    If pathGuess <> "" Then
        If Right(pathGuess, 1) <> "\" Then pathGuess = pathGuess & "\"
        GetOneDriveBasePath = pathGuess
    Else
        GetOneDriveBasePath = ""
    End If
End Function

Private Sub EnsurePathExists(ByVal targetPath As String)
    Dim parts() As String, i As Long, p As String
    parts = Split(targetPath, "\")
    If UBound(parts) < 0 Then Exit Sub
    p = parts(0)
    For i = 1 To UBound(parts)
        p = p & "\" & parts(i)
        If Len(p) > 3 Then
            On Error Resume Next
            If Dir(p, vbDirectory) = "" Then MkDir p
            On Error GoTo 0
        End If
    Next i
End Sub

Private Function DateToFrenchMonthName(ByVal d As Date) As String
    Dim noms As Variant
    noms = Array("", "Janvier", "Février", "Mars", "Avril", "Mai", "Juin", _
                      "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre")
    DateToFrenchMonthName = noms(Month(d))
End Function

Private Function ParseSheetNameToDate(ByVal sheetName As String) As Date
    Dim m As Integer, nameNorm As String
    nameNorm = UCase(Trim(sheetName))
    nameNorm = Replace(nameNorm, "À", "A"): nameNorm = Replace(nameNorm, "Â", "A")
    nameNorm = Replace(nameNorm, "Ä", "A"): nameNorm = Replace(nameNorm, "Á", "A")
    nameNorm = Replace(nameNorm, "É", "E"): nameNorm = Replace(nameNorm, "È", "E")
    nameNorm = Replace(nameNorm, "Ê", "E"): nameNorm = Replace(nameNorm, "Ë", "E")
    nameNorm = Replace(nameNorm, "Í", "I"): nameNorm = Replace(nameNorm, "Î", "I")
    nameNorm = Replace(nameNorm, "Ï", "I"): nameNorm = Replace(nameNorm, "Ó", "O")
    nameNorm = Replace(nameNorm, "Ô", "O"): nameNorm = Replace(nameNorm, "Ö", "O")
    nameNorm = Replace(nameNorm, "Ú", "U"): nameNorm = Replace(nameNorm, "Û", "U")
    nameNorm = Replace(nameNorm, "Ü", "U"): nameNorm = Replace(nameNorm, "Ç", "C")

    Select Case nameNorm
        Case "JANV", "JANVIER":              m = 1
        Case "FEV", "FEVR", "FEVRIER":       m = 2
        Case "MARS":                         m = 3
        Case "AVR", "AVRIL":                 m = 4
        Case "MAI":                          m = 5
        Case "JUIN":                         m = 6
        Case "JUIL", "JUILLET":              m = 7
        Case "AOUT":                         m = 8
        Case "SEPT", "SEPTEMBRE":            m = 9
        Case "OCT", "OCTOBRE":               m = 10
        Case "NOV", "NOVEMBRE":              m = 11
        Case "DEC", "DECEMBRE":              m = 12
        Case Else:                           m = 0
    End Select

    If m = 0 Then
        ParseSheetNameToDate = 0
    Else
        ParseSheetNameToDate = DateSerial(GetPlanningYear(), m, 1)
    End If
End Function
