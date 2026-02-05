' ExportedAt: 2026-01-12 15:37:09 | Workbook: Planning_2026.xlsm
Attribute VB_Name = "ModulePDFGeneration"
Option Explicit

' =================================================================================
' Module         : Module_PDF_Generation_New (VERSION FINALE 2025-11-30 v2)
' Description    :
'   VERSION FINALE CORRIGÉE - Corrections appliquées :
'   1. Lit l'année depuis le paramètre "AnneePlanning" dans Feuil_Config
'   2. Force l'impression sur 1 SEULE PAGE (FitToPagesWide/Tall = 1)
'   3. SOLUTION 3 : Masquage automatique des lignes pour Planning Nuit
'
'   MODIFICATIONS CRITIQUES:
'   -------------------------
'   1. GetPlanningYear() : Lit l'année depuis Feuil_Config (AnneePlanning)
'   2. ParseSheetNameToDate() : Utilise GetPlanningYear() au lieu de Year(Date)
'   3. ExportHorairePDF() :
'      - Pour NUIT : Masque automatiquement les lignes non désirées
'      - Utilise une zone continue $A$1:$AG$73 pour Nuit
'      - Démasque les lignes après l'export
'      - Pour JOUR : Fonctionne normalement
'
'   RÉSULTAT:
'   ---------
'   - Planning 2026 : PDFs dans le bon dossier (Live au lieu d'Archive) ?
'   - PDF Nuit : 1 page au lieu de 4 pages (masquage auto) ?
'   - PDF Jour : 1 page (fonctionne comme avant) ?
'   - Pas d'erreur "formule incorrecte" ?
' =================================================================================


' --- Boutons publics ---

Public Sub Generate_PDF_Jour()
    ProcessPDFGeneration "Jour"
End Sub

Public Sub Generate_PDF_Nuit()
    ProcessPDFGeneration "Nuit"
End Sub

' Outil : normalise tous les anciens fichiers (Jour + Nuit)
Public Sub Normalize_All_Names()
    NormalizeExistingPdfNames "Jour"
    NormalizeExistingPdfNames "Nuit"
    MsgBox "Normalisation des noms terminée pour Jour & Nuit.", vbInformation
End Sub

' Outil : affiche les cibles sans exporter (Jour & Nuit pour l'onglet actif)
Public Sub Diag_AfficherCibles()
    Dim ws As Worksheet, m As Date, base As String, rel As String, fold As String, arch As String
    Dim curStart As Date, isPast As Boolean, monthName As String, cible As String
    Dim overrideBase As String
    
    Set ws = ActiveSheet
    m = ParseSheetNameToDate(ws.Name)
    If m = 0 Then
        MsgBox "Nom d'onglet non reconnu pour déduire le mois : " & ws.Name, vbCritical: Exit Sub
    End If
    monthName = DateToFrenchMonthName(m)
    curStart = DateSerial(Year(Date), Month(Date), 1)
    isPast = (m < curStart) And (UCase(Trim(LireParametre("PDF_AlwaysLive"))) <> "1")
    
    ' Base path avec override
    base = GetOneDriveBasePath()
    overrideBase = Trim(LireParametre("PDF_BasePath_Override"))
    If overrideBase <> "" Then
        If Right(overrideBase, 1) <> "\" Then overrideBase = overrideBase & "\"
        base = overrideBase
    End If
    
    rel = LireParametre("PDF_CheminParentRelatif")
    Dim equipe As Variant
    For Each equipe In Array("Jour", "Nuit")
        fold = LireParametre("PDF_Dossier_" & equipe)
        arch = LireParametre("PDF_Archive_SousDossier_" & equipe)
        If base = "" Or rel = "" Or fold = "" Then
            MsgBox "Paramètres manquants pour " & equipe & " (base/rel/folder).", vbCritical
        Else
            cible = base & rel & fold & IIf(isPast, "\" & arch, "") & "\" & _
                    "Horaire " & monthName & "_" & equipe & ".pdf"
            MsgBox "Destination " & equipe & " :" & vbCrLf & cible, vbInformation, _
                   "Test ( " & IIf(isPast, "ARCHIVE", "LIVE") & " )"
        End If
    Next equipe
End Sub

' Outil : ping ciblé pour Jour (ouvre l'explorateur directement sur le fichier)
Public Sub Ping_Where_Jour()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim base As String, rel As String, fold As String, arch As String
    Dim m As Date, curStart As Date, isPast As Boolean
    Dim monthName As String, targetPath As String, pdfName As String
    Dim overrideBase As String
    
    m = ParseSheetNameToDate(ws.Name)
    If m = 0 Then MsgBox "Onglet non reconnu : " & ws.Name, vbCritical: Exit Sub
    
    base = GetOneDriveBasePath()
    overrideBase = Trim(LireParametre("PDF_BasePath_Override"))
    If overrideBase <> "" Then
        If Right(overrideBase, 1) <> "\" Then overrideBase = overrideBase & "\"
        base = overrideBase
    End If
    
    rel = LireParametre("PDF_CheminParentRelatif")
    fold = LireParametre("PDF_Dossier_Jour")
    arch = LireParametre("PDF_Archive_SousDossier_Jour")
    If base = "" Or rel = "" Or fold = "" Then
        MsgBox "Param manquant (base/rel/fold).", vbCritical: Exit Sub
    End If
    
    curStart = DateSerial(Year(Date), Month(Date), 1)
    isPast = (m < curStart) And (UCase(Trim(LireParametre("PDF_AlwaysLive"))) <> "1")
    monthName = DateToFrenchMonthName(m)
    
    targetPath = base & rel & fold & IIf(isPast, "\" & arch, "") & "\"
    pdfName = "Horaire " & monthName & "_Jour.pdf"
    
    MsgBox "RESOLU :" & vbCrLf & targetPath & pdfName, vbInformation, IIf(isPast, "ARCHIVE", "LIVE")
    On Error Resume Next
    Shell "explorer.exe /select,""" & targetPath & pdfName & """", vbNormalFocus
    On Error GoTo 0
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
    Dim parentPathRel As String: parentPathRel = LireParametre("PDF_CheminParentRelatif")
    Dim teamFolder As String:    teamFolder = LireParametre("PDF_Dossier_" & equipe)
    Dim archiveSub As String:    archiveSub = LireParametre("PDF_Archive_SousDossier_" & equipe)

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
    Dim parentPathRel As String: parentPathRel = LireParametre("PDF_CheminParentRelatif")
    Dim teamFolder As String:    teamFolder = LireParametre("PDF_Dossier_" & equipe)
    Dim archiveSub As String:    archiveSub = LireParametre("PDF_Archive_SousDossier_" & equipe)

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
'                     EXPORT PDF DU MOIS ACTIF (? SOLUTION 3)
' ================================================================================

Private Function ExportHorairePDF(ByVal ws As Worksheet, ByVal equipe As String) As Boolean
    Dim parentPathRel As String: parentPathRel = LireParametre("PDF_CheminParentRelatif")
    Dim folderConfig As String:  folderConfig = LireParametre("PDF_Dossier_" & equipe)
    Dim archiveSub As String:    archiveSub = LireParametre("PDF_Archive_SousDossier_" & equipe)

    Dim printArea As String
    printArea = LireParametre("PDF_PrintArea_" & equipe)
    If printArea = "" Then printArea = LireParametre("PDF_PrintArea")

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
    overrideBase = Trim(LireParametre("PDF_BasePath_Override"))
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
    Dim alwaysLive As String: alwaysLive = UCase(Trim(LireParametre("PDF_AlwaysLive")))
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

    ' ? SOLUTION 3 : Masquage automatique pour Planning Nuit
    Dim rowsToHide As Variant
    Dim wasHidden() As Boolean
    Dim i As Long
    Dim needsHiding As Boolean: needsHiding = False
    
    ' Si Planning Nuit : masquer temporairement les lignes non désirées
    If UCase(equipe) = "NUIT" Then
        ' Zones à garder : 1-4, 31-47, 59, 63, 71-73
        ' Zones à masquer : 5-30, 48-58, 60-62, 64-70
        rowsToHide = Array(5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, _
                           21, 22, 23, 24, 25, 26, 27, 28, 29, 30, _
                           48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, _
                           60, 61, 62, _
                           64, 65, 66, 67, 68, 69, 70)
        
        needsHiding = True
        ReDim wasHidden(LBound(rowsToHide) To UBound(rowsToHide))
        
        ' Sauvegarder l'état actuel et masquer
        For i = LBound(rowsToHide) To UBound(rowsToHide)
            wasHidden(i) = ws.Rows(rowsToHide(i)).Hidden
            ws.Rows(rowsToHide(i)).Hidden = True
        Next i
        
        ' Utiliser une zone continue pour Nuit
        printArea = "$A$1:$AG$73"
    End If
    
    ' Remplacer les point-virgules par des virgules (requis par VBA)
    printArea = Replace(printArea, ";", ",")
    
    ' Configuration complète du PageSetup
    With ws.PageSetup
        .printArea = printArea
        
        ' Forcer l'impression sur 1 page
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        
        ' Orientation et format
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        
        ' Marges réduites
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0.3)
        .FooterMargin = Application.CentimetersToPoints(0.3)
    End With

    On Error GoTo ExportError

    ' Écrasement sûr
    If Dir(fullPdfPath) <> "" Then
        On Error Resume Next
        Kill fullPdfPath
        On Error GoTo 0
    End If

    ' Export
    ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=fullPdfPath, _
        Quality:=xlQualityStandard, OpenAfterPublish:=False

    ' ? Restaurer l'état des lignes après export (Nuit uniquement)
    If needsHiding Then
        For i = LBound(rowsToHide) To UBound(rowsToHide)
            ws.Rows(rowsToHide(i)).Hidden = wasHidden(i)
        Next i
    End If

    ' Trace + ouverture dossier
    MsgBox "PDF généré ici :" & vbCrLf & fullPdfPath, vbInformation, "Destination PDF"
    On Error Resume Next
    Shell "explorer.exe /select,""" & fullPdfPath & """", vbNormalFocus
    On Error GoTo 0

    ' WhatsApp (facultatif)
    NotifyUserAfterExport pdfFileName, formattedMonthName, equipe, isPastMonth

    ExportHorairePDF = True
    Exit Function

ExportError:
    ' Restaurer les lignes même en cas d'erreur
    If needsHiding Then
        For i = LBound(rowsToHide) To UBound(rowsToHide)
            ws.Rows(rowsToHide(i)).Hidden = wasHidden(i)
        Next i
    End If
    
    MsgBox "Erreur export PDF (" & equipe & ") : " & Err.Description & vbCrLf & _
           "Vérifie la zone d'impression dans Feuil_Config.", vbCritical
    ExportHorairePDF = False
End Function


' ================================================================================
'                     NORMALISATION DES ANCIENS FICHIERS
' ================================================================================

Private Sub NormalizeExistingPdfNames(ByVal equipe As String)
    Dim base As String, rel As String, teamFold As String, arch As String
    base = GetOneDriveBasePath()
    Dim overrideBase As String: overrideBase = Trim(LireParametre("PDF_BasePath_Override"))
    If overrideBase <> "" Then
        If Right(overrideBase, 1) <> "\" Then overrideBase = overrideBase & "\"
        base = overrideBase
    End If
    
    rel = LireParametre("PDF_CheminParentRelatif")
    teamFold = LireParametre("PDF_Dossier_" & equipe)
    arch = LireParametre("PDF_Archive_SousDossier_" & equipe)

    If base = "" Or rel = "" Or teamFold = "" Then Exit Sub

    Dim livePath As String, archPath As String
    livePath = base & rel & teamFold & "\"
    archPath = livePath & arch & "\"

    NormalizeInFolder livePath
    NormalizeInFolder archPath
End Sub

Private Sub NormalizeInFolder(ByVal folderPath As String)
    On Error Resume Next
    If folderPath = "" Then Exit Sub
    If Dir(folderPath, vbDirectory) = "" Then Exit Sub
    On Error GoTo 0

    Dim fso As Object, f As Object, fl As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set f = fso.GetFolder(folderPath)
    If f Is Nothing Then Exit Sub
    On Error GoTo 0

    For Each fl In f.Files
        If LCase(Right(fl.Name, 4)) = ".pdf" Then
            If Left(fl.Name, 8) = "Horaire_" Then
                Dim newName As String
                newName = "Horaire " & Mid(fl.Name, 9)
                If newName <> fl.Name Then
                    On Error Resume Next
                    fso.MoveFile fl.Path, f.Path & "\" & newName
                    On Error GoTo 0
                End If
            End If
        End If
    Next fl
End Sub


' ================================================================================
'                     NOTIFICATION WHATSAPP (OPTIONNELLE)
' ================================================================================

Private Sub NotifyUserAfterExport( _
    ByVal pdfName As String, _
    ByVal monthName As String, _
    ByVal equipe As String, _
    ByVal wasArchived As Boolean)

    Dim whatsappNum As String
    whatsappNum = LireParametre("WhatsApp_Numero")

    Dim baseMsg As String
    baseMsg = "Le PDF '" & pdfName & "' a été exporté avec succès."

    If whatsappNum = "" Then
        MsgBox baseMsg, vbInformation, "Exportation terminée"
        Exit Sub
    End If

    Dim ask As VbMsgBoxResult
    ask = MsgBox(baseMsg & vbCrLf & vbCrLf & _
                 "Ouvrir WhatsApp pour prévenir le groupe Team ?", _
                 vbYesNo + vbQuestion, "Exportation terminée")

    If ask = vbYes Then
        Dim messageText As String
        messageText = "Planning " & monthName & " (" & equipe & ") mis à jour – PDF prêt."
        If wasArchived Then messageText = messageText & " (Archivé)"

        Dim encodedText As String
        On Error Resume Next
        encodedText = Application.WorksheetFunction.EncodeURL(messageText)
        If Err.Number <> 0 Then
            Err.Clear: encodedText = Replace(messageText, " ", "%20")
        End If
        On Error GoTo 0

        Dim whatsappLink As String
        whatsappLink = "https://wa.me/" & whatsappNum & "?text=" & encodedText
        ThisWorkbook.FollowHyperlink whatsappLink
    Else
        MsgBox baseMsg, vbInformation, "Exportation terminée"
    End If
End Sub


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

Private Function LireParametre(ByVal param As String) As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Feuil_Config")
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets("Configuration")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    Dim cell As Range
    For Each cell In ws.Range("A1:A" & lastRow)
        If Trim(UCase(cell.value)) = Trim(UCase(param)) Then
            LireParametre = Trim(CStr(cell.offset(0, 1).value))
            Exit Function
        End If
    Next cell
End Function

Private Function GetPlanningYear() As Integer
    Dim anneeStr As String
    anneeStr = Trim(LireParametre("AnneePlanning"))
    
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

