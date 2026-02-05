Attribute VB_Name = "SaisieAnnuelle"
' ExportedAt: 2026-01-12 15:37:10 | Workbook: Planning_2026.xlsm
Option Explicit

'--- Fonctions Utilitaires (Communes et inchangées) -----------------------
Private Function GetSheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
End Function

Private Function GetTable(ws As Worksheet, tableName As String) As ListObject
    On Error Resume Next
    Set GetTable = ws.ListObjects(tableName)
    On Error GoTo 0
End Function

Private Function GetColIndex(tbl As ListObject, ParamArray names() As Variant) As Long
    Dim nm As Variant, col As ListColumn
    For Each nm In names
        For Each col In tbl.ListColumns
            If StrComp(col.Name, CStr(nm), vbTextCompare) = 0 Then
                GetColIndex = col.index
                Exit Function
            End If
        Next col
    Next nm
    GetColIndex = 0
End Function

'==========================================================================
' MACRO 1 : Synchronise la table Saisie Annuelle avec la table Personnel
' (Version avec bug de restauration des données corrigé)
'==========================================================================
Public Sub SynchroniserListePersonnel()
    Dim wsSaisie As Worksheet, wsPersonnel As Worksheet
    Dim tblSaisie As ListObject, tblPersonnel As ListObject
    Dim manualData As Object, srcData As Variant, outData() As Variant
    Dim i As Long, nbRows As Long, colManualStart As Long
    Dim idxMat As Long, idxNom As Long, idxPrenom As Long
    Dim anneeRef As Long, tmp As Variant

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Set wsSaisie = GetSheet("Saisie Annuelle")
    Set wsPersonnel = GetSheet("Personnel")
    If wsSaisie Is Nothing Or wsPersonnel Is Nothing Then
        MsgBox "Feuilles 'Saisie Annuelle' ou 'Personnel' introuvables.", vbCritical: GoTo CleanUp
    End If
    Set tblSaisie = GetTable(wsSaisie, "T_SaisieAnnuelle")
    Set tblPersonnel = GetTable(wsPersonnel, "T_Personnel")
    If tblSaisie Is Nothing Or tblPersonnel Is Nothing Then
        MsgBox "Tables 'T_SaisieAnnuelle' ou 'T_Personnel' introuvables.", vbCritical: GoTo CleanUp
    End If

    tmp = wsSaisie.Range("B1").value
    If IsNumeric(tmp) And tmp <> "" Then
        anneeRef = CLng(tmp)
        If anneeRef < 2020 Then
            MsgBox "L'année en B1 (" & anneeRef & ") semble trop ancienne.", vbExclamation: GoTo CleanUp
        End If
    Else
        MsgBox "La cellule B1 ne contient pas une année valide.", vbCritical: GoTo CleanUp
    End If

    colManualStart = 5
    Set manualData = CreateObject("Scripting.Dictionary")
    If tblSaisie.ListRows.count > 0 Then
        If tblSaisie.ListColumns.count >= colManualStart Then
            srcData = tblSaisie.DataBodyRange.Value2
            For i = 1 To UBound(srcData, 1)
                ' BUG CORRIGÉ : Utiliser COLUMN pour obtenir un tableau horizontal
                manualData(CStr(srcData(i, 1))) = Application.index(srcData, i, Evaluate("COLUMN(" & colManualStart & ":" & tblSaisie.ListColumns.count & ")"))
            Next i
        End If
        tblSaisie.DataBodyRange.Delete
    End If

    idxMat = GetColIndex(tblPersonnel, "Matricule")
    idxNom = GetColIndex(tblPersonnel, "Nom")
    idxPrenom = GetColIndex(tblPersonnel, "Prénom", "Prenom")
    If idxMat = 0 Or idxNom = 0 Or idxPrenom = 0 Then
        MsgBox "Colonnes Matricule, Nom ou Prénom introuvables dans la table Personnel.", vbCritical: GoTo CleanUp
    End If

    nbRows = tblPersonnel.ListRows.count
    If nbRows = 0 Then GoTo FinishFill

    srcData = tblPersonnel.DataBodyRange.Value2
    ReDim outData(1 To nbRows, 1 To 4)
    For i = 1 To nbRows
        outData(i, 1) = CStr(srcData(i, idxMat))
        outData(i, 2) = CStr(srcData(i, idxNom))
        outData(i, 3) = CStr(srcData(i, idxPrenom))
        outData(i, 4) = anneeRef
    Next i

    tblSaisie.Resize tblSaisie.HeaderRowRange.Resize(nbRows + 1)
    tblSaisie.DataBodyRange.value = outData

    If manualData.count > 0 And tblSaisie.ListColumns.count >= colManualStart Then
        For i = 1 To nbRows
            tmp = CStr(tblSaisie.DataBodyRange.Cells(i, 1).value)
            If manualData.Exists(tmp) Then
                tblSaisie.DataBodyRange.Cells(i, colManualStart).Resize(1, UBound(manualData(tmp), 2)).value = manualData(tmp)
            End If
        Next i
    End If

FinishFill:
    tblSaisie.ShowTableStyleRowStripes = False
    wsSaisie.Activate
    MsgBox "La liste des employés a été synchronisée pour l'année " & anneeRef & ".", vbInformation, "Synchronisation Terminée"

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub
ErrorHandler:
    MsgBox "Erreur inattendue dans SynchroniserListePersonnel." & vbCrLf & vbCrLf & "Erreur " & Err.Number & ": " & Err.description, vbCritical
    Resume CleanUp
End Sub

'==============================================================================
' MACRO 2 : Génère les affectations mensuelles
' (Version finale, ultra-rapide et robuste)
'==============================================================================
Public Sub GenererAffectations()
    Dim wsSaisie As Worksheet, wsAffect As Worksheet
    Dim tblSaisie As ListObject, tblAffect As ListObject
    Dim dataSaisie As Variant, dataResult() As Variant
    Dim anneeRef As Long, resultCounter As Long
    Dim i As Long, m As Long, key As Variant, mon As Variant
    Dim monthNames As Variant
    
    Dim idxS As Object, idxA As Object
    Set idxS = CreateObject("Scripting.Dictionary")
    Set idxA = CreateObject("Scripting.Dictionary")

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' --- 1. VÉRIFICATION DES FEUILLES ET TABLEAUX ---
    Set wsSaisie = GetSheet("Saisie Annuelle")
    Set wsAffect = GetSheet("Affectations")
    If wsSaisie Is Nothing Or wsAffect Is Nothing Then MsgBox "Feuilles 'Saisie Annuelle' ou 'Affectations' introuvables.", vbCritical: GoTo CleanUp
    Set tblSaisie = GetTable(wsSaisie, "T_SaisieAnnuelle")
    Set tblAffect = GetTable(wsAffect, "T_Affectations")
    If tblSaisie Is Nothing Or tblAffect Is Nothing Then MsgBox "Tables 'T_SaisieAnnuelle' ou 'T_Affectations' introuvables.", vbCritical: GoTo CleanUp
    
    ' --- 2. VÉRIFICATION DES COLONNES ---
    monthNames = Array("Janv", "Fev", "Mars", "Avril", "Mai", "Juin", "Juillet", "Aout", "Sept", "Oct", "Nov", "Dec")
    idxS.Add "Matricule", GetColIndex(tblSaisie, "Matricule"): idxS.Add "Nom", GetColIndex(tblSaisie, "Nom"): idxS.Add "Prénom", GetColIndex(tblSaisie, "Prénom", "Prenom"): idxS.Add "Position Base", GetColIndex(tblSaisie, "Position Base"): idxS.Add "% Base", GetColIndex(tblSaisie, "% Base")
    For Each mon In monthNames
        idxS.Add "Pos " & mon, GetColIndex(tblSaisie, "Pos " & mon)
        idxS.Add "% " & mon, GetColIndex(tblSaisie, "% " & mon)
    Next mon
    idxA.Add "Matricule", GetColIndex(tblAffect, "Matricule"): idxA.Add "Nom", GetColIndex(tblAffect, "Nom"): idxA.Add "Prénom", GetColIndex(tblAffect, "Prénom", "Prenom"): idxA.Add "Année", GetColIndex(tblAffect, "Année"): idxA.Add "Mois", GetColIndex(tblAffect, "Mois"): idxA.Add "Position", GetColIndex(tblAffect, "Position"): idxA.Add "Pourcentage", GetColIndex(tblAffect, "Pourcentage")
    
    For Each key In idxS.keys
        If idxS(key) = 0 Then MsgBox "Colonne manquante dans T_SaisieAnnuelle : '" & key & "'", vbCritical: GoTo CleanUp
    Next
    For Each key In idxA.keys
        If idxA(key) = 0 Then MsgBox "Colonne manquante dans T_Affectations : '" & key & "'", vbCritical: GoTo CleanUp
    Next

    ' --- 3. EXÉCUTION (TOUT EN MÉMOIRE) ---
    anneeRef = wsSaisie.Range("B1").value
    If anneeRef < 2020 Then MsgBox "L'année de référence en B1 est invalide.", vbCritical: GoTo CleanUp

    ' Vider totalement la table pour repartir proprement
    If Not tblAffect.DataBodyRange Is Nothing Then tblAffect.DataBodyRange.Delete
    
    If tblSaisie.ListRows.count = 0 Then GoTo Finish
    dataSaisie = tblSaisie.DataBodyRange.Value2
    ReDim dataResult(1 To UBound(dataSaisie, 1) * 12, 1 To tblAffect.ListColumns.count)
    resultCounter = 0

    For i = 1 To UBound(dataSaisie, 1)
        If Not IsEmpty(dataSaisie(i, idxS("Matricule"))) And dataSaisie(i, idxS("Matricule")) <> "" Then
            For m = 0 To 11
                resultCounter = resultCounter + 1
                dataResult(resultCounter, idxA("Matricule")) = dataSaisie(i, idxS("Matricule"))
                dataResult(resultCounter, idxA("Nom")) = dataSaisie(i, idxS("Nom"))
                dataResult(resultCounter, idxA("Prénom")) = dataSaisie(i, idxS("Prénom"))
                dataResult(resultCounter, idxA("Année")) = anneeRef
                dataResult(resultCounter, idxA("Mois")) = monthNames(m)
                
                If Not IsEmpty(dataSaisie(i, idxS("Pos " & monthNames(m)))) And dataSaisie(i, idxS("Pos " & monthNames(m))) <> "" Then
                    dataResult(resultCounter, idxA("Position")) = dataSaisie(i, idxS("Pos " & monthNames(m)))
                Else
                    dataResult(resultCounter, idxA("Position")) = dataSaisie(i, idxS("Position Base"))
                End If
                
                If Not IsEmpty(dataSaisie(i, idxS("% " & monthNames(m)))) Then
                    dataResult(resultCounter, idxA("Pourcentage")) = dataSaisie(i, idxS("% " & monthNames(m)))
                Else
                    dataResult(resultCounter, idxA("Pourcentage")) = dataSaisie(i, idxS("% Base"))
                End If
            Next m
        End If
    Next i
    
Finish:
    If resultCounter > 0 Then
        Dim finalData() As Variant
        ReDim finalData(1 To resultCounter, 1 To tblAffect.ListColumns.count)
        Dim c As Long
        For i = 1 To resultCounter
            For c = 1 To tblAffect.ListColumns.count
                finalData(i, c) = dataResult(i, c)
            Next c
        Next i

        Dim startRow As Long
        If tblAffect.ListRows.count = 0 Then
            tblAffect.Resize tblAffect.HeaderRowRange.Resize(resultCounter + 1)
            tblAffect.DataBodyRange.value = finalData
        Else
            startRow = tblAffect.ListRows.count + 1
            tblAffect.Resize tblAffect.HeaderRowRange.Resize(startRow + resultCounter)
            tblAffect.DataBodyRange.Rows(startRow).Resize(resultCounter).value = finalData
        End If
    End If

    MsgBox "La mise à jour des affectations pour l'année " & anneeRef & " est terminée ! (" & resultCounter & " lignes traitées)", vbInformation, "Opération Réussie"

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub
ErrorHandler:
    MsgBox "Erreur inattendue dans GenererAffectations." & vbCrLf & vbCrLf & "Erreur " & Err.Number & ": " & Err.description, vbCritical
    Resume CleanUp
End Sub
