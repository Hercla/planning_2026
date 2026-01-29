Attribute VB_Name = "Module_CodeProcessor"
' ExportedAt: 2026-01-12 15:37:09 | Workbook: Planning_2026.xlsm
Option Explicit

'================================================================================================
' MODULE :          Module_CodeProcessor (Optimized & Robust Cleaning)
' DESCRIPTION :     Reads code configuration with enhanced string normalization to prevent
'                   mismatches due to spacing or hidden characters.
'================================================================================================

Private codeDict As Object ' cache

' --- NOUVELLE FONCTION DE NETTOYAGE ---
Private Function NormalizeString(ByVal inputText As String) As String
    ' Nettoie une chaîne pour garantir une comparaison fiable.
    ' 1. Remplace les espaces insécables par des espaces normaux.
    Dim s As String
    s = Replace(inputText, Chr(160), " ")
    
    ' 2. Boucle pour supprimer les doubles espaces au milieu du texte.
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    
    ' 3. Supprime les espaces au début et à la fin.
    NormalizeString = Trim(s)
End Function


Public Function GetCodeInfo(ByVal code As String) As clsCodeInfo
    ' Returns information about a specific code. Loads definitions if necessary.
    If codeDict Is Nothing Then LoadCodeDefinitions

    ' --- MODIFIÉ : Utilise la nouvelle fonction de nettoyage ---
    Dim cleanCode As String
    cleanCode = NormalizeString(code)

    If codeDict.Exists(cleanCode) Then
        Set GetCodeInfo = codeDict(cleanCode)
    Else
        ' --- AMÉLIORATION DEBUG ---
        ' Si le code n'est pas trouvé, on le signale pour aider au diagnostic futur
        Debug.Print "Code non trouvé dans le dictionnaire : '" & cleanCode & "'"
        
        Set GetCodeInfo = New clsCodeInfo
        GetCodeInfo.code = "INCONNU"
    End If
End Function

Public Sub ReloadCodeDefinitions()
    ' Force reload of the configuration.
    Set codeDict = Nothing
    LoadCodeDefinitions
End Sub
Private Sub LoadCodeDefinitions()
    ' DESCRIPTION: Lit la table "tbl_Codes" et la charge dans le dictionnaire.
    '              (Version de DÉBOGAGE)
    
    Set codeDict = CreateObject("Scripting.Dictionary")
    codeDict.CompareMode = vbTextCompare ' Insensible à la casse

    Dim wsConfig As Worksheet, codeTable As ListObject, dataRange As Range
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Worksheets("Config_Codes")
    If wsConfig Is Nothing Then Exit Sub
    Set codeTable = wsConfig.ListObjects("tbl_Codes")
    If codeTable Is Nothing Then Exit Sub
    If codeTable.DataBodyRange Is Nothing Then Exit Sub
    Set dataRange = codeTable.DataBodyRange
    On Error GoTo 0
    
    Dim allData As Variant: allData = dataRange.value
    
    Dim i As Long, k As Integer
    Dim currentCode As clsCodeInfo
    Dim cellValue As Variant
    Dim codeLu As String ' Variable pour stocker le code lu en tant que texte

    ' On gère le cas d'une seule ligne vs plusieurs
    If dataRange.Rows.count = 1 Then
        Set currentCode = New clsCodeInfo
        With currentCode
            ' On lit la propriété .Text de la première cellule de la table
            codeLu = dataRange.Cells(1, 1).text
            .code = NormalizeString(codeLu) ' Utilise la fonction de nettoyage
            
            .description = CStr(allData(1, 2))
            .typeCode = CStr(allData(1, 3))
            .ColorCategory = CStr(allData(1, 4))
            
            For k = 1 To 13
                cellValue = allData(1, k + 4)
                If IsNumeric(cellValue) And Not IsEmpty(cellValue) Then
                    .Fractions(k) = CDbl(cellValue)
                End If
            Next k
            
            If .code <> "" Then
                ' --- AJOUT DE DÉBOGAGE (PARTIE 3) ---
                Debug.Print "CONFIG -> Ajouté au dictionnaire : '" & .code & "'"
                codeDict.Add .code, currentCode
            End If
        End With
    Else
        For i = 1 To UBound(allData, 1)
            Set currentCode = New clsCodeInfo
            With currentCode
                ' On lit la propriété .Text de la cellule de code correspondante
                codeLu = dataRange.Cells(i, 1).text
                .code = NormalizeString(codeLu) ' Utilise la fonction de nettoyage
            
                .description = CStr(allData(i, 2))
                .typeCode = CStr(allData(i, 3))
                .ColorCategory = CStr(allData(i, 4))
                
                For k = 1 To 13
                    cellValue = allData(i, k + 4)
                    If IsNumeric(cellValue) And Not IsEmpty(cellValue) Then
                        .Fractions(k) = CDbl(cellValue)
                    End If
                Next k
                
                If .code <> "" And Not codeDict.Exists(.code) Then
                    ' --- AJOUT DE DÉBOGAGE (PARTIE 4) ---
                    ' On affiche seulement le code qui nous intéresse pour ne pas tout polluer
                    If InStr(.code, "6:45") > 0 And InStr(.code, "15:15") > 0 Then
                         Debug.Print "CONFIG -> Ajouté au dictionnaire : '" & .code & "'"
                    End If
                    ' --- FIN DE L'AJOUT ---
                    
                    codeDict.Add .code, currentCode
                End If
            End With
        Next i
    End If
End Sub
