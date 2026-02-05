' ExportedAt: 2026-01-12 15:37:10 | Workbook: Planning_2026.xlsm
Attribute VB_Name = "Module_Utilities"
Option Explicit

'================================================================================================
' MODULE :          Module_Utilities
' DESCRIPTION :     Bibliothèque centrale de fonctions utilitaires publiques et robustes.
'                   - Lecture optimisée de la configuration via un cache.
'                   - Fonctions de gestion de dates normalisées.
'                   - Détection fiable des chemins système (OneDrive Pro/Perso).
'================================================================================================

' --- Constantes Privées du Module ---
' Centralise les noms importants pour une maintenance facile.
Private Const CONFIG_SHEET_NAME As String = "Configuration_GenerateNewWorkbo"

' --- Cache de Configuration (pour la performance) ---
' Ce dictionnaire stockera les paramètres après la première lecture.
Private configCache As Object

'================================================================================================
'   FONCTIONS DE GESTION DE LA CONFIGURATION
'================================================================================================

Public Function LireParametre(ByVal nomParametre As String) As String
    ' DESCRIPTION: Lit une valeur depuis la feuille de configuration.
    '              Utilise un cache (dictionnaire) pour des lectures ultra-rapides après le premier appel.
    
    ' Étape 1: Si le cache est vide, on le remplit en lisant la feuille de calcul une seule fois.
    If configCache Is Nothing Then LoadConfiguration

    ' Étape 2: On lit la valeur directement depuis le cache (très rapide).
    If configCache.Exists(nomParametre) Then
        LireParametre = configCache(nomParametre)
    Else
        LireParametre = "" ' Le paramètre n'existe pas dans le cache.
        Debug.Print "Attention : Le paramètre '" & nomParametre & "' n'a pas été trouvé dans la configuration."
    End If
End Function

Private Sub LoadConfiguration()
    ' DESCRIPTION: Procédure interne qui lit la feuille de configuration et remplit le dictionnaire "configCache".
    Dim wsConfig As Worksheet
    Dim configRange As Range
    Dim configData As Variant
    Dim i As Long
    
    ' Initialise le dictionnaire. La comparaison est insensible à la casse.
    Set configCache = CreateObject("Scripting.Dictionary")
    configCache.CompareMode = vbTextCompare

    ' Trouve la feuille de configuration.
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0
    
    If wsConfig Is Nothing Then
        Debug.Print "ALERTE : La feuille de configuration '" & CONFIG_SHEET_NAME & "' est introuvable. Le cache restera vide."
        Exit Sub
    End If
    
    ' Définit la plage à lire (de A1 à la dernière ligne de la colonne B).
    Set configRange = wsConfig.Range("A1:B" & wsConfig.Cells(wsConfig.Rows.Count, "A").End(xlUp).row)
    
    ' Lit toutes les données en une seule opération (très rapide).
    configData = configRange.value
    
    ' Remplit le dictionnaire.
    For i = LBound(configData, 1) To UBound(configData, 1)
        Dim key As String: key = Trim(CStr(configData(i, 1)))
        Dim value As String: value = CStr(configData(i, 2))
        
        If key <> "" And Not configCache.Exists(key) Then
            configCache.Add key, value
        End If
    Next i
End Sub

'================================================================================================
'   FONCTIONS DE GESTION DES DATES
'================================================================================================

Public Function DateToFrenchMonthName(ByVal d As Date) As String
    ' DESCRIPTION: Convertit une date en nom de mois français complet, indépendamment de la langue du système.
    Select Case Month(d)
        Case 1: DateToFrenchMonthName = "Janvier"
        Case 2: DateToFrenchMonthName = "Février"
        Case 3: DateToFrenchMonthName = "Mars"
        Case 4: DateToFrenchMonthName = "Avril"
        Case 5: DateToFrenchMonthName = "Mai"
        Case 6: DateToFrenchMonthName = "Juin"
        Case 7: DateToFrenchMonthName = "Juillet"
        Case 8: DateToFrenchMonthName = "Août"
        Case 9: DateToFrenchMonthName = "Septembre"
        Case 10: DateToFrenchMonthName = "Octobre"
        Case 11: DateToFrenchMonthName = "Novembre"
        Case 12: DateToFrenchMonthName = "Décembre"
    End Select
End Function

Public Function GetMonthDateFromName(ByVal monthNameInput As String) As Date
    ' DESCRIPTION: Interprète un nom d'onglet (ex: "Avril", "Juin 2024", "fev") et retourne le 1er jour du mois.
    Dim monthStr As String, yearStr As String
    Dim m As Integer, y As Integer
    Dim parts() As String
    
    ' Valeur par défaut en cas d'échec
    GetMonthDateFromName = CDate(0)
    
    ' 1. Nettoyer l'entrée
    monthNameInput = Trim(monthNameInput)
    If monthNameInput = "" Then Exit Function
    
    ' 2. Séparer le mois et l'année (ex: "Avril 2024")
    parts = Split(monthNameInput, " ")
    If UBound(parts) >= 1 And IsNumeric(parts(UBound(parts))) Then
        yearStr = parts(UBound(parts))
        monthStr = Trim(Replace(monthNameInput, yearStr, ""))
    Else
        monthStr = monthNameInput
        yearStr = ""
    End If

    ' 3. Trouver le numéro du mois à partir du nom (gère les abréviations)
    Select Case LCase(monthStr)
        Case "janvier", "janv": m = 1
        Case "février", "fevrier", "févr", "fevr", "fev": m = 2
        Case "mars": m = 3
        Case "avril", "avr": m = 4
        Case "mai": m = 5
        Case "juin": m = 6
        Case "juillet", "juil": m = 7
        Case "août", "aout", "aoû", "aou": m = 8
        Case "septembre", "sept": m = 9
        Case "octobre", "oct": m = 10
        Case "novembre", "nov": m = 11
        Case "décembre", "decembre", "déc", "dec": m = 12
        Case Else: Exit Function ' Mois non reconnu, on arrête.
    End Select

    ' 4. Déterminer l'année
    If yearStr <> "" Then
        y = CInt(yearStr)
        ' Assurer une année à 4 chiffres
        If y < 100 Then y = y + 2000
    Else
        y = Year(Date) ' Utilise l'année en cours par défaut
    End If
    
    GetMonthDateFromName = DateSerial(y, m, 1)
End Function
' Convertit une date en nom d'onglet abrégé (ex: "Janv", "Fev").
' Cette fonction est conservée ici car elle n'est pas dans Module_Utilities.
Public Function MonthToSheetName(d As Date) As String
    Dim arr As Variant
    arr = Array("Janv", "Fev", "Mars", "Avril", "Mai", "Juin", "Juil", "Aout", "Sept", "Oct", "Nov", "Dec")
    MonthToSheetName = arr(Month(d) - 1)
End Function

'================================================================================================
'   FONCTIONS SYSTÈME
'================================================================================================

Public Function GetOneDriveBasePath() As String
    ' DESCRIPTION: Trouve le chemin du dossier OneDrive de l'utilisateur de manière fiable.
    '              Teste d'abord les chemins OneDrive Pro, puis Perso, puis un fallback.
    Dim basePath As String
    
    ' Méthode 1: Variable d'environnement OneDrive pour les comptes PRO (le plus fiable pour vous)
    basePath = Environ("OneDriveCommercial")
    If basePath = "" Then basePath = Environ("OneDrive - Hôpital Universitaire de Bruxelles")
    
    ' Méthode 2: Variable d'environnement pour les comptes PERSONNELS
    If basePath = "" Then basePath = Environ("OneDrive")

    ' Méthode 3: Fallback en cherchant dans le profil utilisateur
    If basePath = "" Then
        Dim userProfile As String: userProfile = Environ("USERPROFILE")
        If Dir(userProfile & "\OneDrive", vbDirectory) <> "" Then basePath = userProfile & "\OneDrive"
    End If
    
    ' Vérification finale et formatage du chemin
    If basePath <> "" Then
        If Right(basePath, 1) <> "\" Then basePath = basePath & "\"
        GetOneDriveBasePath = basePath
    Else
        MsgBox "Le chemin du dossier OneDrive n'a pas pu être trouvé automatiquement.", vbCritical, "Erreur de Chemin"
        GetOneDriveBasePath = ""
    End If
End Function
Public Sub EnsurePathExists(ByVal fullPath As String)
    ' DESCRIPTION: S'assure que chaque dossier dans le chemin existe, en le créant si nécessaire.
    '              Contourne la limitation de MkDir qui ne peut créer qu'un niveau à la fois.
    Dim parts() As String
    Dim currentPath As String
    Dim i As Long
    
    ' Gère les chemins réseau (commençant par \\)
    If Left(fullPath, 2) = "\\" Then
        parts = Split(Mid(fullPath, 3), "\")
        currentPath = "\\" & parts(0) & "\" & parts(1) & "\" ' Ex: \\serveur\partage\
        i = 2
    Else
        parts = Split(fullPath, "\")
        currentPath = parts(0) & "\" ' Ex: "C:\"
        i = 1
    End If
    
    ' Boucle sur chaque partie du chemin et crée le dossier s'il est manquant
    For i = i To UBound(parts)
        If parts(i) <> "" Then
            currentPath = currentPath & parts(i) & "\"
            If Dir(currentPath, vbDirectory) = "" Then
                On Error Resume Next
                MkDir currentPath
                If Err.Number <> 0 Then
                    MsgBox "Impossible de créer le dossier : " & currentPath & vbCrLf & _
                           "Vérifiez vos permissions d'accès et le chemin.", vbCritical, "Erreur de création de dossier"
                    Err.Clear
                    Exit Sub ' Arrête la procédure pour éviter d'autres erreurs
                End If
                On Error GoTo 0
            End If
        End If
    Next i
End Sub
' --- AJOUTER CES FONCTIONS À MODULE_UTILITIES DANS UNE NOUVELLE SECTION "ARRAY UTILITIES" ---

Public Function IsInArray(ByVal itemToFind As Variant, ByRef arr As Variant) As Boolean
    ' DESCRIPTION: Vérifie si un élément existe dans un tableau 1D. Insensible à la casse pour les strings.
    Dim element As Variant
    On Error Resume Next
    For Each element In arr
        If StrComp(CStr(itemToFind), CStr(element), vbTextCompare) = 0 Then
            IsInArray = True
            Exit Function
        End If
    Next element
    On Error GoTo 0
End Function

Public Sub QuickSort(ByRef arr As Variant, Optional LB As Long = -1, Optional UB As Long = -1)
    ' DESCRIPTION: Trie un tableau 1D de variants (texte ou numérique) en utilisant l'algorithme QuickSort.
    If Not IsArray(arr) Then Exit Sub
    If LB = -1 Then LB = LBound(arr)
    If UB = -1 Then UB = UBound(arr)
    If LB >= UB Then Exit Sub

    Dim i As Long: i = LB
    Dim j As Long: j = UB
    Dim pivot As Variant: pivot = arr((LB + UB) \ 2)
    Dim temp As Variant

    Do While i <= j
        While StrComp(CStr(arr(i)), CStr(pivot), vbTextCompare) < 0: i = i + 1: Wend
        While StrComp(CStr(arr(j)), CStr(pivot), vbTextCompare) > 0: j = j - 1: Wend
        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop

    If LB < j Then QuickSort arr, LB, j
    If i < UB Then QuickSort arr, i, UB
End Sub
