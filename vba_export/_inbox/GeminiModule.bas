' ExportedAt: 2026-01-12 15:37:08 | Workbook: Planning_2026.xlsm
Attribute VB_Name = "GeminiModule"
Option Explicit

' Module pour intégrer l'API Gemini 2.5 Pro dans Excel/VBA.
' N'oubliez pas d'ajouter une référence à "Microsoft WinHTTP Services 5.1" et
' "Microsoft Scripting Runtime" dans Tools > References du VBA Editor.
' Ce module dépend également de JsonConverter.bas (VBA?JSON) pour analyser le
' JSON retourné par l'API.

Public Function CallGemini(prompt As String) As String
    Dim url As String
    url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-pro:generateContent"

    ' Clé API fournie par l'utilisateur.
    Dim apiKey As String
    apiKey = "AIzaSyDJaPMS4uJVYFpTU7Ivv2NW7o4xezcjz4k"

    ' Construire le corps JSON pour la requête.
    Dim jsonBody As String
    jsonBody = "{\"contents\":[{\"parts\":[{\"text\":\"" & _
               Replace(prompt, "\"", "\\\"") & "\"}]}]}"

    ' Initialiser l'objet HTTP et configurer les en?têtes.
    Dim http As WinHttp.WinHttpRequest
    Set http = New WinHttp.WinHttpRequest
    http.Open "POST", url, False
    http.SetRequestHeader "Content-Type", "application/json"
    http.SetRequestHeader "x-goog-api-key", apiKey

    ' Envoyer la requête JSON.
    http.Send jsonBody

    ' Vérifier la réponse HTTP.
    If http.Status <> 200 Then
        CallGemini = "Erreur API: " & http.Status & " - " & http.StatusText
        Exit Function
    End If

    ' Analyser la réponse JSON.
    Dim response As String
    response = http.ResponseText

    Dim json As Object
    Set json = JsonConverter.ParseJson(response)

    ' Extraire la première réponse dans candidates > content > parts > text
    Dim answer As String
    answer = json("candidates")(1)("content")("parts")(1)("text")

    CallGemini = answer
End Function

' Exemple de macro utilisant CallGemini pour générer une suggestion de planning.
Public Sub ObtenirSuggestion()
    Dim prompt As String
    prompt = "Voici les vacations prévues aujourd'hui pour l'unité neuro-traumato : " & _
             Range("A2:A10").value & _
             ". Donne?moi des suggestions pour optimiser les affectations."

    Dim reponse As String
    reponse = CallGemini(prompt)

    ' Écrire la réponse dans la cellule B2.
    Range("B2").value = reponse
End Sub


