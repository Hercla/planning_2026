VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCodeInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' CLASSE :          clsCodeInfo (Version Finale Corrigée et Robuste)
'================================================================================

' --- Variables Publiques Simples ---
Public code As String
Public description As String
Public typeCode As String
Public ColorCategory As String

' --- Tableau Privé pour les Fractions ---
' On le déclare comme un tableau DYNAMIQUE (parenthèses vides).
Private m_fractions() As Double


'================================================================================
'   MÉTHODES PUBLIQUES ("PORTES") POUR ACCÉDER AUX FRACTIONS
'================================================================================

Public Property Get Fractions(ByVal index As Integer) As Double
    ' Permet de LIRE une valeur.
    If index >= LBound(m_fractions) And index <= UBound(m_fractions) Then
        Fractions = m_fractions(index)
    End If
End Property

Public Property Let Fractions(ByVal index As Integer, ByVal value As Double)
    ' Permet d'ÉCRIRE une valeur.
    If index >= LBound(m_fractions) And index <= UBound(m_fractions) Then
        m_fractions(index) = value
    End If
End Property


'================================================================================
'   INITIALISATION
'================================================================================

Private Sub Class_Initialize()
    ' S'assure qu'un nouvel objet est toujours propre à sa création.
    ' C'est ICI qu'on donne sa taille au tableau dynamique.
    ReDim m_fractions(1 To 13)
End Sub
