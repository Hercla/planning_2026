Attribute VB_Name = "Module_Menu"
Option Explicit

  Public Sub Afficher_cacher_menu()
      Dim uf As UserForm1

      On Error Resume Next
      Set uf = UserForm1
      On Error GoTo 0

      If uf Is Nothing Then Exit Sub

      If uf.Visible Then
          uf.Hide
      Else
          PositionnerUserFormHautDroite uf
          uf.Show
      End If
  End Sub

  Private Sub PositionnerUserFormHautDroite(ByVal uf As Object)
      uf.StartUpPosition = 0 ' Manual

      Dim margin As Double: margin = 10
      Dim offsetLeft As Double: offsetLeft = 160 ' <- augmente pour aller plus à gauche

      Dim leftX As Double, topY As Double
      leftX = Application.ActiveWindow.Left + Application.ActiveWindow.width - uf.width - margin - offsetLeft
      topY = Application.ActiveWindow.Top + margin

      If leftX < Application.ActiveWindow.Left Then leftX = Application.ActiveWindow.Left + margin

      uf.Left = leftX
      uf.Top = topY
  End Sub
