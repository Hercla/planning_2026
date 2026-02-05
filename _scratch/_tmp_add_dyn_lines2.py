from pathlib import Path
p = Path(r"C:\Users\hercl\planning_2026\UserForm1.frm")
text = p.read_text(encoding="latin-1")
block = """Sub AfficherMasquerLignes()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lignes As Range
    Set lignes = ws.Rows("43:44")
    
    If lignes.Hidden = True Then
        lignes.Hidden = False ' Affiche les lignes 43 et 44
    Else
        lignes.Hidden = True ' Masque les lignes 43 et 44
    End If
End Sub"""
if block in text and "AfficherMasquerLignesDynamiques" not in text:
    stub = block + "\n\nPrivate Sub AfficherMasquerLignesDynamiques()\n    ' TODO: remplacer par la vraie logique si besoin\n    AfficherMasquerLignes\nEnd Sub"
    text = text.replace(block, stub, 1)
    p.write_text(text, encoding="latin-1")
    print("inserted")
else:
    print("no change")
